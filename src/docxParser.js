import JSZip from 'jszip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/**
 * Parses a .docx File or ArrayBuffer into an editable document structure
 * containing Q&A table rows (preserving rich formatting, marks, sources, images, and page breaks).
 */
export async function parseDocxDocument(fileOrBuffer, fileName = 'Document.docx') {
  const zip = await JSZip.loadAsync(fileOrBuffer);
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('無效的 .docx 檔案：缺少 word/document.xml');
  }

  const xmlText = await docXmlFile.async('text');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
  const imageMap = await loadDocxImages(zip, parser);

  const body = xmlDoc.getElementsByTagNameNS(W_NS, 'body')[0] || xmlDoc.querySelector('body');
  if (!body) {
    throw new Error('無法讀取 .docx 文件內容');
  }

  const rows = [];
  let docTitle = cleanFileName(fileName);
  let currentSection = docTitle;
  let pendingParagraphs = [];

  const children = Array.from(body.childNodes).filter((n) => n.nodeType === 1);

  for (const node of children) {
    const localName = node.localName || node.nodeName.replace(/^w:/, '');

    if (localName === 'p') {
      const hasPageBreak = nodeHasPageBreak(node);
      if (hasPageBreak && rows.length > 0) {
        rows[rows.length - 1].pageBreakAfter = true;
      }

      const pText = extractParagraphText(node).trim();
      if (pText) {
        const isHeading = checkIsHeading(node, pText);
        if (isHeading && pText.length < 150 && !looksLikeQAInline(pText)) {
          if (rows.length === 0 && docTitle === cleanFileName(fileName)) {
            docTitle = pText;
          }
          currentSection = pText;
        } else {
          pendingParagraphs.push({
            text: pText,
            html: extractParagraphHtml(node),
            images: extractNodeImages(node, imageMap),
            hasPageBreak,
          });
        }
      }
    } else if (localName === 'tbl') {
      const trNodes = getDirectChildElements(node, 'tr');
      for (let rIdx = 0; rIdx < trNodes.length; rIdx++) {
        const tr = trNodes[rIdx];
        const tcNodes = getDirectChildElements(tr, 'tc');
        if (tcNodes.length < 2) continue;

        const leftCell = extractCellData(tcNodes[0], imageMap);
        const rightCell = extractCellData(tcNodes[1], imageMap);

        const qRaw = leftCell.lines.join('\n').trim();
        const aRaw = rightCell.lines.join('\n').trim();
        if (!qRaw && !aRaw) continue;
        if (rIdx === 0 && isHeaderRow(qRaw, aRaw)) continue;

        const hasBreakInRow = nodeHasPageBreak(tr);
        rows.push(
          buildDocRow({
            id: `row-${rows.length + 1}-${hashStr(qRaw.slice(0, 48) + rIdx)}`,
            index: rows.length + 1,
            section: currentSection,
            questionHtml: leftCell.html,
            answerHtml: rightCell.html,
            questionRaw: qRaw || '(No question text)',
            answerRaw: aRaw || '(No answer text)',
            questionLines: leftCell.lines,
            answerLines: rightCell.lines,
            questionImages: leftCell.images,
            answerImages: rightCell.images,
            pageBreakAfter: hasBreakInRow,
          })
        );
      }
    }
  }

  // Fallback for non-table Q:/A: documents
  if (rows.length === 0 && pendingParagraphs.length > 0) {
    const fallbackRows = parseParagraphPairs(pendingParagraphs, currentSection);
    rows.push(...fallbackRows);
  }

  return {
    title: docTitle || cleanFileName(fileName),
    fileName,
    rows,
  };
}

/**
 * Converts a DocRow (possibly freshly edited in the Docs rich-text editor) into
 * clean structured Q&A metadata for both Docs view and simultaneous Flashcards.
 */
export function buildDocRow(raw) {
  const qLines =
    raw.questionLines && raw.questionLines.length
      ? [...raw.questionLines]
      : htmlToLines(raw.questionHtml || raw.questionRaw || '');
  const aLines =
    raw.answerLines && raw.answerLines.length
      ? [...raw.answerLines]
      : htmlToLines(raw.answerHtml || raw.answerRaw || '');

  const qLinesCopy = [...qLines];
  let source = raw.source || '';

  if (!source && qLinesCopy.length > 1) {
    const lastLine = qLinesCopy[qLinesCopy.length - 1].trim();
    if (/^\*|^Source\s*:|^來源\s*[:：]|^Ref\s*:/i.test(lastLine)) {
      source = lastLine.replace(/^\*\s*/, '').trim();
      qLinesCopy.pop();
    }
  }

  const questionFull = qLinesCopy.join('\n').trim() || qLines.join('\n').trim();
  let totalMarks = null;
  const marksMatch = questionFull.match(/[\(（]\s*(\d+)\s*(?:marks?|分|pts?)\s*[\)）]/i);
  if (marksMatch) {
    totalMarks = parseInt(marksMatch[1], 10);
  }

  const questionClean = questionFull
    .replace(/\s*[\(（]\s*\d+\s*(?:marks?|分|pts?)\s*[\)）]\s*$/i, '')
    .trim();

  const answerPoints = [];
  let summedMarks = 0;

  for (const line of aLines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const segments = splitInlineMarkingPoints(trimmed);
    for (const seg of segments) {
      const ptMatch = seg.match(/^(.*?)\s*[\(（](\d+)(?:\s*(?:marks?|分))?[\)）]\s*$/i);
      if (ptMatch && ptMatch[1].trim().length > 2) {
        const ptMarks = parseInt(ptMatch[2], 10);
        summedMarks += ptMarks;
        answerPoints.push({
          text: ptMatch[1].trim().replace(/^[-•*]\s*/, ''),
          marks: ptMarks,
        });
      } else {
        answerPoints.push({
          text: seg.replace(/^[-•*]\s*/, ''),
          marks: null,
        });
      }
    }
  }

  if (totalMarks === null && summedMarks > 0) {
    totalMarks = summedMarks;
  }

  const questionHtml =
    raw.questionHtml ||
    linesToSafeHtml(source ? [...qLinesCopy, `*${source}`] : qLinesCopy);
  const answerHtml = raw.answerHtml || linesToSafeHtml(aLines);

  return {
    id: raw.id || `row-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
    index: raw.index || 1,
    section: raw.section || '',
    question: questionClean || questionFull,
    questionWithMarks: questionFull,
    questionRaw: qLines.join('\n'),
    answerRaw: aLines.join('\n'),
    questionHtml,
    answerHtml,
    source,
    marks: totalMarks,
    answerPoints,
    questionImages: raw.questionImages || [],
    answerImages: raw.answerImages || [],
    pageBreakAfter: Boolean(raw.pageBreakAfter),
  };
}

/**
 * Re-parses a single row after inline editing in the Docs editor.
 */
export function updateRowFromHtml(existingRow, newQuestionHtml, newAnswerHtml) {
  const qLines = htmlToLines(newQuestionHtml);
  const aLines = htmlToLines(newAnswerHtml);
  return buildDocRow({
    ...existingRow,
    questionHtml: newQuestionHtml,
    answerHtml: newAnswerHtml,
    questionLines: qLines,
    answerLines: aLines,
    source: '', // Re-detect from edited text
  });
}

/**
 * Generates a real downloadable / Drive-uploadable .docx Blob from the current edited document,
 * preserving 2-column Q&A borders, rich text formatting (B/I/U/S/highlight), and page breaks!
 */
export async function exportDocToDocxBlob(doc) {
  const zip = new JSZip();

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );

  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  // Split into tables around explicit pageBreakAfter markers so Word renders clean page breaks
  const bodyBlocks = [];
  bodyBlocks.push(`
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${escapeXml(doc.title || 'Q&A Notes')}</w:t></w:r>
    </w:p>
  `);

  const tblPr = `
    <w:tblPr>
      <w:tblW w:w="9360" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="666666"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="666666"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="666666"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="666666"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="666666"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="666666"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid>
      <w:gridCol w:w="4320"/>
      <w:gridCol w:w="5040"/>
    </w:tblGrid>
  `;

  let currentTableRows = [];
  const flushTable = () => {
    if (currentTableRows.length === 0) return;
    bodyBlocks.push(`<w:tbl>${tblPr}${currentTableRows.join('')}</w:tbl>`);
    currentTableRows = [];
  };

  for (let i = 0; i < doc.rows.length; i++) {
    const r = doc.rows[i];
    const leftXml = htmlToWordParagraphsXml(r.questionHtml || linesToSafeHtml([r.questionWithMarks || r.question]));
    const rightXml = htmlToWordParagraphsXml(r.answerHtml || linesToSafeHtml([r.answerRaw || '']));

    currentTableRows.push(`
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>${leftXml}</w:tc>
        <w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/></w:tcPr>${rightXml}</w:tc>
      </w:tr>
    `);

    if (r.pageBreakAfter && i < doc.rows.length - 1) {
      flushTable();
      bodyBlocks.push(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`);
    }
  }
  flushTable();

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyBlocks.join('\n')}
  </w:body>
</w:document>`;

  zip.folder('word').file('document.xml', docXml);
  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/* ==================== Internal Helpers ==================== */

function htmlToWordParagraphsXml(html) {
  const div = document.createElement('div');
  div.innerHTML = html || '';

  const paragraphs = [];
  let currentRuns = [];

  const flushParagraph = () => {
    if (currentRuns.length === 0) {
      paragraphs.push('<w:p/>');
    } else {
      paragraphs.push(`<w:p>${currentRuns.join('')}</w:p>`);
    }
    currentRuns = [];
  };

  const walk = (node, fmt) => {
    if (node.nodeType === 3) {
      const txt = node.nodeValue || '';
      if (!txt) return;
      const rPrParts = [];
      if (fmt.b) rPrParts.push('<w:b/>');
      if (fmt.i) rPrParts.push('<w:i/>');
      if (fmt.u) rPrParts.push('<w:u w:val="single"/>');
      if (fmt.s) rPrParts.push('<w:strike/>');
      if (fmt.hl) rPrParts.push('<w:highlight w:val="yellow"/>');
      const rPr = rPrParts.length ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
      currentRuns.push(`<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(txt)}</w:t></w:r>`);
      return;
    }

    if (node.nodeType !== 1) return;
    const tag = node.tagName.toLowerCase();

    if (tag === 'br') {
      flushParagraph();
      return;
    }

    const nextFmt = {
      b: fmt.b || tag === 'b' || tag === 'strong',
      i: fmt.i || tag === 'i' || tag === 'em',
      u: fmt.u || tag === 'u',
      s: fmt.s || tag === 's' || tag === 'strike' || tag === 'del',
      hl: fmt.hl || tag === 'mark',
    };

    const isBlock = tag === 'p' || tag === 'div' || tag === 'li';
    if (isBlock && currentRuns.length > 0) {
      flushParagraph();
    }

    for (const child of node.childNodes) {
      walk(child, nextFmt);
    }

    if (isBlock) {
      flushParagraph();
    }
  };

  walk(div, { b: false, i: false, u: false, s: false, hl: false });
  if (currentRuns.length > 0) flushParagraph();

  return paragraphs.length > 0 ? paragraphs.join('') : '<w:p/>';
}

function htmlToLines(html) {
  const div = document.createElement('div');
  div.innerHTML = String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li)>/gi, '\n');
  return (div.textContent || '')
    .split('\n')
    .map((s) => s.trim())
    .filter(Boolean);
}

function linesToSafeHtml(lines) {
  return (lines || [])
    .filter((l) => String(l).trim().length > 0)
    .map((l) => `<p>${escapeHtml(l)}</p>`)
    .join('');
}

function extractCellData(tcNode, imageMap) {
  const lines = [];
  const htmlParts = [];
  const images = extractNodeImages(tcNode, imageMap);
  const paragraphs = tcNode.getElementsByTagNameNS(W_NS, 'p');

  for (let i = 0; i < paragraphs.length; i++) {
    const pText = extractParagraphText(paragraphs[i]);
    const pHtml = extractParagraphHtml(paragraphs[i]);
    const subLines = pText
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    lines.push(...subLines);
    if (pHtml.trim()) {
      htmlParts.push(`<p>${pHtml}</p>`);
    }
  }
  return {
    lines,
    html: htmlParts.join('') || linesToSafeHtml(lines),
    images,
  };
}

function extractParagraphHtml(pNode) {
  let out = '';
  const walk = (node) => {
    if (node.nodeType !== 1) return;
    const name = node.localName || node.nodeName.replace(/^w:/, '');
    if (name === 'r') {
      const rPr = getDirectChildElements(node, 'rPr')[0];
      const isB = rPr && hasToggleTag(rPr, 'b');
      const isI = rPr && hasToggleTag(rPr, 'i');
      const isU = rPr && rPr.getElementsByTagNameNS(W_NS, 'u').length > 0;
      const isS = rPr && hasToggleTag(rPr, 'strike');
      const isHl = rPr && rPr.getElementsByTagNameNS(W_NS, 'highlight').length > 0;

      let runText = '';
      for (let i = 0; i < node.childNodes.length; i++) {
        const ch = node.childNodes[i];
        if (ch.nodeType !== 1) continue;
        const chName = ch.localName || ch.nodeName.replace(/^w:/, '');
        if (chName === 't') runText += escapeHtml(ch.textContent || '');
        else if (chName === 'br' || chName === 'cr') {
          const brType = ch.getAttributeNS(W_NS, 'type') || ch.getAttribute('w:type');
          if (brType !== 'page') runText += '<br>';
        } else if (chName === 'tab') runText += ' &nbsp;&nbsp; ';
      }

      if (runText) {
        if (isHl) runText = `<mark>${runText}</mark>`;
        if (isS) runText = `<s>${runText}</s>`;
        if (isU) runText = `<u>${runText}</u>`;
        if (isI) runText = `<i>${runText}</i>`;
        if (isB) runText = `<b>${runText}</b>`;
        out += runText;
      }
      return;
    }

    for (let i = 0; i < node.childNodes.length; i++) {
      walk(node.childNodes[i]);
    }
  };
  walk(pNode);
  return out;
}

function hasToggleTag(rPr, tagName) {
  const el = rPr.getElementsByTagNameNS(W_NS, tagName)[0];
  if (!el) return false;
  const val = el.getAttributeNS(W_NS, 'val') || el.getAttribute('w:val');
  return val !== '0' && val !== 'false';
}

function nodeHasPageBreak(node) {
  const brs = node.getElementsByTagNameNS(W_NS, 'br');
  for (let i = 0; i < brs.length; i++) {
    const t = brs[i].getAttributeNS(W_NS, 'type') || brs[i].getAttribute('w:type');
    if (t === 'page') return true;
  }
  const rendered = node.getElementsByTagNameNS(W_NS, 'lastRenderedPageBreak');
  return rendered.length > 0;
}

function splitInlineMarkingPoints(line) {
  const parts = line.split(/(?<=[\(（]\d+[\)）])\s+(?=[A-Z\u4e00-\u9fa5])/);
  return parts.map((p) => p.trim()).filter(Boolean);
}

function isHeaderRow(q, a) {
  const qNorm = q.toLowerCase().trim();
  const aNorm = a.toLowerCase().trim();
  const qHeaders = ['question', 'questions', 'prompt', 'term', '問題', '题目', '題目', '問'];
  const aHeaders = ['answer', 'answers', 'marking scheme', 'definition', 'solution', '答案', '解答', '評分標準', '答'];
  return qHeaders.includes(qNorm) && aHeaders.includes(aNorm);
}

function looksLikeQAInline(text) {
  return /^(?:Q\d*|Question\s*\d*|問\d*|題目\s*\d*)\s*[:：.]/i.test(text);
}

function parseParagraphPairs(paragraphs, sectionTitle) {
  const rows = [];
  let currentQ = null;
  let currentA = [];

  const flush = () => {
    if (currentQ && currentA.length > 0) {
      const aLines = currentA.map((x) => x.text);
      const aHtml = currentA.map((x) => `<p>${x.html || escapeHtml(x.text)}</p>`).join('');
      rows.push(
        buildDocRow({
          id: `row-${rows.length + 1}-${hashStr(currentQ.text.slice(0, 40) + rows.length)}`,
          index: rows.length + 1,
          section: sectionTitle,
          questionRaw: currentQ.text,
          answerRaw: aLines.join('\n'),
          questionHtml: `<p>${currentQ.html || escapeHtml(currentQ.text)}</p>`,
          answerHtml: aHtml,
          questionLines: [currentQ.text],
          answerLines: aLines,
          questionImages: currentQ.images || [],
          answerImages: currentA.flatMap((x) => x.images || []),
          pageBreakAfter: currentA.some((x) => x.hasPageBreak),
        })
      );
    }
    currentQ = null;
    currentA = [];
  };

  for (const item of paragraphs) {
    const t = item.text;
    const qPrefix = t.match(/^(?:Q\d*|Question\s*\d*|問\d*|問題\s*\d*)\s*[:：.]\s*(.+)$/i);
    const aPrefix = t.match(/^(?:A\d*|Ans(?:wer)?\s*\d*|答\d*|答案\s*\d*)\s*[:：.]\s*(.+)$/i);

    if (qPrefix) {
      flush();
      currentQ = { text: qPrefix[1].trim(), html: escapeHtml(qPrefix[1].trim()), images: item.images };
    } else if (aPrefix) {
      if (!currentQ) currentQ = { text: 'Question', html: 'Question', images: [] };
      currentA.push({ text: aPrefix[1].trim(), html: escapeHtml(aPrefix[1].trim()), images: item.images, hasPageBreak: item.hasPageBreak });
    } else if (t.includes('?') || t.includes('？')) {
      if (currentQ && currentA.length > 0) flush();
      if (!currentQ) currentQ = item;
      else currentA.push(item);
    } else if (currentQ) {
      currentA.push(item);
    }
  }
  flush();

  if (rows.length === 0 && paragraphs.length >= 2) {
    for (let i = 0; i + 1 < paragraphs.length; i += 2) {
      rows.push(
        buildDocRow({
          id: `row-${rows.length + 1}-${hashStr(paragraphs[i].text.slice(0, 32) + i)}`,
          index: rows.length + 1,
          section: sectionTitle,
          questionRaw: paragraphs[i].text,
          answerRaw: paragraphs[i + 1].text,
          questionHtml: `<p>${paragraphs[i].html || escapeHtml(paragraphs[i].text)}</p>`,
          answerHtml: `<p>${paragraphs[i + 1].html || escapeHtml(paragraphs[i + 1].text)}</p>`,
          questionLines: [paragraphs[i].text],
          answerLines: [paragraphs[i + 1].text],
          questionImages: paragraphs[i].images || [],
          answerImages: paragraphs[i + 1].images || [],
        })
      );
    }
  }

  return rows;
}

function getDirectChildElements(parent, localTagName) {
  const result = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1) {
      const name = child.localName || child.nodeName.replace(/^w:/, '');
      if (name === localTagName) result.push(child);
    }
  }
  return result;
}

function checkIsHeading(pNode, text) {
  const pStyles = pNode.getElementsByTagNameNS(W_NS, 'pStyle');
  if (pStyles.length > 0) {
    const val = pStyles[0].getAttributeNS(W_NS, 'val') || pStyles[0].getAttribute('w:val') || '';
    if (/heading|title|subtitle/i.test(val)) return true;
  }
  const runs = pNode.getElementsByTagNameNS(W_NS, 'r');
  if (runs.length > 0 && text.length < 110) {
    let allBold = true;
    for (let i = 0; i < runs.length; i++) {
      if (!(runs[i].textContent || '').trim()) continue;
      if (!runs[i].getElementsByTagNameNS(W_NS, 'b')[0]) {
        allBold = false;
        break;
      }
    }
    if (allBold) return true;
  }
  return false;
}

function extractParagraphText(pNode) {
  let out = '';
  const walk = (node) => {
    if (node.nodeType !== 1) return;
    const name = node.localName || node.nodeName.replace(/^w:/, '');
    if (name === 't') {
      out += node.textContent || '';
      return;
    }
    if (name === 'br' || name === 'cr') {
      out += '\n';
      return;
    }
    if (name === 'tab') {
      out += '\t';
      return;
    }
    for (let i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i]);
  };
  walk(pNode);
  return out;
}

async function loadDocxImages(zip, parser) {
  const map = new Map();
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (!relsFile) return map;
  try {
    const relsXml = await relsFile.async('text');
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');
    const rels = relsDoc.getElementsByTagNameNS(REL_NS, 'Relationship');
    for (let i = 0; i < rels.length; i++) {
      const r = rels[i];
      const id = r.getAttribute('Id');
      const target = r.getAttribute('Target') || '';
      const type = r.getAttribute('Type') || '';
      if (id && type.includes('/image')) {
        const normalizedPath = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\//, '')}`;
        const imgFile = zip.file(normalizedPath);
        if (imgFile) {
          const blob = await imgFile.async('blob');
          map.set(id, URL.createObjectURL(blob));
        }
      }
    }
  } catch {
    // Graceful fallback
  }
  return map;
}

function extractNodeImages(node, imageMap) {
  if (!imageMap || imageMap.size === 0) return [];
  const urls = [];
  const blips = node.getElementsByTagName('a:blip');
  for (let i = 0; i < blips.length; i++) {
    const embedId = blips[i].getAttribute('r:embed');
    if (embedId && imageMap.has(embedId)) urls.push(imageMap.get(embedId));
  }
  return urls;
}

function cleanFileName(name) {
  return String(name || 'Document').replace(/\.docx$/i, '').replace(/[_-]+/g, ' ').trim();
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeHtml(str) {
  return escapeXml(str);
}

function hashStr(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) {
    h = ((h << 5) - h + str.charCodeAt(i)) | 0;
  }
  return Math.abs(h).toString(36).slice(0, 6);
}

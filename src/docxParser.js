import JSZip from 'jszip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/**
 * Extracts Q&A flashcard items from a .docx File or ArrayBuffer.
 * Supports:
 * 1. Two-column (or multi-column) Q&A tables (like HKDSE_Bio_Ch8_QA_Notes.docx)
 * 2. Q: / A: or 問：/ 答： paragraph pairs
 * 3. Heading + body paragraph pairs
 */
export async function parseDocxToCards(fileOrBuffer, fileName = 'Document.docx') {
  const zip = await JSZip.loadAsync(fileOrBuffer);
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('Invalid .docx file: missing word/document.xml');
  }

  const xmlText = await docXmlFile.async('text');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, 'application/xml');

  // Optional: Load image relationships if the docx embeds diagrams in questions/answers
  const imageMap = await loadDocxImages(zip, parser);

  const body = xmlDoc.getElementsByTagNameNS(W_NS, 'body')[0] || xmlDoc.querySelector('body');
  if (!body) {
    throw new Error('Could not read document body from .docx');
  }

  const cards = [];
  let currentSectionTitle = cleanFileName(fileName);
  let pendingParagraphs = [];

  // Iterate direct children of <w:body> in document order so section titles align with tables
  const children = Array.from(body.childNodes).filter((n) => n.nodeType === 1);

  for (const node of children) {
    const localName = node.localName || node.nodeName.replace(/^w:/, '');

    if (localName === 'p') {
      const pText = extractParagraphText(node).trim();
      if (pText) {
        const isHeading = checkIsHeading(node, pText);
        if (isHeading && pText.length < 140 && ! looksLikeQAInline(pText)) {
          currentSectionTitle = pText;
        } else {
          pendingParagraphs.push({ text: pText, images: extractNodeImages(node, imageMap) });
        }
      }
    } else if (localName === 'tbl') {
      const rows = getDirectChildElements(node, 'tr');
      for (let rIdx = 0; rIdx < rows.length; rIdx++) {
        const row = rows[rIdx];
        const cells = getDirectChildElements(row, 'tc');
        if (cells.length < 2) continue;

        const leftCell = extractCellLines(cells[0], imageMap);
        const rightCell = extractCellLines(cells[1], imageMap);

        const qRaw = leftCell.lines.join('\n').trim();
        const aRaw = rightCell.lines.join('\n').trim();

        if (!qRaw && !aRaw) continue;

        // Skip pure table header rows like "Question | Answer" or "問題 | 答案"
        if (rIdx === 0 && isHeaderRow(qRaw, aRaw)) {
          continue;
        }

        if (qRaw || aRaw) {
          cards.push(
            enrichCard({
              id: `card-${cards.length + 1}-${Date.now().toString(36)}`,
              index: cards.length + 1,
              section: currentSectionTitle,
              questionRaw: qRaw || '(No question text)',
              answerRaw: aRaw || '(No answer text)',
              questionLines: leftCell.lines,
              answerLines: rightCell.lines,
              questionImages: leftCell.images,
              answerImages: rightCell.images,
            })
          );
        }
      }
    }
  }

  // Fallback: if no table cards were found, parse paragraph Q&A pairs
  if (cards.length === 0 && pendingParagraphs.length > 0) {
    const fallbackCards = parseParagraphPairs(pendingParagraphs, currentSectionTitle);
    cards.push(...fallbackCards);
  }

  return {
    title: currentSectionTitle || cleanFileName(fileName),
    fileName,
    cards,
  };
}

/**
 * Generates a real downloadable .docx Blob matching the HKDSE Biology Q&A table format
 * so users can export or test uploading a genuine .docx file!
 */
export async function generateSampleDocxBlob(deck) {
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

  const escapeXml = (str) =>
    String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');

  const rowsXml = deck.cards
    .map((c) => {
      const qParas = [
        c.question,
        c.source ? `*${c.source}` : '',
      ]
        .filter(Boolean)
        .map((line) => `<w:p><w:r><w:t>${escapeXml(line)}</w:t></w:r></w:p>`)
        .join('');

      const aParas = (c.answerPoints && c.answerPoints.length
        ? c.answerPoints.map((pt) => `${pt.text}${pt.marks ? ` (${pt.marks})` : ''}`)
        : String(c.answerRaw || '').split('\n').filter(Boolean)
      )
        .map((line) => `<w:p><w:r><w:t>${escapeXml(line)}</w:t></w:r></w:p>`)
        .join('');

      return `<w:tr>
        <w:tc>${qParas}</w:tc>
        <w:tc>${aParas}</w:tc>
      </w:tr>`;
    })
    .join('\n');

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(deck.title)}</w:t></w:r>
    </w:p>
    <w:tbl>
      ${rowsXml}
    </w:tbl>
  </w:body>
</w:document>`;

  zip.folder('word').file('document.xml', docXml);
  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

function enrichCard(rawCard) {
  const qLines = [...rawCard.questionLines];
  let source = '';

  // Detect footnote/source line at the bottom of the Question cell (e.g. "*OCR GCE (AS) ..." or "*Oxford Ch.8 QB (p. 8-95)")
  if (qLines.length > 1) {
    const lastLine = qLines[qLines.length - 1].trim();
    if (/^\*|^Source\s*:|^來源\s*[:：]|^Ref\s*:/i.test(lastLine)) {
      source = lastLine.replace(/^\*\s*/, '').trim();
      qLines.pop();
    }
  }

  const questionFull = qLines.join('\n').trim();
  // Detect marks in question e.g. "(3 marks)" or "(1 mark)" or "（3分）"
  let totalMarks = null;
  const marksMatch = questionFull.match(/[\(（]\s*(\d+)\s*(?:marks?|分|pts?)\s*[\)）]/i);
  if (marksMatch) {
    totalMarks = parseInt(marksMatch[1], 10);
  }

  // Clean display question (keep original text readable, optionally strip trailing marks badge so we can render a crisp pill)
  const questionClean = questionFull
    .replace(/\s*[\(（]\s*\d+\s*(?:marks?|分|pts?)\s*[\)）]\s*$/i, '')
    .trim();

  // Parse answer marking points e.g. "The blood pressure in aorta changes periodically. (1)"
  const answerPoints = [];
  let summedMarks = 0;

  for (const line of rawCard.answerLines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    // Split inline marking points if multiple "(1)" appear in a single paragraph
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

  return {
    ...rawCard,
    question: questionClean || questionFull,
    questionWithMarks: questionFull,
    source,
    marks: totalMarks,
    answerPoints,
    mastered: false,
    bookmarked: false,
  };
}

function splitInlineMarkingPoints(line) {
  // If a paragraph has multiple sentences each ending in (1) e.g. "... periodically. (1) It drops ... (1)"
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
  const cards = [];
  let currentQ = null;
  let currentA = [];

  const flush = () => {
    if (currentQ && currentA.length > 0) {
      const aLines = currentA.map((x) => x.text);
      cards.push(
        enrichCard({
          id: `card-${cards.length + 1}-${Date.now().toString(36)}`,
          index: cards.length + 1,
          section: sectionTitle,
          questionRaw: currentQ.text,
          answerRaw: aLines.join('\n'),
          questionLines: [currentQ.text],
          answerLines: aLines,
          questionImages: currentQ.images || [],
          answerImages: currentA.flatMap((x) => x.images || []),
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
      currentQ = { text: qPrefix[1].trim(), images: item.images };
    } else if (aPrefix) {
      if (!currentQ) {
        currentQ = { text: 'Question', images: [] };
      }
      currentA.push({ text: aPrefix[1].trim(), images: item.images });
    } else if (t.includes('?') || t.includes('？')) {
      if (currentQ && currentA.length > 0) {
        flush();
      }
      if (!currentQ) {
        currentQ = { text: t, images: item.images };
      } else {
        currentA.push(item);
      }
    } else if (currentQ) {
      currentA.push(item);
    }
  }
  flush();

  // Last-resort pair sequential paragraphs 2-by-2 if still empty
  if (cards.length === 0 && paragraphs.length >= 2) {
    for (let i = 0; i + 1 < paragraphs.length; i += 2) {
      cards.push(
        enrichCard({
          id: `card-${cards.length + 1}-${Date.now().toString(36)}`,
          index: cards.length + 1,
          section: sectionTitle,
          questionRaw: paragraphs[i].text,
          answerRaw: paragraphs[i + 1].text,
          questionLines: [paragraphs[i].text],
          answerLines: [paragraphs[i + 1].text],
          questionImages: paragraphs[i].images || [],
          answerImages: paragraphs[i + 1].images || [],
        })
      );
    }
  }

  return cards;
}

function getDirectChildElements(parent, localTagName) {
  const result = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1) {
      const name = child.localName || child.nodeName.replace(/^w:/, '');
      if (name === localTagName) {
        result.push(child);
      }
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
  // Also check if entire short line is bold outside a table (like "HKDSE BIOLOGY — Chapter 8: Transport in Humans")
  const runs = pNode.getElementsByTagNameNS(W_NS, 'r');
  if (runs.length > 0 && text.length < 110) {
    let allBold = true;
    for (let i = 0; i < runs.length; i++) {
      const rText = runs[i].textContent || '';
      if (!rText.trim()) continue;
      const bTag = runs[i].getElementsByTagNameNS(W_NS, 'b')[0];
      if (!bTag) {
        allBold = false;
        break;
      }
    }
    if (allBold) return true;
  }
  return false;
}

function extractCellLines(tcNode, imageMap) {
  const lines = [];
  const images = extractNodeImages(tcNode, imageMap);
  const paragraphs = tcNode.getElementsByTagNameNS(W_NS, 'p');

  for (let i = 0; i < paragraphs.length; i++) {
    const pText = extractParagraphText(paragraphs[i]);
    const subLines = pText
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    lines.push(...subLines);
  }
  return { lines, images };
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
    for (let i = 0; i < node.childNodes.length; i++) {
      walk(node.childNodes[i]);
    }
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
        const normalizedPath = target.startsWith('/')
          ? target.slice(1)
          : `word/${target.replace(/^\.\//, '')}`;
        const imgFile = zip.file(normalizedPath);
        if (imgFile) {
          const blob = await imgFile.async('blob');
          const url = URL.createObjectURL(blob);
          map.set(id, url);
        }
      }
    }
  } catch {
    // Ignore image extraction errors gracefully
  }
  return map;
}

function extractNodeImages(node, imageMap) {
  if (!imageMap || imageMap.size === 0) return [];
  const urls = [];
  const blips = node.getElementsByTagName('a:blip');
  for (let i = 0; i < blips.length; i++) {
    const embedId = blips[i].getAttribute('r:embed');
    if (embedId && imageMap.has(embedId)) {
      urls.push(imageMap.get(embedId));
    }
  }
  return urls;
}

function cleanFileName(name) {
  return name.replace(/\.docx$/i, '').replace(/[_-]+/g, ' ').trim();
}

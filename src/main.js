import { parseDocxToCards, generateSampleDocxBlob } from './docxParser.js';
import { SAMPLE_BIO_DECK } from './sampleData.js';

const state = {
  title: SAMPLE_BIO_DECK.title,
  fileName: SAMPLE_BIO_DECK.fileName,
  cards: structuredCloneCards(SAMPLE_BIO_DECK.cards),
  viewMode: 'stream', // 'stream' (side-by-side list) | 'focus' (single simultaneous Q&A card carousel)
  filter: 'all', // 'all' | 'unmastered' | 'bookmarked'
  searchQuery: '',
  focusIndex: 0,
  checkedPoints: new Set(), // key: `${cardId}:${pointIndex}`
  theme: localStorage.getItem('lumina-theme') || 'light',
};

function structuredCloneCards(cards) {
  return cards.map((c) => ({
    ...c,
    answerPoints: (c.answerPoints || []).map((pt) => ({ ...pt })),
    questionImages: [...(c.questionImages || [])],
    answerImages: [...(c.answerImages || [])],
  }));
}

// DOM References
const els = {
  html: document.documentElement,
  deckTitle: document.getElementById('deck-title'),
  deckFilename: document.getElementById('deck-filename'),
  statMastered: document.getElementById('stat-mastered'),
  statMarks: document.getElementById('stat-marks'),
  masteryBar: document.getElementById('mastery-bar'),
  fileInput: document.getElementById('docx-file-input'),
  btnUpload: document.getElementById('btn-upload-docx'),
  btnToggleDrawer: document.getElementById('btn-toggle-drawer'),
  btnThemeToggle: document.getElementById('btn-theme-toggle'),
  iconSun: document.getElementById('icon-sun'),
  iconMoon: document.getElementById('icon-moon'),
  modeStreamBtn: document.getElementById('mode-stream-btn'),
  modeFocusBtn: document.getElementById('mode-focus-btn'),
  filterBtns: document.querySelectorAll('[data-filter]'),
  searchInput: document.getElementById('search-input'),
  btnClearSearch: document.getElementById('btn-clear-search'),
  dropzoneDrawer: document.getElementById('dropzone-drawer'),
  globalDropVeil: document.getElementById('global-drop-veil'),
  btnDownloadSample: document.getElementById('btn-download-sample-docx'),
  btnResetSample: document.getElementById('btn-reset-sample'),
  cardsStage: document.getElementById('cards-stage'),
  dockPrevBtn: document.getElementById('dock-prev-btn'),
  dockNextBtn: document.getElementById('dock-next-btn'),
  dockCounter: document.getElementById('dock-counter'),
  dockShuffleBtn: document.getElementById('dock-shuffle-btn'),
  dockCheckAllBtn: document.getElementById('dock-checkall-btn'),
  toastPill: document.getElementById('toast-pill'),
  toastText: document.getElementById('toast-text'),
};

let toastTimer = null;

function init() {
  applyTheme(state.theme);
  bindEvents();
  renderAll();
}

function applyTheme(theme) {
  state.theme = theme;
  els.html.setAttribute('data-theme', theme);
  localStorage.setItem('lumina-theme', theme);
  if (theme === 'dark') {
    els.iconSun.style.display = 'block';
    els.iconMoon.style.display = 'none';
  } else {
    els.iconSun.style.display = 'none';
    els.iconMoon.style.display = 'block';
  }
}

function showToast(message) {
  els.toastText.textContent = message;
  els.toastPill.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    els.toastPill.classList.remove('show');
  }, 2800);
}

function getFilteredCards() {
  const q = state.searchQuery.trim().toLowerCase();
  return state.cards.filter((card) => {
    if (state.filter === 'unmastered' && card.mastered) return false;
    if (state.filter === 'bookmarked' && !card.bookmarked) return false;
    if (!q) return true;

    const haystacks = [
      card.question,
      card.answerRaw,
      card.source || '',
      ...(card.answerPoints || []).map((p) => p.text),
    ];
    return haystacks.some((h) => h.toLowerCase().includes(q));
  });
}

function renderAll() {
  // Header Info
  els.deckTitle.textContent = state.title;
  els.deckFilename.textContent = state.fileName;

  // Stats
  const total = state.cards.length;
  const masteredCount = state.cards.filter((c) => c.mastered).length;
  const pct = total > 0 ? Math.round((masteredCount / total) * 100) : 0;
  els.statMastered.textContent = `${masteredCount} / ${total}`;
  els.masteryBar.style.width = `${pct}%`;

  const totalMarks = state.cards.reduce((sum, c) => sum + (Number(c.marks) || 0), 0);
  els.statMarks.textContent = totalMarks > 0 ? `${totalMarks} marks` : `${total} 張閃卡`;

  // View Mode Classes
  els.cardsStage.className = `cards-stage mode-${state.viewMode}`;
  els.modeStreamBtn.classList.toggle('active', state.viewMode === 'stream');
  els.modeStreamBtn.setAttribute('aria-selected', String(state.viewMode === 'stream'));
  els.modeFocusBtn.classList.toggle('active', state.viewMode === 'focus');
  els.modeFocusBtn.setAttribute('aria-selected', String(state.viewMode === 'focus'));

  // Filter Active State
  els.filterBtns.forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.filter === state.filter);
  });

  // Search Clear Button
  els.btnClearSearch.classList.toggle('visible', state.searchQuery.length > 0);

  // Cards Rendering
  const filtered = getFilteredCards();

  if (filtered.length === 0) {
    els.dockCounter.textContent = `0 / 0`;
    els.cardsStage.innerHTML = `
      <div class="empty-state-card squircle-2xl">
        <div class="brand-emblem squircle-md" style="margin: 0 auto">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="11" cy="11" r="8"></circle>
            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          </svg>
        </div>
        <h3>沒有符合目前篩選條件的問答閃卡</h3>
        <p>可嘗試清除搜尋關鍵字、切換篩選標籤，或上傳新的 .docx 檔案。</p>
        <button id="btn-empty-reset" class="neu-btn squircle-md" type="button">顯示全部閃卡</button>
      </div>
    `;
    const resetBtn = document.getElementById('btn-empty-reset');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        state.filter = 'all';
        state.searchQuery = '';
        els.searchInput.value = '';
        renderAll();
      });
    }
    return;
  }

  if (state.focusIndex >= filtered.length) {
    state.focusIndex = Math.max(0, filtered.length - 1);
  }

  if (state.viewMode === 'focus') {
    const activeCard = filtered[state.focusIndex];
    els.dockCounter.textContent = `${state.focusIndex + 1} / ${filtered.length}`;
    els.cardsStage.innerHTML = renderCardHTML(activeCard, state.focusIndex, filtered.length);
  } else {
    els.dockCounter.textContent = `共 ${filtered.length} 張`;
    els.cardsStage.innerHTML = filtered
      .map((card, idx) => renderCardHTML(card, idx, filtered.length))
      .join('');
  }

  attachCardInteractions();
}

function renderCardHTML(card, idx) {
  const qHighlighted = highlightText(card.question, state.searchQuery);
  const sourceBadge = card.source
    ? `<span class="badge-source squircle-sm" title="${escapeAttr(card.source)}">*${highlightText(card.source, state.searchQuery)}</span>`
    : '';
  const marksBadge =
    card.marks != null
      ? `<span class="badge-marks squircle-sm">${card.marks} ${card.marks === 1 ? 'mark' : 'marks'}</span>`
      : '';

  const qImagesHtml =
    card.questionImages && card.questionImages.length
      ? `<div class="docx-images-grid">${card.questionImages
          .map((src) => `<img src="${src}" alt="問題附圖" loading="lazy" />`)
          .join('')}</div>`
      : '';

  const aImagesHtml =
    card.answerImages && card.answerImages.length
      ? `<div class="docx-images-grid">${card.answerImages
          .map((src) => `<img src="${src}" alt="答案附圖" loading="lazy" />`)
          .join('')}</div>`
      : '';

  const points =
    card.answerPoints && card.answerPoints.length > 0
      ? card.answerPoints
      : [{ text: card.answerRaw || '', marks: null }];

  const pointsHtml = points
    .map((pt, pIdx) => {
      const pointKey = `${card.id}:${pIdx}`;
      const isChecked = state.checkedPoints.has(pointKey);
      return `
        <li
          class="answer-point-item squircle-md ${isChecked ? 'checked' : ''}"
          data-action="toggle-point"
          data-card-id="${escapeAttr(card.id)}"
          data-point-idx="${pIdx}"
          role="checkbox"
          aria-checked="${isChecked}"
          tabindex="0"
          title="點擊微交互勾選此給分要點"
        >
          <div class="point-left">
            <span class="point-check-circle" aria-hidden="true">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="20 6 9 17 4 12"></polyline>
              </svg>
            </span>
            <span>${highlightText(pt.text, state.searchQuery)}</span>
          </div>
          ${pt.marks != null ? `<span class="point-mark-tag">(${pt.marks})</span>` : ''}
        </li>
      `;
    })
    .join('');

  const animDelay = state.viewMode === 'stream' ? `style="animation-delay: ${Math.min(idx * 45, 320)}ms"` : '';

  return `
    <article
      class="qa-flashcard squircle-2xl ${card.mastered ? 'is-mastered' : ''}"
      data-card-id="${escapeAttr(card.id)}"
      id="flashcard-${escapeAttr(card.id)}"
      ${animDelay}
    >
      <div class="card-meta-bar">
        <div class="card-badges">
          <span class="badge-index squircle-sm">Q${String(card.index).padStart(2, '0')}</span>
          ${marksBadge}
          ${sourceBadge}
        </div>

        <div class="card-quick-actions">
          <button
            class="micro-chip-btn squircle-sm ${card.bookmarked ? 'active-bookmark' : ''}"
            type="button"
            data-action="toggle-bookmark"
            data-card-id="${escapeAttr(card.id)}"
            title="加入重點星標"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="${card.bookmarked ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round">
              <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
            </svg>
            <span>${card.bookmarked ? '已星標' : '星標'}</span>
          </button>

          <button
            class="micro-chip-btn squircle-sm ${card.mastered ? 'active-mastered' : ''}"
            type="button"
            data-action="toggle-mastered"
            data-card-id="${escapeAttr(card.id)}"
            title="標記為已掌握"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">
              <path d="M20 6L9 17l-5-5"></path>
            </svg>
            <span>${card.mastered ? '已掌握' : '標記掌握'}</span>
          </button>
        </div>
      </div>

      <!-- Simultaneous Question & Answer Dual Panes -->
      <div class="qa-split-body">
        <!-- Left: Raised Extruded Question Surface -->
        <section class="qa-pane-question squircle-xl" aria-label="問題">
          <div class="pane-eyebrow">
            <span class="eyebrow-dot"><span class="dot"></span>QUESTION · 問題</span>
            ${card.marks != null ? `<span>(${card.marks} ${card.marks === 1 ? 'mark' : 'marks'})</span>` : ''}
          </div>
          <p class="question-copy">${qHighlighted}</p>
          ${qImagesHtml}
        </section>

        <!-- Right: Soft Inset Answer Well (Displayed Simultaneously) -->
        <section class="qa-pane-answer squircle-xl" aria-label="答案與給分點">
          <div class="pane-eyebrow">
            <span class="eyebrow-dot"><span class="dot"></span>ANSWER · 答案與給分標準</span>
            <span>點擊單項可核對</span>
          </div>
          <ul class="answer-points-list">
            ${pointsHtml}
          </ul>
          ${aImagesHtml}
        </section>
      </div>
    </article>
  `;
}

function attachCardInteractions() {
  // Cursor Specular Light Tracking Micro-interaction on each Flashcard
  const cardEls = els.cardsStage.querySelectorAll('.qa-flashcard');
  cardEls.forEach((cardEl) => {
    cardEl.addEventListener('pointermove', (e) => {
      const rect = cardEl.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      cardEl.style.setProperty('--spot-x', `${x}px`);
      cardEl.style.setProperty('--spot-y', `${y}px`);
    });
  });
}

function bindEvents() {
  // Upload Button & File Input
  els.btnUpload.addEventListener('click', () => els.fileInput.click());
  els.dropzoneDrawer.addEventListener('click', (e) => {
    if (e.target.closest('button')) return;
    els.fileInput.click();
  });

  els.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (file) {
      await handleDocxFile(file);
      els.fileInput.value = '';
    }
  });

  // Toggle Dropzone Drawer
  els.btnToggleDrawer.addEventListener('click', () => {
    els.dropzoneDrawer.classList.toggle('hidden');
  });

  // Theme Toggle
  els.btnThemeToggle.addEventListener('click', () => {
    applyTheme(state.theme === 'light' ? 'dark' : 'light');
  });

  // View Mode Switch
  els.modeStreamBtn.addEventListener('click', () => {
    state.viewMode = 'stream';
    renderAll();
  });

  els.modeFocusBtn.addEventListener('click', () => {
    state.viewMode = 'focus';
    state.focusIndex = 0;
    renderAll();
  });

  // Filter Tabs
  els.filterBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.filter = btn.dataset.filter;
      state.focusIndex = 0;
      renderAll();
    });
  });

  // Search Input
  els.searchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    state.focusIndex = 0;
    renderAll();
  });

  els.btnClearSearch.addEventListener('click', () => {
    state.searchQuery = '';
    els.searchInput.value = '';
    els.searchInput.focus();
    renderAll();
  });

  // Download Real Sample .docx File (so user can test re-uploading or editing in Word!)
  els.btnDownloadSample.addEventListener('click', async () => {
    try {
      const blob = await generateSampleDocxBlob(SAMPLE_BIO_DECK);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'HKDSE_Bio_Ch8_QA_Notes.docx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showToast('已下載範例 HKDSE_Bio_Ch8_QA_Notes.docx');
    } catch (err) {
      showToast(`產生範例失敗：${err.message}`);
    }
  });

  // Reset to Built-in Biology Sample Deck
  els.btnResetSample.addEventListener('click', () => {
    state.title = SAMPLE_BIO_DECK.title;
    state.fileName = SAMPLE_BIO_DECK.fileName;
    state.cards = structuredCloneCards(SAMPLE_BIO_DECK.cards);
    state.checkedPoints.clear();
    state.focusIndex = 0;
    state.filter = 'all';
    state.searchQuery = '';
    els.searchInput.value = '';
    renderAll();
    showToast('已載入預設 HKDSE Biology Ch.8 問答閃卡');
  });

  // Delegated Card Actions (Mastered, Bookmark, Marking Point Checkbox)
  els.cardsStage.addEventListener('click', (e) => {
    const actionEl = e.target.closest('[data-action]');
    if (!actionEl) return;

    const action = actionEl.dataset.action;
    const cardId = actionEl.dataset.cardId;
    const card = state.cards.find((c) => c.id === cardId);
    if (!card) return;

    if (action === 'toggle-mastered') {
      card.mastered = !card.mastered;
      renderAll();
    } else if (action === 'toggle-bookmark') {
      card.bookmarked = !card.bookmarked;
      renderAll();
    } else if (action === 'toggle-point') {
      const pIdx = Number(actionEl.dataset.pointIdx);
      const key = `${cardId}:${pIdx}`;
      if (state.checkedPoints.has(key)) {
        state.checkedPoints.delete(key);
      } else {
        state.checkedPoints.add(key);
      }

      // Auto-mark card mastered when all its points are checked!
      const totalPts = (card.answerPoints && card.answerPoints.length) || 1;
      let allChecked = true;
      for (let i = 0; i < totalPts; i++) {
        if (!state.checkedPoints.has(`${cardId}:${i}`)) {
          allChecked = false;
          break;
        }
      }
      if (allChecked && !card.mastered) {
        card.mastered = true;
        showToast(`Q${String(card.index).padStart(2, '0')} 所有給分點已核對，自動標記掌握！`);
      }
      renderAll();
    }
  });

  // Keyboard support for point checkboxes
  els.cardsStage.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      const actionEl = e.target.closest('[data-action="toggle-point"]');
      if (actionEl) {
        e.preventDefault();
        actionEl.click();
      }
    }
  });

  // Bottom Dock Navigation
  els.dockPrevBtn.addEventListener('click', () => navigateStep(-1));
  els.dockNextBtn.addEventListener('click', () => navigateStep(1));

  els.dockShuffleBtn.addEventListener('click', () => {
    for (let i = state.cards.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [state.cards[i], state.cards[j]] = [state.cards[j], state.cards[i]];
    }
    state.focusIndex = 0;
    renderAll();
    showToast('已隨機重排閃卡順序');
  });

  els.dockCheckAllBtn.addEventListener('click', () => {
    state.checkedPoints.clear();
    state.cards.forEach((c) => {
      c.mastered = false;
    });
    renderAll();
    showToast('已重置所有核對點與掌握進度');
  });

  // Global Arrow Key Navigation
  window.addEventListener('keydown', (e) => {
    if (e.target.matches('input, textarea')) return;
    if (e.key === 'ArrowRight') {
      navigateStep(1);
    } else if (e.key === 'ArrowLeft') {
      navigateStep(-1);
    }
  });

  // Global Drag & Drop for .docx files
  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => {
    if (e.dataTransfer?.types?.includes('Files')) {
      e.preventDefault();
      dragDepth++;
      els.globalDropVeil.classList.add('active');
      els.dropzoneDrawer.classList.add('drag-active');
    }
  });

  window.addEventListener('dragover', (e) => {
    if (e.dataTransfer?.types?.includes('Files')) {
      e.preventDefault();
    }
  });

  window.addEventListener('dragleave', () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) {
      els.globalDropVeil.classList.remove('active');
      els.dropzoneDrawer.classList.remove('drag-active');
    }
  });

  window.addEventListener('drop', async (e) => {
    e.preventDefault();
    dragDepth = 0;
    els.globalDropVeil.classList.remove('active');
    els.dropzoneDrawer.classList.remove('drag-active');
    const file = e.dataTransfer?.files?.[0];
    if (file) {
      await handleDocxFile(file);
    }
  });
}

function navigateStep(delta) {
  const filtered = getFilteredCards();
  if (filtered.length === 0) return;

  if (state.viewMode === 'focus') {
    state.focusIndex = (state.focusIndex + delta + filtered.length) % filtered.length;
    renderAll();
  } else {
    // In stream mode, smooth scroll between cards and update counter
    state.focusIndex = (state.focusIndex + delta + filtered.length) % filtered.length;
    const targetCard = filtered[state.focusIndex];
    els.dockCounter.textContent = `${state.focusIndex + 1} / ${filtered.length}`;
    const cardDom = document.getElementById(`flashcard-${targetCard.id}`);
    if (cardDom) {
      cardDom.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }
}

async function handleDocxFile(file) {
  if (!file.name.toLowerCase().endsWith('.docx')) {
    showToast('請選擇 .docx 格式的 Word 文件');
    return;
  }

  try {
    showToast(`正在解析 ${file.name}...`);
    const parsed = await parseDocxToCards(file, file.name);
    if (!parsed.cards || parsed.cards.length === 0) {
      showToast('未在此 .docx 偵測到雙欄問答表格或 Q&A 段落');
      return;
    }

    state.title = parsed.title;
    state.fileName = parsed.fileName;
    state.cards = parsed.cards;
    state.checkedPoints.clear();
    state.focusIndex = 0;
    state.filter = 'all';
    state.searchQuery = '';
    els.searchInput.value = '';
    renderAll();
    showToast(`成功讀取 ${parsed.cards.length} 張同步問答閃卡！`);
  } catch (err) {
    console.error(err);
    showToast(`讀取 .docx 失敗：${err.message}`);
  }
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeAttr(str) {
  return escapeHtml(str).replace(/'/g, '&#39;');
}

function highlightText(text, query) {
  const safe = escapeHtml(text);
  const q = query.trim();
  if (!q) return safe;
  const escapedRegex = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return safe.replace(new RegExp(`(${escapedRegex})`, 'gi'), '<mark class="hl">$1</mark>');
}

init();

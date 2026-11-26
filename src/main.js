import { GalaxyS9Wallpaper, hexToHsl } from './wallpaper.js';
import {
  parseDocxDocument,
  updateRowFromHtml,
  exportDocToDocxBlob,
  buildDocRow,
} from './docxParser.js';
import {
  SAMPLE_DOC_DATA,
  buildInitialSeededFlashcards,
  toLocalISO,
  addDaysISO,
} from './sampleData.js';
import { GoogleDriveClient, APPS_SCRIPT_TEMPLATE } from './googleDrive.js';

const STORAGE_KEY = 'lumina-spatial-studio-v2';

const DEFAULT_BENTO_ORDER = ['summary', 'kpi', 'chart', 'goal', 'activity'];
const DEFAULT_BENTO_LAYOUT = {
  summary: { span: 4, minHeight: 280 },
  kpi: { span: 4, minHeight: 280 },
  chart: { span: 4, minHeight: 280 },
  goal: { span: 4, minHeight: 260 },
  activity: { span: 8, minHeight: 260 },
};

const DEFAULT_APPEARANCE = {
  theme: 'light', // 'light' (IMG_9401 VisionOS frosted) | 'dark' (Obsidian glass)
  accentHex: '#1F9D8B',
  sqRadius: 28,
  sqCurvature: 1.48,
  wpPreset: 's9-silk',
  wpSpeed: 1.0,
  wpPaused: false,
  docsPaper: 'dark', // 'dark' (IMG_9402 black paper) | 'light'
};

const DEFAULT_CARD_GEOMETRY = {
  qaSplit: 46,
  streamMaxWidth: 100,
  fontScale: 1.0,
};

const driveClient = new GoogleDriveClient();
let wallpaper = null;

function createInitialState() {
  const saved = loadFromLocalStorage();
  if (saved && saved.doc && Array.isArray(saved.doc.rows) && saved.doc.rows.length > 0) {
    return {
      activeNav: 'overview',
      fcMode: saved.fcMode || 'stream', // 'stream' | 'focus' | 'canvas'
      fcDayFilter: 'all',
      fcFocusIndex: 0,
      searchQuery: '',
      layoutEditMode: false,
      currentDay: saved.currentDay || 3,
      studyStartDate: saved.studyStartDate || addDaysISO(toLocalISO(), -2),
      doc: saved.doc,
      flashcardsMap: saved.flashcardsMap || {},
      flashcardOrder: saved.flashcardOrder || Object.keys(saved.flashcardsMap || {}),
      appearance: { ...DEFAULT_APPEARANCE, ...(saved.appearance || {}) },
      cardGeometry: { ...DEFAULT_CARD_GEOMETRY, ...(saved.cardGeometry || {}) },
      bentoOrder: Array.isArray(saved.bentoOrder) ? saved.bentoOrder : [...DEFAULT_BENTO_ORDER],
      bentoLayout: { ...structuredClone(DEFAULT_BENTO_LAYOUT), ...(saved.bentoLayout || {}) },
      cardBoxLayouts: saved.cardBoxLayouts || {}, // { [rowId]: { x, y, w, h } }
      lastLayoutSavedAt: saved.lastLayoutSavedAt || '',
      driveFiles: [],
      driveBrowseAll: false,
    };
  }

  const startISO = addDaysISO(toLocalISO(), -2);
  const seeded = buildInitialSeededFlashcards(startISO);

  return {
    activeNav: 'overview',
    fcMode: 'stream',
    fcDayFilter: 'all',
    fcFocusIndex: 0,
    searchQuery: '',
    layoutEditMode: false,
    currentDay: 3,
    studyStartDate: startISO,
    doc: structuredClone(SAMPLE_DOC_DATA),
    flashcardsMap: seeded,
    flashcardOrder: Object.keys(seeded),
    appearance: { ...DEFAULT_APPEARANCE },
    cardGeometry: { ...DEFAULT_CARD_GEOMETRY },
    bentoOrder: [...DEFAULT_BENTO_ORDER],
    bentoLayout: structuredClone(DEFAULT_BENTO_LAYOUT),
    cardBoxLayouts: {},
    lastLayoutSavedAt: '',
    driveFiles: [],
    driveBrowseAll: false,
  };
}

const state = createInitialState();

// DOM references
const els = {
  canvas: document.getElementById('s9-wallpaper-canvas'),
  root: document.documentElement,
  mainWindow: document.getElementById('main-spatial-window'),
  fileInput: document.getElementById('docx-file-input'),
  globalDropVeil: document.getElementById('global-drop-veil'),
  toastPill: document.getElementById('toast-pill'),
  toastText: document.getElementById('toast-text'),

  // Top Bar
  btnTopDrivePill: document.getElementById('btn-top-drive-pill'),
  topWelcomeLabel: document.getElementById('top-welcome-label'),
  topDriveDot: document.getElementById('top-drive-dot'),
  globalSearchInput: document.getElementById('global-search-input'),
  topDayDisplay: document.getElementById('top-day-display'),
  btnDayMinus: document.getElementById('btn-day-minus'),
  btnDayPlus: document.getElementById('btn-day-plus'),
  btnOpenAppearance: document.getElementById('btn-open-appearance'),

  // Window Header
  mainHeading: document.getElementById('main-heading'),
  headerDocName: document.getElementById('header-doc-name'),
  headerSyncSubtitle: document.getElementById('header-sync-subtitle'),
  btnToggleLayoutEdit: document.getElementById('btn-toggle-layout-edit'),
  layoutEditLabel: document.getElementById('layout-edit-label'),
  btnSaveLayout: document.getElementById('btn-save-layout'),
  btnResetLayout: document.getElementById('btn-reset-layout'),
  btnHeaderUpload: document.getElementById('btn-header-upload'),
  layoutSavedTimestamp: document.getElementById('layout-saved-timestamp'),

  // Left Rail & Bottom Dock
  navButtons: document.querySelectorAll('[data-nav]'),
  railEditLayoutBtn: document.getElementById('rail-edit-layout-btn'),
  railUploadBtn: document.getElementById('rail-upload-btn'),
  railAppearanceBtn: document.getElementById('rail-appearance-btn'),
  dockQuickUploadBtn: document.getElementById('dock-quick-upload-btn'),

  // View Panels
  viewOverview: document.getElementById('view-overview'),
  viewFlashcards: document.getElementById('view-flashcards'),
  viewDocs: document.getElementById('view-docs'),
  viewDrive: document.getElementById('view-drive'),

  // Overview
  overviewBentoGrid: document.getElementById('overview-bento-grid'),

  // Flashcards
  fcModeBtns: document.querySelectorAll('[data-fc-mode]'),
  fcDayFilterGroup: document.getElementById('fc-day-filter-group'),
  btnFcShuffle: document.getElementById('btn-fc-shuffle'),
  btnFcGotoDocs: document.getElementById('btn-fc-goto-docs'),
  sliderQaSplit: document.getElementById('slider-qa-split'),
  qaSplitVal: document.getElementById('qa-split-val'),
  sliderCardWidth: document.getElementById('slider-card-width'),
  cardWidthVal: document.getElementById('card-width-val'),
  sliderFontScale: document.getElementById('slider-font-scale'),
  fontScaleVal: document.getElementById('font-scale-val'),
  flashcardsStage: document.getElementById('flashcards-stage'),

  // Docs Editor
  docsPaperShell: document.getElementById('docs-paper-shell'),
  docsTitleEditable: document.getElementById('docs-title-editable'),
  docsProgressPill: document.getElementById('docs-progress-pill'),
  docsTableContainer: document.getElementById('docs-table-container'),
  btnFmtHighlight: document.getElementById('btn-fmt-highlight'),
  btnDocsPaperToggle: document.getElementById('btn-docs-paper-toggle'),
  btnDocsAddRow: document.getElementById('btn-docs-add-row'),
  btnDocsExport: document.getElementById('btn-docs-export'),
  btnDocsSaveDrive: document.getElementById('btn-docs-save-drive'),
  btnDocsAddAllRemaining: document.getElementById('btn-docs-add-all-remaining'),

  // Drive Hub
  driveConnBadge: document.getElementById('drive-conn-badge'),
  driveModeBtns: document.querySelectorAll('[data-drive-mode]'),
  driveOauthPanel: document.getElementById('drive-oauth-panel'),
  driveScriptPanel: document.getElementById('drive-script-panel'),
  originHintCode: document.getElementById('origin-hint-code'),
  inputGdriveClientId: document.getElementById('input-gdrive-client-id'),
  btnDriveOauthConnect: document.getElementById('btn-drive-oauth-connect'),
  btnDriveDisconnect: document.getElementById('btn-drive-disconnect'),
  inputAppsScriptUrl: document.getElementById('input-apps-script-url'),
  btnDriveScriptConnect: document.getElementById('btn-drive-script-connect'),
  btnCopyAppsScript: document.getElementById('btn-copy-apps-script'),
  appsScriptPreview: document.getElementById('apps-script-preview'),
  btnDrivePushState: document.getElementById('btn-drive-push-state'),
  btnDrivePullState: document.getElementById('btn-drive-pull-state'),
  btnDriveUploadCurrentDocx: document.getElementById('btn-drive-upload-current-docx'),
  btnDriveRefreshFolder: document.getElementById('btn-drive-refresh-folder'),
  btnDriveBrowseAll: document.getElementById('btn-drive-browse-all'),
  driveFolderInfo: document.getElementById('drive-folder-info'),
  driveFileList: document.getElementById('drive-file-list'),

  // Appearance Popover
  appearancePopover: document.getElementById('appearance-popover'),
  btnCloseAppearance: document.getElementById('btn-close-appearance'),
  accentSwatches: document.querySelectorAll('.accent-swatch'),
  customAccentPicker: document.getElementById('custom-accent-picker'),
  accentHexReadout: document.getElementById('accent-hex-readout'),
  sliderSqRadius: document.getElementById('slider-sq-radius'),
  radiusValLabel: document.getElementById('radius-val-label'),
  sliderSqK: document.getElementById('slider-sq-k'),
  curvatureValLabel: document.getElementById('curvature-val-label'),
  squircleReadout: document.getElementById('squircle-readout'),
  wpPresetBtns: document.querySelectorAll('[data-wp-preset]'),
  sliderWpSpeed: document.getElementById('slider-wp-speed'),
  wpSpeedLabel: document.getElementById('wp-speed-label'),
  btnToggleWpMotion: document.getElementById('btn-toggle-wp-motion'),
  themeLightBtn: document.getElementById('theme-light-btn'),
  themeDarkBtn: document.getElementById('theme-dark-btn'),
};

let toastTimer = null;

function init() {
  // Boot Samsung Galaxy S9 Infinity Live Wallpaper
  wallpaper = new GalaxyS9Wallpaper(els.canvas);
  wallpaper.start();

  // Populate Drive setup hints
  els.originHintCode.textContent = window.location.origin;
  els.inputGdriveClientId.value = driveClient.clientId;
  els.inputAppsScriptUrl.value = driveClient.appsScriptUrl;
  els.appsScriptPreview.textContent = APPS_SCRIPT_TEMPLATE;

  applyAppearance();
  applyCardGeometry();
  bindEvents();
  renderAll();
}

/* ==========================================================================
   Persistence (LocalStorage + Optional Google Drive Auto-Sync)
   ========================================================================== */

function loadFromLocalStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function getSerializableWorkspace() {
  return {
    version: 2,
    savedAt: new Date().toISOString(),
    fcMode: state.fcMode,
    currentDay: state.currentDay,
    studyStartDate: state.studyStartDate,
    doc: state.doc,
    flashcardsMap: state.flashcardsMap,
    flashcardOrder: state.flashcardOrder,
    appearance: state.appearance,
    cardGeometry: state.cardGeometry,
    bentoOrder: state.bentoOrder,
    bentoLayout: state.bentoLayout,
    cardBoxLayouts: state.cardBoxLayouts,
    lastLayoutSavedAt: state.lastLayoutSavedAt,
  };
}

function persistLocal() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(getSerializableWorkspace()));
  } catch (e) {
    console.warn('LocalStorage quota warning:', e);
  }
}

/* ==========================================================================
   Appearance, Accent Color, Continuous Curvature & S9 Wallpaper Sync
   ========================================================================== */

function applyAppearance() {
  const a = state.appearance;
  const hsl = hexToHsl(a.accentHex);

  els.root.setAttribute('data-theme', a.theme);
  els.root.style.setProperty('--accent-h', String(hsl.h));
  els.root.style.setProperty('--accent-s', `${Math.max(hsl.s, 28)}%`);
  els.root.style.setProperty('--accent-l', `${Math.min(Math.max(hsl.l, 32), 54)}%`);

  els.root.style.setProperty('--sq-radius', `${a.sqRadius}px`);
  els.root.style.setProperty('--sq-k', String(a.sqCurvature));

  // Sync UI readouts
  els.accentHexReadout.textContent = a.accentHex.toUpperCase();
  els.customAccentPicker.value = a.accentHex;
  els.accentSwatches.forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.accent.toLowerCase() === a.accentHex.toLowerCase());
  });

  els.sliderSqRadius.value = String(a.sqRadius);
  els.radiusValLabel.textContent = `${a.sqRadius}px`;
  els.sliderSqK.value = String(a.sqCurvature);
  els.curvatureValLabel.textContent = Number(a.sqCurvature).toFixed(2);
  els.squircleReadout.textContent = `${a.sqRadius}px · n=${Number(a.sqCurvature).toFixed(2)}`;

  els.wpPresetBtns.forEach((b) => b.classList.toggle('active', b.dataset.wpPreset === a.wpPreset));
  els.sliderWpSpeed.value = String(a.wpSpeed);
  els.wpSpeedLabel.textContent = `${Number(a.wpSpeed).toFixed(1)}×`;
  els.btnToggleWpMotion.textContent = a.wpPaused ? '恢復流動' : '暫停流動';

  els.themeLightBtn.classList.toggle('active', a.theme === 'light');
  els.themeDarkBtn.classList.toggle('active', a.theme === 'dark');

  // Docs Paper mode
  els.docsPaperShell.classList.toggle('paper-light', a.docsPaper === 'light');
  els.btnDocsPaperToggle.textContent = a.docsPaper === 'dark' ? '🌙 深色原檔' : '☀ 淺色原檔';

  // Live Galaxy S9 Wallpaper sync
  if (wallpaper) {
    wallpaper.setAccentHex(a.accentHex);
    wallpaper.setPreset(a.wpPreset);
    wallpaper.setSpeed(a.wpSpeed);
    wallpaper.setPaused(a.wpPaused);
  }
}

function applyCardGeometry() {
  const g = state.cardGeometry;
  els.root.style.setProperty('--card-qa-split', `${g.qaSplit}%`);
  els.root.style.setProperty('--card-stream-max-w', `${g.streamMaxWidth}%`);
  els.root.style.setProperty('--card-font-scale', String(g.fontScale));

  els.sliderQaSplit.value = String(g.qaSplit);
  els.qaSplitVal.textContent = `${g.qaSplit}%`;
  els.sliderCardWidth.value = String(g.streamMaxWidth);
  els.cardWidthVal.textContent = `${g.streamMaxWidth}%`;
  els.sliderFontScale.value = String(g.fontScale);
  els.fontScaleVal.textContent = `${Number(g.fontScale).toFixed(2)}×`;
}

/* ==========================================================================
   Core Study Helpers (Docs Rows <-> Day-Tracked Flashcards)
   ========================================================================== */

function getRowById(rowId) {
  return state.doc.rows.find((r) => r.id === rowId);
}

function getActiveFlashcardItems() {
  // Preserve saved flashcardOrder while keeping valid rows
  const validIds = state.flashcardOrder.filter(
    (id) => state.flashcardsMap[id] && getRowById(id)
  );
  // Add any map keys not yet in flashcardOrder
  for (const id of Object.keys(state.flashcardsMap)) {
    if (!validIds.includes(id) && getRowById(id)) validIds.push(id);
  }
  state.flashcardOrder = validIds;

  return validIds.map((id) => ({
    row: getRowById(id),
    meta: state.flashcardsMap[id],
  }));
}

function addRowToFlashcards(rowId, customDay = state.currentDay) {
  const row = getRowById(rowId);
  if (!row) return;
  if (state.flashcardsMap[rowId]) {
    showToast(`Q${String(row.index).padStart(2, '0')} 已在 Day ${state.flashcardsMap[rowId].addedDay} 加入過閃卡`);
    return;
  }
  state.flashcardsMap[rowId] = {
    rowId,
    addedDay: customDay,
    addedDate: toLocalISO(),
    mastered: false,
    bookmarked: false,
    checkedPoints: [],
  };
  if (!state.flashcardOrder.includes(rowId)) {
    state.flashcardOrder.unshift(rowId);
  }
  persistLocal();
  renderAll();
  showToast(`✓ 已加入 Q${String(row.index).padStart(2, '0')} 至 Flashcard（記錄為 Day ${customDay}）`);
}

function removeRowFromFlashcards(rowId) {
  if (!state.flashcardsMap[rowId]) return;
  delete state.flashcardsMap[rowId];
  state.flashcardOrder = state.flashcardOrder.filter((id) => id !== rowId);
  persistLocal();
  renderAll();
  showToast('已從閃卡清單移出（原檔題目保留）');
}

function getGreetingByHour() {
  const h = new Date().getHours();
  if (h < 12) return 'Good morning';
  if (h < 18) return 'Good afternoon';
  return 'Good evening';
}

/* ==========================================================================
   Master Render Orchestrator
   ========================================================================== */

function renderAll() {
  const fcItems = getActiveFlashcardItems();
  const totalDocRows = state.doc.rows.length;

  // Update Header & Top Utility Bar
  const connected = driveClient.isConnected();
  els.topDriveDot.classList.toggle('connected', connected);
  els.topWelcomeLabel.textContent = connected
    ? `Drive · ${driveClient.userEmail || 'flashcard 已連線'}`
    : 'Welcome, Study Studio';
  els.topDayDisplay.textContent = `Day ${state.currentDay}`;

  const headingByView = {
    overview: `${getGreetingByHour()}, Study Studio`,
    flashcards: '同步問答閃卡 · Flashcards',
    docs: 'Docs 原檔閱讀與編輯器',
    drive: 'Google Drive · flashcard 同步中心',
  };
  els.mainHeading.textContent = headingByView[state.activeNav] || headingByView.overview;
  els.headerDocName.textContent = state.doc.fileName;
  els.headerSyncSubtitle.textContent = `原檔已收錄 ${fcItems.length} / ${totalDocRows} 題 · 目前為 Day ${state.currentDay}`;

  // Layout Edit State UI
  els.mainWindow.classList.toggle('layout-editing', state.layoutEditMode);
  els.btnToggleLayoutEdit.classList.toggle('editing-active', state.layoutEditMode);
  els.railEditLayoutBtn.classList.toggle('accent-active', state.layoutEditMode);
  els.layoutEditLabel.textContent = state.layoutEditMode ? '退出佈局編輯' : '調整卡片大小與位置';
  els.btnSaveLayout.style.display = state.layoutEditMode ? 'inline-flex' : 'none';
  els.btnResetLayout.style.display = state.layoutEditMode ? 'inline-flex' : 'none';
  if (state.lastLayoutSavedAt) {
    els.layoutSavedTimestamp.textContent = `上次保存：${state.lastLayoutSavedAt}`;
  }

  // Nav active state
  els.navButtons.forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.nav === state.activeNav);
  });

  // View panels
  els.viewOverview.classList.toggle('active', state.activeNav === 'overview');
  els.viewFlashcards.classList.toggle('active', state.activeNav === 'flashcards');
  els.viewDocs.classList.toggle('active', state.activeNav === 'docs');
  els.viewDrive.classList.toggle('active', state.activeNav === 'drive');

  // Render active panels
  renderOverviewBento(fcItems);
  renderFlashcardsView(fcItems);
  renderDocsView();
  renderDriveView();
}

/* ==========================================================================
   VIEW 1 RENDERER: VisionOS Bento Overview (IMG_9401 + Draggable/Resizable)
   ========================================================================== */

function renderOverviewBento(fcItems) {
  const totalDocs = state.doc.rows.length;
  const totalFc = fcItems.length;
  const masteredFc = fcItems.filter((x) => x.meta.mastered).length;
  const pendingFc = totalFc - masteredFc;
  const todayFc = fcItems.filter((x) => x.meta.addedDay === state.currentDay).length;
  const totalMarks = fcItems.reduce((s, x) => s + (Number(x.row.marks) || 0), 0);
  const coveragePct = totalDocs > 0 ? Math.round((totalFc / totalDocs) * 100) : 0;

  // Compute Day distribution for Wallets + Bar Chart
  const maxDay = Math.max(state.currentDay, ...fcItems.map((x) => x.meta.addedDay || 1), 5);
  const dayBuckets = [];
  for (let d = 1; d <= Math.min(maxDay + 1, 7); d++) {
    const items = fcItems.filter((x) => x.meta.addedDay === d);
    dayBuckets.push({
      day: d,
      total: items.length,
      mastered: items.filter((x) => x.meta.mastered).length,
    });
  }
  const maxBucket = Math.max(...dayBuckets.map((b) => b.total), 3);

  // Recent Activity Filtered by Global Search
  const q = state.searchQuery.trim().toLowerCase();
  const recentRows = fcItems
    .filter((item) => {
      if (!q) return true;
      return (
        item.row.question.toLowerCase().includes(q) ||
        item.row.answerRaw.toLowerCase().includes(q) ||
        `day ${item.meta.addedDay}`.includes(q)
      );
    })
    .slice(0, 7);

  const tilesMarkup = {
    summary: () => {
      const top3Days = dayBuckets.slice(0, 3);
      return `
        <div class="tile-kicker">
          <span>Flashcards 總覽</span>
          <span class="pill-tag tag-day">Day ${state.currentDay}</span>
        </div>
        <div class="big-metric-row">
          <span class="big-metric-num">${totalFc} Cards</span>
          <span class="trend-chip">+${todayFc} 今日加入</span>
        </div>
        <div class="dual-cta-row">
          <button class="btn-pill solid-ink" data-jump-nav="flashcards" type="button">進入閃卡複習</button>
          <button class="btn-pill white-glass" data-jump-nav="docs" type="button">閱讀 Docs 原檔</button>
        </div>
        <div class="mini-day-wallets">
          ${top3Days
            .map(
              (b) => `
            <div class="mini-wallet-card sq-sm" data-jump-day="${b.day}" title="篩選 Day ${b.day} 閃卡">
              <span>Day ${b.day}</span>
              <strong>${b.total} 題</strong>
              <em>${b.mastered} 已掌握</em>
            </div>
          `
            )
            .join('')}
        </div>
      `;
    },

    kpi: () => `
      <div class="kpi-matrix-grid">
        <div class="kpi-sub-tile featured-accent sq-md" data-jump-nav="flashcards" style="cursor: pointer">
          <div class="kpi-top">
            <span class="kpi-label">已掌握 Mastered</span>
            <span class="kpi-icon-circle">✓</span>
          </div>
          <div>
            <div class="kpi-value">${masteredFc}</div>
            <div class="kpi-sub">${totalFc > 0 ? Math.round((masteredFc / totalFc) * 100) : 0}% 掌握率</div>
          </div>
        </div>

        <div class="kpi-sub-tile sq-md" data-jump-nav="flashcards" style="cursor: pointer">
          <div class="kpi-top">
            <span class="kpi-label">待複習 Pending</span>
            <span class="kpi-icon-circle">⚡</span>
          </div>
          <div>
            <div class="kpi-value">${pendingFc}</div>
            <div class="kpi-sub">同步問答複習中</div>
          </div>
        </div>

        <div class="kpi-sub-tile sq-md" data-jump-nav="docs" style="cursor: pointer">
          <div class="kpi-top">
            <span class="kpi-label">原檔題庫 Docs</span>
            <span class="kpi-icon-circle">📄</span>
          </div>
          <div>
            <div class="kpi-value">${totalFc}/${totalDocs}</div>
            <div class="kpi-sub">${totalDocs - totalFc} 題尚未加入</div>
          </div>
        </div>

        <div class="kpi-sub-tile sq-md">
          <div class="kpi-top">
            <span class="kpi-label">閃卡總分 Marks</span>
            <span class="kpi-icon-circle">★</span>
          </div>
          <div>
            <div class="kpi-value">${totalMarks}</div>
            <div class="kpi-sub">Marking Scheme</div>
          </div>
        </div>
      </div>
    `,

    chart: () => `
      <div class="tile-kicker">
        <span>每日加入與掌握曲線</span>
        <span style="font-size: 0.74rem; color: var(--ink-muted)">點擊長條篩選天數</span>
      </div>
      <div class="day-chart-wrap sq-md">
        <div class="chart-legend">
          <span><i class="legend-dot" style="background: var(--accent)"></i>新增閃卡</span>
          <span><i class="legend-dot" style="background: var(--ink-solid-btn)"></i>已掌握</span>
        </div>
        <div class="chart-bars-row">
          ${dayBuckets
            .map((b) => {
              const unmastered = Math.max(b.total - b.mastered, 0);
              const addedH = b.total === 0 ? 6 : Math.max(Math.round((unmastered / maxBucket) * 86), unmastered > 0 ? 16 : 0);
              const mastH = b.mastered === 0 ? 4 : Math.max(Math.round((b.mastered / maxBucket) * 86), 18);
              return `
                <div class="chart-col" data-jump-day="${b.day}" title="Day ${b.day}: 共 ${b.total} 題 (${b.mastered} 已掌握)">
                  <div class="chart-bar-stack">
                    ${addedH > 0 ? `<div class="bar-seg-added" style="height: ${addedH}px"></div>` : ''}
                    <div class="bar-seg-mastered" style="height: ${mastH}px"></div>
                  </div>
                  <span class="chart-col-label">D${b.day}</span>
                </div>
              `;
            })
            .join('')}
        </div>
      </div>
    `,

    goal: () => `
      <div class="tile-kicker">
        <span>原檔收錄進度</span>
        <strong style="color: var(--accent)">${coveragePct}%</strong>
      </div>
      <div>
        <div class="goal-track">
          <div class="goal-fill" style="width: ${coveragePct}%"></div>
        </div>
        <div style="display: flex; justify-content: space-between; font-size: 0.76rem; color: var(--ink-muted); margin-top: 7px">
          <span>已加入 ${totalFc} 題</span>
          <span>原檔共 ${totalDocs} 題</span>
        </div>
      </div>
      <div class="mini-cards-preview" style="margin-top: auto">
        <div class="mini-visual-card card-dark sq-md" data-jump-nav="docs">
          <span style="font-size: 0.72rem; opacity: 0.75">DOCS VIEWER</span>
          <div>
            <strong style="font-size: 0.9rem; display: block">原檔編輯與勾選</strong>
            <span style="font-size: 0.72rem; opacity: 0.7">${totalDocs} 題雙欄表格</span>
          </div>
        </div>
        <div class="mini-visual-card card-accent sq-md" data-jump-nav="drive">
          <span style="font-size: 0.72rem; opacity: 0.85">GOOGLE DRIVE</span>
          <div>
            <strong style="font-size: 0.9rem; display: block">/flashcard</strong>
            <span style="font-size: 0.72rem; opacity: 0.85">${driveClient.isConnected() ? '已連線同步' : '點擊設定雲端'}</span>
          </div>
        </div>
      </div>
    `,

    activity: () => `
      <div class="tile-kicker">
        <span>已加入的 Flashcards 紀錄 (Day Log)</span>
        <button class="chip-action-btn" data-jump-nav="docs" type="button">＋ 從 Docs 新增</button>
      </div>
      <div class="activity-table-wrap">
        <table class="activity-table">
          <thead>
            <tr>
              <th>題號</th>
              <th>問題摘要</th>
              <th>加入天數</th>
              <th>狀態</th>
              <th style="text-align: right">原檔</th>
            </tr>
          </thead>
          <tbody>
            ${
              recentRows.length === 0
                ? `<tr><td colspan="5" style="text-align: center; color: var(--ink-muted)">尚未加入任何閃卡，前往「Docs 原檔」勾選題目即可加入！</td></tr>`
                : recentRows
                    .map(
                      ({ row, meta }) => `
                  <tr data-open-card="${escapeAttr(row.id)}">
                    <td>Q${String(row.index).padStart(2, '0')}</td>
                    <td style="max-width: 310px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis">
                      ${escapeHtml(row.question)}
                    </td>
                    <td><span class="pill-tag tag-day" style="padding: 2px 9px">Day ${meta.addedDay}</span></td>
                    <td>
                      <span class="status-dot-label ${meta.mastered ? '' : 'pending'}">
                        ${meta.mastered ? '已掌握' : '複習中'}
                      </span>
                    </td>
                    <td>
                      <button class="chip-action-btn" data-locate-doc="${escapeAttr(row.id)}" type="button">查看原檔</button>
                    </td>
                  </tr>
                `
                    )
                    .join('')
            }
          </tbody>
        </table>
      </div>
    `,
  };

  els.overviewBentoGrid.innerHTML = state.bentoOrder
    .map((tileKey) => {
      const cfg = state.bentoLayout[tileKey] || { span: 4, minHeight: 260 };
      const renderFn = tilesMarkup[tileKey];
      if (!renderFn) return '';
      return `
        <article
          class="bento-tile resizable-spatial-item sq-card"
          data-bento-id="${tileKey}"
          draggable="${state.layoutEditMode}"
          style="--span: ${cfg.span}; --tile-h: ${cfg.minHeight}px"
        >
          <div class="spatial-drag-bar sq-sm">
            <span>⋮⋮ 拖曳調換位置 (${tileKey.toUpperCase()})</span>
            <span>欄寬 ${cfg.span}/12</span>
          </div>
          ${renderFn()}
          <span class="spatial-size-readout">span ${cfg.span}/12 · ${cfg.minHeight}px</span>
          <div class="spatial-resize-handle" data-resize-bento="${tileKey}" title="拖曳調整卡片大小">⤡</div>
        </article>
      `;
    })
    .join('');

  attachBentoDragAndResize();
}

/* ==========================================================================
   VIEW 2 RENDERER: Simultaneous Q&A Flashcards (Stream / Focus / Free Canvas)
   ========================================================================== */

function renderFlashcardsView(fcItems) {
  // Update mode buttons
  els.fcModeBtns.forEach((b) => b.classList.toggle('active', b.dataset.fcMode === state.fcMode));

  // Build Day filter pills dynamically
  const distinctDays = [...new Set(fcItems.map((x) => x.meta.addedDay))].sort((a, b) => a - b);
  const dayPills = [
    { key: 'all', label: `全部 (${fcItems.length})` },
    ...distinctDays.map((d) => ({
      key: `day-${d}`,
      label: `Day ${d} (${fcItems.filter((x) => x.meta.addedDay === d).length})`,
    })),
    { key: 'pending', label: '待複習' },
    { key: 'starred', label: '星標' },
  ];

  els.fcDayFilterGroup.innerHTML = dayPills
    .map(
      (p) => `
      <button class="seg-btn ${state.fcDayFilter === p.key ? 'active' : ''}" data-fc-filter="${p.key}" type="button">
        ${p.label}
      </button>
    `
    )
    .join('');

  // Filter items
  const q = state.searchQuery.trim().toLowerCase();
  const filtered = fcItems.filter(({ row, meta }) => {
    if (state.fcDayFilter === 'pending' && meta.mastered) return false;
    if (state.fcDayFilter === 'starred' && !meta.bookmarked) return false;
    if (state.fcDayFilter.startsWith('day-')) {
      const targetDay = Number(state.fcDayFilter.replace('day-', ''));
      if (meta.addedDay !== targetDay) return false;
    }
    if (!q) return true;
    return (
      row.question.toLowerCase().includes(q) ||
      row.answerRaw.toLowerCase().includes(q) ||
      (row.source || '').toLowerCase().includes(q) ||
      `day ${meta.addedDay}`.includes(q)
    );
  });

  els.flashcardsStage.className = `flashcards-stage mode-${state.fcMode}`;

  if (filtered.length === 0) {
    els.flashcardsStage.innerHTML = `
      <div class="bento-tile sq-card" style="width: 100%; text-align: center; padding: 48px 24px">
        <h3 style="margin: 0 0 8px">目前篩選下沒有閃卡</h3>
        <p style="margin: 0 0 18px; color: var(--ink-secondary)">你可以前往「Docs 原檔」檢視題目，看完一題點一下「＋ 加入 Flashcard」就會自動記錄加入天數並出現在這裡！</p>
        <div style="display: flex; justify-content: center; gap: 10px">
          <button class="btn-pill solid-accent" data-jump-nav="docs" type="button">打開 Docs 原檔挑題</button>
          <button class="btn-pill" id="btn-clear-fc-filter" type="button">顯示全部閃卡</button>
        </div>
      </div>
    `;
    document.getElementById('btn-clear-fc-filter')?.addEventListener('click', () => {
      state.fcDayFilter = 'all';
      state.searchQuery = '';
      els.globalSearchInput.value = '';
      renderAll();
    });
    return;
  }

  if (state.fcFocusIndex >= filtered.length) state.fcFocusIndex = 0;

  if (state.fcMode === 'focus') {
    const active = filtered[state.fcFocusIndex];
    els.flashcardsStage.innerHTML = `
      ${buildFlashcardMarkup(active, state.fcFocusIndex)}
      <div class="focus-nav-bar">
        <button id="btn-focus-prev" class="btn-pill" type="button">← 上一張</button>
        <span class="pill-tag tag-id" style="padding: 8px 16px">${state.fcFocusIndex + 1} / ${filtered.length}</span>
        <button id="btn-focus-next" class="btn-pill solid-ink" type="button">下一張 →</button>
      </div>
    `;
    document.getElementById('btn-focus-prev')?.addEventListener('click', () => {
      state.fcFocusIndex = (state.fcFocusIndex - 1 + filtered.length) % filtered.length;
      renderAll();
    });
    document.getElementById('btn-focus-next')?.addEventListener('click', () => {
      state.fcFocusIndex = (state.fcFocusIndex + 1) % filtered.length;
      renderAll();
    });
  } else {
    els.flashcardsStage.innerHTML = filtered
      .map((item, idx) => buildFlashcardMarkup(item, idx))
      .join('');
  }

  attachFlashcardDragAndResize();
}

function buildFlashcardMarkup({ row, meta }, idx) {
  const box = state.cardBoxLayouts[row.id] || {};
  const isCanvas = state.fcMode === 'canvas';

  // Compute default staggered canvas coordinates if none saved yet
  const defaultX = (idx % 2) * 585 + 16;
  const defaultY = Math.floor(idx / 2) * 320 + 16;
  const leftPos = box.x !== undefined ? box.x : defaultX;
  const topPos = box.y !== undefined ? box.y : defaultY;

  const styleParts = [];
  if (isCanvas) {
    styleParts.push(`left: ${leftPos}px`, `top: ${topPos}px`);
    if (box.w) styleParts.push(`--card-w: ${box.w}px`);
    if (box.h) styleParts.push(`--card-h: ${box.h}px`);
  } else {
    if (box.w) styleParts.push(`--card-w: ${box.w}px`);
    if (box.h) styleParts.push(`--card-h: ${box.h}px`);
  }

  const points =
    row.answerPoints && row.answerPoints.length > 0
      ? row.answerPoints
      : [{ text: row.answerRaw || '', marks: null }];

  const checkedSet = new Set(meta.checkedPoints || []);
  const sizeBadge = `${box.w ? `${Math.round(box.w)}px` : 'auto'} × ${box.h ? `${Math.round(box.h)}px` : 'auto'}`;

  return `
    <article
      class="qa-flashcard resizable-spatial-item sq-card ${meta.mastered ? 'is-mastered' : ''}"
      data-fc-id="${escapeAttr(row.id)}"
      draggable="${state.layoutEditMode && !isCanvas}"
      style="${styleParts.join('; ')}"
    >
      <div class="spatial-drag-bar sq-sm" data-canvas-drag="${escapeAttr(row.id)}">
        <span>⋮⋮ ${isCanvas ? '自由拖曳定位 (X/Y)' : '拖曳調整閃卡順序'}</span>
        <span>${isCanvas ? `(${Math.round(leftPos)}, ${Math.round(topPos)})` : '拖曳右下角可縮放'}</span>
      </div>

      <div class="card-meta-row">
        <div class="card-badge-group">
          <span class="pill-tag tag-id">Q${String(row.index).padStart(2, '0')}</span>
          <span class="pill-tag tag-day" title="於 ${meta.addedDate || ''} 加入閃卡">✓ Day ${meta.addedDay} 加入</span>
          ${row.marks != null ? `<span class="pill-tag tag-marks">${row.marks} ${row.marks === 1 ? 'mark' : 'marks'}</span>` : ''}
        </div>

        <div class="card-actions-group">
          <button class="chip-action-btn" data-locate-doc="${escapeAttr(row.id)}" type="button" title="在 Docs 原檔定位並編輯此題">
            📄 查看原檔
          </button>
          <button class="chip-action-btn ${meta.bookmarked ? 'active-star' : ''}" data-fc-star="${escapeAttr(row.id)}" type="button">
            ${meta.bookmarked ? '★ 已星標' : '☆ 星標'}
          </button>
          <button class="chip-action-btn ${meta.mastered ? 'active-master' : ''}" data-fc-master="${escapeAttr(row.id)}" type="button">
            ${meta.mastered ? '✓ 已掌握' : '標記掌握'}
          </button>
          <button class="chip-action-btn" data-fc-remove="${escapeAttr(row.id)}" type="button" title="移出閃卡（不刪除原檔）">
            ✕
          </button>
        </div>
      </div>

      <!-- Simultaneous Side-by-Side Question & Answer -->
      <div class="qa-dual-panes">
        <div class="qa-q-box sq-md">
          <div class="pane-cap">
            <span>QUESTION · 問題</span>
            ${row.marks != null ? `<span>(${row.marks}m)</span>` : ''}
          </div>
          <p class="q-body-text">${highlightText(row.question, state.searchQuery)}</p>
          ${row.source ? `<div class="q-source-foot">*${highlightText(row.source, state.searchQuery)}</div>` : ''}
          ${
            row.questionImages?.length
              ? `<div style="display:flex;gap:8px;flex-wrap:wrap">${row.questionImages
                  .map((src) => `<img src="${src}" style="max-height:180px;border-radius:10px" alt="附圖"/>`)
                  .join('')}</div>`
              : ''
          }
        </div>

        <div class="qa-a-box sq-md">
          <div class="pane-cap">
            <span>ANSWER · 答案與給分點（同步顯示）</span>
            <span>點擊單條核對</span>
          </div>
          <ul class="a-points-ul">
            ${points
              .map((pt, pIdx) => {
                const isChecked = checkedSet.has(pIdx);
                return `
                  <li class="a-point-li sq-sm ${isChecked ? 'checked' : ''}" data-fc-pt="${escapeAttr(row.id)}" data-pt-idx="${pIdx}">
                    <span class="a-pt-left">
                      <span class="a-pt-box">✓</span>
                      <span>${highlightText(pt.text, state.searchQuery)}</span>
                    </span>
                    ${pt.marks != null ? `<span class="a-pt-mark">(${pt.marks})</span>` : ''}
                  </li>
                `;
              })
              .join('')}
          </ul>
          ${
            row.answerImages?.length
              ? `<div style="display:flex;gap:8px;flex-wrap:wrap">${row.answerImages
                  .map((src) => `<img src="${src}" style="max-height:180px;border-radius:10px" alt="答案附圖"/>`)
                  .join('')}</div>`
              : ''
          }
        </div>
      </div>

      <span class="spatial-size-readout">${sizeBadge}</span>
      <div class="spatial-resize-handle" data-resize-fc="${escapeAttr(row.id)}" title="拖曳調整閃卡寬度與高度">⤡</div>
    </article>
  `;
}

/* ==========================================================================
   VIEW 3 RENDERER: Docs Viewer & Rich-Text Editor (Exact IMG_9402 Style!)
   ========================================================================== */

function renderDocsView() {
  // Only overwrite editable title if not actively focused
  if (document.activeElement !== els.docsTitleEditable) {
    els.docsTitleEditable.textContent = state.doc.title;
  }

  const addedCount = state.doc.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length;
  els.docsProgressPill.textContent = `已加入閃卡 ${addedCount} / ${state.doc.rows.length} 題`;
  els.btnDocsAddAllRemaining.textContent = `將剩餘 ${state.doc.rows.length - addedCount} 題全部加入 Day ${state.currentDay}`;
  els.btnDocsAddAllRemaining.style.display = addedCount < state.doc.rows.length ? 'inline-flex' : 'none';

  const q = state.searchQuery.trim().toLowerCase();
  const visibleRows = state.doc.rows.filter((r) => {
    if (!q) return true;
    return (
      r.questionRaw.toLowerCase().includes(q) ||
      r.answerRaw.toLowerCase().includes(q) ||
      (r.source || '').toLowerCase().includes(q)
    );
  });

  els.docsTableContainer.innerHTML = visibleRows
    .map((row) => {
      const addedMeta = state.flashcardsMap[row.id];
      const isAdded = Boolean(addedMeta);

      return `
        <div class="doc-row-block ${isAdded ? 'is-added-row' : ''}" id="doc-row-${escapeAttr(row.id)}" data-doc-row-id="${escapeAttr(row.id)}">
          <div class="doc-cell">
            <div
              class="doc-cell-editable"
              contenteditable="true"
              spellcheck="false"
              data-edit-cell="q"
              data-row-id="${escapeAttr(row.id)}"
            >${row.questionHtml}</div>
            <span class="doc-row-num-tag">
              #${String(row.index).padStart(2, '0')}
              ${isAdded ? `<strong style="color: var(--accent)">● 已收錄 (Day ${addedMeta.addedDay})</strong>` : ''}
            </span>
          </div>

          <div class="doc-cell">
            <div
              class="doc-cell-editable"
              contenteditable="true"
              spellcheck="false"
              data-edit-cell="a"
              data-row-id="${escapeAttr(row.id)}"
            >${row.answerHtml}</div>

            <div class="doc-row-action-bar">
              ${
                isAdded
                  ? `
                    <button
                      class="doc-added-badge"
                      data-doc-jump-fc="${escapeAttr(row.id)}"
                      type="button"
                      title="於 ${addedMeta.addedDate || ''} 加入，點擊前往閃卡檢視"
                    >
                      <span>✓</span>
                      <span>已加入 · Day ${addedMeta.addedDay}</span>
                    </button>
                    <button
                      class="doc-mini-icon-btn"
                      data-doc-unadd="${escapeAttr(row.id)}"
                      type="button"
                      title="取消加入閃卡"
                    >✕</button>
                  `
                  : `
                    <button
                      class="doc-add-btn"
                      data-doc-add="${escapeAttr(row.id)}"
                      type="button"
                      title="看完此題後加入 Flashcard 並記錄為 Day ${state.currentDay}"
                    >
                      <span>＋ 加入 Flashcard (Day ${state.currentDay})</span>
                    </button>
                  `
              }
              <button
                class="doc-mini-icon-btn"
                data-doc-toggle-break="${escapeAttr(row.id)}"
                type="button"
                title="在此列下方插入或移除分頁符號 (Page break)"
              >⏎</button>
            </div>
          </div>
        </div>
        ${row.pageBreakAfter ? `<div class="doc-page-break">Page break</div>` : ''}
      `;
    })
    .join('');
}

/* ==========================================================================
   VIEW 4 RENDERER: Google Drive ("flashcard" Folder) Hub
   ========================================================================== */

function renderDriveView() {
  const connected = driveClient.isConnected();
  els.driveConnBadge.textContent = connected
    ? `✓ 已連線 (${driveClient.mode === 'oauth' ? 'OAuth 2.0' : 'Apps Script'})`
    : '未連線 (本地自動儲存中)';
  els.driveConnBadge.classList.toggle('tag-day', connected);

  els.driveModeBtns.forEach((b) => b.classList.toggle('active', b.dataset.driveMode === driveClient.mode));
  els.driveOauthPanel.style.display = driveClient.mode === 'oauth' ? 'flex' : 'none';
  els.driveScriptPanel.style.display = driveClient.mode === 'apps-script' ? 'flex' : 'none';
  els.btnDriveDisconnect.style.display = connected ? 'inline-flex' : 'none';

  if (!connected) {
    els.driveFileList.innerHTML = `
      <div style="padding: 24px; text-align: center; color: var(--ink-secondary); font-size: 0.84rem">
        請在左側選擇連線方式以自動建立並開啟 Google Drive 的 <strong>flashcard</strong> 資料夾。
      </div>
    `;
    return;
  }

  if (state.driveFiles.length === 0) {
    els.driveFileList.innerHTML = `
      <div style="padding: 22px; text-align: center; color: var(--ink-secondary); font-size: 0.84rem">
        <code>Google Drive / flashcard/</code> 目前尚無 .docx 檔案。<br/>
        點擊左側「📄 上傳目前 .docx 至 flashcard 資料夾」即可立即建立第一份雲端原檔！
      </div>
    `;
    return;
  }

  els.driveFileList.innerHTML = state.driveFiles
    .map(
      (f) => `
      <div class="drive-file-item sq-sm">
        <div style="min-width: 0">
          <strong style="display: block; font-size: 0.84rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis">
            ${escapeHtml(f.name)}
          </strong>
          <span style="font-size: 0.72rem; color: var(--ink-muted)">
            ${f.mimeType?.includes('google-apps.document') ? 'Google Docs' : '.docx Word 原檔'}
            ${f.modifiedTime ? ` · ${new Date(f.modifiedTime).toLocaleDateString()}` : ''}
          </span>
        </div>
        <button
          class="btn-pill solid-accent"
          style="padding: 6px 13px; font-size: 0.76rem; flex-shrink: 0"
          data-drive-open-id="${escapeAttr(f.id)}"
          data-drive-open-name="${escapeAttr(f.name)}"
          data-drive-open-mime="${escapeAttr(f.mimeType || '')}"
          type="button"
        >
          在 Docs 開啟
        </button>
      </div>
    `
    )
    .join('');
}

/* ==========================================================================
   Spatial Drag-and-Drop & Interactive Corner Resize Engine
   ========================================================================== */

function attachBentoDragAndResize() {
  let draggedTileId = null;

  els.overviewBentoGrid.querySelectorAll('.bento-tile').forEach((tile) => {
    tile.addEventListener('dragstart', (e) => {
      if (!state.layoutEditMode) return;
      draggedTileId = tile.dataset.bentoId;
      e.dataTransfer.effectAllowed = 'move';
      tile.style.opacity = '0.55';
    });

    tile.addEventListener('dragend', () => {
      tile.style.opacity = '';
      draggedTileId = null;
    });

    tile.addEventListener('dragover', (e) => {
      if (!state.layoutEditMode || !draggedTileId) return;
      e.preventDefault();
    });

    tile.addEventListener('drop', (e) => {
      if (!state.layoutEditMode || !draggedTileId) return;
      e.preventDefault();
      const targetId = tile.dataset.bentoId;
      if (targetId && targetId !== draggedTileId) {
        const fromIdx = state.bentoOrder.indexOf(draggedTileId);
        const toIdx = state.bentoOrder.indexOf(targetId);
        if (fromIdx > -1 && toIdx > -1) {
          state.bentoOrder.splice(fromIdx, 1);
          state.bentoOrder.splice(toIdx, 0, draggedTileId);
          renderAll();
        }
      }
    });
  });

  // Resize handles for Bento Tiles (drag diagonally to change grid span & minHeight!)
  els.overviewBentoGrid.querySelectorAll('[data-resize-bento]').forEach((handle) => {
    handle.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const tileId = handle.dataset.resizeBento;
      const tileEl = handle.closest('.bento-tile');
      const gridRect = els.overviewBentoGrid.getBoundingClientRect();
      const colWidth = gridRect.width / 12;
      const startX = e.clientX;
      const startY = e.clientY;
      const startCfg = { ...(state.bentoLayout[tileId] || { span: 4, minHeight: 260 }) };
      const startRect = tileEl.getBoundingClientRect();
      tileEl.classList.add('is-resizing');

      const onMove = (ev) => {
        const nextWidth = startRect.width + (ev.clientX - startX);
        const nextSpan = Math.max(3, Math.min(12, Math.round(nextWidth / colWidth)));
        const nextH = Math.max(190, Math.min(680, Math.round(startCfg.minHeight + (ev.clientY - startY))));
        state.bentoLayout[tileId] = { span: nextSpan, minHeight: nextH };
        tileEl.style.setProperty('--span', String(nextSpan));
        tileEl.style.setProperty('--tile-h', `${nextH}px`);
        const readout = tileEl.querySelector('.spatial-size-readout');
        if (readout) readout.textContent = `span ${nextSpan}/12 · ${nextH}px`;
      };

      const onUp = () => {
        tileEl.classList.remove('is-resizing');
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
      };

      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
    });
  });
}

function attachFlashcardDragAndResize() {
  const isCanvas = state.fcMode === 'canvas';
  let draggedCardId = null;

  // 1. Stream Reorder Drag-and-Drop
  if (!isCanvas) {
    els.flashcardsStage.querySelectorAll('.qa-flashcard').forEach((cardEl) => {
      cardEl.addEventListener('dragstart', (e) => {
        if (!state.layoutEditMode) return;
        draggedCardId = cardEl.dataset.fcId;
        e.dataTransfer.effectAllowed = 'move';
        cardEl.style.opacity = '0.55';
      });
      cardEl.addEventListener('dragend', () => {
        cardEl.style.opacity = '';
        draggedCardId = null;
      });
      cardEl.addEventListener('dragover', (e) => {
        if (!state.layoutEditMode || !draggedCardId) return;
        e.preventDefault();
      });
      cardEl.addEventListener('drop', (e) => {
        if (!state.layoutEditMode || !draggedCardId) return;
        e.preventDefault();
        const targetId = cardEl.dataset.fcId;
        if (targetId && targetId !== draggedCardId) {
          const fromIdx = state.flashcardOrder.indexOf(draggedCardId);
          const toIdx = state.flashcardOrder.indexOf(targetId);
          if (fromIdx > -1 && toIdx > -1) {
            state.flashcardOrder.splice(fromIdx, 1);
            state.flashcardOrder.splice(toIdx, 0, draggedCardId);
            renderAll();
          }
        }
      });
    });
  } else {
    // 2. Freeform Spatial Canvas X/Y Positioning Drag
    els.flashcardsStage.querySelectorAll('[data-canvas-drag]').forEach((bar) => {
      bar.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        const fcId = bar.dataset.canvasDrag;
        const cardEl = bar.closest('.qa-flashcard');
        const startX = e.clientX;
        const startY = e.clientY;
        const initLeft = parseFloat(cardEl.style.left) || 16;
        const initTop = parseFloat(cardEl.style.top) || 16;
        cardEl.classList.add('is-dragging');

        const onMove = (ev) => {
          const nx = Math.max(8, Math.round(initLeft + (ev.clientX - startX)));
          const ny = Math.max(8, Math.round(initTop + (ev.clientY - startY)));
          cardEl.style.left = `${nx}px`;
          cardEl.style.top = `${ny}px`;
          state.cardBoxLayouts[fcId] = {
            ...(state.cardBoxLayouts[fcId] || {}),
            x: nx,
            y: ny,
          };
        };
        const onUp = () => {
          cardEl.classList.remove('is-dragging');
          window.removeEventListener('pointermove', onMove);
          window.removeEventListener('pointerup', onUp);
        };
        window.addEventListener('pointermove', onMove);
        window.addEventListener('pointerup', onUp);
      });
    });
  }

  // 3. Individual Flashcard Corner Resize Handle (Works in Stream, Focus, AND Canvas!)
  els.flashcardsStage.querySelectorAll('[data-resize-fc]').forEach((handle) => {
    handle.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const fcId = handle.dataset.resizeFc;
      const cardEl = handle.closest('.qa-flashcard');
      const rect = cardEl.getBoundingClientRect();
      const startX = e.clientX;
      const startY = e.clientY;
      cardEl.classList.add('is-resizing');

      const onMove = (ev) => {
        const nw = Math.max(360, Math.min(1260, Math.round(rect.width + (ev.clientX - startX))));
        const nh = Math.max(200, Math.min(760, Math.round(rect.height + (ev.clientY - startY))));
        cardEl.style.setProperty('--card-w', `${nw}px`);
        cardEl.style.setProperty('--card-h', `${nh}px`);
        state.cardBoxLayouts[fcId] = {
          ...(state.cardBoxLayouts[fcId] || {}),
          w: nw,
          h: nh,
        };
        const readout = cardEl.querySelector('.spatial-size-readout');
        if (readout) readout.textContent = `${nw}px × ${nh}px`;
      };

      const onUp = () => {
        cardEl.classList.remove('is-resizing');
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
      };

      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
    });
  });
}

/* ==========================================================================
   Event Wiring
   ========================================================================== */

function bindEvents() {
  // Navigation (Left Rail + Bottom Dock)
  els.navButtons.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.activeNav = btn.dataset.nav;
      renderAll();
    });
  });

  // Top Drive status pill jumps to Drive tab
  els.btnTopDrivePill.addEventListener('click', () => {
    state.activeNav = 'drive';
    renderAll();
  });

  // Global Search
  els.globalSearchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    renderAll();
  });

  // Day Stepper (- / +)
  els.btnDayMinus.addEventListener('click', () => {
    state.currentDay = Math.max(1, state.currentDay - 1);
    persistLocal();
    renderAll();
    showToast(`目前學習天數切換為 Day ${state.currentDay}`);
  });
  els.btnDayPlus.addEventListener('click', () => {
    state.currentDay += 1;
    persistLocal();
    renderAll();
    showToast(`目前學習天數切換為 Day ${state.currentDay}`);
  });

  // Appearance Popover Toggle
  const toggleAppearance = () => els.appearancePopover.classList.toggle('open');
  els.btnOpenAppearance.addEventListener('click', toggleAppearance);
  els.railAppearanceBtn.addEventListener('click', toggleAppearance);
  els.btnCloseAppearance.addEventListener('click', () => els.appearancePopover.classList.remove('open'));

  // Accent Swatches + Custom Wheel
  els.accentSwatches.forEach((sw) => {
    sw.addEventListener('click', () => {
      state.appearance.accentHex = sw.dataset.accent;
      applyAppearance();
      persistLocal();
    });
  });
  els.customAccentPicker.addEventListener('input', (e) => {
    state.appearance.accentHex = e.target.value;
    applyAppearance();
    persistLocal();
  });

  // Continuous Curvature Squircle Sliders
  els.sliderSqRadius.addEventListener('input', (e) => {
    state.appearance.sqRadius = Number(e.target.value);
    applyAppearance();
    persistLocal();
  });
  els.sliderSqK.addEventListener('input', (e) => {
    state.appearance.sqCurvature = Number(e.target.value);
    applyAppearance();
    persistLocal();
  });

  // Samsung Galaxy S9 Wallpaper Controls
  els.wpPresetBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.appearance.wpPreset = btn.dataset.wpPreset;
      applyAppearance();
      persistLocal();
    });
  });
  els.sliderWpSpeed.addEventListener('input', (e) => {
    state.appearance.wpSpeed = Number(e.target.value);
    applyAppearance();
    persistLocal();
  });
  els.btnToggleWpMotion.addEventListener('click', () => {
    state.appearance.wpPaused = !state.appearance.wpPaused;
    applyAppearance();
    persistLocal();
  });

  // Theme Toggle (VisionOS Light Glass / Obsidian Dark Glass)
  els.themeLightBtn.addEventListener('click', () => {
    state.appearance.theme = 'light';
    applyAppearance();
    persistLocal();
  });
  els.themeDarkBtn.addEventListener('click', () => {
    state.appearance.theme = 'dark';
    applyAppearance();
    persistLocal();
  });

  // Layout Edit Mode + Save Layout + Reset Layout
  const toggleEditLayout = () => {
    state.layoutEditMode = !state.layoutEditMode;
    renderAll();
    if (state.layoutEditMode) {
      showToast('已開啟卡片大小與位置調整模式，完成後記得點「保存佈局」');
    }
  };
  els.btnToggleLayoutEdit.addEventListener('click', toggleEditLayout);
  els.railEditLayoutBtn.addEventListener('click', toggleEditLayout);

  els.btnSaveLayout.addEventListener('click', async () => {
    state.lastLayoutSavedAt = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    state.layoutEditMode = false;
    persistLocal();
    renderAll();
    showToast('✓ 已保存所有卡片大小、位置與空間佈局！');
    if (driveClient.isConnected()) {
      try {
        await driveClient.saveWorkspaceState(getSerializableWorkspace());
        showToast('✓ 佈局已同步至 Google Drive / flashcard！');
      } catch (err) {
        showToast(`本地已保存（Drive 同步失敗：${err.message}）`);
      }
    }
  });

  els.btnResetLayout.addEventListener('click', () => {
    state.bentoOrder = [...DEFAULT_BENTO_ORDER];
    state.bentoLayout = structuredClone(DEFAULT_BENTO_LAYOUT);
    state.cardBoxLayouts = {};
    state.cardGeometry = { ...DEFAULT_CARD_GEOMETRY };
    applyCardGeometry();
    persistLocal();
    renderAll();
    showToast('已重置所有卡片尺寸與預設位置');
  });

  // Global Card Geometry Sliders
  els.sliderQaSplit.addEventListener('input', (e) => {
    state.cardGeometry.qaSplit = Number(e.target.value);
    applyCardGeometry();
  });
  els.sliderCardWidth.addEventListener('input', (e) => {
    state.cardGeometry.streamMaxWidth = Number(e.target.value);
    applyCardGeometry();
  });
  els.sliderFontScale.addEventListener('input', (e) => {
    state.cardGeometry.fontScale = Number(e.target.value);
    applyCardGeometry();
  });

  // Upload .docx Triggers
  const openUpload = () => els.fileInput.click();
  els.btnHeaderUpload.addEventListener('click', openUpload);
  els.railUploadBtn.addEventListener('click', openUpload);
  els.dockQuickUploadBtn.addEventListener('click', openUpload);

  els.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (file) {
      await handleLocalDocxUpload(file);
      els.fileInput.value = '';
    }
  });

  // Global Drag & Drop .docx
  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => {
    if (e.dataTransfer?.types?.includes('Files')) {
      e.preventDefault();
      dragDepth++;
      els.globalDropVeil.classList.add('active');
    }
  });
  window.addEventListener('dragover', (e) => {
    if (e.dataTransfer?.types?.includes('Files')) e.preventDefault();
  });
  window.addEventListener('dragleave', () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) els.globalDropVeil.classList.remove('active');
  });
  window.addEventListener('drop', async (e) => {
    if (!e.dataTransfer?.files?.length) return;
    e.preventDefault();
    dragDepth = 0;
    els.globalDropVeil.classList.remove('active');
    await handleLocalDocxUpload(e.dataTransfer.files[0]);
  });

  // Flashcards View Mode & Day Filter
  els.fcModeBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.fcMode = btn.dataset.fcMode;
      if (state.fcMode === 'canvas' && !state.layoutEditMode) {
        state.layoutEditMode = true;
        showToast('自由畫布模式：已自動開啟卡片定位與縮放控點');
      }
      persistLocal();
      renderAll();
    });
  });

  els.fcDayFilterGroup.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-fc-filter]');
    if (!btn) return;
    state.fcDayFilter = btn.dataset.fcFilter;
    state.fcFocusIndex = 0;
    renderAll();
  });

  els.btnFcShuffle.addEventListener('click', () => {
    for (let i = state.flashcardOrder.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [state.flashcardOrder[i], state.flashcardOrder[j]] = [
        state.flashcardOrder[j],
        state.flashcardOrder[i],
      ];
    }
    persistLocal();
    renderAll();
    showToast('已隨機重排閃卡順序');
  });

  els.btnFcGotoDocs.addEventListener('click', () => {
    state.activeNav = 'docs';
    renderAll();
  });

  // Delegated Clicks inside Main Spatial Window (Overview jumps, Flashcards actions, Docs actions)
  els.mainWindow.addEventListener('click', (e) => {
    // 1. Jump to nav view
    const jumpNav = e.target.closest('[data-jump-nav]');
    if (jumpNav) {
      state.activeNav = jumpNav.dataset.jumpNav;
      renderAll();
      return;
    }

    // 2. Jump to Day filter in Flashcards
    const jumpDay = e.target.closest('[data-jump-day]');
    if (jumpDay) {
      state.activeNav = 'flashcards';
      state.fcDayFilter = `day-${jumpDay.dataset.jumpDay}`;
      renderAll();
      return;
    }

    // 3. Locate row in Docs原檔 from Overview or Flashcards!
    const locateDoc = e.target.closest('[data-locate-doc]');
    if (locateDoc) {
      e.stopPropagation();
      const rowId = locateDoc.dataset.locateDoc;
      state.activeNav = 'docs';
      renderAll();
      requestAnimationFrame(() => {
        const el = document.getElementById(`doc-row-${rowId}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          el.classList.remove('pulse-locate');
          void el.offsetWidth;
          el.classList.add('pulse-locate');
        }
      });
      return;
    }

    // 4. Open card from Overview Recent Activity table row
    const openCard = e.target.closest('[data-open-card]');
    if (openCard) {
      const rowId = openCard.dataset.openCard;
      state.activeNav = 'flashcards';
      state.fcDayFilter = 'all';
      renderAll();
      requestAnimationFrame(() => {
        const el = els.flashcardsStage.querySelector(`[data-fc-id="${rowId}"]`);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      });
      return;
    }

    // 5. Flashcard Star / Master / Remove / Marking Point Check
    const starBtn = e.target.closest('[data-fc-star]');
    if (starBtn) {
      const id = starBtn.dataset.fcStar;
      if (state.flashcardsMap[id]) {
        state.flashcardsMap[id].bookmarked = !state.flashcardsMap[id].bookmarked;
        persistLocal();
        renderAll();
      }
      return;
    }

    const masterBtn = e.target.closest('[data-fc-master]');
    if (masterBtn) {
      const id = masterBtn.dataset.fcMaster;
      if (state.flashcardsMap[id]) {
        state.flashcardsMap[id].mastered = !state.flashcardsMap[id].mastered;
        persistLocal();
        renderAll();
      }
      return;
    }

    const removeBtn = e.target.closest('[data-fc-remove]');
    if (removeBtn) {
      removeRowFromFlashcards(removeBtn.dataset.fcRemove);
      return;
    }

    const ptEl = e.target.closest('[data-fc-pt]');
    if (ptEl) {
      const id = ptEl.dataset.fcPt;
      const pIdx = Number(ptEl.dataset.ptIdx);
      const meta = state.flashcardsMap[id];
      const row = getRowById(id);
      if (meta && row) {
        const set = new Set(meta.checkedPoints || []);
        if (set.has(pIdx)) set.delete(pIdx);
        else set.add(pIdx);
        meta.checkedPoints = [...set];

        const totalPts = Math.max(row.answerPoints?.length || 1, 1);
        if (meta.checkedPoints.length >= totalPts && !meta.mastered) {
          meta.mastered = true;
          showToast(`✓ Q${String(row.index).padStart(2, '0')} 所有給分點已核對，自動標記已掌握！`);
        }
        persistLocal();
        renderAll();
      }
      return;
    }

    // 6. Docs View Actions: Add to Flashcard, Jump to Flashcard, Remove, Toggle Page Break
    const docAddBtn = e.target.closest('[data-doc-add]');
    if (docAddBtn) {
      addRowToFlashcards(docAddBtn.dataset.docAdd, state.currentDay);
      return;
    }

    const docUnaddBtn = e.target.closest('[data-doc-unadd]');
    if (docUnaddBtn) {
      removeRowFromFlashcards(docUnaddBtn.dataset.docUnadd);
      return;
    }

    const docJumpFc = e.target.closest('[data-doc-jump-fc]');
    if (docJumpFc) {
      const id = docJumpFc.dataset.docJumpFc;
      state.activeNav = 'flashcards';
      state.fcDayFilter = 'all';
      renderAll();
      requestAnimationFrame(() => {
        const el = els.flashcardsStage.querySelector(`[data-fc-id="${id}"]`);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      });
      return;
    }

    const toggleBreak = e.target.closest('[data-doc-toggle-break]');
    if (toggleBreak) {
      const row = getRowById(toggleBreak.dataset.docToggleBreak);
      if (row) {
        row.pageBreakAfter = !row.pageBreakAfter;
        persistLocal();
        renderDocsView();
      }
      return;
    }
  });

  // Docs Rich-Text Formatting Toolbar
  document.querySelectorAll('.fmt-btn[data-cmd]').forEach((btn) => {
    btn.addEventListener('mousedown', (e) => e.preventDefault()); // Keep focus in contenteditable!
    btn.addEventListener('click', () => {
      document.execCommand(btn.dataset.cmd, false, null);
    });
  });

  els.btnFmtHighlight.addEventListener('mousedown', (e) => e.preventDefault());
  els.btnFmtHighlight.addEventListener('click', () => {
    document.execCommand('hiliteColor', false, 'rgba(250, 204, 21, 0.38)');
  });

  els.btnDocsPaperToggle.addEventListener('click', () => {
    state.appearance.docsPaper = state.appearance.docsPaper === 'dark' ? 'light' : 'dark';
    applyAppearance();
    persistLocal();
  });

  // Live Docs Contenteditable Sync (Edits immediately update row & linked Flashcard!)
  els.docsTitleEditable.addEventListener('blur', () => {
    state.doc.title = els.docsTitleEditable.textContent.trim() || state.doc.title;
    persistLocal();
  });

  els.docsTableContainer.addEventListener(
    'blur',
    (e) => {
      const cell = e.target.closest('[data-edit-cell]');
      if (!cell) return;
      const rowId = cell.dataset.rowId;
      const rowIdx = state.doc.rows.findIndex((r) => r.id === rowId);
      if (rowIdx === -1) return;

      const rowBlock = cell.closest('.doc-row-block');
      const qHtml = rowBlock.querySelector('[data-edit-cell="q"]').innerHTML;
      const aHtml = rowBlock.querySelector('[data-edit-cell="a"]').innerHTML;
      state.doc.rows[rowIdx] = updateRowFromHtml(state.doc.rows[rowIdx], qHtml, aHtml);
      persistLocal();
    },
    true
  );

  els.btnDocsAddRow.addEventListener('click', () => {
    const nextIdx = state.doc.rows.length + 1;
    const newRow = buildDocRow({
      id: `row-custom-${Date.now().toString(36)}`,
      index: nextIdx,
      section: state.doc.title,
      questionLines: ['請在此輸入新題目... (2 marks)'],
      answerLines: ['請在此輸入答案要點一。 (1)', '請在此輸入答案要點二。 (1)'],
      pageBreakAfter: false,
    });
    state.doc.rows.push(newRow);
    persistLocal();
    renderAll();
    requestAnimationFrame(() => {
      const el = document.getElementById(`doc-row-${newRow.id}`);
      el?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
    showToast(`已新增第 #${nextIdx} 列問答題`);
  });

  els.btnDocsAddAllRemaining.addEventListener('click', () => {
    let added = 0;
    for (const r of state.doc.rows) {
      if (!state.flashcardsMap[r.id]) {
        state.flashcardsMap[r.id] = {
          rowId: r.id,
          addedDay: state.currentDay,
          addedDate: toLocalISO(),
          mastered: false,
          bookmarked: false,
          checkedPoints: [],
        };
        state.flashcardOrder.push(r.id);
        added++;
      }
    }
    persistLocal();
    renderAll();
    showToast(`✓ 已批次將 ${added} 題加入 Flashcard（Day ${state.currentDay}）`);
  });

  els.btnDocsExport.addEventListener('click', async () => {
    try {
      const blob = await exportDocToDocxBlob(state.doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = state.doc.fileName.endsWith('.docx')
        ? state.doc.fileName
        : `${state.doc.fileName}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showToast(`已匯出更新後的原檔：${a.download}`);
    } catch (err) {
      showToast(`匯出失敗：${err.message}`);
    }
  });

  const saveDocxToDriveAction = async () => {
    if (!driveClient.isConnected()) {
      state.activeNav = 'drive';
      renderAll();
      showToast('請先連接 Google Drive，系統將自動建立 flashcard 資料夾並保存原檔');
      return;
    }
    try {
      showToast('正在打包 .docx 並上傳至 Google Drive / flashcard...');
      const blob = await exportDocToDocxBlob(state.doc);
      const saved = await driveClient.saveDocxToDrive(
        state.doc.fileName,
        blob,
        state.doc.driveFileId
      );
      state.doc.driveFileId = saved.id || state.doc.driveFileId;
      await driveClient.saveWorkspaceState(getSerializableWorkspace());
      persistLocal();
      await refreshDriveFileList(false);
      showToast(`✓ 已將 ${state.doc.fileName} 與閃卡進度保存至 Google Drive / flashcard！`);
    } catch (err) {
      showToast(`Drive 保存失敗：${err.message}`);
    }
  };

  els.btnDocsSaveDrive.addEventListener('click', saveDocxToDriveAction);
  els.btnDriveUploadCurrentDocx.addEventListener('click', saveDocxToDriveAction);

  // Google Drive Hub Controls
  els.driveModeBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      driveClient.setConfig({ mode: btn.dataset.driveMode });
      renderDriveView();
    });
  });

  els.btnDriveOauthConnect.addEventListener('click', async () => {
    try {
      driveClient.setConfig({ mode: 'oauth', clientId: els.inputGdriveClientId.value });
      showToast('正在開啟 Google 登入視窗並建立 flashcard 資料夾...');
      await driveClient.signInWithOAuth(true);
      await refreshDriveFileList(false);
      renderAll();
      showToast('✓ 已連接 Google Drive 並就緒 flashcard 資料夾！');
    } catch (err) {
      showToast(err.message);
    }
  });

  els.btnDriveScriptConnect.addEventListener('click', async () => {
    try {
      driveClient.setConfig({
        mode: 'apps-script',
        appsScriptUrl: els.inputAppsScriptUrl.value,
      });
      showToast('正在連接 Apps Script 並確認 flashcard 資料夾...');
      await driveClient.ensureFlashcardFolder();
      await refreshDriveFileList(false);
      renderAll();
      showToast('✓ Apps Script 連線成功，flashcard 資料夾已就緒！');
    } catch (err) {
      showToast(err.message);
    }
  });

  els.btnCopyAppsScript.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(APPS_SCRIPT_TEMPLATE);
      showToast('已複製 Google Apps Script 程式碼！');
    } catch {
      showToast('請手動全選下方程式碼框複製');
    }
  });

  els.btnDriveDisconnect.addEventListener('click', () => {
    driveClient.disconnect();
    state.driveFiles = [];
    renderAll();
    showToast('已中斷 Google Drive 連線');
  });

  els.btnDrivePushState.addEventListener('click', async () => {
    if (!driveClient.isConnected()) {
      showToast('請先連接 Google Drive');
      return;
    }
    try {
      showToast('正在同步閃卡、Day 紀錄與卡片佈局至 Drive / flashcard...');
      await driveClient.saveWorkspaceState(getSerializableWorkspace());
      showToast('✓ 雲端同步完成！');
    } catch (err) {
      showToast(err.message);
    }
  });

  els.btnDrivePullState.addEventListener('click', async () => {
    if (!driveClient.isConnected()) {
      showToast('請先連接 Google Drive');
      return;
    }
    try {
      showToast('正在從 Google Drive / flashcard 讀取進度...');
      const cloud = await driveClient.loadWorkspaceState();
      if (!cloud) {
        showToast('flashcard 資料夾尚無雲端狀態檔，可先按「立即同步」建立');
        return;
      }
      Object.assign(state, {
        doc: cloud.doc || state.doc,
        flashcardsMap: cloud.flashcardsMap || state.flashcardsMap,
        flashcardOrder: cloud.flashcardOrder || state.flashcardOrder,
        currentDay: cloud.currentDay || state.currentDay,
        appearance: { ...DEFAULT_APPEARANCE, ...(cloud.appearance || {}) },
        cardGeometry: { ...DEFAULT_CARD_GEOMETRY, ...(cloud.cardGeometry || {}) },
        bentoOrder: cloud.bentoOrder || state.bentoOrder,
        bentoLayout: cloud.bentoLayout || state.bentoLayout,
        cardBoxLayouts: cloud.cardBoxLayouts || state.cardBoxLayouts,
      });
      applyAppearance();
      applyCardGeometry();
      persistLocal();
      renderAll();
      showToast('✓ 已從 Google Drive 載入所有閃卡、原檔與卡片佈局！');
    } catch (err) {
      showToast(err.message);
    }
  });

  els.btnDriveRefreshFolder.addEventListener('click', () => refreshDriveFileList(false));
  els.btnDriveBrowseAll.addEventListener('click', () => refreshDriveFileList(true));

  els.driveFileList.addEventListener('click', async (e) => {
    const btn = e.target.closest('[data-drive-open-id]');
    if (!btn) return;
    try {
      const fileId = btn.dataset.driveOpenId;
      const name = btn.dataset.driveOpenName;
      const mime = btn.dataset.driveOpenMime;
      showToast(`正在從 Google Drive 讀取 ${name}...`);
      const buf = await driveClient.downloadDocxArrayBuffer(fileId, mime);
      const parsed = await parseDocxDocument(buf, name);
      if (!parsed.rows.length) {
        showToast('此檔案中未偵測到雙欄問答表格或 Q&A 內容');
        return;
      }
      state.doc = { ...parsed, driveFileId: fileId };
      state.activeNav = 'docs';
      persistLocal();
      renderAll();
      showToast(`✓ 已從 Google Drive 載入 ${parsed.rows.length} 題至 Docs 原檔檢視器！`);
    } catch (err) {
      showToast(err.message);
    }
  });
}

async function refreshDriveFileList(allDrive = false) {
  if (!driveClient.isConnected()) return;
  try {
    state.driveBrowseAll = allDrive;
    els.driveFolderInfo.innerHTML = allDrive
      ? '顯示整個 Google Drive 的 <code>.docx</code> 與 Google Docs（開啟後保存可一鍵存入 <code>flashcard/</code>）'
      : '顯示 <code>Google Drive / flashcard/</code> 資料夾內的文件';
    state.driveFiles = await driveClient.listDocuments(allDrive);
    renderDriveView();
  } catch (err) {
    showToast(err.message);
  }
}

async function handleLocalDocxUpload(file) {
  if (!file.name.toLowerCase().endsWith('.docx')) {
    showToast('請上傳 .docx 格式的 Word 筆記檔案');
    return;
  }
  try {
    showToast(`正在解析 ${file.name}...`);
    const parsed = await parseDocxDocument(file, file.name);
    if (!parsed.rows.length) {
      showToast('未在此 .docx 偵測到雙欄問答表格或 Q&A 段落');
      return;
    }
    state.doc = { ...parsed, driveFileId: null };
    state.activeNav = 'docs';
    persistLocal();
    renderAll();
    showToast(`✓ 成功載入 ${parsed.rows.length} 題！可在 Docs 原檔勾選並加入 Day ${state.currentDay} 閃卡`);
  } catch (err) {
    console.error(err);
    showToast(`讀取失敗：${err.message}`);
  }
}

function showToast(message) {
  els.toastText.textContent = message;
  els.toastPill.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toastPill.classList.remove('show'), 3200);
}

function escapeHtml(str) {
  return String(str || '')
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
  const q = String(query || '').trim();
  if (!q) return safe;
  const escaped = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return safe.replace(new RegExp(`(${escaped})`, 'gi'), '<mark class="hl">$1</mark>');
}

init();

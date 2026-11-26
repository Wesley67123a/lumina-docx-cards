import { GalaxyS9Wallpaper } from './wallpaper.js';
import {
  parseDocxDocument,
  updateRowFromHtml,
  buildUpdatedDocxBlob,
  buildDocRow,
  groupRowsByChapter,
} from './docxParser.js';
import {
  SAMPLE_DOC_DATA,
  buildInitialSeededFlashcards,
  toLocalISO,
  addDaysISO,
} from './sampleData.js';
import { GoogleDriveClient, APPS_SCRIPT_TEMPLATE } from './googleDrive.js';

const STORAGE_KEY = 'wesley-study-studio-v3';

const PALETTE_ACCENTS = {
  terracotta: '#D96F52',
  sage: '#528467',
  ocean: '#4C7EB0',
  lavender: '#7B61B8',
  amber: '#C98F32',
};

const BREAKDOWN_COLORS = ['#D96F52', '#D4A138', '#5C7F67', '#557CA5', '#7B61B8', '#0EA5E9'];

const I18N = {
  zh: {
    navHome: 'Home 首頁',
    navFlashcards: 'Flashcards 閃卡',
    navDocs: 'Docs 原檔',
    navDrive: 'Google Drive',
    navSettings: '外觀與設定',
    promoTitle: '今日隨機抽卡',
    promoDesc: '從已收錄閃卡隨機組一組快篩複習',
    promoCta: '🎲 立即隨機抽卡',
    s9Btn: 'S9 壁紙',
    openDocx: '開啟 .docx',
    heroResume: '繼續複習閃卡',
    heroRandom: '隨機抽題',
    tileFcLabel: '已收錄閃卡',
    tileStudyLabel: "Today's Study",
    tileRandomLabel: '隨機 Flashcard 組',
    tileRandomVal: '一鍵抽組',
    tileDocsLabel: 'Docs 原檔章節',
    chartTitle: '學習紀錄分佈',
    breakdownTitle: '各 Chapter 收錄與掌握度',
    randDeckTitle: '🎲 選 Flashcard · 隨機閃卡組產生器',
    randDeckSub: '隨機挑一個章節，或自選範圍與張數抽出今日隨機練習卡組（問題與答案同步並排顯示）',
    randDeckBtn: '隨機抽出卡組',
    randChBtn: '隨機挑一章',
    fcModeList: '雙欄總覽',
    fcModeRandomQ: '🎲 隨機題目模式',
    randNextBtn: '隨機抽一題 (R)',
    docsChLabel: '章節排序目錄：',
    writebackSort: '寫回原檔時依 Chapter 排序',
    dropTitle: '放開以載入 .docx 問答筆記',
    dropSub: '自動偵測 Chapter 編號排序，並保留原始排版與圖片以供回寫 Google Drive',
  },
  en: {
    navHome: 'Home',
    navFlashcards: 'Flashcards',
    navDocs: 'Docs Editor',
    navDrive: 'Google Drive',
    navSettings: 'Appearance',
    promoTitle: 'Quick Random Set',
    promoDesc: 'Pull a fresh shuffled set from your saved cards',
    promoCta: '🎲 Draw Random Set',
    s9Btn: 'S9 Live',
    openDocx: 'Open .docx',
    heroResume: 'Resume Study',
    heroRandom: 'Random Q',
    tileFcLabel: 'Flashcards',
    tileStudyLabel: "Today's Study",
    tileRandomLabel: 'Random Flashcards',
    tileRandomVal: 'Quick Shuffle',
    tileDocsLabel: 'Docs Chapters',
    chartTitle: 'Study Activity',
    breakdownTitle: 'Chapter Mastery & Coverage',
    randDeckTitle: '🎲 Random Flashcard Deck Builder',
    randDeckSub: 'Pick a random chapter or draw a custom-sized random set (Q & A shown simultaneously)',
    randDeckBtn: 'Build Random Set',
    randChBtn: 'Random Chapter',
    fcModeList: 'All Cards',
    fcModeRandomQ: '🎲 Random Question',
    randNextBtn: 'Next Random (R)',
    docsChLabel: 'Sorted Chapters:',
    writebackSort: 'Sort by Chapter on Drive writeback',
    dropTitle: 'Drop .docx notes to import',
    dropSub: 'Auto-sorts by Chapter and preserves original styles for Google Drive writeback',
  },
};

const DEFAULT_APPEARANCE = {
  lang: 'zh',
  palette: 'terracotta',
  customAccentHex: '#D96F52',
  sqRadius: 24,
  sqCurvature: 1.48,
  s9Mode: false,
  wpPreset: 's9-silk',
  wpSpeed: 1.0,
  docsPaper: 'dark',
  writebackSortByChapter: false,
};

const driveClient = new GoogleDriveClient();
let wallpaper = null;
let rawDocxBuffer = null; // In-memory original .docx ArrayBuffer for lossless in-place XML patching!
let autoSyncTimer = null;
let clockInterval = null;
let toastTimer = null;

function createInitialState() {
  const saved = loadFromLocalStorage();
  if (saved && saved.doc && Array.isArray(saved.doc.rows) && saved.doc.rows.length > 0) {
    // Ensure chapter metadata exists on older saved rows
    saved.doc.rows = saved.doc.rows.map((r, i) =>
      r.chapterKey ? r : buildDocRow({ ...r, originalOrder: i })
    );
    return {
      activeNav: 'home',
      sidebarCollapsed: false,
      chartMode: 'day', // 'day' | 'chapter'
      fcMode: 'list', // 'list' | 'random-q'
      fcScope: 'all', // 'all' | 'day-X' | 'ch-X' | 'pending'
      customRandomDeckIds: null, // array of rowIds when user builds a Random Flashcard Set
      customRandomDeckLabel: '',
      randomQueue: [], // no-repeat shuffled rowIds for Random Question mode
      randomCursor: 0,
      docsChapterFilter: 'all',
      searchQuery: '',
      currentDay: saved.currentDay || 3,
      studyStartDate: saved.studyStartDate || addDaysISO(toLocalISO(), -2),
      doc: saved.doc,
      flashcardsMap: saved.flashcardsMap || {},
      appearance: { ...DEFAULT_APPEARANCE, ...(saved.appearance || {}) },
      driveFiles: [],
      driveSyncStatus: 'idle', // 'idle' | 'pending' | 'syncing' | 'synced' | 'error'
      driveSyncTime: '',
    };
  }

  const startISO = addDaysISO(toLocalISO(), -2);
  return {
    activeNav: 'home',
    sidebarCollapsed: false,
    chartMode: 'day',
    fcMode: 'list',
    fcScope: 'all',
    customRandomDeckIds: null,
    customRandomDeckLabel: '',
    randomQueue: [],
    randomCursor: 0,
    docsChapterFilter: 'all',
    searchQuery: '',
    currentDay: 3,
    studyStartDate: startISO,
    doc: structuredClone(SAMPLE_DOC_DATA),
    flashcardsMap: buildInitialSeededFlashcards(startISO),
    appearance: { ...DEFAULT_APPEARANCE },
    driveFiles: [],
    driveSyncStatus: 'idle',
    driveSyncTime: '',
  };
}

const state = createInitialState();

// DOM Element References
const els = {
  html: document.documentElement,
  body: document.body,
  canvas: document.getElementById('s9-wallpaper-canvas'),
  appLayout: document.getElementById('app-layout'),
  fileInput: document.getElementById('docx-file-input'),
  globalDropVeil: document.getElementById('global-drop-veil'),
  toastPill: document.getElementById('toast-pill'),
  toastText: document.getElementById('toast-text'),

  // Sidebar & Topbar
  btnCollapseSidebar: document.getElementById('btn-collapse-sidebar'),
  btnExpandSidebar: document.getElementById('btn-expand-sidebar'),
  sidebarSubLabel: document.getElementById('sidebar-sub-label'),
  topbarBreadcrumb: document.getElementById('topbar-breadcrumb'),
  sideNavBtns: document.querySelectorAll('[data-nav]'),
  btnSidebarRandom: document.getElementById('btn-sidebar-random'),
  btnSidebarProfile: document.getElementById('btn-sidebar-profile'),
  sidebarDriveCaption: document.getElementById('sidebar-drive-caption'),
  sidebarDriveDot: document.getElementById('sidebar-drive-dot'),
  langBtns: document.querySelectorAll('[data-lang]'),
  paletteDots: document.querySelectorAll('[data-set-palette]'),
  btnS9ModeToggle: document.getElementById('btn-s9-mode-toggle'),
  globalSearchInput: document.getElementById('global-search-input'),
  btnTopAvatar: document.getElementById('btn-top-avatar'),
  pageGreeting: document.getElementById('page-greeting'),
  liveDateLine: document.getElementById('live-date-line'),
  activeFileCaption: document.getElementById('active-file-caption'),
  btnOpenDocx: document.getElementById('btn-open-docx'),

  // Views
  viewHome: document.getElementById('view-home'),
  viewFlashcards: document.getElementById('view-flashcards'),
  viewDocs: document.getElementById('view-docs'),
  viewDrive: document.getElementById('view-drive'),
  viewSettings: document.getElementById('view-settings'),

  // Home Dashboard
  heroDayLabel: document.getElementById('hero-day-label'),
  btnDayMinus: document.getElementById('btn-day-minus'),
  btnDayPlus: document.getElementById('btn-day-plus'),
  heroShortDate: document.getElementById('hero-short-date'),
  heroClock: document.getElementById('hero-clock'),
  heroFullDate: document.getElementById('hero-full-date'),
  btnHeroResume: document.getElementById('btn-hero-resume'),
  btnHeroRandomQ: document.getElementById('btn-hero-random-q'),
  heroRingCircle: document.getElementById('hero-ring-circle'),
  heroRingPct: document.getElementById('hero-ring-pct'),
  heroMasteryTitle: document.getElementById('hero-mastery-title'),
  heroMasterySub: document.getElementById('hero-mastery-sub'),
  orbitalDotSec: document.getElementById('orbital-dot-sec'),
  orbitalDotMin: document.getElementById('orbital-dot-min'),
  orbitalArc: document.getElementById('orbital-arc'),
  cardHomeFlashcards: document.getElementById('card-home-flashcards'),
  cardHomeStudy: document.getElementById('card-home-study'),
  cardHomeRandom: document.getElementById('card-home-random'),
  cardHomeDocs: document.getElementById('card-home-docs'),
  tileFcValue: document.getElementById('tile-fc-value'),
  tileFcSub: document.getElementById('tile-fc-sub'),
  tileStudyValue: document.getElementById('tile-study-value'),
  tileStudySub: document.getElementById('tile-study-sub'),
  tileDocsValue: document.getElementById('tile-docs-value'),
  tileDocsSub: document.getElementById('tile-docs-sub'),
  chartModeBtns: document.querySelectorAll('[data-chart-mode]'),
  homeBarChart: document.getElementById('home-bar-chart'),
  homeSegmentedBar: document.getElementById('home-segmented-bar'),
  homeChapterBreakdown: document.getElementById('home-chapter-breakdown'),
  btnHomeSyncRefresh: document.getElementById('btn-home-sync-refresh'),

  // Flashcards
  randSourceSelect: document.getElementById('rand-source-select'),
  randCountSelect: document.getElementById('rand-count-select'),
  btnBuildRandomDeck: document.getElementById('btn-build-random-deck'),
  btnPickRandomChapter: document.getElementById('btn-pick-random-chapter'),
  fcModeBtns: document.querySelectorAll('[data-fc-mode]'),
  fcScopeTabs: document.getElementById('fc-scope-tabs'),
  btnFcRandomNext: document.getElementById('btn-fc-random-next'),
  btnFcClearRandomDeck: document.getElementById('btn-fc-clear-random-deck'),
  randomQuizBanner: document.getElementById('random-quiz-banner'),
  randomPoolBadge: document.getElementById('random-pool-badge'),
  randomDeckCaption: document.getElementById('random-deck-caption'),
  btnRandomReshuffle: document.getElementById('btn-random-reshuffle'),
  btnRandomStepNext: document.getElementById('btn-random-step-next'),
  fcCardsList: document.getElementById('fc-cards-list'),

  // Docs Editor
  docsChapterPills: document.getElementById('docs-chapter-pills'),
  chkWritebackChapterSort: document.getElementById('chk-writeback-chapter-sort'),
  docsDriveSyncPill: document.getElementById('docs-drive-sync-pill'),
  btnFmtHighlight: document.getElementById('btn-fmt-highlight'),
  btnDocsPaperMode: document.getElementById('btn-docs-paper-mode'),
  btnDocsAddRow: document.getElementById('btn-docs-add-row'),
  btnDocsExport: document.getElementById('btn-docs-export'),
  btnDocsSyncNow: document.getElementById('btn-docs-sync-now'),
  docsPaper: document.getElementById('docs-paper'),
  docsTitleEditable: document.getElementById('docs-title-editable'),
  docsAddedSummary: document.getElementById('docs-added-summary'),
  btnDocsAddAll: document.getElementById('btn-docs-add-all'),
  docsGroupedBody: document.getElementById('docs-grouped-body'),

  // Google Drive
  driveStatusBadge: document.getElementById('drive-status-badge'),
  driveModeBtns: document.querySelectorAll('[data-drive-mode]'),
  driveOauthBox: document.getElementById('drive-oauth-box'),
  driveScriptBox: document.getElementById('drive-script-box'),
  originHintCode: document.getElementById('origin-hint-code'),
  inputGdriveClientId: document.getElementById('input-gdrive-client-id'),
  btnDriveOauthConnect: document.getElementById('btn-drive-oauth-connect'),
  btnDriveDisconnect: document.getElementById('btn-drive-disconnect'),
  inputAppsScriptUrl: document.getElementById('input-apps-script-url'),
  btnDriveScriptConnect: document.getElementById('btn-drive-script-connect'),
  btnCopyAppsScript: document.getElementById('btn-copy-apps-script'),
  appsScriptCode: document.getElementById('apps-script-code'),
  btnDrivePushAll: document.getElementById('btn-drive-push-all'),
  btnDrivePullAll: document.getElementById('btn-drive-pull-all'),
  btnDriveListFolder: document.getElementById('btn-drive-list-folder'),
  btnDriveListAll: document.getElementById('btn-drive-list-all'),
  driveFilesHint: document.getElementById('drive-files-hint'),
  driveFilesContainer: document.getElementById('drive-files-container'),

  // Settings
  settingsCustomAccent: document.getElementById('settings-custom-accent'),
  settingsAccentHex: document.getElementById('settings-accent-hex'),
  settingsRadiusSlider: document.getElementById('settings-radius-slider'),
  settingsRadiusVal: document.getElementById('settings-radius-val'),
  settingsKSlider: document.getElementById('settings-k-slider'),
  settingsKVal: document.getElementById('settings-k-val'),
  settingsS9Checkbox: document.getElementById('settings-s9-checkbox'),
  wpPresetBtns: document.querySelectorAll('[data-wp-preset]'),
  settingsWpSpeed: document.getElementById('settings-wp-speed'),
  settingsWpSpeedVal: document.getElementById('settings-wp-speed-val'),
};

function init() {
  wallpaper = new GalaxyS9Wallpaper(els.canvas);
  wallpaper.start();

  // Auto-collapse sidebar on narrow portrait screens (like IMG_9403)
  if (window.innerWidth < 900) {
    state.sidebarCollapsed = true;
  }

  els.originHintCode.textContent = window.location.origin;
  els.inputGdriveClientId.value = driveClient.clientId;
  els.inputAppsScriptUrl.value = driveClient.appsScriptUrl;
  els.appsScriptCode.textContent = APPS_SCRIPT_TEMPLATE;

  applyAppearance();
  startLiveClock();
  bindEvents();
  renderAll();
}

/* ==========================================================================
   Persistence & Live Debounced Google Drive Writeback
   ========================================================================== */

function loadFromLocalStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function getSerializableState() {
  return {
    version: 3,
    savedAt: new Date().toISOString(),
    currentDay: state.currentDay,
    studyStartDate: state.studyStartDate,
    doc: state.doc,
    flashcardsMap: state.flashcardsMap,
    appearance: state.appearance,
  };
}

function persistLocal() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(getSerializableState()));
  } catch (e) {
    console.warn('LocalStorage write skipped:', e);
  }
}

/**
 * Schedules a 2.5-second debounced in-place writeback of the edited .docx
 * straight back into Google Drive (plus workspace state JSON).
 */
function scheduleDriveAutoSync() {
  persistLocal();
  if (!driveClient.isConnected()) {
    state.driveSyncStatus = 'idle';
    renderDocsSyncStatus();
    return;
  }

  state.driveSyncStatus = 'pending';
  renderDocsSyncStatus();

  clearTimeout(autoSyncTimer);
  autoSyncTimer = setTimeout(() => {
    syncDocAndStateToDrive(false);
  }, 2500);
}

async function syncDocAndStateToDrive(showSuccessToast = true) {
  if (!driveClient.isConnected()) {
    if (showSuccessToast) {
      state.activeNav = 'drive';
      renderAll();
      showToast('請先連接 Google Drive，修改內容即可自動覆寫雲端原檔');
    }
    return;
  }

  clearTimeout(autoSyncTimer);
  state.driveSyncStatus = 'syncing';
  renderDocsSyncStatus();

  try {
    const docxBlob = await buildUpdatedDocxBlob(state.doc, rawDocxBuffer, {
      sortByChapter: state.appearance.writebackSortByChapter,
    });
    const savedFile = await driveClient.saveDocxToDrive(
      state.doc.fileName,
      docxBlob,
      state.doc.driveFileId
    );
    if (savedFile?.id) {
      state.doc.driveFileId = savedFile.id;
    }
    // Refresh in-memory rawDocxBuffer to the newly patched blob so subsequent edits stay cumulative!
    rawDocxBuffer = await docxBlob.arrayBuffer();
    state.doc.rows.forEach((r) => {
      r.dirty = false;
    });

    await driveClient.saveWorkspaceState(getSerializableState());
    persistLocal();

    state.driveSyncStatus = 'synced';
    state.driveSyncTime = new Date().toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    renderDocsSyncStatus();

    if (showSuccessToast) {
      showToast(`✓ 已將修改寫回 Google Drive 原檔 (${state.doc.fileName})`);
    }
  } catch (err) {
    state.driveSyncStatus = 'error';
    renderDocsSyncStatus();
    showToast(`寫回 Google Drive 失敗：${err.message}`);
  }
}

/* ==========================================================================
   Live Clock & Concentric Orbital Disc Animation (IMG_9404 & IMG_9403)
   ========================================================================== */

function startLiveClock() {
  const tick = () => {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');

    els.heroClock.textContent = `${hh}:${mm}:${ss}`;

    const isEn = state.appearance.lang === 'en';
    const dayStr = now.toLocaleDateString(isEn ? 'en-GB' : 'zh-Hant', {
      weekday: 'short',
      day: 'numeric',
      month: 'short',
    });
    const fullStr = now.toLocaleDateString(isEn ? 'en-GB' : 'zh-Hant', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric',
    });
    els.heroShortDate.textContent = dayStr;
    els.heroFullDate.textContent = fullStr;
    els.liveDateLine.textContent = `${fullStr} · ${hh}:${mm}`;

    // Greeting text
    const hr = now.getHours();
    if (isEn) {
      const part = hr < 12 ? 'Good morning' : hr < 18 ? 'Good afternoon' : 'Good evening';
      els.pageGreeting.textContent = `${part}, Wesley`;
    } else {
      const part = hr < 12 ? '早安' : hr < 18 ? '午安' : '晚安';
      els.pageGreeting.textContent = `${part}，Wesley`;
    }

    // Animate Concentric Orbital Disc Satellite Dots!
    const secAngle = ((now.getSeconds() + now.getMilliseconds() / 1000) / 60) * Math.PI * 2 - Math.PI / 2;
    const minAngle = ((now.getMinutes() + now.getSeconds() / 60) / 60) * Math.PI * 2 - Math.PI / 2;

    const sx = 80 + Math.cos(secAngle) * 64;
    const sy = 80 + Math.sin(secAngle) * 64;
    const mx = 80 + Math.cos(minAngle) * 48;
    const my = 80 + Math.sin(minAngle) * 48;

    els.orbitalDotSec.setAttribute('cx', sx.toFixed(1));
    els.orbitalDotSec.setAttribute('cy', sy.toFixed(1));
    els.orbitalDotMin.setAttribute('cx', mx.toFixed(1));
    els.orbitalDotMin.setAttribute('cy', my.toFixed(1));

    // Dynamic orbital trail arc behind second dot
    const startA = secAngle - 1.15;
    const ax1 = 80 + Math.cos(startA) * 64;
    const ay1 = 80 + Math.sin(startA) * 64;
    els.orbitalArc.setAttribute('d', `M ${ax1.toFixed(1)} ${ay1.toFixed(1)} A 64 64 0 0 1 ${sx.toFixed(1)} ${sy.toFixed(1)}`);
  };

  tick();
  clearInterval(clockInterval);
  clockInterval = setInterval(tick, 1000);
}

/* ==========================================================================
   Appearance, Bilingual I18N, Continuous Curvature & Optional S9 Mode
   ========================================================================== */

function applyAppearance() {
  const a = state.appearance;

  els.html.setAttribute('data-palette', a.palette);
  if (a.customAccentHex) {
    els.html.style.setProperty('--accent', a.customAccentHex);
    els.html.style.setProperty('--accent-deep', a.customAccentHex);
    els.html.style.setProperty('--accent-soft', hexToRgba(a.customAccentHex, 0.15));
    els.html.style.setProperty('--accent-glow', hexToRgba(a.customAccentHex, 0.28));
  }

  els.html.style.setProperty('--sq-radius', `${a.sqRadius}px`);
  els.html.style.setProperty('--sq-k', String(a.sqCurvature));

  // Optional Samsung Galaxy S9 Live Wallpaper Mode
  els.body.classList.toggle('s9-wallpaper-mode', Boolean(a.s9Mode));
  els.btnS9ModeToggle.classList.toggle('active', Boolean(a.s9Mode));
  els.settingsS9Checkbox.checked = Boolean(a.s9Mode);

  if (wallpaper) {
    wallpaper.setAccentHex(a.customAccentHex || PALETTE_ACCENTS[a.palette] || '#D96F52');
    wallpaper.setPreset(a.wpPreset);
    wallpaper.setSpeed(a.wpSpeed);
    wallpaper.setPaused(!a.s9Mode);
  }

  // Sync Palette Dots & Settings Controls
  els.paletteDots.forEach((d) => d.classList.toggle('active', d.dataset.setPalette === a.palette));
  els.settingsCustomAccent.value = a.customAccentHex;
  els.settingsAccentHex.textContent = a.customAccentHex.toUpperCase();
  els.settingsRadiusSlider.value = String(a.sqRadius);
  els.settingsRadiusVal.textContent = `${a.sqRadius}px`;
  els.settingsKSlider.value = String(a.sqCurvature);
  els.settingsKVal.textContent = Number(a.sqCurvature).toFixed(2);
  els.wpPresetBtns.forEach((b) => b.classList.toggle('active', b.dataset.wpPreset === a.wpPreset));
  els.settingsWpSpeed.value = String(a.wpSpeed);
  els.settingsWpSpeedVal.textContent = `${Number(a.wpSpeed).toFixed(1)}×`;

  // Docs paper light/dark
  els.docsPaper.classList.toggle('paper-light', a.docsPaper === 'light');
  els.btnDocsPaperMode.textContent = a.docsPaper === 'dark' ? '🌙 深色原檔' : '☀ 淺色原檔';
  els.chkWritebackChapterSort.checked = Boolean(a.writebackSortByChapter);

  // Bilingual I18N
  const dict = I18N[a.lang] || I18N.zh;
  els.langBtns.forEach((b) => b.classList.toggle('active', b.dataset.lang === a.lang));
  document.querySelectorAll('[data-i18n]').forEach((el) => {
    const key = el.dataset.i18n;
    if (dict[key]) el.textContent = dict[key];
  });
}

/* ==========================================================================
   Data Helpers (Rows, Chapters, Flashcards & Random Queues)
   ========================================================================== */

function getRowById(rowId) {
  return state.doc.rows.find((r) => r.id === rowId);
}

function getAddedFlashcardRows() {
  return state.doc.rows
    .filter((r) => Boolean(state.flashcardsMap[r.id]))
    .map((row) => ({ row, meta: state.flashcardsMap[row.id] }));
}

function getChaptersSummary() {
  return groupRowsByChapter(state.doc.rows);
}

function addRowToFlashcards(rowId, day = state.currentDay) {
  const row = getRowById(rowId);
  if (!row) return;
  if (state.flashcardsMap[rowId]) {
    showToast(`Q${String(row.index).padStart(2, '0')} 已在 Day ${state.flashcardsMap[rowId].addedDay} 加入過閃卡`);
    return;
  }
  state.flashcardsMap[rowId] = {
    rowId,
    addedDay: day,
    addedDate: toLocalISO(),
    mastered: false,
    bookmarked: false,
    checkedPoints: [],
  };
  persistLocal();
  renderAll();
  showToast(`✓ 已加入 Q${String(row.index).padStart(2, '0')}（${row.chapterShort} · Day ${day}）`);
  if (driveClient.isConnected()) {
    driveClient.saveWorkspaceState(getSerializableState()).catch(() => {});
  }
}

function removeRowFromFlashcards(rowId) {
  if (!state.flashcardsMap[rowId]) return;
  delete state.flashcardsMap[rowId];
  persistLocal();
  renderAll();
  showToast('已從閃卡清單移出（原檔題目保留）');
}

function shuffleArray(arr) {
  const copy = [...arr];
  for (let i = copy.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

/**
 * ① 選 Flashcard 處：建立隨機閃卡組 (Random Flashcard Set)
 */
function buildRandomFlashcardDeck(source = 'added', countVal = '5') {
  let poolRows = [];
  if (source === 'added') {
    poolRows = state.doc.rows.filter((r) => Boolean(state.flashcardsMap[r.id]));
  } else if (source === 'pending') {
    poolRows = state.doc.rows.filter((r) => state.flashcardsMap[r.id] && !state.flashcardsMap[r.id].mastered);
  } else if (source === 'all-doc') {
    poolRows = [...state.doc.rows];
  } else if (source.startsWith('ch-')) {
    poolRows = state.doc.rows.filter((r) => r.chapterKey === source);
  }

  if (poolRows.length === 0) {
    // Graceful fallback to all doc rows so random button always gives instant delight
    poolRows = [...state.doc.rows];
  }

  const shuffled = shuffleArray(poolRows);
  const picked = countVal === 'all' ? shuffled : shuffled.slice(0, Math.min(Number(countVal) || 5, shuffled.length));

  state.customRandomDeckIds = picked.map((r) => r.id);
  state.customRandomDeckLabel = `隨機抽出 ${picked.length} 張閃卡組`;
  state.randomQueue = [...state.customRandomDeckIds];
  state.randomCursor = 0;
  state.activeNav = 'flashcards';
  state.fcMode = 'list';
  renderAll();
  showToast(`🎲 已為你隨機抽出 ${picked.length} 張閃卡組成練習牌組！`);
}

/**
 * ① 選 Flashcard 處：隨機挑選一個 Chapter 開始複習
 */
function pickRandomChapterDeck() {
  const chapters = getChaptersSummary();
  if (chapters.length === 0) return;
  const chosen = chapters[Math.floor(Math.random() * chapters.length)];
  state.customRandomDeckIds = shuffleArray(chosen.rows).map((r) => r.id);
  state.customRandomDeckLabel = `隨機抽中章節：${chosen.chapterLabel} (${chosen.rows.length} 題)`;
  state.randomQueue = [...state.customRandomDeckIds];
  state.randomCursor = 0;
  state.activeNav = 'flashcards';
  state.fcMode = 'list';
  renderAll();
  showToast(`🎯 隨機挑中 ${chosen.chapterLabel}（共 ${chosen.rows.length} 題）！`);
}

/**
 * ② Flashcard 頁面內：無重複隨機抽下一題 (Random Question Pool)
 */
function stepRandomQuestion(forceFresh = false) {
  const candidates = getCurrentFlashcardCandidates();
  if (candidates.length === 0) {
    showToast('目前範圍沒有閃卡可抽題，請先調整篩選或從 Docs 加入題目');
    return;
  }

  const candidateIds = candidates.map((c) => c.row.id);
  const queueValid =
    !forceFresh &&
    state.randomQueue.length === candidateIds.length &&
    state.randomQueue.every((id) => candidateIds.includes(id));

  if (!queueValid) {
    state.randomQueue = shuffleArray(candidateIds);
    state.randomCursor = 0;
  } else {
    state.randomCursor += 1;
    if (state.randomCursor >= state.randomQueue.length) {
      state.randomQueue = shuffleArray(candidateIds);
      state.randomCursor = 0;
      showToast('🎉 本輪所有題目皆已無重複抽完！已自動開啟全新洗牌輪次');
    }
  }

  state.fcMode = 'random-q';
  renderAll();
}

/* ==========================================================================
   Master Render Orchestrator
   ========================================================================== */

function renderAll() {
  els.appLayout.classList.toggle('sidebar-collapsed', state.sidebarCollapsed);

  const navLabels = {
    home: 'Home',
    flashcards: 'Flashcards',
    docs: 'Docs 原檔',
    drive: 'Google Drive',
    settings: 'Appearance',
  };
  els.topbarBreadcrumb.textContent = navLabels[state.activeNav] || 'Home';
  els.sidebarSubLabel.textContent = navLabels[state.activeNav] || 'Home';
  els.activeFileCaption.textContent = state.doc.fileName;

  // Drive Status
  const connected = driveClient.isConnected();
  els.sidebarDriveDot.classList.toggle('connected', connected);
  els.sidebarDriveCaption.textContent = connected
    ? driveClient.userEmail || 'Drive / flashcard'
    : '@Wesley67123a';

  // Active nav buttons & sections
  els.sideNavBtns.forEach((b) => b.classList.toggle('active', b.dataset.nav === state.activeNav));
  els.viewHome.classList.toggle('active', state.activeNav === 'home');
  els.viewFlashcards.classList.toggle('active', state.activeNav === 'flashcards');
  els.viewDocs.classList.toggle('active', state.activeNav === 'docs');
  els.viewDrive.classList.toggle('active', state.activeNav === 'drive');
  els.viewSettings.classList.toggle('active', state.activeNav === 'settings');

  renderHomeDashboard();
  renderFlashcardsView();
  renderDocsView();
  renderDriveView();
}

/* ==========================================================================
   VIEW 1: HOME DASHBOARD RENDERER (IMG_9404 & IMG_9403)
   ========================================================================== */

function renderHomeDashboard() {
  const addedItems = getAddedFlashcardRows();
  const chapters = getChaptersSummary();
  const totalDocs = state.doc.rows.length;
  const totalAdded = addedItems.length;
  const mastered = addedItems.filter((x) => x.meta.mastered).length;
  const pending = totalAdded - mastered;
  const pct = totalAdded > 0 ? Math.round((mastered / totalAdded) * 100) : 0;

  // Hero Ring & Labels
  els.heroDayLabel.textContent = `Day ${state.currentDay}`;
  const circumference = 2 * Math.PI * 19; // 119.38
  els.heroRingCircle.setAttribute('stroke-dasharray', circumference.toFixed(2));
  els.heroRingCircle.setAttribute('stroke-dashoffset', (circumference * (1 - pct / 100)).toFixed(2));
  els.heroRingPct.textContent = `${pct}%`;
  els.heroMasteryTitle.textContent =
    state.appearance.lang === 'en'
      ? `Mastered ${mastered} / ${totalAdded} Cards`
      : `已掌握 ${mastered} / ${totalAdded} 張閃卡`;
  els.heroMasterySub.textContent =
    state.appearance.lang === 'en'
      ? `${totalAdded} / ${totalDocs} doc questions collected`
      : `原檔已收錄 ${totalAdded} / ${totalDocs} 題`;

  // 4 Pastel Cards
  els.tileFcValue.textContent = `${totalAdded} Cards`;
  els.tileFcSub.textContent = `${mastered} 已掌握 · ${pending} 待複習`;

  // Infer Subject Name from filename/title
  const subjectGuess = /chem/i.test(state.doc.title + state.doc.fileName)
    ? 'Chemistry'
    : /phys/i.test(state.doc.title + state.doc.fileName)
      ? 'Physics'
      : /bio/i.test(state.doc.title + state.doc.fileName)
        ? 'Biology'
        : state.doc.title.slice(0, 14);
  const numericChapters = chapters.filter((c) => c.chapterNum < 9999);
  const chRange =
    numericChapters.length >= 2
      ? `${numericChapters[0].chapterShort} – ${numericChapters[numericChapters.length - 1].chapterShort}`
      : numericChapters[0]?.chapterShort || `${chapters.length} Sections`;
  els.tileStudyValue.textContent = subjectGuess;
  els.tileStudySub.textContent = chRange;

  els.tileDocsValue.textContent = `${chapters.length} Chapters`;
  els.tileDocsSub.textContent = driveClient.isConnected() ? '☁ Drive 自動同步已啟用' : '依 Chapter 自然排序';

  // Bottom Left Capsule Bar Chart (A · Day / B · Chapter)
  els.chartModeBtns.forEach((b) => b.classList.toggle('active', b.dataset.chartMode === state.chartMode));

  let bars = [];
  if (state.chartMode === 'day') {
    const maxDay = Math.max(state.currentDay, ...addedItems.map((x) => x.meta.addedDay || 1), 6);
    for (let d = 1; d <= Math.min(maxDay + 1, 7); d++) {
      const cnt = addedItems.filter((x) => x.meta.addedDay === d).length;
      bars.push({
        label: `D${d}`,
        count: cnt,
        highlight: d === state.currentDay,
        actionType: 'day',
        actionKey: `day-${d}`,
      });
    }
  } else {
    bars = chapters.slice(0, 7).map((ch, idx) => ({
      label: ch.chapterShort,
      count: ch.rows.length,
      highlight: idx === 1 || (chapters.length === 1 && idx === 0),
      actionType: 'chapter',
      actionKey: ch.chapterKey,
    }));
  }

  const maxBar = Math.max(...bars.map((b) => b.count), 1);
  els.homeBarChart.innerHTML = bars
    .map((b) => {
      const hPct = b.count === 0 ? 18 : Math.max(26, Math.round((b.count / maxBar) * 100));
      return `
        <div
          class="capsule-bar-col ${b.highlight ? 'is-highlight' : ''}"
          data-bar-type="${b.actionType}"
          data-bar-key="${escapeAttr(b.actionKey)}"
          title="${escapeAttr(b.label)}: ${b.count} 題"
        >
          <span style="font-size: 0.7rem; font-weight: 700; color: var(--ink-secondary)">${b.count}</span>
          <div class="capsule-bar-pill" style="height: ${hPct}%"></div>
          <span class="capsule-bar-lbl">${escapeHtml(b.label)}</span>
        </div>
      `;
    })
    .join('');

  // Bottom Right Multi-Color Segmented Horizontal Bar + Chapter Mastery Breakdown
  const totalAllRows = Math.max(state.doc.rows.length, 1);
  els.homeSegmentedBar.innerHTML = chapters
    .map((ch, i) => {
      const color = BREAKDOWN_COLORS[i % BREAKDOWN_COLORS.length];
      return `<div class="segmented-hbar-seg" style="flex-grow: ${ch.rows.length}; background: ${color}" title="${escapeAttr(ch.chapterLabel)}: ${ch.rows.length} 題"></div>`;
    })
    .join('');

  els.homeChapterBreakdown.innerHTML = chapters
    .map((ch, i) => {
      const color = BREAKDOWN_COLORS[i % BREAKDOWN_COLORS.length];
      const addedInCh = ch.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length;
      const fillPct = ch.rows.length > 0 ? Math.round((addedInCh / ch.rows.length) * 100) : 0;
      return `
        <div class="chapter-break-row" data-jump-chapter-docs="${escapeAttr(ch.chapterKey)}" title="點擊前往 Docs 檢視 ${escapeAttr(ch.chapterLabel)}">
          <span class="break-dot" style="background: ${color}"></span>
          <strong style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis">${escapeHtml(ch.chapterShort)}</strong>
          <div class="break-track">
            <div class="break-fill" style="width: ${Math.max(fillPct, addedInCh > 0 ? 12 : 4)}%; background: ${color}"></div>
          </div>
          <span style="font-family: 'JetBrains Mono', monospace; font-size: 0.73rem; color: var(--ink-secondary)">
            ${addedInCh}/${ch.rows.length}
          </span>
        </div>
      `;
    })
    .join('');
}

/* ==========================================================================
   VIEW 2: FLASHCARDS VIEW (Random Deck Builder + Simultaneous Q&A + Random Q)
   ========================================================================== */

function getCurrentFlashcardCandidates() {
  const q = state.searchQuery.trim().toLowerCase();

  let list = [];
  if (state.customRandomDeckIds && state.customRandomDeckIds.length > 0) {
    list = state.customRandomDeckIds
      .map((id) => getRowById(id))
      .filter(Boolean)
      .map((row) => ({
        row,
        meta: state.flashcardsMap[row.id] || {
          rowId: row.id,
          addedDay: null,
          mastered: false,
          bookmarked: false,
          checkedPoints: [],
        },
      }));
  } else {
    const added = getAddedFlashcardRows();
    list = added.filter(({ row, meta }) => {
      if (state.fcScope === 'pending') return !meta.mastered;
      if (state.fcScope.startsWith('day-')) return meta.addedDay === Number(state.fcScope.replace('day-', ''));
      if (state.fcScope.startsWith('ch-')) return row.chapterKey === state.fcScope;
      return true;
    });
  }

  if (!q) return list;
  return list.filter(
    ({ row, meta }) =>
      row.question.toLowerCase().includes(q) ||
      row.answerRaw.toLowerCase().includes(q) ||
      (row.chapterLabel || '').toLowerCase().includes(q) ||
      (meta.addedDay ? `day ${meta.addedDay}`.includes(q) : false)
  );
}

function renderFlashcardsView() {
  const chapters = getChaptersSummary();
  const addedItems = getAddedFlashcardRows();

  // Populate Random Deck Source <select> with dynamic Chapters!
  const prevSource = els.randSourceSelect.value || 'added';
  els.randSourceSelect.innerHTML = `
    <option value="added">已收錄的閃卡 (${addedItems.length})</option>
    <option value="pending">待複習閃卡 (${addedItems.filter((x) => !x.meta.mastered).length})</option>
    <option value="all-doc">原檔全部題目 (${state.doc.rows.length})</option>
    ${chapters
      .map((ch) => `<option value="${escapeAttr(ch.chapterKey)}">只抽 ${escapeHtml(ch.chapterShort)} (${ch.rows.length} 題)</option>`)
      .join('')}
  `;
  if ([...els.randSourceSelect.options].some((o) => o.value === prevSource)) {
    els.randSourceSelect.value = prevSource;
  }

  // Mode Tabs
  els.fcModeBtns.forEach((b) => b.classList.toggle('active', b.dataset.fcMode === state.fcMode));

  // Scope Filter Tabs (All, Days, Chapters)
  const days = [...new Set(addedItems.map((x) => x.meta.addedDay))].sort((a, b) => a - b);
  const scopeItems = [
    { key: 'all', label: `全部 (${addedItems.length})` },
    { key: 'pending', label: '待複習' },
    ...days.map((d) => ({ key: `day-${d}`, label: `Day ${d}` })),
    ...chapters.map((ch) => ({ key: ch.chapterKey, label: ch.chapterShort })),
  ];
  els.fcScopeTabs.innerHTML = scopeItems
    .map(
      (t) => `
      <button class="pill-tab ${!state.customRandomDeckIds && state.fcScope === t.key ? 'active' : ''}" data-fc-scope="${escapeAttr(t.key)}" type="button">
        ${escapeHtml(t.label)}
      </button>
    `
    )
    .join('');

  els.btnFcClearRandomDeck.style.display = state.customRandomDeckIds ? 'inline-flex' : 'none';

  const candidates = getCurrentFlashcardCandidates();
  if (candidates.length === 0) {
    els.randomQuizBanner.style.display = 'none';
    els.fcCardsList.innerHTML = `
      <div class="surface-card sq-card" style="text-align: center; padding: 48px 24px">
        <h3 style="margin: 0 0 8px">這個篩選條件下還沒有閃卡</h3>
        <p style="margin: 0 0 18px; color: var(--ink-secondary)">你可以點擊上方「🎲 隨機抽出卡組」，或前往「Docs 原檔」勾選題目加入！</p>
        <div style="display: flex; justify-content: center; gap: 10px">
          <button class="btn-soft accent" data-quick-goto="docs" type="button">前往 Docs 原檔收錄題目</button>
          <button class="btn-soft" id="btn-fc-reset-scope" type="button">顯示全部已收錄閃卡</button>
        </div>
      </div>
    `;
    document.getElementById('btn-fc-reset-scope')?.addEventListener('click', () => {
      state.customRandomDeckIds = null;
      state.fcScope = 'all';
      state.searchQuery = '';
      els.globalSearchInput.value = '';
      renderAll();
    });
    return;
  }

  // Random Question Focus Mode vs All-Cards Simultaneous List
  if (state.fcMode === 'random-q') {
    const candidateIds = candidates.map((c) => c.row.id);
    if (
      state.randomQueue.length !== candidateIds.length ||
      !state.randomQueue.every((id) => candidateIds.includes(id))
    ) {
      state.randomQueue = shuffleArray(candidateIds);
      state.randomCursor = 0;
    }
    const activeId = state.randomQueue[state.randomCursor] || candidateIds[0];
    const activeItem = candidates.find((c) => c.row.id === activeId) || candidates[0];

    els.randomQuizBanner.style.display = 'flex';
    els.randomPoolBadge.textContent = `🎲 無重複隨機題目 · ${state.randomCursor + 1} / ${state.randomQueue.length}`;
    els.randomDeckCaption.textContent = state.customRandomDeckLabel || '按 R 或點擊「隨機下一題」無重複抽下一題';

    els.fcCardsList.innerHTML = buildFlashcardHtml(activeItem);
  } else {
    els.randomQuizBanner.style.display = state.customRandomDeckIds ? 'flex' : 'none';
    if (state.customRandomDeckIds) {
      els.randomPoolBadge.textContent = `🎲 ${state.customRandomDeckLabel}`;
      els.randomDeckCaption.textContent = '目前顯示你剛剛隨機抽出的閃卡組（按「隨機抽一題」可切換單題無重複測驗）';
    }
    els.fcCardsList.innerHTML = candidates.map((item) => buildFlashcardHtml(item)).join('');
  }
}

function buildFlashcardHtml({ row, meta }) {
  const isAdded = Boolean(state.flashcardsMap[row.id]);
  const points =
    row.answerPoints && row.answerPoints.length > 0
      ? row.answerPoints
      : [{ text: row.answerRaw || '', marks: null }];
  const checkedSet = new Set(meta.checkedPoints || []);

  return `
    <article class="qa-flashcard sq-card ${meta.mastered ? 'is-mastered' : ''}" data-fc-id="${escapeAttr(row.id)}">
      <div class="fc-meta-row">
        <div class="fc-badges">
          <span class="tag-chip" style="font-family: 'JetBrains Mono', monospace; color: var(--accent)">Q${String(row.index).padStart(2, '0')}</span>
          <span class="tag-chip chapter-pill">${escapeHtml(row.chapterShort)}</span>
          ${
            isAdded
              ? `<span class="tag-chip accent-solid" title="於 ${meta.addedDate || ''} 加入">✓ Day ${meta.addedDay} 加入</span>`
              : `<button class="doc-add-btn" style="padding: 3px 10px; font-size: 0.72rem; border-color: var(--accent); color: var(--accent)" data-fc-quick-add="${escapeAttr(row.id)}" type="button">＋ 加入 Day ${state.currentDay} 閃卡</button>`
          }
          ${row.marks != null ? `<span class="tag-chip marks-pill">${row.marks} ${row.marks === 1 ? 'mark' : 'marks'}</span>` : ''}
        </div>

        <div class="fc-actions">
          <button class="chip-btn" data-locate-doc="${escapeAttr(row.id)}" type="button">📄 查看 Docs 原檔</button>
          ${
            isAdded
              ? `
                <button class="chip-btn ${meta.bookmarked ? 'is-star' : ''}" data-fc-star="${escapeAttr(row.id)}" type="button">${meta.bookmarked ? '★ 星標' : '☆ 星標'}</button>
                <button class="chip-btn ${meta.mastered ? 'is-on' : ''}" data-fc-master="${escapeAttr(row.id)}" type="button">${meta.mastered ? '✓ 已掌握' : '標記掌握'}</button>
              `
              : ''
          }
        </div>
      </div>

      <!-- Question & Answer Simultaneously Displayed Side-by-Side -->
      <div class="qa-dual-grid">
        <div class="qa-q-pane sq-md">
          <div class="pane-eyebrow">
            <span>QUESTION · 問題</span>
            ${row.marks != null ? `<span>(${row.marks}m)</span>` : ''}
          </div>
          <p class="q-copy">${highlightText(row.question, state.searchQuery)}</p>
          ${row.source ? `<div style="margin-top: auto; font-size: 0.74rem; color: var(--ink-muted)">*${highlightText(row.source, state.searchQuery)}</div>` : ''}
          ${
            row.questionImages?.length
              ? `<div style="display:flex;gap:8px;flex-wrap:wrap">${row.questionImages
                  .map((src) => `<img src="${src}" style="max-height:180px;border-radius:10px" alt="題目附圖"/>`)
                  .join('')}</div>`
              : ''
          }
        </div>

        <div class="qa-a-pane sq-md">
          <div class="pane-eyebrow">
            <span>ANSWER · 答案與給分點（同步顯示）</span>
            <span>點擊單條核對</span>
          </div>
          <ul class="a-points-list">
            ${points
              .map((pt, pIdx) => {
                const isChecked = checkedSet.has(pIdx);
                return `
                  <li class="a-point-item sq-sm ${isChecked ? 'checked' : ''}" data-fc-pt="${escapeAttr(row.id)}" data-pt-idx="${pIdx}">
                    <span style="display: flex; align-items: flex-start; gap: 9px">
                      <span class="pt-check">✓</span>
                      <span>${highlightText(pt.text, state.searchQuery)}</span>
                    </span>
                    ${pt.marks != null ? `<span class="tag-chip marks-pill" style="padding: 2px 7px">(${pt.marks})</span>` : ''}
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
    </article>
  `;
}

/* ==========================================================================
   VIEW 3: DOCS EDITOR (Chapter-Sorted + Live Auto-Sync to Google Drive!)
   ========================================================================== */

function renderDocsSyncStatus() {
  const pill = els.docsDriveSyncPill;
  if (!driveClient.isConnected()) {
    pill.className = 'sync-status-pill';
    pill.textContent = '本地自動儲存中（連線 Drive 後即時回寫）';
    return;
  }
  if (state.driveSyncStatus === 'pending') {
    pill.className = 'sync-status-pill syncing';
    pill.textContent = '✏️ 編輯中… 2.5 秒後自動同步至 Google Drive';
  } else if (state.driveSyncStatus === 'syncing') {
    pill.className = 'sync-status-pill syncing';
    pill.textContent = '☁ 正在以 XML 就地修補寫回 Google Drive…';
  } else if (state.driveSyncStatus === 'synced') {
    pill.className = 'sync-status-pill synced';
    pill.textContent = `✓ 已同步至 Google Drive 原檔 (${state.driveSyncTime})`;
  } else if (state.driveSyncStatus === 'error') {
    pill.className = 'sync-status-pill';
    pill.textContent = '⚠ Drive 同步失敗，請點「立即寫回」重試';
  } else {
    pill.className = 'sync-status-pill synced';
    pill.textContent = '☁ Google Drive 自動回寫已待命';
  }
}

function renderDocsView() {
  renderDocsSyncStatus();

  if (document.activeElement !== els.docsTitleEditable) {
    els.docsTitleEditable.textContent = state.doc.title;
  }

  // Naturally sorted Chapter groups!
  const chapterGroups = groupRowsByChapter(state.doc.rows);
  els.docsChapterPills.innerHTML = `
    <button class="pill-tab ${state.docsChapterFilter === 'all' ? 'active' : ''}" data-doc-ch="all" type="button">
      全部 (${state.doc.rows.length})
    </button>
    ${chapterGroups
      .map((g) => {
        const added = g.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length;
        return `
          <button class="pill-tab ${state.docsChapterFilter === g.chapterKey ? 'active' : ''}" data-doc-ch="${escapeAttr(g.chapterKey)}" type="button">
            ${escapeHtml(g.chapterShort)} (${added}/${g.rows.length})
          </button>
        `;
      })
      .join('')}
  `;

  const q = state.searchQuery.trim().toLowerCase();
  const visibleGroups = chapterGroups
    .filter((g) => state.docsChapterFilter === 'all' || g.chapterKey === state.docsChapterFilter)
    .map((g) => ({
      ...g,
      rows: g.rows.filter((r) => {
        if (!q) return true;
        return (
          r.questionRaw.toLowerCase().includes(q) ||
          r.answerRaw.toLowerCase().includes(q) ||
          (r.chapterLabel || '').toLowerCase().includes(q)
        );
      }),
    }))
    .filter((g) => g.rows.length > 0);

  const visibleRowCount = visibleGroups.reduce((s, g) => s + g.rows.length, 0);
  const visibleAddedCount = visibleGroups.reduce(
    (s, g) => s + g.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length,
    0
  );
  els.docsAddedSummary.textContent = `已加入閃卡 ${visibleAddedCount} / ${visibleRowCount} 題`;
  els.btnDocsAddAll.textContent = `將剩餘 ${visibleRowCount - visibleAddedCount} 題加入 Day ${state.currentDay}`;
  els.btnDocsAddAll.style.display = visibleAddedCount < visibleRowCount ? 'inline-flex' : 'none';

  els.docsGroupedBody.innerHTML = visibleGroups
    .map((group) => {
      const groupAdded = group.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length;
      return `
        <div class="docs-chapter-group-header sq-sm" id="doc-ch-${escapeAttr(group.chapterKey)}">
          <span>📚 ${escapeHtml(group.chapterLabel)}</span>
          <span style="font-size: 0.76rem; opacity: 0.78">已收錄 ${groupAdded} / ${group.rows.length} 題</span>
        </div>
        ${group.rows
          .map((row) => {
            const addedMeta = state.flashcardsMap[row.id];
            const isAdded = Boolean(addedMeta);
            return `
              <div class="doc-row-block ${isAdded ? 'is-added-row' : ''}" id="doc-row-${escapeAttr(row.id)}" data-doc-row-id="${escapeAttr(row.id)}">
                <div class="doc-cell">
                  <div class="doc-cell-editable" contenteditable="true" spellcheck="false" data-edit-cell="q" data-row-id="${escapeAttr(row.id)}">${row.questionHtml}</div>
                  <span class="doc-row-num-tag">
                    #${String(row.index).padStart(2, '0')} · ${escapeHtml(row.chapterShort)}
                    ${isAdded ? `<strong style="color: var(--accent)">✓ 已加入 (Day ${addedMeta.addedDay})</strong>` : ''}
                  </span>
                </div>

                <div class="doc-cell">
                  <div class="doc-cell-editable" contenteditable="true" spellcheck="false" data-edit-cell="a" data-row-id="${escapeAttr(row.id)}">${row.answerHtml}</div>

                  <div class="doc-row-action-bar">
                    ${
                      isAdded
                        ? `
                          <button class="doc-added-badge" data-doc-jump-fc="${escapeAttr(row.id)}" type="button" title="於 ${addedMeta.addedDate || ''} 加入，點擊前往閃卡">
                            ✓ 已加入 · Day ${addedMeta.addedDay}
                          </button>
                          <button class="doc-mini-btn" data-doc-unadd="${escapeAttr(row.id)}" type="button" title="取消加入閃卡">✕</button>
                        `
                        : `
                          <button class="doc-add-btn" data-doc-add="${escapeAttr(row.id)}" type="button">
                            ＋ 加入 Flashcard (Day ${state.currentDay})
                          </button>
                        `
                    }
                    <button class="doc-mini-btn" data-doc-toggle-break="${escapeAttr(row.id)}" type="button" title="插入/移除分頁符號">⏎</button>
                  </div>
                </div>
              </div>
              ${row.pageBreakAfter ? `<div class="doc-page-break">Page break</div>` : ''}
            `;
          })
          .join('')}
      `;
    })
    .join('');
}

/* ==========================================================================
   VIEW 4: GOOGLE DRIVE VIEW
   ========================================================================== */

function renderDriveView() {
  const connected = driveClient.isConnected();
  els.driveStatusBadge.textContent = connected
    ? `✓ 已連線 (${driveClient.mode === 'oauth' ? 'OAuth 2.0' : 'Apps Script'})`
    : '未連線';
  els.driveStatusBadge.classList.toggle('accent-solid', connected);

  els.driveModeBtns.forEach((b) => b.classList.toggle('active', b.dataset.driveMode === driveClient.mode));
  els.driveOauthBox.style.display = driveClient.mode === 'oauth' ? 'flex' : 'none';
  els.driveScriptBox.style.display = driveClient.mode === 'apps-script' ? 'flex' : 'none';
  els.btnDriveDisconnect.style.display = connected ? 'inline-flex' : 'none';

  if (!connected) {
    els.driveFilesContainer.innerHTML = `<div style="padding: 20px; text-align: center; font-size: 0.84rem; color: var(--ink-muted)">尚未連線 Google Drive。連線後會自動建立並讀取 <code>flashcard/</code> 資料夾內的 .docx 原檔。</div>`;
    return;
  }

  if (state.driveFiles.length === 0) {
    els.driveFilesContainer.innerHTML = `<div style="padding: 20px; text-align: center; font-size: 0.84rem; color: var(--ink-muted)"><code>flashcard/</code> 目前沒有文件。點左側「☁ 同步全部閃卡與原檔到 Drive」即可上傳！</div>`;
    return;
  }

  els.driveFilesContainer.innerHTML = state.driveFiles
    .map(
      (f) => `
      <div class="drive-file-item sq-sm">
        <div style="min-width: 0">
          <strong style="display: block; font-size: 0.84rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis">${escapeHtml(f.name)}</strong>
          <span style="font-size: 0.72rem; color: var(--ink-muted)">${f.mimeType?.includes('google-apps.document') ? 'Google Docs' : '.docx 原檔'}</span>
        </div>
        <button
          class="btn-soft accent"
          style="padding: 6px 12px; font-size: 0.75rem; flex-shrink: 0"
          data-drive-open-id="${escapeAttr(f.id)}"
          data-drive-open-name="${escapeAttr(f.name)}"
          data-drive-open-mime="${escapeAttr(f.mimeType || '')}"
          type="button"
        >開啟並同步編輯</button>
      </div>
    `
    )
    .join('');
}

/* ==========================================================================
   Event Wiring
   ========================================================================== */

function bindEvents() {
  // Sidebar Collapse / Expand
  els.btnCollapseSidebar.addEventListener('click', () => {
    state.sidebarCollapsed = true;
    renderAll();
  });
  els.btnExpandSidebar.addEventListener('click', () => {
    state.sidebarCollapsed = !state.sidebarCollapsed;
    renderAll();
  });

  // Navigation
  els.sideNavBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.activeNav = btn.dataset.nav;
      if (window.innerWidth < 860) state.sidebarCollapsed = true;
      renderAll();
    });
  });

  els.btnSidebarProfile.addEventListener('click', () => {
    state.activeNav = 'drive';
    renderAll();
  });
  els.btnTopAvatar.addEventListener('click', () => {
    state.activeNav = 'drive';
    renderAll();
  });

  // Top Bar: 中/EN Language Switch
  els.langBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.appearance.lang = btn.dataset.lang;
      applyAppearance();
      persistLocal();
      renderAll();
    });
  });

  // Top Bar: Inline 5-Color Accent Swatch Pill (IMG_9403)
  els.paletteDots.forEach((dot) => {
    dot.addEventListener('click', () => {
      const p = dot.dataset.setPalette;
      state.appearance.palette = p;
      state.appearance.customAccentHex = PALETTE_ACCENTS[p] || '#D96F52';
      applyAppearance();
      persistLocal();
    });
  });

  // Top Bar & Settings: Samsung Galaxy S9 Live Wallpaper Mode Toggle
  const toggleS9 = () => {
    state.appearance.s9Mode = !state.appearance.s9Mode;
    applyAppearance();
    persistLocal();
    showToast(state.appearance.s9Mode ? '✦ 已開啟 Samsung Galaxy S9 動態壁紙＋毛玻璃模式' : '已切回柔和奶油粉彩介面');
  };
  els.btnS9ModeToggle.addEventListener('click', toggleS9);
  els.settingsS9Checkbox.addEventListener('change', (e) => {
    state.appearance.s9Mode = e.target.checked;
    applyAppearance();
    persistLocal();
  });

  // Settings: Custom Accent, Continuous Curvature, S9 Preset & Speed
  els.settingsCustomAccent.addEventListener('input', (e) => {
    state.appearance.customAccentHex = e.target.value;
    applyAppearance();
    persistLocal();
  });
  els.settingsRadiusSlider.addEventListener('input', (e) => {
    state.appearance.sqRadius = Number(e.target.value);
    applyAppearance();
    persistLocal();
  });
  els.settingsKSlider.addEventListener('input', (e) => {
    state.appearance.sqCurvature = Number(e.target.value);
    applyAppearance();
    persistLocal();
  });
  els.wpPresetBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.appearance.wpPreset = btn.dataset.wpPreset;
      applyAppearance();
      persistLocal();
    });
  });
  els.settingsWpSpeed.addEventListener('input', (e) => {
    state.appearance.wpSpeed = Number(e.target.value);
    applyAppearance();
    persistLocal();
  });

  // Global Search + ⌘K / Ctrl+K + Keyboard Shortcut R for Random Next Question
  els.globalSearchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    renderAll();
  });

  window.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      els.globalSearchInput.focus();
      return;
    }
    if (e.target.matches('input, textarea, select, [contenteditable="true"]')) return;
    if (state.activeNav === 'flashcards' && (e.key.toLowerCase() === 'r' || e.key === 'ArrowRight')) {
      e.preventDefault();
      stepRandomQuestion(false);
    }
  });

  // Day Stepper (- / +)
  els.btnDayMinus.addEventListener('click', () => {
    state.currentDay = Math.max(1, state.currentDay - 1);
    persistLocal();
    renderAll();
  });
  els.btnDayPlus.addEventListener('click', () => {
    state.currentDay += 1;
    persistLocal();
    renderAll();
  });

  // Home Hero & 4 Pastel Cards Actions
  els.btnHeroResume.addEventListener('click', () => {
    state.activeNav = 'flashcards';
    state.fcMode = 'list';
    renderAll();
  });
  els.btnHeroRandomQ.addEventListener('click', () => {
    state.activeNav = 'flashcards';
    stepRandomQuestion(true);
  });
  els.btnSidebarRandom.addEventListener('click', () => buildRandomFlashcardDeck('added', '5'));
  els.cardHomeFlashcards.addEventListener('click', () => {
    state.activeNav = 'flashcards';
    state.customRandomDeckIds = null;
    renderAll();
  });
  els.cardHomeStudy.addEventListener('click', () => {
    state.activeNav = 'docs';
    renderAll();
  });
  els.cardHomeRandom.addEventListener('click', () => buildRandomFlashcardDeck('added', '5'));
  els.cardHomeDocs.addEventListener('click', () => {
    state.activeNav = 'docs';
    renderAll();
  });

  // Home Chart A/B Toggle & Bar Clicks
  els.chartModeBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.chartMode = btn.dataset.chartMode;
      renderHomeDashboard();
    });
  });

  els.homeBarChart.addEventListener('click', (e) => {
    const col = e.target.closest('[data-bar-key]');
    if (!col) return;
    state.customRandomDeckIds = null;
    state.fcScope = col.dataset.barKey;
    state.activeNav = 'flashcards';
    renderAll();
  });

  els.homeChapterBreakdown.addEventListener('click', (e) => {
    const row = e.target.closest('[data-jump-chapter-docs]');
    if (!row) return;
    state.docsChapterFilter = row.dataset.jumpChapterDocs;
    state.activeNav = 'docs';
    renderAll();
  });

  els.btnHomeSyncRefresh.addEventListener('click', () => syncDocAndStateToDrive(true));

  // Upload .docx + Global Drag & Drop
  els.btnOpenDocx.addEventListener('click', () => els.fileInput.click());
  els.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (file) {
      await handleLocalDocxUpload(file);
      els.fileInput.value = '';
    }
  });

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

  // Flashcards View: Random Deck Builder + Random Question Mode Controls
  els.btnBuildRandomDeck.addEventListener('click', () => {
    buildRandomFlashcardDeck(els.randSourceSelect.value, els.randCountSelect.value);
  });
  els.btnPickRandomChapter.addEventListener('click', pickRandomChapterDeck);

  els.fcModeBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.fcMode = btn.dataset.fcMode;
      if (state.fcMode === 'random-q') {
        stepRandomQuestion(true);
      } else {
        renderAll();
      }
    });
  });

  els.fcScopeTabs.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-fc-scope]');
    if (!btn) return;
    state.customRandomDeckIds = null;
    state.fcScope = btn.dataset.fcScope;
    renderAll();
  });

  els.btnFcRandomNext.addEventListener('click', () => stepRandomQuestion(false));
  els.btnRandomStepNext.addEventListener('click', () => stepRandomQuestion(false));
  els.btnRandomReshuffle.addEventListener('click', () => {
    stepRandomQuestion(true);
    showToast('↻ 已重新洗牌本輪隨機題目池');
  });
  els.btnFcClearRandomDeck.addEventListener('click', () => {
    state.customRandomDeckIds = null;
    state.customRandomDeckLabel = '';
    renderAll();
  });

  // Flashcards Delegated Interactions
  els.fcCardsList.addEventListener('click', (e) => {
    const quickGoto = e.target.closest('[data-quick-goto]');
    if (quickGoto) {
      state.activeNav = quickGoto.dataset.quickGoto;
      renderAll();
      return;
    }

    const quickAdd = e.target.closest('[data-fc-quick-add]');
    if (quickAdd) {
      addRowToFlashcards(quickAdd.dataset.fcQuickAdd, state.currentDay);
      return;
    }

    const locate = e.target.closest('[data-locate-doc]');
    if (locate) {
      const rowId = locate.dataset.locateDoc;
      state.docsChapterFilter = 'all';
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

    const pt = e.target.closest('[data-fc-pt]');
    if (pt) {
      const id = pt.dataset.fcPt;
      const pIdx = Number(pt.dataset.ptIdx);
      if (!state.flashcardsMap[id]) {
        addRowToFlashcards(id, state.currentDay);
      }
      const meta = state.flashcardsMap[id];
      const row = getRowById(id);
      if (meta && row) {
        const set = new Set(meta.checkedPoints || []);
        if (set.has(pIdx)) set.delete(pIdx);
        else set.add(pIdx);
        meta.checkedPoints = [...set];

        if (meta.checkedPoints.length >= Math.max(row.answerPoints?.length || 1, 1) && !meta.mastered) {
          meta.mastered = true;
          showToast(`✓ Q${String(row.index).padStart(2, '0')} 所有給分點已核對，標記已掌握！`);
        }
        persistLocal();
        renderAll();
      }
    }
  });

  // Docs View: Chapter Directory Pills, Writeback Sort Toggle, Rich Text & Auto-Sync!
  els.docsChapterPills.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-doc-ch]');
    if (!btn) return;
    state.docsChapterFilter = btn.dataset.docCh;
    renderDocsView();
  });

  els.chkWritebackChapterSort.addEventListener('change', (e) => {
    state.appearance.writebackSortByChapter = e.target.checked;
    persistLocal();
    showToast(
      e.target.checked
        ? '已開啟：寫回 Google Drive / 匯出時將按 Chapter 自然數字重新排序原檔'
        : '已關閉：寫回時保留原檔段落順序，僅就地修補修改文字'
    );
  });

  document.querySelectorAll('.fmt-btn[data-cmd]').forEach((btn) => {
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    btn.addEventListener('click', () => {
      document.execCommand(btn.dataset.cmd, false, null);
      scheduleDriveAutoSync();
    });
  });

  els.btnFmtHighlight.addEventListener('mousedown', (e) => e.preventDefault());
  els.btnFmtHighlight.addEventListener('click', () => {
    document.execCommand('hiliteColor', false, 'rgba(250, 204, 21, 0.38)');
    scheduleDriveAutoSync();
  });

  els.btnDocsPaperMode.addEventListener('click', () => {
    state.appearance.docsPaper = state.appearance.docsPaper === 'dark' ? 'light' : 'dark';
    applyAppearance();
    persistLocal();
  });

  // Live Typing in Docs Title or Cells -> Updates State + Schedules 2.5s Auto-Sync to Google Drive!
  els.docsTitleEditable.addEventListener('input', () => {
    state.doc.title = els.docsTitleEditable.textContent.trim() || state.doc.title;
    scheduleDriveAutoSync();
  });

  els.docsGroupedBody.addEventListener('input', (e) => {
    const cell = e.target.closest('[data-edit-cell]');
    if (!cell) return;
    const rowId = cell.dataset.rowId;
    const rowIdx = state.doc.rows.findIndex((r) => r.id === rowId);
    if (rowIdx === -1) return;

    const rowBlock = cell.closest('.doc-row-block');
    const qHtml = rowBlock.querySelector('[data-edit-cell="q"]').innerHTML;
    const aHtml = rowBlock.querySelector('[data-edit-cell="a"]').innerHTML;
    state.doc.rows[rowIdx] = updateRowFromHtml(state.doc.rows[rowIdx], qHtml, aHtml);
    scheduleDriveAutoSync();
  });

  els.docsGroupedBody.addEventListener('click', (e) => {
    const addBtn = e.target.closest('[data-doc-add]');
    if (addBtn) {
      addRowToFlashcards(addBtn.dataset.docAdd, state.currentDay);
      return;
    }

    const unaddBtn = e.target.closest('[data-doc-unadd]');
    if (unaddBtn) {
      removeRowFromFlashcards(unaddBtn.dataset.docUnadd);
      return;
    }

    const jumpFc = e.target.closest('[data-doc-jump-fc]');
    if (jumpFc) {
      const id = jumpFc.dataset.docJumpFc;
      state.customRandomDeckIds = null;
      state.fcScope = 'all';
      state.fcMode = 'list';
      state.activeNav = 'flashcards';
      renderAll();
      requestAnimationFrame(() => {
        els.fcCardsList.querySelector(`[data-fc-id="${id}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      });
      return;
    }

    const toggleBreak = e.target.closest('[data-doc-toggle-break]');
    if (toggleBreak) {
      const row = getRowById(toggleBreak.dataset.docToggleBreak);
      if (row) {
        row.pageBreakAfter = !row.pageBreakAfter;
        row.dirty = true;
        renderDocsView();
        scheduleDriveAutoSync();
      }
    }
  });

  els.btnDocsAddRow.addEventListener('click', () => {
    const nextIdx = state.doc.rows.length + 1;
    const targetChapter =
      state.docsChapterFilter !== 'all'
        ? getChaptersSummary().find((g) => g.chapterKey === state.docsChapterFilter)?.chapterLabel
        : 'Chapter 8: Transport in Humans';

    const newRow = buildDocRow({
      id: `row-new-${Date.now().toString(36)}`,
      index: nextIdx,
      originalOrder: state.doc.rows.length,
      section: targetChapter || 'Chapter 1',
      questionLines: ['請在此輸入新題目... (2 marks)'],
      answerLines: ['請在此輸入答案要點一。 (1)', '請在此輸入答案要點二。 (1)'],
      dirty: true,
    });
    state.doc.rows.push(newRow);
    renderAll();
    scheduleDriveAutoSync();
    requestAnimationFrame(() => {
      document.getElementById(`doc-row-${newRow.id}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  });

  els.btnDocsAddAll.addEventListener('click', () => {
    const groups = groupRowsByChapter(state.doc.rows).filter(
      (g) => state.docsChapterFilter === 'all' || g.chapterKey === state.docsChapterFilter
    );
    let added = 0;
    for (const g of groups) {
      for (const r of g.rows) {
        if (!state.flashcardsMap[r.id]) {
          state.flashcardsMap[r.id] = {
            rowId: r.id,
            addedDay: state.currentDay,
            addedDate: toLocalISO(),
            mastered: false,
            bookmarked: false,
            checkedPoints: [],
          };
          added++;
        }
      }
    }
    persistLocal();
    renderAll();
    showToast(`✓ 已批次加入 ${added} 題至 Day ${state.currentDay} 閃卡！`);
  });

  els.btnDocsExport.addEventListener('click', async () => {
    try {
      const blob = await buildUpdatedDocxBlob(state.doc, rawDocxBuffer, {
        sortByChapter: state.appearance.writebackSortByChapter,
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = state.doc.fileName.endsWith('.docx') ? state.doc.fileName : `${state.doc.fileName}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showToast(`⬇ 已匯出 ${a.download}`);
    } catch (err) {
      showToast(`匯出失敗：${err.message}`);
    }
  });

  els.btnDocsSyncNow.addEventListener('click', () => syncDocAndStateToDrive(true));

  // Google Drive Hub Events
  els.driveModeBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      driveClient.setConfig({ mode: btn.dataset.driveMode });
      renderDriveView();
    });
  });

  els.btnDriveOauthConnect.addEventListener('click', async () => {
    try {
      driveClient.setConfig({ mode: 'oauth', clientId: els.inputGdriveClientId.value });
      showToast('正在連接 Google Drive 並建立 flashcard 資料夾...');
      await driveClient.signInWithOAuth(true);
      await refreshDriveFiles(false);
      renderAll();
      showToast('✓ Google Drive 已連線，Docs 修改將自動同步回原檔！');
    } catch (err) {
      showToast(err.message);
    }
  });

  els.btnDriveScriptConnect.addEventListener('click', async () => {
    try {
      driveClient.setConfig({ mode: 'apps-script', appsScriptUrl: els.inputAppsScriptUrl.value });
      showToast('正在驗證 Apps Script 與 flashcard 資料夾...');
      await driveClient.ensureFlashcardFolder();
      await refreshDriveFiles(false);
      renderAll();
      showToast('✓ Apps Script 已連線！');
    } catch (err) {
      showToast(err.message);
    }
  });

  els.btnCopyAppsScript.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(APPS_SCRIPT_TEMPLATE);
      showToast('已複製 Google Apps Script 程式碼！');
    } catch {
      showToast('請手動複製下方程式碼');
    }
  });

  els.btnDriveDisconnect.addEventListener('click', () => {
    driveClient.disconnect();
    state.driveFiles = [];
    state.driveSyncStatus = 'idle';
    renderAll();
    showToast('已中斷 Google Drive 連線');
  });

  els.btnDrivePushAll.addEventListener('click', () => syncDocAndStateToDrive(true));

  els.btnDrivePullAll.addEventListener('click', async () => {
    if (!driveClient.isConnected()) {
      showToast('請先連接 Google Drive');
      return;
    }
    try {
      showToast('正在讀取 Google Drive / flashcard 進度...');
      const cloud = await driveClient.loadWorkspaceState();
      if (!cloud) {
        showToast('尚無雲端狀態檔，可先按「同步全部」建立');
        return;
      }
      state.doc = cloud.doc || state.doc;
      state.flashcardsMap = cloud.flashcardsMap || state.flashcardsMap;
      state.currentDay = cloud.currentDay || state.currentDay;
      state.appearance = { ...DEFAULT_APPEARANCE, ...(cloud.appearance || {}) };
      applyAppearance();
      persistLocal();
      renderAll();
      showToast('✓ 已載入雲端閃卡與原檔狀態！');
    } catch (err) {
      showToast(err.message);
    }
  });

  els.btnDriveListFolder.addEventListener('click', () => refreshDriveFiles(false));
  els.btnDriveListAll.addEventListener('click', () => refreshDriveFiles(true));

  els.driveFilesContainer.addEventListener('click', async (e) => {
    const btn = e.target.closest('[data-drive-open-id]');
    if (!btn) return;
    try {
      const fileId = btn.dataset.driveOpenId;
      const name = btn.dataset.driveOpenName;
      const mime = btn.dataset.driveOpenMime;
      showToast(`正在從 Google Drive 下載 ${name}...`);
      const buf = await driveClient.downloadDocxArrayBuffer(fileId, mime);
      const parsed = await parseDocxDocument(buf, name);
      if (!parsed.rows.length) {
        showToast('此檔案內未偵測到雙欄 Q&A 表格或問答內容');
        return;
      }
      rawDocxBuffer = parsed.rawBuffer;
      state.doc = {
        title: parsed.title,
        fileName: parsed.fileName,
        rows: parsed.rows,
        driveFileId: fileId,
      };
      state.docsChapterFilter = 'all';
      state.activeNav = 'docs';
      state.driveSyncStatus = 'synced';
      state.driveSyncTime = '剛載入';
      persistLocal();
      renderAll();
      showToast(`✓ 已載入 ${parsed.rows.length} 題（按 Chapter 排序），編輯將自動同步回此 Google Drive 檔案！`);
    } catch (err) {
      showToast(err.message);
    }
  });
}

async function refreshDriveFiles(allDrive = false) {
  if (!driveClient.isConnected()) return;
  try {
    els.driveFilesHint.textContent = allDrive
      ? '顯示整個 Google Drive 的 .docx 與 Google Docs'
      : '顯示 Google Drive / flashcard/ 資料夾內的文件';
    state.driveFiles = await driveClient.listDocuments(allDrive);
    renderDriveView();
  } catch (err) {
    showToast(err.message);
  }
}

async function handleLocalDocxUpload(file) {
  if (!file.name.toLowerCase().endsWith('.docx')) {
    showToast('請上傳 .docx 格式文件');
    return;
  }
  try {
    showToast(`正在解析並按 Chapter 排序 ${file.name}...`);
    const parsed = await parseDocxDocument(file, file.name);
    if (!parsed.rows.length) {
      showToast('未在此 .docx 偵測到雙欄問答表格或 Q&A 段落');
      return;
    }
    rawDocxBuffer = parsed.rawBuffer;
    state.doc = {
      title: parsed.title,
      fileName: parsed.fileName,
      rows: parsed.rows,
      driveFileId: null,
    };
    state.docsChapterFilter = 'all';
    state.activeNav = 'docs';
    persistLocal();
    renderAll();
    const chCount = getChaptersSummary().length;
    showToast(`✓ 成功載入 ${parsed.rows.length} 題（自動識別並排序 ${chCount} 個 Chapter）！`);
  } catch (err) {
    showToast(`讀取失敗：${err.message}`);
  }
}

function showToast(msg) {
  els.toastText.textContent = msg;
  els.toastPill.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toastPill.classList.remove('show'), 3200);
}

function hexToRgba(hex, alpha) {
  const clean = String(hex || '#D96F52').replace('#', '');
  const full = clean.length === 3 ? clean.split('').map((c) => c + c).join('') : clean;
  const num = parseInt(full, 16);
  const r = (num >> 16) & 255;
  const g = (num >> 8) & 255;
  const b = num & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
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
  const esc = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return safe.replace(new RegExp(`(${esc})`, 'gi'), '<mark class="hl">$1</mark>');
}

init();

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
import { GoogleDriveClient } from './googleDrive.js';

const META_STORAGE_KEY = 'wesley-study-meta-v4';
const DOCS_PAGE_SIZE = 20;
const FC_LIST_PAGE_SIZE = 20;

const PALETTE_ACCENTS = {
  terracotta: '#D96F52',
  sage: '#528467',
  ocean: '#4C7EB0',
  lavender: '#7B61B8',
  amber: '#C98F32',
};

const BREAKDOWN_COLORS = ['#D96F52', '#D4A138', '#5C7F67', '#557CA5', '#7B61B8', '#0EA5E9'];

const DEFAULT_APPEARANCE = {
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
let rawDocxBuffer = null;
let autoSyncTimer = null;
let toastTimer = null;
let clockTimer = null;

function loadMeta() {
  try {
    const raw = localStorage.getItem(META_STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function createInitialState() {
  const saved = loadMeta();
  const startISO = saved?.studyStartDate || addDaysISO(toLocalISO(), -2);
  const initialDoc = structuredClone(SAMPLE_DOC_DATA);
  const initialChapters = groupRowsByChapter(initialDoc.rows);

  return {
    activeNav: 'home',
    chartMode: 'day',
    searchQuery: '',

    // Study & Flashcards
    currentDay: saved?.currentDay || 3,
    studyStartDate: startISO,
    doc: initialDoc,
    lastDriveFileId: saved?.lastDriveFileId || null,
    flashcardsMap: saved?.flashcardsMap || buildInitialSeededFlashcards(startISO),

    fcViewMode: 'single', // 'single' (study-first!) | 'list'
    fcScope: 'all', // 'all' | 'pending' | 'day-X' | 'ch-X'
    fcIndex: 0,
    fcListPage: 0,
    customRandomDeckIds: null,
    customRandomLabel: '',
    randomQueue: [],
    randomCursor: 0,
    selectedRandCount: '5',

    // Docs Read-First + Chapter Pagination (Zero lag even for 500+ Qs!)
    docsSelectedChapter: initialChapters[0]?.chapterKey || 'all',
    docsPage: 0,
    docsEditMode: false,
    editingSingleRowId: null,

    // Appearance & Drive
    appearance: { ...DEFAULT_APPEARANCE, ...(saved?.appearance || {}) },
    driveFiles: [],
    driveBannerState: 'loading', // 'loading' | 'synced' | 'error'
    driveBannerMsg: '☁ 正在自動連線 Google Drive / flashcard 資料夾...',
  };
}

const state = createInitialState();

// DOM references
const els = {
  html: document.documentElement,
  body: document.body,
  canvas: document.getElementById('s9-wallpaper-canvas'),
  fileInput: document.getElementById('docx-file-input'),
  globalDropVeil: document.getElementById('global-drop-veil'),
  toastPill: document.getElementById('toast-pill'),
  toastText: document.getElementById('toast-text'),

  // Nav & Topbar
  navBtns: document.querySelectorAll('[data-nav]'),
  sidebarSubLabel: document.getElementById('sidebar-sub-label'),
  topbarBreadcrumb: document.getElementById('topbar-breadcrumb'),
  sidebarDriveDot: document.getElementById('sidebar-drive-dot'),
  btnSidebarRandom: document.getElementById('btn-sidebar-random'),
  paletteDots: document.querySelectorAll('[data-set-palette]'),
  btnS9ModeToggle: document.getElementById('btn-s9-mode-toggle'),
  globalSearchInput: document.getElementById('global-search-input'),
  driveBannerStrip: document.getElementById('drive-banner-strip'),
  driveBannerText: document.getElementById('drive-banner-text'),
  btnDriveQuickReload: document.getElementById('btn-drive-quick-reload'),

  // Views
  viewHome: document.getElementById('view-home'),
  viewFlashcards: document.getElementById('view-flashcards'),
  viewDocs: document.getElementById('view-docs'),
  viewCloud: document.getElementById('view-cloud'),

  // Home
  pageGreeting: document.getElementById('page-greeting'),
  activeFileCaption: document.getElementById('active-file-caption'),
  heroDayLabel: document.getElementById('hero-day-label'),
  btnDayMinus: document.getElementById('btn-day-minus'),
  btnDayPlus: document.getElementById('btn-day-plus'),
  heroShortDate: document.getElementById('hero-short-date'),
  heroClock: document.getElementById('hero-clock'),
  btnHeroResume: document.getElementById('btn-hero-resume'),
  btnHeroRandomQ: document.getElementById('btn-hero-random-q'),
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
  tileStudyValue: document.getElementById('tile-study-value'),
  tileDocsValue: document.getElementById('tile-docs-value'),
  chartModeBtns: document.querySelectorAll('[data-chart-mode]'),
  homeBarChart: document.getElementById('home-bar-chart'),
  homeSegmentedBar: document.getElementById('home-segmented-bar'),
  homeChapterBreakdown: document.getElementById('home-chapter-breakdown'),

  // Flashcards & Random Sheet
  btnOpenRandomSheet: document.getElementById('btn-open-random-sheet'),
  btnFcModeToggle: document.getElementById('btn-fc-mode-toggle'),
  btnFcClearRandom: document.getElementById('btn-fc-clear-random'),
  fcCounterPill: document.getElementById('fc-counter-pill'),
  fcScopeRail: document.getElementById('fc-scope-rail'),
  fcStudyStage: document.getElementById('fc-study-stage'),
  randomDeckSheet: document.getElementById('random-deck-sheet'),
  btnCloseRandomSheet: document.getElementById('btn-close-random-sheet'),
  randSourceSelect: document.getElementById('rand-source-select'),
  randCountPills: document.querySelectorAll('[data-rand-count]'),
  btnSheetRandomChapter: document.getElementById('btn-sheet-random-chapter'),
  btnSheetStartRandom: document.getElementById('btn-sheet-start-random'),

  // Docs
  btnDocsEditToggle: document.getElementById('btn-docs-edit-toggle'),
  btnDocsPaperMode: document.getElementById('btn-docs-paper-mode'),
  btnDocsAddAllCh: document.getElementById('btn-docs-add-all-ch'),
  btnDocsSyncNow: document.getElementById('btn-docs-sync-now'),
  docsChapterRail: document.getElementById('docs-chapter-rail'),
  docsEditToolbar: document.getElementById('docs-edit-toolbar'),
  btnFmtHighlight: document.getElementById('btn-fmt-highlight'),
  btnDocsAddRow: document.getElementById('btn-docs-add-row'),
  btnDocsExport: document.getElementById('btn-docs-export'),
  chkWritebackSort: document.getElementById('chk-writeback-sort'),
  docsPaper: document.getElementById('docs-paper'),
  docsChapterHeading: document.getElementById('docs-chapter-heading'),
  docsChapterCount: document.getElementById('docs-chapter-count'),
  docsRowsContainer: document.getElementById('docs-rows-container'),
  docsPaginationBar: document.getElementById('docs-pagination-bar'),

  // Cloud & Settings
  inputAppsScriptUrl: document.getElementById('input-apps-script-url'),
  btnCloudReconnect: document.getElementById('btn-cloud-reconnect'),
  btnCloudUploadLocal: document.getElementById('btn-cloud-upload-local'),
  btnCloudPush: document.getElementById('btn-cloud-push'),
  btnDriveBrowseAll: document.getElementById('btn-drive-browse-all'),
  driveFilesContainer: document.getElementById('drive-files-container'),
  settingsS9Checkbox: document.getElementById('settings-s9-checkbox'),
  settingsCustomAccent: document.getElementById('settings-custom-accent'),
  settingsRadiusSlider: document.getElementById('settings-radius-slider'),
  settingsRadiusVal: document.getElementById('settings-radius-val'),
  settingsKSlider: document.getElementById('settings-k-slider'),
  settingsKVal: document.getElementById('settings-k-val'),
  wpPresetBtns: document.querySelectorAll('[data-wp-preset]'),
};

function init() {
  wallpaper = new GalaxyS9Wallpaper(els.canvas);
  wallpaper.start();

  els.inputAppsScriptUrl.value = driveClient.appsScriptUrl;

  applyAppearance();
  startLiveClock();
  bindEvents();
  renderAll();

  // Auto-fetch from Google Drive / flashcard immediately on load
  autoConnectDriveAndLoad();
}

/* ==========================================================================
   Lightweight Persistence (NEVER freezes iPad with 5MB DOM dumps!)
   ========================================================================== */

function persistMeta() {
  try {
    localStorage.setItem(
      META_STORAGE_KEY,
      JSON.stringify({
        currentDay: state.currentDay,
        studyStartDate: state.studyStartDate,
        lastDriveFileId: state.doc.driveFileId || state.lastDriveFileId,
        flashcardsMap: state.flashcardsMap,
        appearance: state.appearance,
      })
    );
  } catch {
    // Safe ignore
  }
}

function getCloudStatePayload() {
  return {
    version: 4,
    savedAt: new Date().toISOString(),
    currentDay: state.currentDay,
    studyStartDate: state.studyStartDate,
    lastDriveFileId: state.doc.driveFileId,
    flashcardsMap: state.flashcardsMap,
    appearance: state.appearance,
  };
}

/* ==========================================================================
   Google Drive Auto-Load & 2.5s Debounced In-Place OOXML Writeback
   ========================================================================== */

function setDriveBanner(status, msg) {
  state.driveBannerState = status;
  state.driveBannerMsg = msg;
  els.driveBannerStrip.className = `drive-banner-strip sq-md is-${status}`;
  els.driveBannerText.textContent = msg;
  els.sidebarDriveDot.classList.toggle('connected', status === 'synced');
}

async function autoConnectDriveAndLoad(allDrive = false) {
  if (!driveClient.isConnected()) {
    setDriveBanner('error', '未設定 Google Drive 橋接網址，請前往「雲端/設定」貼上 Apps Script URL');
    return;
  }

  setDriveBanner('loading', '☁ 正在讀取 Google Drive / flashcard 資料夾...');

  try {
    await driveClient.ensureFlashcardFolder();
    const [files, cloudState] = await Promise.all([
      driveClient.listDocuments(allDrive),
      driveClient.loadWorkspaceState().catch(() => null),
    ]);

    state.driveFiles = files || [];

    if (cloudState) {
      if (cloudState.flashcardsMap) state.flashcardsMap = cloudState.flashcardsMap;
      if (cloudState.currentDay) state.currentDay = cloudState.currentDay;
      if (cloudState.appearance) {
        state.appearance = { ...DEFAULT_APPEARANCE, ...cloudState.appearance };
        applyAppearance();
      }
      if (cloudState.lastDriveFileId) state.lastDriveFileId = cloudState.lastDriveFileId;
    }

    if (state.driveFiles.length > 0) {
      const target =
        state.driveFiles.find((f) => f.id === (state.doc.driveFileId || state.lastDriveFileId)) ||
        state.driveFiles[0];
      await openDriveFile(target);
    } else {
      setDriveBanner('synced', '✓ 已連線 Google Drive / flashcard（目前資料夾內尚無 .docx，使用內建範例）');
      renderAll();
    }
  } catch (err) {
    setDriveBanner('error', `⚠ Google Drive 連線失敗：${err.message}`);
    renderCloudView();
  }
}

async function openDriveFile(fileMeta) {
  setDriveBanner('loading', `☁ 正在下載並解析 ${fileMeta.name}...`);
  const buf = await driveClient.downloadDocxArrayBuffer(fileMeta.id, fileMeta.mimeType);
  const parsed = await parseDocxDocument(buf, fileMeta.name);
  if (!parsed.rows.length) {
    setDriveBanner('error', `⚠ ${fileMeta.name} 中未偵測到雙欄問答表格或 Q&A 內容`);
    return;
  }

  rawDocxBuffer = parsed.rawBuffer;
  state.doc = {
    title: parsed.title,
    fileName: parsed.fileName,
    rows: parsed.rows,
    driveFileId: fileMeta.id,
  };
  state.lastDriveFileId = fileMeta.id;

  const chapters = groupRowsByChapter(parsed.rows);
  state.docsSelectedChapter = chapters[0]?.chapterKey || 'all';
  state.docsPage = 0;
  state.fcIndex = 0;

  persistMeta();
  setDriveBanner(
    'synced',
    `✓ 已連線 Drive · ${fileMeta.name}（${parsed.rows.length} 題 · ${chapters.length} 章）`
  );
  renderAll();
}

function scheduleDriveAutoSync() {
  persistMeta();
  if (!driveClient.isConnected()) return;
  setDriveBanner('loading', '✏️ 編輯中… 停筆 2.5 秒後自動寫回 Google Drive 原檔');
  clearTimeout(autoSyncTimer);
  autoSyncTimer = setTimeout(() => syncToDrive(false), 2500);
}

async function syncToDrive(showAlert = true) {
  if (!driveClient.isConnected()) {
    if (showAlert) showToast('請先在「雲端/設定」連線 Google Drive');
    return;
  }
  clearTimeout(autoSyncTimer);
  setDriveBanner('loading', '☁ 正在寫回 Google Drive 原檔與閃卡進度...');
  try {
    const blob = await buildUpdatedDocxBlob(state.doc, rawDocxBuffer, {
      sortByChapter: state.appearance.writebackSortByChapter,
    });
    const saved = await driveClient.saveDocxToDrive(state.doc.fileName, blob, state.doc.driveFileId);
    if (saved?.id) state.doc.driveFileId = saved.id;
    rawDocxBuffer = await blob.arrayBuffer();
    state.doc.rows.forEach((r) => {
      r.dirty = false;
    });
    await driveClient.saveWorkspaceState(getCloudStatePayload());
    persistMeta();

    const t = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    setDriveBanner('synced', `✓ 已同步寫回 Google Drive 原檔 (${t})`);
    if (showAlert) showToast('✓ 已寫回 Google Drive 原檔與閃卡進度！');
  } catch (err) {
    setDriveBanner('error', `⚠ 寫回 Drive 失敗：${err.message}`);
  }
}

/* ==========================================================================
   Appearance & Live Clock
   ========================================================================== */

function applyAppearance() {
  const a = state.appearance;
  els.html.setAttribute('data-palette', a.palette);
  if (a.customAccentHex) {
    els.html.style.setProperty('--accent', a.customAccentHex);
    els.html.style.setProperty('--accent-deep', a.customAccentHex);
    els.html.style.setProperty('--accent-soft', hexToRgba(a.customAccentHex, 0.14));
    els.html.style.setProperty('--accent-glow', hexToRgba(a.customAccentHex, 0.28));
  }
  els.html.style.setProperty('--sq-radius', `${a.sqRadius}px`);
  els.html.style.setProperty('--sq-k', String(a.sqCurvature));

  els.body.classList.toggle('s9-wallpaper-mode', Boolean(a.s9Mode));
  els.btnS9ModeToggle.classList.toggle('active', Boolean(a.s9Mode));
  els.settingsS9Checkbox.checked = Boolean(a.s9Mode);

  if (wallpaper) {
    wallpaper.setAccentHex(a.customAccentHex || PALETTE_ACCENTS[a.palette] || '#D96F52');
    wallpaper.setPreset(a.wpPreset);
    wallpaper.setSpeed(a.wpSpeed);
    wallpaper.setPaused(!a.s9Mode);
  }

  els.paletteDots.forEach((d) => d.classList.toggle('active', d.dataset.setPalette === a.palette));
  els.settingsCustomAccent.value = a.customAccentHex;
  els.settingsRadiusSlider.value = String(a.sqRadius);
  els.settingsRadiusVal.textContent = `${a.sqRadius}px`;
  els.settingsKSlider.value = String(a.sqCurvature);
  els.settingsKVal.textContent = Number(a.sqCurvature).toFixed(2);
  els.wpPresetBtns.forEach((b) => b.classList.toggle('active', b.dataset.wpPreset === a.wpPreset));

  els.docsPaper.classList.toggle('paper-light', a.docsPaper === 'light');
  els.btnDocsPaperMode.textContent = a.docsPaper === 'dark' ? '🌙 深色紙張' : '☀ 淺色紙張';
  els.chkWritebackSort.checked = Boolean(a.writebackSortByChapter);
}

function startLiveClock() {
  const tick = () => {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');
    els.heroClock.textContent = `${hh}:${mm}:${ss}`;
    els.heroShortDate.textContent = now.toLocaleDateString('en-GB', {
      weekday: 'short',
      day: 'numeric',
      month: 'short',
    });

    const hr = now.getHours();
    els.pageGreeting.textContent = hr < 12 ? '早安，Wesley' : hr < 18 ? '午安，Wesley' : '晚安，Wesley';

    const secA = (now.getSeconds() / 60) * Math.PI * 2 - Math.PI / 2;
    const minA = (now.getMinutes() / 60) * Math.PI * 2 - Math.PI / 2;
    const sx = 80 + Math.cos(secA) * 64;
    const sy = 80 + Math.sin(secA) * 64;
    els.orbitalDotSec.setAttribute('cx', sx.toFixed(1));
    els.orbitalDotSec.setAttribute('cy', sy.toFixed(1));
    els.orbitalDotMin.setAttribute('cx', (80 + Math.cos(minA) * 46).toFixed(1));
    els.orbitalDotMin.setAttribute('cy', (80 + Math.sin(minA) * 46).toFixed(1));
    const ax1 = 80 + Math.cos(secA - 1.1) * 64;
    const ay1 = 80 + Math.sin(secA - 1.1) * 64;
    els.orbitalArc.setAttribute('d', `M ${ax1.toFixed(1)} ${ay1.toFixed(1)} A 64 64 0 0 1 ${sx.toFixed(1)} ${sy.toFixed(1)}`);
  };
  tick();
  clearInterval(clockTimer);
  clockTimer = setInterval(tick, 1000);
}

/* ==========================================================================
   Data Helpers
   ========================================================================== */

function getRowById(id) {
  return state.doc.rows.find((r) => r.id === id);
}

function getAddedFlashcards() {
  return state.doc.rows
    .filter((r) => Boolean(state.flashcardsMap[r.id]))
    .map((row) => ({ row, meta: state.flashcardsMap[row.id] }));
}

function shuffleArray(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function addRowToFlashcards(rowId, day = state.currentDay) {
  const row = getRowById(rowId);
  if (!row || state.flashcardsMap[rowId]) return;
  state.flashcardsMap[rowId] = {
    rowId,
    addedDay: day,
    addedDate: toLocalISO(),
    mastered: false,
    checkedPoints: [],
  };
  persistMeta();
  renderAll();
  showToast(`✓ 已加入 Q${row.index}（Day ${day}）`);
  if (driveClient.isConnected()) {
    driveClient.saveWorkspaceState(getCloudStatePayload()).catch(() => {});
  }
}

function removeRowFromFlashcards(rowId) {
  if (!state.flashcardsMap[rowId]) return;
  delete state.flashcardsMap[rowId];
  persistMeta();
  renderAll();
  showToast('已從閃卡清單移除');
}

function getFlashcardCandidates() {
  const q = state.searchQuery.trim().toLowerCase();
  let list = [];

  if (state.customRandomDeckIds?.length) {
    list = state.customRandomDeckIds
      .map((id) => getRowById(id))
      .filter(Boolean)
      .map((row) => ({
        row,
        meta: state.flashcardsMap[row.id] || { rowId: row.id, addedDay: null, mastered: false, checkedPoints: [] },
      }));
  } else {
    const added = getAddedFlashcards();
    // If user hasn't added any cards from a newly loaded Drive docx yet, gracefully show all doc rows so Flashcards is immediately usable!
    const basePool =
      added.length > 0
        ? added
        : state.doc.rows.map((row) => ({
            row,
            meta: { rowId: row.id, addedDay: null, mastered: false, checkedPoints: [] },
          }));

    list = basePool.filter(({ row, meta }) => {
      if (state.fcScope === 'pending') return !meta.mastered;
      if (state.fcScope.startsWith('day-')) return meta.addedDay === Number(state.fcScope.replace('day-', ''));
      if (state.fcScope.startsWith('ch-')) return row.chapterKey === state.fcScope;
      return true;
    });
  }

  if (!q) return list;
  return list.filter(
    ({ row }) =>
      row.question.toLowerCase().includes(q) ||
      row.answerRaw.toLowerCase().includes(q) ||
      (row.chapterLabel || '').toLowerCase().includes(q)
  );
}

/* ==========================================================================
   Master Render
   ========================================================================== */

function renderAll() {
  const labels = { home: 'Home', flashcards: 'Flashcards', docs: 'Docs 原檔', cloud: '雲端與設定' };
  els.topbarBreadcrumb.textContent = labels[state.activeNav] || 'Home';
  els.sidebarSubLabel.textContent = labels[state.activeNav] || 'Home';
  els.activeFileCaption.textContent = state.doc.fileName;

  els.navBtns.forEach((b) => b.classList.toggle('active', b.dataset.nav === state.activeNav));
  els.viewHome.classList.toggle('active', state.activeNav === 'home');
  els.viewFlashcards.classList.toggle('active', state.activeNav === 'flashcards');
  els.viewDocs.classList.toggle('active', state.activeNav === 'docs');
  els.viewCloud.classList.toggle('active', state.activeNav === 'cloud');

  if (state.activeNav === 'home') renderHomeView();
  if (state.activeNav === 'flashcards') renderFlashcardsView();
  if (state.activeNav === 'docs') renderDocsView();
  if (state.activeNav === 'cloud') renderCloudView();
}

/* ==========================================================================
   VIEW 1: HOME
   ========================================================================== */

function renderHomeView() {
  const added = getAddedFlashcards();
  const chapters = groupRowsByChapter(state.doc.rows);
  const mastered = added.filter((x) => x.meta.mastered).length;

  els.heroDayLabel.textContent = `Day ${state.currentDay}`;
  els.heroMasteryTitle.textContent = `已掌握 ${mastered} / ${added.length} 張`;
  els.heroMasterySub.textContent = `原檔共 ${state.doc.rows.length} 題`;

  els.tileFcValue.textContent = `${added.length} Cards`;
  els.tileStudyValue.textContent = /chem/i.test(state.doc.fileName)
    ? 'Chemistry'
    : /phys/i.test(state.doc.fileName)
      ? 'Physics'
      : 'Biology';
  els.tileDocsValue.textContent = `${chapters.length} Chapters`;

  els.chartModeBtns.forEach((b) => b.classList.toggle('active', b.dataset.chartMode === state.chartMode));

  let bars = [];
  if (state.chartMode === 'day') {
    const maxD = Math.max(state.currentDay, 5);
    for (let d = 1; d <= Math.min(maxD + 1, 7); d++) {
      bars.push({
        label: `D${d}`,
        count: added.filter((x) => x.meta.addedDay === d).length,
        hl: d === state.currentDay,
        scope: `day-${d}`,
      });
    }
  } else {
    bars = chapters.slice(0, 7).map((ch, idx) => ({
      label: ch.chapterShort,
      count: ch.rows.length,
      hl: idx === 0,
      scope: ch.chapterKey,
    }));
  }

  const maxB = Math.max(...bars.map((b) => b.count), 1);
  els.homeBarChart.innerHTML = bars
    .map((b) => {
      const pct = b.count === 0 ? 16 : Math.max(24, Math.round((b.count / maxB) * 100));
      return `
        <div class="capsule-bar-col ${b.hl ? 'is-highlight' : ''}" data-home-bar-scope="${escapeAttr(b.scope)}">
          <span style="font-size: 0.72rem; font-weight: 700; color: var(--ink-secondary)">${b.count}</span>
          <div class="capsule-bar-pill" style="height: ${pct}%"></div>
          <span style="font-size: 0.72rem; color: var(--ink-muted)">${escapeHtml(b.label)}</span>
        </div>
      `;
    })
    .join('');

  els.homeSegmentedBar.innerHTML = chapters
    .slice(0, 10)
    .map((ch, i) => `<div class="segmented-hbar-seg" style="flex-grow: ${ch.rows.length}; background: ${BREAKDOWN_COLORS[i % BREAKDOWN_COLORS.length]}"></div>`)
    .join('');

  els.homeChapterBreakdown.innerHTML = chapters
    .map((ch, i) => {
      const color = BREAKDOWN_COLORS[i % BREAKDOWN_COLORS.length];
      const addedInCh = ch.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length;
      const pct = Math.round((addedInCh / Math.max(ch.rows.length, 1)) * 100);
      return `
        <div class="chapter-break-row" data-home-jump-ch="${escapeAttr(ch.chapterKey)}">
          <span style="width: 8px; height: 8px; border-radius: 50%; background: ${color}"></span>
          <strong style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis">${escapeHtml(ch.chapterShort)}</strong>
          <div class="break-track"><div class="break-fill" style="width: ${Math.max(pct, addedInCh ? 12 : 4)}%; background: ${color}"></div></div>
          <span style="font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: var(--ink-secondary)">${addedInCh}/${ch.rows.length}</span>
        </div>
      `;
    })
    .join('');
}

/* ==========================================================================
   VIEW 2: STUDY-FIRST SIMULTANEOUS FLASHCARDS (Swipeable + Big Controls)
   ========================================================================== */

function renderFlashcardsView() {
  const chapters = groupRowsByChapter(state.doc.rows);
  const added = getAddedFlashcards();
  const days = [...new Set(added.map((x) => x.meta.addedDay))].sort((a, b) => a - b);

  // Horizontal Scrollable Scope Pills
  const scopes = [
    { key: 'all', label: added.length > 0 ? `已收錄 (${added.length})` : `全部題目 (${state.doc.rows.length})` },
    { key: 'pending', label: '待複習' },
    ...days.map((d) => ({ key: `day-${d}`, label: `Day ${d}` })),
    ...chapters.map((c) => ({ key: c.chapterKey, label: c.chapterShort })),
  ];

  els.fcScopeRail.innerHTML = scopes
    .map(
      (s) =>
        `<button class="rail-chip ${!state.customRandomDeckIds && state.fcScope === s.key ? 'active' : ''}" data-fc-scope="${escapeAttr(s.key)}" type="button">${escapeHtml(s.label)}</button>`
    )
    .join('');

  els.btnFcClearRandom.style.display = state.customRandomDeckIds ? 'inline-flex' : 'none';
  els.btnFcModeToggle.textContent = state.fcViewMode === 'single' ? '📋 切換列表總覽' : '🗂 切回單卡專注';

  const list = getFlashcardCandidates();
  if (list.length === 0) {
    els.fcCounterPill.textContent = '0 / 0';
    els.fcStudyStage.innerHTML = `
      <div class="surface-card sq-card" style="text-align: center; padding: 44px 20px">
        <strong style="font-size: 1.05rem">這個分類目前沒有題目</strong>
        <p style="color: var(--ink-secondary); margin: 6px 0 18px">切換其他章節，或前往 Docs 原檔收錄新題目。</p>
        <button class="btn-soft accent" data-nav-jump="docs" type="button">前往 Docs 原檔</button>
      </div>
    `;
    return;
  }

  if (state.fcIndex >= list.length) state.fcIndex = 0;

  if (state.fcViewMode === 'single') {
    const active = list[state.fcIndex];
    els.fcCounterPill.textContent = state.customRandomLabel
      ? `🎲 ${state.fcIndex + 1} / ${list.length}`
      : `${state.fcIndex + 1} / ${list.length}`;

    els.fcStudyStage.innerHTML = `
      ${buildStudyCardHtml(active)}
      <div class="study-action-bar" style="margin-top: 14px">
        <button class="btn-soft" id="btn-fc-prev" type="button">← 上一題</button>
        <button class="btn-soft ${active.meta.mastered ? 'accent' : ''}" id="btn-fc-toggle-master" data-id="${escapeAttr(active.row.id)}" type="button">
          ${active.meta.mastered ? '✓ 已掌握' : '標記掌握'}
        </button>
        <button class="btn-soft dark-ink" id="btn-fc-rand-step" type="button">🎲 隨機題目</button>
        <button class="btn-soft accent" id="btn-fc-next" type="button">下一題 →</button>
      </div>
      <div style="text-align: center; font-size: 0.78rem; color: var(--ink-muted); margin-top: 10px">
        左右滑動卡片亦可快速切換上一題 / 下一題
      </div>
    `;

    attachCardSwipe();
  } else {
    // Paginated List View (20 cards max per page = 60fps smooth!)
    const totalPages = Math.max(1, Math.ceil(list.length / FC_LIST_PAGE_SIZE));
    if (state.fcListPage >= totalPages) state.fcListPage = 0;
    const slice = list.slice(state.fcListPage * FC_LIST_PAGE_SIZE, (state.fcListPage + 1) * FC_LIST_PAGE_SIZE);
    els.fcCounterPill.textContent = `共 ${list.length} 題`;

    els.fcStudyStage.innerHTML = `
      <div style="display: flex; flex-direction: column; gap: 16px">
        ${slice.map((item) => buildStudyCardHtml(item)).join('')}
      </div>
      ${
        totalPages > 1
          ? `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 16px">
          <button class="btn-soft sm" id="btn-fc-list-prev" ${state.fcListPage === 0 ? 'disabled' : ''}>← 上一頁</button>
          <span style="font-size: 0.84rem; font-weight: 700">第 ${state.fcListPage + 1} / ${totalPages} 頁</span>
          <button class="btn-soft sm" id="btn-fc-list-next" ${state.fcListPage >= totalPages - 1 ? 'disabled' : ''}>下一頁 →</button>
        </div>
      `
          : ''
      }
    `;
  }
}

function buildStudyCardHtml({ row, meta }) {
  const isAdded = Boolean(state.flashcardsMap[row.id]);
  const points =
    row.answerPoints && row.answerPoints.length > 0
      ? row.answerPoints
      : [{ text: row.answerRaw || '', marks: null }];
  const checkedSet = new Set(meta.checkedPoints || []);

  return `
    <article class="study-card sq-card ${meta.mastered ? 'is-mastered' : ''}" data-card-id="${escapeAttr(row.id)}">
      <div class="study-card-head">
        <div class="badge-group">
          <span class="pill-badge" style="font-family: 'JetBrains Mono', monospace">Q${row.index}</span>
          <span class="pill-badge chapter">${escapeHtml(row.chapterShort)}</span>
          ${
            isAdded
              ? `<span class="pill-badge accent">✓ Day ${meta.addedDay}</span>`
              : `<button class="btn-soft sm" data-fc-add-now="${escapeAttr(row.id)}" type="button">＋ 加入 Day ${state.currentDay}</button>`
          }
          ${row.marks != null ? `<span class="pill-badge marks">${row.marks}m</span>` : ''}
        </div>
        <button class="btn-soft sm" data-locate-doc="${escapeAttr(row.id)}" type="button">📄 原檔</button>
      </div>

      <div class="study-qa-grid">
        <div class="study-q-pane sq-md">
          <span class="pane-label">QUESTION · 問題</span>
          <p class="study-q-text">${highlightText(row.question, state.searchQuery)}</p>
          ${row.source ? `<div style="margin-top: auto; font-size: 0.78rem; color: var(--ink-muted)">*${escapeHtml(row.source)}</div>` : ''}
        </div>

        <div class="study-a-pane sq-md">
          <span class="pane-label">ANSWER · 答案與給分點（同步顯示）</span>
          <ul class="study-a-list">
            ${points
              .map((pt, pIdx) => {
                const checked = checkedSet.has(pIdx);
                return `
                  <li class="study-a-item sq-sm ${checked ? 'checked' : ''}" data-fc-pt="${escapeAttr(row.id)}" data-pt-idx="${pIdx}">
                    <span style="display: flex; align-items: flex-start; gap: 10px">
                      <span class="pt-box">✓</span>
                      <span>${highlightText(pt.text, state.searchQuery)}</span>
                    </span>
                    ${pt.marks != null ? `<span class="pill-badge marks" style="padding: 2px 8px">(${pt.marks})</span>` : ''}
                  </li>
                `;
              })
              .join('')}
          </ul>
        </div>
      </div>
    </article>
  `;
}

function attachCardSwipe() {
  const card = els.fcStudyStage.querySelector('.study-card');
  if (!card) return;
  let startX = 0;
  let startY = 0;
  card.addEventListener(
    'touchstart',
    (e) => {
      startX = e.touches[0].clientX;
      startY = e.touches[0].clientY;
    },
    { passive: true }
  );
  card.addEventListener(
    'touchend',
    (e) => {
      const dx = e.changedTouches[0].clientX - startX;
      const dy = e.changedTouches[0].clientY - startY;
      if (Math.abs(dx) > 56 && Math.abs(dx) > Math.abs(dy) * 1.4) {
        const list = getFlashcardCandidates();
        if (!list.length) return;
        state.fcIndex = (state.fcIndex + (dx < 0 ? 1 : -1) + list.length) % list.length;
        renderFlashcardsView();
      }
    },
    { passive: true }
  );
}

/* ==========================================================================
   VIEW 3: DOCS READER & EXPLICIT EDIT MODE (1 Chapter at a Time = 60fps!)
   ========================================================================== */

function renderDocsView() {
  const chapters = groupRowsByChapter(state.doc.rows);
  if (!chapters.some((c) => c.chapterKey === state.docsSelectedChapter)) {
    state.docsSelectedChapter = chapters[0]?.chapterKey || 'all';
  }

  // Horizontal scrollable Chapter pills
  els.docsChapterRail.innerHTML = chapters
    .map((ch) => {
      const added = ch.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length;
      const active = ch.chapterKey === state.docsSelectedChapter;
      return `
        <button class="rail-chip ${active ? 'active' : ''}" data-doc-ch="${escapeAttr(ch.chapterKey)}" type="button">
          ${escapeHtml(ch.chapterShort)} (${added}/${ch.rows.length})
        </button>
      `;
    })
    .join('');

  els.btnDocsEditToggle.textContent = state.docsEditMode ? '✓ 完成編輯' : '✏️ 開啟編輯模式';
  els.btnDocsEditToggle.classList.toggle('accent', state.docsEditMode);
  els.docsEditToolbar.classList.toggle('active', state.docsEditMode || Boolean(state.editingSingleRowId));

  const currentGroup = chapters.find((c) => c.chapterKey === state.docsSelectedChapter) || chapters[0];
  if (!currentGroup) {
    els.docsRowsContainer.innerHTML = `<p style="opacity: 0.7">尚無內容</p>`;
    return;
  }

  const q = state.searchQuery.trim().toLowerCase();
  const filteredRows = currentGroup.rows.filter((r) => {
    if (!q) return true;
    return r.questionRaw.toLowerCase().includes(q) || r.answerRaw.toLowerCase().includes(q);
  });

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / DOCS_PAGE_SIZE));
  if (state.docsPage >= totalPages) state.docsPage = 0;
  const pageRows = filteredRows.slice(state.docsPage * DOCS_PAGE_SIZE, (state.docsPage + 1) * DOCS_PAGE_SIZE);

  const addedInCh = currentGroup.rows.filter((r) => Boolean(state.flashcardsMap[r.id])).length;
  els.docsChapterHeading.textContent = currentGroup.chapterLabel;
  els.docsChapterCount.textContent = `已收錄 ${addedInCh} / ${currentGroup.rows.length} 題`;
  els.btnDocsAddAllCh.textContent = `＋ 本章加入 Day ${state.currentDay}`;

  els.docsRowsContainer.innerHTML = pageRows
    .map((row) => {
      const meta = state.flashcardsMap[row.id];
      const isAdded = Boolean(meta);
      const isEditable = state.docsEditMode || state.editingSingleRowId === row.id;

      return `
        <div class="doc-row-card ${isAdded ? 'is-added' : ''} ${isEditable ? 'is-editing' : ''}" id="doc-row-${escapeAttr(row.id)}">
          <div class="doc-row-cells">
            <div class="doc-cell">
              <div
                class="doc-cell-content"
                contenteditable="${isEditable ? 'true' : 'false'}"
                spellcheck="false"
                data-edit-cell="q"
                data-row-id="${escapeAttr(row.id)}"
              >${row.questionHtml}</div>
            </div>
            <div class="doc-cell">
              <div
                class="doc-cell-content"
                contenteditable="${isEditable ? 'true' : 'false'}"
                spellcheck="false"
                data-edit-cell="a"
                data-row-id="${escapeAttr(row.id)}"
              >${row.answerHtml}</div>
            </div>
          </div>

          <div class="doc-row-footer">
            <span style="font-family: 'JetBrains Mono', monospace; font-size: 0.78rem; opacity: 0.72">
              #${row.index} · ${escapeHtml(row.chapterShort)}
            </span>
            <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap">
              <button class="btn-soft sm" data-doc-edit-row="${escapeAttr(row.id)}" type="button">
                ${isEditable ? '✓ 完成' : '✏️ 編輯'}
              </button>
              ${
                isAdded
                  ? `
                    <button class="btn-soft accent sm" data-doc-unadd="${escapeAttr(row.id)}" type="button" title="點擊可取消加入">
                      ✓ 已加入 · Day ${meta.addedDay}
                    </button>
                  `
                  : `
                    <button class="btn-soft sm" data-doc-add="${escapeAttr(row.id)}" type="button">
                      ＋ 加入閃卡 (Day ${state.currentDay})
                    </button>
                  `
              }
            </div>
          </div>
        </div>
        ${row.pageBreakAfter ? `<div class="doc-page-break">Page break</div>` : ''}
      `;
    })
    .join('');

  els.docsPaginationBar.innerHTML =
    totalPages > 1
      ? `
        <button class="btn-soft sm" id="btn-docs-prev-page" ${state.docsPage === 0 ? 'disabled' : ''}>← 上一頁</button>
        <span style="font-size: 0.84rem; font-weight: 700">第 ${state.docsPage + 1} / ${totalPages} 頁</span>
        <button class="btn-soft sm" id="btn-docs-next-page" ${state.docsPage >= totalPages - 1 ? 'disabled' : ''}>下一頁 →</button>
      `
      : '';
}

/* ==========================================================================
   VIEW 4: CLOUD & APPEARANCE
   ========================================================================== */

function renderCloudView() {
  if (!state.driveFiles.length) {
    els.driveFilesContainer.innerHTML = `<div style="padding: 14px; font-size: 0.84rem; color: var(--ink-muted)">尚未列出檔案，可點擊上方「🔄 立即重新連線並讀取」或「看全部 Drive」。</div>`;
    return;
  }

  els.driveFilesContainer.innerHTML = state.driveFiles
    .map(
      (f) => `
      <div class="drive-file-item sq-sm">
        <div style="min-width: 0">
          <strong style="display: block; font-size: 0.86rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis">${escapeHtml(f.name)}</strong>
          <span style="font-size: 0.74rem; color: var(--ink-muted)">${f.id === state.doc.driveFileId ? '✓ 目前已載入' : '.docx / Docs'}</span>
        </div>
        <button
          class="btn-soft accent sm"
          data-open-drive-id="${escapeAttr(f.id)}"
          data-open-drive-name="${escapeAttr(f.name)}"
          data-open-drive-mime="${escapeAttr(f.mimeType || '')}"
          type="button"
        >開啟</button>
      </div>
    `
    )
    .join('');
}

/* ==========================================================================
   Random Sheet & Random Question Engine
   ========================================================================== */

function openRandomSheet() {
  const chapters = groupRowsByChapter(state.doc.rows);
  const added = getAddedFlashcards();
  els.randSourceSelect.innerHTML = `
    <option value="added">已收錄的閃卡 (${added.length} 題)</option>
    <option value="pending">待複習閃卡 (${added.filter((x) => !x.meta.mastered).length} 題)</option>
    <option value="all-doc">整份原檔全部題目 (${state.doc.rows.length} 題)</option>
    ${chapters.map((c) => `<option value="${escapeAttr(c.chapterKey)}">只抽 ${escapeHtml(c.chapterShort)} (${c.rows.length} 題)</option>`).join('')}
  `;
  els.randomDeckSheet.classList.add('open');
}

function startRandomDeck(source, countVal) {
  let pool = [];
  if (source === 'added') pool = state.doc.rows.filter((r) => Boolean(state.flashcardsMap[r.id]));
  else if (source === 'pending') pool = state.doc.rows.filter((r) => state.flashcardsMap[r.id] && !state.flashcardsMap[r.id].mastered);
  else if (source.startsWith('ch-')) pool = state.doc.rows.filter((r) => r.chapterKey === source);
  else pool = [...state.doc.rows];

  if (!pool.length) pool = [...state.doc.rows];
  const shuffled = shuffleArray(pool);
  const picked = countVal === 'all' ? shuffled : shuffled.slice(0, Math.min(Number(countVal) || 5, shuffled.length));

  state.customRandomDeckIds = picked.map((r) => r.id);
  state.customRandomLabel = `隨機 ${picked.length} 題`;
  state.fcIndex = 0;
  state.fcViewMode = 'single';
  state.activeNav = 'flashcards';
  els.randomDeckSheet.classList.remove('open');
  renderAll();
  showToast(`🎲 已隨機抽出 ${picked.length} 張閃卡！`);
}

function stepNoRepeatRandomQuestion() {
  const list = getFlashcardCandidates();
  if (!list.length) return;
  const ids = list.map((x) => x.row.id);

  if (state.randomQueue.length !== ids.length || !state.randomQueue.every((id) => ids.includes(id))) {
    state.randomQueue = shuffleArray(ids);
    state.randomCursor = 0;
  } else {
    state.randomCursor += 1;
    if (state.randomCursor >= state.randomQueue.length) {
      state.randomQueue = shuffleArray(ids);
      state.randomCursor = 0;
      showToast('🎉 本輪全部題目已無重複抽完，重新洗牌！');
    }
  }

  const targetId = state.randomQueue[state.randomCursor];
  const idx = list.findIndex((x) => x.row.id === targetId);
  state.fcIndex = idx >= 0 ? idx : 0;
  state.fcViewMode = 'single';
  renderFlashcardsView();
}

/* ==========================================================================
   Events
   ========================================================================== */

function bindEvents() {
  // Navigation (Desktop Sidebar + iPad Floating Bottom Tab Bar)
  els.navBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      state.activeNav = btn.dataset.nav;
      renderAll();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });

  // Topbar Theme Dots, S9 Toggle, Search, Quick Reload
  els.paletteDots.forEach((dot) => {
    dot.addEventListener('click', () => {
      const p = dot.dataset.setPalette;
      state.appearance.palette = p;
      state.appearance.customAccentHex = PALETTE_ACCENTS[p];
      applyAppearance();
      persistMeta();
    });
  });

  const toggleS9 = () => {
    state.appearance.s9Mode = !state.appearance.s9Mode;
    applyAppearance();
    persistMeta();
  };
  els.btnS9ModeToggle.addEventListener('click', toggleS9);
  els.settingsS9Checkbox.addEventListener('change', (e) => {
    state.appearance.s9Mode = e.target.checked;
    applyAppearance();
    persistMeta();
  });

  els.globalSearchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    state.docsPage = 0;
    state.fcIndex = 0;
    renderAll();
  });

  els.btnDriveQuickReload.addEventListener('click', () => autoConnectDriveAndLoad(false));

  // Home interactions
  els.btnDayMinus.addEventListener('click', () => {
    state.currentDay = Math.max(1, state.currentDay - 1);
    persistMeta();
    renderAll();
  });
  els.btnDayPlus.addEventListener('click', () => {
    state.currentDay += 1;
    persistMeta();
    renderAll();
  });

  els.btnHeroResume.addEventListener('click', () => {
    state.activeNav = 'flashcards';
    renderAll();
  });
  els.btnHeroRandomQ.addEventListener('click', () => {
    state.activeNav = 'flashcards';
    stepNoRepeatRandomQuestion();
    renderAll();
  });

  els.btnSidebarRandom.addEventListener('click', openRandomSheet);
  els.cardHomeRandom.addEventListener('click', openRandomSheet);
  els.cardHomeFlashcards.addEventListener('click', () => {
    state.activeNav = 'flashcards';
    renderAll();
  });
  els.cardHomeStudy.addEventListener('click', () => {
    state.activeNav = 'docs';
    renderAll();
  });
  els.cardHomeDocs.addEventListener('click', () => {
    state.activeNav = 'docs';
    renderAll();
  });

  els.chartModeBtns.forEach((b) => {
    b.addEventListener('click', () => {
      state.chartMode = b.dataset.chartMode;
      renderHomeView();
    });
  });

  els.homeBarChart.addEventListener('click', (e) => {
    const col = e.target.closest('[data-home-bar-scope]');
    if (!col) return;
    state.customRandomDeckIds = null;
    state.fcScope = col.dataset.homeBarScope;
    state.fcIndex = 0;
    state.activeNav = 'flashcards';
    renderAll();
  });

  els.homeChapterBreakdown.addEventListener('click', (e) => {
    const row = e.target.closest('[data-home-jump-ch]');
    if (!row) return;
    state.docsSelectedChapter = row.dataset.homeJumpCh;
    state.docsPage = 0;
    state.activeNav = 'docs';
    renderAll();
  });

  // Random Deck Modal Sheet
  els.btnOpenRandomSheet.addEventListener('click', openRandomSheet);
  els.btnCloseRandomSheet.addEventListener('click', () => els.randomDeckSheet.classList.remove('open'));
  els.randomDeckSheet.addEventListener('click', (e) => {
    if (e.target === els.randomDeckSheet) els.randomDeckSheet.classList.remove('open');
  });

  els.randCountPills.forEach((p) => {
    p.addEventListener('click', () => {
      state.selectedRandCount = p.dataset.randCount;
      els.randCountPills.forEach((x) => x.classList.toggle('active', x === p));
    });
  });

  els.btnSheetStartRandom.addEventListener('click', () => {
    startRandomDeck(els.randSourceSelect.value, state.selectedRandCount);
  });

  els.btnSheetRandomChapter.addEventListener('click', () => {
    const chapters = groupRowsByChapter(state.doc.rows);
    if (!chapters.length) return;
    const ch = chapters[Math.floor(Math.random() * chapters.length)];
    startRandomDeck(ch.chapterKey, 'all');
  });

  // Flashcards Stage Controls
  els.btnFcModeToggle.addEventListener('click', () => {
    state.fcViewMode = state.fcViewMode === 'single' ? 'list' : 'single';
    renderFlashcardsView();
  });

  els.btnFcClearRandom.addEventListener('click', () => {
    state.customRandomDeckIds = null;
    state.customRandomLabel = '';
    state.fcIndex = 0;
    renderFlashcardsView();
  });

  els.fcScopeRail.addEventListener('click', (e) => {
    const chip = e.target.closest('[data-fc-scope]');
    if (!chip) return;
    state.customRandomDeckIds = null;
    state.fcScope = chip.dataset.fcScope;
    state.fcIndex = 0;
    state.fcListPage = 0;
    renderFlashcardsView();
  });

  els.fcStudyStage.addEventListener('click', (e) => {
    if (e.target.closest('#btn-fc-prev')) {
      const list = getFlashcardCandidates();
      state.fcIndex = (state.fcIndex - 1 + list.length) % list.length;
      renderFlashcardsView();
      return;
    }
    if (e.target.closest('#btn-fc-next')) {
      const list = getFlashcardCandidates();
      state.fcIndex = (state.fcIndex + 1) % list.length;
      renderFlashcardsView();
      return;
    }
    if (e.target.closest('#btn-fc-rand-step')) {
      stepNoRepeatRandomQuestion();
      return;
    }
    const masterBtn = e.target.closest('#btn-fc-toggle-master');
    if (masterBtn) {
      const id = masterBtn.dataset.id;
      if (!state.flashcardsMap[id]) addRowToFlashcards(id, state.currentDay);
      state.flashcardsMap[id].mastered = !state.flashcardsMap[id].mastered;
      persistMeta();
      renderFlashcardsView();
      return;
    }
    if (e.target.closest('#btn-fc-list-prev')) {
      state.fcListPage = Math.max(0, state.fcListPage - 1);
      renderFlashcardsView();
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }
    if (e.target.closest('#btn-fc-list-next')) {
      state.fcListPage += 1;
      renderFlashcardsView();
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }
    const navJump = e.target.closest('[data-nav-jump]');
    if (navJump) {
      state.activeNav = navJump.dataset.navJump;
      renderAll();
      return;
    }
    const addNow = e.target.closest('[data-fc-add-now]');
    if (addNow) {
      addRowToFlashcards(addNow.dataset.fcAddNow, state.currentDay);
      return;
    }
    const locate = e.target.closest('[data-locate-doc]');
    if (locate) {
      const row = getRowById(locate.dataset.locateDoc);
      if (row) {
        state.docsSelectedChapter = row.chapterKey;
        state.docsPage = 0;
        state.activeNav = 'docs';
        renderAll();
        requestAnimationFrame(() => {
          document.getElementById(`doc-row-${row.id}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
        });
      }
      return;
    }
    const pt = e.target.closest('[data-fc-pt]');
    if (pt) {
      const id = pt.dataset.fcPt;
      const pIdx = Number(pt.dataset.ptIdx);
      if (!state.flashcardsMap[id]) addRowToFlashcards(id, state.currentDay);
      const meta = state.flashcardsMap[id];
      const set = new Set(meta.checkedPoints || []);
      if (set.has(pIdx)) set.delete(pIdx);
      else set.add(pIdx);
      meta.checkedPoints = [...set];
      persistMeta();
      renderFlashcardsView();
    }
  });

  // Docs Controls
  els.docsChapterRail.addEventListener('click', (e) => {
    const chip = e.target.closest('[data-doc-ch]');
    if (!chip) return;
    state.docsSelectedChapter = chip.dataset.docCh;
    state.docsPage = 0;
    state.editingSingleRowId = null;
    renderDocsView();
  });

  els.btnDocsEditToggle.addEventListener('click', () => {
    state.docsEditMode = !state.docsEditMode;
    state.editingSingleRowId = null;
    renderDocsView();
    if (state.docsEditMode) showToast('已進入編輯模式，點擊文字即可修改並自動寫回 Drive');
  });

  els.btnDocsPaperMode.addEventListener('click', () => {
    state.appearance.docsPaper = state.appearance.docsPaper === 'dark' ? 'light' : 'dark';
    applyAppearance();
    persistMeta();
  });

  els.btnDocsAddAllCh.addEventListener('click', () => {
    const ch = groupRowsByChapter(state.doc.rows).find((g) => g.chapterKey === state.docsSelectedChapter);
    if (!ch) return;
    let cnt = 0;
    for (const r of ch.rows) {
      if (!state.flashcardsMap[r.id]) {
        state.flashcardsMap[r.id] = {
          rowId: r.id,
          addedDay: state.currentDay,
          addedDate: toLocalISO(),
          mastered: false,
          checkedPoints: [],
        };
        cnt++;
      }
    }
    persistMeta();
    renderDocsView();
    showToast(`✓ 已將本章 ${cnt} 題加入 Day ${state.currentDay} 閃卡！`);
  });

  els.btnDocsSyncNow.addEventListener('click', () => syncToDrive(true));
  els.chkWritebackSort.addEventListener('change', (e) => {
    state.appearance.writebackSortByChapter = e.target.checked;
    persistMeta();
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

  els.btnDocsAddRow.addEventListener('click', () => {
    const ch = groupRowsByChapter(state.doc.rows).find((g) => g.chapterKey === state.docsSelectedChapter);
    const newRow = buildDocRow({
      id: `row-new-${Date.now().toString(36)}`,
      index: state.doc.rows.length + 1,
      originalOrder: state.doc.rows.length,
      section: ch?.chapterLabel || 'Chapter 1',
      questionLines: ['請輸入新題目... (2 marks)'],
      answerLines: ['請輸入答案要點... (1)'],
      dirty: true,
    });
    state.doc.rows.push(newRow);
    state.editingSingleRowId = newRow.id;
    renderDocsView();
    scheduleDriveAutoSync();
  });

  els.btnDocsExport.addEventListener('click', async () => {
    const blob = await buildUpdatedDocxBlob(state.doc, rawDocxBuffer, {
      sortByChapter: state.appearance.writebackSortByChapter,
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = state.doc.fileName;
    a.click();
    URL.revokeObjectURL(url);
  });

  els.docsRowsContainer.addEventListener('input', (e) => {
    const cell = e.target.closest('[data-edit-cell]');
    if (!cell) return;
    const rowId = cell.dataset.rowId;
    const idx = state.doc.rows.findIndex((r) => r.id === rowId);
    if (idx === -1) return;
    const card = cell.closest('.doc-row-card');
    state.doc.rows[idx] = updateRowFromHtml(
      state.doc.rows[idx],
      card.querySelector('[data-edit-cell="q"]').innerHTML,
      card.querySelector('[data-edit-cell="a"]').innerHTML
    );
    scheduleDriveAutoSync();
  });

  els.docsRowsContainer.addEventListener('click', (e) => {
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
    const editBtn = e.target.closest('[data-doc-edit-row]');
    if (editBtn) {
      const id = editBtn.dataset.docEditRow;
      state.editingSingleRowId = state.editingSingleRowId === id ? null : id;
      renderDocsView();
    }
  });

  els.docsPaginationBar.addEventListener('click', (e) => {
    if (e.target.closest('#btn-docs-prev-page')) {
      state.docsPage = Math.max(0, state.docsPage - 1);
      renderDocsView();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } else if (e.target.closest('#btn-docs-next-page')) {
      state.docsPage += 1;
      renderDocsView();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  });

  // Cloud & Appearance Controls
  els.btnCloudReconnect.addEventListener('click', () => {
    driveClient.setConfig({ mode: 'apps-script', appsScriptUrl: els.inputAppsScriptUrl.value });
    autoConnectDriveAndLoad(false);
  });
  els.btnDriveBrowseAll.addEventListener('click', () => autoConnectDriveAndLoad(true));
  els.btnCloudPush.addEventListener('click', () => syncToDrive(true));
  els.btnCloudUploadLocal.addEventListener('click', () => els.fileInput.click());

  els.driveFilesContainer.addEventListener('click', async (e) => {
    const btn = e.target.closest('[data-open-drive-id]');
    if (!btn) return;
    try {
      await openDriveFile({
        id: btn.dataset.openDriveId,
        name: btn.dataset.openDriveName,
        mimeType: btn.dataset.openDriveMime,
      });
      state.activeNav = 'docs';
      renderAll();
    } catch (err) {
      showToast(err.message);
    }
  });

  els.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const parsed = await parseDocxDocument(file, file.name);
      rawDocxBuffer = parsed.rawBuffer;
      state.doc = { title: parsed.title, fileName: parsed.fileName, rows: parsed.rows, driveFileId: null };
      state.docsSelectedChapter = groupRowsByChapter(parsed.rows)[0]?.chapterKey || 'all';
      state.docsPage = 0;
      state.activeNav = 'docs';
      renderAll();
      showToast(`✓ 已載入本機檔案 ${file.name}（${parsed.rows.length} 題）`);
    } catch (err) {
      showToast(err.message);
    }
    els.fileInput.value = '';
  });

  els.settingsCustomAccent.addEventListener('input', (e) => {
    state.appearance.customAccentHex = e.target.value;
    applyAppearance();
    persistMeta();
  });
  els.settingsRadiusSlider.addEventListener('input', (e) => {
    state.appearance.sqRadius = Number(e.target.value);
    applyAppearance();
    persistMeta();
  });
  els.settingsKSlider.addEventListener('input', (e) => {
    state.appearance.sqCurvature = Number(e.target.value);
    applyAppearance();
    persistMeta();
  });
  els.wpPresetBtns.forEach((b) => {
    b.addEventListener('click', () => {
      state.appearance.wpPreset = b.dataset.wpPreset;
      applyAppearance();
      persistMeta();
    });
  });

  // Keyboard arrow / R shortcuts on desktop
  window.addEventListener('keydown', (e) => {
    if (e.target.matches('input, textarea, select, [contenteditable="true"]')) return;
    if (state.activeNav === 'flashcards') {
      const list = getFlashcardCandidates();
      if (!list.length) return;
      if (e.key.toLowerCase() === 'r') stepNoRepeatRandomQuestion();
      else if (e.key === 'ArrowRight') {
        state.fcIndex = (state.fcIndex + 1) % list.length;
        renderFlashcardsView();
      } else if (e.key === 'ArrowLeft') {
        state.fcIndex = (state.fcIndex - 1 + list.length) % list.length;
        renderFlashcardsView();
      }
    }
  });
}

function showToast(msg) {
  els.toastText.textContent = msg;
  els.toastPill.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toastPill.classList.remove('show'), 2800);
}

function hexToRgba(hex, alpha) {
  const clean = String(hex || '#D96F52').replace('#', '');
  const full = clean.length === 3 ? clean.split('').map((c) => c + c).join('') : clean;
  const num = parseInt(full, 16);
  return `rgba(${(num >> 16) & 255}, ${(num >> 8) & 255}, ${num & 255}, ${alpha})`;
}

function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeAttr(s) {
  return escapeHtml(s).replace(/'/g, '&#39;');
}

function highlightText(text, q) {
  const safe = escapeHtml(text);
  const query = String(q || '').trim();
  if (!query) return safe;
  const esc = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return safe.replace(new RegExp(`(${esc})`, 'gi'), '<mark class="hl">$1</mark>');
}

init();

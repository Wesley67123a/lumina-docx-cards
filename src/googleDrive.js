/**
 * Google Drive Integration for Lumina Docx Cards
 * - Automatically locates or creates a folder named "flashcard" in the user's Google Drive
 * - Syncs flashcards, Day-added records, checkmarks, layout positions/sizes, and appearance settings
 * - Reads, edits, and saves .docx & Google Docs files directly inside Google Drive / flashcard
 * - Supports both Google Identity Services (OAuth 2.0) AND Zero-Config Apps Script Web App Bridge
 */

const DRIVE_API = 'https://www.googleapis.com/drive/v3';
const DRIVE_UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3';
const STATE_FILENAME = 'lumina_flashcards_state.json';
const FOLDER_NAME = 'flashcard';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const GDOC_MIME = 'application/vnd.google-apps.document';
const FOLDER_MIME = 'application/vnd.google-apps.folder';

export const DEFAULT_APPS_SCRIPT_URL =
  'https://script.google.com/macros/s/AKfycbzpvmaDX1Q_lihjeqxEaSiKhfUbTruKIKGrlcCZ9rd61J7W7YX0jL2aI8LKMmSjEe8/exec';

export class GoogleDriveClient {
  constructor() {
    const savedMode = localStorage.getItem('lumina-drive-mode');
    const savedClientId = localStorage.getItem('lumina-gdrive-client-id') || '';
    const savedScriptUrl = localStorage.getItem('lumina-apps-script-url') || '';

    this.clientId = savedClientId;
    this.appsScriptUrl = savedScriptUrl || DEFAULT_APPS_SCRIPT_URL;
    // Default to apps-script bridge automatically unless user explicitly configured OAuth with a Client ID
    this.mode = savedMode === 'oauth' && savedClientId ? 'oauth' : 'apps-script';
    this.accessToken = sessionStorage.getItem('lumina-gdrive-token') || '';
    this.tokenExpiry = Number(sessionStorage.getItem('lumina-gdrive-token-exp') || '0');
    this.folderId = localStorage.getItem('lumina-gdrive-folder-id') || '';
    this.userEmail = localStorage.getItem('lumina-gdrive-email') || '';
    this.gisLoaded = false;
  }

  isConnected() {
    if (this.mode === 'apps-script') {
      return Boolean(this.appsScriptUrl && this.appsScriptUrl.startsWith('https://script.google.com/'));
    }
    return Boolean(this.accessToken && Date.now() < this.tokenExpiry);
  }

  setConfig({ mode, clientId, appsScriptUrl }) {
    if (mode) {
      this.mode = mode;
      localStorage.setItem('lumina-drive-mode', mode);
    }
    if (clientId !== undefined) {
      this.clientId = clientId.trim();
      localStorage.setItem('lumina-gdrive-client-id', this.clientId);
    }
    if (appsScriptUrl !== undefined) {
      this.appsScriptUrl = appsScriptUrl.trim();
      localStorage.setItem('lumina-apps-script-url', this.appsScriptUrl);
    }
  }

  disconnect() {
    this.accessToken = '';
    this.tokenExpiry = 0;
    this.folderId = '';
    this.userEmail = '';
    sessionStorage.removeItem('lumina-gdrive-token');
    sessionStorage.removeItem('lumina-gdrive-token-exp');
    localStorage.removeItem('lumina-gdrive-folder-id');
    localStorage.removeItem('lumina-gdrive-email');
  }

  async loadGisScript() {
    if (this.gisLoaded || window.google?.accounts?.oauth2) {
      this.gisLoaded = true;
      return;
    }
    await new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://accounts.google.com/gsi/client';
      s.async = true;
      s.defer = true;
      s.onload = () => {
        this.gisLoaded = true;
        resolve();
      };
      s.onerror = () => reject(new Error('無法載入 Google Identity Services SDK'));
      document.head.appendChild(s);
    });
  }

  /**
   * Signs into Google Drive via OAuth 2.0 and immediately ensures the "flashcard" folder exists.
   */
  async signInWithOAuth(fullDriveScope = true) {
    if (!this.clientId) {
      throw new Error('請先填入 Google OAuth Client ID（或切換至 Apps Script 免設定橋接模式）');
    }
    await this.loadGisScript();

    const scope = fullDriveScope
      ? 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/userinfo.email'
      : 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email';

    const tokenResp = await new Promise((resolve, reject) => {
      const tokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: this.clientId,
        scope,
        callback: (resp) => {
          if (resp.error) {
            reject(new Error(resp.error_description || resp.error));
          } else {
            resolve(resp);
          }
        },
      });
      tokenClient.requestAccessToken({ prompt: this.accessToken ? '' : 'consent' });
    });

    this.accessToken = tokenResp.access_token;
    this.tokenExpiry = Date.now() + (Number(tokenResp.expires_in || 3500) - 60) * 1000;
    sessionStorage.setItem('lumina-gdrive-token', this.accessToken);
    sessionStorage.setItem('lumina-gdrive-token-exp', String(this.tokenExpiry));

    // Fetch user email for header pill
    try {
      const infoRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${this.accessToken}` },
      });
      if (infoRes.ok) {
        const info = await infoRes.json();
        this.userEmail = info.email || '';
        localStorage.setItem('lumina-gdrive-email', this.userEmail);
      }
    } catch {
      // Optional userinfo
    }

    const folder = await this.ensureFlashcardFolder();
    return { folderId: folder.id, folderName: folder.name, email: this.userEmail };
  }

  /**
   * Finds or creates a folder named "flashcard" in the user's Google Drive root.
   */
  async ensureFlashcardFolder() {
    if (this.mode === 'apps-script') {
      const res = await this._appsScriptCall({ action: 'init' });
      this.folderId = res.folderId || 'flashcard';
      return { id: this.folderId, name: FOLDER_NAME, webViewLink: res.folderUrl || '' };
    }

    const q = encodeURIComponent(
      `name='${FOLDER_NAME}' and mimeType='${FOLDER_MIME}' and trashed=false`
    );
    const search = await this._driveFetch(
      `${DRIVE_API}/files?q=${q}&fields=files(id,name,webViewLink)&spaces=drive`
    );

    if (search.files && search.files.length > 0) {
      const existing = search.files[0];
      this.folderId = existing.id;
      localStorage.setItem('lumina-gdrive-folder-id', existing.id);
      return existing;
    }

    // Create the new "flashcard" folder
    const created = await this._driveFetch(`${DRIVE_API}/files?fields=id,name,webViewLink`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: FOLDER_NAME,
        mimeType: FOLDER_MIME,
      }),
    });

    this.folderId = created.id;
    localStorage.setItem('lumina-gdrive-folder-id', created.id);
    return created;
  }

  /**
   * Lists .docx and Google Docs files inside the "flashcard" folder (or all Drive docs).
   */
  async listDocuments(allDrive = false) {
    if (this.mode === 'apps-script') {
      const res = await this._appsScriptCall({ action: 'list', allDrive });
      return res.files || [];
    }

    if (!this.folderId) await this.ensureFlashcardFolder();

    const mimeFilter = `(mimeType='${DOCX_MIME}' or mimeType='${GDOC_MIME}')`;
    const parentClause = allDrive ? '' : `'${this.folderId}' in parents and `;
    const q = encodeURIComponent(`${parentClause}${mimeFilter} and trashed=false`);
    const data = await this._driveFetch(
      `${DRIVE_API}/files?q=${q}&orderBy=modifiedTime desc&pageSize=40&fields=files(id,name,mimeType,modifiedTime,size,webViewLink)`
    );
    return data.files || [];
  }

  /**
   * Downloads a .docx (or exports a Google Doc as .docx) ArrayBuffer from Google Drive.
   */
  async downloadDocxArrayBuffer(fileId, mimeType) {
    if (this.mode === 'apps-script') {
      const res = await this._appsScriptCall({ action: 'downloadDocx', fileId });
      const bin = atob(res.base64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return bytes.buffer;
    }

    const url =
      mimeType === GDOC_MIME
        ? `${DRIVE_API}/files/${encodeURIComponent(fileId)}/export?mimeType=${encodeURIComponent(DOCX_MIME)}`
        : `${DRIVE_API}/files/${encodeURIComponent(fileId)}?alt=media`;

    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${this.accessToken}` },
    });
    if (!res.ok) {
      throw new Error(`從 Google Drive 下載檔案失敗 (${res.status})`);
    }
    return await res.arrayBuffer();
  }

  /**
   * Saves or creates a .docx file inside Google Drive / flashcard folder.
   */
  async saveDocxToDrive(fileName, docxBlob, existingFileId = null) {
    const safeName = fileName.endsWith('.docx') ? fileName : `${fileName}.docx`;

    if (this.mode === 'apps-script') {
      const base64 = await blobToBase64(docxBlob);
      return await this._appsScriptCall({
        action: 'saveDocx',
        fileId: existingFileId,
        fileName: safeName,
        base64,
      });
    }

    if (!this.folderId) await this.ensureFlashcardFolder();

    const metadata = existingFileId
      ? { name: safeName, mimeType: DOCX_MIME }
      : { name: safeName, mimeType: DOCX_MIME, parents: [this.folderId] };

    const form = new FormData();
    form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
    form.append('file', docxBlob);

    const endpoint = existingFileId
      ? `${DRIVE_UPLOAD_API}/files/${encodeURIComponent(existingFileId)}?uploadType=multipart&fields=id,name,modifiedTime,webViewLink`
      : `${DRIVE_UPLOAD_API}/files?uploadType=multipart&fields=id,name,modifiedTime,webViewLink`;

    const res = await fetch(endpoint, {
      method: existingFileId ? 'PATCH' : 'POST',
      headers: { Authorization: `Bearer ${this.accessToken}` },
      body: form,
    });

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`上傳 .docx 至 Google Drive 失敗: ${errText.slice(0, 140)}`);
    }
    return await res.json();
  }

  /**
   * Syncs the full workspace state (flashcards, Day log, checkmarks, card sizes/positions, theme)
   * into Google Drive / flashcard / lumina_flashcards_state.json
   */
  async saveWorkspaceState(stateObj) {
    if (this.mode === 'apps-script') {
      return await this._appsScriptCall({
        action: 'saveState',
        state: stateObj,
      });
    }

    if (!this.folderId) await this.ensureFlashcardFolder();

    const q = encodeURIComponent(
      `'${this.folderId}' in parents and name='${STATE_FILENAME}' and trashed=false`
    );
    const existingList = await this._driveFetch(`${DRIVE_API}/files?q=${q}&fields=files(id)`);
    const existingId = existingList.files?.[0]?.id || null;

    const jsonBlob = new Blob([JSON.stringify(stateObj, null, 2)], { type: 'application/json' });
    const metadata = existingId
      ? { name: STATE_FILENAME, mimeType: 'application/json' }
      : { name: STATE_FILENAME, mimeType: 'application/json', parents: [this.folderId] };

    const form = new FormData();
    form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
    form.append('file', jsonBlob);

    const endpoint = existingId
      ? `${DRIVE_UPLOAD_API}/files/${existingId}?uploadType=multipart&fields=id,modifiedTime`
      : `${DRIVE_UPLOAD_API}/files?uploadType=multipart&fields=id,modifiedTime`;

    const res = await fetch(endpoint, {
      method: existingId ? 'PATCH' : 'POST',
      headers: { Authorization: `Bearer ${this.accessToken}` },
      body: form,
    });
    if (!res.ok) throw new Error(`同步狀態檔至 Google Drive 失敗 (${res.status})`);
    return await res.json();
  }

  /**
   * Loads lumina_flashcards_state.json from Google Drive / flashcard if present.
   */
  async loadWorkspaceState() {
    if (this.mode === 'apps-script') {
      const res = await this._appsScriptCall({ action: 'loadState' });
      return res.state || null;
    }

    if (!this.folderId) await this.ensureFlashcardFolder();

    const q = encodeURIComponent(
      `'${this.folderId}' in parents and name='${STATE_FILENAME}' and trashed=false`
    );
    const existingList = await this._driveFetch(`${DRIVE_API}/files?q=${q}&fields=files(id)`);
    const existingId = existingList.files?.[0]?.id;
    if (!existingId) return null;

    const res = await fetch(`${DRIVE_API}/files/${existingId}?alt=media`, {
      headers: { Authorization: `Bearer ${this.accessToken}` },
    });
    if (!res.ok) return null;
    return await res.json();
  }

  async _driveFetch(url, options = {}) {
    if (!this.accessToken || Date.now() >= this.tokenExpiry) {
      throw new Error('Google Drive 登入已過期，請重新連接');
    }
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        ...(options.headers || {}),
      },
    });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(`Google Drive API 錯誤 (${res.status}): ${txt.slice(0, 120)}`);
    }
    return await res.json();
  }

  async _appsScriptCall(payload) {
    if (!this.appsScriptUrl) {
      throw new Error('請先填入 Google Apps Script Web App URL');
    }
    const res = await fetch(this.appsScriptUrl, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      throw new Error(`Apps Script 連線失敗 (${res.status})`);
    }
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    return data;
  }
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const str = String(reader.result || '');
      resolve(str.includes(',') ? str.split(',')[1] : str);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

/**
 * Copy-pasteable Google Apps Script template so users can connect their Google Drive
 * "flashcard" folder in 60 seconds without configuring Google Cloud Console OAuth!
 */
export const APPS_SCRIPT_TEMPLATE = `// 部署步驟：開啟 script.google.com → 新增專案 → 貼上此程式碼 → 部署為「網頁應用程式」(存取權限選「任何人」)
const FOLDER_NAME = 'flashcard';
const STATE_FILE = 'lumina_flashcards_state.json';

function getOrCreateFolder() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const folder = getOrCreateFolder();
    let result = { ok: true, folderId: folder.getId(), folderUrl: folder.getUrl() };

    if (req.action === 'list') {
      const files = [];
      const iter = req.allDrive ? DriveApp.getFiles() : folder.getFiles();
      while (iter.hasNext() && files.length < 40) {
        const f = iter.next();
        const mime = f.getMimeType();
        if (mime.includes('wordprocessingml') || mime === MimeType.GOOGLE_DOCS) {
          files.push({
            id: f.getId(), name: f.getName(), mimeType: mime,
            modifiedTime: f.getLastUpdated().toISOString(), webViewLink: f.getUrl()
          });
        }
      }
      result.files = files;
    } else if (req.action === 'saveState') {
      const content = JSON.stringify(req.state, null, 2);
      const it = folder.getFilesByName(STATE_FILE);
      if (it.hasNext()) it.next().setContent(content);
      else folder.createFile(STATE_FILE, content, MimeType.PLAIN_TEXT);
    } else if (req.action === 'loadState') {
      const it = folder.getFilesByName(STATE_FILE);
      result.state = it.hasNext() ? JSON.parse(it.next().getBlob().getDataAsString()) : null;
    } else if (req.action === 'downloadDocx') {
      const f = DriveApp.getFileById(req.fileId);
      const blob = f.getMimeType() === MimeType.GOOGLE_DOCS
        ? f.getAs('application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        : f.getBlob();
      result.base64 = Utilities.base64Encode(blob.getBytes());
    } else if (req.action === 'saveDocx') {
      const bytes = Utilities.base64Decode(req.base64);
      const blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', req.fileName);
      if (req.fileId) {
        try {
          // If Advanced Drive API is enabled, update in-place keeping the exact same fileId
          const updated = Drive.Files.update({ title: req.fileName }, req.fileId, blob);
          result.id = updated.id;
          result.name = req.fileName;
        } catch (_e) {
          // Fallback with standard DriveApp: put updated file in same parent folder & trash old revision
          const oldFile = DriveApp.getFileById(req.fileId);
          const parents = oldFile.getParents();
          const targetFolder = parents.hasNext() ? parents.next() : folder;
          const created = targetFolder.createFile(blob);
          oldFile.setTrashed(true);
          result.id = created.getId();
          result.name = created.getName();
          result.webViewLink = created.getUrl();
        }
      } else {
        const created = folder.createFile(blob);
        result.id = created.getId();
        result.name = created.getName();
        result.webViewLink = created.getUrl();
      }
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}`;

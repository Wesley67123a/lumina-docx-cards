# Lumina Spatial Studio (`lumina-docx-cards`)

融合 **VisionOS 空間毛玻璃 UI**、**Samsung Galaxy S9 動態絲綢光波壁紙**、**可調強調色與連續曲率圓角**、**卡片大小與位置自由拖曳保存**、**Google Drive `flashcard` 資料夾雙向同步** 與 **Docs 原檔閱讀/編輯器** 的 `.docx` 同步問答閃卡平台。

## 線上使用
- **GitHub Pages**：https://wesley67123a.github.io/lumina-docx-cards/

## 核心功能一覽

1. **VisionOS 空間毛玻璃介面 & Samsung Galaxy S9 動態壁紙**
   - 頂部懸浮功能列（狀態列 + 發光球體搜尋列 + Day 天數控制器）、左側懸浮圓角圖示側欄、底部懸浮藥丸導覽列與重度高斯模糊 (`backdrop-filter: blur(34px)`) 主視窗。
   - 內建 60fps HTML5 Canvas **Galaxy S9 Infinity 動態絲綢光波壁紙**（`S9 Silk` / `S9 Midnight` / `S9 Pearl`），即時跟隨強調色與滑鼠視差流動。
2. **即時外觀客製化**
   - **強調色 (Accent Color)**：多款預設色票與完整自訂色盤，即時連動強調卡片、Day 標籤、原檔打勾狀態與 S9 動態壁紙色系。
   - **連續曲率圓角 (Continuous Curvature Squircle)**：即時滑桿調整圓角半徑 (`12px–44px`) 與超橢圓曲率平滑指數 (`corner-shape: superellipse(n)`)。
3. **卡片大小、位置自由調整並保存**
   - 支援 **Overview 空間 Bento 卡片**（拖曳排序 + 拖曳右下角調整欄寬與高度）與 **Flashcards 閃卡**（雙欄流 / 專注單卡 / 自由畫布 X,Y 任意擺放與個別縮放，並可調整 Q/A 欄寬比例與字體大小）。
   - 點擊「保存佈局」即持久化至本地與 Google Drive。
4. **Docs 原檔閱讀、編輯與 Day 天數閃卡收錄**
   - 原汁原味呈現 `.docx` 雙欄表格與 `Page break` 分頁線，支援深色/淺色紙張切換與富文本編輯（**B** / *I* / <u>U</u> / <s>S</s> / 螢光重點 / 清單 / 新增題目列 / 匯出更新後 `.docx`）。
   - 看完題目後點擊 **`＋ 加入 Flashcard (Day X)`**，自動記錄加入天數並於原檔該列顯示 **`✓ 已加入 · Day X`** 徽章與強調色側邊標記，亦可由閃卡一鍵跳回原檔定位。
5. **Google Drive (`flashcard` 資料夾) 雙向同步**
   - 支援 **OAuth 2.0 直連** 與 **1 分鐘 Google Apps Script 橋接**兩種方式，自動於 Google Drive 建立 `flashcard` 資料夾，直接讀取/覆寫 `.docx` 原檔並同步所有閃卡進度、Day 紀錄與卡片佈局。

## 本地開發

```bash
npm install
npm run dev
```

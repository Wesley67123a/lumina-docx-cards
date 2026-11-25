# Lumina Docx Cards (`lumina-docx-cards`)

一個結合 **擬態設計 (Neumorphism)**、**背景高斯模糊 (Backdrop Blur)**、**半透明界面 (Translucent UI)**、**懸浮式儀表盤 (Floating Dashboard)**、**微交互動畫 (Micro-interactions)** 與 **連續曲率圓角 (Continuous Curvature Squircles)** 的 `.docx` 雙欄閃卡學習工作台。

## 核心特色

- **直接讀取 `.docx` 問答筆記**：純瀏覽器端解析 OpenXML (`word/document.xml`)，支援：
  - **雙欄問答表格**（如 HKDSE 筆記左欄問題 + 出處 `*OCR / *Oxford`、右欄答案與給分點 `(1)`）
  - **段落式 Q&A**（`Q: ... / A: ...` 或 `問：... / 答：...`）
  - **內嵌圖片提取**：若 `.docx` 題目或答案內含圖表亦可自動顯示
- **問題與答案同時並排顯示 (Simultaneous Q&A Flashcards)**：左欄採用擬態凸起卡片呈現題目與配分，右欄採用擬態內凹槽 (Inset Well) 呈現逐條給分標準，點擊單一給分點即可微交互核對，全數勾選自動標記「已掌握」。
- **雙檢視模式**：
  - **雙欄總覽 (Stream View)**：一覽整份章節所有問答閃卡
  - **專注單卡 (Focus Stage)**：單張並排問答搭配底部懸浮控制列與鍵盤 `←` / `→` 切換
- **視覺與工藝細節**：
  - `backdrop-filter: blur(28px) saturate(175%)` 多層次毛玻璃半透明面板
  - 雙光源柔和擬態凸起與內凹陰影（支援淺色珍珠霧白 / 深色冰川岩灰雙主題切換）
  - `corner-shape: superellipse(...)` 連續曲率平滑超橢圓圓角 (Squircle)

## 本地開發與啟動

```bash
npm install
npm run dev
```

建置靜態網頁版本：

```bash
npm run build
```

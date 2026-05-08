# LightAnalyzer - 輕量資料分析工具

純前端應用，無需安裝，瀏覽器開啟 `index.html` 即用。上傳 Excel/CSV 檔案，自動計算基礎統計指標並匯出 PDF 報告。

## 檔案結構

```
E:\New folder\
├── index.html          — HTML 結構（視圖、組件）
├── styles.css          — 自訂樣式（動效、拖曳、選擇器、報告模板）
├── app.js              — 應用邏輯（i18n、解析、統計、雙模式 PDF 匯出）
└── README.md           — 本檔案
```

## index.html

### `<head>` 資源

| 資源 | 來源 | 用途 |
|---|---|---|
| Tailwind CSS | Play CDN | 樣式框架，擴展 spacing 3.5/4.5 |
| Lucide Icons | unpkg 0.460.0 UMD | SVG 圖示 `<i data-lucide="...">` |
| SheetJS | cdn.sheetjs.com 0.20.3 | .xlsx / .csv 解析 |
| jsPDF | cdnjs 2.5.1 UMD | PDF 文件生成 |
| html2canvas | cdnjs 1.4.1 | HTML → Canvas（解決 CJK PDF 亂碼） |
| styles.css | 本地 | 自訂樣式 |

### `<body>` 視圖結構

- **導航欄** — Logo + 首頁/資料分析切換 + 語言下拉（簡體/繁體/English）
- **首頁** `#view-home` — Hero 漸變區 → 3 張功能卡片 → 4 步使用說明 → CTA → Footer
- **分析頁** `#view-analysis`
  - 拖曳上傳區（虛線框，拖入高亮）
  - 檔案資訊條（上傳後顯示，含行列數 + 清除按鈕）
  - 分析模式切換：`[按列] [按行]`
  - 列/行選擇器（根據模式動態切換標籤和選項）
  - 4 張統計卡片（平均值、中位數、眾數、標準差）+ 3 個補充指標
  - 行數選擇器：`[10] [50] [100] [500] [全部]` + 自訂輸入
  - 匯出按鈕組：`[匯出選中分析]` `[匯出全部統計]` + 描述文字
  - `[重新上傳]` 按鈕
  - 資料預覽表格（支援列高亮/行高亮，可捲動）

## styles.css

| 區塊 | 說明 |
|---|---|
| Reset & Base | 捲軸美化、平滑捲動 |
| Utility | `.view-hidden` 全域隱藏 |
| Drag & Drop | `.drag-over` 高亮邊框 + 縮放 |
| Statistics Cards | `.stat-card` hover 上浮 + 陰影 |
| Language | `.lang-btn.active-lang` 選中態 |
| Animations | `fadeInUp` 漸入 + 子元素 stagger |
| Hero | `.hero-gradient` 紫藍漸變 |
| Row Count Selector | `.rc-segmented` / `.rc-btn` / `.rc-active` / `.rc-custom` |
| Report Template | `.report-container` / `.report-stat-card` / `.report-table` 等（用於 PDF 渲染） |
| Responsive | 小螢幕調整 |

## app.js

### 狀態變數

```
currentLang      — localStorage('la-lang')，預設 zh-CN
currentView      — 'home' | 'analysis'
workbookData     — { headers, rows, fileName }
allHeaders[]     — 列名陣列
allRows[][]      — 二維資料陣列
currentColIndex  — 當前選中列/行索引
displayRowCount  — 表格/PDF 顯示行數（-1 = 全部），預設 100
analysisMode     — 'column' | 'row'
```

### 模組

| 模組 | 函式 | 說明 |
|---|---|---|
| **i18n** | `t(key)`, `updateAllI18n()`, `setLang(lang)`, `updateLangMenuUI()` | zh-CN / zh-TW / en 三語言，80+ 翻譯鍵 |
| **視圖** | `showView(view)`, `updateNavActive()` | 首頁 / 分析頁切換 |
| **語言下拉** | `toggleLangMenu()`, 全域 click 監聽 | 展開/收起/外部關閉 |
| **檔案處理** | `setupDragDrop()`, `handleFileSelect()`, `processFile()`, `onFileLoaded()` | 拖曳/瀏覽 → FileReader → XLSX.read() → 提取資料 |
| **資料檢測** | `isColumnNumeric()`, `isRowNumeric()`, `getColumnData()`, `getRowData()`, `getAnalysisData()` | 剝離去噪後判數值、提取數值陣列 |
| **分析模式** | `setAnalysisMode(mode)`, `populateSelector()` | 按列/按行切換，重建下拉選單 |
| **統計** | `calcMean()`, `calcMedian()`, `calcMode()`, `calcStdDev()`, `formatNum()` | 4 項指標 + 格式化 |
| **顯示** | `onSelectionChange()`, `refreshStatsDisplay()`, `renderTable()` | 重算 → 更新卡片 → 渲染表格（列/行高亮） |
| **行數選擇** | `setRowCount(n)`, `onCustomRowCount()` | 預設/自訂行數，聯動表格 + PDF |
| **重置** | `resetAnalysis()` | 清空所有狀態、恢復 UI 預設值 |
| **單次匯出** | `exportPDF()`, `buildReportElement()`, `buildStatCard()`, `buildExtraItem()` | 選中列/行 → 統計卡片 + 資料表 → PDF |
| **批量匯出** | `exportBatchPDF()`, `buildBatchReportElement()` | 全部列/行 → 彙總統計表 + 資料表 → PDF |
| **工具** | `escHtml()` | HTML 轉義 |
| **初始化** | `DOMContentLoaded` | 綁定事件、載入 i18n、渲染圖示 |

### PDF 匯出流程（兩種匯出共用此模式）

```
exportPDF() / exportBatchPDF()
  → buildXxxReportElement()              建立報告 DOM（1100px，系統 CJK 字型）
  → 掛載 body（off-screen）              讓瀏覽器完成字型渲染
  → setTimeout 300ms
  → html2canvas(container, {scale:2})    高 DPI 截圖，中文完美保留
  → 移除暫存 DOM
  → 逐頁切片 canvas → addImage() → PDF  自動跨頁 + 頁碼
  → doc.save('xxx_report.pdf')
```

### 兩種 PDF 匯出對比

| | 匯出選中分析 | 匯出全部統計 |
|---|---|---|
| **按鈕** | 實心 indigo `file-down` | 白底 indigo 邊框 `layers` |
| **報告標題** | 資料分析報告 | 全量統計分析報告 |
| **統計內容** | 當前列/行的 4 張統計卡片 | 彙總表：每列/行 × 7 個指標 |
| **資料表** | 有 | 有 |
| **檔案名稱** | `xxx_analysis_report.pdf` | `xxx_batch_analysis_report.pdf` |

## 資料流

```
使用者拖曳/選擇檔案
  → FileReader.readAsArrayBuffer()
  → XLSX.read(data, { type: 'array' })
  → XLSX.utils.sheet_to_json(sheet, { header: 1 })
  → allHeaders + allRows
  → isColumnNumeric() / isRowNumeric() 檢測
  → populateSelector() 填充下拉
  → onSelectionChange() → getAnalysisData()
  → calcMean/Median/Mode/StdDev
  → 更新統計卡片 + renderTable()
  → [可選] exportPDF() — 單次 / exportBatchPDF() — 批量
```

## 瀏覽器相容性

Chrome / Firefox / Safari / Edge 最新兩個大版本。依賴：
- FileReader API
- ES5+ (var, function, Promise)
- CSS Grid / Flexbox
- Canvas 2D

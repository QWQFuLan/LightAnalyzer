# LightAnalyzer - 轻量数据分析工具

纯前端应用，无需安装，浏览器打开 `index.html` 即用。上传 Excel/CSV 文件，自动计算基础统计指标并导出 PDF 报告。

## 文件结构

```
E:\New folder\
├── index.html          — HTML 结构（视图、组件）
├── styles.css          — 自定义样式（动效、拖拽、选择器、报告模板）
├── app.js              — 应用逻辑（i18n、解析、统计、双模式 PDF 导出）
└── README.md           — 本文件
```

## index.html

### `<head>` 资源

| 资源 | 来源 | 用途 |
|---|---|---|
| Tailwind CSS | Play CDN | 样式框架，扩展 spacing 3.5/4.5 |
| Lucide Icons | unpkg 0.460.0 UMD | SVG 图标 `<i data-lucide="...">` |
| SheetJS | cdn.sheetjs.com 0.20.3 | .xlsx / .csv 解析 |
| jsPDF | cdnjs 2.5.1 UMD | PDF 文档生成 |
| html2canvas | cdnjs 1.4.1 | HTML → Canvas（解决 CJK PDF 乱码） |
| styles.css | 本地 | 自定义样式 |

### `<body>` 视图结构

- **导航栏** — Logo + 首页/数据分析切换 + 语言下拉（简体/繁體/English）
- **首页** `#view-home` — Hero 渐变区 → 3 张功能卡片 → 4 步使用说明 → CTA → Footer
- **分析页** `#view-analysis`
  - 拖拽上传区（虚线框，拖入高亮）
  - 文件信息条（上传后显示，含行列数 + 清除按钮）
  - 分析模式切换：`[按列] [按行]`
  - 列/行选择器（根据模式动态切换标签和选项）
  - 4 张统计卡片（平均值、中位数、众数、标准差）+ 3 个补充指标
  - 行数选择器：`[10] [50] [100] [500] [全部]` + 自定义输入
  - 导出按钮组：`[导出选中分析]` `[导出全部统计]` + 描述文字
  - `[重新上传]` 按钮
  - 数据预览表格（支持列高亮/行高亮，可滚动）

## styles.css

| 区块 | 说明 |
|---|---|
| Reset & Base | 滚动条美化、平滑滚动 |
| Utility | `.view-hidden` 全局隐藏 |
| Drag & Drop | `.drag-over` 高亮边框 + 缩放 |
| Statistics Cards | `.stat-card` hover 上浮 + 阴影 |
| Language | `.lang-btn.active-lang` 选中态 |
| Animations | `fadeInUp` 渐入 + 子元素 stagger |
| Hero | `.hero-gradient` 紫蓝渐变 |
| Row Count Selector | `.rc-segmented` / `.rc-btn` / `.rc-active` / `.rc-custom` |
| Report Template | `.report-container` / `.report-stat-card` / `.report-table` 等（用于 PDF 渲染） |
| Responsive | 小屏幕调整 |

## app.js

### 状态变量

```
currentLang      — localStorage('la-lang')，默认 zh-CN
currentView      — 'home' | 'analysis'
workbookData     — { headers, rows, fileName }
allHeaders[]     — 列名数组
allRows[][]      — 二维数据数组
currentColIndex  — 当前选中列/行索引
displayRowCount  — 表格/PDF 显示行数（-1 = 全部），默认 100
analysisMode     — 'column' | 'row'
```

### 模块

| 模块 | 函数 | 说明 |
|---|---|---|
| **i18n** | `t(key)`, `updateAllI18n()`, `setLang(lang)`, `updateLangMenuUI()` | zh-CN / zh-TW / en 三语言，80+ 翻译键 |
| **视图** | `showView(view)`, `updateNavActive()` | 首页 / 分析页切换 |
| **语言下拉** | `toggleLangMenu()`, 全局 click 监听 | 展开/收起/外部关闭 |
| **文件处理** | `setupDragDrop()`, `handleFileSelect()`, `processFile()`, `onFileLoaded()` | 拖拽/浏览 → FileReader → XLSX.read() → 提取数据 |
| **数据检测** | `isColumnNumeric()`, `isRowNumeric()`, `getColumnData()`, `getRowData()`, `getAnalysisData()` | 剥离去噪后判数值、提取数值数组 |
| **分析模式** | `setAnalysisMode(mode)`, `populateSelector()` | 按列/按行切换，重建下拉菜单 |
| **统计** | `calcMean()`, `calcMedian()`, `calcMode()`, `calcStdDev()`, `formatNum()` | 4 项指标 + 格式化 |
| **显示** | `onSelectionChange()`, `refreshStatsDisplay()`, `renderTable()` | 重算 → 更新卡片 → 渲染表格（列/行高亮） |
| **行数选择** | `setRowCount(n)`, `onCustomRowCount()` | 预设/自定义行数，联动表格 + PDF |
| **重置** | `resetAnalysis()` | 清空所有状态、恢复 UI 默认值 |
| **单次导出** | `exportPDF()`, `buildReportElement()`, `buildStatCard()`, `buildExtraItem()` | 选中列/行 → 统计卡片 + 数据表 → PDF |
| **批量导出** | `exportBatchPDF()`, `buildBatchReportElement()` | 全部列/行 → 汇总统计表 + 数据表 → PDF |
| **工具** | `escHtml()` | HTML 转义 |
| **初始化** | `DOMContentLoaded` | 绑定事件、加载 i18n、渲染图标 |

### PDF 导出流程（两种导出共用此模式）

```
exportPDF() / exportBatchPDF()
  → buildXxxReportElement()              创建报告 DOM（1100px，系统 CJK 字体）
  → 挂载 body（off-screen）              让浏览器完成字体渲染
  → setTimeout 300ms
  → html2canvas(container, {scale:2})    高 DPI 截图，中文完美保留
  → 移除临时 DOM
  → 逐页切片 canvas → addImage() → PDF  自动跨页 + 页码
  → doc.save('xxx_report.pdf')
```

### 两种 PDF 导出对比

| | 导出选中分析 | 导出全部统计 |
|---|---|---|
| **按钮** | 实心 indigo `file-down` | 白底 indigo 边框 `layers` |
| **报告标题** | 数据分析报告 | 全量统计分析报告 |
| **统计内容** | 当前列/行的 4 张统计卡片 | 汇总表：每列/行 × 7 个指标 |
| **数据表** | 有 | 有 |
| **文件名** | `xxx_analysis_report.pdf` | `xxx_batch_analysis_report.pdf` |

## 数据流

```
用户拖拽/选择文件
  → FileReader.readAsArrayBuffer()
  → XLSX.read(data, { type: 'array' })
  → XLSX.utils.sheet_to_json(sheet, { header: 1 })
  → allHeaders + allRows
  → isColumnNumeric() / isRowNumeric() 检测
  → populateSelector() 填充下拉
  → onSelectionChange() → getAnalysisData()
  → calcMean/Median/Mode/StdDev
  → 更新统计卡片 + renderTable()
  → [可选] exportPDF() — 单次 / exportBatchPDF() — 批量
```

## 浏览器兼容

Chrome / Firefox / Safari / Edge 最新两个大版本。依赖：
- FileReader API
- ES5+ (var, function, Promise)
- CSS Grid / Flexbox
- Canvas 2D

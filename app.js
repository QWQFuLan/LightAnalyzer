/* ============================================
   LightAnalyzer - 轻量数据分析工具
   Application Logic
   ============================================ */

// ─── State ────────────────────────────────────────────
var currentLang = localStorage.getItem('la-lang') || 'zh-CN';
var currentView = 'home';
var workbookData = null;
var allHeaders = [];
var allRows = [];
var currentColIndex = -1;
var displayRowCount = 100;  // -1 means all rows
var analysisMode = 'column'; // 'column' or 'row'

// ─── i18n Translations ────────────────────────────────
const i18n = {
  'zh-CN': {
    'app.name': 'LightAnalyzer',
    'nav.home': '首页',
    'nav.analysis': '数据分析',
    'hero.badge': '轻量 · 快速 · 隐私优先',
    'hero.title': '几秒钟内，从数据到洞察',
    'hero.subtitle': 'LightAnalyzer 是一款轻量化的在线数据分析工具。只需拖拽您的 Excel 或 CSV 文件，即可自动计算关键统计指标，并导出专业的 PDF 分析报告。所有数据均在您的浏览器本地处理，不会上传到任何服务器。',
    'hero.cta': '开始分析数据',
    'hero.learn': '了解更多',
    'features.title': '核心功能',
    'features.subtitle': '专为快速数据分析而设计，无需安装任何软件',
    'features.f1.title': '拖拽上传',
    'features.f1.desc': '支持 .xlsx 和 .csv 格式的表格文件。直接拖拽到上传区域，或点击选择文件即可开始分析。',
    'features.f2.title': '即时统计分析',
    'features.f2.desc': '自动计算平均值、中位数、众数和标准差四项核心统计指标，帮助您快速理解数据分布特征。',
    'features.f3.title': '导出 PDF 报告',
    'features.f3.desc': '一键将分析结果导出为排版精美的 PDF 报告，包含统计数据摘要和完整数据表格，方便分享和存档。',
    'howto.title': '如何使用',
    'howto.subtitle': '只需简单四步，即可完成数据分析',
    'howto.s1.title': '上传文件',
    'howto.s1.desc': '拖拽或选择您的 Excel (.xlsx) 或 CSV 文件',
    'howto.s2.title': '选择列',
    'howto.s2.desc': '从下拉菜单中选择您想要分析的数值列',
    'howto.s3.title': '查看结果',
    'howto.s3.desc': '查看四项统计指标和完整的数据明细表格',
    'howto.s4.title': '导出报告',
    'howto.s4.desc': '点击导出按钮，下载包含分析结果的 PDF 报告',
    'cta.title': '准备好分析您的数据了吗？',
    'cta.subtitle': '完全免费，无需注册，数据不上传服务器',
    'cta.button': '开始使用',
    'footer.text': 'LightAnalyzer - 轻量数据分析工具 · 所有数据仅在浏览器本地处理',
    'upload.title': '拖拽文件到此处',
    'upload.hint': '支持 .xlsx 和 .csv 格式',
    'upload.browse': '浏览文件',
    'analysis.selectCol': '选择分析列：',
    'analysis.numeric': '数值列',
    'analysis.nonnumeric': '非数值列',
    'stat.mean': '平均值',
    'stat.meanDesc': '所有数值的算术平均',
    'stat.median': '中位数',
    'stat.medianDesc': '排序后居中的数值',
    'stat.mode': '众数',
    'stat.modeDesc': '出现频率最高的值',
    'stat.stddev': '标准差',
    'stat.stddevDesc': '数据离散程度的度量',
    'stat.count': '数据量',
    'stat.min': '最小值',
    'stat.max': '最大值',
    'stat.noMode': '无众数',
    'stat.multiMode': '个众数',
    'export.button': '导出 PDF 报告',
    'export.single': '导出选中分析',
    'export.singleDesc': '仅当前所选列/行的统计结果与数据表',
    'export.batch': '导出全部统计',
    'export.batchDesc': '所有列/行的统计摘要与完整数据表',
    'export.reset': '重新上传',
    'table.title': '数据预览',
    'table.rows': '行',
    'table.cols': '列',
    'table.showing': '显示前 {n} 行，共 {total} 行',
    'report.title': '数据分析报告',
    'report.batchTitle': '全量统计分析报告',
    'report.mode': '分析模式',
    'report.file': '分析文件',
    'report.column': '分析列',
    'report.row': '分析行',
    'report.date': '生成时间',
    'report.summary': '统计摘要',
    'report.dataTable': '数据明细表',
    'report.exporting': '正在生成报告...',
    'display.rows': '显示行数：',
    'display.all': '全部',
    'analysis.mode': '分析模式：',
    'analysis.byColumn': '按列',
    'analysis.byRow': '按行',
    'analysis.selectRow': '选择分析行：',
    'validation.noNumeric': '所选列不包含足够的数值数据，请选择其他列。',
    'validation.empty': '文件中未找到有效数据，请检查文件格式。',
    'validation.noNumericRow': '所选行不包含足够的数值数据，请选择其他行。',
  },
  'zh-TW': {
    'app.name': 'LightAnalyzer',
    'nav.home': '首頁',
    'nav.analysis': '數據分析',
    'hero.badge': '輕量 · 快速 · 隱私優先',
    'hero.title': '幾秒鐘內，從數據到洞察',
    'hero.subtitle': 'LightAnalyzer 是一款輕量化的線上數據分析工具。只需拖曳您的 Excel 或 CSV 檔案，即可自動計算關鍵統計指標，並匯出專業的 PDF 分析報告。所有數據均在您的瀏覽器本機處理，不會上傳到任何伺服器。',
    'hero.cta': '開始分析數據',
    'hero.learn': '了解更多',
    'features.title': '核心功能',
    'features.subtitle': '專為快速數據分析而設計，無需安裝任何軟體',
    'features.f1.title': '拖曳上傳',
    'features.f1.desc': '支援 .xlsx 和 .csv 格式的表格檔案。直接拖曳到上傳區域，或點擊選擇檔案即可開始分析。',
    'features.f2.title': '即時統計分析',
    'features.f2.desc': '自動計算平均值、中位數、眾數和標準差四項核心統計指標，幫助您快速理解數據分布特徵。',
    'features.f3.title': '匯出 PDF 報告',
    'features.f3.desc': '一鍵將分析結果匯出為排版精美的 PDF 報告，包含統計數據摘要和完整數據表格，方便分享和存檔。',
    'howto.title': '如何使用',
    'howto.subtitle': '只需簡單四步，即可完成數據分析',
    'howto.s1.title': '上傳檔案',
    'howto.s1.desc': '拖曳或選擇您的 Excel (.xlsx) 或 CSV 檔案',
    'howto.s2.title': '選擇欄位',
    'howto.s2.desc': '從下拉選單中選擇您想要分析的數值欄位',
    'howto.s3.title': '檢視結果',
    'howto.s3.desc': '檢視四項統計指標和完整的數據明細表格',
    'howto.s4.title': '匯出報告',
    'howto.s4.desc': '點擊匯出按鈕，下載包含分析結果的 PDF 報告',
    'cta.title': '準備好分析您的數據了嗎？',
    'cta.subtitle': '完全免費，無需註冊，數據不上傳伺服器',
    'cta.button': '開始使用',
    'footer.text': 'LightAnalyzer - 輕量數據分析工具 · 所有數據僅在瀏覽器本機處理',
    'upload.title': '拖曳檔案到此處',
    'upload.hint': '支援 .xlsx 和 .csv 格式',
    'upload.browse': '瀏覽檔案',
    'analysis.selectCol': '選擇分析欄位：',
    'analysis.numeric': '數值欄位',
    'analysis.nonnumeric': '非數值欄位',
    'stat.mean': '平均值',
    'stat.meanDesc': '所有數值的算術平均',
    'stat.median': '中位數',
    'stat.medianDesc': '排序後居中的數值',
    'stat.mode': '眾數',
    'stat.modeDesc': '出現頻率最高的值',
    'stat.stddev': '標準差',
    'stat.stddevDesc': '數據離散程度的度量',
    'stat.count': '數據量',
    'stat.min': '最小值',
    'stat.max': '最大值',
    'stat.noMode': '無眾數',
    'stat.multiMode': '個眾數',
    'export.button': '匯出 PDF 報告',
    'export.single': '匯出選中分析',
    'export.singleDesc': '僅當前所選欄/行的統計結果與數據表',
    'export.batch': '匯出全部統計',
    'export.batchDesc': '所有欄/行的統計摘要與完整數據表',
    'export.reset': '重新上傳',
    'table.title': '數據預覽',
    'table.rows': '列',
    'table.cols': '欄',
    'table.showing': '顯示前 {n} 列，共 {total} 列',
    'report.title': '數據分析報告',
    'report.batchTitle': '全量統計分析報告',
    'report.mode': '分析模式',
    'report.file': '分析檔案',
    'report.column': '分析欄位',
    'report.row': '分析行',
    'report.date': '生成時間',
    'report.summary': '統計摘要',
    'report.dataTable': '數據明細表',
    'report.exporting': '正在生成報告...',
    'display.rows': '顯示行數：',
    'display.all': '全部',
    'analysis.mode': '分析模式：',
    'analysis.byColumn': '按欄',
    'analysis.byRow': '按行',
    'analysis.selectRow': '選擇分析行：',
    'validation.noNumeric': '所選欄位不包含足夠的數值數據，請選擇其他欄位。',
    'validation.empty': '檔案中未找到有效數據，請檢查檔案格式。',
    'validation.noNumericRow': '所選行不包含足夠的數值數據，請選擇其他行。',
  },
  'en': {
    'app.name': 'LightAnalyzer',
    'nav.home': 'Home',
    'nav.analysis': 'Analysis',
    'hero.badge': 'Lightweight · Fast · Privacy First',
    'hero.title': 'From data to insights in seconds',
    'hero.subtitle': 'LightAnalyzer is a lightweight online data analysis tool. Simply drag and drop your Excel or CSV file to automatically compute key statistical metrics and export a professional PDF report. All data is processed locally in your browser — nothing is ever uploaded to any server.',
    'hero.cta': 'Start Analyzing',
    'hero.learn': 'Learn More',
    'features.title': 'Core Features',
    'features.subtitle': 'Designed for rapid data analysis, no software installation required',
    'features.f1.title': 'Drag & Drop Upload',
    'features.f1.desc': 'Supports .xlsx and .csv spreadsheet files. Drag your file directly onto the upload area or click to browse and start analyzing instantly.',
    'features.f2.title': 'Instant Statistics',
    'features.f2.desc': 'Automatically computes the four key statistical measures — mean, median, mode, and standard deviation — to help you quickly understand your data distribution.',
    'features.f3.title': 'Export PDF Report',
    'features.f3.desc': 'One-click export of your analysis results into a beautifully formatted PDF report, including a statistical summary and full data table for easy sharing and archiving.',
    'howto.title': 'How to Use',
    'howto.subtitle': 'Just four simple steps to complete your data analysis',
    'howto.s1.title': 'Upload File',
    'howto.s1.desc': 'Drag and drop or browse for your Excel (.xlsx) or CSV file',
    'howto.s2.title': 'Select Column',
    'howto.s2.desc': 'Choose the numeric column you want to analyze from the dropdown',
    'howto.s3.title': 'View Results',
    'howto.s3.desc': 'Review the four statistical metrics and the complete data table',
    'howto.s4.title': 'Export Report',
    'howto.s4.desc': 'Click the export button to download a PDF report with your analysis',
    'cta.title': 'Ready to analyze your data?',
    'cta.subtitle': 'Completely free, no registration, data never leaves your browser',
    'cta.button': 'Get Started',
    'footer.text': 'LightAnalyzer - Lightweight Data Analysis Tool · All data processed locally',
    'upload.title': 'Drop your file here',
    'upload.hint': 'Supports .xlsx and .csv formats',
    'upload.browse': 'Browse Files',
    'analysis.selectCol': 'Select column to analyze:',
    'analysis.numeric': 'Numeric',
    'analysis.nonnumeric': 'Non-numeric',
    'stat.mean': 'Mean',
    'stat.meanDesc': 'Arithmetic average of all values',
    'stat.median': 'Median',
    'stat.medianDesc': 'The middle value when sorted',
    'stat.mode': 'Mode',
    'stat.modeDesc': 'The most frequently occurring value',
    'stat.stddev': 'Std Deviation',
    'stat.stddevDesc': 'Measure of data dispersion',
    'stat.count': 'Count',
    'stat.min': 'Minimum',
    'stat.max': 'Maximum',
    'stat.noMode': 'No mode',
    'stat.multiMode': 'modes',
    'export.button': 'Export PDF Report',
    'export.single': 'Export Selected',
    'export.singleDesc': 'Stats & data table for current selection only',
    'export.batch': 'Export All Stats',
    'export.batchDesc': 'Summary of all columns/rows with full data table',
    'export.reset': 'Upload New File',
    'table.title': 'Data Preview',
    'table.rows': 'rows',
    'table.cols': 'cols',
    'table.showing': 'Showing first {n} rows of {total}',
    'report.title': 'Data Analysis Report',
    'report.batchTitle': 'Complete Statistical Report',
    'report.mode': 'Analysis Mode',
    'report.file': 'Analyzed File',
    'report.column': 'Analyzed Column',
    'report.row': 'Analyzed Row',
    'report.date': 'Generated',
    'report.summary': 'Statistical Summary',
    'report.dataTable': 'Data Table',
    'report.exporting': 'Generating report...',
    'display.rows': 'Display rows:',
    'display.all': 'All',
    'analysis.mode': 'Analysis mode:',
    'analysis.byColumn': 'By Column',
    'analysis.byRow': 'By Row',
    'analysis.selectRow': 'Select row to analyze:',
    'validation.noNumeric': 'The selected column does not contain enough numeric data. Please choose another column.',
    'validation.noNumericRow': 'The selected row does not contain enough numeric data. Please choose another row.',
    'validation.empty': 'No valid data found in the file. Please check the file format.',
  }
};

function t(key) {
  return i18n[currentLang]?.[key] || i18n['zh-CN']?.[key] || key;
}

// ─── i18n Update ──────────────────────────────────────
function updateAllI18n() {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const key = el.getAttribute('data-i18n');
    const text = t(key);
    if (text) el.textContent = text;
  });
  document.documentElement.lang = currentLang;
  updateNavActive();
  updateLangMenuUI();
  lucide.createIcons();
}

function setLang(lang) {
  currentLang = lang;
  localStorage.setItem('la-lang', lang);
  updateAllI18n();
  if (allRows.length > 0 && currentColIndex >= 0) {
    refreshStatsDisplay();
  }
  document.getElementById('lang-menu').classList.add('hidden');
}

function updateLangMenuUI() {
  const labelMap = { 'zh-CN': '简体中文', 'zh-TW': '繁體中文', 'en': 'English' };
  document.getElementById('current-lang-label').textContent = labelMap[currentLang] || '简体中文';

  document.querySelectorAll('#lang-menu .lang-btn').forEach(btn => {
    const lang = btn.getAttribute('data-lang');
    const checkMark = btn.querySelector('.check-mark');
    if (lang === currentLang) {
      btn.classList.add('active-lang');
      if (checkMark) checkMark.classList.remove('hidden');
    } else {
      btn.classList.remove('active-lang');
      if (checkMark) checkMark.classList.add('hidden');
    }
  });
}

// ─── View Switching ──────────────────────────────────
function showView(view) {
  currentView = view;
  document.getElementById('view-home').classList.toggle('view-hidden', view !== 'home');
  document.getElementById('view-analysis').classList.toggle('view-hidden', view !== 'analysis');
  updateNavActive();
  window.scrollTo({ top: 0, behavior: 'smooth' });
  setTimeout(() => lucide.createIcons(), 100);
}

function updateNavActive() {
  const homeBtn = document.getElementById('nav-home');
  const analysisBtn = document.getElementById('nav-analysis');
  if (currentView === 'home') {
    homeBtn.className = 'px-3.5 py-2 rounded-lg text-sm font-medium transition-colors bg-indigo-50 text-indigo-700';
    analysisBtn.className = 'px-3.5 py-2 rounded-lg text-sm font-medium text-gray-600 hover:text-gray-900 hover:bg-gray-100 transition-colors';
  } else {
    homeBtn.className = 'px-3.5 py-2 rounded-lg text-sm font-medium text-gray-600 hover:text-gray-900 hover:bg-gray-100 transition-colors';
    analysisBtn.className = 'px-3.5 py-2 rounded-lg text-sm font-medium transition-colors bg-indigo-50 text-indigo-700';
  }
  homeBtn.textContent = t('nav.home');
  analysisBtn.textContent = t('nav.analysis');
}

// ─── Language Dropdown ────────────────────────────────
function toggleLangMenu() {
  const menu = document.getElementById('lang-menu');
  menu.classList.toggle('hidden');
  updateLangMenuUI();
}

document.addEventListener('click', function(e) {
  const sw = document.getElementById('lang-switcher');
  if (sw && !sw.contains(e.target)) {
    document.getElementById('lang-menu').classList.add('hidden');
  }
});

// ─── File Handling ────────────────────────────────────
function setupDragDrop() {
  const dropZone = document.getElementById('drop-zone');

  ['dragenter', 'dragover'].forEach(function(evt) {
    dropZone.addEventListener(evt, function(e) {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(function(evt) {
    dropZone.addEventListener(evt, function(e) {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('drag-over');
    });
  });

  dropZone.addEventListener('drop', function(e) {
    var files = e.dataTransfer.files;
    if (files.length > 0) processFile(files[0]);
  });
}

function handleFileSelect(event) {
  var file = event.target.files[0];
  if (file) processFile(file);
}

function processFile(file) {
  var ext = file.name.split('.').pop().toLowerCase();
  if (['xlsx', 'xls', 'csv'].indexOf(ext) === -1) {
    alert(t('validation.empty'));
    return;
  }

  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: 'array' });
      var sheetName = workbook.SheetNames[0];
      var sheet = workbook.Sheets[sheetName];
      var json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      if (!json || json.length === 0) {
        alert(t('validation.empty'));
        return;
      }

      allHeaders = json[0].map(function(h, i) {
        return String(h || '').trim() || 'Column ' + (i + 1);
      });
      allRows = json.slice(1).filter(function(row) {
        return row.some(function(cell) { return String(cell).trim() !== ''; });
      });

      if (allRows.length === 0) {
        alert(t('validation.empty'));
        return;
      }

      workbookData = { headers: allHeaders, rows: allRows, fileName: file.name };
      onFileLoaded();
    } catch (err) {
      console.error(err);
      alert(t('validation.empty'));
    }
  };
  reader.readAsArrayBuffer(file);
}

function onFileLoaded() {
  // Reset to defaults
  displayRowCount = 100;
  document.getElementById('custom-row-count').value = '';
  setRowCount(100);
  analysisMode = 'column';
  document.getElementById('mode-col').classList.add('rc-active');
  document.getElementById('mode-row').classList.remove('rc-active');

  document.getElementById('file-info').classList.remove('hidden');
  document.getElementById('file-name').textContent = workbookData.fileName;
  document.getElementById('file-rows').textContent = allRows.length + ' ' + t('table.rows') + ' · ' + allHeaders.length + ' ' + t('table.cols');

  document.getElementById('analysis-section').classList.remove('hidden');
  document.getElementById('upload-section').style.opacity = '0.6';

  populateSelector();
  onSelectionChange();
  lucide.createIcons();
}

// ─── Analysis Mode ───────────────────────────────────
function setAnalysisMode(mode) {
  analysisMode = mode;
  document.getElementById('mode-col').classList.toggle('rc-active', mode === 'column');
  document.getElementById('mode-row').classList.toggle('rc-active', mode === 'row');

  var label = document.getElementById('selector-label');
  var key = mode === 'column' ? 'analysis.selectCol' : 'analysis.selectRow';
  label.setAttribute('data-i18n', key);
  label.textContent = t(key);

  populateSelector();
  onSelectionChange();
}

function populateSelector() {
  var select = document.getElementById('column-select');
  select.innerHTML = '';
  var firstTargetIdx = -1;

  if (analysisMode === 'column') {
    allHeaders.forEach(function(h, i) {
      var opt = document.createElement('option');
      opt.value = i;
      opt.textContent = h;
      select.appendChild(opt);
      if (firstTargetIdx === -1 && isColumnNumeric(i)) firstTargetIdx = i;
    });
  } else {
    allRows.forEach(function(row, i) {
      var opt = document.createElement('option');
      opt.value = i;
      var label = '#' + (i + 1);
      var firstVal = row[0];
      if (firstVal !== '' && firstVal !== null && firstVal !== undefined) {
        label += ' · ' + String(firstVal).substring(0, 28);
      }
      opt.textContent = label;
      select.appendChild(opt);
      if (firstTargetIdx === -1 && isRowNumeric(i)) firstTargetIdx = i;
    });
  }

  if (firstTargetIdx >= 0) select.value = firstTargetIdx;
}

function getRowData(rowIdx) {
  var nums = [];
  var row = allRows[rowIdx];
  if (!row) return nums;
  for (var ci = 0; ci < allHeaders.length; ci++) {
    var val = row[ci];
    if (val === '' || val === null || val === undefined) continue;
    var num = parseFloat(String(val).replace(/[,\s%￥$€£¥]/g, ''));
    if (!isNaN(num)) nums.push(num);
  }
  return nums;
}

function isRowNumeric(rowIdx) {
  var row = allRows[rowIdx];
  if (!row) return false;
  var numericCount = 0;
  var totalCount = 0;
  for (var ci = 0; ci < allHeaders.length; ci++) {
    var val = row[ci];
    if (val === '' || val === null || val === undefined) continue;
    totalCount++;
    var num = parseFloat(String(val).replace(/[,\s%￥$€£¥]/g, ''));
    if (!isNaN(num)) numericCount++;
  }
  return totalCount > 0 && (numericCount / totalCount) > 0.5;
}

function getAnalysisData(idx) {
  return analysisMode === 'column' ? getColumnData(idx) : getRowData(idx);
}

function isColumnNumeric(colIdx) {
  var numericCount = 0;
  var totalCount = 0;
  for (var i = 0; i < allRows.length; i++) {
    var val = allRows[i][colIdx];
    if (val === '' || val === null || val === undefined) continue;
    totalCount++;
    var num = parseFloat(String(val).replace(/[,\s%￥$€£¥]/g, ''));
    if (!isNaN(num)) numericCount++;
  }
  return totalCount > 0 && (numericCount / totalCount) > 0.5;
}

function getColumnData(colIdx) {
  var nums = [];
  for (var i = 0; i < allRows.length; i++) {
    var val = allRows[i][colIdx];
    if (val === '' || val === null || val === undefined) continue;
    var num = parseFloat(String(val).replace(/[,\s%￥$€£¥]/g, ''));
    if (!isNaN(num)) nums.push(num);
  }
  return nums;
}

// ─── Statistics ──────────────────────────────────────
function calcMean(arr) {
  return arr.reduce(function(s, v) { return s + v; }, 0) / arr.length;
}

function calcMedian(arr) {
  var sorted = arr.slice().sort(function(a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function calcMode(arr) {
  var freq = {};
  arr.forEach(function(v) {
    var key = parseFloat(v.toFixed(2));
    freq[key] = (freq[key] || 0) + 1;
  });
  var maxFreq = 0;
  var modes = [];
  for (var key in freq) {
    if (freq.hasOwnProperty(key)) {
      if (freq[key] > maxFreq) {
        maxFreq = freq[key];
        modes = [parseFloat(key)];
      } else if (freq[key] === maxFreq) {
        modes.push(parseFloat(key));
      }
    }
  }
  if (maxFreq <= 1) return { values: null, count: 0 };
  return { values: modes.slice(0, 3), count: maxFreq };
}

function calcStdDev(arr) {
  var mean = calcMean(arr);
  var variance = arr.reduce(function(s, v) { return s + (v - mean) * (v - mean); }, 0) / arr.length;
  return Math.sqrt(variance);
}

function formatNum(n) {
  if (Number.isInteger(n)) return n.toLocaleString();
  return parseFloat(n.toFixed(4)).toString();
}

// ─── Row Count Selector ──────────────────────────────
function setRowCount(n) {
  displayRowCount = n;
  // Update segmented button active states
  document.querySelectorAll('.rc-btn').forEach(function(btn) {
    btn.classList.remove('rc-active');
  });
  var idMap = { '10': 'rc-10', '50': 'rc-50', '100': 'rc-100', '500': 'rc-500', '-1': 'rc-all' };
  var activeId = idMap[String(n)];
  if (activeId) {
    var btn = document.getElementById(activeId);
    if (btn) btn.classList.add('rc-active');
  }
  // Clear custom input if a preset was clicked
  if (n !== -2) {
    document.getElementById('custom-row-count').value = '';
  }
  // Re-render
  renderTable();
}

function onCustomRowCount() {
  var input = document.getElementById('custom-row-count');
  var val = parseInt(input.value);
  if (!isNaN(val) && val > 0) {
    displayRowCount = val;
    document.querySelectorAll('.rc-btn').forEach(function(btn) { btn.classList.remove('rc-active'); });
    renderTable();
  }
}

// ─── Display ─────────────────────────────────────────
function onSelectionChange() {
  var idx = parseInt(document.getElementById('column-select').value);
  currentColIndex = idx;
  var data = getAnalysisData(idx);

  if (data.length < 2) {
    document.getElementById('stat-mean').textContent = '--';
    document.getElementById('stat-median').textContent = '--';
    document.getElementById('stat-mode').textContent = '--';
    document.getElementById('stat-stddev').textContent = '--';
    document.getElementById('stat-count').textContent = data.length;
    document.getElementById('stat-min').textContent = data.length > 0 ? formatNum(data[0]) : '--';
    document.getElementById('stat-max').textContent = data.length > 0 ? formatNum(data[data.length - 1]) : '--';
    var warnKey = analysisMode === 'column' ? 'validation.noNumeric' : 'validation.noNumericRow';
    document.getElementById('col-type-badge').textContent = t(warnKey);
    document.getElementById('col-type-badge').className = 'px-2.5 py-1 rounded-full text-xs font-medium bg-red-50 text-red-600';
    renderTable();
    return;
  }

  document.getElementById('col-type-badge').textContent = t('analysis.numeric') + ' (' + data.length + ' ' + t('stat.count').toLowerCase() + ')';
  document.getElementById('col-type-badge').className = 'px-2.5 py-1 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700';

  var mean = calcMean(data);
  var median = calcMedian(data);
  var modeResult = calcMode(data);
  var stddev = calcStdDev(data);
  var sorted = data.slice().sort(function(a, b) { return a - b; });

  document.getElementById('stat-mean').textContent = formatNum(mean);
  document.getElementById('stat-median').textContent = formatNum(median);
  document.getElementById('stat-mode').textContent = modeResult.values
    ? modeResult.values.map(formatNum).join(', ')
    : t('stat.noMode');
  document.getElementById('stat-stddev').textContent = formatNum(stddev);
  document.getElementById('stat-count').textContent = data.length;
  document.getElementById('stat-min').textContent = formatNum(sorted[0]);
  document.getElementById('stat-max').textContent = formatNum(sorted[sorted.length - 1]);

  renderTable();
  lucide.createIcons();
}

function refreshStatsDisplay() {
  onSelectionChange();
  updateAllI18n();
}

function renderTable() {
  var thead = document.querySelector('#data-table thead');
  var tbody = document.querySelector('#data-table tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (allHeaders.length === 0) return;

  var tr = document.createElement('tr');
  var thNum = document.createElement('th');
  thNum.className = 'px-4 py-3 text-left text-xs font-semibold text-gray-500 uppercase tracking-wider w-12';
  thNum.textContent = '#';
  tr.appendChild(thNum);
  allHeaders.forEach(function(h) {
    var th = document.createElement('th');
    th.className = 'px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider whitespace-nowrap';
    th.textContent = h;
    tr.appendChild(th);
  });
  thead.appendChild(tr);

  var maxRows = displayRowCount === -1 ? allRows.length : Math.min(allRows.length, displayRowCount);
  for (var i = 0; i < maxRows; i++) {
    var rowTr = document.createElement('tr');
    rowTr.className = 'hover:bg-gray-50/50';
    var tdNum = document.createElement('td');
    tdNum.className = 'px-4 py-2.5 text-xs text-gray-400';
    tdNum.textContent = i + 1;
    rowTr.appendChild(tdNum);
    allHeaders.forEach(function(_, ci) {
      var td = document.createElement('td');
      td.className = 'px-4 py-2.5 text-sm text-gray-700 whitespace-nowrap max-w-xs truncate';
      var val = allRows[i] ? allRows[i][ci] : '';
      td.textContent = val !== undefined && val !== null ? String(val) : '';
      if (analysisMode === 'column' && ci === currentColIndex) {
        td.className += ' bg-indigo-50/30 font-medium';
      }
      rowTr.appendChild(td);
    });
    if (analysisMode === 'row' && i === currentColIndex) {
      rowTr.className += ' bg-indigo-50/30 font-medium';
    }
    tbody.appendChild(rowTr);
  }

  var infoEl = document.getElementById('table-info');
  if (allRows.length > maxRows) {
    infoEl.textContent = t('table.showing').replace('{n}', String(maxRows)).replace('{total}', String(allRows.length));
  } else {
    infoEl.textContent = allRows.length + ' ' + t('table.rows') + ' · ' + allHeaders.length + ' ' + t('table.cols');
  }
}

function resetAnalysis() {
  allHeaders = [];
  allRows = [];
  workbookData = null;
  currentColIndex = -1;
  displayRowCount = 100;
  analysisMode = 'column';
  document.getElementById('custom-row-count').value = '';
  setRowCount(100);
  document.getElementById('mode-col').classList.add('rc-active');
  document.getElementById('mode-row').classList.remove('rc-active');
  var label = document.getElementById('selector-label');
  label.setAttribute('data-i18n', 'analysis.selectCol');
  label.textContent = t('analysis.selectCol');
  document.getElementById('analysis-section').classList.add('hidden');
  document.getElementById('upload-section').style.opacity = '1';
  document.getElementById('file-info').classList.add('hidden');
  document.getElementById('file-input').value = '';
  document.querySelector('#data-table thead').innerHTML = '';
  document.querySelector('#data-table tbody').innerHTML = '';
  document.getElementById('table-info').textContent = '';
  document.getElementById('stat-mean').textContent = '--';
  document.getElementById('stat-median').textContent = '--';
  document.getElementById('stat-mode').textContent = '--';
  document.getElementById('stat-stddev').textContent = '--';
  document.getElementById('stat-count').textContent = '--';
  document.getElementById('stat-min').textContent = '--';
  document.getElementById('stat-max').textContent = '--';
  document.getElementById('column-select').innerHTML = '';
}

// ─── PDF Export (with html2canvas for CJK support) ───
function exportPDF() {
  if (!workbookData || currentColIndex < 0) return;
  if (typeof html2canvas === 'undefined') {
    alert('html2canvas not loaded. Please check your network connection.');
    return;
  }

  var exportBtn = document.getElementById('export-btn');
  var originalText = exportBtn ? exportBtn.textContent : '';
  if (exportBtn) {
    exportBtn.textContent = t('report.exporting');
    exportBtn.disabled = true;
  }

  // Build report HTML
  var container = buildReportElement();

  // Position off-screen
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.top = '0';
  container.style.zIndex = '-1';
  document.body.appendChild(container);

  // Wait for fonts to settle, then capture
  setTimeout(function() {
    html2canvas(container, {
      scale: 2,
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false,
    }).then(function(canvas) {
      document.body.removeChild(container);

      var pdfW = 297;  // A4 landscape mm
      var pdfH = 210;
      var margin = 12;
      var contentW = pdfW - margin * 2;
      var contentH = pdfH - margin * 2;

      // Image size in mm at the PDF scale
      var imgW_mm = contentW;
      var imgH_mm = (canvas.height / canvas.width) * imgW_mm;

      var doc = new jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

      var srcY_px = 0;
      var pageNum = 0;

      while (srcY_px < canvas.height) {
        if (pageNum > 0) doc.addPage();

        var remainingH_mm = imgH_mm - (srcY_px / canvas.height) * imgH_mm;
        var sliceH_mm = Math.min(contentH, remainingH_mm);
        var sliceH_px = Math.ceil((sliceH_mm / imgH_mm) * canvas.height);

        // Create slice canvas
        var sliceCanvas = document.createElement('canvas');
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = sliceH_px;
        var ctx = sliceCanvas.getContext('2d');
        ctx.drawImage(canvas, 0, srcY_px, canvas.width, sliceH_px, 0, 0, canvas.width, sliceH_px);

        var imgData = sliceCanvas.toDataURL('image/png');
        doc.addImage(imgData, 'PNG', margin, margin, imgW_mm, sliceH_mm, undefined, 'FAST');

        // Page number footer
        doc.setFontSize(7);
        doc.setTextColor(160, 160, 160);
        doc.text(String(pageNum + 1), pdfW - margin, pdfH - 4, { align: 'right' });

        srcY_px += sliceH_px;
        pageNum++;
      }

      var safeName = workbookData.fileName.replace(/\.[^.]+$/, '');
      doc.save(safeName + '_analysis_report.pdf');

      if (exportBtn) {
        exportBtn.textContent = originalText;
        exportBtn.disabled = false;
      }
    }).catch(function(err) {
      console.error(err);
      document.body.removeChild(container);
      if (exportBtn) {
        exportBtn.textContent = originalText;
        exportBtn.disabled = false;
      }
      alert('PDF generation failed: ' + err.message);
    });
  }, 300);
}

function buildReportElement() {
  var data = getAnalysisData(currentColIndex);
  var mean = calcMean(data);
  var median = calcMedian(data);
  var modeResult = calcMode(data);
  var stddev = calcStdDev(data);
  var sorted = data.slice().sort(function(a, b) { return a - b; });
  var modeText = modeResult.values ? modeResult.values.map(formatNum).join(', ') : t('stat.noMode');

  // Build a readable label for what was analyzed
  var analyzedLabel;
  if (analysisMode === 'column') {
    analyzedLabel = escHtml(allHeaders[currentColIndex]);
  } else {
    analyzedLabel = '#' + (currentColIndex + 1);
    var firstVal = allRows[currentColIndex] ? allRows[currentColIndex][0] : '';
    if (firstVal !== '' && firstVal !== null && firstVal !== undefined) {
      analyzedLabel += ' · ' + escHtml(String(firstVal).substring(0, 30));
    }
  }

  var div = document.createElement('div');
  div.className = 'report-container';

  var html = '';

  // Title
  html += '<div class="report-title">LightAnalyzer — ' + t('report.title') + '</div>';
  html += '<div class="report-divider"></div>';

  // Metadata
  html += '<div class="report-meta">';
  html += '<span><b>' + t('report.file') + ':</b> ' + escHtml(workbookData.fileName) + '</span>';
  html += '<span><b>' + (analysisMode === 'column' ? t('report.column') : t('report.row')) + ':</b> ' + analyzedLabel + '</span>';
  html += '<span><b>' + t('report.date') + ':</b> ' + new Date().toLocaleString() + '</span>';
  html += '</div>';
  html += '<div class="report-divider"></div>';

  // Stats section
  html += '<div class="report-section-title">' + t('report.summary') + '</div>';

  // Four main stat cards
  html += '<div class="report-stats-grid">';
  html += buildStatCard('card-blue', t('stat.mean'), formatNum(mean), t('stat.meanDesc'));
  html += buildStatCard('card-green', t('stat.median'), formatNum(median), t('stat.medianDesc'));
  html += buildStatCard('card-purple', t('stat.mode'), modeText, t('stat.modeDesc'));
  html += buildStatCard('card-amber', t('stat.stddev'), formatNum(stddev), t('stat.stddevDesc'));
  html += '</div>';

  // Extra stats row
  html += '<div class="report-extra-row">';
  html += buildExtraItem(t('stat.count'), String(data.length));
  html += buildExtraItem(t('stat.min'), formatNum(sorted[0]));
  html += buildExtraItem(t('stat.max'), formatNum(sorted[sorted.length - 1]));
  html += '</div>';

  html += '<div class="report-divider"></div>';

  // Data table
  html += '<div class="report-section-title">' + t('report.dataTable') + '</div>';
  html += '<div class="report-table-wrap">';
  html += '<table class="report-table"><thead><tr><th>#</th>';

  allHeaders.forEach(function(h) {
    html += '<th>' + escHtml(h) + '</th>';
  });
  html += '</tr></thead><tbody>';

  var pdfMaxRows = displayRowCount === -1 ? allRows.length : Math.min(allRows.length, displayRowCount);
  for (var i = 0; i < pdfMaxRows; i++) {
    html += '<tr><td>' + (i + 1) + '</td>';
    allHeaders.forEach(function(_, ci) {
      var val = allRows[i] ? allRows[i][ci] : '';
      var str = val !== undefined && val !== null ? String(val) : '';
      html += '<td>' + escHtml(str) + '</td>';
    });
    html += '</tr>';
  }

  if (allRows.length > pdfMaxRows) {
    html += '<tr><td colspan="' + (allHeaders.length + 1) + '" style="text-align:center;color:#9ca3af;padding:12px;">';
    html += t('table.showing').replace('{n}', String(pdfMaxRows)).replace('{total}', String(allRows.length));
    html += '</td></tr>';
  }

  html += '</tbody></table></div>';

  // Footer
  html += '<div class="report-divider"></div>';
  html += '<div style="font-size:11px;color:#9ca3af;text-align:center;">LightAnalyzer · ' + t('report.title') + ' · ' + new Date().toLocaleDateString() + '</div>';

  div.innerHTML = html;
  return div;
}

function buildStatCard(colorClass, label, value, desc) {
  return '<div class="report-stat-card ' + colorClass + '">' +
    '<div class="rsc-label">' + escHtml(label) + '</div>' +
    '<div class="rsc-value">' + escHtml(value) + '</div>' +
    '<div style="font-size:10px;color:#9ca3af;margin-top:4px;">' + escHtml(desc) + '</div>' +
    '</div>';
}

function buildExtraItem(label, value) {
  return '<div class="report-extra-item">' +
    '<div class="rei-label">' + escHtml(label) + '</div>' +
    '<div class="rei-value">' + escHtml(value) + '</div>' +
    '</div>';
}

function escHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ─── Batch PDF Export (all columns / all rows) ─────
function exportBatchPDF() {
  if (!workbookData) return;
  if (typeof html2canvas === 'undefined') {
    alert('html2canvas not loaded. Please check your network connection.');
    return;
  }

  var batchBtn = document.getElementById('export-batch-btn');
  var originalText = batchBtn ? batchBtn.textContent : '';
  if (batchBtn) {
    batchBtn.textContent = t('report.exporting');
    batchBtn.disabled = true;
  }

  var container = buildBatchReportElement();
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.top = '0';
  container.style.zIndex = '-1';
  document.body.appendChild(container);

  setTimeout(function() {
    html2canvas(container, {
      scale: 2,
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false,
    }).then(function(canvas) {
      document.body.removeChild(container);

      var pdfW = 297;
      var pdfH = 210;
      var margin = 12;
      var contentW = pdfW - margin * 2;
      var contentH = pdfH - margin * 2;
      var imgW_mm = contentW;
      var imgH_mm = (canvas.height / canvas.width) * imgW_mm;

      var doc = new jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      var srcY_px = 0;
      var pageNum = 0;

      while (srcY_px < canvas.height) {
        if (pageNum > 0) doc.addPage();
        var remainingH_mm = imgH_mm - (srcY_px / canvas.height) * imgH_mm;
        var sliceH_mm = Math.min(contentH, remainingH_mm);
        var sliceH_px = Math.ceil((sliceH_mm / imgH_mm) * canvas.height);

        var sliceCanvas = document.createElement('canvas');
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = sliceH_px;
        var ctx = sliceCanvas.getContext('2d');
        ctx.drawImage(canvas, 0, srcY_px, canvas.width, sliceH_px, 0, 0, canvas.width, sliceH_px);

        doc.addImage(sliceCanvas.toDataURL('image/png'), 'PNG', margin, margin, imgW_mm, sliceH_mm, undefined, 'FAST');
        doc.setFontSize(7);
        doc.setTextColor(160, 160, 160);
        doc.text(String(pageNum + 1), pdfW - margin, pdfH - 4, { align: 'right' });

        srcY_px += sliceH_px;
        pageNum++;
      }

      var safeName = workbookData.fileName.replace(/\.[^.]+$/, '');
      doc.save(safeName + '_batch_analysis_report.pdf');

      if (batchBtn) {
        batchBtn.textContent = originalText;
        batchBtn.disabled = false;
      }
    }).catch(function(err) {
      console.error(err);
      document.body.removeChild(container);
      if (batchBtn) {
        batchBtn.textContent = originalText;
        batchBtn.disabled = false;
      }
      alert('PDF generation failed: ' + err.message);
    });
  }, 300);
}

function buildBatchReportElement() {
  var div = document.createElement('div');
  div.className = 'report-container';

  var html = '';

  // Title
  html += '<div class="report-title">LightAnalyzer — ' + t('report.batchTitle') + '</div>';
  html += '<div class="report-divider"></div>';

  // Metadata
  var modeLabel = analysisMode === 'column' ? t('analysis.byColumn') : t('analysis.byRow');
  html += '<div class="report-meta">';
  html += '<span><b>' + t('report.file') + ':</b> ' + escHtml(workbookData.fileName) + '</span>';
  html += '<span><b>' + t('report.mode') + ':</b> ' + modeLabel + '</span>';
  html += '<span><b>' + t('report.date') + ':</b> ' + new Date().toLocaleString() + '</span>';
  html += '</div>';
  html += '<div class="report-divider"></div>';

  // Build stats for each column or row
  var entries = [];
  if (analysisMode === 'column') {
    for (var ci = 0; ci < allHeaders.length; ci++) {
      var colData = getColumnData(ci);
      entries.push({
        label: allHeaders[ci],
        data: colData,
      });
    }
  } else {
    for (var ri = 0; ri < allRows.length; ri++) {
      var rowData = getRowData(ri);
      var rowLabel = '#' + (ri + 1);
      var firstVal = allRows[ri] ? allRows[ri][0] : '';
      if (firstVal !== '' && firstVal !== null && firstVal !== undefined) {
        rowLabel += ' · ' + String(firstVal).substring(0, 20);
      }
      entries.push({
        label: rowLabel,
        data: rowData,
      });
    }
  }

  // Summary table
  var entityLabel = analysisMode === 'column' ? t('report.column') : t('report.row');
  html += '<div class="report-section-title">' + t('report.summary') + '（' + t('display.all') + ' ' + entityLabel + '）</div>';
  html += '<div class="report-table-wrap">';
  html += '<table class="report-table"><thead><tr>';
  html += '<th>' + entityLabel + '</th>';
  html += '<th>' + t('stat.count') + '</th>';
  html += '<th>' + t('stat.mean') + '</th>';
  html += '<th>' + t('stat.median') + '</th>';
  html += '<th>' + t('stat.mode') + '</th>';
  html += '<th>' + t('stat.stddev') + '</th>';
  html += '<th>' + t('stat.min') + '</th>';
  html += '<th>' + t('stat.max') + '</th>';
  html += '</tr></thead><tbody>';

  for (var e = 0; e < entries.length; e++) {
    var entry = entries[e];
    var d = entry.data;
    if (d.length < 2) {
      html += '<tr><td>' + escHtml(entry.label) + '</td>';
      html += '<td>' + d.length + '</td>';
      html += '<td colspan="6" style="color:#9ca3af;">' + t('validation.noNumeric') + '</td></tr>';
      continue;
    }
    var m = calcMean(d);
    var med = calcMedian(d);
    var modR = calcMode(d);
    var sd = calcStdDev(d);
    var srt = d.slice().sort(function(a, b) { return a - b; });
    var modeStr = modR.values ? modR.values.map(formatNum).join(', ') : t('stat.noMode');

    html += '<tr>';
    html += '<td><b>' + escHtml(entry.label) + '</b></td>';
    html += '<td>' + d.length + '</td>';
    html += '<td>' + formatNum(m) + '</td>';
    html += '<td>' + formatNum(med) + '</td>';
    html += '<td>' + modeStr + '</td>';
    html += '<td>' + formatNum(sd) + '</td>';
    html += '<td>' + formatNum(srt[0]) + '</td>';
    html += '<td>' + formatNum(srt[srt.length - 1]) + '</td>';
    html += '</tr>';
  }

  html += '</tbody></table></div>';
  html += '<div class="report-divider"></div>';

  // Data table
  html += '<div class="report-section-title">' + t('report.dataTable') + '</div>';
  html += '<div class="report-table-wrap">';
  html += '<table class="report-table"><thead><tr><th>#</th>';
  allHeaders.forEach(function(h) { html += '<th>' + escHtml(h) + '</th>'; });
  html += '</tr></thead><tbody>';

  var maxRows = displayRowCount === -1 ? allRows.length : Math.min(allRows.length, displayRowCount);
  for (var i = 0; i < maxRows; i++) {
    html += '<tr><td>' + (i + 1) + '</td>';
    allHeaders.forEach(function(_, ci) {
      var val = allRows[i] ? allRows[i][ci] : '';
      html += '<td>' + escHtml(String(val !== undefined && val !== null ? val : '')) + '</td>';
    });
    html += '</tr>';
  }

  if (allRows.length > maxRows) {
    html += '<tr><td colspan="' + (allHeaders.length + 1) + '" style="text-align:center;color:#9ca3af;padding:12px;">';
    html += t('table.showing').replace('{n}', String(maxRows)).replace('{total}', String(allRows.length));
    html += '</td></tr>';
  }

  html += '</tbody></table></div>';
  html += '<div class="report-divider"></div>';
  html += '<div style="font-size:11px;color:#9ca3af;text-align:center;">LightAnalyzer · ' + t('report.batchTitle') + ' · ' + new Date().toLocaleDateString() + '</div>';

  div.innerHTML = html;
  return div;
}

// ─── Init ────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function() {
  setupDragDrop();
  updateAllI18n();
  updateNavActive();
  updateLangMenuUI();
  lucide.createIcons();
});

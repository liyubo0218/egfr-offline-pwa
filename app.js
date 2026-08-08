
/* global XLSX */
if (typeof XLSX === 'undefined') {
  console.error('SheetJS XLSX 未加载。请检查 vendor/xlsx.full.min.js 是否存在且能正常加载。');
}
const $ = (s) => document.querySelector(s);
const ui = { fileInput: $('#fileInput'), dropZone: $('#dropZone'), fileName: $('#fileName'), genes: $('#geneOptions'), configState: $('#configState'), start: $('#startButton'), progressPanel: $('#progressPanel'), progressText: $('#progressText'), progressBar: $('#progressBar'), statusText: $('#statusText'), summaryPanel: $('#summaryPanel'), summaryContent: $('#summaryContent'), downloadsPanel: $('#downloadsPanel'), downloadButtons: $('#downloadButtons') };
let selectedFile = null;
let configRegistry = []; // 仅含 {id, name, file}，不含关键词，用于渲染基因勾选列表
const configCache = new Map(); // id -> 完整 config（懒加载，仅加载用户勾选过的基因）

/* 以下七个函数按用户提供的 build_egfr.mjs 原样迁移（仅把关键词参数化）。 */
function normalizeText(value) {
  if (value == null) return '';
  return String(value).replace(/\u3000/g, ' ').replace(/[－–—]/g, '-').replace(/\s+/g, '').toUpperCase();
}
function displayText(value) { if (value == null) return ''; return String(value).trim(); }
function columnLetter(index) { let n = index + 1; let out = ''; while (n > 0) { const rem = (n - 1) % 26; out = String.fromCharCode(65 + rem) + out; n = Math.floor((n - 1) / 26); } return out; }
function findHeaderRow(values) { const searchRows = Math.min(values.length, 20); for (let rowIndex = 0; rowIndex < searchRows; rowIndex += 1) { const headers = values[rowIndex].map(displayText); const impressionIndex = headers.findIndex((header) => header === '印象'); if (impressionIndex >= 0) return { rowIndex, headers, impressionIndex }; } return null; }
function keywordMatch(text, keyword) { return text.includes(normalizeText(keyword)); }
function rowMatches(row, impressionIndex, includeKeywords, excludeKeywords, configOverride) {
  const config = configOverride || {
    includeKeywords: includeKeywords || [],
    excludeKeywords: excludeKeywords || []
  };
  const text = Object.values(row).map(v => String(v ?? "")).join(" ");

  const anyGroups = Array.isArray(config.includeKeywordGroupsAny)
    ? config.includeKeywordGroupsAny : [];
  if (anyGroups.length) {
    const matched = anyGroups.some(group =>
      Array.isArray(group) && group.length &&
      group.every(k => text.includes(String(k)))
    );
    if (!matched) return false;
  } else {
    const groups = Array.isArray(config.includeKeywordGroups)
      ? config.includeKeywordGroups : [];
    if (groups.length) {
      const matched = groups.every(group =>
        Array.isArray(group) && group.length &&
        group.some(k => text.includes(String(k)))
      );
      if (!matched) return false;
    } else {
      const includes = config.includeKeywords || [];
      if (includes.length && !includes.some(k => text.includes(String(k)))) return false;
    }
  }

  const excludes = config.excludeKeywords || [];
  if (excludes.some(k => text.includes(String(k)))) return false;

  return true;
}
function nonEmptyRow(row) { return row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== ''); }

/* 启动时只读取 configs/index.json 里的 {id, name} 用于渲染勾选列表。
   不在此处拉取任何单个基因的完整配置（含关键词），避免“遍历全部 configs 做筛选”。 */
async function loadConfigurations() {
  configRegistry = await (await fetch('configs/index.json')).json();
  ui.genes.replaceChildren(...configRegistry.map((entry) => { const label = document.createElement('label'); label.className = 'gene-option'; label.innerHTML = `<input type="checkbox" value="${escapeHtml(entry.id)}"><span>${escapeHtml(entry.name)}</span>`; return label; }));
  ui.configState.textContent = `已加载 ${configRegistry.length} 个基因`;
  ui.genes.addEventListener('change', updateStartState); updateStartState();
}
function validateConfig(config) {
  for (const key of ['id', 'name', 'outputFile', 'includeKeywords', 'excludeKeywords']) if (!(key in config)) throw new Error(`配置缺少 ${key}`);
  if (!Array.isArray(config.includeKeywords) || !Array.isArray(config.excludeKeywords)) throw new Error('关键词必须是数组');
  return config;
}
/* 仅按需（懒加载）拉取用户勾选的那一个基因的完整配置（含 includeKeywords/excludeKeywords）。
   未勾选的基因（如 ALK、KRAS、HER2、ROS1、RET、MET、BRAF……）的配置文件完全不会被读取。 */
async function loadConfigById(id) {
  if (configCache.has(id)) return configCache.get(id);
  const entry = configRegistry.find((item) => item.id === id);
  if (!entry) throw new Error(`未找到基因配置：${id}`);
  const config = validateConfig(await (await fetch(`configs/${entry.file}`)).json());
  configCache.set(id, config);
  return config;
}
function selectedGeneIds() { return [...ui.genes.querySelectorAll('input:checked')].map((input) => input.value); }
function updateStartState() { ui.start.disabled = !selectedFile || selectedGeneIds().length === 0; }
function setProgress(value, text) { ui.progressPanel.classList.remove('hidden'); ui.progressBar.style.width = `${value}%`; ui.progressText.textContent = `${Math.round(value)}%`; ui.statusText.textContent = text; }
function setFile(file) { if (!file) return; if (!/\.xlsx$/i.test(file.name)) { setProgress(0, '请选择 .xlsx 文件。'); return; } selectedFile = file; ui.fileName.textContent = file.name; updateStartState(); }

/* 单基因模式：一次调用只处理一个 config，产出一个只含“筛选结果”单一工作表的 Workbook。
   禁止在此函数内部循环 append 多个 Sheet，也不会混入其它基因的数据。
   逐工作表执行原脚本的行扫描、includeHits、excludedByKeyword 与匹配输出（数据来源是用户上传的原始 Excel 的多个 Sheet，
   但输出始终只汇总进同一个“筛选结果”工作表，与本函数只处理单一 config 是两回事，不要混淆）。 */
function buildSingleGeneWorkbook(sourceWorkbook, config, onProgress) {
  const workbook = XLSX.utils.book_new();
  const outputRows = []; const sheetSummaries = []; let outputHeaders = null;
  const { includeKeywords, excludeKeywords } = config;
  sourceWorkbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sourceSheet = sourceWorkbook.Sheets[sheetName];
    const values = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, defval: null, raw: true });
    const header = findHeaderRow(values);
    if (!header) { sheetSummaries.push({ sheet: sheetName, headerRow: null, impressionColumn: null, rowsScanned: 0, includeHits: 0, excludedByKeyword: 0, matched: 0, skipped: true }); onProgress((sheetIndex + 1) / sourceWorkbook.SheetNames.length); return; }
    const { rowIndex: headerRowIndex, headers, impressionIndex } = header;
    if (!outputHeaders) outputHeaders = headers;
    let includeHits = 0; let excludedByKeyword = 0; let matched = 0;
    for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      if (!nonEmptyRow(row)) continue;
      const impression = normalizeText(row[impressionIndex]);
      const includes = includeKeywords.some((keyword) => keywordMatch(impression, keyword));
      const excludes = excludeKeywords.some((keyword) => keywordMatch(impression, keyword));
      if (includes) includeHits += 1;
      if (includes && excludes) excludedByKeyword += 1;
      if (rowMatches(row, impressionIndex, includeKeywords, excludeKeywords)) { matched += 1; outputRows.push(row.slice(0, outputHeaders.length)); }
    }
    sheetSummaries.push({ sheet: sheetName, headerRow: headerRowIndex + 1, impressionColumn: columnLetter(impressionIndex), rowsScanned: values.length - headerRowIndex - 1, includeHits, excludedByKeyword, matched, skipped: false });
    onProgress((sheetIndex + 1) / sourceWorkbook.SheetNames.length);
  });
  if (!outputHeaders) throw new Error('未在工作簿前 20 行找到名为“印象”的列。');

  // 将所有原始 Sheet 中筛选出的数据合并后彻底随机打乱。
  // Fisher-Yates 洗牌：不再保留原 Excel Sheet 的分类/先后顺序。
  for (let i = outputRows.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [outputRows[i], outputRows[j]] = [outputRows[j], outputRows[i]];
  }

  const resultSheet = XLSX.utils.aoa_to_sheet([outputHeaders, ...outputRows]);
  decorateSheet(resultSheet, outputHeaders.length, outputRows.length + 1);
  // 整个 Workbook 生命周期内，book_append_sheet 只调用这一次；不允许出现循环 append 多个 Sheet 的写法。
  XLSX.utils.book_append_sheet(workbook, resultSheet, "筛选结果");
  if (workbook.SheetNames.length !== 1 || workbook.SheetNames[0] !== '筛选结果') {
    throw new Error('内部错误：单基因导出必须只包含一个“筛选结果”工作表。');
  }
  return { config, workbook, sheetSummaries };
}
function displayWidth(value) { return [...String(value ?? '')].reduce((sum, char) => sum + (char.charCodeAt(0) > 255 ? 2 : 1), 0); }
function decorateSheet(sheet, columnCount, rowCount) {
  sheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, rowCount - 1), c: Math.max(0, columnCount - 1) } }) };
  sheet['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
  sheet['!cols'] = Array.from({ length: columnCount }, (_, c) => { let width = 10; for (let r = 0; r < rowCount; r += 1) width = Math.max(width, displayWidth(sheet[XLSX.utils.encode_cell({ r, c })]?.v)); return { wch: Math.min(width + 2, 86) }; });
}

async function processWorkbook() {
  const geneIds = selectedGeneIds(); if (!selectedFile || !geneIds.length) return;
  ui.start.disabled = true; ui.summaryPanel.classList.add('hidden'); ui.downloadsPanel.classList.add('hidden');
  try {
    setProgress(3, '正在读取 Excel…'); const sourceWorkbook = XLSX.read(await selectedFile.arrayBuffer(), { type: 'array', cellText: true, cellDates: true }); const results = [];
    for (let i = 0; i < geneIds.length; i += 1) {
      // 只加载当前这一个被勾选的基因的 config；未勾选的基因配置不会在这里被读取。
      const config = await loadConfigById(geneIds[i]);
      const result = buildSingleGeneWorkbook(sourceWorkbook, config, (part) => setProgress(5 + ((i + part) / geneIds.length) * 85, `正在筛选 ${config.name}…`));
      results.push(result);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    renderResults(results); setProgress(100, '处理完成，可下载结果。');
  } catch (error) { console.error(error); setProgress(0, `处理失败：${error.message || '无法读取此文件'}`); }
  finally { updateStartState(); }
}
function totals(stats) { return stats.reduce((total, stat) => ({ scanned: total.scanned + stat.rowsScanned, matched: total.matched + stat.matched, excluded: total.excluded + stat.excludedByKeyword }), { scanned: 0, matched: 0, excluded: 0 }); }
function renderResults(results) {
  ui.summaryContent.replaceChildren(...results.map((result) => { const t = totals(result.sheetSummaries); const block = document.createElement('div'); block.innerHTML = `<h3>${escapeHtml(result.config.name)}</h3><table class="stats"><thead><tr><th>工作表</th><th>扫描</th><th>命中</th><th>排除</th></tr></thead><tbody>${result.sheetSummaries.map((s) => `<tr><td>${escapeHtml(s.sheet)}${s.skipped ? '（跳过）' : ''}</td><td>${s.rowsScanned}</td><td>${s.matched}</td><td>${s.excludedByKeyword}</td></tr>`).join('')}<tr class="total-row"><td>总计</td><td>${t.scanned}</td><td>${t.matched}</td><td>${t.excluded}</td></tr></tbody></table>`; return block; }));
  ui.summaryPanel.classList.remove('hidden'); ui.downloadButtons.replaceChildren();
  results.forEach((result) => addDownload(result.config.outputFile, result.workbook)); addDownload('筛选总统计.xlsx', summaryWorkbook(results)); ui.downloadsPanel.classList.remove('hidden');
}
function summaryWorkbook(results) { const rows = [['基因', '工作表', '表头行', '印象列', '扫描数量', '包含命中', '排除数量', '最终命中', '状态']]; results.forEach((result) => result.sheetSummaries.forEach((s) => rows.push([result.config.name, s.sheet, s.headerRow ?? '', s.impressionColumn ?? '', s.rowsScanned, s.includeHits, s.excludedByKeyword, s.matched, s.skipped ? '跳过：未找到印象列' : '已处理']))); const book = XLSX.utils.book_new(); const sheet = XLSX.utils.aoa_to_sheet(rows); decorateSheet(sheet, 9, rows.length); XLSX.utils.book_append_sheet(book, sheet, '总统计'); return book; }
function addDownload(name, workbook) { const button = document.createElement('button'); button.className = 'download'; button.type = 'button'; button.textContent = `下载 ${name}`; button.addEventListener('click', () => XLSX.writeFile(workbook, name, { compression: true })); ui.downloadButtons.append(button); }
function escapeHtml(value) { return String(value).replace(/[&<>'"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c])); }

ui.dropZone.addEventListener('click', () => ui.fileInput.click()); ui.dropZone.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') ui.fileInput.click(); }); ui.fileInput.addEventListener('change', (e) => setFile(e.target.files[0]));
['dragenter', 'dragover'].forEach((type) => ui.dropZone.addEventListener(type, (e) => { e.preventDefault(); ui.dropZone.classList.add('dragging'); })); ['dragleave', 'drop'].forEach((type) => ui.dropZone.addEventListener(type, (e) => { e.preventDefault(); ui.dropZone.classList.remove('dragging'); })); ui.dropZone.addEventListener('drop', (e) => setFile(e.dataTransfer.files[0])); ui.start.addEventListener('click', processWorkbook);
if ('serviceWorker' in navigator) window.addEventListener('load', () => navigator.serviceWorker.register('service-worker.js'));
loadConfigurations().catch((error) => { console.error(error); ui.configState.textContent = '配置读取失败'; setProgress(0, '无法读取基因配置，请通过 HTTP/HTTPS 服务打开本应用。'); });

/* Unified gene-screening UI */
document.addEventListener("DOMContentLoaded", async () => {
  const geneButtons = document.getElementById("geneButtons");
  const panel = document.getElementById("screeningPanel");
  const selectedGeneEl = document.getElementById("selectedGene");
  const fileInput = document.getElementById("fileInput");
  const startBtn = document.getElementById("startBtn");
  const downloadBtn = document.getElementById("downloadBtn");
  const status = document.getElementById("status");
  if (!geneButtons) return;

  let configs = [];
  let selected = null;
  let lastBlob = null;

  try {
    const idx = await fetch("configs/index.json").then(r => r.json());
    configs = await Promise.all(idx.map(async item => ({
      ...item,
      config: await fetch("configs/" + item.file).then(r => r.json())
    })));
  } catch (e) {
    status.textContent = "配置加载失败：" + e.message;
    return;
  }

  configs.forEach(item => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = item.name;
    b.addEventListener("click", () => {
      selected = item;
      selectedGeneEl.textContent = "当前筛查：" + item.name;
      panel.hidden = false;
      downloadBtn.hidden = true;
      lastBlob = null;
      status.textContent = "";
    });
    geneButtons.appendChild(b);
  });

  startBtn.addEventListener("click", async () => {
    if (!selected) return alert("请先选择筛查项目");
    const file = fileInput.files && fileInput.files[0];
    if (!file) return alert("请选择 Excel 文件");
    if (typeof XLSX === "undefined") {
      status.textContent = "处理失败：XLSX 未加载，请检查 vendor/xlsx.full.min.js";
      return;
    }

    try {
      status.textContent = "正在读取并筛查，请稍候……";
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, {type:"array"});
      const rows = [];

      wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        const arr = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        if (!arr.length) return;
        const header = arr[0];
        for (let i=1;i<arr.length;i++) {
          const vals = arr[i];
          const row = {};
          header.forEach((h,j)=>{ row[h || ("列"+(j+1))] = vals[j] ?? ""; });
          if (rowMatches(row, 0, null, null, selected.config)) rows.push(vals);
        }
      });

      const outWb = XLSX.utils.book_new();
      // Keep the first non-empty sheet's header as the common header.
      let header = null;
      for (const name of wb.SheetNames) {
        const arr = XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:""});
        if (arr.length) { header = arr[0]; break; }
      }
      header = header || [];
      for (let i=rows.length-1;i>0;i--) {
        const j = Math.floor(Math.random()*(i+1));
        [rows[i],rows[j]] = [rows[j],rows[i]];
      }
      const outWs = XLSX.utils.aoa_to_sheet([header, ...rows]);
      XLSX.utils.book_append_sheet(outWb, outWs, "筛选结果");
      const bytes = XLSX.write(outWb, {bookType:"xlsx", type:"array"});
      lastBlob = new Blob([bytes], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      downloadBtn.hidden = false;
      status.textContent = `筛查完成：共找到 ${rows.length} 条记录。`;
    } catch (e) {
      status.textContent = "处理失败：" + e.message;
    }
  });

  downloadBtn.addEventListener("click", () => {
    if (!lastBlob || !selected) return;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(lastBlob);
    a.download = selected.config.outputFile || (selected.id + ".xlsx");
    a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),1000);
  });
});

#!/usr/bin/env node
/**
 * build_gene.mjs — 单基因关键字筛选，命令行版
 *
 * 用法：
 *   node build_gene.mjs <源Excel路径> <基因配置json路径> [输出路径]
 *
 * 示例：
 *   node build_gene.mjs ./样本.xlsx ./configs/egfr.json ./EGFR.xlsx
 *   （不传输出路径时，默认用 config.outputFile，如 EGFR.xlsx，写到当前目录）
 *
 * 行为：
 *   - 从源 Excel 的所有 Sheet 中按“印象”列做关键词包含/排除筛选。
 *   - 输出 Workbook 只包含一个名为“筛选结果”的 Sheet。
 *   - 输出 Sheet 里只写入命中的行；未命中的行完全不会出现在输出文件中
 *    （不是先写全部数据再筛选/隐藏，而是从未被写入过）。
 *   - 筛选逻辑（normalizeText / findHeaderRow / keywordMatch / rowMatches / nonEmptyRow）
 *     与离线 PWA 的 app.js 保持完全一致，确保浏览器端和命令行端结果一致。
 */

import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
// 复用 PWA 内置的 SheetJS（vendor/xlsx.full.min.js），避免额外的网络依赖。
const XLSX = require(
  path.join(path.dirname(fileURLToPath(import.meta.url)), 'vendor', 'xlsx.full.min.js')
);

/* ---- 以下过滤函数与 app.js 完全一致，勿单独修改一处而漏改另一处 ---- */
function normalizeText(value) {
  if (value == null) return '';
  return String(value).replace(/\u3000/g, ' ').replace(/[－–—]/g, '-').replace(/\s+/g, '').toUpperCase();
}
function displayText(value) { if (value == null) return ''; return String(value).trim(); }
function columnLetter(index) {
  let n = index + 1; let out = '';
  while (n > 0) { const rem = (n - 1) % 26; out = String.fromCharCode(65 + rem) + out; n = Math.floor((n - 1) / 26); }
  return out;
}
function findHeaderRow(values) {
  const searchRows = Math.min(values.length, 20);
  for (let rowIndex = 0; rowIndex < searchRows; rowIndex += 1) {
    const headers = values[rowIndex].map(displayText);
    const impressionIndex = headers.findIndex((header) => header === '印象');
    if (impressionIndex >= 0) return { rowIndex, headers, impressionIndex };
  }
  return null;
}
function keywordMatch(text, keyword) { return text.includes(normalizeText(keyword)); }
function rowMatches(row, impressionIndex, includeKeywords, excludeKeywords) {
  const impression = normalizeText(row[impressionIndex]);
  if (!impression) return false;
  const includes = includeKeywords.some((keyword) => keywordMatch(impression, keyword));
  const excludes = excludeKeywords.some((keyword) => keywordMatch(impression, keyword));
  return includes && !excludes;
}
function nonEmptyRow(row) { return row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== ''); }
function displayWidth(value) { return [...String(value ?? '')].reduce((sum, char) => sum + (char.charCodeAt(0) > 255 ? 2 : 1), 0); }
function decorateSheet(sheet, columnCount, rowCount) {
  sheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, rowCount - 1), c: Math.max(0, columnCount - 1) } }) };
  sheet['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
  sheet['!cols'] = Array.from({ length: columnCount }, (_, c) => {
    let width = 10;
    for (let r = 0; r < rowCount; r += 1) width = Math.max(width, displayWidth(sheet[XLSX.utils.encode_cell({ r, c })]?.v));
    return { wch: Math.min(width + 2, 86) };
  });
}
function validateConfig(config) {
  for (const key of ['id', 'name', 'outputFile', 'includeKeywords', 'excludeKeywords']) {
    if (!(key in config)) throw new Error(`配置缺少 ${key}`);
  }
  if (!Array.isArray(config.includeKeywords) || !Array.isArray(config.excludeKeywords)) {
    throw new Error('关键词必须是数组');
  }
  return config;
}

/**
 * 核心筛选：只把命中的行放进 outputRows，未命中的行永远不会被写入输出表。
 */
function filterToMatchedRowsOnly(sourceWorkbook, config) {
  const { includeKeywords, excludeKeywords } = config;
  const outputRows = [];
  const sheetSummaries = [];
  let outputHeaders = null;

  sourceWorkbook.SheetNames.forEach((sheetName) => {
    const sourceSheet = sourceWorkbook.Sheets[sheetName];
    const values = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, defval: null, raw: true });
    const header = findHeaderRow(values);
    if (!header) {
      sheetSummaries.push({ sheet: sheetName, headerRow: null, impressionColumn: null, rowsScanned: 0, includeHits: 0, excludedByKeyword: 0, matched: 0, skipped: true });
      return;
    }
    const { rowIndex: headerRowIndex, headers, impressionIndex } = header;
    if (!outputHeaders) outputHeaders = headers;

    let includeHits = 0;
    let excludedByKeyword = 0;
    let matched = 0;
    for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      if (!nonEmptyRow(row)) continue;
      const impression = normalizeText(row[impressionIndex]);
      const includes = includeKeywords.some((keyword) => keywordMatch(impression, keyword));
      const excludes = excludeKeywords.some((keyword) => keywordMatch(impression, keyword));
      if (includes) includeHits += 1;
      if (includes && excludes) excludedByKeyword += 1;

      // 只有命中才 push；未命中的行在这里就被彻底丢弃，不会进入输出 workbook。
      if (rowMatches(row, impressionIndex, includeKeywords, excludeKeywords)) {
        matched += 1;
        outputRows.push(row.slice(0, outputHeaders.length));
      }
    }
    sheetSummaries.push({
      sheet: sheetName,
      headerRow: headerRowIndex + 1,
      impressionColumn: columnLetter(impressionIndex),
      rowsScanned: values.length - headerRowIndex - 1,
      includeHits,
      excludedByKeyword,
      matched,
      skipped: false,
    });
  });

  if (!outputHeaders) throw new Error('未在工作簿前 20 行找到名为"印象"的列。');
  return { outputHeaders, outputRows, sheetSummaries };
}

function buildOutputWorkbook(outputHeaders, outputRows) {
  const workbook = XLSX.utils.book_new();
  // 输出内容 = [表头, ...仅命中的行]。没有任何未命中的行混入其中。
  const resultSheet = XLSX.utils.aoa_to_sheet([outputHeaders, ...outputRows]);
  decorateSheet(resultSheet, outputHeaders.length, outputRows.length + 1);
  // 整个 Workbook 生命周期只 append 一次，只有一个"筛选结果"工作表。
  XLSX.utils.book_append_sheet(workbook, resultSheet, '筛选结果');
  if (workbook.SheetNames.length !== 1 || workbook.SheetNames[0] !== '筛选结果') {
    throw new Error('内部错误：导出必须只包含一个"筛选结果"工作表。');
  }
  return workbook;
}

async function main() {
  const [, , sourcePathArg, configPathArg, outputPathArg] = process.argv;
  if (!sourcePathArg || !configPathArg) {
    console.error('用法: node build_gene.mjs <源Excel路径> <基因配置json路径> [输出路径]');
    process.exit(1);
  }

  const sourcePath = path.resolve(sourcePathArg);
  const configPath = path.resolve(configPathArg);
  const config = validateConfig(JSON.parse(await fs.readFile(configPath, 'utf8')));
  const outputPath = path.resolve(outputPathArg || config.outputFile);

  const sourceBuffer = await fs.readFile(sourcePath);
  const sourceWorkbook = XLSX.read(sourceBuffer, { type: 'buffer', cellText: true, cellDates: true });

  const { outputHeaders, outputRows, sheetSummaries } = filterToMatchedRowsOnly(sourceWorkbook, config);
  const outputWorkbook = buildOutputWorkbook(outputHeaders, outputRows);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const outBuffer = XLSX.write(outputWorkbook, { type: 'buffer', bookType: 'xlsx', compression: true });
  await fs.writeFile(outputPath, outBuffer);

  console.log('FILTER_SUMMARY');
  console.log(JSON.stringify({
    sourcePath,
    outputPath,
    gene: config.name,
    rowsWritten: outputRows.length,
    columnsWritten: outputHeaders.length,
    includeKeywords: config.includeKeywords,
    excludeKeywords: config.excludeKeywords,
    sheets: sheetSummaries,
  }, null, 2));
}

await main();

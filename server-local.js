const childProcess = require('child_process');
const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const zlib = require('zlib');

const rootDir = __dirname;
const cacheDir = path.join(rootDir, '.word-memory-cache');
const stateFile = path.join(cacheDir, 'local-state.json');
const args = parseArgs(process.argv.slice(2));
const defaultExcelFile = args.excel || process.env.VOCAB_EXCEL_FILE ? path.resolve(args.excel || process.env.VOCAB_EXCEL_FILE) : '';
const selectedExcelFile = path.join(cacheDir, 'selected-workbook.xlsx');
const selectedExcelMetaFile = path.join(cacheDir, 'selected-workbook.json');
const pendingExcelFile = path.join(cacheDir, 'pending-workbook.xlsx');
const pendingExcelMetaFile = path.join(cacheDir, 'pending-workbook.json');
const host = '127.0.0.1';
const startPort = Number(args.port || process.env.VOCAB_PORT || 3765);
const defaultOnly = args['default-only'] === true || process.env.VOCAB_DEFAULT_EXCEL_ONLY === '1';

function parseArgs(values) {
  const parsed = {};
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value === '--open' || value === '--default-only') {
      parsed[value.slice(2)] = true;
    } else if (value.startsWith('--')) {
      parsed[value.slice(2)] = values[index + 1];
      index += 1;
    }
  }
  return parsed;
}

function ensureCacheDir() {
  fs.mkdirSync(cacheDir, { recursive: true });
}

function readJson(filePath, fallback) {
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch (error) {
    return fallback;
  }
}

function writeJson(filePath, value) {
  ensureCacheDir();
  fs.writeFileSync(filePath, JSON.stringify(value, null, 2), 'utf8');
}

function loadState() {
  const rawState = readJson(stateFile, {});
  return {
    wordWeightMap: isObject(rawState.wordWeightMap) ? rawState.wordWeightMap : {},
    wordKnowRecord: isObject(rawState.wordKnowRecord) ? rawState.wordKnowRecord : {},
    wordFirstUnselected: isObject(rawState.wordFirstUnselected) ? rawState.wordFirstUnselected : {},
    lastSelectNum: Number(rawState.lastSelectNum) > 0 ? Number(rawState.lastSelectNum) : 50
  };
}

function saveState(nextState) {
  const current = loadState();
  const state = {
    wordWeightMap: isObject(nextState.wordWeightMap) ? nextState.wordWeightMap : current.wordWeightMap,
    wordKnowRecord: isObject(nextState.wordKnowRecord) ? nextState.wordKnowRecord : current.wordKnowRecord,
    wordFirstUnselected: isObject(nextState.wordFirstUnselected) ? nextState.wordFirstUnselected : current.wordFirstUnselected,
    lastSelectNum: Number(nextState.lastSelectNum) > 0 ? Number(nextState.lastSelectNum) : current.lastSelectNum
  };
  writeJson(stateFile, state);
  return state;
}

function isObject(value) {
  return value && typeof value === 'object' && !Array.isArray(value);
}

function loadWorkbook() {
  const activeExcel = getActiveExcel();
  return buildWorkbook(activeExcel.filePath, activeExcel.displayName, activeExcel.mapping);
}

function getActiveExcel() {
  const selectedMeta = readJson(selectedExcelMetaFile, {});
  if (!defaultOnly && fs.existsSync(selectedExcelFile)) {
    return {
      filePath: selectedExcelFile,
      displayName: selectedMeta.originalName ? selectedMeta.originalName + '（自定义文件）' : selectedExcelFile,
      mapping: selectedMeta.mapping
    };
  }
  return {
    filePath: defaultExcelFile,
    displayName: defaultExcelFile
  };
}

function buildWorkbook(excelFile, displayName, mapping = null) {
  if (!excelFile) {
    throw new Error('请选择xlsx文件');
  }
  if (!fs.existsSync(excelFile)) {
    throw new Error('未找到Excel文件：' + excelFile);
  }
  const workbook = parseXlsx(excelFile);
  let totalRows;
  let summaryRows;
  if (isObject(mapping) && mapping.total && mapping.total.sheetName) {
    const mappedRows = parseMappedWorkbook(workbook, mapping);
    totalRows = mappedRows.totalRows;
    summaryRows = mappedRows.summaryRows;
  } else {
    const sheetEntries = getWorkbookSheetEntries(workbook);
    if (sheetEntries.length >= 2) {
    const selectedSheets = selectWorkbookSheets(sheetEntries);
    totalRows = parseTotalRows(selectedSheets.total.rows, selectedSheets.total.totalMapping);
    summaryRows = parseSummaryRows(selectedSheets.summary.rows, selectedSheets.summary.summaryMapping);
    } else {
    const fallbackRows = sheetEntries[0] ? sheetEntries[0].rows : getSingleSheetRows(workbook);
    totalRows = parseTotalRows(fallbackRows);
    summaryRows = [];
    }
  }

  const words = mergeWords(totalRows, summaryRows);
  if (!words.length) {
    throw new Error('两张工作表无任何有效单词数据');
  }

  return {
    filePath: displayName || excelFile,
    totalRows,
    summaryRows,
    words,
    totalUnits: uniqueNonEmpty(totalRows.map((row) => row.unit)),
    summaryUnits: uniqueNonEmpty(summaryRows.map((row) => row.unit || row.word.charAt(0).toUpperCase()))
  };
}

function parseMappedWorkbook(workbook, mapping) {
  const totalSheet = workbook.sheets[mapping.total.sheetName];
  if (!totalSheet) {
    throw new Error('映射错误：找不到总表工作表 ' + mapping.total.sheetName);
  }
  const summarySheet = mapping.summary && mapping.summary.sheetName ? workbook.sheets[mapping.summary.sheetName] : null;
  if (mapping.summary && mapping.summary.sheetName && !summarySheet) {
    throw new Error('映射错误：找不到汇总表工作表 ' + mapping.summary.sheetName);
  }
  return {
    totalRows: parseTotalRows(totalSheet, normalizeMapping(mapping.total, 'total')),
    summaryRows: summarySheet ? parseSummaryRows(summarySheet, normalizeMapping(mapping.summary, 'summary')) : []
  };
}

function normalizeMapping(rawMapping, role) {
  const normalized = {
    unitIndex: toColumnIndex(rawMapping.unitIndex),
    wordIndex: toColumnIndex(rawMapping.wordIndex),
    meaningIndex: toColumnIndex(rawMapping.meaningIndex)
  };
  if (role === 'summary') {
    normalized.phoneticIndex = toColumnIndex(rawMapping.phoneticIndex);
    normalized.partOfSpeechIndex = toColumnIndex(rawMapping.partOfSpeechIndex);
  }
  return normalized;
}

function toColumnIndex(value) {
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 ? number : -1;
}

function inspectWorkbook(excelFile, originalName) {
  const workbook = parseXlsx(excelFile);
  const sheets = Object.entries(workbook.sheets).map(([name, rows]) => {
    const columnCount = getEffectiveColumnCount(rows);
    return {
      name,
      columnCount,
      columns: createColumnOptions(rows, columnCount),
      totalMapping: inferSheetMapping(rows, 'total'),
      summaryMapping: inferSheetMapping(rows, 'summary')
    };
  }).filter((sheet) => sheet.columnCount >= 1);
  const selectable = getWorkbookSheetEntries(workbook);
  let suggested = null;
  if (selectable.length >= 2) {
    const selectedSheets = selectWorkbookSheets(selectable);
    suggested = {
      totalSheetName: selectedSheets.total.name,
      summarySheetName: selectedSheets.summary.name
    };
  } else if (selectable.length === 1) {
    suggested = {
      totalSheetName: selectable[0].name,
      summarySheetName: ''
    };
  }
  return {
    mappingRequired: true,
    originalName,
    sheets,
    suggested
  };
}

function createColumnOptions(rows, columnCount) {
  const headers = rows[0] || [];
  return Array.from({ length: columnCount }, (_, index) => {
    const header = normalize(headers[index]);
    const samples = rows.slice(1, 4).map((row) => normalize(row[index])).filter(Boolean);
    return {
      index,
      label: header || '第' + (index + 1) + '列',
      sample: samples.join(' / ')
    };
  });
}

function getSingleSheetRows(workbook) {
  const sheetNames = Object.keys(workbook.sheets);
  if (sheetNames.length !== 1) {
    throw new Error('文件格式错误：未找到支持的单词工作表');
  }
  return workbook.sheets[sheetNames[0]];
}

function getWorkbookSheetEntries(workbook) {
  return Object.entries(workbook.sheets).map(([name, rows]) => {
    const columnCount = getEffectiveColumnCount(rows);
    return {
      name,
      rows,
      columnCount,
      totalMapping: inferSheetMapping(rows, 'total'),
      summaryMapping: inferSheetMapping(rows, 'summary')
    };
  }).filter((entry) => entry.rows.length > 1 && entry.columnCount >= 3);
}

function selectWorkbookSheets(sheetEntries) {
  const sortedByColumns = sheetEntries.slice().sort((left, right) => {
    if (right.columnCount !== left.columnCount) {
      return right.columnCount - left.columnCount;
    }
    return getSummaryHeaderScore(right.rows) - getSummaryHeaderScore(left.rows);
  });
  const summary = sortedByColumns[0];
  const total = sortedByColumns.slice(1).sort((left, right) => {
    if (left.columnCount !== right.columnCount) {
      return left.columnCount - right.columnCount;
    }
    return getTotalHeaderScore(right.rows) - getTotalHeaderScore(left.rows);
  })[0];
  if (!total || !summary || total === summary) {
    throw new Error('文件格式错误：未找到可识别的两张单词工作表');
  }
  return { total, summary };
}

function parseXlsx(filePath) {
  const entries = readZipEntries(fs.readFileSync(filePath));
  const workbookXml = entries.get('xl/workbook.xml');
  const workbookRelsXml = entries.get('xl/_rels/workbook.xml.rels');
  if (!workbookXml || !workbookRelsXml) {
    throw new Error('Excel文件读取解析失败：缺少 workbook.xml');
  }

  const sharedStrings = parseSharedStrings(entries.get('xl/sharedStrings.xml'));
  const relationships = parseRelationships(workbookRelsXml.toString('utf8'));
  const sheets = {};
  const workbookText = workbookXml.toString('utf8');
  for (const sheetMatch of workbookText.matchAll(/<sheet\b([^>]*)\/?>(?:<\/sheet>)?/g)) {
    const attrs = parseAttrs(sheetMatch[1]);
    const relId = attrs['r:id'] || attrs.id;
    const target = relationships[relId];
    if (!attrs.name || !target) {
      continue;
    }
    const sheetPath = normalizeZipPath('xl/' + target);
    const sheetXml = entries.get(sheetPath);
    if (sheetXml) {
      sheets[decodeXml(attrs.name)] = parseSheet(sheetXml.toString('utf8'), sharedStrings);
    }
  }
  return { sheets };
}

function readZipEntries(buffer) {
  const entries = new Map();
  let offset = 0;
  while (offset < buffer.length - 30) {
    const signature = buffer.readUInt32LE(offset);
    if (signature !== 0x04034b50) {
      offset += 1;
      continue;
    }
    const method = buffer.readUInt16LE(offset + 8);
    const compressedSize = buffer.readUInt32LE(offset + 18);
    const fileNameLength = buffer.readUInt16LE(offset + 26);
    const extraLength = buffer.readUInt16LE(offset + 28);
    const fileNameStart = offset + 30;
    const fileName = normalizeZipPath(buffer.slice(fileNameStart, fileNameStart + fileNameLength).toString('utf8'));
    const dataStart = fileNameStart + fileNameLength + extraLength;
    const dataEnd = dataStart + compressedSize;
    const compressed = buffer.slice(dataStart, dataEnd);
    if (method === 0) {
      entries.set(fileName, compressed);
    } else if (method === 8) {
      entries.set(fileName, zlib.inflateRawSync(compressed));
    }
    offset = dataEnd;
  }
  return entries;
}

function normalizeZipPath(value) {
  const parts = [];
  value.replace(/\\/g, '/').split('/').forEach((part) => {
    if (!part || part === '.') {
      return;
    }
    if (part === '..') {
      parts.pop();
      return;
    }
    parts.push(part);
  });
  return parts.join('/');
}

function parseRelationships(xml) {
  const relationships = {};
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?>(?:<\/Relationship>)?/g)) {
    const attrs = parseAttrs(match[1]);
    if (attrs.Id && attrs.Target) {
      relationships[attrs.Id] = attrs.Target;
    }
  }
  return relationships;
}

function parseSharedStrings(buffer) {
  if (!buffer) {
    return [];
  }
  const xml = buffer.toString('utf8');
  const values = [];
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    values.push(extractTextNodes(match[1]));
  }
  return values;
}

function parseSheet(xml, sharedStrings) {
  const rows = [];
  for (const rowMatch of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
    const row = [];
    for (const cellMatch of rowMatch[1].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
      const attrs = parseAttrs(cellMatch[1] || cellMatch[3] || '');
      const columnIndex = getColumnIndex(attrs.r || '') - 1;
      if (columnIndex < 0) {
        continue;
      }
      row[columnIndex] = getCellValue(attrs, cellMatch[2] || '', sharedStrings);
    }
    rows.push(row);
  }
  return rows;
}

function getCellValue(attrs, content, sharedStrings) {
  if (attrs.t === 'inlineStr') {
    return extractTextNodes(content);
  }
  const valueMatch = content.match(/<v[^>]*>([\s\S]*?)<\/v>/);
  const rawValue = valueMatch ? decodeXml(valueMatch[1]).trim() : '';
  if (attrs.t === 's') {
    const index = Number(rawValue);
    return Number.isInteger(index) ? (sharedStrings[index] || '') : '';
  }
  return rawValue;
}

function getColumnIndex(cellRef) {
  const match = cellRef.match(/^[A-Z]+/i);
  if (!match) {
    return 0;
  }
  let index = 0;
  for (const char of match[0].toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index;
}

function parseAttrs(text) {
  const attrs = {};
  for (const match of text.matchAll(/([\w:-]+)="([^"]*)"/g)) {
    attrs[match[1]] = decodeXml(match[2]);
  }
  return attrs;
}

function extractTextNodes(text) {
  const parts = [];
  for (const match of text.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) {
    parts.push(decodeXml(match[1]));
  }
  return parts.join('').trim();
}

function decodeXml(value) {
  return String(value)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function parseTotalRows(rows, mapping = inferSheetMapping(rows, 'total')) {
  const dataStart = hasHeaderRow(rows) ? 1 : 0;
  return rows.slice(dataStart).map((row) => ({
    unit: normalize(cell(row, mapping.unitIndex)),
    word: normalizeWord(cell(row, mapping.wordIndex)),
    combinedMeaning: normalize(cell(row, mapping.meaningIndex))
  })).filter((row) => row.word);
}

function parseSummaryRows(rows, mapping = inferSheetMapping(rows, 'summary')) {
  const dataStart = hasHeaderRow(rows) ? 1 : 0;
  return rows.slice(dataStart).map((row) => ({
    unit: normalize(cell(row, mapping.unitIndex)),
    word: normalizeWord(cell(row, mapping.wordIndex)),
    phonetic: normalize(cell(row, mapping.phoneticIndex)),
    partOfSpeech: normalize(cell(row, mapping.partOfSpeechIndex)),
    meaning: normalize(cell(row, mapping.meaningIndex))
  })).filter((row) => row.word);
}

function inferSheetMapping(rows, role) {
  const headers = rows[0] || [];
  const columnCount = getEffectiveColumnCount(rows);
  const unitIndex = findColumnIndex(headers, ['day', 'unit', '单元', '章节', '课时'], fallbackIndex(columnCount, 0));
  const wordIndex = findColumnIndex(headers, ['单词', '英文', 'english', 'word', '词汇'], getWordFallbackIndex(columnCount));
  const meaningIndex = findColumnIndex(headers, ['中文意思', '中文释义', '汉语意思', '意思', '释义', 'meaning', 'chinese'], getMeaningFallbackIndex(columnCount, role));
  if (role === 'total') {
    return { unitIndex, wordIndex, meaningIndex };
  }
  return {
    unitIndex,
    wordIndex,
    phoneticIndex: findColumnIndex(headers, ['音标', 'phonetic', 'pronunciation'], fallbackIndex(columnCount, 3)),
    partOfSpeechIndex: findColumnIndex(headers, ['词性', 'partofspeech', 'part of speech', 'pos'], fallbackIndex(columnCount, 4)),
    meaningIndex
  };
}

function getEffectiveColumnCount(rows) {
  return rows.slice(0, 20).reduce((maxCount, row) => {
    let lastValueIndex = -1;
    row.forEach((value, index) => {
      if (normalize(value)) {
        lastValueIndex = index;
      }
    });
    return Math.max(maxCount, lastValueIndex + 1);
  }, 0);
}

function findColumnIndex(headers, names, fallback) {
  const normalizedNames = names.map(normalizeHeader);
  const index = headers.findIndex((header) => {
    const normalizedHeader = normalizeHeader(header);
    return normalizedHeader && normalizedNames.some((name) => normalizedHeader === name || normalizedHeader.includes(name));
  });
  return index >= 0 ? index : fallback;
}

function getWordFallbackIndex(columnCount) {
  if (columnCount >= 4) {
    return 2;
  }
  return fallbackIndex(columnCount, 1);
}

function getMeaningFallbackIndex(columnCount, role) {
  if (role === 'summary') {
    if (columnCount >= 6) {
      return 5;
    }
    if (columnCount >= 4) {
      return 3;
    }
  }
  if (columnCount >= 4) {
    return 3;
  }
  return fallbackIndex(columnCount, 2);
}

function fallbackIndex(columnCount, index) {
  return index >= 0 && index < columnCount ? index : -1;
}

function hasHeaderRow(rows) {
  return (rows[0] || []).some((value) => {
    const header = normalizeHeader(value);
    return ['day', 'unit', 'word', 'english', 'meaning', 'chinese', 'phonetic', 'pronunciation', 'pos', '单元', '单词', '英文', '意思', '释义', '音标', '词性'].some((name) => header.includes(normalizeHeader(name)));
  });
}

function getTotalHeaderScore(rows) {
  const headers = rows[0] || [];
  return ['day', '单元', '单词', 'word', '中文意思', '释义'].filter((name) => headers.some((header) => normalizeHeader(header).includes(normalizeHeader(name)))).length;
}

function getSummaryHeaderScore(rows) {
  const headers = rows[0] || [];
  return ['音标', 'phonetic', '词性', 'pos', 'page', '总表里有'].filter((name) => headers.some((header) => normalizeHeader(header).includes(normalizeHeader(name)))).length;
}

function cell(row, index) {
  return index >= 0 ? row[index] : '';
}

function normalizeHeader(value) {
  return normalize(value).toLowerCase().replace(/[\s_\-()（）/]+/g, '');
}

function mergeWords(totalRows, summaryRows) {
  const wordMap = new Map();
  totalRows.forEach((row) => {
    const existing = wordMap.get(row.word) || createWord(row.word);
    existing.unitsFromTotal.add(row.unit);
    existing.fromTotal = true;
    if (!existing.partOfSpeech && !existing.meaning && row.combinedMeaning) {
      const parts = splitCombinedMeaning(row.combinedMeaning);
      existing.phonetic = parts.phonetic;
      existing.partOfSpeech = parts.partOfSpeech;
      existing.meaning = parts.meaning;
    }
    wordMap.set(row.word, existing);
  });
  summaryRows.forEach((row) => {
    const existing = wordMap.get(row.word) || createWord(row.word);
    existing.fromSummary = true;
    existing.phonetic = row.phonetic;
    existing.partOfSpeech = row.partOfSpeech;
    existing.meaning = row.meaning;
    existing.unitsFromSummary.add(row.unit || row.word.charAt(0).toUpperCase());
    wordMap.set(row.word, existing);
  });
  return Array.from(wordMap.values()).map((wordItem) => ({
    word: wordItem.word,
    phonetic: wordItem.phonetic,
    partOfSpeech: wordItem.partOfSpeech,
    meaning: wordItem.meaning,
    totalUnits: Array.from(wordItem.unitsFromTotal).filter(Boolean),
    summaryUnits: Array.from(wordItem.unitsFromSummary).filter(Boolean),
    isDuplicate: wordItem.fromTotal && wordItem.fromSummary
  }));
}

function splitCombinedMeaning(value) {
  const text = normalize(value);
  const phoneticMatch = text.match(/(?:\/[^\/]+\/|\[[^\]]+\])/);
  const phonetic = phoneticMatch ? phoneticMatch[0] : '';
  const textWithoutPhonetic = phonetic ? text.replace(phonetic, '').trim() : text;
  const partMatches = textWithoutPhonetic.match(/\b(?:n|v|adj|adv|prep|pron|conj|int|num|art|aux|vi|vt|abbr|det|interj)\./gi) || [];
  const partOfSpeech = Array.from(new Set(partMatches.map((item) => item.toLowerCase()))).join(' ');
  const meaning = partMatches.length ? textWithoutPhonetic.replace(/\b(?:n|v|adj|adv|prep|pron|conj|int|num|art|aux|vi|vt|abbr|det|interj)\./gi, '').replace(/\s+/g, ' ').trim() : textWithoutPhonetic;
  return {
    phonetic,
    partOfSpeech,
    meaning: meaning || text
  };
}

function createWord(word) {
  return {
    word,
    phonetic: '',
    partOfSpeech: '',
    meaning: '',
    unitsFromTotal: new Set(),
    unitsFromSummary: new Set(),
    fromTotal: false,
    fromSummary: false
  };
}

function uniqueNonEmpty(values) {
  return Array.from(new Set(values.map(normalize).filter(Boolean)));
}

function normalize(value) {
  if (value === undefined || value === null) {
    return '';
  }
  return String(value).trim();
}

function normalizeWord(value) {
  return normalize(value).toLowerCase();
}

function ensureFirstUnselected(words) {
  const state = loadState();
  let changed = false;
  words.forEach((wordItem) => {
    if (state.wordFirstUnselected[wordItem.word] === undefined) {
      state.wordFirstUnselected[wordItem.word] = true;
      changed = true;
    }
  });
  if (changed) {
    saveState(state);
  }
  return loadState();
}

function sendJson(response, statusCode, value) {
  response.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store'
  });
  response.end(JSON.stringify(value));
}

function sendFile(response, filePath, contentType) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      sendText(response, 404, 'Not found');
      return;
    }
    response.writeHead(200, { 'Content-Type': contentType });
    response.end(content);
  });
}

function sendText(response, statusCode, text) {
  response.writeHead(statusCode, { 'Content-Type': 'text/plain; charset=utf-8' });
  response.end(text);
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    request.on('data', (chunk) => chunks.push(chunk));
    request.on('end', () => {
      try {
        const raw = Buffer.concat(chunks).toString('utf8');
        resolve(raw ? JSON.parse(raw) : {});
      } catch (error) {
        reject(error);
      }
    });
    request.on('error', reject);
  });
}

function readRawBody(request, limitBytes = 30 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let totalLength = 0;
    request.on('data', (chunk) => {
      totalLength += chunk.length;
      if (totalLength > limitBytes) {
        reject(new Error('Excel文件过大，请选择小于30MB的文件'));
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });
    request.on('end', () => resolve(Buffer.concat(chunks)));
    request.on('error', reject);
  });
}

function getUploadFileName(request) {
  const rawName = request.headers['x-file-name'] || '自定义词库.xlsx';
  try {
    return decodeURIComponent(String(rawName));
  } catch (error) {
    return String(rawName);
  }
}

function renderIndex() {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>背单词工具</title>
  <link rel="stylesheet" href="/style.css">
</head>
<body>
  <main id="app" class="app-shell"></main>
  <script src="/local-app.js"></script>
</body>
</html>`;
}

const server = http.createServer(async (request, response) => {
  const parsedUrl = new URL(request.url, `http://${request.headers.host || host}`);
  try {
    if (request.method === 'GET' && parsedUrl.pathname === '/') {
      response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      response.end(renderIndex());
      return;
    }
    if (request.method === 'GET' && parsedUrl.pathname === '/style.css') {
      sendFile(response, path.join(rootDir, 'style.css'), 'text/css; charset=utf-8');
      return;
    }
    if (request.method === 'GET' && parsedUrl.pathname === '/local-app.js') {
      sendFile(response, path.join(rootDir, 'local-app.js'), 'application/javascript; charset=utf-8');
      return;
    }
    if (request.method === 'GET' && parsedUrl.pathname === '/api/workbook') {
      const workbook = loadWorkbook();
      ensureFirstUnselected(workbook.words);
      sendJson(response, 200, workbook);
      return;
    }
    if (request.method === 'POST' && parsedUrl.pathname === '/api/workbook-file') {
      ensureCacheDir();
      const originalName = getUploadFileName(request);
      const uploadBuffer = await readRawBody(request);
      fs.writeFileSync(pendingExcelFile, uploadBuffer);
      writeJson(pendingExcelMetaFile, {
        originalName,
        uploadedAt: new Date().toISOString()
      });
      sendJson(response, 200, inspectWorkbook(pendingExcelFile, originalName));
      return;
    }
    if (request.method === 'POST' && parsedUrl.pathname === '/api/workbook-mapping') {
      if (!fs.existsSync(pendingExcelFile)) {
        sendJson(response, 400, { error: '请先选择xlsx文件' });
        return;
      }
      const body = await readBody(request);
      const pendingMeta = readJson(pendingExcelMetaFile, {});
      const originalName = pendingMeta.originalName || '自定义词库.xlsx';
      const mapping = body.mapping || body;
      const workbook = buildWorkbook(pendingExcelFile, originalName + '（自定义文件）', mapping);
      fs.copyFileSync(pendingExcelFile, selectedExcelFile);
      writeJson(selectedExcelMetaFile, {
        originalName,
        mapping,
        selectedAt: new Date().toISOString()
      });
      if (fs.existsSync(pendingExcelFile)) {
        fs.unlinkSync(pendingExcelFile);
      }
      if (fs.existsSync(pendingExcelMetaFile)) {
        fs.unlinkSync(pendingExcelMetaFile);
      }
      ensureFirstUnselected(workbook.words);
      sendJson(response, 200, workbook);
      return;
    }
    if (request.method === 'POST' && parsedUrl.pathname === '/api/workbook-file-auto') {
      ensureCacheDir();
      const originalName = getUploadFileName(request);
      const uploadBuffer = await readRawBody(request);
      fs.writeFileSync(pendingExcelFile, uploadBuffer);
      try {
        const workbook = buildWorkbook(pendingExcelFile, originalName + '（自定义文件）');
        fs.copyFileSync(pendingExcelFile, selectedExcelFile);
        writeJson(selectedExcelMetaFile, {
          originalName,
          selectedAt: new Date().toISOString()
        });
        ensureFirstUnselected(workbook.words);
        sendJson(response, 200, workbook);
      } finally {
        if (fs.existsSync(pendingExcelFile)) {
          fs.unlinkSync(pendingExcelFile);
        }
      }
      return;
    }
    if (request.method === 'GET' && parsedUrl.pathname === '/api/state') {
      try {
        const workbook = loadWorkbook();
        sendJson(response, 200, ensureFirstUnselected(workbook.words));
      } catch (error) {
        sendJson(response, 200, loadState());
      }
      return;
    }
    if (request.method === 'POST' && parsedUrl.pathname === '/api/state') {
      const body = await readBody(request);
      sendJson(response, 200, saveState(body));
      return;
    }
    sendText(response, 404, 'Not found');
  } catch (error) {
    sendJson(response, 500, { error: error.message });
  }
});

function listen(port) {
  server.once('error', (error) => {
    if (error.code === 'EADDRINUSE' && port < startPort + 20) {
      listen(port + 1);
      return;
    }
    throw error;
  });
  server.listen(port, host, () => {
    const address = `http://${host}:${port}/`;
    console.log('背单词工具已启动：' + address);
    console.log(defaultExcelFile ? '指定Excel文件：' + defaultExcelFile : '未指定默认Excel文件，请在页面选择xlsx文件');
    if (args.open) {
      const opener = process.platform === 'win32' ? 'start' : process.platform === 'darwin' ? 'open' : 'xdg-open';
      childProcess.exec(`${opener} "" "${address}"`);
    }
  });
}

ensureCacheDir();
listen(startPort);
const xlsx = require('node-xlsx');

const TOTAL_SHEET_NAME = '单词总表';
const SUMMARY_SHEET_NAME = '八年级上下册英语全单词汇总';
const TOTAL_HEADERS = ['单元', '单词', '词性和中文意思'];
const SUMMARY_HEADERS = ['单词', '音标', '词性', '中文意思'];

function parseWorkbook(filePath) {
  let sheets;
  try {
    sheets = xlsx.parse(filePath);
  } catch (error) {
    throw new Error('Excel文件读取解析失败：' + error.message);
  }

  const totalSheet = findSheet(sheets, TOTAL_SHEET_NAME);
  const summarySheet = findSheet(sheets, SUMMARY_SHEET_NAME);
  if (!totalSheet || !summarySheet) {
    throw new Error('文件格式错误：缺少「单词总表」或「八年级上下册英语全单词汇总」工作表');
  }

  validateHeaders(totalSheet.data, TOTAL_HEADERS, TOTAL_SHEET_NAME);
  validateHeaders(summarySheet.data, SUMMARY_HEADERS, SUMMARY_SHEET_NAME);

  const totalRows = parseTotalRows(totalSheet.data);
  const summaryRows = parseSummaryRows(summarySheet.data);
  const words = mergeWords(totalRows, summaryRows);

  if (!words.length) {
    throw new Error('两张工作表无任何有效单词数据');
  }

  return {
    totalRows,
    summaryRows,
    words,
    totalUnits: uniqueNonEmpty(totalRows.map((row) => row.unit)),
    summaryUnits: uniqueNonEmpty(summaryRows.map((row) => row.unit || row.word.charAt(0).toUpperCase()))
  };
}

function findSheet(sheets, sheetName) {
  return sheets.find((sheet) => sheet.name === sheetName);
}

function validateHeaders(data, expectedHeaders, sheetName) {
  const headers = data[0] || [];
  expectedHeaders.forEach((header, index) => {
    if (normalize(headers[index]) !== header) {
      throw new Error('文件格式错误：「' + sheetName + '」表头必须依次为' + expectedHeaders.join('、'));
    }
  });
}

function parseTotalRows(data) {
  return data.slice(1).map((row) => ({
    unit: normalize(row[0]),
    word: normalizeWord(row[1]),
    combinedMeaning: normalize(row[2])
  })).filter((row) => row.word);
}

function parseSummaryRows(data) {
  return data.slice(1).map((row) => ({
    word: normalizeWord(row[0]),
    phonetic: normalize(row[1]),
    partOfSpeech: normalize(row[2]),
    meaning: normalize(row[3]),
    unit: normalize(row[4])
  })).filter((row) => row.word);
}

function mergeWords(totalRows, summaryRows) {
  const wordMap = new Map();

  totalRows.forEach((row) => {
    const existing = wordMap.get(row.word) || createWord(row.word);
    existing.unitsFromTotal.add(row.unit);
    existing.fromTotal = true;
    if (!existing.partOfSpeech && !existing.meaning && row.combinedMeaning) {
      existing.meaning = row.combinedMeaning;
    }
    wordMap.set(row.word, existing);
  });

  summaryRows.forEach((row) => {
    const existing = wordMap.get(row.word) || createWord(row.word);
    existing.fromSummary = true;
    existing.phonetic = row.phonetic;
    existing.partOfSpeech = row.partOfSpeech;
    existing.meaning = row.meaning;
    if (row.unit) {
      existing.unitsFromSummary.add(row.unit);
    } else {
      existing.unitsFromSummary.add(row.word.charAt(0).toUpperCase());
    }
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

function getWordsByUnits(parsedWorkbook, totalUnits, summaryUnits) {
  const totalSet = new Set(totalUnits);
  const summarySet = new Set(summaryUnits);
  return parsedWorkbook.words.filter((wordItem) => {
    const matchedTotal = wordItem.totalUnits.some((unit) => totalSet.has(unit));
    const matchedSummary = wordItem.summaryUnits.some((unit) => summarySet.has(unit));
    return matchedTotal || matchedSummary;
  });
}

function uniqueNonEmpty(values) {
  return Array.from(new Set(values.filter(Boolean)));
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

module.exports = {
  parseWorkbook,
  getWordsByUnits
};
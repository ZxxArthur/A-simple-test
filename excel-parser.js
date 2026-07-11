const xlsx = require('node-xlsx');

function parseWorkbook(filePath) {
  let sheets;
  try {
    sheets = xlsx.parse(filePath);
  } catch (error) {
    throw new Error('Excel文件读取解析失败：' + error.message);
  }

  const sheetEntries = getWorkbookSheetEntries(sheets);
  let totalRows;
  let summaryRows;
  if (sheetEntries.length >= 2) {
    const selectedSheets = selectWorkbookSheets(sheetEntries);
    totalRows = parseTotalRows(selectedSheets.total.data, selectedSheets.total.totalMapping);
    summaryRows = parseSummaryRows(selectedSheets.summary.data, selectedSheets.summary.summaryMapping);
  } else if (sheetEntries.length === 1) {
    totalRows = parseTotalRows(sheetEntries[0].data);
    summaryRows = [];
  } else {
    throw new Error('文件格式错误：未找到可识别的单词工作表');
  }
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

function getWorkbookSheetEntries(sheets) {
  return sheets.map((sheet) => {
    const columnCount = getEffectiveColumnCount(sheet.data || []);
    return {
      name: sheet.name,
      data: sheet.data || [],
      columnCount,
      totalMapping: inferSheetMapping(sheet.data || [], 'total'),
      summaryMapping: inferSheetMapping(sheet.data || [], 'summary')
    };
  }).filter((entry) => entry.data.length > 1 && entry.columnCount >= 3);
}

function selectWorkbookSheets(sheetEntries) {
  const sortedByColumns = sheetEntries.slice().sort((left, right) => {
    if (right.columnCount !== left.columnCount) {
      return right.columnCount - left.columnCount;
    }
    return getSummaryHeaderScore(right.data) - getSummaryHeaderScore(left.data);
  });
  const summary = sortedByColumns[0];
  const total = sortedByColumns.slice(1).sort((left, right) => {
    if (left.columnCount !== right.columnCount) {
      return left.columnCount - right.columnCount;
    }
    return getTotalHeaderScore(right.data) - getTotalHeaderScore(left.data);
  })[0];
  if (!total || !summary || total === summary) {
    throw new Error('文件格式错误：未找到可识别的两张单词工作表');
  }
  return { total, summary };
}

function parseTotalRows(data, mapping = inferSheetMapping(data, 'total')) {
  const dataStart = hasHeaderRow(data) ? 1 : 0;
  return data.slice(dataStart).map((row) => ({
    unit: normalize(cell(row, mapping.unitIndex)),
    word: normalizeWord(cell(row, mapping.wordIndex)),
    combinedMeaning: normalize(cell(row, mapping.meaningIndex))
  })).filter((row) => row.word);
}

function parseSummaryRows(data, mapping = inferSheetMapping(data, 'summary')) {
  const dataStart = hasHeaderRow(data) ? 1 : 0;
  return data.slice(dataStart).map((row) => ({
    unit: normalize(cell(row, mapping.unitIndex)),
    word: normalizeWord(cell(row, mapping.wordIndex)),
    phonetic: normalize(cell(row, mapping.phoneticIndex)),
    partOfSpeech: normalize(cell(row, mapping.partOfSpeechIndex)),
    meaning: normalize(cell(row, mapping.meaningIndex))
  })).filter((row) => row.word);
}

function inferSheetMapping(data, role) {
  const headers = data[0] || [];
  const columnCount = getEffectiveColumnCount(data);
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

function getEffectiveColumnCount(data) {
  return data.slice(0, 20).reduce((maxCount, row) => {
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

function hasHeaderRow(data) {
  return (data[0] || []).some((value) => {
    const header = normalizeHeader(value);
    return ['day', 'unit', 'word', 'english', 'meaning', 'chinese', 'phonetic', 'pronunciation', 'pos', '单元', '单词', '英文', '意思', '释义', '音标', '词性'].some((name) => header.includes(normalizeHeader(name)));
  });
}

function getTotalHeaderScore(data) {
  const headers = data[0] || [];
  return ['day', '单元', '单词', 'word', '中文意思', '释义'].filter((name) => headers.some((header) => normalizeHeader(header).includes(normalizeHeader(name)))).length;
}

function getSummaryHeaderScore(data) {
  const headers = data[0] || [];
  return ['音标', 'phonetic', '词性', 'pos', 'page', '总表里有'].filter((name) => headers.some((header) => normalizeHeader(header).includes(normalizeHeader(name)))).length;
}

function cell(row, index) {
  return index >= 0 ? row[index] : '';
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

function normalizeHeader(value) {
  return normalize(value).toLowerCase().replace(/[\s_\-()（）/]+/g, '');
}

module.exports = {
  parseWorkbook,
  getWordsByUnits
};
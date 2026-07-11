const { LocalStorage } = require('node-localstorage');
const path = require('path');

const storage = new LocalStorage(path.join(__dirname, '.word-memory-cache'));

const defaults = {
  excelFilePath: '',
  wordWeightMap: {},
  wordKnowRecord: {},
  wordFirstUnselected: {},
  lastSelectNum: 50
};

function readValue(key) {
  try {
    const rawValue = storage.getItem(key);
    if (rawValue === null || rawValue === undefined) {
      return defaults[key];
    }
    return JSON.parse(rawValue);
  } catch (error) {
    notifyStorageError('缓存读取失败：' + error.message);
    return defaults[key];
  }
}

function writeValue(key, value) {
  try {
    storage.setItem(key, JSON.stringify(value));
    return true;
  } catch (error) {
    notifyStorageError('缓存写入失败：' + error.message);
    return false;
  }
}

function notifyStorageError(message) {
  if (typeof window !== 'undefined' && typeof window.alert === 'function') {
    window.alert(message);
  }
}

function getExcelFilePath() {
  return readValue('excelFilePath');
}

function setExcelFilePath(filePath) {
  return writeValue('excelFilePath', filePath);
}

function getWordWeightMap() {
  return readValue('wordWeightMap');
}

function setWordWeight(word, weight) {
  const weightMap = getWordWeightMap();
  weightMap[word] = weight;
  return writeValue('wordWeightMap', weightMap);
}

function getWordKnowRecord() {
  return readValue('wordKnowRecord');
}

function setWordKnowRecord(word, isKnown) {
  const knowRecord = getWordKnowRecord();
  knowRecord[word] = isKnown;
  return writeValue('wordKnowRecord', knowRecord);
}

function getWordFirstUnselected() {
  return readValue('wordFirstUnselected');
}

function markWordsSelected(words) {
  const firstUnselected = getWordFirstUnselected();
  words.forEach((wordItem) => {
    firstUnselected[wordItem.word] = false;
  });
  return writeValue('wordFirstUnselected', firstUnselected);
}

function ensureFirstUnselected(words) {
  const firstUnselected = getWordFirstUnselected();
  let changed = false;
  words.forEach((wordItem) => {
    if (firstUnselected[wordItem.word] === undefined) {
      firstUnselected[wordItem.word] = true;
      changed = true;
    }
  });
  if (changed) {
    writeValue('wordFirstUnselected', firstUnselected);
  }
  return firstUnselected;
}

function getLastSelectNum() {
  const value = Number(readValue('lastSelectNum'));
  return Number.isFinite(value) && value > 0 ? value : defaults.lastSelectNum;
}

function setLastSelectNum(value) {
  return writeValue('lastSelectNum', Number(value));
}

module.exports = {
  getExcelFilePath,
  setExcelFilePath,
  getWordWeightMap,
  setWordWeight,
  getWordKnowRecord,
  setWordKnowRecord,
  getWordFirstUnselected,
  markWordsSelected,
  ensureFirstUnselected,
  getLastSelectNum,
  setLastSelectNum
};
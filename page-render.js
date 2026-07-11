const { ipcRenderer } = require('electron');
const { parseWorkbook, getWordsByUnits } = require('./excel-parser');
const { pickWeightedWords, KNOWN_WEIGHT, UNKNOWN_DUPLICATE_WEIGHT } = require('./weight-calc');
const storage = require('./storage');

const appRoot = document.getElementById('app');
const selectOptions = [20, 30, 40, 50, 60, 70, 80, 90, 100];

const state = {
  parsedWorkbook: null,
  currentWords: [],
  currentMode: '',
  currentIndex: 0,
  revealedUnknown: false,
  roundKnown: {}
};

document.addEventListener('DOMContentLoaded', renderHomePage);

function renderHomePage() {
  state.currentWords = [];
  state.currentMode = '';
  state.currentIndex = 0;
  state.revealedUnknown = false;
  state.roundKnown = {};

  const filePath = storage.getExcelFilePath();
  const lastSelectNum = storage.getLastSelectNum();
  appRoot.innerHTML = `
    <h1 class="page-title">背单词工具</h1>
    <section class="section">
      <h2 class="section-title">Excel文件选择</h2>
      <div class="file-row">
        <button id="chooseFileButton" class="primary-button" type="button">选择xlsx文件</button>
        <span id="filePathText" class="path-text">${escapeHtml(filePath || '请选择单词Excel文件')}</span>
      </div>
    </section>
    <section class="section">
      <h2 class="section-title">功能操作</h2>
      <div class="action-row">
        <select id="selectNum" class="select-input">${renderSelectOptions(lastSelectNum)}</select>
        <button id="randomButton" class="primary-button" type="button">随机生成背诵</button>
        <button id="unitButton" class="secondary-button" type="button">按单元选择生成单词</button>
      </div>
    </section>
  `;

  document.getElementById('chooseFileButton').addEventListener('click', chooseFile);
  document.getElementById('randomButton').addEventListener('click', startRandomFlow);
  document.getElementById('unitButton').addEventListener('click', startUnitFlow);
}

async function chooseFile() {
  const filePath = await ipcRenderer.invoke('select-xlsx-file');
  if (!filePath) {
    return;
  }
  try {
    state.parsedWorkbook = parseWorkbook(filePath);
    storage.setExcelFilePath(filePath);
    document.getElementById('filePathText').textContent = filePath;
  } catch (error) {
    alert(error.message);
  }
}

function getParsedWorkbook() {
  const filePath = storage.getExcelFilePath();
  if (!filePath) {
    alert('请选择单词Excel文件');
    return null;
  }

  try {
    state.parsedWorkbook = parseWorkbook(filePath);
    storage.ensureFirstUnselected(state.parsedWorkbook.words);
    return state.parsedWorkbook;
  } catch (error) {
    alert(error.message);
    return null;
  }
}

function startRandomFlow() {
  const parsedWorkbook = getParsedWorkbook();
  if (!parsedWorkbook) {
    return;
  }

  const selectElement = document.getElementById('selectNum');
  let selectNum = Number(selectElement.value);
  storage.setLastSelectNum(selectNum);

  if (selectNum > parsedWorkbook.words.length) {
    alert('当前词库总单词数量为' + parsedWorkbook.words.length + '，已自动修改抽取数量为全部单词');
    selectNum = parsedWorkbook.words.length;
  }

  const selectedWords = pickWeightedWords(
    parsedWorkbook.words,
    selectNum,
    storage.getWordKnowRecord(),
    storage.getWordFirstUnselected(),
    storage.getWordWeightMap()
  );
  storage.markWordsSelected(selectedWords);
  state.currentWords = selectedWords;
  renderModePage();
}

function startUnitFlow() {
  const parsedWorkbook = getParsedWorkbook();
  if (!parsedWorkbook) {
    return;
  }
  renderUnitPage(parsedWorkbook);
}

function renderUnitPage(parsedWorkbook) {
  appRoot.innerHTML = `
    <h1 class="page-title">按单元选择生成单词</h1>
    <div class="unit-grid">
      <section class="section">
        <h2 class="section-title">单词总表单元</h2>
        <div class="checkbox-list">${renderCheckboxes('totalUnit', parsedWorkbook.totalUnits)}</div>
      </section>
      <section class="section">
        <h2 class="section-title">八年级汇总分组</h2>
        <div class="checkbox-list">${renderCheckboxes('summaryUnit', parsedWorkbook.summaryUnits)}</div>
      </section>
    </div>
    <div class="bottom-actions">
      <button id="confirmUnitsButton" class="primary-button" type="button">确认筛选</button>
      <button id="backHomeButton" class="secondary-button" type="button">返回主页</button>
    </div>
  `;

  document.getElementById('confirmUnitsButton').addEventListener('click', () => confirmUnits(parsedWorkbook));
  document.getElementById('backHomeButton').addEventListener('click', renderHomePage);
}

function confirmUnits(parsedWorkbook) {
  const totalUnits = getCheckedValues('totalUnit');
  const summaryUnits = getCheckedValues('summaryUnit');
  if (!totalUnits.length && !summaryUnits.length) {
    alert('至少选择一个单元');
    return;
  }

  const words = getWordsByUnits(parsedWorkbook, totalUnits, summaryUnits);
  if (!words.length) {
    alert('无匹配单词');
    return;
  }

  storage.markWordsSelected(words);
  state.currentWords = words;
  renderModePage();
}

function renderModePage() {
  appRoot.innerHTML = `
    <h1 class="page-title">选择背诵模式</h1>
    <section class="section">
      <div class="mode-row">
        <label class="choice-label"><input type="radio" name="mode" value="englishToChinese">显示英文回忆中文意思</label>
        <label class="choice-label"><input type="radio" name="mode" value="chineseToEnglish">显示词性+中文意思回忆英文</label>
      </div>
    </section>
    <button id="startMemoryButton" class="primary-button" type="button" disabled>开始背诵</button>
  `;

  const startButton = document.getElementById('startMemoryButton');
  document.querySelectorAll('input[name="mode"]').forEach((radio) => {
    radio.addEventListener('change', () => {
      state.currentMode = radio.value;
      startButton.disabled = false;
    });
  });
  startButton.addEventListener('click', startMemory);
}

function startMemory() {
  state.currentIndex = 0;
  state.revealedUnknown = false;
  state.roundKnown = {};
  renderMemoryPage();
}

function renderMemoryPage() {
  if (isRoundComplete()) {
    renderCompletePage();
    return;
  }

  const currentWord = state.currentWords[state.currentIndex];
  appRoot.innerHTML = `
    <div class="progress-text">${state.currentIndex + 1}/${state.currentWords.length}</div>
    <section class="word-stage">
      ${renderWord(currentWord, state.revealedUnknown ? 'full' : state.currentMode)}
    </section>
    <div class="bottom-actions">
      <button id="previousButton" class="secondary-button" type="button" ${state.currentIndex === 0 ? 'disabled' : ''}>上一个</button>
      ${state.revealedUnknown ? '<button id="confirmUnknownButton" class="primary-button" type="button">确定</button>' : '<button id="knowButton" class="primary-button" type="button">知道</button><button id="unknownButton" class="danger-button" type="button">不知道</button>'}
      <button id="nextButton" class="secondary-button" type="button" ${state.currentIndex === state.currentWords.length - 1 ? 'disabled' : ''}>下一个</button>
    </div>
  `;

  document.getElementById('previousButton').addEventListener('click', goPreviousWord);
  document.getElementById('nextButton').addEventListener('click', goNextWord);
  if (state.revealedUnknown) {
    document.getElementById('confirmUnknownButton').addEventListener('click', confirmUnknown);
  } else {
    document.getElementById('knowButton').addEventListener('click', markKnown);
    document.getElementById('unknownButton').addEventListener('click', markUnknown);
  }
}

function markKnown() {
  const currentWord = state.currentWords[state.currentIndex];
  storage.setWordKnowRecord(currentWord.word, true);
  storage.setWordWeight(currentWord.word, KNOWN_WEIGHT);
  state.roundKnown[currentWord.word] = true;
  state.revealedUnknown = false;

  if (isRoundComplete()) {
    renderCompletePage();
    return;
  }

  state.currentIndex = getNextIndex();
  renderMemoryPage();
}

function markUnknown() {
  const currentWord = state.currentWords[state.currentIndex];
  storage.setWordKnowRecord(currentWord.word, false);
  storage.setWordWeight(currentWord.word, UNKNOWN_DUPLICATE_WEIGHT);
  state.revealedUnknown = true;
  renderMemoryPage();
}

function confirmUnknown() {
  state.revealedUnknown = false;
  state.currentIndex = getNextIndex();
  renderMemoryPage();
}

function goPreviousWord() {
  if (state.currentIndex <= 0) {
    return;
  }
  state.revealedUnknown = false;
  state.currentIndex -= 1;
  renderMemoryPage();
}

function goNextWord() {
  if (state.currentIndex >= state.currentWords.length - 1) {
    return;
  }
  state.revealedUnknown = false;
  state.currentIndex += 1;
  renderMemoryPage();
}

function isRoundComplete() {
  return state.currentWords.length > 0 && state.currentWords.every((wordItem) => state.roundKnown[wordItem.word] === true);
}

function moveToNextUnknownIfNeeded() {
  if (state.roundKnown[state.currentWords[state.currentIndex].word] !== true) {
    return;
  }
  state.currentIndex = getNextIndex();
}

function getNextIndex() {
  for (let offset = 1; offset <= state.currentWords.length; offset += 1) {
    const nextIndex = (state.currentIndex + offset) % state.currentWords.length;
    if (state.roundKnown[state.currentWords[nextIndex].word] !== true) {
      return nextIndex;
    }
  }
  return state.currentIndex;
}

function renderCompletePage() {
  appRoot.innerHTML = `
    <h1 class="page-title">本轮背诵完成</h1>
    <div class="bottom-actions">
      <button id="answerButton" class="primary-button" type="button">正确答案</button>
      <button id="againButton" class="primary-button" type="button">再来一次</button>
      <button id="homeButton" class="secondary-button" type="button">返回主页</button>
      <button id="exitButton" class="danger-button" type="button">退出</button>
    </div>
  `;

  document.getElementById('answerButton').addEventListener('click', renderAnswerPage);
  document.getElementById('againButton').addEventListener('click', startMemory);
  document.getElementById('homeButton').addEventListener('click', renderHomePage);
  document.getElementById('exitButton').addEventListener('click', () => ipcRenderer.invoke('exit-app'));
}

function renderAnswerPage() {
  appRoot.innerHTML = `
    <h1 class="page-title">正确答案</h1>
    <section class="section">
      <h2 class="section-title">本次测试 ${state.currentWords.length} 个单词</h2>
      <div class="answer-table-wrap">
        <table class="answer-table">
          <thead>
            <tr>
              <th>序号</th>
              <th>英文</th>
              <th>中文</th>
              <th>词性</th>
            </tr>
          </thead>
          <tbody>${renderAnswerRows(state.currentWords)}</tbody>
        </table>
      </div>
    </section>
    <div class="bottom-actions">
      <button id="completeButton" class="secondary-button" type="button">返回完成页</button>
      <button id="homeButton" class="secondary-button" type="button">返回主页</button>
    </div>
  `;

  document.getElementById('completeButton').addEventListener('click', renderCompletePage);
  document.getElementById('homeButton').addEventListener('click', renderHomePage);
}

function renderAnswerRows(words) {
  return words.map((wordItem, index) => `
    <tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(wordItem.word || '-')}</td>
      <td>${escapeHtml(wordItem.meaning || '-')}</td>
      <td>${escapeHtml(wordItem.partOfSpeech || '-')}</td>
    </tr>
  `).join('');
}

function renderWord(wordItem, mode) {
  const parts = [];
  if (mode === 'englishToChinese' || mode === 'full') {
    parts.push(`<div class="word-text">${escapeHtml(wordItem.word)}</div>`);
    if (wordItem.phonetic) {
      parts.push(`<div class="phonetic-text">${escapeHtml(wordItem.phonetic)}</div>`);
    }
  }
  if (mode === 'chineseToEnglish' || mode === 'full') {
    if (wordItem.partOfSpeech) {
      parts.push(`<div class="part-text">${escapeHtml(wordItem.partOfSpeech)}</div>`);
    }
    if (wordItem.meaning) {
      parts.push(`<div class="meaning-text">${escapeHtml(wordItem.meaning)}</div>`);
    }
  }
  return `<div class="word-content">${parts.join('')}</div>`;
}

function renderSelectOptions(selectedValue) {
  return selectOptions.map((value) => `<option value="${value}" ${value === selectedValue ? 'selected' : ''}>${value}</option>`).join('');
}

function renderCheckboxes(name, values) {
  if (!values.length) {
    return '<div class="empty-state">无可选单元</div>';
  }
  return values.map((value) => `
    <label class="checkbox-label">
      <input type="checkbox" name="${name}" value="${escapeHtml(value)}">
      <span>${escapeHtml(value)}</span>
    </label>
  `).join('');
}

function getCheckedValues(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map((input) => input.value);
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"]/g, (char) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;'
  }[char]));
}
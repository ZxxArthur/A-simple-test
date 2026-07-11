const BASE_WEIGHT = 10;
const FIRST_UNSELECTED_WEIGHT = 15;
const UNKNOWN_DUPLICATE_WEIGHT = 15;
const KNOWN_WEIGHT = 5;
const selectOptions = [20, 30, 40, 50, 60, 70, 80, 90, 100];

const appRoot = document.getElementById('app');
const state = {
  workbook: null,
  workbookError: '',
  mappingInspect: null,
  mappingDraft: null,
  persisted: null,
  currentWords: [],
  currentMode: '',
  currentIndex: 0,
  revealedUnknown: false,
  roundKnown: {}
};

document.addEventListener('DOMContentLoaded', initApp);

async function initApp() {
  state.persisted = await loadPersistedState();
  try {
    state.workbook = await apiGet('/api/workbook');
    state.workbookError = '';
  } catch (error) {
    state.workbook = null;
    state.workbookError = error.message;
  }
  renderHomePage();
}

async function loadPersistedState() {
  try {
    return await apiGet('/api/state');
  } catch (error) {
    return createDefaultPersistedState();
  }
}

function createDefaultPersistedState() {
  return {
    wordWeightMap: {},
    wordKnowRecord: {},
    wordFirstUnselected: {},
    lastSelectNum: 50
  };
}

async function apiGet(path) {
  const response = await fetch(path, { cache: 'no-store' });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || '请求失败');
  }
  return payload;
}

async function savePersisted() {
  const response = await fetch('/api/state', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(state.persisted)
  });
  const payload = await response.json();
  if (response.ok) {
    state.persisted = payload;
  }
}

async function apiPostJson(path, body) {
  const response = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || '请求失败');
  }
  return payload;
}

function renderHomePage() {
  state.currentWords = [];
  state.currentMode = '';
  state.currentIndex = 0;
  state.revealedUnknown = false;
  state.roundKnown = {};

  const lastSelectNum = Number((state.persisted || {}).lastSelectNum) || 50;
  const hasWorkbook = Boolean(state.workbook);
  const fileText = hasWorkbook ? state.workbook.filePath : (state.workbookError || '请选择xlsx文件');
  appRoot.innerHTML = `
    <h1 class="page-title">背单词工具</h1>
    <section class="section">
      <h2 class="section-title">Excel文件</h2>
      <div class="file-row">
        <input id="excelFileInput" type="file" accept=".xlsx" hidden>
        <button id="chooseExcelButton" class="primary-button" type="button">选择xlsx文件</button>
        <button id="reloadButton" class="secondary-button" type="button">刷新页面</button>
        <span class="path-text">${escapeHtml(fileText)}</span>
      </div>
      ${hasWorkbook ? '' : '<div class="empty-state">未加载词库，请选择符合要求的 xlsx 文件。</div>'}
    </section>
    <section class="section">
      <h2 class="section-title">功能操作</h2>
      <div class="action-row">
        <select id="selectNum" class="select-input">${renderSelectOptions(lastSelectNum)}</select>
        <button id="randomButton" class="primary-button" type="button" ${hasWorkbook ? '' : 'disabled'}>随机生成背诵</button>
        <button id="unitButton" class="secondary-button" type="button" ${hasWorkbook ? '' : 'disabled'}>按单元选择生成单词</button>
      </div>
    </section>
  `;

  document.getElementById('chooseExcelButton').addEventListener('click', () => document.getElementById('excelFileInput').click());
  document.getElementById('excelFileInput').addEventListener('change', uploadExcelFile);
  document.getElementById('reloadButton').addEventListener('click', () => window.location.reload());
  document.getElementById('randomButton').addEventListener('click', startRandomFlow);
  document.getElementById('unitButton').addEventListener('click', startUnitFlow);
}

async function uploadExcelFile(event) {
  const file = event.target.files[0];
  if (!file) {
    return;
  }
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    alert('请选择xlsx文件');
    event.target.value = '';
    return;
  }

  const button = document.getElementById('chooseExcelButton');
  button.disabled = true;
  button.textContent = '正在读取...';
  try {
    const response = await fetch('/api/workbook-file', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'X-File-Name': encodeURIComponent(file.name)
      },
      body: file
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Excel文件读取失败');
    }
    state.mappingInspect = payload;
    state.mappingDraft = createDefaultMappingDraft(payload);
    renderMappingPage();
  } catch (error) {
    alert(error.message);
  } finally {
    button.disabled = false;
    button.textContent = '选择xlsx文件';
    event.target.value = '';
  }
}

function renderMappingPage() {
  const inspect = state.mappingInspect;
  if (!inspect || !inspect.sheets || !inspect.sheets.length) {
    alert('未找到可映射的工作表');
    renderHomePage();
    return;
  }
  const draft = state.mappingDraft || createDefaultMappingDraft(inspect);
  state.mappingDraft = draft;
  const totalSheet = getInspectSheet(draft.total.sheetName) || inspect.sheets[0];
  const summarySheet = getInspectSheet(draft.summary.sheetName) || inspect.sheets.find((sheet) => sheet.name !== totalSheet.name) || null;
  draft.total.sheetName = totalSheet.name;
  draft.summary.sheetName = summarySheet ? summarySheet.name : '';

  appRoot.innerHTML = `
    <h1 class="page-title">Excel字段映射</h1>
    <section class="section">
      <h2 class="section-title">选择工作表</h2>
      <div class="mapping-grid">
        <label class="mapping-field">
          <span>总表逻辑 sheet</span>
          <select id="totalSheetSelect" class="select-input">${renderSheetOptions(inspect.sheets, draft.total.sheetName)}</select>
        </label>
        <label class="mapping-field">
          <span>汇总表逻辑 sheet</span>
          <select id="summarySheetSelect" class="select-input">${renderSheetOptions(inspect.sheets, draft.summary.sheetName, true)}</select>
        </label>
      </div>
    </section>
    <section class="section">
      <h2 class="section-title">总表字段</h2>
      <div class="mapping-grid">
        ${renderColumnField('总表单元', 'totalUnitIndex', totalSheet, draft.total.unitIndex)}
        ${renderColumnField('英文单词', 'totalWordIndex', totalSheet, draft.total.wordIndex)}
        ${renderColumnField('组合释义/中文意思', 'totalMeaningIndex', totalSheet, draft.total.meaningIndex)}
      </div>
    </section>
    <section class="section">
      <h2 class="section-title">汇总表字段</h2>
      ${summarySheet ? `
        <div class="mapping-grid">
          ${renderColumnField('汇总单元', 'summaryUnitIndex', summarySheet, draft.summary.unitIndex)}
          ${renderColumnField('英文单词', 'summaryWordIndex', summarySheet, draft.summary.wordIndex)}
          ${renderColumnField('音标', 'summaryPhoneticIndex', summarySheet, draft.summary.phoneticIndex, true)}
          ${renderColumnField('词性', 'summaryPartOfSpeechIndex', summarySheet, draft.summary.partOfSpeechIndex, true)}
          ${renderColumnField('中文意思', 'summaryMeaningIndex', summarySheet, draft.summary.meaningIndex)}
        </div>
      ` : '<div class="empty-state">不使用汇总表，只按总表逻辑导入。</div>'}
    </section>
    <section class="section">
      <div class="bottom-actions">
        <button id="applyMappingButton" class="primary-button" type="button">按这个映射导入</button>
        <button id="cancelMappingButton" class="secondary-button" type="button">返回首页</button>
      </div>
    </section>
  `;

  document.getElementById('totalSheetSelect').addEventListener('change', handleMappingSheetChange);
  document.getElementById('summarySheetSelect').addEventListener('change', handleMappingSheetChange);
  document.getElementById('applyMappingButton').addEventListener('click', submitWorkbookMapping);
  document.getElementById('cancelMappingButton').addEventListener('click', renderHomePage);
}

function createDefaultMappingDraft(inspect) {
  const totalSheet = inspect.sheets.find((sheet) => sheet.name === (inspect.suggested || {}).totalSheetName) || inspect.sheets[0];
  const summarySheet = inspect.sheets.find((sheet) => sheet.name === (inspect.suggested || {}).summarySheetName) || inspect.sheets.find((sheet) => sheet.name !== totalSheet.name) || null;
  return {
    total: createRoleMapping(totalSheet, 'total'),
    summary: summarySheet ? createRoleMapping(summarySheet, 'summary') : { sheetName: '' }
  };
}

function createRoleMapping(sheet, role) {
  const mapping = role === 'summary' ? sheet.summaryMapping : sheet.totalMapping;
  if (role === 'summary') {
    return {
      sheetName: sheet.name,
      unitIndex: mapping.unitIndex,
      wordIndex: mapping.wordIndex,
      phoneticIndex: mapping.phoneticIndex,
      partOfSpeechIndex: mapping.partOfSpeechIndex,
      meaningIndex: mapping.meaningIndex
    };
  }
  return {
    sheetName: sheet.name,
    unitIndex: mapping.unitIndex,
    wordIndex: mapping.wordIndex,
    meaningIndex: mapping.meaningIndex
  };
}

function handleMappingSheetChange() {
  const totalSheet = getInspectSheet(document.getElementById('totalSheetSelect').value);
  const summarySheet = getInspectSheet(document.getElementById('summarySheetSelect').value);
  state.mappingDraft = {
    total: createRoleMapping(totalSheet, 'total'),
    summary: summarySheet ? createRoleMapping(summarySheet, 'summary') : { sheetName: '' }
  };
  renderMappingPage();
}

async function submitWorkbookMapping() {
  const mapping = readMappingForm();
  if (!mapping.total.sheetName || mapping.total.wordIndex < 0 || mapping.total.meaningIndex < 0) {
    alert('总表逻辑至少需要选择 sheet、英文单词和中文意思字段');
    return;
  }
  if (mapping.summary.sheetName && (mapping.summary.wordIndex < 0 || mapping.summary.meaningIndex < 0)) {
    alert('汇总表逻辑至少需要选择英文单词和中文意思字段');
    return;
  }
  const button = document.getElementById('applyMappingButton');
  button.disabled = true;
  button.textContent = '正在导入...';
  try {
    const workbook = await apiPostJson('/api/workbook-mapping', { mapping });
    state.workbook = workbook;
    state.workbookError = '';
    state.mappingInspect = null;
    state.mappingDraft = null;
    state.persisted = await apiGet('/api/state');
    renderHomePage();
  } catch (error) {
    alert(error.message);
    button.disabled = false;
    button.textContent = '按这个映射导入';
  }
}

function readMappingForm() {
  return {
    total: {
      sheetName: document.getElementById('totalSheetSelect').value,
      unitIndex: readColumnIndex('totalUnitIndex'),
      wordIndex: readColumnIndex('totalWordIndex'),
      meaningIndex: readColumnIndex('totalMeaningIndex')
    },
    summary: {
      sheetName: document.getElementById('summarySheetSelect').value,
      unitIndex: readColumnIndex('summaryUnitIndex'),
      wordIndex: readColumnIndex('summaryWordIndex'),
      phoneticIndex: readColumnIndex('summaryPhoneticIndex'),
      partOfSpeechIndex: readColumnIndex('summaryPartOfSpeechIndex'),
      meaningIndex: readColumnIndex('summaryMeaningIndex')
    }
  };
}

function readColumnIndex(id) {
  const element = document.getElementById(id);
  return element ? Number(element.value) : -1;
}

function getInspectSheet(sheetName) {
  return (state.mappingInspect || {}).sheets.find((sheet) => sheet.name === sheetName);
}

function renderSheetOptions(sheets, selectedName, includeEmpty = false) {
  const options = includeEmpty ? ['<option value="">不使用汇总表</option>'] : [];
  return options.concat(sheets.map((sheet) => `<option value="${escapeHtml(sheet.name)}" ${sheet.name === selectedName ? 'selected' : ''}>${escapeHtml(sheet.name)}（${sheet.columnCount}列）</option>`)).join('');
}

function renderColumnField(label, id, sheet, selectedIndex, includeEmpty = false) {
  return `
    <label class="mapping-field">
      <span>${escapeHtml(label)}</span>
      <select id="${id}" class="select-input">${renderColumnOptions(sheet, selectedIndex, includeEmpty)}</select>
    </label>
  `;
}

function renderColumnOptions(sheet, selectedIndex, includeEmpty) {
  const options = includeEmpty ? ['<option value="-1">不使用</option>'] : [];
  return options.concat(sheet.columns.map((column) => {
    const sample = column.sample ? '：' + column.sample : '';
    const text = '第' + (column.index + 1) + '列 ' + column.label + sample;
    return `<option value="${column.index}" ${Number(selectedIndex) === column.index ? 'selected' : ''}>${escapeHtml(text)}</option>`;
  })).join('');
}

async function startRandomFlow() {
  if (!state.workbook) {
    alert('请先选择xlsx文件');
    return;
  }
  let selectNum = Number(document.getElementById('selectNum').value);
  state.persisted.lastSelectNum = selectNum;

  if (selectNum > state.workbook.words.length) {
    alert('当前词库总单词数量为' + state.workbook.words.length + '，已自动修改抽取数量为全部单词');
    selectNum = state.workbook.words.length;
  }

  const selectedWords = pickWeightedWords(
    state.workbook.words,
    selectNum,
    state.persisted.wordKnowRecord,
    state.persisted.wordFirstUnselected,
    state.persisted.wordWeightMap
  );
  markWordsSelected(selectedWords);
  await savePersisted();
  state.currentWords = selectedWords;
  renderModePage();
}

function startUnitFlow() {
  if (!state.workbook) {
    alert('请先选择xlsx文件');
    return;
  }
  renderUnitPage();
}

function renderUnitPage() {
  appRoot.innerHTML = `
    <h1 class="page-title">按单元选择生成单词</h1>
    <div class="unit-grid">
      <section class="section">
        <h2 class="section-title">单词总表单元</h2>
        <div class="checkbox-list">${renderCheckboxes('totalUnit', state.workbook.totalUnits)}</div>
      </section>
      <section class="section">
        <h2 class="section-title">八年级汇总分组</h2>
        <div class="checkbox-list">${renderCheckboxes('summaryUnit', state.workbook.summaryUnits)}</div>
      </section>
    </div>
    <div class="bottom-actions">
      <button id="confirmUnitsButton" class="primary-button" type="button">确认筛选</button>
      <button id="backHomeButton" class="secondary-button" type="button">返回主页</button>
    </div>
  `;

  document.getElementById('confirmUnitsButton').addEventListener('click', confirmUnits);
  document.getElementById('backHomeButton').addEventListener('click', renderHomePage);
}

async function confirmUnits() {
  const totalUnits = getCheckedValues('totalUnit');
  const summaryUnits = getCheckedValues('summaryUnit');
  if (!totalUnits.length && !summaryUnits.length) {
    alert('至少选择一个单元');
    return;
  }

  const words = getWordsByUnits(totalUnits, summaryUnits);
  if (!words.length) {
    alert('无匹配单词');
    return;
  }

  markWordsSelected(words);
  await savePersisted();
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

async function markKnown() {
  const currentWord = state.currentWords[state.currentIndex];
  state.persisted.wordKnowRecord[currentWord.word] = true;
  state.persisted.wordWeightMap[currentWord.word] = KNOWN_WEIGHT;
  state.roundKnown[currentWord.word] = true;
  state.revealedUnknown = false;
  await savePersisted();

  if (isRoundComplete()) {
    renderCompletePage();
    return;
  }

  state.currentIndex = getNextIndex();
  renderMemoryPage();
}

async function markUnknown() {
  const currentWord = state.currentWords[state.currentIndex];
  state.persisted.wordKnowRecord[currentWord.word] = false;
  state.persisted.wordWeightMap[currentWord.word] = UNKNOWN_DUPLICATE_WEIGHT;
  state.revealedUnknown = true;
  await savePersisted();
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

function renderCompletePage() {
  appRoot.innerHTML = `
    <h1 class="page-title">本轮背诵完成</h1>
    <div class="bottom-actions">
      <button id="answerButton" class="primary-button" type="button">正确答案</button>
      <button id="againButton" class="primary-button" type="button">再来一次</button>
      <button id="homeButton" class="secondary-button" type="button">返回主页</button>
    </div>
  `;

  document.getElementById('answerButton').addEventListener('click', renderAnswerPage);
  document.getElementById('againButton').addEventListener('click', startMemory);
  document.getElementById('homeButton').addEventListener('click', renderHomePage);
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

function getWordsByUnits(totalUnits, summaryUnits) {
  const totalSet = new Set(totalUnits);
  const summarySet = new Set(summaryUnits);
  return state.workbook.words.filter((wordItem) => {
    const matchedTotal = wordItem.totalUnits.some((unit) => totalSet.has(unit));
    const matchedSummary = wordItem.summaryUnits.some((unit) => summarySet.has(unit));
    return matchedTotal || matchedSummary;
  });
}

function getWordWeight(wordItem, knowRecord, firstUnselected, weightMap) {
  if (firstUnselected[wordItem.word] === true) {
    return FIRST_UNSELECTED_WEIGHT;
  }
  if (knowRecord[wordItem.word] === true) {
    return KNOWN_WEIGHT;
  }
  if (knowRecord[wordItem.word] === false && wordItem.isDuplicate) {
    return UNKNOWN_DUPLICATE_WEIGHT;
  }
  const cachedWeight = Number(weightMap[wordItem.word]);
  if (Number.isFinite(cachedWeight) && cachedWeight > 0) {
    return cachedWeight;
  }
  return BASE_WEIGHT;
}

function pickWeightedWords(words, count, knowRecord, firstUnselected, weightMap) {
  const pool = words.map((wordItem) => ({
    wordItem,
    weight: getWordWeight(wordItem, knowRecord, firstUnselected, weightMap)
  }));
  const selected = [];
  const targetCount = Math.min(count, pool.length);

  while (selected.length < targetCount && pool.length) {
    const totalWeight = pool.reduce((sum, item) => sum + item.weight, 0);
    let randomPoint = Math.random() * totalWeight;
    let selectedIndex = 0;

    for (let index = 0; index < pool.length; index += 1) {
      randomPoint -= pool[index].weight;
      if (randomPoint <= 0) {
        selectedIndex = index;
        break;
      }
    }

    selected.push(pool[selectedIndex].wordItem);
    pool.splice(selectedIndex, 1);
  }

  return selected;
}

function markWordsSelected(words) {
  words.forEach((wordItem) => {
    state.persisted.wordFirstUnselected[wordItem.word] = false;
  });
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
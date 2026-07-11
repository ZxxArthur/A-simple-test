const BASE_WEIGHT = 10;
const FIRST_UNSELECTED_WEIGHT = 15;
const UNKNOWN_DUPLICATE_WEIGHT = 15;
const KNOWN_WEIGHT = 5;

function getWordWeight(wordItem, knowRecord, firstUnselected, weightMap = {}) {
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

function pickWeightedWords(words, count, knowRecord, firstUnselected, weightMap = {}) {
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

module.exports = {
  getWordWeight,
  pickWeightedWords,
  KNOWN_WEIGHT,
  UNKNOWN_DUPLICATE_WEIGHT
};
import { defaultWords, defaultDatasetMeta } from "./data/defaultWords.js";

const STORAGE_KEY = "sew-word-studio.store.v1";
const HISTORY_LIMIT = 400;
const AUTO_ADVANCE_DELAY = 720;
const PERSISTABLE_BACKGROUND_LIMIT = 650_000;

const MODE_LABELS = {
  multiple: "4-Choice",
  input: "Input",
  flashcard: "Flashcards",
};

const DIRECTION_LABELS = {
  "en-to-ja": "English → Japanese",
  "ja-to-en": "Japanese → English",
  mixed: "Mixed",
};

const DEFAULT_SETTINGS = {
  theme: "aurora",
  activeDataset: defaultDatasetMeta.id,
  mode: "multiple",
  direction: "en-to-ja",
  focus: "all",
  sessionLength: 20,
  rangeStart: "",
  rangeEnd: "",
  autoPronounce: false,
  autoAdvance: true,
  customBackground: "",
};

const DEFAULT_PROGRESS = {
  attempts: [],
  wordStats: {},
  favorites: {},
  bestStreak: 0,
};

const DEFAULT_DATASET = {
  meta: defaultDatasetMeta,
  words: prepareWords(defaultDatasetMeta.id, defaultWords),
};

const dateFormatter = new Intl.DateTimeFormat("ja-JP", {
  dateStyle: "short",
  timeStyle: "short",
});

const el = {
  answerArea: document.querySelector("#answerArea"),
  autoAdvanceToggle: document.querySelector("#autoAdvanceToggle"),
  autoPronounceToggle: document.querySelector("#autoPronounceToggle"),
  backgroundFileInput: document.querySelector("#backgroundFileInput"),
  backgroundHint: document.querySelector("#backgroundHint"),
  backToSetupButton: document.querySelector("#backToSetupButton"),
  clearBackgroundButton: document.querySelector("#clearBackgroundButton"),
  confusionList: document.querySelector("#confusionList"),
  csvFileInput: document.querySelector("#csvFileInput"),
  dashboardWeakWords: document.querySelector("#dashboardWeakWords"),
  datasetCount: document.querySelector("#datasetCount"),
  datasetSelect: document.querySelector("#datasetSelect"),
  datasetStatusText: document.querySelector("#datasetStatusText"),
  directionBreakdown: document.querySelector("#directionBreakdown"),
  directionChip: document.querySelector("#directionChip"),
  directionSelect: document.querySelector("#directionSelect"),
  exportHistoryButton: document.querySelector("#exportHistoryButton"),
  favoriteButton: document.querySelector("#favoriteButton"),
  feedbackBox: document.querySelector("#feedbackBox"),
  focusSelect: document.querySelector("#focusSelect"),
  historyFilterSelect: document.querySelector("#historyFilterSelect"),
  historyList: document.querySelector("#historyList"),
  metricAccuracy: document.querySelector("#metricAccuracy"),
  metricAttempts: document.querySelector("#metricAttempts"),
  metricBestStreak: document.querySelector("#metricBestStreak"),
  metricWeakCount: document.querySelector("#metricWeakCount"),
  modeBreakdown: document.querySelector("#modeBreakdown"),
  modeChip: document.querySelector("#modeChip"),
  modeSelect: document.querySelector("#modeSelect"),
  nextQuestionButton: document.querySelector("#nextQuestionButton"),
  overallAccuracy: document.querySelector("#overallAccuracy"),
  promptText: document.querySelector("#promptText"),
  questionBadge: document.querySelector("#questionBadge"),
  quizAccuracy: document.querySelector("#quizAccuracy"),
  quizMetaList: document.querySelector("#quizMetaList"),
  quizProgress: document.querySelector("#quizProgress"),
  quizStatusBadge: document.querySelector("#quizStatusBadge"),
  quizStreak: document.querySelector("#quizStreak"),
  rangeEndInput: document.querySelector("#rangeEndInput"),
  rangeHint: document.querySelector("#rangeHint"),
  rangeStartInput: document.querySelector("#rangeStartInput"),
  revealButton: document.querySelector("#revealButton"),
  restartSessionButton: document.querySelector("#restartSessionButton"),
  resetAllButton: document.querySelector("#resetAllButton"),
  resetProgressButton: document.querySelector("#resetProgressButton"),
  returnToSetupButton: document.querySelector("#returnToSetupButton"),
  scoreAccuracyValue: document.querySelector("#scoreAccuracyValue"),
  scoreAnsweredCount: document.querySelector("#scoreAnsweredCount"),
  scoreBestStreak: document.querySelector("#scoreBestStreak"),
  scoreCorrectCount: document.querySelector("#scoreCorrectCount"),
  scoreDatasetChip: document.querySelector("#scoreDatasetChip"),
  scoreMistakeBadge: document.querySelector("#scoreMistakeBadge"),
  scoreMistakeList: document.querySelector("#scoreMistakeList"),
  scoreNoteList: document.querySelector("#scoreNoteList"),
  scoreRangeChip: document.querySelector("#scoreRangeChip"),
  scoreSummaryText: document.querySelector("#scoreSummaryText"),
  sessionAccuracy: document.querySelector("#sessionAccuracy"),
  sessionLengthSelect: document.querySelector("#sessionLengthSelect"),
  sessionModeBadge: document.querySelector("#sessionModeBadge"),
  sessionProgress: document.querySelector("#sessionProgress"),
  sessionStreak: document.querySelector("#sessionStreak"),
  setupDatasetChip: document.querySelector("#setupDatasetChip"),
  setupRangeChip: document.querySelector("#setupRangeChip"),
  setupSummaryText: document.querySelector("#setupSummaryText"),
  speakButton: document.querySelector("#speakButton"),
  startSessionButton: document.querySelector("#startSessionButton"),
  studyScreens: document.querySelectorAll(".study-screen"),
  tabButtons: document.querySelectorAll(".tab-button"),
  themeSelect: document.querySelector("#themeSelect"),
  totalAnswered: document.querySelector("#totalAnswered"),
  wordIdChip: document.querySelector("#wordIdChip"),
  jumpToSettingsButton: document.querySelector("#jumpToSettingsButton"),
};

const state = {
  store: loadStore(),
  currentQuestion: null,
  currentView: "study",
  pendingAutoAdvance: null,
  studyScreen: "setup",
  lastSessionSummary: null,
  session: createSessionState(),
  temporaryBackground: "",
};

syncStore();
bindEvents();
hydrateControls();
applyTheme();
renderAll();
registerServiceWorker();

function prepareWords(datasetId, words) {
  return words
    .map((word, index) => {
      const id = String(word.id ?? index + 1).trim() || String(index + 1);
      const english = String(word.english ?? "").trim();
      const japanese = String(word.japanese ?? "").trim();
      if (!english || !japanese) {
        return null;
      }
      return {
        id,
        uid: `${datasetId}:${id}`,
        english,
        japanese,
        order: index + 1,
        sectionNumber: getSectionNumber(id, index),
      };
    })
    .filter(Boolean);
}

function createSessionState() {
  return {
    answered: 0,
    correct: 0,
    streak: 0,
    bestStreak: 0,
    limit: 0,
    attempts: [],
    recentIds: [],
    active: false,
    completed: false,
  };
}

function createWordStat() {
  return {
    asked: 0,
    correct: 0,
    incorrect: 0,
    byDirection: {},
    byMode: {},
    confusions: {},
    lastSeenAt: null,
  };
}

function getSectionNumber(id, index) {
  const direct = Number(id);
  if (Number.isFinite(direct)) {
    return direct;
  }
  const matched = String(id).match(/\d+/);
  if (matched) {
    return Number(matched[0]);
  }
  return index + 1;
}

function loadStore() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return {
        settings: { ...DEFAULT_SETTINGS },
        progress: structuredClone(DEFAULT_PROGRESS),
        customDataset: null,
      };
    }

    const parsed = JSON.parse(raw);
    return {
      settings: {
        ...DEFAULT_SETTINGS,
        ...(parsed.settings || {}),
      },
      progress: {
        ...structuredClone(DEFAULT_PROGRESS),
        ...(parsed.progress || {}),
        attempts: Array.isArray(parsed.progress?.attempts) ? parsed.progress.attempts : [],
        wordStats: parsed.progress?.wordStats || {},
        favorites: parsed.progress?.favorites || {},
        bestStreak: Number(parsed.progress?.bestStreak || 0),
      },
      customDataset: parsed.customDataset || null,
    };
  } catch (error) {
    console.warn("Failed to load local data:", error);
    return {
      settings: { ...DEFAULT_SETTINGS },
      progress: structuredClone(DEFAULT_PROGRESS),
      customDataset: null,
    };
  }
}

function saveStore() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.store));
}

function syncStore() {
  const datasets = getDatasets();
  const availableIds = new Set(datasets.map((dataset) => dataset.meta.id));
  if (!availableIds.has(state.store.settings.activeDataset)) {
    state.store.settings.activeDataset = defaultDatasetMeta.id;
  }
  state.store.settings.sessionLength = Number(state.store.settings.sessionLength || 20);
}

function getDatasets() {
  const datasets = [DEFAULT_DATASET];
  if (state.store.customDataset?.words?.length) {
    datasets.push({
      meta: state.store.customDataset.meta,
      words: prepareWords(state.store.customDataset.meta.id, state.store.customDataset.words),
    });
  }
  return datasets;
}

function getDatasetById(datasetId) {
  return getDatasets().find((dataset) => dataset.meta.id === datasetId) || DEFAULT_DATASET;
}

function getActiveDataset() {
  return getDatasetById(state.store.settings.activeDataset);
}

function getWordLookup() {
  const map = new Map();
  for (const dataset of getDatasets()) {
    for (const word of dataset.words) {
      map.set(word.uid, word);
    }
  }
  return map;
}

function hydrateControls() {
  el.modeSelect.value = state.store.settings.mode;
  el.directionSelect.value = state.store.settings.direction;
  el.focusSelect.value = state.store.settings.focus;
  el.sessionLengthSelect.value = String(state.store.settings.sessionLength);
  el.rangeStartInput.value = state.store.settings.rangeStart;
  el.rangeEndInput.value = state.store.settings.rangeEnd;
  el.autoPronounceToggle.checked = Boolean(state.store.settings.autoPronounce);
  el.autoAdvanceToggle.checked = Boolean(state.store.settings.autoAdvance);
  el.themeSelect.value = state.store.settings.theme;
  populateDatasetSelect();
}

function populateDatasetSelect() {
  const datasets = getDatasets();
  el.datasetSelect.replaceChildren();
  for (const dataset of datasets) {
    const option = document.createElement("option");
    option.value = dataset.meta.id;
    option.textContent = `${dataset.meta.name} (${dataset.words.length}語)`;
    if (dataset.meta.id === state.store.settings.activeDataset) {
      option.selected = true;
    }
    el.datasetSelect.append(option);
  }
}

function bindEvents() {
  el.tabButtons.forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.viewTarget || "study"));
  });

  el.modeSelect.addEventListener("change", () => updateSetting("mode", el.modeSelect.value));
  el.directionSelect.addEventListener("change", () => updateSetting("direction", el.directionSelect.value));
  el.focusSelect.addEventListener("change", () => updateSetting("focus", el.focusSelect.value));
  el.sessionLengthSelect.addEventListener("change", () => updateSetting("sessionLength", Number(el.sessionLengthSelect.value)));
  el.rangeStartInput.addEventListener("input", () => updateRangeSetting("rangeStart", el.rangeStartInput.value));
  el.rangeEndInput.addEventListener("input", () => updateRangeSetting("rangeEnd", el.rangeEndInput.value));
  el.autoPronounceToggle.addEventListener("change", () => updateSetting("autoPronounce", el.autoPronounceToggle.checked));
  el.autoAdvanceToggle.addEventListener("change", () => updateSetting("autoAdvance", el.autoAdvanceToggle.checked));
  el.themeSelect.addEventListener("change", () => {
    updateSetting("theme", el.themeSelect.value);
    applyTheme();
    renderAll();
  });
  el.datasetSelect.addEventListener("change", () => {
    updateSetting("activeDataset", el.datasetSelect.value);
    state.currentQuestion = null;
    state.session = createSessionState();
    state.studyScreen = "setup";
    state.lastSessionSummary = null;
    setFeedback("データセットを切り替えました。新しいセッションで出題されます。");
    renderAll();
  });

  el.startSessionButton.addEventListener("click", () => startSession());
  el.nextQuestionButton.addEventListener("click", () => {
    if (state.session.completed || !state.session.active) {
      startSession();
      return;
    }
    nextQuestion();
  });
  el.jumpToSettingsButton.addEventListener("click", () => switchView("settings"));
  el.backToSetupButton.addEventListener("click", () => moveToSetupScreen());
  el.restartSessionButton.addEventListener("click", () => startSession());
  el.returnToSetupButton.addEventListener("click", () => moveToSetupScreen());
  el.revealButton.addEventListener("click", revealCurrentAnswer);
  el.speakButton.addEventListener("click", () => {
    if (state.currentQuestion) {
      speakText(state.currentQuestion.word.english);
    }
  });
  el.favoriteButton.addEventListener("click", toggleFavoriteForCurrentWord);

  el.historyFilterSelect.addEventListener("change", renderHistory);
  el.exportHistoryButton.addEventListener("click", exportStudyData);
  document.addEventListener("keydown", handleKeyboardShortcuts);
}

function updateSetting(key, value) {
  state.store.settings[key] = value;
  saveStore();
  renderAll();
}

function updateRangeSetting(key, value) {
  const sanitized = value.replace(/[^\d]/g, "");
  state.store.settings[key] = sanitized;
  if (key === "rangeStart") {
    el.rangeStartInput.value = sanitized;
  }
  if (key === "rangeEnd") {
    el.rangeEndInput.value = sanitized;
  }
  saveStore();
  renderAll();
}

function moveToSetupScreen() {
  clearPendingAutoAdvance();
  state.studyScreen = "setup";
  state.session = createSessionState();
  state.currentQuestion = null;
  renderAll();
}

function getFocusLabel(focus) {
  const labels = {
    all: "All",
    weak: "Weak words",
    favorites: "Favorites",
    "recent-mistakes": "Recent mistakes",
  };
  return labels[focus] || focus;
}

function switchView(viewName) {
  state.currentView = viewName;
  document.querySelectorAll(".view").forEach((view) => {
    view.classList.toggle("is-active", view.id === `view-${viewName}`);
  });
  el.tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.viewTarget === viewName);
  });
}

function startSession() {
  const rangeWords = getWordsInRange(getActiveDataset().words);
  if (!rangeWords.length) {
    state.studyScreen = "setup";
    state.currentQuestion = null;
    setFeedback("指定した範囲に単語がありません。開始番号と終了番号を見直してください。", "wrong");
    renderAll();
    return;
  }

  clearPendingAutoAdvance();
  state.session = {
    answered: 0,
    correct: 0,
    streak: 0,
    bestStreak: 0,
    limit: Number(state.store.settings.sessionLength || 0),
    attempts: [],
    recentIds: [],
    active: true,
    completed: false,
  };
  state.lastSessionSummary = null;
  state.studyScreen = "quiz";
  switchView("study");
  nextQuestion(true);
}

function nextQuestion(isFreshStart = false) {
  clearPendingAutoAdvance();
  if (!state.session.active) {
    startSession();
    return;
  }
  if (!isFreshStart && state.session.completed) {
    return;
  }

  const question = buildQuestion();
  state.currentQuestion = question;

  if (!question) {
    setFeedback("出題できる単語が見つかりませんでした。条件を変えるか、CSVを読み込んでみてください。");
    renderAll();
    return;
  }

  if (state.store.settings.autoPronounce) {
    speakText(question.word.english);
  }

  setFeedback("回答すると、その結果がこの端末に保存されます。");
  renderAll();
}

function buildQuestion() {
  const dataset = getActiveDataset();
  const words = getWordsInRange(dataset.words);
  if (!words.length) {
    return null;
  }

  const direction = state.store.settings.direction === "mixed"
    ? (Math.random() < 0.5 ? "en-to-ja" : "ja-to-en")
    : state.store.settings.direction;
  const candidates = getCandidatesForFocus(words);
  const pool = candidates.length ? candidates : words;
  const word = selectWeightedWord(pool, direction);
  if (!word) {
    return null;
  }

  state.session.recentIds.unshift(word.uid);
  state.session.recentIds = state.session.recentIds.slice(0, 8);

  const mode = state.store.settings.mode;
  return {
    word,
    datasetName: dataset.meta.name,
    prompt: direction === "en-to-ja" ? word.english : word.japanese,
    direction,
    mode,
    answered: false,
    correct: false,
    revealed: false,
    selectedChoiceUid: "",
    options: mode === "multiple" ? buildMultipleChoiceOptions(word, pool, direction) : [],
  };
}

function getCandidatesForFocus(words) {
  const focus = state.store.settings.focus;
  if (focus === "favorites") {
    return words.filter((word) => state.store.progress.favorites[word.uid]);
  }
  if (focus === "recent-mistakes") {
    const recentMistakes = new Set(
      state.store.progress.attempts
        .filter((attempt) => !attempt.correct)
        .slice(0, 40)
        .map((attempt) => attempt.uid)
    );
    return words.filter((word) => recentMistakes.has(word.uid));
  }
  if (focus === "weak") {
    const weakIds = new Set(getWeakWords(120).map((entry) => entry.word.uid));
    return words.filter((word) => weakIds.has(word.uid));
  }
  return words;
}

function getWordsInRange(words) {
  const { start, end } = getSelectedRange();
  if (start === null && end === null) {
    return words;
  }

  return words.filter((word) => {
    if (start !== null && word.sectionNumber < start) {
      return false;
    }
    if (end !== null && word.sectionNumber > end) {
      return false;
    }
    return true;
  });
}

function getSelectedRange() {
  const rawStart = state.store.settings.rangeStart;
  const rawEnd = state.store.settings.rangeEnd;
  const parsedStart = rawStart ? Number(rawStart) : null;
  const parsedEnd = rawEnd ? Number(rawEnd) : null;

  if (parsedStart !== null && parsedEnd !== null && parsedStart > parsedEnd) {
    return { start: parsedEnd, end: parsedStart };
  }
  return { start: parsedStart, end: parsedEnd };
}

function formatRangeLabel(dataset = getActiveDataset()) {
  const { start, end } = getSelectedRange();
  const min = dataset.words[0]?.sectionNumber ?? 1;
  const max = dataset.words[dataset.words.length - 1]?.sectionNumber ?? dataset.words.length;
  if (start === null && end === null) {
    return `${min} - ${max}`;
  }
  if (start !== null && end !== null) {
    return `${start} - ${end}`;
  }
  if (start !== null) {
    return `${start} - ${max}`;
  }
  return `${min} - ${end}`;
}

function selectWeightedWord(words, direction) {
  const recentSet = new Set(state.session.recentIds);
  const filtered = words.filter((word) => !recentSet.has(word.uid));
  const source = filtered.length >= 4 ? filtered : words;
  if (!source.length) {
    return null;
  }

  const totalWeight = source.reduce((sum, word) => sum + getWordWeight(word.uid, direction), 0);
  let threshold = Math.random() * totalWeight;
  for (const word of source) {
    threshold -= getWordWeight(word.uid, direction);
    if (threshold <= 0) {
      return word;
    }
  }
  return source[source.length - 1];
}

function getWordWeight(uid, direction) {
  const stats = state.store.progress.wordStats[uid] || createWordStat();
  const directionStats = stats.byDirection?.[direction] || { asked: 0, incorrect: 0, correct: 0 };
  const overallAccuracy = stats.asked ? stats.correct / stats.asked : 0;
  const directionAccuracy = directionStats.asked ? directionStats.correct / directionStats.asked : 0;
  return (
    1 +
    stats.incorrect * 1.8 +
    directionStats.incorrect * 1.4 +
    (stats.asked ? (1 - overallAccuracy) * 2.2 : 1.3) +
    (directionStats.asked ? (1 - directionAccuracy) * 1.6 : 0.8)
  );
}

function buildMultipleChoiceOptions(word, pool, direction) {
  const choiceLabel = (item) => (direction === "en-to-ja" ? item.japanese : item.english);
  const seen = new Set([normalizeBasic(choiceLabel(word))]);
  const distractors = [];
  const shuffled = shuffle([...pool].filter((candidate) => candidate.uid !== word.uid));

  for (const candidate of shuffled) {
    const label = normalizeBasic(choiceLabel(candidate));
    if (seen.has(label)) {
      continue;
    }
    seen.add(label);
    distractors.push(candidate);
    if (distractors.length === 3) {
      break;
    }
  }

  return shuffle([word, ...distractors]).map((candidate, index) => ({
    index: index + 1,
    uid: candidate.uid,
    label: choiceLabel(candidate),
    isCorrect: candidate.uid === word.uid,
  }));
}

function renderAll() {
  renderHeader();
  renderStudyFlow();
  renderQuestion();
  renderSession();
  renderScore();
  renderQuickAnalytics();
  renderDashboard();
  renderHistory();
  renderSettings();
}

function renderHeader() {
  const dataset = getActiveDataset();
  const analytics = computeAnalytics();
  el.datasetCount.textContent = String(dataset.words.length);
  el.totalAnswered.textContent = String(analytics.totalAttempts);
  el.overallAccuracy.textContent = formatPercent(analytics.accuracy);
}

function renderStudyFlow() {
  const dataset = getActiveDataset();
  const selectedWords = getWordsInRange(dataset.words);
  const rangeLabel = formatRangeLabel(dataset);
  const setupCopy = `${MODE_LABELS[state.store.settings.mode]} / ${DIRECTION_LABELS[state.store.settings.direction]} / ${selectedWords.length} words selected`;

  el.studyScreens.forEach((screen) => {
    screen.classList.toggle("is-active", screen.id === `study-screen-${state.studyScreen}`);
  });

  el.setupDatasetChip.textContent = dataset.meta.name;
  el.setupRangeChip.textContent = `Range ${rangeLabel}`;
  el.setupSummaryText.textContent = `${setupCopy}. Setup → Quiz → Score.`;
  el.rangeHint.textContent = `Available range: ${dataset.words[0]?.sectionNumber ?? 1} - ${dataset.words[dataset.words.length - 1]?.sectionNumber ?? dataset.words.length} / Current selection: ${selectedWords.length} words`;
}

function renderQuestion() {
  const question = state.currentQuestion;
  const dataset = getActiveDataset();

  if (!question) {
    el.questionBadge.textContent = "Ready";
    el.promptText.textContent = "セッションを開始すると問題が表示されます。";
    el.directionChip.textContent = DIRECTION_LABELS[state.store.settings.direction];
    el.modeChip.textContent = MODE_LABELS[state.store.settings.mode];
    el.wordIdChip.textContent = dataset.meta.name;
    el.favoriteButton.textContent = "☆";
    el.answerArea.replaceChildren(makeEmptyState("左の設定を選んでから、セッション開始を押してください。"));
    el.revealButton.disabled = true;
    el.speakButton.disabled = true;
    if (state.studyScreen === "setup") {
      setFeedback("左の設定を選んでから、セッション開始を押してください。");
    }
    return;
  }

  el.questionBadge.textContent = `Q${state.session.answered + 1}`;
  el.promptText.textContent = question.prompt;
  el.directionChip.textContent = DIRECTION_LABELS[question.direction];
  el.modeChip.textContent = MODE_LABELS[question.mode];
  el.wordIdChip.textContent = `${question.datasetName} / No.${question.word.id}`;
  el.favoriteButton.textContent = state.store.progress.favorites[question.word.uid] ? "★" : "☆";
  el.revealButton.disabled = false;
  el.speakButton.disabled = false;
  el.answerArea.replaceChildren();

  if (question.mode === "multiple") {
    renderMultipleChoice(question);
  } else if (question.mode === "input") {
    renderInputMode(question);
  } else {
    renderFlashcardMode(question);
  }

  el.revealButton.textContent = question.mode === "flashcard" && question.revealed ? "答えを隠す" : "答えを見る";
  el.nextQuestionButton.textContent = state.session.completed ? "新しいセッション" : "次の問題";
}

function renderMultipleChoice(question) {
  const grid = document.createElement("div");
  grid.className = "option-grid";

  question.options.forEach((option) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "option-button";
    button.textContent = `${option.index}. ${option.label}`;
    button.disabled = question.answered;

    if (question.answered) {
      if (option.isCorrect) {
        button.classList.add("is-correct");
      }
      if (option.uid === question.selectedChoiceUid && !option.isCorrect) {
        button.classList.add("is-wrong");
      }
    }

    button.addEventListener("click", () => submitMultipleChoice(option.uid, option.label));
    grid.append(button);
  });

  el.answerArea.append(grid);
}

function renderInputMode(question) {
  const block = document.createElement("div");
  block.className = "input-block";

  const hint = document.createElement("p");
  hint.textContent = question.direction === "en-to-ja"
    ? "日本語で入力してください。複数の言い換えのうち一つが合っていれば正解です。"
    : "英語で入力してください。大文字小文字は区別しません。";
  block.append(hint);

  const row = document.createElement("div");
  row.className = "input-row";

  const input = document.createElement("input");
  input.className = "answer-input";
  input.type = "text";
  input.placeholder = "ここに答えを入力";
  input.autocomplete = "off";
  input.spellcheck = false;
  input.disabled = question.answered;

  const submit = document.createElement("button");
  submit.type = "button";
  submit.className = "primary-button";
  submit.textContent = question.answered ? "採点済み" : "採点";
  submit.disabled = question.answered;

  submit.addEventListener("click", () => submitTypedAnswer(input.value));
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      submitTypedAnswer(input.value);
    }
  });

  row.append(input, submit);
  block.append(row);
  el.answerArea.append(block);

  if (!question.answered) {
    requestAnimationFrame(() => input.focus());
  }
}

function renderFlashcardMode(question) {
  const block = document.createElement("div");
  block.className = "flash-block";

  const lead = document.createElement("p");
  lead.textContent = question.revealed
    ? "答えを確認できたら、自分の感触で自己採点できます。"
    : "答えを見るボタンで裏面を開いてください。";
  block.append(lead);

  const answer = document.createElement("div");
  answer.className = `flash-answer${question.revealed ? " is-visible" : ""}`;
  answer.textContent = `${question.word.english}\n${question.word.japanese}`;
  block.append(answer);

  if (question.revealed) {
    const row = document.createElement("div");
    row.className = "button-row";

    const known = document.createElement("button");
    known.type = "button";
    known.className = "primary-button";
    known.textContent = "覚えていた";
    known.disabled = question.answered;
    known.addEventListener("click", () => submitFlashcardResult(true));

    const unsure = document.createElement("button");
    unsure.type = "button";
    unsure.className = "secondary-button";
    unsure.textContent = "まだあやしい";
    unsure.disabled = question.answered;
    unsure.addEventListener("click", () => submitFlashcardResult(false));

    row.append(known, unsure);
    block.append(row);
  }

  el.answerArea.append(block);
}

function renderSession() {
  const dataset = getActiveDataset();
  const rangeLabel = formatRangeLabel(dataset);
  const limit = state.session.limit;
  const progressLabel = limit > 0 ? `${state.session.answered} / ${limit}` : `${state.session.answered} / ∞`;
  const accuracy = state.session.answered ? state.session.correct / state.session.answered : 0;
  const modeLabel = `${MODE_LABELS[state.store.settings.mode]} / ${DIRECTION_LABELS[state.store.settings.direction]}`;
  el.sessionProgress.textContent = progressLabel;
  el.sessionAccuracy.textContent = state.session.answered ? formatPercent(accuracy) : "-";
  el.sessionStreak.textContent = String(state.session.streak);
  el.sessionModeBadge.textContent = state.session.active ? modeLabel : "Idle";

  el.quizProgress.textContent = progressLabel;
  el.quizAccuracy.textContent = state.session.answered ? formatPercent(accuracy) : "-";
  el.quizStreak.textContent = String(state.session.streak);
  el.quizStatusBadge.textContent = state.session.active ? "In progress" : "Idle";

  const metaItems = [
    `Mode: ${MODE_LABELS[state.store.settings.mode]}`,
    `Direction: ${DIRECTION_LABELS[state.store.settings.direction]}`,
    `Focus: ${getFocusLabel(state.store.settings.focus)}`,
    `Range: ${rangeLabel}`,
    `Dataset: ${dataset.meta.name}`,
  ];
  el.quizMetaList.replaceChildren();
  metaItems.forEach((text) => {
    const item = document.createElement("li");
    item.textContent = text;
    el.quizMetaList.append(item);
  });
}

function renderScore() {
  const summary = state.lastSessionSummary;
  const dataset = getActiveDataset();
  const rangeLabel = formatRangeLabel(dataset);

  if (!summary) {
    el.scoreDatasetChip.textContent = dataset.meta.name;
    el.scoreRangeChip.textContent = `Range ${rangeLabel}`;
    el.scoreSummaryText.textContent = "Session results appear here after you finish.";
    el.scoreCorrectCount.textContent = "0";
    el.scoreAnsweredCount.textContent = "0問中 0問";
    el.scoreAccuracyValue.textContent = "0%";
    el.scoreBestStreak.textContent = "0";
    el.scoreMistakeBadge.textContent = "0件";
    renderWordList(el.scoreMistakeList, [], () => ({ title: "", detail: "" }));
    renderWordList(el.scoreNoteList, [], () => ({ title: "", detail: "" }));
    return;
  }

  el.scoreDatasetChip.textContent = summary.datasetName;
  el.scoreRangeChip.textContent = `Range ${summary.rangeLabel}`;
  el.scoreSummaryText.textContent = `${summary.correct} correct out of ${summary.answered}. Review missed words and go again.`;
  el.scoreCorrectCount.textContent = String(summary.correct);
  el.scoreAnsweredCount.textContent = `${summary.answered}問中 ${summary.correct}問`;
  el.scoreAccuracyValue.textContent = formatPercent(summary.accuracy);
  el.scoreBestStreak.textContent = String(summary.bestStreak);
  el.scoreMistakeBadge.textContent = `${summary.mistakes.length}件`;

  renderWordList(el.scoreMistakeList, summary.mistakes, (attempt) => ({
    title: `${attempt.english} / ${attempt.japanese}`,
    detail: `回答: ${attempt.userAnswer || "未入力"} / 正解: ${attempt.expectedAnswer}`,
  }));

  const notes = buildScoreNotes(summary);
  renderWordList(el.scoreNoteList, notes, (note) => note);
}

function renderQuickAnalytics() {
  if (!el.dashboardWeakWords || !el.confusionList) {
    return;
  }
  const analytics = computeAnalytics();
  state.latestAnalytics = analytics;
}

function buildScoreNotes(summary) {
  const notes = [];
  if (summary.accuracy >= 0.85) {
    notes.push({ title: "Strong result", detail: "Try Input or Japanese → English next for a tougher round." });
  } else if (summary.accuracy >= 0.6) {
    notes.push({ title: "Almost there", detail: "Running the same range one more time should help it stick." });
  } else {
    notes.push({ title: "Review mode", detail: "Mark missed words as favorites and run them again." });
  }

  notes.push({
    title: "Suggested next setup",
    detail: `${MODE_LABELS[state.store.settings.mode]} / ${DIRECTION_LABELS[state.store.settings.direction]} / Range ${summary.rangeLabel}`,
  });

  if (summary.mistakes.length) {
    notes.push({
      title: "Start here",
      detail: `${summary.mistakes[0].english} looks like the best word to review first.`,
    });
  }

  return notes;
}

function renderDashboard() {
  const analytics = computeAnalytics();
  el.metricAttempts.textContent = String(analytics.totalAttempts);
  el.metricAccuracy.textContent = formatPercent(analytics.accuracy);
  el.metricBestStreak.textContent = String(state.store.progress.bestStreak || 0);
  el.metricWeakCount.textContent = String(analytics.weakWords.length);

  renderBreakdown(el.modeBreakdown, analytics.modeBreakdown, MODE_LABELS);
  renderBreakdown(el.directionBreakdown, analytics.directionBreakdown, DIRECTION_LABELS);
  renderWordList(el.dashboardWeakWords, analytics.weakWords.slice(0, 10), (entry) => {
    const accuracy = entry.stats.asked ? entry.stats.correct / entry.stats.asked : 0;
    return {
      title: entry.word.english,
      detail: `${entry.word.japanese} / ${entry.stats.incorrect}ミス / 正答率 ${formatPercent(accuracy)}`,
    };
  });
  renderWordList(el.confusionList, analytics.confusions.slice(0, 10), (entry) => ({
    title: `${entry.source.english} と ${entry.target.english}`,
    detail: `${entry.count}回混同`,
  }));
}

function renderHistory() {
  const filter = el.historyFilterSelect.value;
  const attempts = state.store.progress.attempts.filter((attempt) => {
    if (filter === "correct") {
      return attempt.correct;
    }
    if (filter === "incorrect") {
      return !attempt.correct;
    }
    return true;
  });

  el.historyList.replaceChildren();
  if (!attempts.length) {
    el.historyList.append(makeEmptyState("まだ履歴がありません。", "li"));
    return;
  }

  attempts.forEach((attempt) => {
    const item = document.createElement("li");
    item.className = `history-item ${attempt.correct ? "correct" : "incorrect"}`;

    const topline = document.createElement("div");
    topline.className = "history-item-topline";
    const word = document.createElement("strong");
    word.textContent = `${attempt.english} / ${attempt.japanese}`;
    const time = document.createElement("time");
    time.textContent = dateFormatter.format(attempt.at);
    topline.append(word, time);

    const meta = document.createElement("div");
    meta.className = "history-item-meta";
    meta.textContent = `${MODE_LABELS[attempt.mode]} / ${DIRECTION_LABELS[attempt.direction]} / ${attempt.correct ? "正解" : "不正解"}`;

    const detail = document.createElement("div");
    detail.textContent = `回答: ${attempt.userAnswer || "未入力"} / 正解: ${attempt.expectedAnswer}`;

    item.append(topline, meta, detail);
    el.historyList.append(item);
  });
}

function renderSettings() {
  populateDatasetSelect();
  const activeDataset = getActiveDataset();
  el.datasetStatusText.textContent = state.store.customDataset?.words?.length
    ? `現在のデータセット: ${activeDataset.meta.name} (${activeDataset.words.length}語)`
    : "既定データとして system.xlsx 由来の単語帳を読み込んでいます。";
  el.backgroundHint.textContent = state.store.settings.customBackground || state.temporaryBackground
    ? "背景画像を適用中です。テーマ変更と組み合わせて使えます。"
    : "テーマだけでも使えます。画像が大きい場合は保存せず、その場だけで適用します。";
}

function submitMultipleChoice(selectedUid, label) {
  const question = state.currentQuestion;
  if (!question || question.answered) {
    return;
  }

  const correct = selectedUid === question.word.uid;
  question.answered = true;
  question.correct = correct;
  question.selectedChoiceUid = selectedUid;

  recordAttempt({
    correct,
    direction: question.direction,
    mode: question.mode,
    question,
    selectedChoiceUid: selectedUid,
    userAnswer: label,
  });

  setFeedback(
    `${correct
      ? `正解です。\n${question.word.english} = ${question.word.japanese}`
      : `不正解です。\n正解: ${question.word.english} = ${question.word.japanese}`}${getSessionCompletionSuffix()}`,
    correct ? "correct" : "wrong"
  );

  renderAll();

  if (correct && state.store.settings.autoAdvance && !state.session.completed) {
    state.pendingAutoAdvance = window.setTimeout(() => nextQuestion(), AUTO_ADVANCE_DELAY);
  }
}

function submitTypedAnswer(value) {
  const question = state.currentQuestion;
  if (!question || question.answered) {
    return;
  }

  const result = evaluateTypedAnswer(value, question.word, question.direction);
  question.answered = true;
  question.correct = result.correct;

  recordAttempt({
    correct: result.correct,
    direction: question.direction,
    mode: question.mode,
    question,
    userAnswer: value.trim(),
  });

  setFeedback(
    `${result.correct
      ? `正解です。\n${question.word.english} = ${question.word.japanese}`
      : `不正解です。\n正解候補: ${result.accepted.join(" / ") || formatExpectedAnswer(question.word, question.direction)}`}${getSessionCompletionSuffix()}`,
    result.correct ? "correct" : "wrong"
  );
  renderAll();
}

function submitFlashcardResult(correct) {
  const question = state.currentQuestion;
  if (!question || question.answered || !question.revealed) {
    return;
  }

  question.answered = true;
  question.correct = correct;

  recordAttempt({
    correct,
    direction: question.direction,
    mode: question.mode,
    question,
    userAnswer: correct ? "覚えていた" : "まだあやしい",
  });

  setFeedback(
    `${correct
      ? "感触を記録しました。\nこの単語は覚えていたことにします。"
      : "感触を記録しました。\nこの単語は苦手候補として重み付けされます。"}${getSessionCompletionSuffix()}`,
    correct ? "correct" : "wrong"
  );
  renderAll();
}

function revealCurrentAnswer() {
  const question = state.currentQuestion;
  if (!question) {
    return;
  }

  if (question.mode === "flashcard") {
    question.revealed = !question.revealed;
    renderQuestion();
    if (question.revealed) {
      setFeedback("裏面を表示しました。覚えていたかどうかを下のボタンで残せます。");
    } else {
      setFeedback("表面に戻しました。");
    }
    return;
  }

  setFeedback(`答え: ${question.word.english} = ${question.word.japanese}`);
}

function recordAttempt({ correct, direction, mode, question, userAnswer, selectedChoiceUid = "" }) {
  const at = Date.now();
  const expectedAnswer = formatExpectedAnswer(question.word, direction);
  const attempt = {
    uid: question.word.uid,
    datasetId: state.store.settings.activeDataset,
    wordId: question.word.id,
    english: question.word.english,
    japanese: question.word.japanese,
    direction,
    mode,
    correct,
    userAnswer,
    expectedAnswer,
    selectedChoiceUid,
    at,
  };

  state.store.progress.attempts.unshift(attempt);
  state.store.progress.attempts = state.store.progress.attempts.slice(0, HISTORY_LIMIT);

  const stat = state.store.progress.wordStats[question.word.uid] || createWordStat();
  stat.asked += 1;
  stat.correct += correct ? 1 : 0;
  stat.incorrect += correct ? 0 : 1;
  stat.lastSeenAt = at;

  stat.byDirection[direction] = stat.byDirection[direction] || { asked: 0, correct: 0, incorrect: 0 };
  stat.byDirection[direction].asked += 1;
  stat.byDirection[direction].correct += correct ? 1 : 0;
  stat.byDirection[direction].incorrect += correct ? 0 : 1;

  stat.byMode[mode] = stat.byMode[mode] || { asked: 0, correct: 0, incorrect: 0 };
  stat.byMode[mode].asked += 1;
  stat.byMode[mode].correct += correct ? 1 : 0;
  stat.byMode[mode].incorrect += correct ? 0 : 1;

  if (!correct && mode === "multiple" && selectedChoiceUid && selectedChoiceUid !== question.word.uid) {
    stat.confusions[selectedChoiceUid] = (stat.confusions[selectedChoiceUid] || 0) + 1;
  }

  state.store.progress.wordStats[question.word.uid] = stat;
  state.session.answered += 1;
  state.session.correct += correct ? 1 : 0;
  state.session.streak = correct ? state.session.streak + 1 : 0;
  state.session.bestStreak = Math.max(state.session.bestStreak, state.session.streak);
  state.session.attempts.unshift(attempt);
  state.store.progress.bestStreak = Math.max(state.store.progress.bestStreak || 0, state.session.streak);

  if (state.session.limit > 0 && state.session.answered >= state.session.limit) {
    state.session.completed = true;
    state.session.active = false;
    state.lastSessionSummary = summarizeSession();
    state.studyScreen = "score";
    state.currentQuestion = null;
    clearPendingAutoAdvance();
  }

  saveStore();
}

function summarizeSession() {
  const dataset = getActiveDataset();
  const attempts = [...state.session.attempts];
  const mistakes = attempts.filter((attempt) => !attempt.correct);
  return {
    answered: state.session.answered,
    correct: state.session.correct,
    accuracy: state.session.answered ? state.session.correct / state.session.answered : 0,
    bestStreak: state.session.bestStreak,
    mistakes,
    rangeLabel: formatRangeLabel(dataset),
    datasetName: dataset.meta.name,
  };
}

function evaluateTypedAnswer(value, word, direction) {
  const guess = value.trim();
  const accepted = direction === "en-to-ja" ? getJapaneseAnswerCandidates(word.japanese) : getEnglishAnswerCandidates(word.english);
  if (!guess) {
    return { correct: false, accepted };
  }

  const normalizedGuess = direction === "en-to-ja" ? normalizeJapanese(guess) : normalizeEnglish(guess);
  const correct = accepted.some((candidate) => {
    const normalizedCandidate = direction === "en-to-ja" ? normalizeJapanese(candidate) : normalizeEnglish(candidate);
    if (normalizedGuess === normalizedCandidate) {
      return true;
    }
    if (direction === "en-to-ja") {
      return normalizedCandidate.includes(normalizedGuess) || normalizedGuess.includes(normalizedCandidate);
    }
    return false;
  });

  return { correct, accepted };
}

function getEnglishAnswerCandidates(text) {
  const base = text.replace(/[()]/g, " ");
  const segments = [base, ...base.split(/[;,/]/g)];
  return [...new Set(segments.map((segment) => segment.trim()).filter(Boolean))];
}

function getJapaneseAnswerCandidates(text) {
  const base = text
    .replace(/[（(][^）)]*[）)]/g, "")
    .replace(/・/g, " ")
    .trim();
  const segments = [base, ...base.split(/[、,;；/]/g)];
  return [...new Set(segments.map((segment) => segment.trim().replace(/^～/, "")).filter(Boolean))];
}

function normalizeEnglish(text) {
  return text
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[“”"'`]/g, "")
    .replace(/[._!?]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeJapanese(text) {
  return text
    .normalize("NFKC")
    .replace(/[~〜～]/g, "")
    .replace(/[「」『』【】\[\]()（）]/g, "")
    .replace(/[　\s]/g, "")
    .replace(/[。．、,・]/g, "")
    .trim();
}

function normalizeBasic(text) {
  return text.normalize("NFKC").replace(/[　\s]/g, "").trim();
}

function formatExpectedAnswer(word, direction) {
  return direction === "en-to-ja" ? word.japanese : word.english;
}

function computeAnalytics() {
  const attempts = state.store.progress.attempts;
  const totalAttempts = attempts.length;
  const totalCorrect = attempts.filter((attempt) => attempt.correct).length;
  return {
    totalAttempts,
    accuracy: totalAttempts ? totalCorrect / totalAttempts : 0,
    weakWords: getWeakWords(20),
    confusions: getConfusionPairs(20),
    modeBreakdown: aggregateAttemptsBy(attempts, "mode"),
    directionBreakdown: aggregateAttemptsBy(attempts, "direction"),
  };
}

function aggregateAttemptsBy(attempts, key) {
  const stats = new Map();
  for (const attempt of attempts) {
    const current = stats.get(attempt[key]) || { asked: 0, correct: 0 };
    current.asked += 1;
    current.correct += attempt.correct ? 1 : 0;
    stats.set(attempt[key], current);
  }
  return [...stats.entries()].map(([name, value]) => ({
    name,
    asked: value.asked,
    accuracy: value.asked ? value.correct / value.asked : 0,
  }));
}

function getWeakWords(limit) {
  const lookup = getWordLookup();
  return Object.entries(state.store.progress.wordStats)
    .map(([uid, stats]) => {
      const word = lookup.get(uid);
      if (!word || !stats.asked || !stats.incorrect) {
        return null;
      }
      const score = stats.incorrect * 2.4 + (1 - stats.correct / stats.asked) * stats.asked;
      return { word, stats, score };
    })
    .filter(Boolean)
    .sort((left, right) => right.score - left.score)
    .slice(0, limit);
}

function getConfusionPairs(limit) {
  const lookup = getWordLookup();
  const pairs = [];
  for (const [uid, stats] of Object.entries(state.store.progress.wordStats)) {
    for (const [targetUid, count] of Object.entries(stats.confusions || {})) {
      const source = lookup.get(uid);
      const target = lookup.get(targetUid);
      if (!source || !target || !count) {
        continue;
      }
      pairs.push({ source, target, count });
    }
  }
  return pairs.sort((left, right) => right.count - left.count).slice(0, limit);
}

function renderBreakdown(container, items, labels) {
  container.replaceChildren();
  if (!items.length) {
    container.append(makeEmptyState("まだ分析できるデータがありません。"));
    return;
  }

  items.forEach((item) => {
    const block = document.createElement("div");
    block.className = "breakdown-item";

    const topline = document.createElement("div");
    topline.className = "breakdown-topline";
    const label = document.createElement("strong");
    label.textContent = labels[item.name] || item.name;
    const value = document.createElement("span");
    value.textContent = formatPercent(item.accuracy);
    topline.append(label, value);

    const meter = document.createElement("div");
    meter.className = "meter";
    const fill = document.createElement("span");
    fill.style.setProperty("--ratio", String(item.accuracy * 100));
    meter.append(fill);

    const note = document.createElement("div");
    note.textContent = `${item.asked}問`;

    block.append(topline, meter, note);
    container.append(block);
  });
}

function renderWordList(container, items, mapItem) {
  container.replaceChildren();
  if (!items.length) {
    container.append(makeEmptyState("まだデータがありません。", "li"));
    return;
  }

  items.forEach((item) => {
    const row = document.createElement("li");
    const mapped = mapItem(item);
    const title = document.createElement("strong");
    title.textContent = mapped.title;
    const detail = document.createElement("span");
    detail.textContent = mapped.detail;
    row.append(title, detail);
    container.append(row);
  });
}

function renderAttemptList(container, attempts, compact = false) {
  container.replaceChildren();
  if (!attempts.length) {
    container.append(makeEmptyState("まだ履歴がありません。", "li"));
    return;
  }

  attempts.forEach((attempt) => {
    const row = document.createElement("li");
    const title = document.createElement("strong");
    title.textContent = `${attempt.correct ? "○" : "×"} ${attempt.english}`;
    const detail = document.createElement("span");
    detail.textContent = compact
      ? `${MODE_LABELS[attempt.mode]} / ${DIRECTION_LABELS[attempt.direction]}`
      : `${attempt.userAnswer || "未入力"} / ${attempt.expectedAnswer}`;
    row.append(title, detail);
    container.append(row);
  });
}

function toggleFavoriteForCurrentWord() {
  const question = state.currentQuestion;
  if (!question) {
    return;
  }
  const favorites = state.store.progress.favorites;
  favorites[question.word.uid] = !favorites[question.word.uid];
  if (!favorites[question.word.uid]) {
    delete favorites[question.word.uid];
  }
  saveStore();
  renderQuestion();
}

function bindExtendedEvents() {
  el.csvFileInput.addEventListener("change", async () => {
    const file = el.csvFileInput.files?.[0];
    if (!file) {
      return;
    }
    await importCsvFile(file);
    el.csvFileInput.value = "";
  });

  el.backgroundFileInput.addEventListener("change", async () => {
    const file = el.backgroundFileInput.files?.[0];
    if (!file) {
      return;
    }
    await applyBackgroundFromFile(file);
    el.backgroundFileInput.value = "";
  });

  el.clearBackgroundButton.addEventListener("click", clearBackgroundImage);

  el.resetProgressButton.addEventListener("click", () => {
    if (!confirm("学習履歴を消去します。よろしいですか？")) {
      return;
    }
    state.store.progress = structuredClone(DEFAULT_PROGRESS);
    state.session = createSessionState();
    state.currentQuestion = null;
    state.lastSessionSummary = null;
    state.studyScreen = "setup";
    saveStore();
    renderAll();
  });

  el.resetAllButton.addEventListener("click", () => {
    if (!confirm("履歴・設定・カスタムCSVをまとめて消去します。よろしいですか？")) {
      return;
    }
    state.store = {
      settings: { ...DEFAULT_SETTINGS },
      progress: structuredClone(DEFAULT_PROGRESS),
      customDataset: null,
    };
    state.temporaryBackground = "";
    syncStore();
    hydrateControls();
    applyTheme();
    state.session = createSessionState();
    state.currentQuestion = null;
    state.lastSessionSummary = null;
    state.studyScreen = "setup";
    saveStore();
    renderAll();
  });
}

bindExtendedEvents();

async function importCsvFile(file) {
  try {
    const text = await file.text();
    const parsedWords = parseWordCsv(text);
    if (!parsedWords.length) {
      alert("CSV から単語を読み取れませんでした。");
      return;
    }

    const datasetId = `custom-${Date.now()}`;
    state.store.customDataset = {
      meta: {
        id: datasetId,
        name: file.name.replace(/\.[^.]+$/, "") || "カスタムCSV",
        source: file.name,
        count: parsedWords.length,
      },
      words: parsedWords.map((word, index) => ({
        id: String(word.id || index + 1),
        english: word.english,
        japanese: word.japanese,
      })),
    };
    state.store.settings.activeDataset = datasetId;
    state.studyScreen = "setup";
    state.lastSessionSummary = null;
    saveStore();
    renderAll();
    switchView("study");
  } catch (error) {
    console.error(error);
    alert("CSV の読み込みに失敗しました。列は 英語 / 日本語 の形になっているか確認してください。");
  }
}

function parseWordCsv(text) {
  const sanitized = text.replace(/^\uFEFF/, "");
  const rows = parseDelimitedText(sanitized).filter((row) => row.some((cell) => cell.trim()));
  if (!rows.length) {
    return [];
  }

  const header = rows[0].map((cell) => cell.trim().toLowerCase());
  const hasHeader = header.some((cell) => /(english|英語|japanese|日本語|meaning|意味|word|単語|id|番号)/.test(cell));

  const pickIndex = (patterns, fallback) => {
    const index = header.findIndex((cell) => patterns.some((pattern) => pattern.test(cell)));
    return index >= 0 ? index : fallback;
  };

  const englishIndex = hasHeader ? pickIndex([/english/, /英語/, /word/, /単語/], rows[0].length >= 3 ? 1 : 0) : rows[0].length >= 3 ? 1 : 0;
  const japaneseIndex = hasHeader ? pickIndex([/japanese/, /日本語/, /meaning/, /意味/], rows[0].length >= 3 ? 2 : 1) : rows[0].length >= 3 ? 2 : 1;
  const idIndex = hasHeader ? pickIndex([/^id$/, /番号/], -1) : rows[0].length >= 3 ? 0 : -1;
  const dataRows = hasHeader ? rows.slice(1) : rows;

  return dataRows
    .map((row, index) => ({
      id: row[idIndex]?.trim() || String(index + 1),
      english: row[englishIndex]?.trim() || "",
      japanese: row[japaneseIndex]?.trim() || "",
    }))
    .filter((row) => row.english && row.japanese);
}

function parseDelimitedText(text) {
  const delimiter = detectDelimiter(text);
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const next = text[index + 1];

    if (character === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && character === delimiter) {
      row.push(cell);
      cell = "";
      continue;
    }

    if (!inQuotes && (character === "\n" || character === "\r")) {
      if (character === "\r" && next === "\n") {
        index += 1;
      }
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += character;
  }

  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }

  return rows;
}

function detectDelimiter(text) {
  const sample = text.split(/\r?\n/).slice(0, 8).join("\n");
  const counts = [
    { delimiter: ",", count: (sample.match(/,/g) || []).length },
    { delimiter: "\t", count: (sample.match(/\t/g) || []).length },
    { delimiter: ";", count: (sample.match(/;/g) || []).length },
  ];
  return counts.sort((left, right) => right.count - left.count)[0].delimiter;
}

async function applyBackgroundFromFile(file) {
  const dataUrl = await readFileAsDataUrl(file);
  state.temporaryBackground = dataUrl;

  if (dataUrl.length <= PERSISTABLE_BACKGROUND_LIMIT) {
    state.store.settings.customBackground = dataUrl;
    saveStore();
  } else {
    state.store.settings.customBackground = "";
    saveStore();
  }

  applyTheme();
  renderSettings();
}

function clearBackgroundImage() {
  state.temporaryBackground = "";
  state.store.settings.customBackground = "";
  saveStore();
  applyTheme();
  renderSettings();
}

function applyTheme() {
  document.body.dataset.theme = state.store.settings.theme;
  const backgroundData = state.temporaryBackground || state.store.settings.customBackground;
  const backgroundImage = backgroundData
    ? `linear-gradient(135deg, rgba(7, 17, 31, 0.38), rgba(7, 17, 31, 0.68)), url("${backgroundData}") center / cover no-repeat fixed`
    : "none";
  document.body.style.setProperty("--background-image", backgroundImage);
}

function exportStudyData() {
  const blob = new Blob([JSON.stringify(state.store, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `sew-word-studio-${Date.now()}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function setFeedback(message, tone = "") {
  el.feedbackBox.textContent = message;
  el.feedbackBox.classList.remove("is-correct", "is-wrong");
  if (tone === "correct") {
    el.feedbackBox.classList.add("is-correct");
  }
  if (tone === "wrong") {
    el.feedbackBox.classList.add("is-wrong");
  }
}

function formatPercent(value) {
  return `${Math.round(value * 100)}%`;
}

function getSessionCompletionSuffix() {
  if (!state.session.completed) {
    return "";
  }
  return `\n\nセッション完了: ${state.session.answered}問中 ${state.session.correct}問正解`;
}

function clearPendingAutoAdvance() {
  if (state.pendingAutoAdvance) {
    clearTimeout(state.pendingAutoAdvance);
    state.pendingAutoAdvance = null;
  }
}

function shuffle(items) {
  const result = [...items];
  for (let index = result.length - 1; index > 0; index -= 1) {
    const randomIndex = Math.floor(Math.random() * (index + 1));
    [result[index], result[randomIndex]] = [result[randomIndex], result[index]];
  }
  return result;
}

function makeEmptyState(text, tagName = "div") {
  const item = document.createElement(tagName);
  item.className = "empty-state";
  item.textContent = text;
  return item;
}

async function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function speakText(text) {
  if (!("speechSynthesis" in window)) {
    setFeedback("このブラウザでは読み上げに対応していません。");
    return;
  }
  window.speechSynthesis.cancel();
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = "en-US";
  utterance.rate = 0.92;
  window.speechSynthesis.speak(utterance);
}

function handleKeyboardShortcuts(event) {
  if (state.currentView !== "study" || state.studyScreen !== "quiz" || !state.currentQuestion) {
    return;
  }

  const activeTag = document.activeElement?.tagName;
  const typing = activeTag === "INPUT" || activeTag === "TEXTAREA";
  if (typing && event.key !== "Enter") {
    return;
  }

  if (state.currentQuestion.mode === "multiple" && !state.currentQuestion.answered) {
    const index = Number(event.key);
    if (index >= 1 && index <= state.currentQuestion.options.length) {
      const option = state.currentQuestion.options[index - 1];
      submitMultipleChoice(option.uid, option.label);
      return;
    }
  }

  if (state.currentQuestion.mode === "flashcard" && event.code === "Space") {
    event.preventDefault();
    revealCurrentAnswer();
    return;
  }

  if (event.key === "Enter" && !typing) {
    event.preventDefault();
    if (state.session.completed) {
      startSession();
      return;
    }
    nextQuestion();
  }
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) {
    return;
  }
  if (!window.isSecureContext && location.hostname !== "127.0.0.1" && location.hostname !== "localhost") {
    return;
  }
  navigator.serviceWorker.register("./sw.js").then((registration) => {
    registration.update().catch(() => {});
  }).catch((error) => {
    console.warn("Service worker registration failed:", error);
  });
}

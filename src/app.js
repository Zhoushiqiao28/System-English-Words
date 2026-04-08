import { defaultWords, defaultDatasetMeta } from "./data/defaultWords.js";

const STORAGE_KEY = "sew-word-studio.store.v2";
const HISTORY_LIMIT = 600;
const SESSION_LOG_LIMIT = 180;
const AUTO_ADVANCE_DELAY = 720;
const PERSISTABLE_BACKGROUND_LIMIT = 650_000;
const CARD_SWIPE_THRESHOLD = 72;
const ADVANCE_SWIPE_THRESHOLD = 68;
const CALENDAR_DAY_COUNT = 35;
const RESULTS_PAGE_SIZE = 12;
const PRIMARY_VIEWS = new Set(["system", "details", "study", "results", "dashboard"]);

const MODE_LABELS = {
  multiple: "4-Choice",
  input: "Input",
  flashcard: "Flashcards",
  card: "Word Cards",
};

const DIRECTION_LABELS = {
  "en-to-ja": "English to Japanese",
  "ja-to-en": "Japanese to English",
  mixed: "Mixed",
};

const FOCUS_LABELS = {
  all: "All",
  known: "Known",
  weak: "Weak",
  favorites: "Favorites",
  "recent-mistakes": "Recent Mistakes",
};

const TRANSITION_LABELS = {
  off: "Manual",
  correct: "After Correct",
  always: "After Every Answer",
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
  autoAdvanceMode: "correct",
  customBackground: "",
};

const DEFAULT_PROGRESS = {
  attempts: [],
  sessions: [],
  wordStats: {},
  memorized: {},
  favorites: {},
  bestStreak: 0,
  studyTimeMs: 0,
};

const DEFAULT_DATASET = {
  meta: defaultDatasetMeta,
  words: prepareWords(defaultDatasetMeta.id, defaultWords),
};

const dateFormatter = new Intl.DateTimeFormat("ja-JP", {
  month: "numeric",
  day: "numeric",
  hour: "2-digit",
  minute: "2-digit",
});

const el = {
  answerArea: document.querySelector("#answerArea"),
  autoAdvanceSelect: document.querySelector("#autoAdvanceSelect"),
  autoPronounceToggle: document.querySelector("#autoPronounceToggle"),
  backToDetailsButton: document.querySelector("#backToDetailsButton"),
  backToSystemButton: document.querySelector("#backToSystemButton"),
  backgroundFileInput: document.querySelector("#backgroundFileInput"),
  backgroundHint: document.querySelector("#backgroundHint"),
  clearBackgroundButton: document.querySelector("#clearBackgroundButton"),
  clearCacheButton: document.querySelector("#clearCacheButton"),
  closeWeakWordsButton: document.querySelector("#closeWeakWordsButton"),
  closeWordBankButton: document.querySelector("#closeWordBankButton"),
  csvFileInput: document.querySelector("#csvFileInput"),
  datasetSelect: document.querySelector("#datasetSelect"),
  datasetStatusText: document.querySelector("#datasetStatusText"),
  detailsDirectionText: document.querySelector("#detailsDirectionText"),
  detailsFocusText: document.querySelector("#detailsFocusText"),
  detailsModeText: document.querySelector("#detailsModeText"),
  detailsRangeHint: document.querySelector("#detailsRangeHint"),
  detailsSelectionCount: document.querySelector("#detailsSelectionCount"),
  detailsTransitionText: document.querySelector("#detailsTransitionText"),
  directionSelect: document.querySelector("#directionSelect"),
  exitSessionButton: document.querySelector("#exitSessionButton"),
  favoriteButton: document.querySelector("#favoriteButton"),
  feedbackBox: document.querySelector("#feedbackBox"),
  focusSelect: document.querySelector("#focusSelect"),
  goToDetailsButton: document.querySelector("#goToDetailsButton"),
  headerDatasetName: document.querySelector("#headerDatasetName"),
  headerWordCount: document.querySelector("#headerWordCount"),
  metricAttempts: document.querySelector("#metricAttempts"),
  metricBestStreak: document.querySelector("#metricBestStreak"),
  metricStudyTime: document.querySelector("#metricStudyTime"),
  modeSelect: document.querySelector("#modeSelect"),
  nextQuestionButton: document.querySelector("#nextQuestionButton"),
  openWeakWordsButton: document.querySelector("#openWeakWordsButton"),
  openWordBankButton: document.querySelector("#openWordBankButton"),
  promptText: document.querySelector("#promptText"),
  questionBadge: document.querySelector("#questionBadge"),
  rangeEndInput: document.querySelector("#rangeEndInput"),
  rangeStartInput: document.querySelector("#rangeStartInput"),
  repeatSessionButton: document.querySelector("#repeatSessionButton"),
  resetAllButton: document.querySelector("#resetAllButton"),
  resetProgressButton: document.querySelector("#resetProgressButton"),
  resultAccuracyValue: document.querySelector("#resultAccuracyValue"),
  resultAnsweredCount: document.querySelector("#resultAnsweredCount"),
  resultCorrectCount: document.querySelector("#resultCorrectCount"),
  resultDatasetChip: document.querySelector("#resultDatasetChip"),
  resultDuration: document.querySelector("#resultDuration"),
  resultHistoryCount: document.querySelector("#resultHistoryCount"),
  resultHistoryList: document.querySelector("#resultHistoryList"),
  resultNextPageButton: document.querySelector("#resultNextPageButton"),
  resultPageLabel: document.querySelector("#resultPageLabel"),
  resultPrevPageButton: document.querySelector("#resultPrevPageButton"),
  resultRangeChip: document.querySelector("#resultRangeChip"),
  resultSummaryText: document.querySelector("#resultSummaryText"),
  resultBestStreak: document.querySelector("#resultBestStreak"),
  revealButton: document.querySelector("#revealButton"),
  sessionLengthSelect: document.querySelector("#sessionLengthSelect"),
  speakButton: document.querySelector("#speakButton"),
  startSessionButton: document.querySelector("#startSessionButton"),
  studyAccuracyText: document.querySelector("#studyAccuracyText"),
  studyProgressText: document.querySelector("#studyProgressText"),
  studyStreakText: document.querySelector("#studyStreakText"),
  studyStatusCorner: document.querySelector("#studyStatusCorner"),
  systemDatasetName: document.querySelector("#systemDatasetName"),
  systemRangePreview: document.querySelector("#systemRangePreview"),
  systemWordCount: document.querySelector("#systemWordCount"),
  tabButtons: document.querySelectorAll(".tab-button"),
  themeSelect: document.querySelector("#themeSelect"),
  views: document.querySelectorAll(".view"),
  weakCandidateCount: document.querySelector("#weakCandidateCount"),
  weakCandidatePreview: document.querySelector("#weakCandidatePreview"),
  weakWordsDetailList: document.querySelector("#weakWordsDetailList"),
  wordBankList: document.querySelector("#wordBankList"),
  dashboardCalendar: document.querySelector("#dashboardCalendar"),
};

const state = {
  store: loadStore(),
  currentView: "system",
  primaryView: "system",
  returnView: "system",
  currentQuestion: null,
  lastSessionSummary: null,
  pendingAutoAdvance: null,
  temporaryBackground: "",
  latestAnalytics: null,
  resultHistoryPage: 0,
  session: createSessionState(),
};

syncStore();
bindEvents();
hydrateControls();
applyTheme();
renderAll();
applyActiveViewState();
registerServiceWorker();

function cloneData(value) {
  return JSON.parse(JSON.stringify(value));
}

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
    id: "",
    active: false,
    completed: false,
    startedAt: 0,
    completedAt: 0,
    durationMs: 0,
    answered: 0,
    correct: 0,
    streak: 0,
    bestStreak: 0,
    limit: 0,
    attempts: [],
    poolWords: [],
    cycleQueue: [],
    settings: null,
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
        progress: cloneData(DEFAULT_PROGRESS),
        customDataset: null,
      };
    }

    const parsed = JSON.parse(raw);
    const progress = parsed.progress || {};
    return {
      settings: {
        ...DEFAULT_SETTINGS,
        ...(parsed.settings || {}),
      },
      progress: {
        ...cloneData(DEFAULT_PROGRESS),
        ...progress,
        attempts: Array.isArray(progress.attempts) ? progress.attempts : [],
        sessions: Array.isArray(progress.sessions) ? progress.sessions : [],
        wordStats: progress.wordStats || {},
        memorized: progress.memorized || {},
        favorites: progress.favorites || {},
        bestStreak: Number(progress.bestStreak || 0),
        studyTimeMs: Number(progress.studyTimeMs || 0),
      },
      customDataset: parsed.customDataset || null,
    };
  } catch (error) {
    console.warn("Failed to load local data:", error);
    return {
      settings: { ...DEFAULT_SETTINGS },
      progress: cloneData(DEFAULT_PROGRESS),
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

  if (!MODE_LABELS[state.store.settings.mode]) {
    state.store.settings.mode = DEFAULT_SETTINGS.mode;
  }
  if (!DIRECTION_LABELS[state.store.settings.direction]) {
    state.store.settings.direction = DEFAULT_SETTINGS.direction;
  }
  if (!FOCUS_LABELS[state.store.settings.focus]) {
    state.store.settings.focus = DEFAULT_SETTINGS.focus;
  }
  if (!TRANSITION_LABELS[state.store.settings.autoAdvanceMode]) {
    state.store.settings.autoAdvanceMode = DEFAULT_SETTINGS.autoAdvanceMode;
  }

  state.store.settings.sessionLength = normalizeSessionLength(state.store.settings.sessionLength);
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
  const lookup = new Map();
  for (const dataset of getDatasets()) {
    for (const word of dataset.words) {
      lookup.set(word.uid, word);
    }
  }
  return lookup;
}

function displayDatasetName(meta) {
  if (!meta) {
    return "Default Set";
  }
  if (meta.id === defaultDatasetMeta.id) {
    return "Default Set";
  }
  return String(meta.name || meta.source || "Custom Set").trim() || "Custom Set";
}

function hydrateControls() {
  populateDatasetSelect();
  refreshSessionLengthOptions();
  el.modeSelect.value = state.store.settings.mode;
  el.datasetSelect.value = state.store.settings.activeDataset;
  el.themeSelect.value = state.store.settings.theme;
  el.directionSelect.value = state.store.settings.direction;
  el.focusSelect.value = state.store.settings.focus;
  el.autoAdvanceSelect.value = state.store.settings.autoAdvanceMode;
  el.autoPronounceToggle.checked = Boolean(state.store.settings.autoPronounce);
  el.sessionLengthSelect.value = String(state.store.settings.sessionLength);
  el.rangeStartInput.value = state.store.settings.rangeStart;
  el.rangeEndInput.value = state.store.settings.rangeEnd;
}

function populateDatasetSelect() {
  const datasets = getDatasets();
  el.datasetSelect.replaceChildren();
  datasets.forEach((dataset) => {
    const option = document.createElement("option");
    option.value = dataset.meta.id;
    option.textContent = `${displayDatasetName(dataset.meta)} (${dataset.words.length})`;
    el.datasetSelect.append(option);
  });
}

function refreshSessionLengthOptions() {
  const dataset = getActiveDataset();
  const rangeWords = getWordsInRange(dataset.words);
  const selectedWords = getCandidatesForFocus(rangeWords);
  const count = selectedWords.length;
  const options = [
    { value: "10", label: "10" },
    { value: "20", label: "20" },
    { value: "50", label: "50" },
    { value: "-1", label: `Selected Range (${count})` },
    { value: "0", label: "Unlimited" },
  ];
  const currentValue = String(normalizeSessionLength(state.store.settings.sessionLength));

  el.sessionLengthSelect.replaceChildren();
  options.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    option.selected = item.value === currentValue;
    el.sessionLengthSelect.append(option);
  });
}

function bindEvents() {
  el.tabButtons.forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.viewTarget || "system"));
  });

  el.modeSelect.addEventListener("change", () => updateSetting("mode", el.modeSelect.value));
  el.datasetSelect.addEventListener("change", () => handleDatasetChange());
  el.themeSelect.addEventListener("change", () => {
    updateSetting("theme", el.themeSelect.value);
    applyTheme();
  });
  el.directionSelect.addEventListener("change", () => updateSetting("direction", el.directionSelect.value));
  el.focusSelect.addEventListener("change", () => updateSetting("focus", el.focusSelect.value));
  el.autoAdvanceSelect.addEventListener("change", () => updateSetting("autoAdvanceMode", el.autoAdvanceSelect.value));
  el.autoPronounceToggle.addEventListener("change", () => updateSetting("autoPronounce", el.autoPronounceToggle.checked));
  el.sessionLengthSelect.addEventListener("change", () => updateSetting("sessionLength", Number(el.sessionLengthSelect.value)));
  el.rangeStartInput.addEventListener("input", () => updateRangeSetting("rangeStart", el.rangeStartInput.value));
  el.rangeEndInput.addEventListener("input", () => updateRangeSetting("rangeEnd", el.rangeEndInput.value));

  el.goToDetailsButton.addEventListener("click", () => switchView("details"));
  el.backToSystemButton.addEventListener("click", () => switchView("system"));
  el.startSessionButton.addEventListener("click", () => startSession());
  el.backToDetailsButton.addEventListener("click", () => switchView("details"));
  el.repeatSessionButton.addEventListener("click", () => repeatDisplayedSession());
  el.resultPrevPageButton.addEventListener("click", () => changeResultHistoryPage(-1));
  el.resultNextPageButton.addEventListener("click", () => changeResultHistoryPage(1));
  el.exitSessionButton.addEventListener("click", () => switchView("details"));

  el.nextQuestionButton.addEventListener("click", () => {
    if (state.session.completed || !state.session.active) {
      startSession();
      return;
    }
    nextQuestion();
  });

  el.revealButton.addEventListener("click", () => revealCurrentAnswer());
  el.speakButton.addEventListener("click", () => {
    const question = state.currentQuestion;
    if (question) {
      speakText(question.word.english);
    }
  });
  el.favoriteButton.addEventListener("click", () => toggleFavoriteForCurrentWord());
  el.feedbackBox.addEventListener("click", () => {
    advanceQuestionFromGesture();
  });
  bindAdvanceSwipeSurface(el.answerArea);
  bindAdvanceSwipeSurface(el.feedbackBox);

  el.openWeakWordsButton.addEventListener("click", () => openAuxView("weak-words"));
  el.closeWeakWordsButton.addEventListener("click", () => closeAuxView());
  el.openWordBankButton.addEventListener("click", () => openAuxView("word-bank", "system"));
  el.closeWordBankButton.addEventListener("click", () => closeAuxView());

  el.csvFileInput.addEventListener("change", async () => {
    const file = el.csvFileInput.files?.[0];
    if (!file) {
      return;
    }
    await importWordFile(file);
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

  el.clearBackgroundButton.addEventListener("click", () => clearBackgroundImage());
  el.clearCacheButton.addEventListener("click", () => clearApplicationCache());
  el.resetProgressButton.addEventListener("click", () => resetProgressOnly());
  el.resetAllButton.addEventListener("click", () => resetAllLocalData());
  document.addEventListener("keydown", handleKeyboardShortcuts);
}

function updateSetting(key, value) {
  state.store.settings[key] = value;
  saveStore();
  renderAll();
}

function updateRangeSetting(key, value) {
  const sanitized = String(value).replace(/[^\d]/g, "");
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

function handleDatasetChange() {
  updateSetting("activeDataset", el.datasetSelect.value);
  state.currentQuestion = null;
  state.session = createSessionState();
  state.lastSessionSummary = null;
  clearPendingAutoAdvance();
  renderAll();
}

function switchView(viewName) {
  state.currentView = viewName;
  if (PRIMARY_VIEWS.has(viewName)) {
    state.primaryView = viewName;
    state.returnView = viewName;
  }

  el.views.forEach((view) => {
    view.classList.toggle("is-active", view.id === `view-${viewName}`);
  });

  el.tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.viewTarget === state.primaryView);
  });
  applyActiveViewState();
}

function openAuxView(viewName, returnView = state.primaryView) {
  state.returnView = returnView;
  state.currentView = viewName;
  el.views.forEach((view) => {
    view.classList.toggle("is-active", view.id === `view-${viewName}`);
  });
  el.tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.viewTarget === state.returnView);
  });
  applyActiveViewState();
}

function closeAuxView() {
  switchView(state.returnView || "dashboard");
}

function applyActiveViewState() {
  document.body.dataset.activeView = state.currentView;
}

function normalizeSessionLength(value) {
  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : DEFAULT_SETTINGS.sessionLength;
}

function snapshotSettings(dataset, selectedWords) {
  return {
    activeDataset: state.store.settings.activeDataset,
    datasetName: displayDatasetName(dataset.meta),
    datasetId: dataset.meta.id,
    mode: state.store.settings.mode,
    direction: state.store.settings.direction,
    focus: state.store.settings.focus,
    sessionLength: normalizeSessionLength(state.store.settings.sessionLength),
    rangeStart: state.store.settings.rangeStart,
    rangeEnd: state.store.settings.rangeEnd,
    rangeLabel: formatRangeLabel(dataset, state.store.settings.rangeStart, state.store.settings.rangeEnd),
    autoPronounce: Boolean(state.store.settings.autoPronounce),
    autoAdvanceMode: state.store.settings.autoAdvanceMode,
    poolSize: selectedWords.length,
  };
}

function restoreSettingsFromSnapshot(snapshot) {
  if (!snapshot) {
    return;
  }

  state.store.settings.activeDataset = snapshot.activeDataset || defaultDatasetMeta.id;
  state.store.settings.mode = snapshot.mode || DEFAULT_SETTINGS.mode;
  state.store.settings.direction = snapshot.direction || DEFAULT_SETTINGS.direction;
  state.store.settings.focus = snapshot.focus || DEFAULT_SETTINGS.focus;
  state.store.settings.sessionLength = normalizeSessionLength(snapshot.sessionLength);
  state.store.settings.rangeStart = String(snapshot.rangeStart || "");
  state.store.settings.rangeEnd = String(snapshot.rangeEnd || "");
  state.store.settings.autoPronounce = Boolean(snapshot.autoPronounce);
  state.store.settings.autoAdvanceMode = snapshot.autoAdvanceMode || DEFAULT_SETTINGS.autoAdvanceMode;

  syncStore();
  saveStore();
  hydrateControls();
  renderAll();
}

function startSession() {
  const dataset = getActiveDataset();
  const rangeWords = getWordsInRange(dataset.words);
  const selectedWords = getCandidatesForFocus(rangeWords);
  const poolWords = selectedWords;

  if (!poolWords.length) {
    state.currentQuestion = null;
    state.session = createSessionState();
    setFeedback("No words match the current setup.", "wrong");
    switchView("details");
    renderAll();
    return;
  }

  clearPendingAutoAdvance();
  const settings = snapshotSettings(dataset, poolWords);
  state.resultHistoryPage = 0;
  state.session = {
    ...createSessionState(),
    id: createSessionId(),
    active: true,
    startedAt: Date.now(),
    limit: resolveSessionLimit(settings.sessionLength, poolWords.length),
    poolWords,
    cycleQueue: buildCycleQueue(poolWords),
    settings,
  };
  state.currentQuestion = null;
  state.lastSessionSummary = null;
  switchView("study");
  nextQuestion(true);
}

function createSessionId() {
  return `session-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function resolveSessionLimit(sessionLength, selectedCount) {
  const numericLength = normalizeSessionLength(sessionLength);
  if (numericLength === -1) {
    return selectedCount;
  }
  return numericLength;
}

function buildCycleQueue(words) {
  return shuffle(words);
}

function nextQuestion(isFreshStart = false) {
  clearPendingAutoAdvance();

  if (!state.session.active) {
    startSession();
    return;
  }

  if (!isFreshStart && state.currentQuestion && !state.currentQuestion.answered && !state.session.completed) {
    setFeedback("Answer or reveal first.", "");
    return;
  }

  if (!isFreshStart && state.session.completed) {
    return;
  }

  const question = buildQuestion();
  state.currentQuestion = question;

  if (!question) {
    setFeedback("No question available.", "wrong");
    renderAll();
    return;
  }

  if (state.session.settings?.autoPronounce) {
    speakText(question.word.english);
  }

  setFeedback(getPromptFeedback(question), "");
  renderAll();
}

function getPromptFeedback(question) {
  return question ? "" : "";
}

function buildQuestion() {
  if (!state.session.poolWords.length) {
    if (state.session.active && state.session.answered > 0 && !state.session.completed) {
      completeSession();
    }
    return null;
  }

  if (!state.session.cycleQueue.length) {
    state.session.cycleQueue = buildCycleQueue(state.session.poolWords);
  }

  const word = state.session.cycleQueue.shift();
  if (!word) {
    if (state.session.active && state.session.answered > 0 && !state.session.completed) {
      completeSession();
    }
    return null;
  }

  const settings = state.session.settings;
  const direction = settings.direction === "mixed"
    ? (Math.random() < 0.5 ? "en-to-ja" : "ja-to-en")
    : settings.direction;
  const mode = settings.mode;

  return {
    word,
    datasetName: settings.datasetName,
    direction,
    mode,
    prompt: direction === "en-to-ja" ? word.english : word.japanese,
    answered: false,
    correct: false,
    revealed: false,
    selectedChoiceUid: "",
    options: mode === "multiple" ? buildMultipleChoiceOptions(word, state.session.poolWords, direction) : [],
  };
}

function buildMultipleChoiceOptions(word, pool, direction) {
  const labelFor = (candidate) => (direction === "en-to-ja" ? candidate.japanese : candidate.english);
  const seen = new Set([normalizeBasic(labelFor(word))]);
  const distractors = [];

  shuffle(pool.filter((candidate) => candidate.uid !== word.uid)).forEach((candidate) => {
    if (distractors.length >= 3) {
      return;
    }
    const label = normalizeBasic(labelFor(candidate));
    if (seen.has(label)) {
      return;
    }
    seen.add(label);
    distractors.push(candidate);
  });

  return shuffle([word, ...distractors]).map((candidate, index) => ({
    index: index + 1,
    uid: candidate.uid,
    label: labelFor(candidate),
    isCorrect: candidate.uid === word.uid,
  }));
}

function getCandidatesForFocus(words, focus = state.store.settings.focus) {
  if (focus === "known") {
    return words.filter((word) => isWordMemorized(word.uid));
  }

  const availableWords = words.filter((word) => !isWordMemorized(word.uid));

  if (focus === "favorites") {
    return availableWords.filter((word) => state.store.progress.favorites[word.uid]);
  }

  if (focus === "recent-mistakes") {
    const recentMistakes = new Set(
      state.store.progress.attempts
        .filter((attempt) => !attempt.correct)
        .slice(0, 80)
        .map((attempt) => attempt.uid)
    );
    return availableWords.filter((word) => recentMistakes.has(word.uid));
  }

  if (focus === "weak") {
    const weakIds = new Set(getWeakWords(200).map((entry) => entry.word.uid));
    return availableWords.filter((word) => weakIds.has(word.uid));
  }

  return availableWords;
}

function getSelectedRange(rawStart = state.store.settings.rangeStart, rawEnd = state.store.settings.rangeEnd) {
  const start = rawStart ? Number(rawStart) : null;
  const end = rawEnd ? Number(rawEnd) : null;

  if (start !== null && end !== null && start > end) {
    return { start: end, end: start };
  }
  return { start, end };
}

function getWordsInRange(words, rawStart = state.store.settings.rangeStart, rawEnd = state.store.settings.rangeEnd) {
  const { start, end } = getSelectedRange(rawStart, rawEnd);
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

function getDatasetRange(dataset = getActiveDataset()) {
  return {
    min: dataset.words[0]?.sectionNumber ?? 1,
    max: dataset.words[dataset.words.length - 1]?.sectionNumber ?? dataset.words.length,
  };
}

function formatRangeLabel(dataset = getActiveDataset(), rawStart = state.store.settings.rangeStart, rawEnd = state.store.settings.rangeEnd) {
  const { min, max } = getDatasetRange(dataset);
  const { start, end } = getSelectedRange(rawStart, rawEnd);
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

function getAutoAdvanceMode(settings = state.store.settings) {
  return TRANSITION_LABELS[settings.autoAdvanceMode] ? settings.autoAdvanceMode : DEFAULT_SETTINGS.autoAdvanceMode;
}

function shouldAutoAdvanceAfterAnswer(correct) {
  if (state.session.completed) {
    return false;
  }
  const mode = getAutoAdvanceMode(state.session.settings || state.store.settings);
  return mode === "always" || (mode === "correct" && correct);
}

function focusNextQuestionButton() {
  if (!state.session.active || state.session.completed || !state.currentQuestion?.answered) {
    return;
  }
  requestAnimationFrame(() => el.nextQuestionButton?.focus());
}

function canAdvanceFromGesture() {
  return (
    state.currentView === "study" &&
    state.session.active &&
    !state.session.completed &&
    Boolean(state.currentQuestion?.answered)
  );
}

function advanceQuestionFromGesture() {
  if (!canAdvanceFromGesture()) {
    return false;
  }
  nextQuestion();
  return true;
}

function bindAdvanceSwipeSurface(element) {
  let startX = 0;
  let startY = 0;

  element.addEventListener("touchstart", (event) => {
    const touch = event.changedTouches[0];
    if (!touch) {
      return;
    }
    startX = touch.clientX;
    startY = touch.clientY;
  }, { passive: true });

  element.addEventListener("touchend", (event) => {
    const touch = event.changedTouches[0];
    if (!touch) {
      return;
    }
    const deltaX = touch.clientX - startX;
    const deltaY = touch.clientY - startY;
    if (Math.abs(deltaX) >= ADVANCE_SWIPE_THRESHOLD && Math.abs(deltaX) > Math.abs(deltaY) * 1.2) {
      advanceQuestionFromGesture();
    }
  });
}

function handlePostAnswer(correct) {
  renderAll();
  if (state.session.completed) {
    return;
  }
  if (shouldAutoAdvanceAfterAnswer(correct)) {
    state.pendingAutoAdvance = window.setTimeout(() => nextQuestion(), AUTO_ADVANCE_DELAY);
    return;
  }
  focusNextQuestionButton();
}

function revealCurrentAnswer() {
  const question = state.currentQuestion;
  if (!question) {
    return;
  }

  if (question.mode === "card") {
    toggleWordCardReveal();
    return;
  }

  if (question.mode === "flashcard") {
    question.revealed = !question.revealed;
    renderQuestion();
    setFeedback(question.revealed ? "Mark the card when ready." : "Hidden.", "");
    return;
  }

  setFeedback(`Answer: ${formatExpectedAnswer(question.word, question.direction)}`, "");
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
    correct
      ? `${question.word.english} = ${question.word.japanese}`
      : `Correct: ${question.word.english} = ${question.word.japanese}${getSessionCompletionSuffix()}`,
    correct ? "correct" : "wrong"
  );
  handlePostAnswer(correct);
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
    result.correct
      ? `${question.word.english} = ${question.word.japanese}`
      : `Answer: ${result.accepted.join(" / ") || formatExpectedAnswer(question.word, question.direction)}${getSessionCompletionSuffix()}`,
    result.correct ? "correct" : "wrong"
  );
  handlePostAnswer(result.correct);
}

function submitFlashcardResult(correct) {
  const question = state.currentQuestion;
  if (!question || question.answered || !question.revealed) {
    return;
  }

  question.answered = true;
  question.correct = correct;
  setWordMemorized(question.word.uid, correct);

  recordAttempt({
    correct,
    direction: question.direction,
    mode: question.mode,
    question,
    userAnswer: correct ? "Known" : "Again",
  });

  if (!state.session.completed && state.session.active && !state.session.poolWords.length) {
    completeSession();
  }

  setFeedback(
    correct ? `Saved: ${question.word.english}` : `Review: ${question.word.english}${getSessionCompletionSuffix()}`,
    correct ? "correct" : "wrong"
  );
  handlePostAnswer(correct);
}

function recordAttempt({ correct, direction, mode, question, userAnswer, selectedChoiceUid = "" }) {
  const at = Date.now();
  const expectedAnswer = formatExpectedAnswer(question.word, direction);
  const attempt = {
    sessionId: state.session.id,
    uid: question.word.uid,
    datasetId: state.session.settings?.datasetId || state.store.settings.activeDataset,
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
  state.session.attempts.push(attempt);
  state.store.progress.bestStreak = Math.max(state.store.progress.bestStreak || 0, state.session.streak);

  if (state.session.limit > 0 && state.session.answered >= state.session.limit) {
    completeSession();
  }

  saveStore();
}

function completeSession() {
  if (state.session.completed) {
    return;
  }

  state.session.completed = true;
  state.session.active = false;
  state.session.completedAt = Date.now();
  state.session.durationMs = Math.max(state.session.completedAt - state.session.startedAt, 0);
  state.store.progress.studyTimeMs += state.session.durationMs;

  const summary = summarizeSession(state.session);
  state.lastSessionSummary = summary;
  state.store.progress.sessions.unshift({
    id: summary.id,
    startedAt: summary.startedAt,
    completedAt: summary.completedAt,
    durationMs: summary.durationMs,
    answered: summary.answered,
    correct: summary.correct,
    accuracy: summary.accuracy,
    bestStreak: summary.bestStreak,
    datasetName: summary.datasetName,
    datasetId: summary.datasetId,
    rangeLabel: summary.rangeLabel,
    settingsSnapshot: summary.settingsSnapshot,
  });
  state.store.progress.sessions = state.store.progress.sessions.slice(0, SESSION_LOG_LIMIT);
  state.currentQuestion = null;
  clearPendingAutoAdvance();
  saveStore();
  switchView("results");
}

function summarizeSession(session) {
  const attempts = [...session.attempts];
  const mistakes = attempts.filter((attempt) => !attempt.correct);
  const settingsSnapshot = cloneData(session.settings || DEFAULT_SETTINGS);

  return {
    id: session.id,
    startedAt: session.startedAt,
    completedAt: session.completedAt,
    durationMs: session.durationMs,
    answered: session.answered,
    correct: session.correct,
    accuracy: session.answered ? session.correct / session.answered : 0,
    bestStreak: session.bestStreak,
    attempts,
    mistakes,
    datasetName: session.settings?.datasetName || "Default Set",
    datasetId: session.settings?.datasetId || defaultDatasetMeta.id,
    rangeLabel: session.settings?.rangeLabel || formatRangeLabel(),
    settingsSnapshot,
  };
}

function getDisplayedSessionSummary() {
  if (state.lastSessionSummary) {
    return state.lastSessionSummary;
  }

  const latestSession = state.store.progress.sessions[0];
  if (!latestSession) {
    return null;
  }

  const attempts = state.store.progress.attempts
    .filter((attempt) => attempt.sessionId === latestSession.id)
    .sort((left, right) => left.at - right.at);

  return {
    id: latestSession.id,
    startedAt: latestSession.startedAt,
    completedAt: latestSession.completedAt,
    durationMs: latestSession.durationMs,
    answered: latestSession.answered,
    correct: latestSession.correct,
    accuracy: latestSession.accuracy,
    bestStreak: latestSession.bestStreak,
    attempts,
    mistakes: attempts.filter((attempt) => !attempt.correct),
    datasetName: latestSession.datasetName,
    datasetId: latestSession.datasetId,
    rangeLabel: latestSession.rangeLabel,
    settingsSnapshot: latestSession.settingsSnapshot || null,
  };
}

function repeatDisplayedSession() {
  const summary = getDisplayedSessionSummary();
  if (summary?.settingsSnapshot) {
    restoreSettingsFromSnapshot(summary.settingsSnapshot);
  }
  startSession();
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
  const base = String(text).replace(/[()]/g, " ");
  const segments = [base, ...base.split(/[;,/]/g)];
  return [...new Set(segments.map((segment) => segment.trim()).filter(Boolean))];
}

function getJapaneseAnswerCandidates(text) {
  const base = String(text)
    .replace(/[()（）]/g, " ")
    .replace(/・/g, " ")
    .trim();
  const segments = [base, ...base.split(/[;,/、，]/g)];
  return [...new Set(segments.map((segment) => segment.trim()).filter(Boolean))];
}

function normalizeEnglish(text) {
  return String(text)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/['’`]/g, "")
    .replace(/[._!?-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeJapanese(text) {
  return String(text)
    .normalize("NFKC")
    .replace(/[ 　\t\r\n]/g, "")
    .replace(/[・,.;/()（）「」『』\[\]]/g, "")
    .trim();
}

function normalizeBasic(text) {
  return String(text)
    .normalize("NFKC")
    .replace(/[ 　\t\r\n]/g, "")
    .trim();
}

function formatExpectedAnswer(word, direction) {
  return direction === "en-to-ja" ? word.japanese : word.english;
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
  renderWordBank();
}

function isWordMemorized(wordUid) {
  return Boolean(state.store.progress.memorized?.[wordUid]);
}

function setWordMemorized(wordUid, memorized) {
  if (!state.store.progress.memorized) {
    state.store.progress.memorized = {};
  }

  if (memorized) {
    state.store.progress.memorized[wordUid] = true;
    if (state.session.active) {
      state.session.poolWords = state.session.poolWords.filter((word) => word.uid !== wordUid);
      state.session.cycleQueue = state.session.cycleQueue.filter((word) => word.uid !== wordUid);
    }
    return;
  }

  delete state.store.progress.memorized[wordUid];
}

function renderAll() {
  renderHeader();
  renderSystem();
  renderDetails();
  renderStudy();
  renderResults();
  renderDashboard();
  renderWeakWords();
  renderWordBank();
}

function renderHeader() {
  const dataset = getActiveDataset();
  el.headerDatasetName.textContent = displayDatasetName(dataset.meta);
  el.headerWordCount.textContent = `${dataset.words.length} words`;
}

function renderSystem() {
  const dataset = getActiveDataset();
  const rangeWords = getWordsInRange(dataset.words);
  const previewWords = getCandidatesForFocus(rangeWords);
  const wordCount = previewWords.length;

  el.systemWordCount.textContent = `${wordCount} words`;
  el.systemDatasetName.textContent = displayDatasetName(dataset.meta);
  el.systemRangePreview.textContent = formatRangeLabel(dataset);
  el.datasetStatusText.textContent = state.store.customDataset?.words?.length
    ? `Loaded: ${displayDatasetName(dataset.meta)} / ${dataset.words.length} words`
    : "Built-in data: system.xlsx";
}

function renderDetails() {
  const dataset = getActiveDataset();
  const rangeWords = getWordsInRange(dataset.words);
  const selectedWords = getCandidatesForFocus(rangeWords);
  const count = selectedWords.length;

  refreshSessionLengthOptions();
  el.detailsSelectionCount.textContent = `${count}`;
  el.detailsRangeHint.textContent = formatRangeLabel(dataset);
  el.detailsModeText.textContent = MODE_LABELS[state.store.settings.mode];
  el.detailsDirectionText.textContent = DIRECTION_LABELS[state.store.settings.direction];
  el.detailsFocusText.textContent = FOCUS_LABELS[state.store.settings.focus];
  el.detailsTransitionText.textContent = TRANSITION_LABELS[getAutoAdvanceMode()];
}

function renderStudy() {
  renderQuestion();
  renderStudyCorner();
}

function renderQuestion() {
  const question = state.currentQuestion;

  if (!question) {
    el.questionBadge.textContent = state.session.completed ? "Complete" : "Ready";
    el.promptText.textContent = state.session.completed
      ? "Session complete. Open Results for the score."
      : "Set your session and press Start.";
    el.favoriteButton.textContent = "☆";
    el.favoriteButton.disabled = true;
    el.answerArea.replaceChildren(makeEmptyState("The active question appears here."));
    el.feedbackBox.classList.remove("is-correct", "is-wrong", "is-next-ready");
    el.feedbackBox.textContent = "";
    el.revealButton.disabled = true;
    el.speakButton.disabled = true;
    el.nextQuestionButton.textContent = state.session.active ? "Next" : "Start";
    return;
  }

  const nextIndex = state.session.answered + 1;
  const totalLabel = state.session.limit > 0 ? ` / ${state.session.limit}` : "";
  el.questionBadge.textContent = `Q ${nextIndex}${totalLabel}`;
  el.promptText.textContent = question.mode === "card" ? "Word Card" : question.prompt;
  el.favoriteButton.textContent = state.store.progress.favorites[question.word.uid] ? "★" : "☆";
  el.favoriteButton.disabled = false;
  el.answerArea.replaceChildren();
  el.feedbackBox.classList.toggle("is-next-ready", question.answered && !state.session.completed);
  el.revealButton.disabled = false;
  el.speakButton.disabled = false;
  el.nextQuestionButton.textContent = state.session.completed ? "Restart" : "Next";

  if (question.mode === "multiple") {
    renderMultipleChoice(question);
    el.revealButton.textContent = "Reveal";
  } else if (question.mode === "input") {
    renderInputMode(question);
    el.revealButton.textContent = "Reveal";
  } else if (question.mode === "card") {
    renderWordCardMode(question);
    el.revealButton.textContent = question.revealed ? "Flip Back" : "Flip Card";
  } else {
    renderFlashcardMode(question);
    el.revealButton.textContent = question.revealed ? "Hide" : "Reveal";
  }
}

function renderStudyCorner() {
  const activeDataset = getActiveDataset();
  const datasetWords = getWordsInRange(activeDataset.words);
  const settings = state.session.settings || snapshotSettings(activeDataset, datasetWords);
  const wordId = state.currentQuestion ? `No.${state.currentQuestion.word.id}` : "No.-";
  const accuracy = state.session.answered ? formatPercent(state.session.correct / state.session.answered) : "-";
  const progress = state.session.limit > 0 ? `${state.session.answered} / ${state.session.limit}` : `${state.session.answered} / -`;

  el.studyProgressText.textContent = progress;
  el.studyAccuracyText.textContent = accuracy;
  el.studyStreakText.textContent = String(state.session.streak);
  el.studyStatusCorner.replaceChildren();

  [
    MODE_LABELS[settings.mode] || MODE_LABELS[DEFAULT_SETTINGS.mode],
    DIRECTION_LABELS[settings.direction] || DIRECTION_LABELS[DEFAULT_SETTINGS.direction],
    `${settings.datasetName || "Default Set"} / ${settings.rangeLabel || formatRangeLabel()}`,
    wordId,
  ].forEach((text) => {
    const line = document.createElement("span");
    line.className = "corner-meta";
    line.textContent = text;
    el.studyStatusCorner.append(line);
  });
}

function renderMultipleChoice(question) {
  const grid = document.createElement("div");
  grid.className = "option-grid";

  question.options.forEach((option) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "option-button";
    button.textContent = `${option.index}. ${option.label}`;

    if (question.answered) {
      button.classList.add("is-next-ready");
      if (option.isCorrect) {
        button.classList.add("is-correct");
      }
      if (option.uid === question.selectedChoiceUid && !option.isCorrect) {
        button.classList.add("is-wrong");
      }
    }

    button.addEventListener("click", () => {
      if (question.answered) {
        advanceQuestionFromGesture();
        return;
      }
      submitMultipleChoice(option.uid, option.label);
    });
    grid.append(button);
  });

  el.answerArea.append(grid);
}

function renderInputMode(question) {
  const block = document.createElement("div");
  block.className = "input-block";

  const row = document.createElement("div");
  row.className = "input-row";

  const input = document.createElement("input");
  input.className = "answer-input";
  input.type = "text";
  input.placeholder = "Type your answer";
  input.autocomplete = "off";
  input.spellcheck = false;
  input.disabled = question.answered;

  const submit = document.createElement("button");
  submit.type = "button";
  submit.className = "primary-button";
  submit.textContent = question.answered ? "Answered" : "Submit";
  submit.disabled = question.answered;

  submit.addEventListener("click", () => submitTypedAnswer(input.value));
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
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
    known.textContent = "Known";
    known.disabled = question.answered;
    known.addEventListener("click", () => submitFlashcardResult(true));

    const again = document.createElement("button");
    again.type = "button";
    again.className = "secondary-button";
    again.textContent = "Again";
    again.disabled = question.answered;
    again.addEventListener("click", () => submitFlashcardResult(false));

    row.append(known, again);
    block.append(row);
  }

  el.answerArea.append(block);
}

function renderWordCardMode(question) {
  const faces = question.direction === "ja-to-en"
    ? { front: question.word.japanese, back: question.word.english }
    : { front: question.word.english, back: question.word.japanese };

  const block = document.createElement("div");
  block.className = "card-block";

  const card = document.createElement("button");
  card.type = "button";
  card.className = `word-card${question.revealed ? " is-flipped" : ""}`;
  card.setAttribute("aria-label", "Flip word card");
  bindWordCardGestures(card);
  card.addEventListener("click", () => {
    if (card.dataset.ignoreClick === "true") {
      card.dataset.ignoreClick = "false";
      return;
    }
    if (question.answered) {
      advanceQuestionFromGesture();
      return;
    }
    toggleWordCardReveal();
  });

  const inner = document.createElement("span");
  inner.className = "word-card-inner";

  const front = document.createElement("span");
  front.className = "word-card-face word-card-front";
  const frontTag = document.createElement("span");
  frontTag.className = "word-card-tag";
  frontTag.textContent = "FRONT";
  const frontText = document.createElement("strong");
  frontText.textContent = faces.front;
  front.append(frontTag, frontText);

  const back = document.createElement("span");
  back.className = "word-card-face word-card-back";
  const backTag = document.createElement("span");
  backTag.className = "word-card-tag";
  backTag.textContent = "BACK";
  const backText = document.createElement("strong");
  backText.textContent = faces.back;
  back.append(backTag, backText);

  inner.append(front, back);
  card.append(inner);
  block.append(card);

  const row = document.createElement("div");
  row.className = `button-row card-actions${question.revealed ? " is-visible" : ""}`;

  const known = document.createElement("button");
  known.type = "button";
  known.className = "primary-button";
  known.textContent = "Known";
  known.disabled = question.answered;
  known.addEventListener("click", () => submitFlashcardResult(true));

  const again = document.createElement("button");
  again.type = "button";
  again.className = "secondary-button";
  again.textContent = "Again";
  again.disabled = question.answered;
  again.addEventListener("click", () => submitFlashcardResult(false));

  row.append(known, again);
  block.append(row);

  el.answerArea.append(block);
}

function bindWordCardGestures(card) {
  let startX = 0;
  let startY = 0;

  card.addEventListener("touchstart", (event) => {
    const touch = event.changedTouches[0];
    if (!touch) {
      return;
    }
    startX = touch.clientX;
    startY = touch.clientY;
  }, { passive: true });

  card.addEventListener("touchend", (event) => {
    const touch = event.changedTouches[0];
    if (!touch) {
      return;
    }
    const deltaX = touch.clientX - startX;
    const deltaY = touch.clientY - startY;
    if (Math.abs(deltaX) >= CARD_SWIPE_THRESHOLD && Math.abs(deltaX) > Math.abs(deltaY) * 1.2) {
      card.dataset.ignoreClick = "true";
      handleWordCardSwipe();
    }
  });
}

function toggleWordCardReveal(force) {
  const question = state.currentQuestion;
  if (!question || question.mode !== "card") {
    return;
  }

  question.revealed = typeof force === "boolean" ? force : !question.revealed;
  const card = el.answerArea.querySelector(".word-card");
  const actions = el.answerArea.querySelector(".card-actions");
  card?.classList.toggle("is-flipped", question.revealed);
  actions?.classList.toggle("is-visible", question.revealed);
  el.revealButton.textContent = question.revealed ? "Flip Back" : "Flip Card";
}

function handleWordCardSwipe() {
  const question = state.currentQuestion;
  if (!question || question.mode !== "card" || !question.answered) {
    return;
  }
  advanceQuestionFromGesture();
}

function renderResults() {
  const summary = getDisplayedSessionSummary();
  if (!summary) {
    el.resultDatasetChip.textContent = displayDatasetName(getActiveDataset().meta);
    el.resultRangeChip.textContent = formatRangeLabel();
    el.resultSummaryText.textContent = "Finish a session to see the score.";
    el.resultCorrectCount.textContent = "0";
    el.resultAnsweredCount.textContent = "0 answered";
    el.resultAccuracyValue.textContent = "0%";
    el.resultBestStreak.textContent = "0";
    el.resultDuration.textContent = "0m";
    el.resultHistoryCount.textContent = "0";
    updateResultHistoryPager(0);
    renderResultHistoryList([]);
    return;
  }

  const totalPages = Math.max(1, Math.ceil(summary.attempts.length / RESULTS_PAGE_SIZE));
  state.resultHistoryPage = Math.min(state.resultHistoryPage, totalPages - 1);

  el.resultDatasetChip.textContent = summary.datasetName;
  el.resultRangeChip.textContent = summary.rangeLabel;
  el.resultSummaryText.textContent = `${summary.correct} / ${summary.answered} correct`;
  el.resultCorrectCount.textContent = String(summary.correct);
  el.resultAnsweredCount.textContent = `${summary.answered} answered`;
  el.resultAccuracyValue.textContent = formatPercent(summary.accuracy);
  el.resultBestStreak.textContent = String(summary.bestStreak);
  el.resultDuration.textContent = formatDuration(summary.durationMs);
  el.resultHistoryCount.textContent = `${summary.attempts.length}`;
  updateResultHistoryPager(totalPages);
  renderResultHistoryList(summary.attempts);
}

function renderResultHistoryList(attempts) {
  el.resultHistoryList.replaceChildren();
  if (!attempts.length) {
    el.resultHistoryList.append(makeEmptyState("No session history yet.", "li"));
    return;
  }

  const start = state.resultHistoryPage * RESULTS_PAGE_SIZE;
  const pageAttempts = attempts.slice(start, start + RESULTS_PAGE_SIZE);

  pageAttempts.forEach((attempt, index) => {
    const item = document.createElement("li");
    item.className = `result-history-item${attempt.correct ? " is-correct" : " is-wrong"}`;

    const top = document.createElement("div");
    top.className = "result-history-top";

    const order = document.createElement("span");
    order.className = "result-order";
    order.textContent = `#${start + index + 1}`;

    const badge = document.createElement("span");
    badge.className = "result-badge";
    badge.textContent = attempt.correct ? "OK" : "MISS";

    top.append(order, badge);

    const english = document.createElement("strong");
    english.textContent = attempt.english;

    const japanese = document.createElement("span");
    japanese.className = "result-meaning";
    japanese.textContent = attempt.japanese;

    item.append(top, english, japanese);

    if (!attempt.correct && shouldShowAttemptDetail(attempt)) {
      const detail = document.createElement("span");
      detail.className = "result-detail";
      detail.textContent = `Your: ${attempt.userAnswer || "-"} / Answer: ${attempt.expectedAnswer}`;
      item.append(detail);
    }

    el.resultHistoryList.append(item);
  });
}

function changeResultHistoryPage(delta) {
  const summary = getDisplayedSessionSummary();
  if (!summary) {
    return;
  }
  const totalPages = Math.max(1, Math.ceil(summary.attempts.length / RESULTS_PAGE_SIZE));
  state.resultHistoryPage = Math.max(0, Math.min(totalPages - 1, state.resultHistoryPage + delta));
  renderResults();
}

function updateResultHistoryPager(totalPages) {
  if (totalPages <= 0) {
    el.resultPageLabel.textContent = "0 / 0";
    el.resultPrevPageButton.disabled = true;
    el.resultNextPageButton.disabled = true;
    return;
  }
  el.resultPageLabel.textContent = `${state.resultHistoryPage + 1} / ${totalPages}`;
  el.resultPrevPageButton.disabled = state.resultHistoryPage <= 0;
  el.resultNextPageButton.disabled = state.resultHistoryPage >= totalPages - 1;
}

function shouldShowAttemptDetail(attempt) {
  return attempt.mode === "input" || attempt.mode === "multiple";
}

function renderDashboard() {
  const analytics = computeAnalytics();
  state.latestAnalytics = analytics;

  el.metricStudyTime.textContent = formatDuration(analytics.studyTimeMs);
  el.metricAttempts.textContent = String(analytics.totalAttempts);
  el.metricBestStreak.textContent = String(state.store.progress.bestStreak || 0);
  el.weakCandidateCount.textContent = String(analytics.weakWords.length);

  renderCalendar(analytics.calendarDays);
  renderWeakPreview(analytics.weakWords.slice(0, 4));
}

function renderCalendar(days) {
  el.dashboardCalendar.replaceChildren();
  if (!days.length) {
    el.dashboardCalendar.append(makeEmptyState("No study data."));
    return;
  }

  days.forEach((day) => {
    const cell = document.createElement("div");
    cell.className = `calendar-cell level-${day.level}${day.isToday ? " is-today" : ""}`;
    cell.title = `${day.label}: ${day.attempts} answers / ${formatDuration(day.studyTimeMs)}`;

    const label = document.createElement("span");
    label.className = "calendar-label";
    label.textContent = String(day.day);

    const bar = document.createElement("span");
    bar.className = "calendar-bar";

    cell.append(label, bar);
    el.dashboardCalendar.append(cell);
  });
}

function renderWeakPreview(items) {
  el.weakCandidatePreview.replaceChildren();
  if (!items.length) {
    const item = document.createElement("li");
    item.className = "preview-chip empty-chip";
    item.textContent = "No weak words";
    el.weakCandidatePreview.append(item);
    return;
  }

  items.forEach((entry) => {
    const item = document.createElement("li");
    item.className = "preview-chip";
    item.textContent = entry.word.english;
    el.weakCandidatePreview.append(item);
  });
}

function renderWeakWords() {
  const analytics = state.latestAnalytics || computeAnalytics();
  const items = analytics.weakWords;
  el.weakWordsDetailList.replaceChildren();

  if (!items.length) {
    el.weakWordsDetailList.append(makeEmptyState("No weak words yet.", "li"));
    return;
  }

  items.forEach((entry) => {
    const row = document.createElement("li");

    const top = document.createElement("div");
    top.className = "list-topline";

    const title = document.createElement("strong");
    title.textContent = entry.word.english;

    const badge = document.createElement("span");
    badge.className = "status-pill";
    badge.textContent = formatPercent(entry.stats.asked ? entry.stats.correct / entry.stats.asked : 0);

    top.append(title, badge);

    const meaning = document.createElement("span");
    meaning.textContent = entry.word.japanese;

    const meta = document.createElement("span");
    meta.className = "list-meta";
    meta.textContent = `${entry.stats.incorrect} miss / ${entry.stats.asked} asked`;

    row.append(top, meaning, meta);
    el.weakWordsDetailList.append(row);
  });
}

function renderWordBank() {
  const dataset = getActiveDataset();
  el.wordBankList.replaceChildren();

  if (!dataset.words.length) {
    el.wordBankList.append(makeEmptyState("No words found.", "li"));
    return;
  }

  dataset.words.forEach((word) => {
    const row = document.createElement("li");
    row.className = "word-bank-item";

    const index = document.createElement("span");
    index.className = "word-bank-index";
    index.textContent = `No.${word.id}`;

    const copy = document.createElement("div");
    copy.className = "word-bank-copy";

    const english = document.createElement("strong");
    english.textContent = word.english;

    const japanese = document.createElement("span");
    japanese.textContent = word.japanese;

    copy.append(english, japanese);

    const meta = document.createElement("div");
    meta.className = "word-bank-meta";

    if (isWordMemorized(word.uid)) {
      const known = document.createElement("span");
      known.className = "status-pill";
      known.textContent = "Known";
      meta.append(known);
    }

    if (state.store.progress.favorites[word.uid]) {
      const star = document.createElement("span");
      star.className = "status-pill";
      star.textContent = "★";
      meta.append(star);
    }

    const stats = state.store.progress.wordStats[word.uid];
    if (stats?.incorrect) {
      const weak = document.createElement("span");
      weak.className = "status-pill";
      weak.textContent = `${stats.incorrect} miss`;
      meta.append(weak);
    }

    row.append(index, copy, meta);
    el.wordBankList.append(row);
  });
}

function computeAnalytics() {
  return {
    totalAttempts: state.store.progress.attempts.length,
    studyTimeMs: Number(state.store.progress.studyTimeMs || 0),
    weakWords: getWeakWords(50),
    calendarDays: buildCalendarDays(CALENDAR_DAY_COUNT),
  };
}

function buildCalendarDays(dayCount) {
  const attemptsByDay = new Map();
  const studyTimeByDay = new Map();

  state.store.progress.attempts.forEach((attempt) => {
    const key = getDayKey(attempt.at);
    attemptsByDay.set(key, (attemptsByDay.get(key) || 0) + 1);
  });

  state.store.progress.sessions.forEach((session) => {
    const key = getDayKey(session.completedAt || session.startedAt);
    studyTimeByDay.set(key, (studyTimeByDay.get(key) || 0) + Number(session.durationMs || 0));
  });

  const days = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let offset = dayCount - 1; offset >= 0; offset -= 1) {
    const day = new Date(today);
    day.setDate(today.getDate() - offset);
    const key = getDayKey(day.getTime());
    const attempts = attemptsByDay.get(key) || 0;
    const studyTimeMs = studyTimeByDay.get(key) || 0;
    days.push({
      key,
      day: day.getDate(),
      label: `${day.getMonth() + 1}/${day.getDate()}`,
      attempts,
      studyTimeMs,
      level: getCalendarLevel(attempts),
      isToday: offset === 0,
    });
  }

  return days;
}

function getCalendarLevel(attempts) {
  if (!attempts) {
    return 0;
  }
  if (attempts < 10) {
    return 1;
  }
  if (attempts < 25) {
    return 2;
  }
  if (attempts < 45) {
    return 3;
  }
  return 4;
}

function getDayKey(value) {
  const date = new Date(value);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
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

function importWordFile(file) {
  const fileName = String(file.name || "");
  const importer = isSpreadsheetFile(fileName)
    ? readFileAsArrayBuffer(file).then((buffer) => parseWordWorkbook(buffer))
    : file.text().then((text) => parseWordCsv(text));

  return importer.then((parsedWords) => {
    if (!parsedWords.length) {
      alert("No word rows were found in the selected file.");
      return;
    }

    const datasetId = `custom-${Date.now()}`;
    state.store.customDataset = {
      meta: {
        id: datasetId,
        name: file.name.replace(/\.[^.]+$/, "") || "Custom Set",
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
    state.currentQuestion = null;
    state.session = createSessionState();
    state.lastSessionSummary = null;
    syncStore();
    saveStore();
    hydrateControls();
    renderAll();
    switchView("system");
  }).catch((error) => {
    console.error(error);
    alert("File import failed.");
  });
}

function isSpreadsheetFile(fileName) {
  return /\.(xlsx|xls)$/i.test(fileName);
}

function parseWordCsv(text) {
  const sanitized = text.replace(/^\uFEFF/, "");
  const rows = parseDelimitedText(sanitized).filter((row) => row.some((cell) => cell.trim()));
  return parseWordRows(rows);
}

function parseWordWorkbook(buffer) {
  if (typeof XLSX === "undefined") {
    throw new Error("XLSX library is not loaded.");
  }

  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    return [];
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  }).map((row) => row.map((cell) => String(cell ?? "")));

  return parseWordRows(rows);
}

function parseWordRows(rows) {
  if (!rows.length) {
    return [];
  }

  const header = rows[0].map((cell) => cell.trim().toLowerCase());
  const hasHeader = header.some((cell) => /(english|japanese|meaning|word|id|number)/.test(cell));

  const pickIndex = (patterns, fallback) => {
    const index = header.findIndex((cell) => patterns.some((pattern) => pattern.test(cell)));
    return index >= 0 ? index : fallback;
  };

  const englishIndex = hasHeader ? pickIndex([/english/, /word/], rows[0].length >= 3 ? 1 : 0) : rows[0].length >= 3 ? 1 : 0;
  const japaneseIndex = hasHeader ? pickIndex([/japanese/, /meaning/], rows[0].length >= 3 ? 2 : 1) : rows[0].length >= 3 ? 2 : 1;
  const idIndex = hasHeader ? pickIndex([/^id$/, /number/], -1) : rows[0].length >= 3 ? 0 : -1;
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

function applyBackgroundFromFile(file) {
  return readFileAsDataUrl(file).then((dataUrl) => {
    state.temporaryBackground = dataUrl;
    if (dataUrl.length <= PERSISTABLE_BACKGROUND_LIMIT) {
      state.store.settings.customBackground = dataUrl;
    } else {
      state.store.settings.customBackground = "";
    }
    saveStore();
    applyTheme();
    renderAll();
  });
}

function clearBackgroundImage() {
  state.temporaryBackground = "";
  state.store.settings.customBackground = "";
  saveStore();
  applyTheme();
  renderAll();
}

function clearApplicationCache() {
  if (!("caches" in window)) {
    setFeedback("Cache API is not available here.", "wrong");
    return;
  }

  caches.keys()
    .then((keys) => Promise.all(keys.map((key) => caches.delete(key))))
    .then(async () => {
      if ("serviceWorker" in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.all(registrations.map((registration) => registration.update().catch(() => {})));
      }
      setFeedback("Cache cleared.", "correct");
    })
    .catch((error) => {
      console.warn("Failed to clear cache:", error);
      setFeedback("Cache clear failed.", "wrong");
    });
}

function resetProgressOnly() {
  if (!window.confirm("Clear all saved progress on this device?")) {
    return;
  }

  state.store.progress = cloneData(DEFAULT_PROGRESS);
  state.session = createSessionState();
  state.currentQuestion = null;
  state.lastSessionSummary = null;
  clearPendingAutoAdvance();
  saveStore();
  renderAll();
  switchView("system");
}

function resetAllLocalData() {
  if (!window.confirm("Clear settings, progress, custom CSV data, and local background image?")) {
    return;
  }

  state.store = {
    settings: { ...DEFAULT_SETTINGS },
    progress: cloneData(DEFAULT_PROGRESS),
    customDataset: null,
  };
  state.temporaryBackground = "";
  state.session = createSessionState();
  state.currentQuestion = null;
  state.lastSessionSummary = null;
  clearPendingAutoAdvance();
  syncStore();
  saveStore();
  hydrateControls();
  applyTheme();
  renderAll();
  switchView("system");
}

function applyTheme() {
  document.body.dataset.theme = state.store.settings.theme;
  const backgroundData = state.temporaryBackground || state.store.settings.customBackground;
  const backgroundImage = backgroundData
    ? `linear-gradient(135deg, rgba(7, 17, 31, 0.38), rgba(7, 17, 31, 0.68)), url("${backgroundData}") center / cover no-repeat fixed`
    : "none";
  document.body.style.setProperty("--background-image", backgroundImage);
  el.backgroundHint.textContent = backgroundData ? "Background active" : "Theme only";
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

function formatDuration(milliseconds) {
  const totalMinutes = Math.max(0, Math.round(Number(milliseconds || 0) / 60000));
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  if (hours > 0) {
    return `${hours}h ${minutes}m`;
  }
  return `${minutes}m`;
}

function getSessionCompletionSuffix() {
  if (!state.session.completed) {
    return "";
  }
  return `\n${state.session.correct} / ${state.session.answered}`;
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

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function speakText(text) {
  if (!("speechSynthesis" in window)) {
    setFeedback("Speech is not available in this browser.", "wrong");
    return;
  }
  window.speechSynthesis.cancel();
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = "en-US";
  utterance.rate = 0.92;
  window.speechSynthesis.speak(utterance);
}

function handleKeyboardShortcuts(event) {
  if (state.currentView !== "study" || !state.currentQuestion) {
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

  if ((state.currentQuestion.mode === "flashcard" || state.currentQuestion.mode === "card") && event.code === "Space") {
    event.preventDefault();
    revealCurrentAnswer();
    return;
  }

  if (event.key === "Enter") {
    if (!state.currentQuestion.answered) {
      return;
    }
    event.preventDefault();
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

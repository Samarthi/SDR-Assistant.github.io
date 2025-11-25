const companyEl = document.getElementById('company');
const locationEl = document.getElementById('location');
const productEl = document.getElementById('product');
const fileInput = document.getElementById('fileInput');
const saveKeyBtn = document.getElementById('saveKey');
const newResearchBtn = document.getElementById('newResearch');
const newTargetResearchBtn = document.getElementById('newTargetResearch');
const apiKeyInput = document.getElementById('apiKey');
const generateBtn = document.getElementById('generate');
const status = document.getElementById('status');
const resultDiv = document.getElementById('result');
const briefDiv = document.getElementById('brief');
const briefSummarySection = document.getElementById('briefSummarySection');
const briefHqValue = document.getElementById('briefHqValue');
const briefRevenueValue = document.getElementById('briefRevenueValue');
const briefIndustryValue = document.getElementById('briefIndustryValue');
const topNewsList = document.getElementById('topNewsList');
const topNewsHint = document.getElementById('topNewsHint');
const legacyBriefContainer = document.getElementById('legacyBriefContainer');
const personasDiv = document.getElementById('personas');
const emailOut = document.getElementById('emailOut');
const telePitchOut = document.getElementById('telePitchOut');
const copyEmailBtn = document.getElementById('copyEmail');
const personaTabs = document.getElementById('personaTabs');
const viewDocs = document.getElementById('viewDocs');
const historyList = document.getElementById('historyList');
const historyEmpty = document.getElementById('historyEmpty');
const targetHistoryList = document.getElementById('targetHistoryList');
const targetHistoryEmpty = document.getElementById('targetHistoryEmpty');
const exportTrigger = document.getElementById('exportTrigger');
const settingsTrigger = document.getElementById('settingsTrigger');
const modalRoot = document.getElementById('modalRoot');
const modalTitle = document.getElementById('modalTitle');
const modalBody = document.getElementById('modalBody');
const modalFooter = document.getElementById('modalFooter');
const modalClose = document.getElementById('modalClose');
const modeTabs = document.querySelectorAll('[data-mode-tab]');
const modeViews = document.querySelectorAll('[data-mode-view]');
const historySections = document.querySelectorAll('[data-history-section]');
const targetForm = document.getElementById('targetForm');
const targetProductInput = document.getElementById('targetProduct');
const targetDocInput = document.getElementById('targetDocInput');
const targetSectorsInput = document.getElementById('targetSectorsInput');
const targetSectorsChips = document.getElementById('targetSectorsChips');
const targetLocationInput = document.getElementById('targetLocation');
const targetStatusEl = document.getElementById('targetStatus');
const targetResultsSection = document.getElementById('targetResults');
const targetResultsList = document.getElementById('targetResultsList');
const copyTargetsBtn = document.getElementById('copyTargets');
const briefDocPickerBtn = document.getElementById('briefDocPicker');
const targetDocPickerBtn = document.getElementById('targetDocPicker');
const briefDocChips = document.getElementById('briefDocChips');
const targetDocChips = document.getElementById('targetDocChips');

const EXPORT_TEMPLATE_STORAGE_KEY = 'exportTemplate';
const EXPORT_TEMPLATE_DRAFT_STORAGE_KEY = 'exportTemplateDraft';
const EXPORT_TEMPLATES_STORAGE_KEY = 'exportTemplates';
const EXPORT_MODAL_STATE_STORAGE_KEY = 'exportModalState';
const EXPORT_PAGE_SIZE = 8;
const PITCH_FROM_COMPANY_KEY = 'pitchFromCompany';

const Mode = {
  TARGETED_BRIEF: 'targetedBrief',
  TARGET_GENERATION: 'targetGeneration',
};

const ThemePreference = {
  SYSTEM: 'system',
  LIGHT: 'light',
  DARK: 'dark',
};
const THEME_PREFERENCE_STORAGE_KEY = 'themePreference';
const systemThemeMedia = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;

let historyEntries = [];
let currentHistoryId = null;
let targetHistoryEntries = [];
let currentTargetHistoryId = null;
let personaEmailDrafts = [];
let selectedPersonaIndex = -1;
let exportTemplates = [];
let selectedTemplateId = null;
let exportTemplateDraft = { name: '', format: 'xlsx', columns: [] };
let availableDateFields = [];
let defaultDateFieldPath = 'createdAt';
let activeModalCleanup = null;
let activeExportState = null;
let activeMode = null;
let latestTargetResultsText = '';
let targetSectors = [];
let telephonicPitchErrorMessage = '';
let telephonicPitchDebugAttempts = [];
let storedDocs = [];
let activeBriefData = null;
const docLabelCache = new Map();
const selectedDocsByMode = {
  [Mode.TARGETED_BRIEF]: [],
  [Mode.TARGET_GENERATION]: [],
};
const briefProgressState = { runId: null, total: 0, current: 0 };
let themePreference = ThemePreference.SYSTEM;
let cachedGeminiKey = '';

function debounce(fn, wait = 150) {
  let timeoutId = null;
  return (...args) => {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
    timeoutId = setTimeout(() => {
      timeoutId = null;
      try {
        fn(...args);
      } catch (err) {
        console.warn('Debounced function failed', err);
      }
    }, wait);
  };
}

function normalizeThemePreference(value) {
  if (value === ThemePreference.DARK) return ThemePreference.DARK;
  if (value === ThemePreference.LIGHT) return ThemePreference.LIGHT;
  return ThemePreference.SYSTEM;
}

function resolveTheme(preference = themePreference) {
  const pref = normalizeThemePreference(preference);
  if (pref === ThemePreference.LIGHT) return ThemePreference.LIGHT;
  if (pref === ThemePreference.DARK) return ThemePreference.DARK;
  return systemThemeMedia?.matches ? ThemePreference.DARK : ThemePreference.LIGHT;
}

function applyTheme(preference = themePreference) {
  themePreference = normalizeThemePreference(preference);
  const resolved = resolveTheme(themePreference);
  if (document?.body) {
    document.body.classList.toggle('theme-dark', resolved === ThemePreference.DARK);
    document.body.classList.toggle('theme-light', resolved === ThemePreference.LIGHT);
  }
  return resolved;
}

async function saveThemePreference(preference) {
  const normalized = normalizeThemePreference(preference);
  applyTheme(normalized);
  try {
    await chrome.storage.local.set({ [THEME_PREFERENCE_STORAGE_KEY]: normalized });
  } catch (err) {
    console.warn('Failed to save theme preference', err);
  }
  return normalized;
}

async function loadThemePreferenceFromStorage() {
  let storedPreference = ThemePreference.SYSTEM;
  try {
    const data = await chrome.storage.local.get([THEME_PREFERENCE_STORAGE_KEY]);
    const stored = data && data[THEME_PREFERENCE_STORAGE_KEY];
    storedPreference = normalizeThemePreference(stored);
  } catch (err) {
    console.warn('Failed to load theme preference', err);
  }
  return applyTheme(storedPreference);
}

const handleSystemThemeChange = () => {
  if (themePreference === ThemePreference.SYSTEM) {
    applyTheme(themePreference);
  }
};

if (systemThemeMedia?.addEventListener) {
  systemThemeMedia.addEventListener('change', handleSystemThemeChange);
} else if (systemThemeMedia?.addListener) {
  systemThemeMedia.addListener(handleSystemThemeChange);
}

applyTheme(themePreference);
const themePreferenceLoadPromise = loadThemePreferenceFromStorage();
const geminiKeyLoadPromise = chrome.storage.local
  .get(['geminiKey'])
  .then((data) => {
    const val = data && typeof data.geminiKey === 'string' ? data.geminiKey : '';
    cachedGeminiKey = val;
    if (apiKeyInput) {
      apiKeyInput.value = val;
    }
    return val;
  })
  .catch(() => '');

function setActiveMode(mode) {
  const validModes = Object.values(Mode);
  if (!validModes.includes(mode)) {
    mode = Mode.TARGETED_BRIEF;
  }
  if (activeMode === mode) {
    return;
  }
  activeMode = mode;

  Array.from(modeTabs).forEach((tab) => {
    const isActive = tab.dataset.modeTab === mode;
    tab.classList.toggle('active', isActive);
    tab.setAttribute('aria-selected', isActive ? 'true' : 'false');
    tab.removeAttribute('tabindex');
  });

  Array.from(modeViews).forEach((view) => {
    const isActive = view.dataset.modeView === mode;
    if (isActive) {
      view.removeAttribute('hidden');
      view.setAttribute('aria-hidden', 'false');
    } else {
      view.setAttribute('hidden', 'hidden');
      view.setAttribute('aria-hidden', 'true');
    }
  });

  Array.from(historySections).forEach((section) => {
    const isActive = section.dataset.historySection === mode;
    if (isActive) {
      section.removeAttribute('hidden');
      section.setAttribute('aria-hidden', 'false');
    } else {
      section.setAttribute('hidden', 'hidden');
      section.setAttribute('aria-hidden', 'true');
    }
  });
}

function sendMessagePromise(payload) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(payload, (resp) => {
      const err = chrome.runtime.lastError;
      if (err) {
        reject(err);
        return;
      }
      resolve(resp);
    });
  });
}

async function loadStoredDocs() {
  try {
    const resp = await sendMessagePromise({ action: 'listDocs' });
    storedDocs = Array.isArray(resp?.docs) ? resp.docs : [];
    storedDocs.forEach((doc) => {
      if (doc && doc.id && doc.name) {
        docLabelCache.set(String(doc.id), doc.name);
      }
    });
  } catch (err) {
    console.warn('Failed to load docs', err);
    storedDocs = [];
  }
  return storedDocs;
}

function setTargetStatus(message, options = {}) {
  if (!targetStatusEl) return;
  targetStatusEl.textContent = message || '';
  targetStatusEl.classList.toggle('error', !!options.error);
}

function getSelectedDocObjects(mode) {
  const ids = selectedDocsByMode[mode] || [];
  return ids
    .map((id) => {
      const doc = storedDocs.find((d) => String(d.id) === String(id));
      return doc || null;
    })
    .filter(Boolean);
}

function renderDocChips(mode) {
  const container = mode === Mode.TARGETED_BRIEF ? briefDocChips : targetDocChips;
  if (!container) return;
  container.innerHTML = '';

  const ids = selectedDocsByMode[mode] || [];
  ids.forEach((id) => {
    const doc = storedDocs.find((d) => String(d.id) === String(id));
    const fallbackName = docLabelCache.get(String(id));
    const name = doc?.name || fallbackName || 'Document';

    const chip = document.createElement('span');
    chip.className = 'doc-chip';

    const label = document.createElement('span');
    label.textContent = name;
    chip.appendChild(label);

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'doc-chip-remove';
    removeBtn.dataset.removeDocId = String(id);
    removeBtn.setAttribute('aria-label', `Remove ${name}`);
    removeBtn.textContent = 'A-';
    chip.appendChild(removeBtn);

    container.appendChild(chip);
  });
}

function setSelectedDocs(mode, ids = []) {
  const deduped = Array.from(new Set((Array.isArray(ids) ? ids : []).filter(Boolean).map(String)));
  selectedDocsByMode[mode] = deduped;
  renderDocChips(mode);
}

function setSelectedDocsFromRefs(mode, refs = []) {
  const ids = [];
  (Array.isArray(refs) ? refs : []).forEach((ref) => {
    if (!ref) return;
    if (typeof ref === 'string') {
      ids.push(ref);
      return;
    }
    const id = ref.id || ref.name || '';
    if (id) {
      ids.push(String(id));
    }
    if (ref.name && id) {
      docLabelCache.set(String(id), ref.name);
    }
  });
  setSelectedDocs(mode, ids);
}

function clearSelectedDocs(mode) {
  setSelectedDocs(mode, []);
}

function resetTargetResults() {
  if (targetResultsList) {
    targetResultsList.innerHTML = '';
  }
  if (targetResultsSection) {
    targetResultsSection.setAttribute('hidden', 'hidden');
  }
  latestTargetResultsText = '';
  if (copyTargetsBtn) {
    copyTargetsBtn.disabled = true;
  }
}

function setTelePitchOutput(message, { isError = false } = {}) {
  if (!telePitchOut) return;
  const text = typeof message === 'string' ? message : '';
  telePitchOut.innerText = text;
  telePitchOut.style.color = isError ? '#b91c1c' : '';
}

function normalizeTargetSector(value) {
  if (typeof value !== 'string') return '';
  return value.replace(/\s+/g, ' ').trim();
}

function renderTargetSectorChips() {
  if (!targetSectorsChips) return;
  targetSectorsChips.innerHTML = '';

  targetSectors.forEach((sector) => {
    const chip = document.createElement('span');
    chip.className = 'sector-chip';

    const label = document.createElement('span');
    label.textContent = sector;
    chip.appendChild(label);

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'sector-chip-remove';
    removeBtn.dataset.action = 'remove-sector';
    removeBtn.dataset.value = sector;
    removeBtn.setAttribute('aria-label', `Remove ${sector}`);
    removeBtn.textContent = '×';
    chip.appendChild(removeBtn);

    targetSectorsChips.appendChild(chip);
  });
}

function addTargetSector(value) {
  const normalized = normalizeTargetSector(value);
  if (!normalized) return;
  const exists = targetSectors.some((sector) => sector.toLowerCase() === normalized.toLowerCase());
  if (exists) return;
  targetSectors.push(normalized);
  renderTargetSectorChips();
}

function removeTargetSector(value) {
  const normalized = normalizeTargetSector(value);
  if (!normalized) return;
  const idx = targetSectors.findIndex((sector) => sector.toLowerCase() === normalized.toLowerCase());
  if (idx === -1) return;
  targetSectors.splice(idx, 1);
  renderTargetSectorChips();
}

function setTargetSectors(values = []) {
  const next = [];
  (Array.isArray(values) ? values : []).forEach((value) => {
    const normalized = normalizeTargetSector(value);
    if (!normalized) return;
    if (next.some((sector) => sector.toLowerCase() === normalized.toLowerCase())) return;
    next.push(normalized);
  });

  targetSectors = next;
  renderTargetSectorChips();
}

function commitTargetSectorInput() {
  if (!targetSectorsInput) return;
  const pending = normalizeTargetSector(targetSectorsInput.value);
  if (pending) {
    addTargetSector(pending);
  }
  targetSectorsInput.value = '';
}

setTargetSectors([]);
loadStoredDocs().finally(() => {
  renderDocChips(Mode.TARGETED_BRIEF);
  renderDocChips(Mode.TARGET_GENERATION);
});

function resetBriefProgress() {
  briefProgressState.runId = null;
  briefProgressState.total = 0;
  briefProgressState.current = 0;
}

function createEmptyBriefData() {
  return {
    brief_html: '',
    company_name: '',
    revenue_estimate: '',
    industry_sector: '',
    top_5_news: [],
    hq_location: '',
    hq_lookup_error: '',
    personas: [],
    personaEmails: [],
    telephonicPitches: [],
    telephonicPitchError: '',
    telephonicPitchAttempts: [],
    email: {},
    overview_error: '',
    persona_error: '',
  };
}

function mergeBriefData(partial = {}) {
  if (!activeBriefData) activeBriefData = createEmptyBriefData();
  const next = { ...activeBriefData };

  const overview = partial.overview && typeof partial.overview === 'object' ? partial.overview : null;
  if (overview) {
    if (overview.company_name && !partial.company_name) partial.company_name = overview.company_name;
    if (overview.revenue_estimate && !partial.revenue_estimate) partial.revenue_estimate = overview.revenue_estimate;
    if (overview.industry_sector && !partial.industry_sector) partial.industry_sector = overview.industry_sector;
    if (Array.isArray(overview.top_5_news) && !Array.isArray(partial.top_5_news)) {
      partial.top_5_news = overview.top_5_news;
    }
    if (!partial.top_5_news && Array.isArray(overview.top_news)) {
      partial.top_5_news = overview.top_news;
    }
  }
  const assign = (key) => {
    if (partial[key] !== undefined && partial[key] !== null) {
      next[key] = partial[key];
    }
  };

  assign('brief_html');
  assign('company_name');
  assign('revenue_estimate');
  assign('industry_sector');
  assign('hq_location');
  assign('hq_lookup_error');
  assign('overview_error');
  assign('persona_error');
  if (!partial.brief_html && partial.brief && !next.brief_html) {
    next.brief_html = partial.brief;
  }

  if (Array.isArray(partial.top_5_news)) next.top_5_news = partial.top_5_news;
  if (Array.isArray(partial.personas)) next.personas = partial.personas;
  if (Array.isArray(partial.personaEmails)) next.personaEmails = partial.personaEmails;
  if (Array.isArray(partial.telephonicPitches)) next.telephonicPitches = partial.telephonicPitches;
  if (Array.isArray(partial.telephonicPitchAttempts)) next.telephonicPitchAttempts = partial.telephonicPitchAttempts;
  assign('telephonicPitchError');
  if (partial.email !== undefined) next.email = partial.email || {};

  activeBriefData = next;
  return activeBriefData;
}

function renderTopNewsEntries(newsItems = [], options = {}) {
  if (!topNewsList) return;
  topNewsList.innerHTML = '';
  const errorText = options.errorText || '';
  const message = options.message || '';

  if (errorText) {
    const li = document.createElement('li');
    li.textContent = errorText;
    topNewsList.appendChild(li);
    if (topNewsHint) topNewsHint.textContent = 'Overview call failed';
    return;
  }

  if (!newsItems.length) {
    const li = document.createElement('li');
    li.textContent = message || 'Waiting for news...';
    topNewsList.appendChild(li);
    if (topNewsHint) topNewsHint.textContent = '';
    return;
  }

  newsItems.slice(0, 5).forEach((item) => {
    const li = document.createElement('li');
    const title = document.createElement('div');
    title.className = 'news-title';
    title.textContent = item.title || 'Untitled headline';
    const summary = document.createElement('div');
    summary.className = 'news-summary';
    summary.textContent = item.summary || '';
    li.appendChild(title);
    li.appendChild(summary);
    topNewsList.appendChild(li);
  });
  if (topNewsHint) topNewsHint.textContent = '';
}

function renderBriefFromState() {
  if (!briefDiv) return;
  if (!activeBriefData) activeBriefData = createEmptyBriefData();

  const hqText = activeBriefData.hq_location || activeBriefData.hq_lookup_error || 'Locating headquarters...';
  if (briefHqValue) briefHqValue.textContent = hqText;

  const revenueText = activeBriefData.revenue_estimate || (activeBriefData.overview_error ? 'Unavailable' : 'Generating revenue...');
  if (briefRevenueValue) briefRevenueValue.textContent = revenueText;

  const industryText = activeBriefData.industry_sector || (activeBriefData.overview_error ? 'Unavailable' : 'Identifying sector...');
  if (briefIndustryValue) briefIndustryValue.textContent = industryText;

  const newsItems = Array.isArray(activeBriefData.top_5_news) ? activeBriefData.top_5_news : [];
  renderTopNewsEntries(newsItems, {
    errorText: activeBriefData.overview_error,
    message: newsItems.length ? '' : 'Waiting for overview...',
  });

  if (legacyBriefContainer) {
    const shouldShowLegacy =
      !newsItems.length &&
      !activeBriefData.revenue_estimate &&
      !activeBriefData.industry_sector &&
      !!activeBriefData.brief_html;
    if (briefSummarySection) {
      briefSummarySection.hidden = shouldShowLegacy ? true : false;
    }
    legacyBriefContainer.hidden = !shouldShowLegacy;
    legacyBriefContainer.innerHTML = shouldShowLegacy ? activeBriefData.brief_html : '';
  }
}

function renderPersonaSectionsFromState() {
  if (!activeBriefData) return;
  const personasData = Array.isArray(activeBriefData.personas) ? activeBriefData.personas : [];

  telephonicPitchErrorMessage = typeof activeBriefData.telephonicPitchError === 'string'
    ? activeBriefData.telephonicPitchError.trim()
    : '';
  telephonicPitchDebugAttempts = Array.isArray(activeBriefData.telephonicPitchAttempts)
    ? activeBriefData.telephonicPitchAttempts
    : [];
  if (telephonicPitchErrorMessage && telephonicPitchDebugAttempts.length) {
    console.warn('Telephonic pitch generation attempts', telephonicPitchDebugAttempts);
  }

  if (activeBriefData.persona_error && !personasData.length) {
    if (personasDiv) {
      personasDiv.innerHTML = '';
      const p = document.createElement('p');
      p.textContent = `Persona generation failed: ${activeBriefData.persona_error}`;
      personasDiv.appendChild(p);
    }
  } else {
    clearAndRenderPersonas(personasData);
  }

  const personaEmailsData = Array.isArray(activeBriefData.personaEmails) ? activeBriefData.personaEmails : [];
  const telephonicPitchesData = Array.isArray(activeBriefData.telephonicPitches) ? activeBriefData.telephonicPitches : [];
  renderPersonaEmailDrafts(personasData, personaEmailsData, telephonicPitchesData, activeBriefData.email);
}

function prepareResultShell() {
  activeBriefData = createEmptyBriefData();
  if (resultDiv) resultDiv.style.display = 'block';
  if (briefSummarySection) briefSummarySection.hidden = false;
  if (briefHqValue) briefHqValue.textContent = 'Locating headquarters...';
  if (briefRevenueValue) briefRevenueValue.textContent = 'Generating revenue estimate...';
  if (briefIndustryValue) briefIndustryValue.textContent = 'Identifying industry sector...';
  renderTopNewsEntries([], { message: 'Waiting for overview...' });
  if (legacyBriefContainer) {
    legacyBriefContainer.hidden = true;
    legacyBriefContainer.innerHTML = '';
  }

  if (personasDiv) personasDiv.innerHTML = '<p>Fetching personas...</p>';
  personaEmailDrafts = [];
  selectedPersonaIndex = -1;
  if (personaTabs) {
    personaTabs.innerHTML = '';
    personaTabs.style.display = 'none';
  }
  if (emailOut) emailOut.innerText = 'Generating email drafts...';
  telephonicPitchErrorMessage = '';
  telephonicPitchDebugAttempts = [];
  setTelePitchOutput('Generating telephonic pitches...');
  updateCopyEmailButtonState('');
}

function renderBriefProgress(label = 'Working...') {
  if (!status) return;
  status.innerHTML = '';
  const spinner = document.createElement('span');
  spinner.className = 'spinner';
  status.appendChild(spinner);
  const text = document.createElement('span');
  text.style.marginLeft = '8px';
  const total = briefProgressState.total || '?';
  const current = typeof briefProgressState.current === 'number' ? briefProgressState.current : 0;
  text.textContent = `${label} (${current}/${total})`;
  status.appendChild(text);
  status.style.color = '';
}

function startBriefProgress(runId, initialLabel = 'Starting...', initialTotal = 3) {
  briefProgressState.runId = runId;
  briefProgressState.total = initialTotal;
  briefProgressState.current = 0;
  renderBriefProgress(initialLabel);
}

chrome.runtime.onMessage.addListener((message) => {
  if (!message) return;

  if (message.action === 'briefProgress') {
    if (!briefProgressState.runId || message.runId !== briefProgressState.runId) return;

    const nextTotal = Number(message.total);
    if (Number.isFinite(nextTotal) && nextTotal > 0) {
      briefProgressState.total = nextTotal;
    }

    const nextCurrent = Number(message.current);
    if (Number.isFinite(nextCurrent)) {
      const total = briefProgressState.total || nextCurrent;
      briefProgressState.current = Math.min(nextCurrent, total);
    }

    renderBriefProgress(typeof message.label === 'string' && message.label ? message.label : 'Working...');
    return;
  }

  if (message.action === 'briefPartialUpdate') {
    if (!briefProgressState.runId || message.runId !== briefProgressState.runId) return;
    const payload = message.payload || {};
    mergeBriefData(payload);
    renderBriefFromState();
    if (
      payload.personas ||
      payload.personaEmails ||
      payload.telephonicPitches ||
      payload.email ||
      payload.persona_error
    ) {
      renderPersonaSectionsFromState();
    }
  }
});

function buildTargetsCopyText(companies = []) {
  return companies
    .map((company, idx) => {
      const parts = [];
      const name = company?.name ? String(company.name).trim() : '';
      const website = company?.website ? String(company.website).trim() : '';
      const revenue = company?.revenue ? String(company.revenue).trim() : '';
      if (name) {
        parts.push(`${idx + 1}. ${name}`);
      } else {
        parts.push(`${idx + 1}.`);
      }
      if (website) {
        parts.push(`Website: ${website}`);
      }
      if (revenue) {
        parts.push(`Revenue: ${revenue}`);
      }
      if (company?.notes) {
        parts.push(`Notes: ${String(company.notes).trim()}`);
      }
      return parts.join(' | ');
    })
    .join('\n');
}

function renderTargetResults(companies = []) {
  if (!targetResultsSection || !targetResultsList) {
    latestTargetResultsText = '';
    return;
  }

  targetResultsList.innerHTML = '';

  if (!companies.length) {
    const empty = document.createElement('li');
    empty.className = 'target-results-empty';
    empty.textContent = 'No companies returned by Gemini.';
    targetResultsList.appendChild(empty);
    targetResultsSection.removeAttribute('hidden');
    latestTargetResultsText = '';
    if (copyTargetsBtn) {
      copyTargetsBtn.disabled = true;
    }
    return;
  }

  companies.forEach((company) => {
    const item = document.createElement('li');
    item.className = 'target-results-item';

    const title = document.createElement('p');
    title.className = 'target-company-name';
    title.textContent = company?.name || 'Unnamed Company';
    item.appendChild(title);

    const meta = document.createElement('div');
    meta.className = 'target-company-meta';
    if (company?.website) {
      try {
        const trimmed = String(company.website).trim();
        const normalized = trimmed && /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
        const url = new URL(normalized);
        const link = document.createElement('a');
        link.href = url.href;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = url.hostname || url.href;
        meta.appendChild(link);
      } catch (err) {
        const fallback = document.createElement('span');
        fallback.textContent = company.website;
        meta.appendChild(fallback);
      }
    }
    if (company?.revenue) {
      const revenueEl = document.createElement('span');
      revenueEl.textContent = `Revenue: ${company.revenue}`;
      meta.appendChild(revenueEl);
    }
    if (meta.children.length) {
      item.appendChild(meta);
    }

    if (company?.notes) {
      const notes = document.createElement('p');
      notes.className = 'target-company-notes';
      notes.textContent = company.notes;
      item.appendChild(notes);
    }

    targetResultsList.appendChild(item);
  });

  targetResultsSection.removeAttribute('hidden');
  latestTargetResultsText = buildTargetsCopyText(companies);

  if (copyTargetsBtn) {
    copyTargetsBtn.disabled = !latestTargetResultsText;
  }
}

function arrayBufferToBase64(buffer) {
  if (!buffer) return '';
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const sub = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, sub);
  }
  try {
    return btoa(binary);
  } catch (err) {
    return '';
  }
}

async function fileToBase64(file) {
  if (!file) return '';
  const buffer = await file.arrayBuffer();
  return arrayBufferToBase64(buffer);
}

async function uploadDocs(files = []) {
  const uploadedIds = [];
  for (const file of files) {
    try {
      const base64 = await fileToBase64(file);
      const resp = await sendMessagePromise({ action: 'storeDoc', name: file.name || 'uploaded-document', content_b64: base64 });
      if (resp && resp.ok && resp.id) {
        uploadedIds.push(String(resp.id));
      }
    } catch (err) {
      console.warn('Failed to upload document', err);
    }
  }
  await loadStoredDocs();
  return uploadedIds;
}

async function readFileContentForTargets(file) {
  if (!file) {
    return { name: '', text: '', base64: '' };
  }

  const name = file.name || 'uploaded-document';

  const text = await new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
    reader.onerror = () => resolve('');
    try {
      reader.readAsText(file);
    } catch (err) {
      resolve('');
    }
  });

  const arrayBuffer = await new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => resolve(null);
    try {
      reader.readAsArrayBuffer(file);
    } catch (err) {
      resolve(null);
    }
  });

  const base64 = arrayBufferToBase64(arrayBuffer);
  const trimmedText = typeof text === 'string' ? text.slice(0, 8000) : '';

  return { name, text: trimmedText, base64 };
}
resetTargetResults();

updateCopyEmailButtonState('');

const modeTabList = Array.from(modeTabs);
if (modeTabList.length) {
  modeTabList.forEach((tab) => {
    tab.addEventListener('click', () => {
      const targetMode = tab.dataset.modeTab || Mode.TARGETED_BRIEF;
      setActiveMode(targetMode);
    });
  });
  const initialTab = modeTabList.find((tab) => tab.classList.contains('active') || tab.getAttribute('aria-selected') === 'true');
  const initialMode = initialTab?.dataset.modeTab || Mode.TARGETED_BRIEF;
  setActiveMode(initialMode);
}

saveKeyBtn?.addEventListener('click', async () => {
  const key = apiKeyInput.value.trim();
  if (!key) { status.innerText = 'API key required'; status.style.color = '#b91c1c'; return; }
  await chrome.storage.local.set({ geminiKey: key });
  status.innerText = 'API key saved.';
  status.style.color = '';
});

fileInput?.addEventListener('change', async (e) => {
  const files = e.target.files;
  if (!files || !files.length) return;
  const ids = await uploadDocs(files);
  if (ids.length) {
    setSelectedDocs(Mode.TARGETED_BRIEF, [...selectedDocsByMode[Mode.TARGETED_BRIEF], ...ids]);
  }
  status.innerText = `Uploaded ${ids.length || files.length} file(s).`;
  status.style.color = '';
  e.target.value = '';
});

targetDocInput?.addEventListener('change', async (e) => {
  const files = e.target.files;
  if (!files || !files.length) return;
  const ids = await uploadDocs(files);
  if (ids.length) {
    setSelectedDocs(Mode.TARGET_GENERATION, [...selectedDocsByMode[Mode.TARGET_GENERATION], ...ids]);
  }
  setTargetStatus(`Uploaded ${ids.length || files.length} file(s).`, { error: false });
  e.target.value = '';
});

viewDocs?.addEventListener('click', async () => {
  chrome.runtime.sendMessage({ action: 'listDocs' }, (resp) => {
    const docs = resp && resp.docs ? resp.docs : [];
    if (!docs.length) {
      alert('No docs stored.');
      return;
    }
    const names = docs.map(d => d.name).join('\n');
    alert('Stored docs:\n' + names);
  });
});

targetForm?.addEventListener('submit', async (evt) => {
  evt.preventDefault();
  const product = targetProductInput?.value.trim() || '';
  const location = targetLocationInput?.value.trim() || '';
  const sectors = [...targetSectors];

  if (!product) {
    setTargetStatus('Product name is required to generate targets.', { error: true });
    return;
  }

  setTargetStatus('Generating targets...', { error: false });
  resetTargetResults();

  const submitBtn = targetForm.querySelector('#generateTargets');
  if (submitBtn) {
    submitBtn.disabled = true;
  }

  let docPayload = { name: '', text: '', base64: '' };
  const file = targetDocInput?.files?.[0] || null;
  if (file) {
    docPayload = await readFileContentForTargets(file);
  }

  try {
    await loadStoredDocs();
    const docs = getSelectedDocObjects(Mode.TARGET_GENERATION);

    const resp = await sendMessagePromise({
      action: 'generateTargets',
      product,
      location,
      sectors,
      docs,
      docName: docPayload.name,
      docText: docPayload.text,
      docBase64: docPayload.base64,
    });

    if (!resp) {
      setTargetStatus('No response received from background script.', { error: true });
      return;
    }

    if (resp.error) {
      setTargetStatus(`Error: ${resp.error}`, { error: true });
      return;
    }

    const companies = Array.isArray(resp.companies) ? resp.companies : [];
    renderTargetResults(companies);
    if (resp && resp.ok) {
      loadTargetHistory({ selectLatest: true, autoShow: false, updateForm: false, statusText: '' });
    }

    if (companies.length) {
      setTargetStatus(`Found ${companies.length} potential companies.`, { error: false });
    } else {
      setTargetStatus('No companies returned. Try refining the inputs.', { error: true });
    }
  } catch (err) {
    setTargetStatus(`Failed to generate targets: ${err?.message || err}`, { error: true });
  } finally {
    if (submitBtn) {
      submitBtn.disabled = false;
    }
  }
});

targetSectorsInput?.addEventListener('input', (evt) => {
  const value = evt.target.value;
  if (!value.includes(',')) return;
  const parts = value.split(',');
  const remainder = parts.pop();
  parts.forEach((part) => addTargetSector(part));
  evt.target.value = remainder ? remainder.trimStart() : '';
});

targetSectorsInput?.addEventListener('keydown', (evt) => {
  if (evt.key === 'Enter') {
    evt.preventDefault();
    commitTargetSectorInput();
  }
});

targetSectorsInput?.addEventListener('blur', () => {
  commitTargetSectorInput();
});

targetSectorsChips?.addEventListener('click', (evt) => {
  const button = evt.target.closest('[data-action="remove-sector"]');
  if (!button) return;
  removeTargetSector(button.dataset.value || '');
  targetSectorsInput?.focus();
});

copyTargetsBtn?.addEventListener('click', async () => {
  if (!latestTargetResultsText) {
    setTargetStatus('No results to copy yet.', { error: true });
    return;
  }

  try {
    await navigator.clipboard.writeText(latestTargetResultsText);
    setTargetStatus('Target list copied to clipboard.', { error: false });
  } catch (err) {
    const helper = document.createElement('textarea');
    helper.value = latestTargetResultsText;
    helper.setAttribute('readonly', '');
    helper.style.position = 'absolute';
    helper.style.left = '-9999px';
    document.body.appendChild(helper);
    helper.select();
    try {
      const ok = document.execCommand('copy');
      document.body.removeChild(helper);
      if (ok) {
        setTargetStatus('Target list copied to clipboard.', { error: false });
      } else {
        throw new Error('Copy command failed');
      }
    } catch (copyErr) {
      document.body.removeChild(helper);
      setTargetStatus('Unable to copy results automatically.', { error: true });
    }
  }
});

newResearchBtn?.addEventListener('click', () => {
  companyEl && (companyEl.value = '');
  locationEl && (locationEl.value = '');
  productEl && (productEl.value = '');
  if (fileInput) fileInput.value = '';
  clearSelectedDocs(Mode.TARGETED_BRIEF);

  currentHistoryId = null;
  setActiveHistoryItem('');

  if (status) {
    status.innerText = 'Ready for a new research brief.';
    status.style.color = '';
  }

  activeBriefData = createEmptyBriefData();

  if (resultDiv) resultDiv.style.display = 'none';
  if (briefHqValue) briefHqValue.textContent = 'Headquarters not generated yet.';
  if (briefRevenueValue) briefRevenueValue.textContent = 'Revenue not generated yet.';
  if (briefIndustryValue) briefIndustryValue.textContent = 'Industry sector not generated yet.';
  renderTopNewsEntries([], { message: 'Top news not generated yet.' });
  if (legacyBriefContainer) {
    legacyBriefContainer.hidden = true;
    legacyBriefContainer.innerHTML = '';
  }
  if (personasDiv) personasDiv.innerHTML = '';

  personaEmailDrafts = [];
  selectedPersonaIndex = -1;

  if (personaTabs) {
    personaTabs.innerHTML = '';
    personaTabs.style.display = 'none';
  }

  if (emailOut) emailOut.innerText = 'No email generated yet.';
  telephonicPitchErrorMessage = '';
  telephonicPitchDebugAttempts = [];
  setTelePitchOutput('No telephonic pitch generated yet.');
  updateCopyEmailButtonState('');
});

newTargetResearchBtn?.addEventListener('click', () => {
  if (targetProductInput) targetProductInput.value = '';
  if (targetLocationInput) targetLocationInput.value = '';
  if (targetDocInput) targetDocInput.value = '';
  if (targetSectorsInput) targetSectorsInput.value = '';
  setTargetSectors([]);
  clearSelectedDocs(Mode.TARGET_GENERATION);

  currentTargetHistoryId = null;
  setActiveTargetHistoryItem('');

  resetTargetResults();
  setTargetStatus('Ready for a new target search.', { error: false });
});

briefDocPickerBtn?.addEventListener('click', () => openDocPickerModal(Mode.TARGETED_BRIEF));
targetDocPickerBtn?.addEventListener('click', () => openDocPickerModal(Mode.TARGET_GENERATION));

briefDocChips?.addEventListener('click', (evt) => {
  const btn = evt.target.closest('[data-remove-doc-id]');
  if (!btn) return;
  const id = btn.dataset.removeDocId;
  const next = (selectedDocsByMode[Mode.TARGETED_BRIEF] || []).filter((docId) => docId !== id);
  setSelectedDocs(Mode.TARGETED_BRIEF, next);
});

targetDocChips?.addEventListener('click', (evt) => {
  const btn = evt.target.closest('[data-remove-doc-id]');
  if (!btn) return;
  const id = btn.dataset.removeDocId;
  const next = (selectedDocsByMode[Mode.TARGET_GENERATION] || []).filter((docId) => docId !== id);
  setSelectedDocs(Mode.TARGET_GENERATION, next);
});

personaTabs?.addEventListener('click', (evt) => {
  const button = evt.target.closest('.persona-tab');
  if (!button) return;
  const { index } = button.dataset;
  const idx = Number(index);
  if (!Number.isNaN(idx)) {
    activatePersonaTab(idx);
  }
});

copyEmailBtn?.addEventListener('click', async () => {
  const text = emailOut?.innerText || '';
  if (!text.trim()) {
    if (status) {
      status.innerText = 'No email to copy yet.';
      status.style.color = '#b91c1c';
    }
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    if (status) {
      const personaName = getActivePersonaName();
      status.innerText = personaName ? `Email for ${personaName} copied to clipboard.` : 'Email copied to clipboard.';
      status.style.color = '';
    }
  } catch (err) {
    const helper = document.createElement('textarea');
    helper.value = text;
    helper.setAttribute('readonly', '');
    helper.style.position = 'absolute';
    helper.style.left = '-9999px';
    document.body.appendChild(helper);
    helper.select();
    try {
      const ok = document.execCommand('copy');
      if (ok) {
        if (status) {
          const personaName = getActivePersonaName();
          status.innerText = personaName ? `Email for ${personaName} copied to clipboard.` : 'Email copied to clipboard.';
          status.style.color = '';
        }
      } else {
        throw new Error('Copy command failed');
      }
    } catch (fallbackErr) {
      if (status) {
        status.innerText = 'Unable to copy email.';
        status.style.color = '#b91c1c';
      }
      console.warn('Copy failed', fallbackErr);
    } finally {
      document.body.removeChild(helper);
    }
  }
});

function closeModal() {
  if (modalRoot?.classList.contains('hidden')) return;
  if (typeof activeModalCleanup === 'function') {
    try {
      activeModalCleanup();
    } catch (err) {
      console.warn('modal cleanup failed', err);
    }
  }
  activeModalCleanup = null;
  if (activeExportState) {
    persistExportModalState(activeExportState).catch((err) => console.warn('Failed to save modal state on close', err));
  }
  if (modalTitle) modalTitle.textContent = '';
  if (modalBody) modalBody.innerHTML = '';
  if (modalFooter) modalFooter.innerHTML = '';
  modalRoot?.classList.add('hidden');
  modalRoot?.setAttribute('aria-hidden', 'true');
}

function openModal({ title = '', render }) {
  if (!modalRoot) return;
  if (!modalRoot.classList.contains('hidden')) {
    closeModal();
  }
  if (modalTitle) modalTitle.textContent = title;
  if (modalBody) modalBody.innerHTML = '';
  if (modalFooter) modalFooter.innerHTML = '';
  modalRoot.classList.remove('hidden');
  modalRoot.setAttribute('aria-hidden', 'false');
  if (typeof render === 'function') {
    const cleanup = render({ body: modalBody, footer: modalFooter, close: closeModal });
    if (typeof cleanup === 'function') {
      activeModalCleanup = cleanup;
    }
  }
}

function openDocPickerModal(mode) {
  const modeLabel = mode === Mode.TARGET_GENERATION ? 'target generation' : 'brief generation';
  loadStoredDocs().then(() => {
    const selectedSet = new Set(selectedDocsByMode[mode] || []);
    let searchTerm = '';

    openModal({
      title: 'Select documents',
      render: ({ body, footer, close }) => {
        if (!body || !footer) return;

        body.innerHTML = '';
        footer.innerHTML = '';

        const helper = document.createElement('div');
        helper.className = 'modal-helper';
        helper.textContent = `Choose documents to attach to your ${modeLabel} run. Upload new files or pick from existing ones.`;
        body.appendChild(helper);

        const searchInput = document.createElement('input');
        searchInput.type = 'search';
        searchInput.className = 'doc-picker-search';
        searchInput.placeholder = 'Search documents by name';
        body.appendChild(searchInput);

        const uploadRow = document.createElement('div');
        uploadRow.className = 'doc-picker-upload';
        const uploadLabel = document.createElement('span');
        uploadLabel.textContent = 'Upload new:';
        const uploadInput = document.createElement('input');
        uploadInput.type = 'file';
        uploadInput.multiple = true;
        uploadInput.accept = '.txt,.pdf,.doc,.docx,.md,.html';
        uploadRow.appendChild(uploadLabel);
        uploadRow.appendChild(uploadInput);
        body.appendChild(uploadRow);

        const list = document.createElement('div');
        list.className = 'doc-picker-list';
        body.appendChild(list);

        const useBtn = document.createElement('button');
        useBtn.type = 'button';
        useBtn.className = 'primary';

        const cancelBtn = document.createElement('button');
        cancelBtn.type = 'button';
        cancelBtn.textContent = 'Cancel';
        cancelBtn.addEventListener('click', () => close());

        const renderList = () => {
          list.innerHTML = '';
          const filtered = (storedDocs || []).filter((doc) => {
            if (!searchTerm) return true;
            const name = (doc?.name || '').toLowerCase();
            return name.includes(searchTerm);
          });

          if (!filtered.length) {
            const empty = document.createElement('div');
            empty.className = 'doc-picker-empty';
            empty.textContent = storedDocs.length
              ? 'No documents match your search.'
              : 'No documents uploaded yet.';
            list.appendChild(empty);
            return;
          }

          filtered.forEach((doc) => {
            const item = document.createElement('div');
            item.className = 'doc-picker-item';
            const label = document.createElement('label');

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = selectedSet.has(String(doc.id));
            checkbox.addEventListener('change', () => {
              const docId = String(doc.id);
              if (checkbox.checked) {
                selectedSet.add(docId);
              } else {
                selectedSet.delete(docId);
              }
              updateUseButton();
            });

            const nameEl = document.createElement('span');
            nameEl.className = 'doc-picker-name';
            nameEl.textContent = doc.name || 'Untitled document';

            const meta = document.createElement('span');
            meta.className = 'doc-picker-meta';
            meta.textContent = doc.id ? `ID: ${doc.id}` : 'Stored document';

            label.appendChild(checkbox);
            label.appendChild(nameEl);
            label.appendChild(meta);
            item.appendChild(label);
            list.appendChild(item);
          });
        };

        const updateUseButton = () => {
          const count = selectedSet.size;
          useBtn.textContent = count ? `Use ${count} document${count === 1 ? '' : 's'}` : 'Select documents';
          useBtn.disabled = count === 0;
        };

        uploadInput.addEventListener('change', async (evt) => {
          const files = evt.target.files;
          if (!files || !files.length) return;
          uploadInput.disabled = true;
          const ids = await uploadDocs(files);
          ids.forEach((id) => selectedSet.add(String(id)));
          await loadStoredDocs();
          renderList();
          updateUseButton();
          uploadInput.value = '';
          uploadInput.disabled = false;
        });

        searchInput.addEventListener('input', () => {
          searchTerm = (searchInput.value || '').toLowerCase();
          renderList();
        });

        useBtn.addEventListener('click', () => {
          setSelectedDocs(mode, Array.from(selectedSet));
          close();
        });

        footer.appendChild(cancelBtn);
        footer.appendChild(useBtn);

        renderList();
        updateUseButton();
      },
    });
  });
}

modalClose?.addEventListener('click', () => closeModal());
modalRoot?.addEventListener('click', (evt) => {
  if (!evt.target) return;
  if (evt.target === modalRoot || evt.target.classList.contains('modal-backdrop')) {
    closeModal();
  }
});

document.addEventListener('keydown', (evt) => {
  if (evt.key === 'Escape' && !modalRoot?.classList.contains('hidden')) {
    closeModal();
  }
});

  settingsTrigger?.addEventListener('click', () => {
    Promise.all([loadExportTemplateFromStorage(), themePreferenceLoadPromise, geminiKeyLoadPromise]).finally(() => {
      openModal({ title: 'Settings', render: renderSettingsModal });
    });
  });

exportTrigger?.addEventListener('click', async () => {
  await loadExportTemplateFromStorage();
  const savedState = await loadExportModalStateFromStorage();
  activeExportState = savedState || null;
  openModal({ title: 'Export Research', render: renderExportModal });
});

function normalizeUrl(u) {
  if (!u) return '';
  u = String(u).trim();

  if (/^https?:\/\//i.test(u)) return u;

  if (/^(www\.)?google\.[a-z]{2,}\/search/i.test(u) || /^google\.[a-z]{2,}\/search/i.test(u)) {
    return 'https://' + u;
  }

  if (!/\s/.test(u) && /\.[a-z]{2,}$/i.test(u)) {
    return u.startsWith('www.') ? 'https://' + u : 'https://' + u;
  }

  if (/site:|zoominfo\.|cognism\.|linkedin\.com\/in|["\s]/i.test(u)) {
    return 'https://www.google.com/search?q=' + encodeURIComponent(u) + '&num=20';
  }

  return 'https://www.google.com/search?q=' + encodeURIComponent(u) + '&num=20';
}

function sanitizeTemplateColumns(columns) {
  if (!Array.isArray(columns)) return [];
  return columns
    .map((col, idx) => {
      const header = col && col.header ? String(col.header).trim() : '';
      const description = col && col.description ? String(col.description).trim() : '';
      if (!header) return null;
      return {
        header,
        description: description || `Column ${idx + 1}`,
      };
    })
    .filter(Boolean);
}

function getTemplateDraftDefaults() {
  return { name: '', format: 'xlsx', columns: [] };
}

function sanitizeTemplateDraft(draft) {
  if (Array.isArray(draft)) {
    return { ...getTemplateDraftDefaults(), columns: sanitizeTemplateColumns(draft) };
  }
  const base = draft && typeof draft === 'object' ? draft : {};
  const name = base.name ? String(base.name).trim() : '';
  const format = base.format === 'md' ? 'md' : 'xlsx';
  return {
    name,
    format,
    columns: sanitizeTemplateColumns(base.columns || []),
  };
}

function normalizeTemplate(template, fallbackIndex = 0) {
  if (!template || typeof template !== 'object') return null;
  const columns = sanitizeTemplateColumns(template.columns || []);
  const format = template.format === 'md' ? 'md' : 'xlsx';
  const nameRaw = template.name ? String(template.name).trim() : '';
  const name = nameRaw || `Template ${fallbackIndex + 1}`;
  const updatedAt = template.updatedAt && !Number.isNaN(new Date(template.updatedAt).getTime())
    ? new Date(template.updatedAt).toISOString()
    : new Date().toISOString();
  const id = template.id || `tpl-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  return { id, name, format, columns, updatedAt };
}

function findTemplateById(id) {
  if (!id) return null;
  return exportTemplates.find((tpl) => tpl.id === id) || null;
}

function getActiveTemplate() {
  const chosen = findTemplateById(selectedTemplateId);
  if (chosen) return chosen;
  return exportTemplates.length ? exportTemplates[0] : null;
}

function setSelectedTemplate(templateId) {
  if (!templateId) {
    selectedTemplateId = exportTemplates[0]?.id || null;
    return selectedTemplateId;
  }
  const exists = exportTemplates.some((tpl) => tpl.id === templateId);
  selectedTemplateId = exists ? templateId : exportTemplates[0]?.id || null;
  return selectedTemplateId;
}

async function persistTemplatesState(options = {}) {
  try {
    await chrome.storage.local.set({
      [EXPORT_TEMPLATES_STORAGE_KEY]: {
        templates: exportTemplates,
        selectedTemplateId,
      },
    });
    if (!options.skipLegacySync) {
      const active = getActiveTemplate();
      if (active) {
        await chrome.storage.local.set({ [EXPORT_TEMPLATE_STORAGE_KEY]: active });
      }
    }
  } catch (err) {
    console.warn('Failed to persist templates', err);
  }
}

function getTemplateEditorInitialDraft() {
  if (exportTemplateDraft && exportTemplateDraft.columns && exportTemplateDraft.columns.length) {
    return exportTemplateDraft;
  }
  const active = getActiveTemplate();
  if (active) {
    return { name: active.name, format: active.format, columns: active.columns || [] };
  }
  return getTemplateDraftDefaults();
}

function getTemplateEditorInitialColumns() {
  const draft = getTemplateEditorInitialDraft();
  return draft.columns || [];
}

function serializeExportModalState(state) {
  if (!state || typeof state !== 'object') return null;
  const allowedSelectionTypes = ['all', 'date', 'custom'];
  const selectionType = allowedSelectionTypes.includes(state.selectionType) ? state.selectionType : 'all';
  return {
    selectionType,
    dateFieldPath: state.dateFieldPath || defaultDateFieldPath,
    dateFrom: state.dateFrom || '',
    dateTo: state.dateTo || '',
    customSearch: state.customSearch || '',
    customPage: Math.max(1, Number(state.customPage) || 1),
    selectedIds: Array.from(state.selectedIds instanceof Set ? state.selectedIds : new Set(state.selectedIds || [])),
    format: state.format === 'md' ? 'md' : 'xlsx',
    templateId: state.templateId || '',
  };
}

async function persistExportModalState(state) {
  const payload = serializeExportModalState(state);
  if (!payload) return;
  try {
    await chrome.storage.local.set({ [EXPORT_MODAL_STATE_STORAGE_KEY]: payload });
  } catch (err) {
    console.warn('Failed to persist export modal state', err);
  }
}

async function loadExportModalStateFromStorage() {
  try {
    const data = await chrome.storage.local.get([EXPORT_MODAL_STATE_STORAGE_KEY]);
    const stored = data && data[EXPORT_MODAL_STATE_STORAGE_KEY];
    if (stored && typeof stored === 'object') {
      return {
        ...serializeExportModalState(stored),
        selectedIds: new Set(Array.isArray(stored.selectedIds) ? stored.selectedIds : []),
        templateId: stored.templateId || '',
        inProgress: false,
        result: null,
        error: '',
      };
    }
  } catch (err) {
    console.warn('Failed to load export modal state', err);
  }
  return null;
}

async function persistExportTemplateDraft(draftInput) {
  const draft = sanitizeTemplateDraft(draftInput);
  exportTemplateDraft = draft;
  try {
    await chrome.storage.local.set({ [EXPORT_TEMPLATE_DRAFT_STORAGE_KEY]: draft });
  } catch (err) {
    console.warn('Failed to save export template draft', err);
  }
  return draft;
}

async function clearExportTemplateDraft() {
  exportTemplateDraft = getTemplateDraftDefaults();
  try {
    await chrome.storage.local.remove(EXPORT_TEMPLATE_DRAFT_STORAGE_KEY);
  } catch (err) {
    console.warn('Failed to clear export template draft', err);
  }
}

const persistTemplateDraftDebounced = debounce((draft) => {
  persistExportTemplateDraft(draft).catch((err) => console.warn('Draft save failed', err));
}, 200);

function setupTemplateEditor(container, initialColumns = [], options = {}) {
  if (!container) {
    return {
      getColumns: () => [],
    };
  }

  let columns = sanitizeTemplateColumns(initialColumns);
  const onChange = typeof options.onChange === 'function' ? options.onChange : null;
  let draggingIndex = null;

  const emitChange = () => {
    if (onChange) {
      onChange(columns.map((col) => ({ ...col })));
    }
  };

  const moveColumn = (fromIdx, toIdx) => {
    if (
      fromIdx === toIdx
      || fromIdx < 0
      || toIdx < 0
      || fromIdx >= columns.length
      || toIdx >= columns.length
    ) {
      return;
    }
    const [moved] = columns.splice(fromIdx, 1);
    columns.splice(toIdx, 0, moved);
    render();
    emitChange();
  };

  const listEl = document.createElement('div');
  listEl.className = 'template-columns';

  const clearDragState = () => {
    Array.from(listEl.querySelectorAll('.template-column-row')).forEach((row) => {
      row.classList.remove('drag-over', 'dragging');
    });
  };

  const actionsEl = document.createElement('div');
  actionsEl.className = 'template-actions';
  const addBtn = document.createElement('button');
  addBtn.type = 'button';
  addBtn.textContent = 'Add column';
  actionsEl.appendChild(addBtn);

  function render() {
    listEl.innerHTML = '';
    if (!columns.length) {
      const empty = document.createElement('div');
      empty.className = 'template-empty';
      empty.textContent = 'No columns configured yet. Add one to start.';
      listEl.appendChild(empty);
      return;
    }

    columns.forEach((col, idx) => {
      const row = document.createElement('div');
      row.className = 'template-column-row';
      row.dataset.index = String(idx);

      const handle = document.createElement('button');
      handle.type = 'button';
      handle.className = 'template-drag-handle';
      handle.setAttribute('aria-label', 'Drag to reorder column');
      handle.draggable = true;
      handle.innerHTML = '&#8942;&#8942;';
      handle.addEventListener('dragstart', (evt) => {
        draggingIndex = idx;
        row.classList.add('dragging');
        if (evt && evt.dataTransfer) {
          evt.dataTransfer.effectAllowed = 'move';
          evt.dataTransfer.setData('text/plain', String(idx));
        }
      });
      handle.addEventListener('dragend', () => {
        draggingIndex = null;
        clearDragState();
      });
      row.appendChild(handle);

      row.addEventListener('dragover', (evt) => {
        if (draggingIndex === null || draggingIndex === idx) return;
        evt.preventDefault();
        clearDragState();
        row.classList.add('drag-over');
      });

      row.addEventListener('dragleave', () => {
        row.classList.remove('drag-over');
      });

      row.addEventListener('drop', (evt) => {
        if (draggingIndex === null) return;
        evt.preventDefault();
        const targetIdx = Number(row.dataset.index);
        moveColumn(draggingIndex, Number.isNaN(targetIdx) ? -1 : targetIdx);
        draggingIndex = null;
      });

      const headerInput = document.createElement('input');
      headerInput.placeholder = 'Column header';
      headerInput.value = col.header || '';
      headerInput.addEventListener('input', (evt) => {
        columns[idx].header = evt.target.value;
        emitChange();
      });
      row.appendChild(headerInput);

      const descInput = document.createElement('input');
      descInput.placeholder = 'Description';
      descInput.value = col.description || '';
      descInput.addEventListener('input', (evt) => {
        columns[idx].description = evt.target.value;
        emitChange();
      });
      row.appendChild(descInput);

      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'template-remove';
      removeBtn.textContent = 'Remove';
      removeBtn.addEventListener('click', () => {
        columns.splice(idx, 1);
        render();
        emitChange();
      });
      row.appendChild(removeBtn);

      listEl.appendChild(row);
    });
  }

  addBtn.addEventListener('click', () => {
    columns.push({ header: '', description: '' });
    render();
    emitChange();
  });

  container.appendChild(listEl);
  container.appendChild(actionsEl);
  render();

  return {
    getColumns: () => sanitizeTemplateColumns(columns),
    setColumns: (cols) => {
      columns = sanitizeTemplateColumns(cols);
      render();
      emitChange();
    },
  };
}

function formatFieldLabel(path) {
  return path
    .split('.')
    .map((part) => part.replace(/\b\w/g, (char) => char.toUpperCase()).replace(/_/g, ' '))
    .join(' > ');
}

function discoverDateFields(entries) {
  const candidates = new Map();

  const walk = (value, path) => {
    if (value === null || value === undefined) return;
    if (typeof value === 'string') {
      const ts = Date.parse(value);
      if (!Number.isNaN(ts)) {
        const existing = candidates.get(path) || { path, count: 0 };
        existing.count += 1;
        candidates.set(path, existing);
      }
      return;
    }
    if (typeof value !== 'object' || Array.isArray(value)) return;
    Object.keys(value).forEach((key) => {
      const nextPath = path ? `${path}.${key}` : key;
      walk(value[key], nextPath);
    });
  };

  entries.forEach((entry) => walk(entry, ''));

  const results = Array.from(candidates.values()).map((item) => ({
    path: item.path,
    label: `${formatFieldLabel(item.path)} (${item.path})`,
    count: item.count,
  }));

  results.sort((a, b) => {
    if (a.path === 'createdAt') return -1;
    if (b.path === 'createdAt') return 1;
    return b.count - a.count;
  });

  return results;
}

function paginate(array, page = 1, pageSize = 10) {
  const start = (page - 1) * pageSize;
  return array.slice(start, start + pageSize);
}

async function loadExportTemplateFromStorage() {
  try {
    const data = await chrome.storage.local.get([EXPORT_TEMPLATES_STORAGE_KEY, EXPORT_TEMPLATE_STORAGE_KEY, EXPORT_TEMPLATE_DRAFT_STORAGE_KEY]);
    const draft = data && data[EXPORT_TEMPLATE_DRAFT_STORAGE_KEY];
    exportTemplateDraft = sanitizeTemplateDraft(draft);

    const collection = data && data[EXPORT_TEMPLATES_STORAGE_KEY];
    const legacyTemplate = data && data[EXPORT_TEMPLATE_STORAGE_KEY];
    let templates = [];
    if (collection && Array.isArray(collection.templates)) {
      templates = collection.templates.map((tpl, idx) => normalizeTemplate(tpl, idx)).filter(Boolean);
      selectedTemplateId = collection.selectedTemplateId || collection.activeTemplateId || null;
    }

    if (!templates.length && legacyTemplate && Array.isArray(legacyTemplate.columns)) {
      const migrated = normalizeTemplate({
        ...legacyTemplate,
        id: legacyTemplate.id || 'legacy-template',
        name: legacyTemplate.name || 'Saved template',
        format: legacyTemplate.format,
        updatedAt: legacyTemplate.updatedAt,
      });
      if (migrated && migrated.columns.length) {
        templates = [migrated];
        selectedTemplateId = migrated.id;
      }
    }

    exportTemplates = templates;
    setSelectedTemplate(selectedTemplateId);

    if (templates.length && (!collection || !collection.templates)) {
      await persistTemplatesState({ skipLegacySync: true });
    }
  } catch (err) {
    console.warn('Failed to load export template', err);
    exportTemplates = [];
    selectedTemplateId = null;
    exportTemplateDraft = getTemplateDraftDefaults();
  }
}

async function persistExportTemplate(template, options = {}) {
  const baseTemplate = (!options.forceNewId && template?.id)
    ? findTemplateById(template.id)
    : (!options.forceNewId ? getActiveTemplate() : null);
  const payload = normalizeTemplate({
    ...baseTemplate,
    ...template,
    id: options.forceNewId ? null : (template?.id || baseTemplate?.id),
    updatedAt: new Date().toISOString(),
  }, exportTemplates.length);
  if (!payload) return null;

  const existingIdx = exportTemplates.findIndex((tpl) => tpl.id === payload.id);
  if (existingIdx !== -1 && !options.forceNewId) {
    exportTemplates[existingIdx] = payload;
  } else {
    exportTemplates = [...exportTemplates, payload];
  }
  setSelectedTemplate(payload.id);
  await persistTemplatesState();
  await clearExportTemplateDraft();
  return payload;
}

function base64ToBlob(base64, mimeType) {
  const binaryString = atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return new Blob([bytes], { type: mimeType });
}

function triggerDownload(downloadInfo) {
  if (!downloadInfo || !downloadInfo.base64) return;
  try {
    const blob = base64ToBlob(downloadInfo.base64, downloadInfo.mimeType || 'application/octet-stream');
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = downloadInfo.filename || 'export';
    document.body.appendChild(anchor);
    anchor.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      document.body.removeChild(anchor);
    }, 200);
  } catch (err) {
    console.error('Download failed', err);
    if (status) {
      status.innerText = 'Unable to start download.';
      status.style.color = '#b91c1c';
    }
  }
}

function renderSettingsModal({ body, footer, close }) {
  if (!body || !footer) return;

  const form = document.createElement('form');
  form.id = 'settingsForm';
  form.noValidate = true;
  const themeInputs = [];

  const errorEl = document.createElement('div');
  errorEl.className = 'modal-error';
  errorEl.style.display = 'none';
  form.appendChild(errorEl);

  const pitchingSection = document.createElement('div');
  pitchingSection.className = 'modal-section';
  const pitchingHeader = document.createElement('h4');
  pitchingHeader.textContent = 'Pitching company';
  pitchingSection.appendChild(pitchingHeader);
  const pitchingHelper = document.createElement('p');
  pitchingHelper.className = 'modal-helper';
  pitchingHelper.textContent = 'Set the company you represent when generating pitches.';
  pitchingSection.appendChild(pitchingHelper);
  const pitchingInput = document.createElement('input');
  pitchingInput.type = 'text';
  pitchingInput.placeholder = 'Your company name (sender)';
  pitchingSection.appendChild(pitchingInput);
  try {
    chrome.storage.local.get([PITCH_FROM_COMPANY_KEY]).then((data) => {
      const stored = data && data[PITCH_FROM_COMPANY_KEY];
      if (typeof stored === 'string') {
        pitchingInput.value = stored;
      }
    });
  } catch (err) {
    console.warn('Failed to load pitching company for settings', err);
  }
  form.appendChild(pitchingSection);

  const themeSection = document.createElement('div');
  themeSection.className = 'modal-section';
  const themeHeader = document.createElement('h4');
  themeHeader.textContent = 'Theme';
  themeSection.appendChild(themeHeader);
  const themeHelper = document.createElement('p');
  themeHelper.className = 'modal-helper';
  themeHelper.textContent = 'Switch between light and dark, or match your device\'s theme.';
  themeSection.appendChild(themeHelper);
  const themeOptions = document.createElement('div');
  themeOptions.className = 'range-radios';
  themeSection.appendChild(themeOptions);
  const currentThemeSetting = themePreference || ThemePreference.SYSTEM;
  [
    { value: ThemePreference.SYSTEM, label: 'Match system' },
    { value: ThemePreference.LIGHT, label: 'Light' },
    { value: ThemePreference.DARK, label: 'Dark' },
  ].forEach((option) => {
    const label = document.createElement('label');
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'themePreference';
    radio.value = option.value;
    radio.checked = currentThemeSetting === option.value;
    label.appendChild(radio);
    label.appendChild(document.createTextNode(option.label));
    themeOptions.appendChild(label);
    themeInputs.push(radio);
  });
  form.appendChild(themeSection);

  const apiSection = document.createElement('div');
  apiSection.className = 'modal-section';
  const apiHeader = document.createElement('h4');
  apiHeader.textContent = 'Add API key';
  apiSection.appendChild(apiHeader);
  const apiHelper = document.createElement('p');
  apiHelper.className = 'modal-helper';
  apiHelper.textContent = 'Provide your Gemini API key. This is required for AI powered exports and brief generation.';
  apiSection.appendChild(apiHelper);
  const apiInputField = document.createElement('input');
  apiInputField.type = 'text';
  apiInputField.placeholder = 'Gemini API key';
    // Pre-fill with cached key or stored value if available.
    apiInputField.value = (apiKeyInput?.value?.trim() || cachedGeminiKey || '').trim();
  apiSection.appendChild(apiInputField);
  form.appendChild(apiSection);

  body.appendChild(form);

  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.textContent = 'Cancel';
  cancelBtn.addEventListener('click', () => close());

  const saveBtn = document.createElement('button');
  saveBtn.type = 'submit';
  saveBtn.textContent = 'Save';
  saveBtn.classList.add('primary');
  saveBtn.setAttribute('form', form.id);

  footer.appendChild(cancelBtn);
  footer.appendChild(saveBtn);

  const onSubmit = async (evt) => {
    evt.preventDefault();
    errorEl.style.display = 'none';
    errorEl.textContent = '';

    const key = apiInputField.value.trim();
    const selectedTheme = (themeInputs.find((input) => input.checked)?.value) || ThemePreference.SYSTEM;
    const pitchingCompany = pitchingInput.value.trim();

    try {
      if (key) {
        await chrome.storage.local.set({ geminiKey: key });
      } else {
        await chrome.storage.local.remove('geminiKey');
      }
      if (pitchingCompany) {
        await chrome.storage.local.set({ [PITCH_FROM_COMPANY_KEY]: pitchingCompany });
      } else {
        await chrome.storage.local.remove(PITCH_FROM_COMPANY_KEY);
      }
      await saveThemePreference(selectedTheme);
      if (apiKeyInput) apiKeyInput.value = key;
      if (status) {
        status.innerText = 'Settings saved.';
        status.style.color = '';
      }
      close();
    } catch (err) {
      errorEl.textContent = err?.message || 'Failed to save settings.';
      errorEl.style.display = 'block';
    }
  };

  form.addEventListener('submit', onSubmit);

  return () => {
    form.removeEventListener('submit', onSubmit);
  };
}

function renderTemplateSetupModal({ body, footer, close }, options = {}) {
  if (!body || !footer) return;

  const initialDraft = getTemplateEditorInitialDraft();
  let currentDraft = { ...initialDraft };

  const container = document.createElement('div');
  container.className = 'modal-section';

  const heading = document.createElement('h4');
  heading.textContent = options.heading || 'Create a new export template';
  container.appendChild(heading);

  const helper = document.createElement('p');
  helper.className = 'modal-helper';
  helper.textContent = 'Name your template, pick a default export format, and add the column headers you want in the export.';
  container.appendChild(helper);

  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.placeholder = 'Template name';
  nameInput.value = currentDraft.name || `Template ${exportTemplates.length + 1}`;
  nameInput.addEventListener('input', (evt) => {
    currentDraft = { ...currentDraft, name: evt.target.value };
    persistTemplateDraftDebounced(currentDraft);
  });
  container.appendChild(nameInput);

  const formatFieldset = document.createElement('fieldset');
  const formatLegend = document.createElement('legend');
  formatLegend.textContent = 'Default format';
  formatFieldset.appendChild(formatLegend);
  const formatOptions = document.createElement('div');
  formatOptions.className = 'range-radios';
  formatFieldset.appendChild(formatOptions);
  const formatInputs = [];
  [
    { value: 'xlsx', label: 'Excel (.xlsx)' },
    { value: 'md', label: 'Markdown (.md)' },
  ].forEach((option) => {
    const label = document.createElement('label');
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'templateFormat';
    radio.value = option.value;
    radio.checked = (currentDraft.format || 'xlsx') === option.value;
    radio.addEventListener('change', () => {
      currentDraft = { ...currentDraft, format: option.value };
      persistTemplateDraftDebounced(currentDraft);
    });
    label.appendChild(radio);
    label.appendChild(document.createTextNode(option.label));
    formatInputs.push(radio);
    formatOptions.appendChild(label);
  });
  container.appendChild(formatFieldset);

  const templateHost = document.createElement('div');
  container.appendChild(templateHost);

  body.appendChild(container);

  const templateEditor = setupTemplateEditor(templateHost, currentDraft.columns || [], {
    onChange: (cols) => {
      currentDraft = { ...currentDraft, columns: cols };
      persistTemplateDraftDebounced(currentDraft);
    },
  });

  const errorEl = document.createElement('div');
  errorEl.className = 'modal-error';
  errorEl.style.display = 'none';
  body.insertBefore(errorEl, container);

  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.textContent = 'Cancel';
  const handleCancel = () => {
    if (typeof options.onCancel === 'function') {
      options.onCancel();
    }
    close && close();
  };
  cancelBtn.addEventListener('click', handleCancel);

  const saveBtn = document.createElement('button');
  saveBtn.type = 'button';
  saveBtn.textContent = 'Save template';
  saveBtn.classList.add('primary');
  const handleSave = async () => {
    errorEl.style.display = 'none';
    errorEl.textContent = '';
    const name = nameInput.value.trim();
    const columns = templateEditor.getColumns();
    const format = formatInputs.find((input) => input.checked)?.value || 'xlsx';
    if (!name) {
      errorEl.textContent = 'Add a template name to continue.';
      errorEl.style.display = 'block';
      return;
    }
    if (!columns.length) {
      errorEl.textContent = 'Add at least one column to continue.';
      errorEl.style.display = 'block';
      return;
    }
    try {
      const saved = await persistExportTemplate({ name, columns, format }, { forceNewId: true });
      currentDraft = getTemplateDraftDefaults();
      if (status) {
        status.innerText = 'Export template saved.';
        status.style.color = '';
      }
      close && close();
      if (typeof options.onSave === 'function') {
        options.onSave(saved);
      }
    } catch (err) {
      errorEl.textContent = err?.message || 'Failed to save template.';
      errorEl.style.display = 'block';
    }
  };
  saveBtn.addEventListener('click', handleSave);

  footer.appendChild(cancelBtn);
  footer.appendChild(saveBtn);
}

function renderTemplateEditModal({ body, footer, close }, template, options = {}) {
  if (!body || !footer) return;

  let handled = false;

  const container = document.createElement('div');
  container.className = 'modal-section';

  const heading = document.createElement('h4');
  heading.textContent = 'Edit export template';
  container.appendChild(heading);

  const helper = document.createElement('p');
  helper.className = 'modal-helper';
  helper.textContent = 'Update the column headers, descriptions, or order. Changes apply to future exports.';
  container.appendChild(helper);

  const targetTemplate = template && template.id ? (findTemplateById(template.id) || template) : getActiveTemplate();
  if (!targetTemplate) {
    const emptyState = document.createElement('p');
    emptyState.className = 'modal-helper';
    emptyState.textContent = 'No template selected. Create one to edit it.';
    container.appendChild(emptyState);
    body.appendChild(container);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = 'Close';
    closeBtn.addEventListener('click', () => close());
    footer.appendChild(closeBtn);
    return;
  }

  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.placeholder = 'Template name';
  nameInput.value = targetTemplate.name || '';
  container.appendChild(nameInput);

  const formatFieldset = document.createElement('fieldset');
  const formatLegend = document.createElement('legend');
  formatLegend.textContent = 'Default format';
  formatFieldset.appendChild(formatLegend);
  const formatOptions = document.createElement('div');
  formatOptions.className = 'range-radios';
  formatFieldset.appendChild(formatOptions);
  const formatInputs = [];
  [
    { value: 'xlsx', label: 'Excel (.xlsx)' },
    { value: 'md', label: 'Markdown (.md)' },
  ].forEach((option) => {
    const label = document.createElement('label');
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'editTemplateFormat';
    radio.value = option.value;
    radio.checked = (targetTemplate.format || 'xlsx') === option.value;
    formatInputs.push(radio);
    label.appendChild(radio);
    label.appendChild(document.createTextNode(option.label));
    formatOptions.appendChild(label);
  });
  container.appendChild(formatFieldset);

  const templateHost = document.createElement('div');
  container.appendChild(templateHost);

  body.appendChild(container);

  const templateEditor = setupTemplateEditor(templateHost, targetTemplate.columns || []);

  const errorEl = document.createElement('div');
  errorEl.className = 'modal-error';
  errorEl.style.display = 'none';
  body.insertBefore(errorEl, container);

  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.textContent = 'Cancel';
  const handleCancel = () => {
    handled = true;
    close();
    if (typeof options.onCancel === 'function') {
      options.onCancel();
    }
  };
  cancelBtn.addEventListener('click', handleCancel);

  const saveBtn = document.createElement('button');
  saveBtn.type = 'button';
  saveBtn.textContent = 'Save changes';
  saveBtn.classList.add('primary');
  const handleSave = async () => {
    errorEl.style.display = 'none';
    errorEl.textContent = '';

    const columns = templateEditor.getColumns();
    if (!columns.length) {
      errorEl.textContent = 'Add at least one column to continue.';
      errorEl.style.display = 'block';
      return;
    }

    try {
      const format = formatInputs.find((input) => input.checked)?.value || targetTemplate.format || 'xlsx';
      const name = nameInput.value.trim() || targetTemplate.name || 'Saved template';
      const saved = await persistExportTemplate({ ...targetTemplate, name, columns, format });
      if (activeExportState) {
        activeExportState.result = null;
        activeExportState.error = '';
        if (activeExportState.templateId && activeExportState.templateId === saved.id) {
          activeExportState.format = saved.format;
        }
      }
      handled = true;
      close();
      if (typeof options.onSave === 'function') {
        options.onSave(saved);
      }
    } catch (err) {
      errorEl.textContent = err?.message || 'Failed to save template.';
      errorEl.style.display = 'block';
    }
  };
  saveBtn.addEventListener('click', handleSave);

  footer.appendChild(cancelBtn);
  footer.appendChild(saveBtn);

  return () => {
    cancelBtn.removeEventListener('click', handleCancel);
    saveBtn.removeEventListener('click', handleSave);
    if (!handled && typeof options.onClose === 'function') {
      options.onClose();
    }
  };
}

function renderExportFlowModal({ body, footer, close }) {
  if (!body || !footer) return;

  if (!historyEntries.length) {
    const emptyMessage = document.createElement('p');
    emptyMessage.className = 'modal-helper';
    emptyMessage.textContent = 'No research history available. Generate a brief first to create exportable data.';
    body.appendChild(emptyMessage);

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = 'Close';
    closeBtn.addEventListener('click', () => close());
    footer.appendChild(closeBtn);
    return;
  }

  const initialState = {
    selectionType: 'all',
    dateFieldPath: defaultDateFieldPath,
    dateFrom: '',
    dateTo: '',
    customSearch: '',
    customPage: 1,
    selectedIds: new Set(),
    format: 'xlsx',
    templateId: selectedTemplateId || (getActiveTemplate()?.id || ''),
    inProgress: false,
    result: null,
    error: '',
  };

  if (!activeExportState) {
    activeExportState = initialState;
  } else {
    activeExportState = {
      ...initialState,
      ...activeExportState,
      selectedIds: activeExportState.selectedIds instanceof Set ? activeExportState.selectedIds : new Set(activeExportState.selectedIds || []),
      templateId: activeExportState.templateId || initialState.templateId,
    };
  }

  const state = activeExportState;
  const persistStateDebounced = debounce(() => persistExportModalState(state), 200);
  const getSelectedTemplate = () => findTemplateById(state.templateId) || getActiveTemplate();
  const ensureSelectedTemplate = () => {
    const found = getSelectedTemplate();
    if (found) {
      state.templateId = found.id;
      setSelectedTemplate(found.id);
      persistTemplatesState({ skipLegacySync: true });
      return found;
    }
    state.templateId = '';
    return null;
  };
  const currentTemplate = ensureSelectedTemplate();
  if (currentTemplate && (!state.format || (state.format !== 'md' && state.format !== 'xlsx'))) {
    state.format = currentTemplate.format || 'xlsx';
  }
  if (!availableDateFields.length) {
    state.selectionType = state.selectionType === 'date' ? 'all' : state.selectionType;
  }
  if (!availableDateFields.find((field) => field.path === state.dateFieldPath)) {
    state.dateFieldPath = availableDateFields[0]?.path || 'createdAt';
  }
  persistStateDebounced();

  body.innerHTML = '';
  footer.innerHTML = '';
  const formatRadioInputs = [];

  const openCreateTemplateModal = () => {
    let reopened = false;
    const reopenExportModal = (savedTemplate) => {
      if (reopened) return;
      reopened = true;
      if (savedTemplate?.id) {
        state.templateId = savedTemplate.id;
        setSelectedTemplate(savedTemplate.id);
        state.format = savedTemplate.format || state.format;
        persistTemplatesState();
      }
      state.result = null;
      state.error = '';
      openModal({ title: 'Export Research', render: renderExportModal });
    };
    openModal({
      title: 'New Export Template',
      render: (ctx) => renderTemplateSetupModal(ctx, {
        onSave: reopenExportModal,
        onCancel: reopenExportModal,
      }),
    });
  };

  const openEditTemplateFromExport = (templateToEdit) => {
    let reopened = false;
    const reopenExportModal = (savedTemplate) => {
      if (reopened) return;
      reopened = true;
      const nextTemplateId = savedTemplate?.id || templateToEdit?.id;
      if (nextTemplateId) {
        state.templateId = nextTemplateId;
        setSelectedTemplate(nextTemplateId);
        if (savedTemplate?.format) {
          state.format = savedTemplate.format;
        }
        persistTemplatesState();
      }
      openModal({ title: 'Export Research', render: renderExportModal });
    };
    openModal({
      title: 'Edit Export Template',
      render: (ctx) => renderTemplateEditModal(ctx, templateToEdit, {
        onSave: reopenExportModal,
        onCancel: reopenExportModal,
        onClose: reopenExportModal,
      }),
    });
  };

  const container = document.createElement('div');
  container.className = 'export-options';
  body.appendChild(container);

  const templateBlock = document.createElement('div');
  templateBlock.className = 'export-section export-section-template';
  const templateHeading = document.createElement('div');
  templateHeading.className = 'export-section-heading';
  const templateStep = document.createElement('span');
  templateStep.className = 'export-step';
  templateStep.textContent = 'Step 1';
  const templateTitleGroup = document.createElement('div');
  templateTitleGroup.className = 'export-section-title-group';
  const templateTitle = document.createElement('div');
  templateTitle.className = 'export-section-title';
  templateTitle.textContent = 'Choose a template';
  const templateHint = document.createElement('p');
  templateHint.className = 'export-section-hint';
  templateHint.textContent = 'Pick a saved template or create a new one to shape the export columns.';
  templateTitleGroup.appendChild(templateTitle);
  templateTitleGroup.appendChild(templateHint);
  templateHeading.appendChild(templateStep);
  templateHeading.appendChild(templateTitleGroup);
  templateBlock.appendChild(templateHeading);

  const templateSection = document.createElement('div');
  templateSection.className = 'export-template-summary';
  templateBlock.appendChild(templateSection);
  container.appendChild(templateBlock);

  const renderTemplateSummary = () => {
    templateSection.innerHTML = '';

    const headerRow = document.createElement('div');
    headerRow.className = 'export-template-summary-header';
    const title = document.createElement('h4');
    title.textContent = 'Saved templates';
    headerRow.appendChild(title);
    const headerActions = document.createElement('div');
    const newBtn = document.createElement('button');
    newBtn.type = 'button';
    newBtn.className = 'ghost';
    newBtn.textContent = 'New template';
    newBtn.addEventListener('click', openCreateTemplateModal);
    headerActions.appendChild(newBtn);
    headerRow.appendChild(headerActions);
    templateSection.appendChild(headerRow);

    if (!exportTemplates.length) {
      const empty = document.createElement('div');
      empty.className = 'template-empty';
      empty.textContent = 'No templates saved yet. Create one to start.';
      templateSection.appendChild(empty);
      return;
    }

    exportTemplates.forEach((template) => {
      const card = document.createElement('div');
      card.className = 'export-template-card export-template-card-selectable';

      const cardHeader = document.createElement('div');
      cardHeader.className = 'export-template-card-header';

      const nameWrap = document.createElement('div');
      nameWrap.className = 'export-template-card-title';
      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'templateChoice';
      radio.value = template.id;
      radio.checked = state.templateId === template.id;
      radio.addEventListener('change', () => {
        state.templateId = template.id;
        setSelectedTemplate(template.id);
        state.format = template.format || state.format;
        state.result = null;
        state.error = '';
        if (notesEl) notesEl.style.display = 'none';
        if (previewWrapper) previewWrapper.innerHTML = '';
        refreshFooterButtons();
        syncFormatRadios();
        persistTemplatesState();
        persistStateDebounced();
      });
      nameWrap.appendChild(radio);
      const name = document.createElement('strong');
      name.textContent = template.name || 'Saved template';
      nameWrap.appendChild(name);
      cardHeader.appendChild(nameWrap);

      const editBtn = document.createElement('button');
      editBtn.type = 'button';
      editBtn.className = 'ghost';
      editBtn.textContent = 'Edit template';
      editBtn.addEventListener('click', (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        openEditTemplateFromExport(template);
      });
      cardHeader.appendChild(editBtn);
      card.appendChild(cardHeader);

      const meta = document.createElement('div');
      meta.className = 'export-template-meta';
      const formatLabel = template.format === 'md' ? 'Markdown (.md)' : 'Excel (.xlsx)';
      const updatedLabel = formatHistoryTimestamp(template.updatedAt) || 'Recently updated';
      meta.textContent = `${formatLabel} • Updated ${updatedLabel}`;
      card.appendChild(meta);

      const columnsList = document.createElement('div');
      columnsList.className = 'export-template-pill-list';
      (template.columns || []).forEach((col) => {
        const pill = document.createElement('span');
        pill.className = 'export-template-pill';
        pill.textContent = col.header || 'Untitled column';
        if (col.description) {
          pill.title = col.description;
        }
        columnsList.appendChild(pill);
      });
      card.appendChild(columnsList);

      templateSection.appendChild(card);
    });
  };

  renderTemplateSummary();

  const rangeBlock = document.createElement('div');
  rangeBlock.className = 'export-section export-section-range';
  const rangeHeading = document.createElement('div');
  rangeHeading.className = 'export-section-heading';
  const rangeStep = document.createElement('span');
  rangeStep.className = 'export-step';
  rangeStep.textContent = 'Step 2';
  const rangeTitleGroup = document.createElement('div');
  rangeTitleGroup.className = 'export-section-title-group';
  const rangeTitle = document.createElement('div');
  rangeTitle.className = 'export-section-title';
  rangeTitle.textContent = 'Choose what to export';
  const rangeHint = document.createElement('p');
  rangeHint.className = 'export-section-hint';
  rangeHint.textContent = 'Send everything, filter by date, or hand-pick specific runs.';
  rangeTitleGroup.appendChild(rangeTitle);
  rangeTitleGroup.appendChild(rangeHint);
  rangeHeading.appendChild(rangeStep);
  rangeHeading.appendChild(rangeTitleGroup);
  rangeBlock.appendChild(rangeHeading);

  const rangeFieldset = document.createElement('fieldset');
  const rangeLegend = document.createElement('legend');
  rangeLegend.textContent = 'Select range';
  rangeFieldset.appendChild(rangeLegend);
  const radioGroup = document.createElement('div');
  radioGroup.className = 'range-radios';
  rangeFieldset.appendChild(radioGroup);

  const rangeOptions = [
    { value: 'all', label: 'All history', disabled: false },
    { value: 'date', label: 'Date range', disabled: !availableDateFields.length },
    { value: 'custom', label: 'Custom selection', disabled: false },
  ];

  rangeOptions.forEach((option) => {
    const radioLabel = document.createElement('label');
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'exportRange';
    radio.value = option.value;
    radio.checked = state.selectionType === option.value;
    radio.disabled = option.disabled;
    radio.addEventListener('change', () => {
      state.selectionType = option.value;
      refreshVisibility();
      refreshFooterButtons();
      if (option.value === 'custom') {
        renderCustomList();
      }
      persistStateDebounced();
    });
    radioLabel.appendChild(radio);
    radioLabel.appendChild(document.createTextNode(option.label));
    radioGroup.appendChild(radioLabel);
  });

  rangeBlock.appendChild(rangeFieldset);

  const dateContainer = document.createElement('div');
  dateContainer.className = 'date-range-inputs';

  if (availableDateFields.length) {
    const dateFieldWrapper = document.createElement('label');
    dateFieldWrapper.textContent = 'Date field';
    const dateSelect = document.createElement('select');
    availableDateFields.forEach((field) => {
      const option = document.createElement('option');
      option.value = field.path;
      option.textContent = field.label;
      dateSelect.appendChild(option);
    });
    dateSelect.value = state.dateFieldPath;
    dateSelect.addEventListener('change', (evt) => {
      state.dateFieldPath = evt.target.value;
      persistStateDebounced();
    });
    dateFieldWrapper.appendChild(dateSelect);
    dateContainer.appendChild(dateFieldWrapper);
  }

  const fromWrapper = document.createElement('label');
  fromWrapper.textContent = 'From';
  const fromInput = document.createElement('input');
  fromInput.type = 'date';
  fromInput.value = state.dateFrom || '';
  fromInput.addEventListener('change', (evt) => {
    state.dateFrom = evt.target.value;
    persistStateDebounced();
  });
  fromWrapper.appendChild(fromInput);
  dateContainer.appendChild(fromWrapper);

  const toWrapper = document.createElement('label');
  toWrapper.textContent = 'To';
  const toInput = document.createElement('input');
  toInput.type = 'date';
  toInput.value = state.dateTo || '';
  toInput.addEventListener('change', (evt) => {
    state.dateTo = evt.target.value;
    persistStateDebounced();
  });
  toWrapper.appendChild(toInput);
  dateContainer.appendChild(toWrapper);

  rangeBlock.appendChild(dateContainer);

  const customContainer = document.createElement('div');
  customContainer.className = 'custom-selection';

  const searchRow = document.createElement('div');
  searchRow.className = 'search-row';
  const searchInput = document.createElement('input');
  searchInput.type = 'search';
  searchInput.placeholder = 'Search company, product, location, or date';
  searchInput.value = state.customSearch || '';
  searchInput.addEventListener('input', (evt) => {
    state.customSearch = evt.target.value;
    state.customPage = 1;
    renderCustomList();
    persistStateDebounced();
  });
  searchRow.appendChild(searchInput);

  const clearSearchBtn = document.createElement('button');
  clearSearchBtn.type = 'button';
  clearSearchBtn.textContent = 'Clear';
  clearSearchBtn.addEventListener('click', () => {
    state.customSearch = '';
    state.customPage = 1;
    searchInput.value = '';
    renderCustomList();
    persistStateDebounced();
  });
  searchRow.appendChild(clearSearchBtn);

  customContainer.appendChild(searchRow);

  const listEl = document.createElement('div');
  listEl.className = 'custom-results';
  customContainer.appendChild(listEl);

  const paginationEl = document.createElement('div');
  paginationEl.className = 'pagination-controls';
  const prevBtn = document.createElement('button');
  prevBtn.type = 'button';
  prevBtn.textContent = 'Prev';
  prevBtn.addEventListener('click', () => {
    if (state.customPage > 1) {
      state.customPage -= 1;
      renderCustomList();
      persistStateDebounced();
    }
  });
  const pageInfo = document.createElement('span');
  const nextBtn = document.createElement('button');
  nextBtn.type = 'button';
  nextBtn.textContent = 'Next';
  nextBtn.addEventListener('click', () => {
    state.customPage += 1;
    renderCustomList();
    persistStateDebounced();
  });
  paginationEl.appendChild(prevBtn);
  paginationEl.appendChild(pageInfo);
  paginationEl.appendChild(nextBtn);
  customContainer.appendChild(paginationEl);

  const customSummary = document.createElement('div');
  customSummary.className = 'export-status';
  customContainer.appendChild(customSummary);

  rangeBlock.appendChild(customContainer);
  container.appendChild(rangeBlock);

  const formatFieldset = document.createElement('fieldset');
  const formatLegend = document.createElement('legend');
  formatLegend.textContent = 'Export format';
  formatFieldset.appendChild(formatLegend);
  const formatRadios = document.createElement('div');
  formatRadios.className = 'range-radios';
  formatFieldset.appendChild(formatRadios);

  [
    { value: 'xlsx', label: 'Excel (.xlsx)' },
    { value: 'md', label: 'Markdown (.md)' },
  ].forEach((option) => {
    const radioLabel = document.createElement('label');
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'exportFormat';
    radio.value = option.value;
    radio.checked = state.format === option.value;
    radio.addEventListener('change', () => {
      state.format = option.value;
      refreshFooterButtons();
      persistStateDebounced();
    });
    formatRadioInputs.push(radio);
    radioLabel.appendChild(radio);
    radioLabel.appendChild(document.createTextNode(option.label));
    formatRadios.appendChild(radioLabel);
  });

  container.appendChild(formatFieldset);

  const statusEl = document.createElement('div');
  statusEl.className = 'export-status';
  container.appendChild(statusEl);

  const notesEl = document.createElement('div');
  notesEl.className = 'notes-box';
  notesEl.style.display = 'none';
  container.appendChild(notesEl);

  const previewWrapper = document.createElement('div');
  container.appendChild(previewWrapper);

  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.textContent = 'Close';
  cancelBtn.addEventListener('click', () => {
    persistStateDebounced();
    close();
  });
  footer.appendChild(cancelBtn);

  const saveAndExitBtn = document.createElement('button');
  saveAndExitBtn.type = 'button';
  saveAndExitBtn.textContent = 'Save and exit';
  saveAndExitBtn.addEventListener('click', async () => {
    if (state.inProgress) return;
    try {
      await persistExportModalState(state);
    } catch (err) {
      console.warn('Failed to save export state before exit', err);
    }
    close();
  });
  footer.appendChild(saveAndExitBtn);

  const exportBtn = document.createElement('button');
  exportBtn.type = 'button';
  exportBtn.classList.add('primary');
  footer.appendChild(exportBtn);

  const downloadBtn = document.createElement('button');
  downloadBtn.type = 'button';
  downloadBtn.textContent = 'Download';
  downloadBtn.disabled = !state.result;
  downloadBtn.addEventListener('click', () => {
    if (state.result?.download) {
      triggerDownload(state.result.download);
    }
  });
  footer.appendChild(downloadBtn);

  function canRunExport() {
    if (!historyEntries.length) return false;
    const template = getSelectedTemplate();
    if (!template || !template.columns || !template.columns.length) return false;
    if (state.selectionType === 'custom' && state.selectedIds.size === 0) return false;
    return true;
  }

  function syncFormatRadios() {
    formatRadioInputs.forEach((input) => {
      input.checked = state.format === input.value;
    });
  }

  function refreshFooterButtons() {
    exportBtn.textContent = state.inProgress ? 'Processing...' : state.result ? 'Run again' : 'Start export';
    exportBtn.disabled = state.inProgress || !canRunExport();
    downloadBtn.disabled = !state.result;
    const downloadLabel = (state.result?.format || state.format) === 'md' ? 'Download Markdown' : 'Download Excel';
    downloadBtn.textContent = downloadLabel;
    saveAndExitBtn.disabled = !!state.inProgress;
  }

  function refreshVisibility() {
    dateContainer.style.display = state.selectionType === 'date' ? 'grid' : 'none';
    customContainer.style.display = state.selectionType === 'custom' ? 'flex' : 'none';
  }

  function formatCustomCandidates() {
    const term = state.customSearch.trim().toLowerCase();
    return historyEntries.map((entry) => {
      const title = entry?.request?.company || 'Untitled brief';
      const subtitle = [entry?.request?.product, entry?.request?.location].filter(Boolean).join(' - ');
      const meta = formatHistoryTimestamp(entry.createdAt) || '';
      const haystack = `${title} ${subtitle} ${meta}`.toLowerCase();
      const matches = !term || haystack.includes(term);
      return matches ? { id: entry.id, title, subtitle, meta } : null;
    }).filter(Boolean);
  }

  function renderCustomList() {
    listEl.innerHTML = '';
    const candidates = formatCustomCandidates();
    const total = candidates.length;
    const totalPages = Math.max(1, Math.ceil(total / EXPORT_PAGE_SIZE));
    if (state.customPage > totalPages) state.customPage = totalPages;
    if (state.customPage < 1) state.customPage = 1;
    const pageItems = paginate(candidates, state.customPage, EXPORT_PAGE_SIZE);

    if (!pageItems.length) {
      const empty = document.createElement('div');
      empty.className = 'template-empty';
      empty.textContent = 'No matches. Adjust the search or try another page.';
      listEl.appendChild(empty);
    } else {
      pageItems.forEach((item) => {
        const wrapper = document.createElement('label');
        wrapper.className = 'custom-result-item';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = state.selectedIds.has(item.id);
        checkbox.addEventListener('change', () => {
          if (checkbox.checked) {
            state.selectedIds.add(item.id);
          } else {
            state.selectedIds.delete(item.id);
          }
          refreshSummary();
          refreshFooterButtons();
          persistStateDebounced();
        });
        wrapper.appendChild(checkbox);

        const metaWrapper = document.createElement('div');
        metaWrapper.className = 'custom-result-meta';
        const titleEl = document.createElement('strong');
        titleEl.textContent = item.title;
        metaWrapper.appendChild(titleEl);
        if (item.subtitle) {
          const subtitleEl = document.createElement('span');
          subtitleEl.textContent = item.subtitle;
          metaWrapper.appendChild(subtitleEl);
        }
        if (item.meta) {
          const metaEl = document.createElement('span');
          metaEl.textContent = item.meta;
          metaWrapper.appendChild(metaEl);
        }
        wrapper.appendChild(metaWrapper);
        listEl.appendChild(wrapper);
      });
    }

    prevBtn.disabled = state.customPage <= 1;
    nextBtn.disabled = state.customPage >= totalPages;
    pageInfo.textContent = `Page ${totalPages ? state.customPage : 0} of ${totalPages}`;
    refreshSummary(total);
    persistStateDebounced();
  }

  function refreshSummary(totalMatches = formatCustomCandidates().length) {
    const selectedCount = state.selectedIds.size;
    const selectionText = selectedCount === 1 ? '1 item selected' : `${selectedCount} items selected`;
    const matchText = state.selectionType === 'custom' ? ` | ${totalMatches} matches` : '';
    customSummary.textContent = `Custom selection${matchText ? matchText : ''}. ${selectionText}.`;
  }

  function setStatus(message, { type = 'info', loading = false } = {}) {
    statusEl.innerHTML = '';
    if (!message) return;
    if (loading) {
      const spinner = document.createElement('span');
      spinner.className = 'spinner';
      statusEl.appendChild(spinner);
      const text = document.createElement('span');
      text.style.marginLeft = '8px';
      text.textContent = message;
      statusEl.appendChild(text);
    } else {
      statusEl.textContent = message;
    }
    statusEl.style.color = type === 'error' ? '#b91c1c' : '#374151';
  }

  function renderPreview() {
    previewWrapper.innerHTML = '';
    notesEl.style.display = 'none';
    if (!state.result?.preview) return;

    const { headers = [], rows = [] } = state.result.preview;
    const totalRows = state.result.totalRows || rows.length;

    const caption = document.createElement('div');
    caption.className = 'modal-helper';
    caption.textContent = `Showing first ${Math.min(rows.length, 10)} of ${totalRows} rows.`;
    previewWrapper.appendChild(caption);

    const table = document.createElement('table');
    table.className = 'preview-table';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    headers.forEach((header) => {
      const th = document.createElement('th');
      th.textContent = header;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    rows.forEach((row) => {
      const tr = document.createElement('tr');
      headers.forEach((header) => {
        const td = document.createElement('td');
        td.textContent = row[header] || '';
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    previewWrapper.appendChild(table);

    if (state.result.notes) {
      notesEl.textContent = state.result.notes;
      notesEl.style.display = 'block';
    }
  }

  async function runExport() {
    state.inProgress = true;
    state.error = '';
    setStatus('Formatting export...', { loading: true });
    refreshFooterButtons();
    notesEl.style.display = 'none';
    previewWrapper.innerHTML = '';

    const template = getSelectedTemplate();
    if (!template || !template.columns || !template.columns.length) {
      state.inProgress = false;
      state.error = 'Add a template with at least one column before exporting.';
      setStatus(state.error, { type: 'error' });
      refreshFooterButtons();
      return;
    }

    const selectionPayload = (() => {
      if (state.selectionType === 'custom') {
        return { type: 'custom', selectedIds: Array.from(state.selectedIds) };
      }
      if (state.selectionType === 'date') {
        return {
          type: 'date',
          dateFieldPath: state.dateFieldPath,
          from: state.dateFrom || null,
          to: state.dateTo || null,
        };
      }
      return { type: 'all' };
    })();

    const response = await new Promise((resolve) => {
      chrome.runtime.sendMessage(
        {
          action: 'exportResearch',
          selection: selectionPayload,
          format: state.format,
          template,
          templateId: template.id,
        },
        (resp) => {
          const err = chrome.runtime.lastError;
          if (err) {
            resolve({ error: err.message || 'Export failed.' });
            return;
          }
          resolve(resp);
        }
      );
    });

    state.inProgress = false;

    if (!response || response.error) {
      state.error = response?.error || 'Export failed.';
      state.result = null;
      setStatus(`Error: ${state.error}`, { type: 'error' });
      refreshFooterButtons();
      persistStateDebounced();
      return;
    }

    state.result = {
      preview: response.preview,
      download: response.download,
      notes: response.notes,
      totalRows: response.totalRows,
      format: response.download?.format || state.format,
    };
    state.error = '';

    setStatus('Export ready. Review the preview and download the file when ready.', { type: 'info' });
    renderPreview();
    refreshFooterButtons();
    persistStateDebounced();
  }

  exportBtn.addEventListener('click', () => {
    if (state.inProgress || !canRunExport()) return;
    runExport();
  });

  refreshVisibility();
  renderCustomList();
  refreshFooterButtons();
  if (state.result) {
    renderPreview();
    setStatus('Previous export preview is available below.', { type: 'info' });
  }

  return () => {
    state.inProgress = false;
    persistStateDebounced();
  };
}

function renderExportModal(context) {
  renderExportFlowModal(context);
}

function clearAndRenderPersonas(personas) {
  while (personasDiv.firstChild) personasDiv.removeChild(personasDiv.firstChild);

  if (!personas || !personas.length) {
    const p = document.createElement('p');
    p.textContent = 'No personas generated.';
    personasDiv.appendChild(p);
    return;
  }

  personas.forEach(p => {
    const wrapper = document.createElement('div');
    wrapper.className = 'persona';

    const title = document.createElement('div');
    const nameEl = document.createElement('strong');
    nameEl.textContent = p.name || '';
    title.appendChild(nameEl);

    if (p.designation) {
      const des = document.createTextNode(' — ' + p.designation);
      title.appendChild(des);
    }
    if (p.department) {
      const dept = document.createTextNode(' (' + p.department + ')');
      title.appendChild(dept);
    }
    wrapper.appendChild(title);

    const rawLink = p.zoominfo_link || p.zoomInfo || p.zoominfo || p.zoom || '';
    const link = normalizeUrl(rawLink);

    if (link) {
      const linkWrap = document.createElement('div');
      const a = document.createElement('a');
      a.href = link;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.textContent = 'Google Search';
      linkWrap.appendChild(a);
      wrapper.appendChild(linkWrap);
    }

    personasDiv.appendChild(wrapper);
  });
}

function mergePersonaEmails(personas = [], personaEmails = [], personaPitches = []) {
  const maxLen = Math.max(personas.length, personaEmails.length, personaPitches.length);
  if (!maxLen) return [];

  const merged = [];
  for (let i = 0; i < maxLen; i += 1) {
    const persona = personas[i] || {};
    const email = personaEmails[i] || {};
    const personaName = email.personaName || persona.name || `Persona ${i + 1}`;
    const personaDesignation = email.personaDesignation || persona.designation || '';
    const personaDepartment = email.personaDepartment || persona.department || '';
    const subject = email.subject || (persona.email && persona.email.subject) || '';
    const body = email.body || (persona.email && persona.email.body) || '';
    let telePitchEntry = personaPitches[i] || null;
    if ((!telePitchEntry || typeof telePitchEntry !== 'object' || !Object.keys(telePitchEntry).length) && personaName) {
      const target = personaName.toLowerCase();
      telePitchEntry = personaPitches.find((pitch) => {
        const matchName = pitch?.personaName || pitch?.persona_name || '';
        return typeof matchName === 'string' && matchName.toLowerCase() === target;
      }) || null;
    }
    const telePitch = formatTelephonicPitchText(telePitchEntry || {});

    merged.push({
      personaName,
      personaDesignation,
      personaDepartment,
      subject,
      body,
      telePitch,
    });
  }

  return merged;
}

function formatEmailDraftText(draft = {}) {
  const subject = typeof draft.subject === 'string' ? draft.subject.trim() : '';
  const body = typeof draft.body === 'string' ? draft.body.trim() : '';
  return [subject, body].filter(Boolean).join('\n\n');
}

function formatTelephonicPitchText(pitch = {}) {
  if (!pitch || typeof pitch !== 'object') return '';
  const sections = [];
  const callGoal = typeof pitch.callGoal === 'string' ? pitch.callGoal.trim() : (typeof pitch.call_goal === 'string' ? pitch.call_goal.trim() : '');
  if (callGoal) sections.push(`Call Goal: ${callGoal}`);
  const opener = typeof pitch.opener === 'string' ? pitch.opener.trim() : '';
  if (opener) sections.push(`Opener: ${opener}`);
  const discovery = typeof pitch.discoveryQuestion === 'string'
    ? pitch.discoveryQuestion.trim()
    : (typeof pitch.discovery_question === 'string' ? pitch.discovery_question.trim() : '');
  if (discovery) sections.push(`Discovery: ${discovery}`);
  const valueStatement = typeof pitch.valueStatement === 'string'
    ? pitch.valueStatement.trim()
    : (typeof pitch.value_statement === 'string' ? pitch.value_statement.trim() : '');
  if (valueStatement) sections.push(`Value Statement: ${valueStatement}`);
  const proof = typeof pitch.proofPoint === 'string'
    ? pitch.proofPoint.trim()
    : (typeof pitch.proof_point === 'string' ? pitch.proof_point.trim() : '');
  if (proof) sections.push(`Proof Point: ${proof}`);
  const cta = typeof pitch.cta === 'string'
    ? pitch.cta.trim()
    : (typeof pitch.closingPrompt === 'string'
      ? pitch.closingPrompt.trim()
      : (typeof pitch.closing_prompt === 'string' ? pitch.closing_prompt.trim() : ''));
  if (cta) sections.push(`CTA: ${cta}`);
  const script = typeof pitch.script === 'string'
    ? pitch.script.trim()
    : (typeof pitch.full_pitch === 'string' ? pitch.full_pitch.trim() : '');
  if (script) {
    if (sections.length) sections.push('');
    sections.push(script);
  }
  return sections.join('\n\n').trim();
}

function formatFallbackEmailText(email) {
  if (!email) return '';
  if (typeof email === 'string') return email;
  if (typeof email === 'object') {
    const subject = typeof email.subject === 'string' ? email.subject.trim() : '';
    const body = typeof email.body === 'string' ? email.body.trim() : '';
    return [subject, body].filter(Boolean).join('\n\n');
  }
  return '';
}

function updateCopyEmailButtonState(text = '') {
  if (!copyEmailBtn) return;
  const hasText = typeof text === 'string' && text.trim().length > 0;
  copyEmailBtn.disabled = !hasText;
}

function getActivePersonaName() {
  if (selectedPersonaIndex < 0) return '';
  const draft = personaEmailDrafts[selectedPersonaIndex];
  return draft && draft.personaName ? draft.personaName : '';
}

function activatePersonaTab(index) {
  if (!personaTabs || !personaEmailDrafts.length) return;
  const safeIndex = Math.max(0, Math.min(index, personaEmailDrafts.length - 1));
  selectedPersonaIndex = safeIndex;

  const buttons = personaTabs.querySelectorAll('.persona-tab');
  buttons.forEach((btn, btnIdx) => {
    const isActive = btnIdx === safeIndex;
    btn.classList.toggle('active', isActive);
    btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
    btn.tabIndex = isActive ? 0 : -1;
  });

  const draft = personaEmailDrafts[safeIndex] || {};
  const text = formatEmailDraftText(draft);
  if (text) {
    emailOut.innerText = text;
    updateCopyEmailButtonState(text);
  } else {
    emailOut.innerText = 'No email available for this persona yet.';
    updateCopyEmailButtonState('');
  }

  if (telePitchOut) {
    if (telephonicPitchErrorMessage) {
      setTelePitchOutput(`Telephonic pitch failed: ${telephonicPitchErrorMessage}`, { isError: true });
    } else {
      const pitchText = typeof draft.telePitch === 'string' ? draft.telePitch.trim() : '';
      setTelePitchOutput(pitchText || 'No telephonic pitch available for this persona yet.');
    }
  }
}

function renderPersonaEmailDrafts(personasData = [], personaEmailsData = [], telephonicPitchesData = [], fallbackEmail) {
  personaEmailDrafts = mergePersonaEmails(personasData, personaEmailsData, telephonicPitchesData);

  if (personaTabs) {
    personaTabs.innerHTML = '';
  }

  if (personaEmailDrafts.length && personaTabs) {
    personaTabs.style.display = '';

    personaEmailDrafts.forEach((draft, idx) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'persona-tab';
      btn.dataset.index = String(idx);
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', idx === 0 ? 'true' : 'false');
      btn.tabIndex = idx === 0 ? 0 : -1;

      const labelBits = [];
      const name = draft.personaName && draft.personaName.trim() ? draft.personaName.trim() : `Persona ${idx + 1}`;
      labelBits.push(name);
      if (draft.personaDesignation) labelBits.push(draft.personaDesignation);

      btn.textContent = labelBits.join(' \u2013 ');
      personaTabs.appendChild(btn);
    });

    activatePersonaTab(0);
  } else {
    if (personaTabs) personaTabs.style.display = 'none';
    personaEmailDrafts = [];
    selectedPersonaIndex = -1;
    const fallbackText = formatFallbackEmailText(fallbackEmail);
    if (fallbackText) {
      emailOut.innerText = fallbackText;
      updateCopyEmailButtonState(fallbackText);
    } else {
      emailOut.innerText = 'No email generated yet.';
      updateCopyEmailButtonState('');
    }
    if (telephonicPitchErrorMessage) {
      setTelePitchOutput(`Telephonic pitch failed: ${telephonicPitchErrorMessage}`, { isError: true });
    } else {
      setTelePitchOutput('No telephonic pitch generated yet.');
    }
  }
}

function renderResultView(data = {}) {
  if (!resultDiv || !briefDiv || !emailOut) return;

  resultDiv.style.display = 'block';
  mergeBriefData(data);
  renderBriefFromState();
  renderPersonaSectionsFromState();

  if (typeof requestAnimationFrame === 'function') {
    requestAnimationFrame(() => {
      resultDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  } else {
    resultDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function formatHistoryTimestamp(isoString) {
  if (!isoString) return '';
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) return '';
  return date.toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
}

function renderHistory(entries, opts = {}) {
  if (!historyList || !historyEmpty) return null;

  const toSortValue = (entry) => {
    if (!entry) return 0;
    const ts = entry.createdAt ? new Date(entry.createdAt).getTime() : NaN;
    if (!Number.isNaN(ts)) return ts;
    const idNum = Number(entry.id);
    return Number.isNaN(idNum) ? 0 : idNum;
  };

  historyEntries = Array.isArray(entries) ? [...entries] : [];
  historyEntries.sort((a, b) => toSortValue(b) - toSortValue(a));

  if (opts.selectEntryId) {
    currentHistoryId = opts.selectEntryId;
  } else if (opts.selectLatest && historyEntries.length) {
    currentHistoryId = historyEntries[0].id;
  } else if (currentHistoryId && !historyEntries.some(item => item.id === currentHistoryId)) {
    currentHistoryId = historyEntries.length ? historyEntries[0].id : null;
  }

  const hasEntries = historyEntries.length > 0;
  historyEmpty.style.display = hasEntries ? 'none' : 'block';

  historyList.innerHTML = '';
  if (!hasEntries) return currentHistoryId;

  historyEntries.forEach((entry) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'history-item';
    btn.setAttribute('role', 'listitem');
    btn.dataset.id = entry.id;
    if (entry.id === currentHistoryId) btn.classList.add('active');

    const title = document.createElement('span');
    title.className = 'history-title';
    title.textContent = entry?.request?.company || 'Untitled brief';
    btn.appendChild(title);

    const subtitleText = [entry?.request?.product, entry?.request?.location].filter(Boolean).join(' - ');
    if (subtitleText) {
      const subtitle = document.createElement('span');
      subtitle.className = 'history-subtitle';
      subtitle.textContent = subtitleText;
      btn.appendChild(subtitle);
    }

    const meta = document.createElement('span');
    meta.className = 'history-meta';
    meta.textContent = formatHistoryTimestamp(entry.createdAt) || '';
    btn.appendChild(meta);

    historyList.appendChild(btn);
  });

  return currentHistoryId;
}

function setActiveHistoryItem(id) {
  currentHistoryId = id;
  if (!historyList) return;
  const buttons = historyList.querySelectorAll('.history-item');
  buttons.forEach(btn => {
    btn.classList.toggle('active', btn.dataset.id === id);
  });
}

function showHistoryEntry(entry, options = {}) {
  if (!entry) return;
  const opts = { updateForm: true, statusText: 'Loaded previous brief.', ...options };

  if (opts.updateForm) {
    if (companyEl) companyEl.value = entry?.request?.company || '';
    if (locationEl) locationEl.value = entry?.request?.location || '';
    if (productEl) productEl.value = entry?.request?.product || '';
  }

  if (entry.result) {
    renderResultView(entry.result);
  }

  if (status) {
    const teleError = typeof entry?.result?.telephonicPitchError === 'string'
      ? entry.result.telephonicPitchError.trim()
      : '';
    if (teleError) {
      status.innerText = 'Telephonic pitch failed: ' + teleError;
      status.style.color = '#b91c1c';
    } else if (typeof opts.statusText === 'string') {
      status.innerText = opts.statusText;
      status.style.color = '';
    }
  }
}

function loadHistory(options = {}) {
  const opts = { autoShow: false, selectLatest: false, updateForm: true, statusText: 'Loaded previous brief.', ...options };
  chrome.runtime.sendMessage({ action: 'getResearchHistory' }, (resp) => {
    const err = chrome.runtime.lastError;
    if (err) {
      console.warn('Failed to load history', err);
      return;
    }
    const entries = resp && Array.isArray(resp.history) ? resp.history : [];
    availableDateFields = discoverDateFields(entries);
    defaultDateFieldPath = availableDateFields[0]?.path || 'createdAt';
    if (activeExportState) {
      if (!availableDateFields.find((field) => field.path === activeExportState.dateFieldPath)) {
        activeExportState.dateFieldPath = defaultDateFieldPath;
      }
    }
    const selectedId = renderHistory(entries, opts);

    if (opts.autoShow && selectedId) {
      const entry = historyEntries.find(item => item.id === selectedId);
      if (entry) {
        showHistoryEntry(entry, { updateForm: opts.updateForm, statusText: opts.statusText });
        setActiveHistoryItem(selectedId);
      }
    } else if (typeof selectedId === 'string') {
      setActiveHistoryItem(selectedId);
    }
  });
}

function renderTargetHistory(entries, opts = {}) {
  if (!targetHistoryList || !targetHistoryEmpty) return null;

  const toSortValue = (entry) => {
    if (!entry) return 0;
    const ts = entry.createdAt ? new Date(entry.createdAt).getTime() : NaN;
    if (!Number.isNaN(ts)) return ts;
    const idNum = Number(entry.id);
    return Number.isNaN(idNum) ? 0 : idNum;
  };

  targetHistoryEntries = Array.isArray(entries) ? [...entries] : [];
  targetHistoryEntries.sort((a, b) => toSortValue(b) - toSortValue(a));

  if (opts.selectEntryId) {
    currentTargetHistoryId = opts.selectEntryId;
  } else if (opts.selectLatest && targetHistoryEntries.length) {
    currentTargetHistoryId = targetHistoryEntries[0].id;
  } else if (currentTargetHistoryId && !targetHistoryEntries.some(item => item.id === currentTargetHistoryId)) {
    currentTargetHistoryId = targetHistoryEntries.length ? targetHistoryEntries[0].id : null;
  }

  const hasEntries = targetHistoryEntries.length > 0;
  targetHistoryEmpty.style.display = hasEntries ? 'none' : 'block';

  targetHistoryList.innerHTML = '';
  if (!hasEntries) return currentTargetHistoryId;

  targetHistoryEntries.forEach((entry) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'history-item';
    btn.setAttribute('role', 'listitem');
    btn.dataset.id = entry.id;
    if (entry.id === currentTargetHistoryId) btn.classList.add('active');

    const title = document.createElement('span');
    title.className = 'history-title';
    title.textContent = entry?.request?.product || 'Untitled search';
    btn.appendChild(title);

    const subtitleParts = [];
    if (entry?.request?.location) subtitleParts.push(entry.request.location);
    if (Array.isArray(entry?.result?.companies)) {
      subtitleParts.push(`${entry.result.companies.length} companies`);
    }
    if (subtitleParts.length) {
      const subtitle = document.createElement('span');
      subtitle.className = 'history-subtitle';
      subtitle.textContent = subtitleParts.join(' • ');
      btn.appendChild(subtitle);
    }

    const meta = document.createElement('span');
    meta.className = 'history-meta';
    meta.textContent = formatHistoryTimestamp(entry.createdAt) || '';
    btn.appendChild(meta);

    targetHistoryList.appendChild(btn);
  });

  return currentTargetHistoryId;
}

function setActiveTargetHistoryItem(id) {
  currentTargetHistoryId = id;
  if (!targetHistoryList) return;
  const buttons = targetHistoryList.querySelectorAll('.history-item');
  buttons.forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.id === id);
  });
}

function showTargetHistoryEntry(entry, options = {}) {
  if (!entry) return;
  const opts = { updateForm: true, statusText: 'Loaded previous target search.', ...options };

  if (opts.updateForm) {
    if (targetProductInput) targetProductInput.value = entry?.request?.product || '';
    if (targetLocationInput) targetLocationInput.value = entry?.request?.location || '';
    if (targetDocInput) targetDocInput.value = '';
    if (targetSectorsInput) targetSectorsInput.value = '';
    setTargetSectors(entry?.request?.sectors || []);
  }
  setSelectedDocsFromRefs(Mode.TARGET_GENERATION, entry?.request?.docs || []);

  const companies = Array.isArray(entry?.result?.companies) ? entry.result.companies : [];
  renderTargetResults(companies);

  if (typeof opts.statusText === 'string') {
    setTargetStatus(opts.statusText, { error: false });
  }
}

function loadTargetHistory(options = {}) {
  const opts = { autoShow: false, selectLatest: false, updateForm: true, statusText: 'Loaded previous target search.', ...options };
  chrome.runtime.sendMessage({ action: 'getTargetHistory' }, (resp) => {
    const err = chrome.runtime.lastError;
    if (err) {
      console.warn('Failed to load target history', err);
      return;
    }
    const entries = resp && Array.isArray(resp.history) ? resp.history : [];
    const selectedId = renderTargetHistory(entries, opts);

    if (opts.autoShow && selectedId) {
      const entry = targetHistoryEntries.find(item => item.id === selectedId);
      if (entry) {
        showTargetHistoryEntry(entry, { updateForm: opts.updateForm, statusText: opts.statusText });
        setActiveTargetHistoryItem(selectedId);
      }
    } else if (typeof selectedId === 'string') {
      setActiveTargetHistoryItem(selectedId);
    }
  });
}
// ---- End helpers ----

generateBtn?.addEventListener('click', async () => {
  const company = companyEl.value.trim();
  const location = locationEl.value.trim();
  const product = productEl.value.trim();
  if (!company || !product) {
    status.innerText = 'Company and Product required';
    status.style.color = '#b91c1c';
    return;
  }
  const runId = String(Date.now());
  startBriefProgress(runId, 'Loading docs...');
  prepareResultShell();
  currentHistoryId = null;

  try {
    await loadStoredDocs();
    let docs = getSelectedDocObjects(Mode.TARGETED_BRIEF);
    if (!docs.length) {
      try {
        const resp = await sendMessagePromise({ action: 'getDocsForProduct', product });
        docs = resp && Array.isArray(resp.docs) ? resp.docs : [];
        if (docs.length) {
          setSelectedDocsFromRefs(Mode.TARGETED_BRIEF, docs);
        }
      } catch (docErr) {
        resetBriefProgress();
        status.innerText = 'Failed to load docs: ' + (docErr?.message || docErr);
        status.style.color = '#b91c1c';
        return;
      }
    }

    startBriefProgress(runId, 'Generating brief...', 3);
    const result = await sendMessagePromise({ action: 'generateBrief', company, location, product, docs, runId });
    if (!result) {
      resetBriefProgress();
      status.innerText = 'Generation failed';
      status.style.color = '#b91c1c';
      return;
    }
    if (result.error) {
      resetBriefProgress();
      status.innerText = 'Error: ' + result.error;
      status.style.color = '#b91c1c';
      return;
    }
    renderResultView(result);
    const overviewError = typeof result.overview_error === 'string' ? result.overview_error.trim() : '';
    const personaError = typeof result.persona_error === 'string' ? result.persona_error.trim() : '';
    const teleError = typeof result.telephonicPitchError === 'string' ? result.telephonicPitchError.trim() : '';
    if (overviewError || personaError || teleError) {
      const parts = [];
      if (overviewError) parts.push('Overview: ' + overviewError);
      if (personaError) parts.push('Personas: ' + personaError);
      if (teleError) parts.push('Telephonic pitch: ' + teleError);
      status.innerText = parts.join(' | ');
      status.style.color = '#b91c1c';
    } else {
      status.innerText = 'Done.';
      status.style.color = '';
    }
    resetBriefProgress();
    loadHistory({ selectLatest: true, autoShow: false, updateForm: false, statusText: '' });
  } catch (err) {
    resetBriefProgress();
    status.innerText = 'Error: ' + (err?.message || err);
    status.style.color = '#b91c1c';
  }
});

historyList?.addEventListener('click', (evt) => {
  const button = evt.target.closest('.history-item');
  if (!button) return;
  const { id } = button.dataset;
  if (!id) return;
  const entry = historyEntries.find(item => item.id === id);
  if (!entry) return;
  evt.preventDefault();
  setActiveHistoryItem(id);
  showHistoryEntry(entry, { updateForm: true });
});

loadHistory({ selectLatest: true, autoShow: true, statusText: '' });

targetHistoryList?.addEventListener('click', (evt) => {
  const button = evt.target.closest('.history-item');
  if (!button) return;
  const { id } = button.dataset;
  if (!id) return;
  const entry = targetHistoryEntries.find(item => item.id === id);
  if (!entry) return;
  evt.preventDefault();
  setActiveTargetHistoryItem(id);
  showTargetHistoryEntry(entry, { updateForm: true });
});

loadTargetHistory({ selectLatest: true, autoShow: true, statusText: '' });



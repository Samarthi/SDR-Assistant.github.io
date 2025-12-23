const companyEl = document.getElementById('company');
const locationEl = document.getElementById('location');
const productEl = document.getElementById('product');
const fileInput = document.getElementById('fileInput');
const saveKeyBtn = document.getElementById('saveKey');
const newResearchBtn = document.getElementById('newResearch');
const apiKeyInput = document.getElementById('apiKey');
const generateBtn = document.getElementById('generate');
const status = document.getElementById('status');
const resultDiv = document.getElementById('result');
const briefSectionsToggle = document.getElementById('briefSectionsToggle');
const briefSectionsButton = document.getElementById('briefSectionsButton');
const briefDiv = document.getElementById('brief');
const briefSummarySection = document.getElementById('briefSummarySection');
const briefHqValue = document.getElementById('briefHqValue');
const briefRevenueValue = document.getElementById('briefRevenueValue');
const briefIndustryValue = document.getElementById('briefIndustryValue');
const topNewsSection = document.getElementById('topNewsSection');
const topNewsList = document.getElementById('topNewsList');
const topNewsHint = document.getElementById('topNewsHint');
const legacyBriefContainer = document.getElementById('legacyBriefContainer');
const personasDiv = document.getElementById('personas');
const emailOut = document.getElementById('emailOut');
const telePitchOut = document.getElementById('telePitchOut');
const copyEmailBtn = document.getElementById('copyEmail');
const personaTabs = document.getElementById('personaTabs');
const telePersonaTabs = document.getElementById('telePersonaTabs');
const reviseEmailPersonaBtn = document.getElementById('reviseEmailPersona');
const reviseEmailAllBtn = document.getElementById('reviseEmailAll');
const revisePitchPersonaBtn = document.getElementById('revisePitchPersona');
const revisePitchAllBtn = document.getElementById('revisePitchAll');
const emailVersionControls = document.getElementById('emailVersionControls');
const pitchVersionControls = document.getElementById('pitchVersionControls');
const viewDocs = document.getElementById('viewDocs');
const historyList = document.getElementById('historyList');
const historyEmpty = document.getElementById('historyEmpty');
const targetHistoryList = document.getElementById('targetHistoryList');
const targetHistoryEmpty = document.getElementById('targetHistoryEmpty');
const historySidebarTitle = document.querySelector('.history-scroll-header');
const historySearchBtn = document.getElementById('historySearch');
const historySearchLabel = historySearchBtn?.querySelector('.history-action-label');
const navRail = document.getElementById('navRail');
const historyCollapseBtn = document.getElementById('historyCollapse');
const railLogoBtn = document.getElementById('railLogo');
const railNewResearchBtn = document.getElementById('railNewResearch');
const railHistoryBtn = document.getElementById('railHistory');
const exportTrigger = document.getElementById('exportTrigger');
const settingsTriggers = Array.from(document.querySelectorAll('[data-settings-trigger]'));
const heroBriefBtn = document.getElementById('heroBriefBtn');
const heroTargetsBtn = document.getElementById('heroTargetsBtn');
const heroDocsBtn = document.getElementById('heroDocsBtn');
const heroExportBtn = document.getElementById('heroExportBtn');
const heroSettingsBtn = document.getElementById('heroSettingsBtn');
const modalRoot = document.getElementById('modalRoot');
const modalTitle = document.getElementById('modalTitle');
const modalBody = document.getElementById('modalBody');
const modalFooter = document.getElementById('modalFooter');
const modalClose = document.getElementById('modalClose');
const modalContainer = document.querySelector('.modal-container');
const modalHeader = modalRoot?.querySelector('.modal-header');
const modalMaximizeBtn = modalHeader ? (() => {
  const btn = document.createElement('button');
  btn.id = 'modalMaximize';
  btn.type = 'button';
  btn.className = 'ghost modal-icon-button';
  btn.title = 'Toggle full height';
  btn.setAttribute('aria-label', 'Maximize modal');
  btn.setAttribute('aria-pressed', 'false');
  setModalButtonIcon(btn, 'maximize-2', 'Max');
  modalHeader.insertBefore(btn, modalClose);
  return btn;
})() : null;
const modeTabs = document.querySelectorAll('[data-mode-tab]');
const modeViews = document.querySelectorAll('[data-mode-view]');
const historySections = document.querySelectorAll('[data-history-section]');
const targetForm = document.getElementById('targetForm');
const targetProductInput = document.getElementById('targetProduct');
const targetDocInput = document.getElementById('targetDocInput');
const targetSectorsInput = document.getElementById('targetSectorsInput');
const targetSectorsChips = document.getElementById('targetSectorsChips');
const briefFormPane = document.getElementById('formPane');
const briefCommandBar = document.getElementById('briefCommandBar');
const briefCommandBody = document.getElementById('briefCommandBody');
const briefBarToggle = document.getElementById('briefBarToggle');
const briefBarStatus = document.getElementById('briefBarStatus');
const briefChipCompany = document.getElementById('briefChipCompany');
const briefChipLocation = document.getElementById('briefChipLocation');
const briefChipProduct = document.getElementById('briefChipProduct');
const briefChipDocs = document.getElementById('briefChipDocs');
const briefChipSections = document.getElementById('briefChipSections');
const briefEditToggle = document.getElementById('briefEditToggle');
const briefRegenerateBtn = document.getElementById('briefRegenerate');
const briefVersionPill = document.getElementById('briefVersionPill');
const briefAutosavePill = document.getElementById('briefAutosavePill');
const briefEditDrawer = document.getElementById('briefEditDrawer');
const briefVersionSelect = document.getElementById('briefVersionSelect');
const targetLocationInput = document.getElementById('targetLocation');
const targetStatusEl = document.getElementById('targetStatus');
const targetResultsSection = document.getElementById('targetResults');
const targetResultsList = document.getElementById('targetResultsList');
const copyTargetsBtn = document.getElementById('copyTargets');
const briefDocPickerBtn = document.getElementById('briefDocPicker');
const targetDocPickerBtn = document.getElementById('targetDocPicker');
const briefDocChips = document.getElementById('briefDocChips');
const targetDocChips = document.getElementById('targetDocChips');
const targetedBriefView = document.querySelector('.mode-view[data-mode-view="targetedBrief"]');

const EXPORT_TEMPLATE_STORAGE_KEY = 'exportTemplate';
const EXPORT_TEMPLATE_DRAFT_STORAGE_KEY = 'exportTemplateDraft';
const EXPORT_TEMPLATES_STORAGE_KEY = 'exportTemplates';
const EXPORT_MODAL_STATE_STORAGE_KEY = 'exportModalState';
const EXPORT_PAGE_SIZE = 8;
const PITCH_FROM_COMPANY_KEY = 'pitchFromCompany';
const HISTORY_SIDEBAR_COLLAPSED_KEY = 'historySidebarCollapsed';
const SIDEBAR_ICON_HTML = '<span class="sidebar-icon" aria-hidden="true"></span>';
const LLMProvider = {
  GEMINI: 'gemini',
  GROQ: 'groq',
};
const LLM_PROVIDER_STORAGE_KEY = 'llmProvider';
const LLM_MODEL_STORAGE_KEY = 'llmModel';
const GROQ_KEY_STORAGE_KEY = 'groqKey';
const GROQ_COMPOUND_MINI_MODEL = 'groq/compound-mini';
const DEFAULT_LLM_MODELS = {
  [LLMProvider.GEMINI]: 'gemini-2.5-flash',
  [LLMProvider.GROQ]: GROQ_COMPOUND_MINI_MODEL,
};
const GROQ_LEGACY_MODEL = 'gpt-oss-20b';

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
const BRIEF_MODULES_STORAGE_KEY = 'briefModules';
const BriefModule = {
  OVERVIEW: 'overview',
  TOP_NEWS: 'topNews',
  PERSONAS: 'personas',
  EMAILS: 'emails',
  TELE_PITCH: 'telePitch',
};
const BRIEF_MODULES = [
  {
    id: BriefModule.OVERVIEW,
    label: 'Company overview',
    description: 'HQ, revenue, and industry sector',
  },
  {
    id: BriefModule.TOP_NEWS,
    label: 'Top news',
    description: 'Recent headlines and summaries',
  },
  {
    id: BriefModule.PERSONAS,
    label: 'Personas',
    description: 'Buyer personas and search links',
  },
  {
    id: BriefModule.EMAILS,
    label: 'Generated emails',
    description: 'Email drafts per persona',
    dependsOn: [BriefModule.PERSONAS],
  },
  {
    id: BriefModule.TELE_PITCH,
    label: 'Telephonic pitch',
    description: 'Call scripts per persona',
    dependsOn: [BriefModule.PERSONAS],
  },
];
const BRIEF_MODULE_ICONS = {
  [BriefModule.OVERVIEW]:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="8" height="16" rx="2"></rect><rect x="13" y="4" width="8" height="9" rx="2"></rect><path d="M13 15h8M13 19h4"></path></svg>',
  [BriefModule.TOP_NEWS]:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V5z"></path><path d="M16 3v2"></path><path d="M8 9h10"></path><path d="M8 13h8M8 17h6"></path></svg>',
  [BriefModule.PERSONAS]:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>',
  [BriefModule.EMAILS]:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2z"></path><path d="m22 7-10 7L2 7"></path></svg>',
  [BriefModule.TELE_PITCH]:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.12.9.35 1.78.69 2.62a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.46-1.26a2 2 0 0 1 2.11-.45c.84.34 1.72.57 2.62.69a2 2 0 0 1 1.72 2.03z"></path></svg>',
};
const systemThemeMedia = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;

let currentBriefRequest = {
  company: '',
  location: '',
  product: '',
  docs: [],
  modules: [],
};
let briefVersionCounter = 0;
let briefVersions = [];
let briefUiMode = 'form';
let briefBarCollapsed = false;
let historyEntries = [];
let currentHistoryId = null;
let targetHistoryEntries = [];
let currentTargetHistoryId = null;
let personaEmailDrafts = [];
let personaEmailVersions = [];
let telephonicPitchVersions = [];
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
let revisionModalState = null;
const docLabelCache = new Map();
let hasLaunchedSettingsOnboarding = false;
const selectedDocsByMode = {
  [Mode.TARGETED_BRIEF]: [],
  [Mode.TARGET_GENERATION]: [],
};
const briefProgressState = { runId: null, total: 0, current: 0 };
let themePreference = ThemePreference.SYSTEM;
let cachedGeminiKey = '';
let cachedGroqKey = '';
let cachedLlmProvider = LLMProvider.GEMINI;
let cachedLlmModel = DEFAULT_LLM_MODELS[LLMProvider.GEMINI];
let cachedPitchingCompany = '';
let briefModuleConfig = null;
let activeBriefModules = null;
let briefModuleOverride = null;
let briefSectionsListActive = null;
let briefSectionsListActiveHandler = null;

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

function getDefaultBriefModules() {
  return BRIEF_MODULES.reduce((acc, module) => {
    acc[module.id] = true;
    return acc;
  }, {});
}

function normalizeBriefModules(value) {
  const defaults = getDefaultBriefModules();
  if (!value || typeof value !== 'object') {
    return { ...defaults };
  }
  const next = { ...defaults };
  Object.keys(defaults).forEach((key) => {
    if (typeof value[key] === 'boolean') {
      next[key] = value[key];
    }
  });
  return next;
}

function resolveBriefModules(value) {
  const next = normalizeBriefModules(value);
  BRIEF_MODULES.forEach((module) => {
    if (next[module.id] && Array.isArray(module.dependsOn)) {
      module.dependsOn.forEach((dep) => {
        next[dep] = true;
      });
    }
  });
  return next;
}

function getEffectiveBriefModules() {
  return resolveBriefModules(briefModuleOverride || briefModuleConfig || null);
}

function getBriefModuleDefinition(id) {
  return BRIEF_MODULES.find((module) => module.id === id);
}

function computeBriefRequiredBy(modules) {
  const requiredBy = {};
  BRIEF_MODULES.forEach((module) => {
    if (modules[module.id] && Array.isArray(module.dependsOn)) {
      module.dependsOn.forEach((dep) => {
        if (!requiredBy[dep]) requiredBy[dep] = [];
        requiredBy[dep].push(module.id);
      });
    }
  });
  return requiredBy;
}

function renderBriefModuleControls(modules, container, options = {}) {
  const host = container || briefSectionsListActive;
  if (!host) return;
  host.innerHTML = '';
  const resolved = resolveBriefModules(modules);
  const requiredBy = computeBriefRequiredBy(resolved);
  const onToggle = typeof options.onToggle === 'function'
    ? options.onToggle
    : (briefSectionsListActive === host && typeof briefSectionsListActiveHandler === 'function'
      ? briefSectionsListActiveHandler
      : handleBriefModuleToggle);

  BRIEF_MODULES.forEach((module) => {
    const wrapper = document.createElement('label');
    wrapper.className = 'brief-module-option';
    const isRequired = Array.isArray(requiredBy[module.id]) && requiredBy[module.id].length > 0;
    const isChecked = !!resolved[module.id];
    if (isRequired) wrapper.classList.add('is-required');
    if (isChecked) wrapper.classList.add('is-active');

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.dataset.moduleId = module.id;
    checkbox.checked = isChecked;
    checkbox.disabled = isRequired;
    checkbox.className = 'visually-hidden';
    checkbox.addEventListener('change', (evt) => {
      wrapper.classList.toggle('is-active', evt.target.checked);
      onToggle(module.id, evt.target.checked);
    });

    const icon = document.createElement('div');
    icon.className = 'brief-module-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.innerHTML = BRIEF_MODULE_ICONS[module.id] || BRIEF_MODULE_ICONS[BriefModule.OVERVIEW];

    const copy = document.createElement('div');
    copy.className = 'brief-module-copy';
    const title = document.createElement('div');
    title.className = 'brief-module-title';
    title.textContent = module.label;
    const desc = document.createElement('div');
    desc.className = 'brief-module-desc';
    desc.textContent = module.description;
    copy.appendChild(title);
    copy.appendChild(desc);

    if (isRequired) {
      const required = document.createElement('div');
      required.className = 'brief-module-required';
      const labels = requiredBy[module.id]
        .map((depId) => getBriefModuleDefinition(depId)?.label || depId);
      required.textContent = `Required for ${labels.join(' and ')}`;
      copy.appendChild(required);
    }

    const toggle = document.createElement('div');
    toggle.className = 'brief-module-toggle';
    toggle.setAttribute('aria-hidden', 'true');

    wrapper.appendChild(checkbox);
    wrapper.appendChild(icon);
    wrapper.appendChild(copy);
    wrapper.appendChild(toggle);
    host.appendChild(wrapper);
  });
}

function applyBriefModulesToView(modules, options = {}) {
  const resolved = resolveBriefModules(modules);
  activeBriefModules = resolved;
  const containers = document.querySelectorAll('[data-brief-module]');
  containers.forEach((container) => {
    const id = container.dataset.briefModule;
    const isEnabled = id ? !!resolved[id] : true;
    if (isEnabled) {
      container.removeAttribute('hidden');
      container.setAttribute('aria-hidden', 'false');
    } else {
      container.setAttribute('hidden', 'hidden');
      container.setAttribute('aria-hidden', 'true');
    }
  });
  if (options.updateControls !== false) {
    renderBriefModuleControls(resolved);
  }
  return resolved;
}

function applyEffectiveBriefModulesToView(options = {}) {
  return applyBriefModulesToView(getEffectiveBriefModules(), options);
}

function setBriefModulesConfig(nextConfig, options = {}) {
  const resolved = resolveBriefModules(nextConfig);
  briefModuleConfig = resolved;
  if (options.persist !== false) {
    chrome.storage.local
      .set({ [BRIEF_MODULES_STORAGE_KEY]: resolved })
      .catch((err) => console.warn('Failed to save brief module settings', err));
  }
  const updateControls = options.updateControls !== false;
  applyEffectiveBriefModulesToView({ updateControls });
  return resolved;
}

function getBriefProgressTotal(modules) {
  const resolved = resolveBriefModules(modules);
  return BRIEF_MODULES.reduce((sum, module) => sum + (resolved[module.id] ? 1 : 0), 0);
}

function setBriefModulesOverride(nextConfig, options = {}) {
  const resolved = resolveBriefModules(nextConfig);
  briefModuleOverride = resolved;
  applyEffectiveBriefModulesToView({ updateControls: options.updateControls !== false });
  return resolved;
}

function clearBriefModulesOverride(options = {}) {
  briefModuleOverride = null;
  applyEffectiveBriefModulesToView({ updateControls: options.updateControls !== false });
}

function handleBriefModuleToggle(moduleId, checked) {
  if (!moduleId) return;
  const base = briefModuleConfig || activeBriefModules || getDefaultBriefModules();
  const next = { ...base, [moduleId]: !!checked };
  setBriefModulesConfig(next);
}

function openBriefSectionsModal(options = {}) {
  const mode = options.mode === 'global' ? 'global' : 'override';
  const isGlobal = mode === 'global';
  const baseModules = options.baseModules || null;
  openModal({
    title: isGlobal ? 'Global brief sections' : 'Brief sections',
    render: ({ body, footer, close }) => {
      const panel = document.createElement('div');
      panel.className = 'brief-sections-panel';
      const hint = document.createElement('p');
      hint.className = 'brief-sections-hint';
      hint.textContent = isGlobal
        ? 'Set the default sections for new briefs.'
        : 'Override global defaults for this brief. Disabled sections will not run.';
      panel.appendChild(hint);

      const list = document.createElement('div');
      list.className = 'brief-sections-list';
      panel.appendChild(list);
      if (!isGlobal) {
        const status = document.createElement('div');
        status.className = 'brief-sections-status';
        status.textContent = briefModuleOverride ? 'Override active for this brief' : 'Using global defaults';
        panel.insertBefore(status, list);
      }
      body.appendChild(panel);

      briefSectionsListActive = list;
      const resolveBase = () => (isGlobal
        ? (briefModuleConfig || getDefaultBriefModules())
        : (briefModuleOverride || baseModules || briefModuleConfig || getDefaultBriefModules()));
      const applyAndRender = (nextConfig) => {
        if (isGlobal) {
          const resolved = setBriefModulesConfig(nextConfig, { updateControls: false });
          renderBriefModuleControls(resolved, list, { onToggle: handleToggle });
        } else {
          const resolved = setBriefModulesOverride(nextConfig, { updateControls: false });
          renderBriefModuleControls(resolved, list, { onToggle: handleToggle });
        }
      };
      const handleToggle = (moduleId, checked) => {
        const base = resolveBase();
        const next = { ...base, [moduleId]: !!checked };
        applyAndRender(next);
        if (!isGlobal) {
          const status = panel.querySelector('.brief-sections-status');
          if (status) {
            status.textContent = 'Override active for this brief';
          }
        }
      };
      briefSectionsListActiveHandler = handleToggle;
      renderBriefModuleControls(resolveBase(), list, { onToggle: handleToggle });

      if (!isGlobal) {
        const clearBtn = document.createElement('button');
        clearBtn.type = 'button';
        clearBtn.textContent = 'Use global defaults';
        clearBtn.addEventListener('click', () => {
          clearBriefModulesOverride({ updateControls: false });
          renderBriefModuleControls(briefModuleConfig || getDefaultBriefModules(), list, { onToggle: handleToggle });
          const status = panel.querySelector('.brief-sections-status');
          if (status) {
            status.textContent = 'Using global defaults';
          }
        });
        footer.appendChild(clearBtn);
      }

      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = 'Done';
      closeBtn.classList.add('primary');
      closeBtn.addEventListener('click', () => close());
      footer.appendChild(closeBtn);

      return () => {
        if (briefSectionsListActive === list) {
          briefSectionsListActive = null;
          briefSectionsListActiveHandler = null;
        }
      };
    },
  });
}

function normalizeLlmProvider(provider, hasGeminiKey, hasGroqKey) {
  if (provider === LLMProvider.GEMINI || provider === LLMProvider.GROQ) return provider;
  if (hasGroqKey) return LLMProvider.GROQ;
  if (hasGeminiKey) return LLMProvider.GEMINI;
  return LLMProvider.GROQ;
}

function normalizeLlmModel(model, provider) {
  if (typeof model === 'string' && model.trim()) return model.trim();
  return DEFAULT_LLM_MODELS[provider] || DEFAULT_LLM_MODELS[LLMProvider.GEMINI];
}

function coerceModelForProvider(provider, model) {
  const fallback = DEFAULT_LLM_MODELS[provider] || DEFAULT_LLM_MODELS[LLMProvider.GEMINI];
  if (!model || typeof model !== 'string') return fallback;
  const normalized = model.trim();
  if (provider === LLMProvider.GEMINI) {
    return normalized.toLowerCase().startsWith('gemini') ? normalized : fallback;
  }
  if (provider === LLMProvider.GROQ) {
    const lower = normalized.toLowerCase();
    const isLegacyGroq20b = lower.includes(GROQ_LEGACY_MODEL);
    // Fix misordered/unknown Groq ids like "openai/oss-gpt-20b" by upgrading to primary.
    if (isLegacyGroq20b || lower.includes('oss-gpt')) {
      return DEFAULT_LLM_MODELS[LLMProvider.GROQ];
    }
    return lower.startsWith('gemini') ? fallback : normalized;
  }
  return fallback;
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

function updateThemeAssets(resolvedTheme = themePreference) {
  const isDark = resolvedTheme === ThemePreference.DARK;
  document.querySelectorAll('[data-theme-logo]').forEach((img) => {
    if (!(img instanceof HTMLImageElement)) return;
    const lightSrc = img.getAttribute('data-logo-light') || 'logo-light.png';
    const darkSrc = img.getAttribute('data-logo-dark') || 'logo-dark.png';
    img.src = isDark ? darkSrc : lightSrc;
  });
  const favicon = document.querySelector('link[rel="icon"]');
  if (favicon) {
    favicon.href = isDark ? 'logo-dark.png' : 'logo-light.png';
  }
}

function applyTheme(preference = themePreference) {
  themePreference = normalizeThemePreference(preference);
  const resolved = resolveTheme(themePreference);
  if (document?.body) {
    document.body.classList.toggle('theme-dark', resolved === ThemePreference.DARK);
    document.body.classList.toggle('theme-light', resolved === ThemePreference.LIGHT);
  }
  updateThemeAssets(resolved);
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
const llmSettingsLoadPromise = chrome.storage.local
  .get([LLM_PROVIDER_STORAGE_KEY, LLM_MODEL_STORAGE_KEY, 'geminiKey', GROQ_KEY_STORAGE_KEY])
  .then((data) => {
    const geminiKey = data && typeof data.geminiKey === 'string' ? data.geminiKey : '';
    const groqKey = data && typeof data[GROQ_KEY_STORAGE_KEY] === 'string' ? data[GROQ_KEY_STORAGE_KEY] : '';
    const provider = normalizeLlmProvider(data && data[LLM_PROVIDER_STORAGE_KEY], !!geminiKey, !!groqKey);
    let model = coerceModelForProvider(provider, normalizeLlmModel(data && data[LLM_MODEL_STORAGE_KEY], provider));
    if (
      provider === LLMProvider.GROQ &&
      typeof model === 'string' &&
      model.toLowerCase().includes(GROQ_LEGACY_MODEL)
    ) {
      model = DEFAULT_LLM_MODELS[LLMProvider.GROQ];
    }
    cachedGeminiKey = geminiKey;
    cachedGroqKey = groqKey;
    cachedLlmProvider = provider;
    cachedLlmModel = model;
    if (apiKeyInput) {
      apiKeyInput.value = provider === LLMProvider.GROQ ? groqKey : geminiKey;
    }
    return { provider, model, geminiKey, groqKey };
  })
  .catch(() => ({
    provider: LLMProvider.GEMINI,
    model: DEFAULT_LLM_MODELS[LLMProvider.GEMINI],
    geminiKey: '',
    groqKey: '',
  }));
const briefModulesLoadPromise = chrome.storage.local
  .get([BRIEF_MODULES_STORAGE_KEY])
  .then((data) => {
    const stored = data && data[BRIEF_MODULES_STORAGE_KEY];
    briefModuleConfig = resolveBriefModules(stored);
    applyEffectiveBriefModulesToView({ updateControls: true });
    return briefModuleConfig;
  })
  .catch(() => {
    briefModuleConfig = resolveBriefModules(null);
    applyEffectiveBriefModulesToView({ updateControls: true });
    return briefModuleConfig;
  });
const pitchingCompanyLoadPromise = chrome.storage.local
  .get([PITCH_FROM_COMPANY_KEY])
  .then((data) => {
    const val = data && typeof data[PITCH_FROM_COMPANY_KEY] === 'string' ? data[PITCH_FROM_COMPANY_KEY] : '';
    cachedPitchingCompany = val;
    return val;
  })
  .catch(() => '');

function updateHistorySidebarTitle(mode) {
  if (!historySidebarTitle) return;
  historySidebarTitle.textContent = mode === Mode.TARGET_GENERATION ? 'Targets' : 'Briefs';
}

function updateHistorySearchLabel(mode) {
  if (!historySearchLabel) return;
  const isTargetMode = mode === Mode.TARGET_GENERATION;
  const labelText = isTargetMode ? 'Search Targets' : 'Search Briefs';
  historySearchLabel.textContent = labelText;
  if (historySearchBtn) {
    historySearchBtn.setAttribute('aria-label', labelText);
  }
}

function setActiveMode(mode) {
  const validModes = Object.values(Mode);
  const normalizedMode = validModes.includes(mode) ? mode : Mode.TARGETED_BRIEF;
  const modeChanged = activeMode !== normalizedMode;
  activeMode = normalizedMode;
  if (normalizedMode === Mode.TARGETED_BRIEF) {
    setBriefUiMode(briefUiMode || 'form');
  }
  updateHistorySidebarTitle(normalizedMode);
  updateHistorySearchLabel(normalizedMode);
  if (!modeChanged) {
    return;
  }

  Array.from(modeTabs).forEach((tab) => {
    const isActive = tab.dataset.modeTab === normalizedMode;
    tab.classList.toggle('active', isActive);
    tab.setAttribute('aria-selected', isActive ? 'true' : 'false');
    tab.removeAttribute('tabindex');
  });

  Array.from(modeViews).forEach((view) => {
    const isActive = view.dataset.modeView === normalizedMode;
    if (isActive) {
      view.removeAttribute('hidden');
      view.setAttribute('aria-hidden', 'false');
    } else {
      view.setAttribute('hidden', 'hidden');
      view.setAttribute('aria-hidden', 'true');
    }
  });

  Array.from(historySections).forEach((section) => {
    const isActive = section.dataset.historySection === normalizedMode;
    if (isActive) {
      section.removeAttribute('hidden');
      section.setAttribute('aria-hidden', 'false');
    } else {
      section.setAttribute('hidden', 'hidden');
      section.setAttribute('aria-hidden', 'true');
    }
  });
}

function scrollToElementSmooth(el) {
  if (!el || typeof el.scrollIntoView !== 'function') return;
  window.requestAnimationFrame(() => {
    try {
      el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (err) {
      el.scrollIntoView();
    }
  });
}

function focusMode(mode, targetEl) {
  setActiveMode(mode);
  if (targetEl) {
    scrollToElementSmooth(targetEl);
  }
}

function normalizeModeFromString(str = '') {
  const v = str.toLowerCase();
  if (v.includes('target') && v.includes('brief')) return Mode.TARGETED_BRIEF;
  if (v === 'brief') return Mode.TARGETED_BRIEF;
  if (v === 'targets' || v === 'targetgeneration' || v === 'target_generation') return Mode.TARGET_GENERATION;
  if (v.includes('target') && v.includes('gen')) return Mode.TARGET_GENERATION;
  return null;
}

function getInitialModeFromUrl() {
  try {
    const url = new URL(window.location.href);
    const qsMode = normalizeModeFromString(url.searchParams.get('mode') || '');
    if (qsMode) return qsMode;
    const hashMode = normalizeModeFromString((url.hash || '').replace('#', ''));
    if (hashMode) return hashMode;
  } catch (err) {
    // ignore parse errors
  }
  return null;
}

function getOnboardingTargetFromUrl() {
  try {
    const url = new URL(window.location.href);
    const onboarding = (url.searchParams.get('onboarding') || '').toLowerCase();
    if (onboarding === 'settings') return 'settings';
  } catch (err) {
    // ignore parse errors
  }
  return null;
}

function clearOnboardingFlagFromUrl() {
  try {
    const url = new URL(window.location.href);
    if (!url.searchParams.has('onboarding')) return;
    url.searchParams.delete('onboarding');
    const nextUrl = url.pathname + (url.searchParams.toString() ? `?${url.searchParams.toString()}` : '') + url.hash;
    window.history.replaceState({}, document.title, nextUrl);
  } catch (err) {
    // ignore parse errors
  }
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
    removeBtn.textContent = '-';
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

function setTelePitchOutput(message, { isError = false, allowMarkdown = false } = {}) {
  if (!telePitchOut) return;
  const text = typeof message === 'string' ? message : '';
  if (allowMarkdown && !isError) {
    telePitchOut.innerHTML = markdownToHtml(text);
  } else {
    telePitchOut.textContent = text;
  }
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
  updateRevisionButtonsState();
});
briefSectionsToggle?.addEventListener('click', () => {
  openBriefSectionsModal({ mode: 'override', baseModules: activeBriefData?.modules });
});
briefSectionsButton?.addEventListener('click', () => {
  openBriefSectionsModal({ mode: 'override' });
});

function resetBriefProgress() {
  briefProgressState.runId = null;
  briefProgressState.total = 0;
  briefProgressState.current = 0;
}

function setBriefEditDrawerCollapsed(collapsed) {
  if (!briefEditDrawer) return;
  if (collapsed) {
    briefEditDrawer.setAttribute('data-collapsed', 'true');
    briefEditDrawer.hidden = true;
  } else {
    briefEditDrawer.removeAttribute('data-collapsed');
    briefEditDrawer.hidden = false;
  }
}

function getActiveBriefStatusEl() {
  if (briefCommandBar && !briefCommandBar.hidden && briefBarStatus) {
    return briefBarStatus;
  }
  return status;
}

function setBriefStatusMessage(text = '', { error = false } = {}) {
  const target = getActiveBriefStatusEl();
  if (target) {
    target.innerText = text;
    target.style.color = error ? '#b91c1c' : '';
  }
  if (status && target !== status) {
    status.innerText = text;
    status.style.color = error ? '#b91c1c' : '';
  }
}

function updateBriefVersionPill() {
  if (!briefVersionPill) return;
  if (!briefVersionCounter) {
    briefVersionPill.hidden = true;
    return;
  }
  briefVersionPill.textContent = `v${briefVersionCounter}`;
  briefVersionPill.hidden = false;
}

function updateBriefVersionSelect() {
  if (!briefVersionSelect) return;
  const options = ['<option value=\"\">Current run</option>'];
  briefVersions.forEach((version) => {
    const label = `${version.id} — ${version.timestamp}`;
    options.push(`<option value=\"${version.id}\">${label}</option>`);
  });
  briefVersionSelect.innerHTML = options.join('');
}

function updateBriefCommandBarFromRequest(request = currentBriefRequest) {
  if (!briefCommandBar) return;
  const req = request || {};
  const docCount = Array.isArray(req.docs) ? req.docs.length : 0;
  const modulesCount = Array.isArray(req.modules) ? req.modules.length : 0;
  const moduleLabel = modulesCount ? `${modulesCount} section${modulesCount === 1 ? '' : 's'}` : 'Default';
  const assignValue = (el, val) => {
    if (el) el.textContent = val || 'Not set';
  };
  assignValue(briefChipCompany?.querySelector('.chip-value'), req.company);
  assignValue(briefChipLocation?.querySelector('.chip-value'), req.location);
  assignValue(briefChipProduct?.querySelector('.chip-value'), req.product);
  assignValue(briefChipDocs?.querySelector('.chip-value'), String(docCount));
  assignValue(briefChipSections?.querySelector('.chip-value'), moduleLabel);
}

function setBriefBarCollapsed(collapsed) {
  briefBarCollapsed = !!collapsed;
  if (briefCommandBar) {
    briefCommandBar.classList.toggle('is-collapsed', briefBarCollapsed);
  }
  if (briefCommandBody) {
    briefCommandBody.hidden = briefBarCollapsed;
  }
  if (briefBarToggle) {
    const label = briefBarCollapsed ? 'Expand input bar' : 'Collapse input bar';
    briefBarToggle.setAttribute('aria-expanded', briefBarCollapsed ? 'false' : 'true');
    briefBarToggle.setAttribute('aria-label', label);
    briefBarToggle.title = label;
  }
}

function toggleBriefBarCollapsed() {
  setBriefBarCollapsed(!briefBarCollapsed);
}

function showBriefCommandBar(show) {
  if (!briefCommandBar) return;
  briefCommandBar.hidden = !show;
  setBriefBarCollapsed(briefBarCollapsed);
  if (show) {
    setBriefEditDrawerCollapsed(true);
    if (briefEditToggle) briefEditToggle.setAttribute('aria-expanded', 'false');
  }
}

function setBriefUiMode(mode) {
  briefUiMode = mode === 'bar' ? 'bar' : 'form';
  const isBar = briefUiMode === 'bar';
  if (targetedBriefView) {
    targetedBriefView.classList.toggle('brief-ui-bar', isBar);
    targetedBriefView.classList.toggle('brief-ui-form', !isBar);
  }
  if (briefCommandBar) briefCommandBar.hidden = !isBar;
  setBriefBarCollapsed(briefBarCollapsed);
  setBriefEditDrawerCollapsed(isBar);
  if (briefEditToggle) briefEditToggle.setAttribute('aria-expanded', isBar ? 'false' : 'true');
}

function showBriefForm() {
  setBriefUiMode('form');
}

function showBriefBar() {
  setBriefUiMode('bar');
}

function toggleBriefEditDrawer() {
  const formVisible = briefEditDrawer && !briefEditDrawer.hidden;
  if (formVisible) {
    showBriefBar();
  } else {
    showBriefForm();
  }
}

function buildBriefRequestFromForm() {
  return {
    company: companyEl?.value?.trim() || '',
    location: locationEl?.value?.trim() || '',
    product: productEl?.value?.trim() || '',
    docs: getSelectedDocObjects(Mode.TARGETED_BRIEF),
    modules: getEffectiveBriefModules(),
  };
}

async function runBriefGeneration(request) {
  const req = request || {};
  const company = (req.company || '').trim();
  const location = (req.location || '').trim();
  const product = (req.product || '').trim();
  const modules = Array.isArray(req.modules) && req.modules.length ? req.modules : getEffectiveBriefModules();
  if (!company || !product) {
    setBriefStatusMessage('Company and Product required', { error: true });
    setBriefEditDrawerCollapsed(false);
    return;
  }
  const totalSteps = getBriefProgressTotal(modules);
  if (!totalSteps) {
    setBriefStatusMessage('Enable at least one brief section to generate.', { error: true });
    return;
  }
  showBriefBar();
  currentBriefRequest = { ...req, company, location, product, modules };
  updateBriefCommandBarFromRequest();
  const runId = String(Date.now());
  startBriefProgress(runId, 'Loading docs...', totalSteps);
  prepareResultShell(modules);
  currentHistoryId = null;

  try {
    await loadStoredDocs();
    let docs = Array.isArray(req.docs) && req.docs.length ? req.docs : getSelectedDocObjects(Mode.TARGETED_BRIEF);
    if (!docs.length) {
      try {
        const resp = await sendMessagePromise({ action: 'getDocsForProduct', product });
        docs = resp && Array.isArray(resp.docs) ? resp.docs : [];
        if (docs.length) {
          setSelectedDocsFromRefs(Mode.TARGETED_BRIEF, docs);
        }
      } catch (docErr) {
        resetBriefProgress();
        setBriefStatusMessage('Failed to load docs: ' + (docErr?.message || docErr), { error: true });
        return;
      }
    }
    currentBriefRequest.docs = docs;
    updateBriefCommandBarFromRequest();

    startBriefProgress(runId, 'Generating brief...', totalSteps);
    const result = await sendMessagePromise({
      action: 'generateBrief',
      company,
      location,
      product,
      docs,
      runId,
      modules,
    });
    if (!result) {
      resetBriefProgress();
      setBriefStatusMessage('Generation failed', { error: true });
      return;
    }
    if (result.error) {
      resetBriefProgress();
      setBriefStatusMessage('Error: ' + result.error, { error: true });
      return;
    }
    renderResultView(result);
    briefVersionCounter += 1;
    const versionId = `v${briefVersionCounter}`;
    briefVersions = [
      { id: versionId, request: { ...currentBriefRequest }, result, timestamp: new Date().toLocaleString([], { hour: '2-digit', minute: '2-digit', month: 'short', day: 'numeric' }) },
      ...briefVersions.filter((v) => v.id !== versionId),
    ];
    updateBriefVersionPill();
    updateBriefVersionSelect();
    const overviewError = typeof result.overview_error === 'string' ? result.overview_error.trim() : '';
    const personaError = typeof result.persona_error === 'string' ? result.persona_error.trim() : '';
    const teleError = typeof result.telephonicPitchError === 'string' ? result.telephonicPitchError.trim() : '';
    if (overviewError || personaError || teleError) {
      const parts = [];
      if (overviewError) parts.push('Overview: ' + overviewError);
      if (personaError) parts.push('Personas: ' + personaError);
      if (teleError) parts.push('Telephonic pitch: ' + teleError);
      setBriefStatusMessage(parts.join(' | '), { error: true });
    } else {
      setBriefStatusMessage('Done.', { error: false });
    }
    resetBriefProgress();
    loadHistory({ selectLatest: true, autoShow: false, updateForm: false, statusText: '' });
  } catch (err) {
    resetBriefProgress();
    setBriefStatusMessage('Error: ' + (err?.message || err), { error: true });
  }
}

function createEmptyBriefData() {
  return {
    brief_html: '',
    modules: getEffectiveBriefModules(),
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
    telephonicPitchVersions: [],
    telephonicPitchVersionIndexes: [],
    personaEmailVersions: [],
    personaEmailVersionIndexes: [],
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

  assign('modules');
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
  if (Array.isArray(partial.personaEmailVersions)) next.personaEmailVersions = partial.personaEmailVersions;
  if (Array.isArray(partial.personaEmailVersionIndexes)) next.personaEmailVersionIndexes = partial.personaEmailVersionIndexes;
  if (Array.isArray(partial.telephonicPitches)) next.telephonicPitches = partial.telephonicPitches;
  if (Array.isArray(partial.telephonicPitchAttempts)) next.telephonicPitchAttempts = partial.telephonicPitchAttempts;
  if (Array.isArray(partial.telephonicPitchVersions)) next.telephonicPitchVersions = partial.telephonicPitchVersions;
  if (Array.isArray(partial.telephonicPitchVersionIndexes)) next.telephonicPitchVersionIndexes = partial.telephonicPitchVersionIndexes;
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

function buildVersionEntry(baseDraft = {}) {
  return {
    versions: [baseDraft],
    activeIndex: 0,
  };
}

function hasContent(obj) {
  if (!obj) return false;
  if (typeof obj === 'string') return obj.trim().length > 0;
  return typeof obj === 'object' && Object.keys(obj).length > 0;
}

function clampIndex(idx, arr) {
  if (!Array.isArray(arr) || !arr.length) return -1;
  return Math.max(0, Math.min(idx, arr.length - 1));
}

function isLatestVersion(entry) {
  if (!entry || !Array.isArray(entry.versions) || !entry.versions.length) return true;
  const active = typeof entry.activeIndex === 'number' ? entry.activeIndex : 0;
  return active >= entry.versions.length - 1;
}

function syncVersionState() {
  if (!activeBriefData) return;
  const personas = Array.isArray(activeBriefData.personas) ? activeBriefData.personas : [];

  // Emails
  const personaEmails = Array.isArray(activeBriefData.personaEmails) ? activeBriefData.personaEmails : [];
  if (!Array.isArray(activeBriefData.personaEmailVersions)) activeBriefData.personaEmailVersions = [];
  const emailVersions = activeBriefData.personaEmailVersions;
  const emailVersionIndexes = Array.isArray(activeBriefData.personaEmailVersionIndexes) ? activeBriefData.personaEmailVersionIndexes : [];
  for (let i = 0; i < Math.max(personaEmails.length, personas.length, emailVersions.length); i += 1) {
    if (!emailVersions[i]) {
      const base = personaEmails[i] || {};
      emailVersions[i] = buildVersionEntry(base);
    } else if (hasContent(personaEmails[i])) {
      const activeIdx = typeof emailVersions[i].activeIndex === 'number' ? emailVersions[i].activeIndex : 0;
      const currentDraft = emailVersions[i].versions?.[activeIdx] || {};
      if (!hasContent(currentDraft)) {
        emailVersions[i] = buildVersionEntry(personaEmails[i]);
      }
    } else if (!Array.isArray(emailVersions[i].versions) || !emailVersions[i].versions.length) {
      emailVersions[i].versions = [personaEmails[i] || {}];
      emailVersions[i].activeIndex = 0;
    }
    if (emailVersionIndexes[i] !== undefined) {
      emailVersions[i].activeIndex = emailVersionIndexes[i];
    }
    emailVersions[i].activeIndex = clampIndex(
      typeof emailVersions[i].activeIndex === 'number' ? emailVersions[i].activeIndex : 0,
      emailVersions[i].versions
    );
    const activeDraft = emailVersions[i].versions[emailVersions[i].activeIndex] || {};
    personaEmails[i] = activeDraft;
  }
  activeBriefData.personaEmails = personaEmails;
  personaEmailVersions = emailVersions;
  activeBriefData.personaEmailVersionIndexes = personaEmailVersions.map((v) => v.activeIndex);

  // Telephonic pitches
  const telePitches = Array.isArray(activeBriefData.telephonicPitches) ? activeBriefData.telephonicPitches : [];
  if (!Array.isArray(activeBriefData.telephonicPitchVersions)) activeBriefData.telephonicPitchVersions = [];
  const pitchVersions = activeBriefData.telephonicPitchVersions;
  const pitchVersionIndexes = Array.isArray(activeBriefData.telephonicPitchVersionIndexes) ? activeBriefData.telephonicPitchVersionIndexes : [];
  for (let i = 0; i < Math.max(telePitches.length, personas.length, pitchVersions.length); i += 1) {
    if (!pitchVersions[i]) {
      const base = telePitches[i] || {};
      pitchVersions[i] = buildVersionEntry(base);
    } else if (hasContent(telePitches[i])) {
      const activeIdx = typeof pitchVersions[i].activeIndex === 'number' ? pitchVersions[i].activeIndex : 0;
      const currentDraft = pitchVersions[i].versions?.[activeIdx] || {};
      if (!hasContent(currentDraft)) {
        pitchVersions[i] = buildVersionEntry(telePitches[i]);
      }
    } else if (!Array.isArray(pitchVersions[i].versions) || !pitchVersions[i].versions.length) {
      pitchVersions[i].versions = [telePitches[i] || {}];
      pitchVersions[i].activeIndex = 0;
    }
    if (pitchVersionIndexes[i] !== undefined) {
      pitchVersions[i].activeIndex = pitchVersionIndexes[i];
    }
    pitchVersions[i].activeIndex = clampIndex(
      typeof pitchVersions[i].activeIndex === 'number' ? pitchVersions[i].activeIndex : 0,
      pitchVersions[i].versions
    );
    const activePitch = pitchVersions[i].versions[pitchVersions[i].activeIndex] || {};
    telePitches[i] = activePitch;
  }
  activeBriefData.telephonicPitches = telePitches;
  telephonicPitchVersions = pitchVersions;
  activeBriefData.telephonicPitchVersionIndexes = telephonicPitchVersions.map((v) => v.activeIndex);
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
    const topNewsModule = topNewsSection?.closest('[data-brief-module="topNews"]');
    if (topNewsModule) {
      if (shouldShowLegacy) {
        topNewsModule.setAttribute('data-legacy-hidden', 'true');
      } else {
        topNewsModule.removeAttribute('data-legacy-hidden');
      }
    }
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

  syncVersionState();

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

function prepareResultShell(modules) {
  const resolvedModules = applyBriefModulesToView(
    modules || getEffectiveBriefModules(),
    { updateControls: true },
  );
  const topNewsModule = topNewsSection?.closest('[data-brief-module="topNews"]');
  topNewsModule?.removeAttribute('data-legacy-hidden');
  activeBriefData = createEmptyBriefData();
  activeBriefData.modules = resolvedModules;
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
  personaEmailVersions = [];
  telephonicPitchVersions = [];
  selectedPersonaIndex = -1;
  if (emailVersionControls) emailVersionControls.innerHTML = '';
  if (pitchVersionControls) pitchVersionControls.innerHTML = '';
  resetPersonaTabContainers();
  setEmailOutput('Generating email drafts...', { copyTextOverride: '' });
  telephonicPitchErrorMessage = '';
  telephonicPitchDebugAttempts = [];
  setTelePitchOutput('Generating telephonic pitches...');
  updateRevisionButtonsState();
}

function renderBriefProgress(label = 'Working...') {
  const target = getActiveBriefStatusEl();
  if (!target) return;
  target.innerHTML = '';
  const spinner = document.createElement('span');
  spinner.className = 'spinner';
  target.appendChild(spinner);
  const text = document.createElement('span');
  text.style.marginLeft = '8px';
  const total = briefProgressState.total || '?';
  const current = typeof briefProgressState.current === 'number' ? briefProgressState.current : 0;
  text.textContent = `${label} (${current}/${total})`;
  target.appendChild(text);
  target.style.color = '';
  if (status && target !== status) {
    status.innerHTML = target.innerHTML;
    status.style.color = '';
  }
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
  const urlPreferredMode = getInitialModeFromUrl();
  const initialMode = urlPreferredMode || initialTab?.dataset.modeTab || Mode.TARGETED_BRIEF;
  setActiveMode(initialMode);
}

const onboardingTarget = getOnboardingTargetFromUrl();
const launchSettingsOnboarding = () => {
  if (hasLaunchedSettingsOnboarding) return;
  hasLaunchedSettingsOnboarding = true;
  openSettingsModal({
    onSave: () => {
      clearOnboardingFlagFromUrl();
      focusMode(Mode.TARGETED_BRIEF, briefFormPane);
    },
  });
};
Promise.all([llmSettingsLoadPromise, pitchingCompanyLoadPromise])
  .then(([llmSettings, storedPitch]) => {
    const hasKey = Boolean(llmSettings && (llmSettings.geminiKey || llmSettings.groqKey));
    const hasPitch = Boolean(storedPitch);
    if (onboardingTarget === 'settings') {
      if (hasKey && hasPitch) {
        clearOnboardingFlagFromUrl();
        return;
      }
      launchSettingsOnboarding();
      return;
    }
    if (!hasKey) {
      launchSettingsOnboarding();
    }
  })
  .catch(() => {
    // even if storage read fails, try to open settings once to unblock the user
    launchSettingsOnboarding();
  });

heroBriefBtn?.addEventListener('click', () => {
  focusMode(Mode.TARGETED_BRIEF, briefFormPane);
});

heroTargetsBtn?.addEventListener('click', () => {
  focusMode(Mode.TARGET_GENERATION, targetForm);
});

heroDocsBtn?.addEventListener('click', () => {
  const mode = activeMode || Mode.TARGETED_BRIEF;
  openDocPickerModal(mode);
});

heroExportBtn?.addEventListener('click', () => {
  exportTrigger?.click();
});

heroSettingsBtn?.addEventListener('click', () => {
  openSettingsModal();
});

saveKeyBtn?.addEventListener('click', async () => {
  const key = apiKeyInput.value.trim();
  if (!key) { status.innerText = 'API key required'; status.style.color = '#b91c1c'; return; }
  const provider = cachedLlmProvider === LLMProvider.GROQ ? LLMProvider.GROQ : LLMProvider.GEMINI;
  const payload = {
    [LLM_PROVIDER_STORAGE_KEY]: provider,
    [LLM_MODEL_STORAGE_KEY]: coerceModelForProvider(provider, cachedLlmModel),
  };
  const keysToRemove = [];
  if (provider === LLMProvider.GROQ) {
    payload[GROQ_KEY_STORAGE_KEY] = key;
    keysToRemove.push('geminiKey');
    cachedGroqKey = key;
  } else {
    payload.geminiKey = key;
    keysToRemove.push(GROQ_KEY_STORAGE_KEY);
    cachedGeminiKey = key;
  }
  cachedLlmProvider = provider;
  await chrome.storage.local.set(payload);
  if (keysToRemove.length) {
    await chrome.storage.local.remove(keysToRemove);
  }
  status.innerText = `${provider === LLMProvider.GROQ ? 'Groq' : 'Gemini'} API key saved.`;
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

function startNewResearch(mode = activeMode || Mode.TARGETED_BRIEF) {
  const normalizedMode = mode || Mode.TARGETED_BRIEF;
  if (normalizedMode === Mode.TARGET_GENERATION) {
    resetTargetSearch();
    return;
  }
  currentBriefRequest = { company: '', location: '', product: '', docs: [], modules: [] };
  briefVersionCounter = 0;
  briefVersions = [];
  briefUiMode = 'form';
  if (briefVersionPill) {
    briefVersionPill.hidden = true;
    briefVersionPill.textContent = 'v1';
  }
  if (briefVersionSelect) {
    briefVersionSelect.innerHTML = '<option value="">Current run</option>';
  }
  showBriefForm();
  if (briefBarStatus) {
    briefBarStatus.innerHTML = '';
    briefBarStatus.style.color = '';
  }
  briefModuleOverride = null;
  applyEffectiveBriefModulesToView({ updateControls: true });

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

  resetPersonaTabContainers();

  setEmailOutput('No email generated yet.', { copyTextOverride: '' });
  telephonicPitchErrorMessage = '';
  telephonicPitchDebugAttempts = [];
  setTelePitchOutput('No telephonic pitch generated yet.');
}

newResearchBtn?.addEventListener('click', () => {
  startNewResearch();
});

function resetTargetSearch() {
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
}

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

telePersonaTabs?.addEventListener('click', (evt) => {
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
    const personaLabel = getActivePersonaLabel();
    status.innerText = personaLabel ? `Email for ${personaLabel} copied to clipboard.` : 'Email copied to clipboard.';
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
        const personaLabel = getActivePersonaLabel();
        status.innerText = personaLabel ? `Email for ${personaLabel} copied to clipboard.` : 'Email copied to clipboard.';
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

reviseEmailPersonaBtn?.addEventListener('click', () => openRevisionPlayground({ type: 'email', scope: 'single' }));
reviseEmailAllBtn?.addEventListener('click', () => openRevisionPlayground({ type: 'email', scope: 'all' }));
revisePitchPersonaBtn?.addEventListener('click', () => openRevisionPlayground({ type: 'pitch', scope: 'single' }));
revisePitchAllBtn?.addEventListener('click', () => openRevisionPlayground({ type: 'pitch', scope: 'all' }));

function setModalButtonIcon(button, iconName, fallbackText = '') {
  if (!button) return;
  const icon = window.feather?.icons?.[iconName];
  if (icon) {
    button.innerHTML = icon.toSvg({ 'stroke-width': 1.9, width: 18, height: 18, 'aria-hidden': 'true' });
    return;
  }
  button.textContent = fallbackText;
}

function closeModal() {
  if (modalRoot?.classList.contains('hidden')) return;
  if (modalContainer) modalContainer.classList.remove('modal-maximized');
  if (modalMaximizeBtn) {
    setModalButtonIcon(modalMaximizeBtn, 'maximize-2', 'Max');
    modalMaximizeBtn.setAttribute('aria-label', 'Maximize modal');
    modalMaximizeBtn.setAttribute('aria-pressed', 'false');
  }
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

function toggleModalMaximize(force) {
  if (!modalContainer || !modalMaximizeBtn) return;
  const next = force !== undefined ? force : !modalContainer.classList.contains('modal-maximized');
  modalContainer.classList.toggle('modal-maximized', next);
  setModalButtonIcon(modalMaximizeBtn, next ? 'minimize-2' : 'maximize-2', next ? 'Res' : 'Max');
  modalMaximizeBtn.setAttribute('aria-label', next ? 'Restore modal size' : 'Maximize modal');
  modalMaximizeBtn.setAttribute('aria-pressed', next ? 'true' : 'false');
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
  toggleModalMaximize(false);
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

function openHistorySearchModal(mode = activeMode || Mode.TARGETED_BRIEF) {
  const resolvedMode = mode === Mode.TARGET_GENERATION ? Mode.TARGET_GENERATION : Mode.TARGETED_BRIEF;
  const isTargetMode = resolvedMode === Mode.TARGET_GENERATION;
  const modalTitleText = isTargetMode ? 'Search Targets' : 'Search Briefs';

  const getEntries = () => sortHistoryEntries(isTargetMode ? targetHistoryEntries : historyEntries);
  const formatSubtitle = (entry) => {
    if (!entry) return '';
    if (isTargetMode) {
      const sectors = Array.isArray(entry?.request?.sectors) ? entry.request.sectors.filter(Boolean).join(', ') : '';
      return [entry?.request?.product, entry?.request?.location, sectors].filter(Boolean).join(' • ');
    }
    return [entry?.request?.product, entry?.request?.location].filter(Boolean).join(' • ');
  };

  openModal({
    title: modalTitleText,
    render: ({ body, footer, close }) => {
      if (!body || !footer) return;
      body.innerHTML = '';
      body.classList.add('history-search-body');
      const previousFooterDisplay = footer.style.display;
      footer.style.display = 'none';

      const helper = document.createElement('div');
      helper.className = 'modal-helper';
      helper.textContent = isTargetMode
        ? 'Search your past target runs by product, location, or sector.'
        : 'Search your past briefs by company, product, or location.';
      body.appendChild(helper);

      const searchInput = document.createElement('input');
      searchInput.type = 'search';
      searchInput.className = 'doc-picker-search history-search-input';
      searchInput.placeholder = isTargetMode
        ? 'Search targets (product, location, sector)'
        : 'Search briefs (company, product, location)';
      body.appendChild(searchInput);

      const results = document.createElement('div');
      results.className = 'history-search-results';
      results.setAttribute('role', 'list');
      body.appendChild(results);

      const renderResults = () => {
        const rawTerm = (searchInput.value || '').toLowerCase();
        const tokens = rawTerm.split(/\s+/).filter(Boolean);
        const sortedEntries = getEntries();
        results.innerHTML = '';

        if (!sortedEntries.length) {
          const empty = document.createElement('div');
          empty.className = 'modal-helper';
          empty.textContent = isTargetMode ? 'No targets generated yet.' : 'No briefs generated yet.';
          results.appendChild(empty);
          return;
        }

        const filtered = sortedEntries.filter((entry) => {
          const title = isTargetMode ? getTargetHistoryTitle(entry) : getBriefHistoryTitle(entry);
          const subtitle = formatSubtitle(entry);
          const haystackParts = [
            title,
            subtitle,
            formatHistoryTimestamp(entry.createdAt) || '',
            entry?.request?.company,
            entry?.request?.product,
            entry?.request?.location,
            entry?.customTitle,
          ];
          if (isTargetMode && Array.isArray(entry?.request?.sectors)) {
            haystackParts.push(entry.request.sectors.join(' '));
          }
          const haystack = haystackParts.filter(Boolean).join(' ').toLowerCase();
          return !tokens.length || tokens.every((token) => haystack.includes(token));
        });

        if (!filtered.length) {
          const empty = document.createElement('div');
          empty.className = 'modal-helper';
          empty.textContent = 'No matches. Try a different search.';
          results.appendChild(empty);
          return;
        }

        filtered.forEach((entry) => {
          const item = document.createElement('button');
          item.type = 'button';
          item.className = 'history-item history-search-item';
          item.dataset.id = entry.id;
          item.setAttribute('role', 'listitem');

          const content = document.createElement('div');
          content.className = 'history-content';

          const title = document.createElement('span');
          title.className = 'history-title';
          title.textContent = isTargetMode ? getTargetHistoryTitle(entry) : getBriefHistoryTitle(entry);
          content.appendChild(title);

          const subtitleText = formatSubtitle(entry);
          if (subtitleText) {
            const subtitle = document.createElement('span');
            subtitle.className = 'history-subtitle';
            subtitle.textContent = subtitleText;
            content.appendChild(subtitle);
          }

          const meta = document.createElement('span');
          meta.className = 'history-meta';
          meta.textContent = formatHistoryTimestamp(entry.createdAt) || '';
          content.appendChild(meta);

          item.appendChild(content);

          item.addEventListener('click', () => {
            close();
            setActiveMode(resolvedMode);
            if (isTargetMode) {
              setActiveTargetHistoryItem(entry.id);
              showTargetHistoryEntry(entry, { updateForm: true, statusText: 'Loaded from search.' });
            } else {
              setActiveHistoryItem(entry.id);
              showHistoryEntry(entry, { updateForm: true, statusText: 'Loaded from search.' });
            }
            setHistorySidebarCollapsed(false);
          });

          results.appendChild(item);
        });
      };

      searchInput.addEventListener('input', renderResults);
      setTimeout(() => searchInput.focus(), 0);
      renderResults();

      return () => {
        body.classList.remove('history-search-body');
        footer.style.display = previousFooterDisplay;
      };
    },
  });
}

modalClose?.addEventListener('click', () => closeModal());
modalMaximizeBtn?.addEventListener('click', () => toggleModalMaximize());
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

function openSettingsModal(options = {}) {
  Promise.all([loadExportTemplateFromStorage(), themePreferenceLoadPromise, llmSettingsLoadPromise]).finally(() => {
    openModal({
      title: 'Settings',
      render: (ctx) => renderSettingsModal(ctx, options),
    });
  });
}

settingsTriggers.forEach((btn) => {
  btn.addEventListener('click', () => openSettingsModal());
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

function toLinkedInPeopleSearchUrl(searchString) {
  const q = (searchString ?? '').trim();
  if (!q) {
    return 'https://www.linkedin.com/search/results/people/?keywords=&origin=SWITCH_SEARCH_VERTICAL';
  }

  const keywords = encodeURIComponent(q);
  return `https://www.linkedin.com/search/results/people/?keywords=${keywords}&origin=SWITCH_SEARCH_VERTICAL`;
}

function getLinkedInPeopleSearchLink(persona = {}) {
  const rawUrl = persona.linkedin_search_url || persona.linkedinSearchUrl || '';
  const keywords = persona.linkedin_keywords || persona.linkedinKeywords || '';

  if (rawUrl && /^https?:\/\//i.test(rawUrl)) return rawUrl;
  if (rawUrl && !/\s/.test(rawUrl)) return rawUrl;
  if (keywords) return toLinkedInPeopleSearchUrl(keywords);
  if (rawUrl) return toLinkedInPeopleSearchUrl(rawUrl);

  return '';
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

function renderSettingsModal({ body, footer, close }, options = {}) {
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
  themeOptions.className = 'choice-grid theme-choice-grid';
  themeSection.appendChild(themeOptions);
  const currentThemeSetting = themePreference || ThemePreference.SYSTEM;
  let selectedTheme = currentThemeSetting;
  const themeChoiceLabels = [];
  const themeIcons = {
    [ThemePreference.SYSTEM]:
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2"></rect><line x1="9" y1="20" x2="15" y2="20"></line><line x1="12" y1="16" x2="12" y2="20"></line></svg>',
    [ThemePreference.LIGHT]:
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>',
    [ThemePreference.DARK]:
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 0 1 11.21 3 7 7 0 0 0 12 21a7 7 0 0 0 9-8.21Z"></path></svg>',
  };
  const updateThemeChoiceState = () => {
    themeChoiceLabels.forEach((label) => {
      const input = label.querySelector('input[type="radio"]');
      const isActive = input?.value === selectedTheme;
      if (input) input.checked = isActive;
      label.classList.toggle('active', isActive);
    });
  };
  [
    { value: ThemePreference.SYSTEM, label: 'Match system', description: 'Follow your OS appearance' },
    { value: ThemePreference.LIGHT, label: 'Light mode', description: 'Bright, high-contrast surface' },
    { value: ThemePreference.DARK, label: 'Dark mode', description: 'Dimmed UI for low light' },
  ].forEach((option) => {
    const label = document.createElement('label');
    label.className = 'choice-card theme-choice';
    label.dataset.value = option.value;
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'themePreference';
    radio.value = option.value;
    radio.className = 'visually-hidden';
    radio.checked = selectedTheme === option.value;
    themeInputs.push(radio);
    const icon = document.createElement('div');
    icon.className = 'choice-icon';
    icon.innerHTML = themeIcons[option.value];
    const textWrap = document.createElement('div');
    textWrap.className = 'choice-text';
    const title = document.createElement('div');
    title.className = 'choice-label';
    title.textContent = option.label;
    const desc = document.createElement('div');
    desc.className = 'choice-sub';
    desc.textContent = option.description;
    textWrap.appendChild(title);
    textWrap.appendChild(desc);
    label.appendChild(radio);
    label.appendChild(icon);
    label.appendChild(textWrap);
    label.addEventListener('click', () => {
      selectedTheme = option.value;
      updateThemeChoiceState();
    });
    radio.addEventListener('change', () => {
      selectedTheme = option.value;
      updateThemeChoiceState();
    });
    themeChoiceLabels.push(label);
    themeOptions.appendChild(label);
  });
  updateThemeChoiceState();
  form.appendChild(themeSection);

  const sectionsSection = document.createElement('div');
  sectionsSection.className = 'modal-section';
  const sectionsHeader = document.createElement('h4');
  sectionsHeader.textContent = 'Brief sections';
  sectionsSection.appendChild(sectionsHeader);
  const sectionsHelper = document.createElement('p');
  sectionsHelper.className = 'modal-helper';
  sectionsHelper.textContent = 'Set the default sections for new brief generation.';
  sectionsSection.appendChild(sectionsHelper);
  const sectionsList = document.createElement('div');
  sectionsList.className = 'brief-sections-list';
  sectionsSection.appendChild(sectionsList);
  const handleSectionToggle = (moduleId, checked) => {
    const base = briefModuleConfig || getDefaultBriefModules();
    const next = { ...base, [moduleId]: !!checked };
    const resolved = setBriefModulesConfig(next, { updateControls: false });
    renderBriefModuleControls(resolved, sectionsList, { onToggle: handleSectionToggle });
  };
  renderBriefModuleControls(briefModuleConfig || getDefaultBriefModules(), sectionsList, { onToggle: handleSectionToggle });
  form.appendChild(sectionsSection);

  const apiSection = document.createElement('div');
  apiSection.className = 'modal-section';
  const apiHeader = document.createElement('h4');
  apiHeader.textContent = 'Groq API key';
  apiSection.appendChild(apiHeader);
  const apiHelper = document.createElement('p');
  apiHelper.className = 'modal-helper';
  apiHelper.textContent = 'Use Groq (OpenAI-compatible). Add your API key.';
  apiSection.appendChild(apiHelper);
  const apiFlow = document.createElement('div');
  apiFlow.className = 'api-flow';
  const providerInstruction = document.createElement('div');
  providerInstruction.className = 'provider-instructions';
  const instructionTitle = document.createElement('div');
  instructionTitle.className = 'instructions-title';
  instructionTitle.textContent = 'Get a Groq key';
  providerInstruction.appendChild(instructionTitle);
  const instructionList = document.createElement('ol');
  [
    'Sign in to the Groq console and open the API Keys page.',
    'Create a new key or reuse an existing one for your project.',
    'Keys are OpenAI-compatible; paste it below to enable Groq.',
  ].forEach((step) => {
    const li = document.createElement('li');
    li.textContent = step;
    instructionList.appendChild(li);
  });
  providerInstruction.appendChild(instructionList);
  const groqLink = document.createElement('a');
  groqLink.href = 'https://console.groq.com/keys';
  groqLink.target = '_blank';
  groqLink.rel = 'noopener noreferrer';
  groqLink.className = 'instruction-link';
  groqLink.textContent = 'Open Groq console';
  providerInstruction.appendChild(groqLink);

  providerInstruction.style.marginBottom = '12px';
  apiFlow.appendChild(providerInstruction);

  const eyeIcon = '<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M1 12s4-7 11-7 11 7 11 7-4 7-11 7-11-7-11-7z\"></path><circle cx=\"12\" cy=\"12\" r=\"3\"></circle></svg>';
  const eyeOffIcon = '<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M17.94 17.94A10.94 10.94 0 0 1 12 19c-7 0-11-7-11-7a21.5 21.5 0 0 1 5.06-5.94\"></path><path d=\"M1 1l22 22\"></path><path d=\"M9.53 9.53a3 3 0 0 0 4.24 4.24\"></path><path d=\"M14.47 14.47 9.53 9.53\"></path></svg>';
  const buildKeyField = (placeholder, initialValue) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'input-with-toggle';
    const input = document.createElement('input');
    input.type = 'password';
    input.placeholder = placeholder;
    input.value = initialValue || '';
    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'input-toggle';
    const updateIcon = (isVisible) => {
      toggle.innerHTML = isVisible ? eyeIcon : eyeOffIcon;
      toggle.setAttribute('aria-label', isVisible ? 'Hide API key' : 'Show API key');
    };
    updateIcon(false);
    toggle.addEventListener('click', () => {
      const isHidden = input.type === 'password';
      input.type = isHidden ? 'text' : 'password';
      updateIcon(isHidden);
    });
    wrapper.appendChild(input);
    wrapper.appendChild(toggle);
    return { wrapper, input };
  };

  const groqKeyField = buildKeyField('Groq API key', (cachedGroqKey || '').trim());
  groqKeyField.input.style.padding = '12px 14px';

  const groqKeyLabel = document.createElement('p');
  groqKeyLabel.className = 'modal-helper';
  groqKeyLabel.textContent = 'Groq key (OpenAI-compatible API)';

  const groqKeyContainer = document.createElement('div');
  groqKeyContainer.appendChild(groqKeyLabel);
  groqKeyContainer.appendChild(groqKeyField.wrapper);
  apiFlow.appendChild(groqKeyContainer);
  apiSection.appendChild(apiFlow);

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

    const provider = LLMProvider.GROQ;
    const model = coerceModelForProvider(
      provider,
      normalizeLlmModel(cachedLlmModel || DEFAULT_LLM_MODELS[LLMProvider.GROQ] || '', provider),
    );
    const groqKey = groqKeyField.input.value.trim();
    const selectedThemeValue = selectedTheme || (themeInputs.find((input) => input.checked)?.value) || ThemePreference.SYSTEM;
    const pitchingCompany = pitchingInput.value.trim();

    try {
      const payload = {
        [LLM_PROVIDER_STORAGE_KEY]: provider,
        [LLM_MODEL_STORAGE_KEY]: model,
      };
      const keysToRemove = ['geminiKey'];
      if (groqKey) {
        payload[GROQ_KEY_STORAGE_KEY] = groqKey;
      } else {
        keysToRemove.push(GROQ_KEY_STORAGE_KEY);
      }
      await chrome.storage.local.set(payload);
      if (keysToRemove.length) {
        await chrome.storage.local.remove(keysToRemove);
      }
      if (pitchingCompany) {
        await chrome.storage.local.set({ [PITCH_FROM_COMPANY_KEY]: pitchingCompany });
        cachedPitchingCompany = pitchingCompany;
      } else {
        await chrome.storage.local.remove(PITCH_FROM_COMPANY_KEY);
        cachedPitchingCompany = '';
      }
      cachedGeminiKey = '';
      cachedGroqKey = groqKey;
      cachedLlmProvider = provider;
      cachedLlmModel = model;
      await saveThemePreference(selectedThemeValue);
      if (apiKeyInput) apiKeyInput.value = groqKey;
      if (status) {
        status.innerText = 'Settings saved.';
        status.style.color = '';
      }
      close();
      if (typeof options.onSave === 'function') {
        try {
          options.onSave({ provider, model, groqKey, theme: selectedThemeValue, pitchingCompany });
        } catch (err) {
          console.warn('onSave handler failed', err);
        }
      }
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
  templateHost.className = 'template-editor-container';
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
  templateHost.className = 'template-editor-container';
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

  const createStepPill = (num) => {
    const step = document.createElement('span');
    step.className = 'export-step';
    const label = document.createElement('span');
    label.className = 'export-step-label';
    label.textContent = 'Step';
    const value = document.createElement('span');
    value.className = 'export-step-number';
    value.textContent = num;
    step.appendChild(label);
    step.appendChild(value);
    return step;
  };

  const container = document.createElement('div');
  container.className = 'export-options';
  body.appendChild(container);

  const templateBlock = document.createElement('div');
  templateBlock.className = 'export-section export-section-template';
  const templateHeading = document.createElement('div');
  templateHeading.className = 'export-section-heading';
  const templateStep = createStepPill(1);
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
  const rangeStep = createStepPill(2);
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
  rangeFieldset.classList.add('export-subsection');
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
  dateContainer.className = 'date-range-inputs export-subsection';

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
  customContainer.className = 'custom-selection export-subsection';

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

  const outputBlock = document.createElement('div');
  outputBlock.className = 'export-section export-section-output';
  const outputHeading = document.createElement('div');
  outputHeading.className = 'export-section-heading';
  const outputStep = createStepPill(3);
  const outputTitleGroup = document.createElement('div');
  outputTitleGroup.className = 'export-section-title-group';
  const outputTitle = document.createElement('div');
  outputTitle.className = 'export-section-title';
  outputTitle.textContent = 'Output format';
  const outputHint = document.createElement('p');
  outputHint.className = 'export-section-hint';
  outputHint.textContent = 'Choose how the export is packaged. Keep the preview on the right to spot issues quickly.';
  outputTitleGroup.appendChild(outputTitle);
  outputTitleGroup.appendChild(outputHint);
  outputHeading.appendChild(outputStep);
  outputHeading.appendChild(outputTitleGroup);
  outputBlock.appendChild(outputHeading);

  const formatFieldset = document.createElement('fieldset');
  formatFieldset.classList.add('export-subsection');
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

  outputBlock.appendChild(formatFieldset);
  container.appendChild(outputBlock);

  const sidePanel = document.createElement('div');
  sidePanel.className = 'export-side-card';
  const sideHeader = document.createElement('div');
  sideHeader.className = 'export-side-card-header';
  const sideTitle = document.createElement('div');
  sideTitle.className = 'export-side-card-title';
  sideTitle.textContent = 'Status & preview';
  const sideHint = document.createElement('p');
  sideHint.className = 'export-section-hint';
  sideHint.textContent = 'Monitor progress, capture notes, and review the last export without leaving this modal.';
  sideHeader.appendChild(sideTitle);
  sideHeader.appendChild(sideHint);
  sidePanel.appendChild(sideHeader);

  const statusEl = document.createElement('div');
  statusEl.className = 'export-status';
  sidePanel.appendChild(statusEl);

  const notesEl = document.createElement('div');
  notesEl.className = 'notes-box';
  notesEl.style.display = 'none';
  sidePanel.appendChild(notesEl);

  const previewWrapper = document.createElement('div');
  previewWrapper.className = 'export-preview';
  sidePanel.appendChild(previewWrapper);
  container.appendChild(sidePanel);

  const footerActions = document.createElement('div');
  footerActions.className = 'export-footer-actions';
  const footerLeft = document.createElement('div');
  footerLeft.className = 'export-footer-left';
  const footerRight = document.createElement('div');
  footerRight.className = 'export-footer-right';
  footerActions.appendChild(footerLeft);
  footerActions.appendChild(footerRight);
  footer.appendChild(footerActions);

  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'ghost';
  cancelBtn.textContent = 'Close';
  cancelBtn.addEventListener('click', () => {
    persistStateDebounced();
    close();
  });
  footerLeft.appendChild(cancelBtn);

  const saveAndExitBtn = document.createElement('button');
  saveAndExitBtn.type = 'button';
  saveAndExitBtn.className = 'ghost';
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
  footerLeft.appendChild(saveAndExitBtn);

  const downloadBtn = document.createElement('button');
  downloadBtn.type = 'button';
  downloadBtn.textContent = 'Download';
  downloadBtn.disabled = !state.result;
  downloadBtn.addEventListener('click', () => {
    if (state.result?.download) {
      triggerDownload(state.result.download);
    }
  });
  footerRight.appendChild(downloadBtn);

  const exportBtn = document.createElement('button');
  exportBtn.type = 'button';
  exportBtn.classList.add('primary');
  footerRight.appendChild(exportBtn);

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

  personas.forEach((p, idx) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'persona';

    const title = document.createElement('div');
    const nameEl = document.createElement('strong');
    nameEl.textContent = formatPersonaLabel(p, idx);
    title.appendChild(nameEl);
    wrapper.appendChild(title);

    const rawLink = p.zoominfo_link || p.zoomInfo || p.zoominfo || p.zoom || '';
    const link = normalizeUrl(rawLink);
    const linkedinLink = getLinkedInPeopleSearchLink(p);

    if (link || linkedinLink) {
      const linkWrap = document.createElement('div');
      if (link) {
        const a = document.createElement('a');
        a.href = link;
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        a.textContent = 'Google Search';
        linkWrap.appendChild(a);
      }
      if (link && linkedinLink) {
        linkWrap.appendChild(document.createTextNode(' | '));
      }
      if (linkedinLink) {
        const linkedinAnchor = document.createElement('a');
        linkedinAnchor.href = linkedinLink;
        linkedinAnchor.target = '_blank';
        linkedinAnchor.rel = 'noopener noreferrer';
        linkedinAnchor.textContent = 'LinkedIn Search';
        linkWrap.appendChild(linkedinAnchor);
      }
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

function escapeHtml(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeUrl(url = '') {
  if (typeof url !== 'string') return '';
  const trimmed = url.trim();
  if (!trimmed) return '';
  const lowered = trimmed.toLowerCase();
  if (lowered.startsWith('javascript:') || lowered.startsWith('data:') || lowered.startsWith('vbscript:')) return '';
  return escapeHtml(trimmed);
}

function formatInlineMarkdown(text = '') {
  let escaped = escapeHtml(text);
  escaped = escaped.replace(/`([^`]+)`/g, (_, code) => `<code>${escapeHtml(code)}</code>`);
  escaped = escaped.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  escaped = escaped.replace(/__(.+?)__/g, '<strong>$1</strong>');
  escaped = escaped.replace(/(\s|^)\*(.+?)\*(\s|$)/g, (_, lead, val, trail) => `${lead}<em>${val}</em>${trail}`);
  escaped = escaped.replace(/(\s|^)_(.+?)_(\s|$)/g, (_, lead, val, trail) => `${lead}<em>${val}</em>${trail}`);
  escaped = escaped.replace(/\[([^\]]+)\]\(([^)]+)\)/g, (match, label, url) => {
    const safeUrl = escapeUrl(url);
    if (!safeUrl) return escapeHtml(label);
    return `<a href="${safeUrl}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`;
  });
  return escaped;
}

function markdownToHtml(md = '') {
  if (typeof md !== 'string' || !md.trim()) return escapeHtml(md || '');
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  const html = [];
  let inCode = false;
  let codeBuffer = [];
  let listType = null;
  let pendingParagraph = [];

  const closeList = () => {
    if (listType) {
      html.push(`</${listType}>`);
      listType = null;
    }
  };

  const flushParagraph = () => {
    if (!pendingParagraph.length) return;
    html.push(`<p>${formatInlineMarkdown(pendingParagraph.join(' '))}</p>`);
    pendingParagraph = [];
  };

  for (const rawLine of lines) {
    const line = rawLine.trimEnd();

    if (line.trim().startsWith('```')) {
      flushParagraph();
      closeList();
      if (inCode) {
        html.push(`<pre><code>${escapeHtml(codeBuffer.join('\n'))}</code></pre>`);
        codeBuffer = [];
        inCode = false;
      } else {
        inCode = true;
        codeBuffer = [];
      }
      continue;
    }

    if (inCode) {
      codeBuffer.push(rawLine);
      continue;
    }

    const headingMatch = line.match(/^(#{1,3})\s+(.*)$/);
    if (headingMatch) {
      flushParagraph();
      closeList();
      const level = headingMatch[1].length;
      html.push(`<h${level}>${formatInlineMarkdown(headingMatch[2])}</h${level}>`);
      continue;
    }

    const ulMatch = rawLine.match(/^\s*[-*]\s+(.*)$/);
    const olMatch = rawLine.match(/^\s*\d+\.\s+(.*)$/);
    if (ulMatch || olMatch) {
      flushParagraph();
      const targetList = olMatch ? 'ol' : 'ul';
      if (listType && listType !== targetList) {
        closeList();
      }
      if (!listType) {
        listType = targetList;
        html.push(`<${targetList}>`);
      }
      const itemContent = formatInlineMarkdown((ulMatch ? ulMatch[1] : olMatch[1]) || '');
      html.push(`<li>${itemContent}</li>`);
      continue;
    }

    if (!line.trim()) {
      flushParagraph();
      closeList();
      continue;
    }

    if (listType) {
      closeList();
    }

    pendingParagraph.push(line.trim());
  }

  if (inCode) {
    html.push(`<pre><code>${escapeHtml(codeBuffer.join('\n'))}</code></pre>`);
  }
  flushParagraph();
  closeList();

  return html.join('\n');
}

function setEmailOutput(text = '', { allowMarkdown = false, copyTextOverride } = {}) {
  if (!emailOut) return;
  const value = typeof text === 'string' ? text : '';
  if (allowMarkdown) {
    emailOut.innerHTML = markdownToHtml(value);
  } else {
    emailOut.textContent = value;
  }
  const copyValue = copyTextOverride !== undefined ? copyTextOverride : value;
  updateCopyEmailButtonState(copyValue);
}

function updateCopyEmailButtonState(text = '') {
  if (!copyEmailBtn) return;
  const hasText = typeof text === 'string' && text.trim().length > 0;
  copyEmailBtn.disabled = !hasText;
}

function getPersonaTabContainers() {
  return [personaTabs, telePersonaTabs].filter(Boolean);
}

function resetPersonaTabContainers() {
  getPersonaTabContainers().forEach((container) => {
    container.innerHTML = '';
    container.style.display = 'none';
  });
}

function buildPersonaTabButton(draft, idx, activeIndex = 0) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'persona-tab';
  btn.dataset.index = String(idx);
  btn.setAttribute('role', 'tab');
  const isActive = idx === activeIndex;
  btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
  btn.tabIndex = isActive ? 0 : -1;

  btn.textContent = formatPersonaLabel(draft, idx);
  return btn;
}

function renderPersonaTabsForContainers(drafts = [], activeIndex = 0) {
  getPersonaTabContainers().forEach((container) => {
    container.innerHTML = '';
    if (!drafts.length) {
      container.style.display = 'none';
      return;
    }
    container.style.display = '';
    drafts.forEach((draft, idx) => {
      container.appendChild(buildPersonaTabButton(draft, idx, activeIndex));
    });
  });
}

function syncPersonaTabActiveStates(activeIndex) {
  getPersonaTabContainers().forEach((container) => {
    const buttons = container.querySelectorAll('.persona-tab');
    buttons.forEach((btn, btnIdx) => {
      const isActive = btnIdx === activeIndex;
      btn.classList.toggle('active', isActive);
      btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
      btn.tabIndex = isActive ? 0 : -1;
    });
  });
}

function getActivePersonaLabel() {
  if (selectedPersonaIndex < 0) return '';
  const draft = personaEmailDrafts[selectedPersonaIndex];
  return draft ? formatPersonaLabel(draft, selectedPersonaIndex) : '';
}

function activatePersonaTab(index) {
  if (!personaEmailDrafts.length) return;
  const safeIndex = Math.max(0, Math.min(index, personaEmailDrafts.length - 1));
  selectedPersonaIndex = safeIndex;

  syncPersonaTabActiveStates(safeIndex);

  const draft = personaEmailDrafts[safeIndex] || {};
  const text = formatEmailDraftText(draft);
  if (text) {
    setEmailOutput(text, { allowMarkdown: true });
  } else {
    setEmailOutput('No email available for this persona yet.', { copyTextOverride: '' });
  }

  if (telePitchOut) {
    if (telephonicPitchErrorMessage) {
      setTelePitchOutput(`Telephonic pitch failed: ${telephonicPitchErrorMessage}`, { isError: true });
    } else {
      const pitchText = typeof draft.telePitch === 'string' ? draft.telePitch.trim() : '';
      setTelePitchOutput(pitchText || 'No telephonic pitch available for this persona yet.', { allowMarkdown: !!pitchText });
    }
  }

  renderEmailVersionControls();
  renderPitchVersionControls();
  updateRevisionButtonsState();
}

function renderPersonaEmailDrafts(personasData = [], personaEmailsData = [], telephonicPitchesData = [], fallbackEmail) {
  personaEmailDrafts = mergePersonaEmails(personasData, personaEmailsData, telephonicPitchesData);

  const hasDrafts = personaEmailDrafts.length > 0;
  const initialIndex = hasDrafts
    ? (selectedPersonaIndex >= 0 ? Math.min(selectedPersonaIndex, personaEmailDrafts.length - 1) : 0)
    : 0;

  renderPersonaTabsForContainers(personaEmailDrafts, initialIndex);

  if (hasDrafts) {
    activatePersonaTab(initialIndex);
  } else {
    personaEmailDrafts = [];
    selectedPersonaIndex = -1;
    const fallbackText = formatFallbackEmailText(fallbackEmail);
    if (fallbackText) {
      setEmailOutput(fallbackText, { allowMarkdown: true });
    } else {
      setEmailOutput('No email generated yet.', { copyTextOverride: '' });
    }
    if (telephonicPitchErrorMessage) {
      setTelePitchOutput(`Telephonic pitch failed: ${telephonicPitchErrorMessage}`, { isError: true });
    } else {
      setTelePitchOutput('No telephonic pitch generated yet.');
    }
    renderEmailVersionControls();
    renderPitchVersionControls();
  }

  updateRevisionButtonsState();
}

function renderEmailVersionControls() {
  if (!emailVersionControls) return;
  emailVersionControls.innerHTML = '';
  emailVersionControls.style.display = 'none';
  if (selectedPersonaIndex < 0 || !Array.isArray(personaEmailVersions) || !personaEmailVersions[selectedPersonaIndex]) return;
  const entry = personaEmailVersions[selectedPersonaIndex];
  if (!entry || !Array.isArray(entry.versions) || entry.versions.length <= 1) return;

  emailVersionControls.style.display = 'flex';
  const prev = document.createElement('button');
  prev.type = 'button';
  prev.textContent = '\u2190';
  prev.disabled = entry.activeIndex <= 0;
  prev.addEventListener('click', () => {
    setEmailVersion(selectedPersonaIndex, entry.activeIndex - 1);
  });

  const next = document.createElement('button');
  next.type = 'button';
  next.textContent = '\u2192';
  next.disabled = entry.activeIndex >= entry.versions.length - 1;
  next.addEventListener('click', () => {
    setEmailVersion(selectedPersonaIndex, entry.activeIndex + 1);
  });

  const label = document.createElement('span');
  label.className = 'version-label';
  label.textContent = `v${entry.activeIndex + 1}/${entry.versions.length}`;

  emailVersionControls.appendChild(prev);
  emailVersionControls.appendChild(label);
  emailVersionControls.appendChild(next);
}

function renderPitchVersionControls() {
  if (!pitchVersionControls) return;
  pitchVersionControls.innerHTML = '';
  pitchVersionControls.style.display = 'none';
  if (selectedPersonaIndex < 0 || !Array.isArray(telephonicPitchVersions) || !telephonicPitchVersions[selectedPersonaIndex]) return;
  const entry = telephonicPitchVersions[selectedPersonaIndex];
  if (!entry || !Array.isArray(entry.versions) || entry.versions.length <= 1) return;

  pitchVersionControls.style.display = 'flex';
  const prev = document.createElement('button');
  prev.type = 'button';
  prev.textContent = '\u2190';
  prev.disabled = entry.activeIndex <= 0;
  prev.addEventListener('click', () => {
    setPitchVersion(selectedPersonaIndex, entry.activeIndex - 1);
  });

  const next = document.createElement('button');
  next.type = 'button';
  next.textContent = '\u2192';
  next.disabled = entry.activeIndex >= entry.versions.length - 1;
  next.addEventListener('click', () => {
    setPitchVersion(selectedPersonaIndex, entry.activeIndex + 1);
  });

  const label = document.createElement('span');
  label.className = 'version-label';
  label.textContent = `v${entry.activeIndex + 1}/${entry.versions.length}`;

  pitchVersionControls.appendChild(prev);
  pitchVersionControls.appendChild(label);
  pitchVersionControls.appendChild(next);
}

function updateRevisionButtonsState() {
  const hasPersonas = personaEmailDrafts && personaEmailDrafts.length > 0;
  const emailEntry = hasPersonas && selectedPersonaIndex >= 0
    ? personaEmailVersions?.[selectedPersonaIndex]
    : null;
  const canReviseEmail = hasPersonas && isLatestVersion(emailEntry);
  if (reviseEmailPersonaBtn) {
    reviseEmailPersonaBtn.disabled = !canReviseEmail;
    reviseEmailPersonaBtn.style.display = '';
  }
  if (reviseEmailAllBtn) {
    reviseEmailAllBtn.disabled = !canReviseEmail;
    reviseEmailAllBtn.style.display = '';
  }

  const hasPitches = Array.isArray(activeBriefData?.personas) && activeBriefData.personas.length > 0;
  const pitchEntry = hasPitches && selectedPersonaIndex >= 0
    ? telephonicPitchVersions?.[selectedPersonaIndex]
    : null;
  const canRevisePitch = hasPitches && isLatestVersion(pitchEntry);
  if (revisePitchPersonaBtn) {
    revisePitchPersonaBtn.disabled = !canRevisePitch;
    revisePitchPersonaBtn.style.display = '';
  }
  if (revisePitchAllBtn) {
    revisePitchAllBtn.disabled = !canRevisePitch;
    revisePitchAllBtn.style.display = '';
  }
}

function setEmailVersion(personaIdx, versionIdx) {
  if (!activeBriefData || !Array.isArray(personaEmailVersions)) return;
  if (personaIdx < 0 || personaIdx >= personaEmailVersions.length) return;
  const entry = personaEmailVersions[personaIdx];
  if (!entry || !Array.isArray(entry.versions)) return;
  entry.activeIndex = clampIndex(versionIdx, entry.versions);
  activeBriefData.personaEmailVersions = personaEmailVersions;
  activeBriefData.personaEmailVersionIndexes = personaEmailVersions.map((v) => {
    const idx = clampIndex(
      typeof v?.activeIndex === 'number' ? v.activeIndex : 0,
      Array.isArray(v?.versions) ? v.versions : []
    );
    return idx < 0 ? 0 : idx;
  });
  syncVersionState();
  const personasData = Array.isArray(activeBriefData.personas) ? activeBriefData.personas : [];
  const telephonicPitchesData = Array.isArray(activeBriefData.telephonicPitches) ? activeBriefData.telephonicPitches : [];
  renderPersonaEmailDrafts(personasData, activeBriefData.personaEmails, telephonicPitchesData, activeBriefData.email);
  activatePersonaTab(personaIdx);
  persistRevisionToHistory();
}

function setPitchVersion(personaIdx, versionIdx) {
  if (!activeBriefData || !Array.isArray(telephonicPitchVersions)) return;
  if (personaIdx < 0 || personaIdx >= telephonicPitchVersions.length) return;
  const entry = telephonicPitchVersions[personaIdx];
  if (!entry || !Array.isArray(entry.versions)) return;
  entry.activeIndex = clampIndex(versionIdx, entry.versions);
  activeBriefData.telephonicPitchVersions = telephonicPitchVersions;
  activeBriefData.telephonicPitchVersionIndexes = telephonicPitchVersions.map((v) => {
    const idx = clampIndex(
      typeof v?.activeIndex === 'number' ? v.activeIndex : 0,
      Array.isArray(v?.versions) ? v.versions : []
    );
    return idx < 0 ? 0 : idx;
  });
  syncVersionState();
  const personasData = Array.isArray(activeBriefData.personas) ? activeBriefData.personas : [];
  const telephonicPitchesData = Array.isArray(activeBriefData.telephonicPitches) ? activeBriefData.telephonicPitches : [];
  renderPersonaEmailDrafts(personasData, activeBriefData.personaEmails, telephonicPitchesData, activeBriefData.email);
  activatePersonaTab(personaIdx);
  renderPitchVersionControls();
  persistRevisionToHistory();
}

function formatPersonaLabel(persona = {}, idx = 0) {
  const role =
    (typeof persona.designation === 'string' && persona.designation.trim()) ||
    (typeof persona.personaDesignation === 'string' && persona.personaDesignation.trim()) ||
    (typeof persona.persona_designation === 'string' && persona.persona_designation.trim()) ||
    (typeof persona.title === 'string' && persona.title.trim()) ||
    '';
  const dept =
    (typeof persona.department === 'string' && persona.department.trim()) ||
    (typeof persona.personaDepartment === 'string' && persona.personaDepartment.trim()) ||
    (typeof persona.persona_department === 'string' && persona.persona_department.trim()) ||
    '';
  return role || dept || 'Persona';
}

function formatRevisionCurrentText(type, indices = []) {
  if (!activeBriefData) return '';
  const personas = Array.isArray(activeBriefData.personas) ? activeBriefData.personas : [];
  const emails = Array.isArray(activeBriefData.personaEmails) ? activeBriefData.personaEmails : [];
  const pitches = Array.isArray(activeBriefData.telephonicPitches) ? activeBriefData.telephonicPitches : [];
  const parts = [];

  indices.forEach((idx) => {
    const persona = personas[idx] || {};
    const label = formatPersonaLabel(persona, idx);
    if (type === 'email') {
      const text = formatEmailDraftText(emails[idx] || {}) || 'No email available.';
      parts.push(`--- ${label} ---\n${text}`);
    } else {
      const text = formatTelephonicPitchText(pitches[idx] || {}) || 'No pitch available.';
      parts.push(`--- ${label} ---\n${text}`);
    }
  });

  return parts.join('\n\n');
}

function formatRevisionPreviewText(type, previewDrafts = []) {
  const personas = Array.isArray(activeBriefData?.personas) ? activeBriefData.personas : [];
  const parts = previewDrafts.map(({ personaIndex, draft }) => {
    const label = formatPersonaLabel(personas[personaIndex] || {}, personaIndex);
    if (type === 'email') {
      return `--- ${label} ---\n${formatEmailDraftText(draft || {}) || 'No content returned.'}`;
    }
    return `--- ${label} ---\n${formatTelephonicPitchText(draft || {}) || 'No content returned.'}`;
  });
  return parts.join('\n\n');
}

async function persistRevisionToHistory() {
  if (!currentHistoryId || !activeBriefData) return;
  try {
    await sendMessagePromise({
      action: 'updateResearchHistoryEntry',
      id: currentHistoryId,
      result: {
        personaEmails: Array.isArray(activeBriefData.personaEmails) ? activeBriefData.personaEmails : [],
        telephonicPitches: Array.isArray(activeBriefData.telephonicPitches) ? activeBriefData.telephonicPitches : [],
        personaEmailVersions: Array.isArray(personaEmailVersions) ? personaEmailVersions : [],
        telephonicPitchVersions: Array.isArray(telephonicPitchVersions) ? telephonicPitchVersions : [],
        personaEmailVersionIndexes: Array.isArray(personaEmailVersions)
          ? personaEmailVersions.map((v) => {
            const idx = clampIndex(v.activeIndex || 0, Array.isArray(v.versions) ? v.versions : []);
            return idx < 0 ? 0 : idx;
          })
          : [],
        telephonicPitchVersionIndexes: Array.isArray(telephonicPitchVersions)
          ? telephonicPitchVersions.map((v) => {
            const idx = clampIndex(v.activeIndex || 0, Array.isArray(v.versions) ? v.versions : []);
            return idx < 0 ? 0 : idx;
          })
          : [],
      },
    });
  } catch (err) {
    console.warn('Failed to persist revision to history', err);
  }
}

function applyRevisionPreview(preview, type) {
  if (!preview || !Array.isArray(preview.drafts) || !activeBriefData) return;
  const targetPersonaIdx = preview.drafts[0]?.personaIndex ?? 0;
  preview.drafts.forEach(({ personaIndex, draft }) => {
    if (type === 'email') {
      if (!personaEmailVersions[personaIndex]) personaEmailVersions[personaIndex] = buildVersionEntry(draft || {});
      else {
        personaEmailVersions[personaIndex].versions.push(draft || {});
        personaEmailVersions[personaIndex].activeIndex = personaEmailVersions[personaIndex].versions.length - 1;
      }
    } else {
      if (!telephonicPitchVersions[personaIndex]) telephonicPitchVersions[personaIndex] = buildVersionEntry(draft || {});
      else {
        telephonicPitchVersions[personaIndex].versions.push(draft || {});
        telephonicPitchVersions[personaIndex].activeIndex = telephonicPitchVersions[personaIndex].versions.length - 1;
      }
    }
  });

  activeBriefData.personaEmailVersions = personaEmailVersions;
  activeBriefData.telephonicPitchVersions = telephonicPitchVersions;
  activeBriefData.personaEmailVersionIndexes = personaEmailVersions.map((v) => {
    const idx = clampIndex(
      typeof v?.activeIndex === 'number' ? v.activeIndex : 0,
      Array.isArray(v?.versions) ? v.versions : []
    );
    return idx < 0 ? 0 : idx;
  });
  activeBriefData.telephonicPitchVersionIndexes = telephonicPitchVersions.map((v) => {
    const idx = clampIndex(
      typeof v?.activeIndex === 'number' ? v.activeIndex : 0,
      Array.isArray(v?.versions) ? v.versions : []
    );
    return idx < 0 ? 0 : idx;
  });
  if (type === 'pitch') {
    activeBriefData.telephonicPitchError = '';
    telephonicPitchErrorMessage = '';
    activeBriefData.telephonicPitchAttempts = [];
    telephonicPitchDebugAttempts = [];
  }
  syncVersionState();
  const personasData = Array.isArray(activeBriefData.personas) ? activeBriefData.personas : [];
  renderPersonaEmailDrafts(personasData, activeBriefData.personaEmails, activeBriefData.telephonicPitches, activeBriefData.email);
  const safeIdx = clampIndex(targetPersonaIdx, personaEmailDrafts);
  if (safeIdx >= 0) activatePersonaTab(safeIdx);
  renderPitchVersionControls();
  updateRevisionButtonsState();
  persistRevisionToHistory();
}

function openRevisionPlayground({ type = 'email', scope = 'single' } = {}) {
  if (!activeBriefData) return;
  const personas = Array.isArray(activeBriefData.personas) ? activeBriefData.personas : [];
  const personasAvailable = personas.length;
  if (!personasAvailable) return;
  const personaIdx = selectedPersonaIndex >= 0 ? selectedPersonaIndex : 0;
  const indices = scope === 'all' ? personas.map((_, idx) => idx) : [personaIdx];
  if (!indices.length) return;

  const currentText = formatRevisionCurrentText(type, indices);
  revisionModalState = {
    type,
    scope,
    instructions: '',
    loading: false,
    error: '',
    preview: null,
    indices,
    status: '',
    messageHistory: [{
      role: 'assistant',
      title: 'Initial version',
      content: currentText || 'Nothing to show yet.',
    }],
  };

  openModal({
    title: `Revision Playground (${type === 'email' ? 'Email' : 'Telephonic pitch'})`,
    render: ({ body, footer, close }) => {
      if (!body || !footer) return;
      body.innerHTML = '';
      footer.innerHTML = '';
      body.classList.add('revision-body');

      const scopeHelper = document.createElement('div');
      scopeHelper.className = 'modal-helper';
      const personaNames = indices.map((idx) => formatPersonaLabel(personas[idx] || {}, idx)).join(', ');
      scopeHelper.textContent = scope === 'all'
        ? `Revising all personas (${personaNames}).`
        : `Revising: ${personaNames}.`;

      const chatWrapper = document.createElement('div');
      chatWrapper.className = 'revision-chat';

      const messageList = document.createElement('div');
      messageList.className = 'revision-messages';
      chatWrapper.appendChild(messageList);

      const inputRow = document.createElement('div');
      inputRow.className = 'revision-input-row';
      const instructionsLabel = document.createElement('label');
      instructionsLabel.textContent = 'Ask for a revision';
      const instructionsInput = document.createElement('textarea');
      instructionsInput.rows = 3;
      instructionsInput.className = 'revision-instructions';
      instructionsInput.placeholder = 'e.g., shorten to 3 sentences, highlight security compliance, keep CTA strong.';
      inputRow.appendChild(instructionsLabel);
      inputRow.appendChild(instructionsInput);
      const sendBtn = document.createElement('button');
      sendBtn.type = 'button';
      sendBtn.className = 'primary';
      sendBtn.textContent = 'Send';
      inputRow.appendChild(sendBtn);

      const statusEl = document.createElement('div');
      statusEl.className = 'modal-helper revision-status';

      body.appendChild(scopeHelper);
      body.appendChild(chatWrapper);
      body.appendChild(inputRow);
      body.appendChild(statusEl);

      const applyBtn = document.createElement('button');
      applyBtn.type = 'button';
      applyBtn.textContent = 'Update all drafts';
      applyBtn.disabled = true;

      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel';
      cancelBtn.addEventListener('click', () => close());

      footer.appendChild(cancelBtn);
      footer.appendChild(applyBtn);

      const renderMessages = () => {
        if (!revisionModalState) return;
        const history = Array.isArray(revisionModalState.messageHistory) ? revisionModalState.messageHistory : [];
        messageList.innerHTML = '';
        if (!history.length) {
          const empty = document.createElement('div');
          empty.className = 'modal-helper';
          empty.textContent = 'Messages will appear here after you send a revision.';
          messageList.appendChild(empty);
          return;
        }
        history.forEach((msg) => {
          const bubble = document.createElement('div');
          bubble.className = `revision-message ${msg.role === 'user' ? 'user' : 'assistant'}`;
          const heading = document.createElement('div');
          heading.className = 'revision-message-title';
          heading.textContent = msg.title || (msg.role === 'user' ? 'Instruction' : 'Response');
          const content = document.createElement('div');
          content.className = 'revision-message-content outputBox richText';
          content.innerHTML = markdownToHtml(msg.content || '');
          bubble.appendChild(heading);
          bubble.appendChild(content);

          if (msg.role === 'assistant' && msg.preview && Array.isArray(msg.preview.drafts) && msg.preview.drafts.length) {
            const actions = document.createElement('div');
            actions.className = 'revision-preview-actions';
            const updateBtn = document.createElement('button');
            updateBtn.type = 'button';
            updateBtn.className = 'primary';
            updateBtn.textContent = 'Update draft';
            updateBtn.disabled = revisionModalState?.loading;
            updateBtn.addEventListener('click', () => {
              applyRevisionPreview({ drafts: msg.preview.drafts }, type);
              close();
            });
            actions.appendChild(updateBtn);
            bubble.appendChild(actions);
          }

          messageList.appendChild(bubble);
        });
        messageList.scrollTop = messageList.scrollHeight;
      };

      const appendUserMessage = (history, instructionText) => {
        const next = Array.isArray(history) ? [...history] : [];
        next.push({
          role: 'user',
          title: 'Instruction',
          content: instructionText || 'No instructions provided.',
        });
        return next;
      };

      const appendAssistantMessage = (history, previewText, previewPayload) => {
        const next = Array.isArray(history) ? [...history] : [];
        next.push({
          role: 'assistant',
          title: 'Revised preview',
          content: previewText || 'No content returned.',
          preview: previewPayload,
        });
        return next;
      };

      const refreshUI = () => {
        if (!revisionModalState) return;
        instructionsInput.value = revisionModalState.instructions || '';
        sendBtn.disabled = revisionModalState.loading;
        if (revisionModalState.loading) {
          sendBtn.innerHTML = '<span class="spinner" aria-hidden="true"></span>Sending...';
        } else {
          sendBtn.textContent = 'Send';
        }
        applyBtn.disabled = revisionModalState.loading || !revisionModalState.preview;
        if (revisionModalState.error) {
          statusEl.textContent = revisionModalState.error;
          statusEl.style.color = '#b91c1c';
        } else if (revisionModalState.loading) {
          statusEl.textContent = 'Revising...';
          statusEl.style.color = '';
        } else if (revisionModalState.status) {
          statusEl.textContent = revisionModalState.status;
          statusEl.style.color = '';
        } else {
          statusEl.textContent = '';
        }
        renderMessages();
      };

      instructionsInput.addEventListener('input', (evt) => {
        revisionModalState = { ...revisionModalState, instructions: evt.target.value };
      });

      instructionsInput.addEventListener('keydown', (evt) => {
        if (evt.key !== 'Enter' || evt.shiftKey || evt.ctrlKey || evt.metaKey || evt.altKey || evt.isComposing) return;
        if (sendBtn.disabled || revisionModalState?.loading) return;
        evt.preventDefault();
        runRevision();
      });

      const runRevision = async () => {
        if (!revisionModalState) return;
        const instructions = (revisionModalState.instructions || '').trim();
        const pendingHistory = appendUserMessage(revisionModalState.messageHistory, instructions);
        revisionModalState = {
          ...revisionModalState,
          loading: true,
          error: '',
          status: '',
          instructions: '',
          messageHistory: pendingHistory,
        };
        refreshUI();
        try {
          const company = (companyEl?.value || activeBriefData?.company_name || '').trim();
          const product = (productEl?.value || '').trim();
          const location = (locationEl?.value || '').trim();
          const pitchingOrg = ((cachedPitchingCompany || await pitchingCompanyLoadPromise) || '').trim();
          const payloadBase = {
            company,
            product,
            location,
            instructions,
            pitchingOrg,
          };
          let result = null;
          if (type === 'email') {
            if (scope === 'all') {
              const personasSubset = revisionModalState.indices.map((idx) => personas[idx] || {});
              const emailsSubset = revisionModalState.indices.map((idx) => {
                const entry = personaEmailVersions[idx];
                if (entry && Array.isArray(entry.versions)) return entry.versions[entry.activeIndex] || {};
                return activeBriefData?.personaEmails?.[idx] || {};
              });
              result = await sendMessagePromise({
                action: 'reviseAllPersonaEmails',
                personas: personasSubset,
                emails: emailsSubset,
                ...payloadBase,
              });
              const drafts = Array.isArray(result?.results)
                ? result.results.filter((r) => !r.error && r.draft).map((r) => ({
                  personaIndex: revisionModalState.indices[r.personaIndex] ?? r.personaIndex ?? 0,
                  draft: r.draft,
                }))
                : [];
              if (!drafts.length) {
                throw new Error(result?.results?.[0]?.error || 'Revision failed.');
              }
              revisionModalState = {
                ...revisionModalState,
                preview: { drafts },
                loading: false,
                status: 'Preview ready.',
                messageHistory: appendAssistantMessage(
                  pendingHistory,
                  formatRevisionPreviewText(type, drafts),
                  { drafts },
                ),
              };
            } else {
              const idx = revisionModalState.indices[0];
              const persona = personas[idx] || {};
              const entry = personaEmailVersions[idx];
              const emailDraft = entry?.versions?.[entry.activeIndex] || activeBriefData?.personaEmails?.[idx] || {};
              result = await sendMessagePromise({
                action: 'revisePersonaEmail',
                persona,
                email: emailDraft,
                ...payloadBase,
              });
              if (result?.error || !result?.draft) {
                throw new Error(result?.error || 'Revision failed.');
              }
              revisionModalState = {
                ...revisionModalState,
                preview: { drafts: [{ personaIndex: idx, draft: result.draft }] },
                loading: false,
                status: 'Preview ready.',
                messageHistory: appendAssistantMessage(
                  pendingHistory,
                  formatRevisionPreviewText(type, [{ personaIndex: idx, draft: result.draft }]),
                  { drafts: [{ personaIndex: idx, draft: result.draft }] },
                ),
              };
            }
          } else {
            if (scope === 'all') {
              const personasSubset = revisionModalState.indices.map((idx) => personas[idx] || {});
              const pitchSubset = revisionModalState.indices.map((idx) => {
                const entry = telephonicPitchVersions[idx];
                if (entry && Array.isArray(entry.versions)) return entry.versions[entry.activeIndex] || {};
                return activeBriefData?.telephonicPitches?.[idx] || {};
              });
              result = await sendMessagePromise({
                action: 'reviseAllPersonaPitches',
                personas: personasSubset,
                pitches: pitchSubset,
                ...payloadBase,
              });
              const drafts = Array.isArray(result?.results)
                ? result.results.filter((r) => !r.error && r.draft).map((r) => ({
                  personaIndex: revisionModalState.indices[r.personaIndex] ?? r.personaIndex ?? 0,
                  draft: r.draft,
                }))
                : [];
              if (!drafts.length) {
                throw new Error(result?.results?.[0]?.error || 'Revision failed.');
              }
              revisionModalState = {
                ...revisionModalState,
                preview: { drafts },
                loading: false,
                status: 'Preview ready.',
                messageHistory: appendAssistantMessage(
                  pendingHistory,
                  formatRevisionPreviewText(type, drafts),
                  { drafts },
                ),
              };
            } else {
              const idx = revisionModalState.indices[0];
              const persona = personas[idx] || {};
              const entry = telephonicPitchVersions[idx];
              const pitchDraft = entry?.versions?.[entry.activeIndex] || activeBriefData?.telephonicPitches?.[idx] || {};
              result = await sendMessagePromise({
                action: 'revisePersonaPitch',
                persona,
                pitch: pitchDraft,
                ...payloadBase,
              });
              if (result?.error || !result?.draft) {
                throw new Error(result?.error || 'Revision failed.');
              }
              revisionModalState = {
                ...revisionModalState,
                preview: { drafts: [{ personaIndex: idx, draft: result.draft }] },
                loading: false,
                status: 'Preview ready.',
                messageHistory: appendAssistantMessage(
                  pendingHistory,
                  formatRevisionPreviewText(type, [{ personaIndex: idx, draft: result.draft }]),
                  { drafts: [{ personaIndex: idx, draft: result.draft }] },
                ),
              };
            }
          }
        } catch (err) {
          revisionModalState = { ...revisionModalState, loading: false, error: err?.message || String(err) };
        }
        refreshUI();
      };

      sendBtn.addEventListener('click', runRevision);
      applyBtn.addEventListener('click', () => {
        if (!revisionModalState?.preview) return;
        applyRevisionPreview(revisionModalState.preview, type);
        close();
      });

      refreshUI();
      return () => {
        revisionModalState = null;
        body.classList.remove('revision-body');
      };
    },
  });
}

function renderResultView(data = {}) {
  if (!resultDiv || !briefDiv || !emailOut) return;

  resultDiv.style.display = 'block';
  const resolvedModules = applyBriefModulesToView(
    data?.modules || briefModuleConfig || null,
    { updateControls: true },
  );
  mergeBriefData({ ...data, modules: resolvedModules });
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

function getHistorySortValue(entry) {
  if (!entry) return 0;
  const ts = entry.createdAt ? new Date(entry.createdAt).getTime() : NaN;
  if (!Number.isNaN(ts)) return ts;
  const idNum = Number(entry.id);
  return Number.isNaN(idNum) ? 0 : idNum;
}

function sortHistoryEntries(entries = []) {
  if (!Array.isArray(entries)) return [];
  return [...entries].sort((a, b) => getHistorySortValue(b) - getHistorySortValue(a));
}

function getBriefHistoryTitle(entry) {
  if (!entry) return 'Untitled brief';
  const custom = typeof entry.customTitle === 'string' ? entry.customTitle.trim() : '';
  if (custom) return custom;
  return entry?.request?.company || 'Untitled brief';
}

function getTargetHistoryTitle(entry) {
  if (!entry) return 'Untitled search';
  const custom = typeof entry.customTitle === 'string' ? entry.customTitle.trim() : '';
  if (custom) return custom;
  return entry?.request?.product || 'Untitled search';
}

function setHistorySidebarCollapsed(collapsed) {
  if (!navRail || !historyCollapseBtn) return;
  navRail.classList.toggle('expanded', !collapsed);
  historyCollapseBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
  historyCollapseBtn.innerHTML = SIDEBAR_ICON_HTML;
  try {
    window.localStorage?.setItem(HISTORY_SIDEBAR_COLLAPSED_KEY, collapsed ? '1' : '0');
  } catch (err) {
    // Ignore storage errors (private mode, etc.)
  }
}

function initHistorySidebarCollapse() {
  if (!navRail || !historyCollapseBtn) return;
  let startCollapsed = true;
  try {
    startCollapsed = window.localStorage?.getItem(HISTORY_SIDEBAR_COLLAPSED_KEY) === '1';
  } catch (err) {
    startCollapsed = true;
  }
  setHistorySidebarCollapsed(startCollapsed);
  historyCollapseBtn.addEventListener('click', () => {
    setHistorySidebarCollapsed(true);
  });
}

function getHistoryTypeFromMode(mode) {
  return mode === Mode.TARGET_GENERATION ? 'target' : 'research';
}

function handleHistoryRename(id, mode) {
  const source = mode === Mode.TARGET_GENERATION ? targetHistoryEntries : historyEntries;
  const entry = source.find(item => item.id === id);
  if (!entry) return;
  const currentTitle = mode === Mode.TARGET_GENERATION ? getTargetHistoryTitle(entry) : getBriefHistoryTitle(entry);
  const nextTitle = window.prompt('Rename this history entry', currentTitle);
  if (nextTitle === null) return;
  const trimmed = nextTitle.trim();
  chrome.runtime.sendMessage(
    { action: 'renameHistoryEntry', id, title: trimmed, historyType: getHistoryTypeFromMode(mode) },
    (resp) => {
      const err = chrome.runtime.lastError;
      if (err || !resp?.ok) {
        console.warn('Failed to rename history entry', err || resp?.error);
        return;
      }
      if (mode === Mode.TARGET_GENERATION) {
        loadTargetHistory({ selectEntryId: id, autoShow: false, updateForm: false, statusText: '' });
      } else {
        loadHistory({ selectEntryId: id, autoShow: false, updateForm: false, statusText: '' });
      }
    }
  );
}

function handleHistoryDelete(id, mode) {
  if (!id) return;
  const confirmed = window.confirm('Delete this history entry? This cannot be undone.');
  if (!confirmed) return;
  const isTarget = mode === Mode.TARGET_GENERATION;
  const wasActive = isTarget ? currentTargetHistoryId === id : currentHistoryId === id;
  if (isTarget && wasActive) {
    currentTargetHistoryId = null;
  } else if (!isTarget && wasActive) {
    currentHistoryId = null;
  }
  chrome.runtime.sendMessage(
    { action: 'deleteHistoryEntry', id, historyType: getHistoryTypeFromMode(mode) },
    (resp) => {
      const err = chrome.runtime.lastError;
      if (err || !resp?.ok) {
        console.warn('Failed to delete history entry', err || resp?.error);
        return;
      }
      const refreshOpts = { selectLatest: true, autoShow: wasActive, updateForm: false, statusText: '' };
      if (isTarget) {
        loadTargetHistory({ ...refreshOpts, selectEntryId: currentTargetHistoryId });
      } else {
        loadHistory({ ...refreshOpts, selectEntryId: currentHistoryId });
      }
    }
  );
}

function handleHistoryAction(action, id, mode) {
  if (!action || !id) return;
  if (action === 'rename') {
    handleHistoryRename(id, mode);
  } else if (action === 'delete') {
    handleHistoryDelete(id, mode);
  }
}

function triggerRailNewResearch() {
  startNewResearch(activeMode || Mode.TARGETED_BRIEF);
}

function closeAllHistoryMenus(container) {
  if (!container) return;
  const menus = container.querySelectorAll('.history-menu.open');
  menus.forEach((menu) => {
    menu.classList.remove('open');
    const parentItem = menu.closest('.history-item');
    parentItem?.classList.remove('menu-open');
    const trigger = menu.querySelector('[data-history-menu="trigger"]');
    const list = menu.querySelector('.history-menu-list');
    trigger?.setAttribute('aria-expanded', 'false');
    if (list) list.hidden = true;
  });
}

function toggleHistoryMenu(trigger, container) {
  if (!trigger) return;
  const menu = trigger.closest('.history-menu');
  if (!menu) return;
  const list = menu.querySelector('.history-menu-list');
  const item = menu.closest('.history-item');
  const willOpen = !menu.classList.contains('open');

  closeAllHistoryMenus(container);

  if (willOpen) {
    menu.classList.add('open');
    item?.classList.add('menu-open');
    trigger.setAttribute('aria-expanded', 'true');
    if (list) list.hidden = false;
  } else {
    menu.classList.remove('open');
    item?.classList.remove('menu-open');
    trigger.setAttribute('aria-expanded', 'false');
    if (list) list.hidden = true;
  }
}

function buildHistoryActionMenu(entryId) {
  const actions = document.createElement('div');
  actions.className = 'history-actions';

  const menuWrapper = document.createElement('div');
  menuWrapper.className = 'history-menu';

  const trigger = document.createElement('button');
  trigger.type = 'button';
  trigger.className = 'history-menu-trigger';
  trigger.dataset.historyMenu = 'trigger';
  trigger.dataset.id = entryId;
  trigger.setAttribute('aria-haspopup', 'true');
  trigger.setAttribute('aria-expanded', 'false');
  trigger.setAttribute('aria-label', 'Open history actions');
  trigger.textContent = '...';
  menuWrapper.appendChild(trigger);

  const list = document.createElement('div');
  list.className = 'history-menu-list';
  list.setAttribute('role', 'menu');
  list.hidden = true;

  const renameBtn = document.createElement('button');
  renameBtn.type = 'button';
  renameBtn.className = 'history-menu-item';
  renameBtn.dataset.historyAction = 'rename';
  renameBtn.dataset.id = entryId;
  renameBtn.setAttribute('role', 'menuitem');
  renameBtn.textContent = 'Rename';
  list.appendChild(renameBtn);

  const deleteBtn = document.createElement('button');
  deleteBtn.type = 'button';
  deleteBtn.className = 'history-menu-item destructive';
  deleteBtn.dataset.historyAction = 'delete';
  deleteBtn.dataset.id = entryId;
  deleteBtn.setAttribute('role', 'menuitem');
  deleteBtn.textContent = 'Delete';
  list.appendChild(deleteBtn);

  menuWrapper.appendChild(list);
  actions.appendChild(menuWrapper);
  return actions;
}

function renderHistory(entries, opts = {}) {
  if (!historyList || !historyEmpty) return null;

  historyEntries = sortHistoryEntries(entries);

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
    const item = document.createElement('div');
    item.className = 'history-item';
    item.setAttribute('role', 'listitem');
    item.dataset.id = entry.id;
    item.tabIndex = 0;
    if (entry.id === currentHistoryId) item.classList.add('active');

    const content = document.createElement('div');
    content.className = 'history-content';

    const title = document.createElement('span');
    title.className = 'history-title';
    title.textContent = getBriefHistoryTitle(entry);
    content.appendChild(title);

    const subtitleText = [entry?.request?.product, entry?.request?.location].filter(Boolean).join(' - ');
    if (subtitleText) {
      const subtitle = document.createElement('span');
      subtitle.className = 'history-subtitle';
      subtitle.textContent = subtitleText;
      content.appendChild(subtitle);
    }

    const meta = document.createElement('span');
    meta.className = 'history-meta';
    meta.textContent = formatHistoryTimestamp(entry.createdAt) || '';
    content.appendChild(meta);

    item.appendChild(content);
    item.appendChild(buildHistoryActionMenu(entry.id));
    historyList.appendChild(item);
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
    currentBriefRequest = {
      company: entry?.request?.company || '',
      location: entry?.request?.location || '',
      product: entry?.request?.product || '',
      docs: entry?.request?.docs || [],
      modules: entry?.request?.modules || entry?.result?.modules || getEffectiveBriefModules(),
    };
    briefVersions = [];
    briefVersionCounter = 0;
    updateBriefVersionSelect();
    updateBriefCommandBarFromRequest();
    showBriefBar();
    renderResultView(entry.result);
    briefVersionCounter = Math.max(briefVersionCounter, 1);
    updateBriefVersionPill();
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

  targetHistoryEntries = sortHistoryEntries(entries);

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
    const item = document.createElement('div');
    item.className = 'history-item';
    item.setAttribute('role', 'listitem');
    item.dataset.id = entry.id;
    item.tabIndex = 0;
    if (entry.id === currentTargetHistoryId) item.classList.add('active');

    const content = document.createElement('div');
    content.className = 'history-content';

    const title = document.createElement('span');
    title.className = 'history-title';
    title.textContent = getTargetHistoryTitle(entry);
    content.appendChild(title);

    const subtitleParts = [];
    if (entry?.request?.location) subtitleParts.push(entry.request.location);
    if (Array.isArray(entry?.result?.companies)) {
      subtitleParts.push(`${entry.result.companies.length} companies`);
    }
    if (subtitleParts.length) {
      const subtitle = document.createElement('span');
      subtitle.className = 'history-subtitle';
      subtitle.textContent = subtitleParts.join(' | ');
      content.appendChild(subtitle);
    }

    const meta = document.createElement('span');
    meta.className = 'history-meta';
    meta.textContent = formatHistoryTimestamp(entry.createdAt) || '';
    content.appendChild(meta);

    item.appendChild(content);
    item.appendChild(buildHistoryActionMenu(entry.id));
    targetHistoryList.appendChild(item);
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
  const request = buildBriefRequestFromForm();
  await runBriefGeneration(request);
});

historyList?.addEventListener('click', (evt) => {
  const menuTrigger = evt.target.closest('[data-history-menu="trigger"]');
  if (menuTrigger) {
    evt.preventDefault();
    evt.stopPropagation();
    toggleHistoryMenu(menuTrigger, historyList);
    return;
  }

  const actionBtn = evt.target.closest('[data-history-action]');
  if (actionBtn) {
    evt.preventDefault();
    evt.stopPropagation();
    closeAllHistoryMenus(historyList);
    const { historyAction: action, id } = actionBtn.dataset;
    handleHistoryAction(action, id, Mode.TARGETED_BRIEF);
    return;
  }
  const button = evt.target.closest('.history-item');
  if (!button) return;
  const { id } = button.dataset;
  if (!id) return;
  const entry = historyEntries.find(item => item.id === id);
  if (!entry) return;
  evt.preventDefault();
  closeAllHistoryMenus(historyList);
  setActiveHistoryItem(id);
  showHistoryEntry(entry, { updateForm: true });
});

historyList?.addEventListener('keydown', (evt) => {
  if (evt.defaultPrevented) return;
  if (evt.target.closest('[data-history-action]')) return;
  if (evt.target.closest('[data-history-menu="trigger"]')) return;
  const button = evt.target.closest('.history-item');
  if (!button) return;
  if (evt.key === 'Enter' || evt.key === ' ') {
    evt.preventDefault();
    const { id } = button.dataset;
    if (!id) return;
    const entry = historyEntries.find(item => item.id === id);
    if (!entry) return;
    setActiveHistoryItem(id);
    showHistoryEntry(entry, { updateForm: true });
  }
});

railHistoryBtn?.addEventListener('click', () => {
  openHistorySearchModal(activeMode || Mode.TARGETED_BRIEF);
});

railNewResearchBtn?.addEventListener('click', () => {
  triggerRailNewResearch();
});

railLogoBtn?.addEventListener('click', () => {
  setHistorySidebarCollapsed(false);
});

historySearchBtn?.addEventListener('click', () => {
  openHistorySearchModal(activeMode || Mode.TARGETED_BRIEF);
});

initHistorySidebarCollapse();
loadHistory({ selectLatest: true, autoShow: true, statusText: '' });

targetHistoryList?.addEventListener('click', (evt) => {
  const menuTrigger = evt.target.closest('[data-history-menu="trigger"]');
  if (menuTrigger) {
    evt.preventDefault();
    evt.stopPropagation();
    toggleHistoryMenu(menuTrigger, targetHistoryList);
    return;
  }

  const actionBtn = evt.target.closest('[data-history-action]');
  if (actionBtn) {
    evt.preventDefault();
    evt.stopPropagation();
    closeAllHistoryMenus(targetHistoryList);
    const { historyAction: action, id } = actionBtn.dataset;
    handleHistoryAction(action, id, Mode.TARGET_GENERATION);
    return;
  }
  const button = evt.target.closest('.history-item');
  if (!button) return;
  const { id } = button.dataset;
  if (!id) return;
  const entry = targetHistoryEntries.find(item => item.id === id);
  if (!entry) return;
  evt.preventDefault();
  closeAllHistoryMenus(targetHistoryList);
  setActiveTargetHistoryItem(id);
  showTargetHistoryEntry(entry, { updateForm: true });
});

targetHistoryList?.addEventListener('keydown', (evt) => {
  if (evt.defaultPrevented) return;
  if (evt.target.closest('[data-history-action]')) return;
  if (evt.target.closest('[data-history-menu="trigger"]')) return;
  const button = evt.target.closest('.history-item');
  if (!button) return;
  if (evt.key === 'Enter' || evt.key === ' ') {
    evt.preventDefault();
    const { id } = button.dataset;
    if (!id) return;
    const entry = targetHistoryEntries.find(item => item.id === id);
    if (!entry) return;
    setActiveTargetHistoryItem(id);
    showTargetHistoryEntry(entry, { updateForm: true });
  }
});

loadTargetHistory({ selectLatest: true, autoShow: true, statusText: '' });

document.addEventListener('click', (evt) => {
  const withinMenu = evt.target.closest('.history-menu');
  if (withinMenu) return;
  closeAllHistoryMenus(historyList);
  closeAllHistoryMenus(targetHistoryList);
});

briefBarToggle?.addEventListener('click', () => {
  toggleBriefBarCollapsed();
});

briefChipCompany?.addEventListener('click', () => showBriefForm());
briefChipLocation?.addEventListener('click', () => showBriefForm());
briefChipProduct?.addEventListener('click', () => showBriefForm());
briefChipDocs?.addEventListener('click', () => {
  showBriefForm();
  if (briefDocPickerBtn) {
    briefDocPickerBtn.click();
  }
});
briefChipSections?.addEventListener('click', () => {
  showBriefForm();
  openBriefSectionsModal({ mode: 'override' });
});

briefRegenerateBtn?.addEventListener('click', async () => {
  if (!currentBriefRequest.company || !currentBriefRequest.product) {
    setBriefStatusMessage('Set company and product before regenerating.', { error: true });
    toggleBriefEditDrawer();
    return;
  }
  await runBriefGeneration(currentBriefRequest);
});

briefVersionSelect?.addEventListener('change', (evt) => {
  const { value } = evt.target;
  if (!value) return;
  const version = briefVersions.find((v) => v.id === value);
  if (!version) return;
  currentBriefRequest = { ...(version.request || {}) };
  updateBriefCommandBarFromRequest();
  showBriefBar();
  renderResultView(version.result || {});
  setBriefStatusMessage(`Loaded ${value}.`, { error: false });
  if (briefVersionPill) {
    briefVersionPill.textContent = value;
    briefVersionPill.hidden = false;
  }
});

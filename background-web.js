(function backgroundWebScope() {
// Wrapped to avoid global name collisions when running service-worker logic in-page.
const EXPORT_TEMPLATE_KEY = "exportTemplate";
const EXPORT_TEMPLATES_KEY = "exportTemplates";
const TARGET_HISTORY_KEY = "targetHistory";
const PITCH_FROM_COMPANY_KEY = "pitchFromCompany";
const RESEARCH_HISTORY_KEY = "researchHistory";
const TARGET_COMPANY_GOAL = 100;
const TELEMETRY_QUEUE_KEY = "telemetryQueue";
const TELEMETRY_ENDPOINT_KEY = "telemetryEndpoint";
const TELEMETRY_DEFAULT_ENDPOINT = "https://jumblingly-turfiest-audie.ngrok-free.dev/logs";
const TELEMETRY_FLUSH_ALARM = "telemetryFlushAlarm";
const TELEMETRY_FLUSH_INTERVAL_MINUTES = 5;
const TELEMETRY_BATCH_SIZE = 10;
const TELEMETRY_QUEUE_LIMIT = 200;
const INSTALLATION_ID_KEY = "installationId";
const LAST_ACTIVE_PING_DATE_KEY = "lastActivePingDate";
const LLM_PROVIDER_STORAGE_KEY = "llmProvider";
const LLM_MODEL_STORAGE_KEY = "llmModel";
const GROQ_KEY_STORAGE_KEY = "groqKey";
const LLMProvider = {
  GEMINI: "gemini",
  GROQ: "groq",
};
const GROQ_COMPOUND_MODEL = "groq/compound";
const GROQ_COMPOUND_MINI_MODEL = "groq/compound-mini";
const DEFAULT_LLM_MODELS = {
  [LLMProvider.GEMINI]: "gemini-2.5-flash",
  [LLMProvider.GROQ]: GROQ_COMPOUND_MINI_MODEL,
};
const GROQ_LEGACY_MODEL = "gpt-oss-20b";
const GROQ_SECONDARY_MODEL = DEFAULT_LLM_MODELS[LLMProvider.GROQ];
const LLAMA_33_MODEL = "llama-3.3-70b-versatile";
const BRIEF_MODULES = {
  OVERVIEW: "overview",
  TOP_NEWS: "topNews",
  PERSONAS: "personas",
  EMAILS: "emails",
  TELE_PITCH: "telePitch",
};
const DEFAULT_BRIEF_MODULES = {
  [BRIEF_MODULES.OVERVIEW]: true,
  [BRIEF_MODULES.TOP_NEWS]: true,
  [BRIEF_MODULES.PERSONAS]: true,
  [BRIEF_MODULES.EMAILS]: true,
  [BRIEF_MODULES.TELE_PITCH]: true,
};

function loadPromptBuilders() {
  if (typeof self !== "undefined" && self.Prompts) return self.Prompts;

  if (typeof importScripts === "function") {
    try {
      importScripts("prompts.js");
      if (typeof self !== "undefined" && self.Prompts) return self.Prompts;
    } catch (err) {
      // Ignore missing prompts in environments without importScripts.
    }
  }

  if (typeof require === "function") {
    try {
      // eslint-disable-next-line global-require
      return require("./prompts");
    } catch (err) {
      // Ignore missing prompts in Node-less contexts.
    }
  }

  return {};
}

const promptBuilders = (() => {
  const loaded = loadPromptBuilders();
  if (loaded && Object.keys(loaded).length) return loaded;
  return {};
})();

function requirePromptBuilder(name) {
  const builder = promptBuilders && typeof promptBuilders[name] === "function" ? promptBuilders[name] : null;
  if (builder) return builder;
  throw new Error(`Prompt builder "${name}" is unavailable. Ensure prompts.js is loaded.`);
}

function normalizeBriefModules(value) {
  if (!value || typeof value !== "object") {
    return { ...DEFAULT_BRIEF_MODULES };
  }
  const next = { ...DEFAULT_BRIEF_MODULES };
  Object.keys(DEFAULT_BRIEF_MODULES).forEach((key) => {
    if (typeof value[key] === "boolean") {
      next[key] = value[key];
    }
  });
  return next;
}

function resolveBriefModules(value) {
  const next = normalizeBriefModules(value);
  if (next[BRIEF_MODULES.EMAILS]) {
    next[BRIEF_MODULES.PERSONAS] = true;
  }
  if (next[BRIEF_MODULES.TELE_PITCH]) {
    next[BRIEF_MODULES.PERSONAS] = true;
  }
  return next;
}

function getBriefProgressTotal(modules) {
  const resolved = resolveBriefModules(modules);
  return Object.keys(DEFAULT_BRIEF_MODULES).reduce((sum, key) => sum + (resolved[key] ? 1 : 0), 0);
}

function emitBriefProgress({ runId, current, total, label }) {
  if (!runId) return;
  try {
    chrome.runtime.sendMessage({
      action: "briefProgress",
      runId,
      current,
      total,
      label,
    });
  } catch (err) {
    // Ignore transient messaging errors
  }
}

function emitBriefPartialUpdate(runId, payload = {}) {
  if (!runId) return;
  try {
    chrome.runtime.sendMessage({
      action: "briefPartialUpdate",
      runId,
      payload,
    });
  } catch (err) {
    // Ignore transient messaging errors
  }
}

function normalizeLlmProvider(provider, hasGeminiKey, hasGroqKey) {
  if (provider === LLMProvider.GEMINI || provider === LLMProvider.GROQ) return provider;
  if (hasGroqKey) return LLMProvider.GROQ;
  if (hasGeminiKey) return LLMProvider.GEMINI;
  return LLMProvider.GROQ;
}

function normalizeLlmModel(model, provider) {
  if (typeof model === "string" && model.trim()) return model.trim();
  return DEFAULT_LLM_MODELS[provider] || DEFAULT_LLM_MODELS[LLMProvider.GEMINI];
}

function coerceModelForProvider(provider, model) {
  const fallback = DEFAULT_LLM_MODELS[provider] || DEFAULT_LLM_MODELS[LLMProvider.GEMINI];
  if (!model || typeof model !== "string") return fallback;
  const normalized = model.trim();
  if (provider === LLMProvider.GEMINI) {
    return normalized.toLowerCase().startsWith("gemini") ? normalized : fallback;
  }
  if (provider === LLMProvider.GROQ) {
    const lower = normalized.toLowerCase();
    const isLegacyGroq20b = lower.includes(GROQ_LEGACY_MODEL);
    // Fix misordered/unknown Groq ids like "openai/oss-gpt-20b" by upgrading to primary.
    if (isLegacyGroq20b || lower.includes("oss-gpt")) {
      return DEFAULT_LLM_MODELS[LLMProvider.GROQ];
    }
    return lower.startsWith("gemini") ? fallback : normalized;
  }
  return fallback;
}

async function loadLlmSettings() {
  try {
    const data = await chrome.storage.local.get([
      LLM_PROVIDER_STORAGE_KEY,
      LLM_MODEL_STORAGE_KEY,
      "geminiKey",
      GROQ_KEY_STORAGE_KEY,
    ]);
    const geminiKey = data && data.geminiKey ? String(data.geminiKey) : "";
    const groqKey = data && data[GROQ_KEY_STORAGE_KEY] ? String(data[GROQ_KEY_STORAGE_KEY]) : "";
    const provider = normalizeLlmProvider(data && data[LLM_PROVIDER_STORAGE_KEY], !!geminiKey, !!groqKey);
    let model = coerceModelForProvider(provider, normalizeLlmModel(data && data[LLM_MODEL_STORAGE_KEY], provider));
    // Promote old Groq default to the new primary model.
    if (
      provider === LLMProvider.GROQ &&
      typeof model === "string" &&
      model.toLowerCase().includes(GROQ_LEGACY_MODEL)
    ) {
      model = DEFAULT_LLM_MODELS[LLMProvider.GROQ];
    }
    return { provider, model, geminiKey, groqKey };
  } catch (err) {
    return {
      provider: LLMProvider.GEMINI,
      model: DEFAULT_LLM_MODELS[LLMProvider.GEMINI],
      geminiKey: "",
      groqKey: "",
    };
  }
}

async function callGeminiDirect(promptText, opts = {}, providerSettings = {}) {
  // Gemini is routed through Groq for consistency with the project-wide Groq usage.
  const candidateModel =
    coerceModelForProvider(
      LLMProvider.GROQ,
      (opts && opts.model && typeof opts.model === "string" ? opts.model : null) ||
        DEFAULT_LLM_MODELS[LLMProvider.GROQ]
    ) || DEFAULT_LLM_MODELS[LLMProvider.GROQ];
  const groqModel = candidateModel && candidateModel.toLowerCase().startsWith("openai/")
    ? candidateModel
    : "openai/gpt-oss-120b";

  const shimmedOpts = {
    ...opts,
    model: groqModel,
  };

  return callGroqDirect(promptText, shimmedOpts, providerSettings);
}

async function callGeminiWithRetryInternal(promptText, opts = {}, providerSettings = {}) {
  // Route through Groq retry logic for consistency; Gemini transport is deprecated here.
  return callGroqWithRetry(promptText, { ...opts }, providerSettings);
}

function normalizeGroqRole(role) {
  if (role === "model" || role === "assistant") return "assistant";
  if (role === "system") return "system";
  if (role === "tool") return "tool";
  return "user";
}

function partsToText(parts = []) {
  const arr = Array.isArray(parts) ? parts : [parts];
  return arr
    .map((part) => {
      if (typeof part === "string" && part.trim()) return part.trim();
      if (typeof part?.text === "string" && part.text.trim()) return part.text.trim();
      if (typeof part?.inlineData?.data === "string" && part.inlineData.data.trim()) {
        return decodeBase64Text(part.inlineData.data);
      }
      if (part?.functionResponse?.response) {
        try {
          return JSON.stringify(part.functionResponse.response);
        } catch (err) {
          return "";
        }
      }
      if (part?.functionCall?.args) {
        try {
          return JSON.stringify(part.functionCall.args);
        } catch (err) {
          return "";
        }
      }
      return "";
    })
    .filter(Boolean)
    .join("\n");
}

function extractInstructionText(instruction) {
  if (!instruction) return "";
  if (typeof instruction === "string") return instruction.trim();
  if (Array.isArray(instruction)) return partsToText(instruction);
  if (Array.isArray(instruction.parts)) return partsToText(instruction.parts);
  if (typeof instruction.text === "string") return instruction.text.trim();
  return "";
}

function buildGroqMessages({ promptText, contents, systemInstruction, contextText }) {
  const messages = [];
  const systemSegments = [];

  const normalizedInstruction = extractInstructionText(systemInstruction);
  if (normalizedInstruction) systemSegments.push(normalizedInstruction);
  if (typeof contextText === "string" && contextText.trim()) {
    systemSegments.push(contextText.trim());
  }

  const systemContent = systemSegments.filter(Boolean).join("\n\n");
  if (systemContent) {
    messages.push({ role: "system", content: systemContent });
  }

  if (Array.isArray(contents) && contents.length) {
    contents.forEach((item) => {
      const role = normalizeGroqRole(item && typeof item.role === "string" ? item.role : "user");
      const text = partsToText(item?.parts || item?.content || []);
      if (text) {
        messages.push({ role, content: text });
      }
    });
  }

  const trimmedPrompt = typeof promptText === "string" ? promptText.trim() : "";
  if (trimmedPrompt && !messages.some((msg) => msg.role === "user" && msg.content === trimmedPrompt)) {
    messages.push({ role: "user", content: trimmedPrompt });
  }

  if (!messages.length) {
    messages.push({ role: "user", content: trimmedPrompt || "" });
  }

  return messages;
}

function translateToolsToGroq(tools, disableDefaultTools) {
  if (disableDefaultTools) return { tools: [], choice: undefined };
  if (!Array.isArray(tools) || !tools.length) return { tools: [], choice: undefined };

  const alreadyOpenAi = tools.every((tool) => tool && typeof tool === "object" && tool.type);
  if (alreadyOpenAi) {
    return { tools, choice: "auto" };
  }

  const translated = [];
  let addedBrowserSearch = false;
  tools.forEach((tool) => {
    if (!tool || typeof tool !== "object") return;
    const hasGoogleSearch =
      Object.prototype.hasOwnProperty.call(tool, "google_search") ||
      Object.prototype.hasOwnProperty.call(tool, "googleSearch");
    const hasGoogleMaps =
      Object.prototype.hasOwnProperty.call(tool, "googleMaps") ||
      Object.prototype.hasOwnProperty.call(tool, "google_maps");

    if (hasGoogleSearch) {
      if (!addedBrowserSearch) {
        translated.push({ type: "browser_search" });
        addedBrowserSearch = true;
      }
    }

    if (hasGoogleMaps) {
      // No dedicated Maps tool; rely on browser_search for the page.
      if (!addedBrowserSearch) {
        translated.push({ type: "browser_search" });
        addedBrowserSearch = true;
      }
    }
  });

  return { tools: translated, choice: translated.length ? "auto" : undefined };
}

function parseRetryAfterMs(headers) {
  if (!headers || typeof headers.get !== "function") return null;
  const raw = headers.get("retry-after");
  if (!raw) return null;
  const num = Number(raw);
  if (Number.isFinite(num)) return Math.max(0, num * 1000);
  const ts = Date.parse(raw);
  return Number.isFinite(ts) ? Math.max(0, ts - Date.now()) : null;
}

function parseHeaderInt(headers, name) {
  if (!headers || typeof headers.get !== "function") return null;
  const raw = headers.get(name);
  if (raw === null || raw === undefined) return null;
  const num = Number(raw);
  return Number.isFinite(num) ? num : null;
}

function parseSecondsHeaderToMs(headers, name) {
  const val = parseHeaderInt(headers, name);
  if (val === null) return null;
  return Math.max(0, val * 1000);
}

function parseRateLimitHeaders(resp) {
  const headers = resp && resp.headers;
  if (!headers || typeof headers.get !== "function") return null;

  const rateLimit = {
    retryAfterMs: parseRetryAfterMs(headers),
    limitRequests: parseHeaderInt(headers, "x-ratelimit-limit-requests"),
    limitTokens: parseHeaderInt(headers, "x-ratelimit-limit-tokens"),
    remainingRequests: parseHeaderInt(headers, "x-ratelimit-remaining-requests"),
    remainingTokens: parseHeaderInt(headers, "x-ratelimit-remaining-tokens"),
    resetRequestsMs: parseSecondsHeaderToMs(headers, "x-ratelimit-reset-requests"),
    resetTokensMs: parseSecondsHeaderToMs(headers, "x-ratelimit-reset-tokens"),
  };

  const hasAny = Object.values(rateLimit).some((v) => v !== null && v !== undefined);
  return hasAny ? rateLimit : null;
}

function addJitter(ms) {
  if (!Number.isFinite(ms)) return 0;
  return Math.max(0, ms + Math.floor(Math.random() * 251));
}

async function callGroqDirect(promptText, opts = {}, providerSettings = {}) {
  const storedKey = providerSettings && providerSettings.groqKey ? providerSettings.groqKey : "";
  let groqKey = storedKey;
  if (!groqKey) {
    const data = await chrome.storage.local.get([GROQ_KEY_STORAGE_KEY]);
    groqKey = data && data[GROQ_KEY_STORAGE_KEY];
  }
  if (!groqKey) {
    return { error: "No Groq API key found. Please add it in the popup." };
  }

  const modelName =
    coerceModelForProvider(
      LLMProvider.GROQ,
      (opts && opts.model && typeof opts.model === "string" ? opts.model : null) ||
        (providerSettings && providerSettings.model) ||
        DEFAULT_LLM_MODELS[LLMProvider.GROQ]
    );

  const allowTools = modelName !== LLAMA_33_MODEL;
  const userCfg = opts.generationConfig || {};
  const temperature = Math.max(0, Math.min(userCfg.temperature ?? 0.3, 1));
  const maxTokens = Math.max(1, Math.min(userCfg.maxOutputTokens ?? userCfg.maxTokens ?? 4000, 8192));
  const wantsJson = userCfg.responseMimeType === "application/json";
  const messages = buildGroqMessages({
    promptText,
    contents: opts.contents,
    systemInstruction: opts.systemInstruction || opts.systemPrompt || opts.system,
    contextText: opts.context,
  });
  const { tools: groqTools, choice: groqToolChoice } = translateToolsToGroq(opts.tools, opts.disableDefaultTools);
  const topP = Math.max(0, Math.min(typeof userCfg.topP === "number" ? userCfg.topP : 1, 1));
  const stopSequences = Array.isArray(userCfg.stopSequences || userCfg.stop)
    ? (userCfg.stopSequences || userCfg.stop).filter((s) => typeof s === "string" && s.length)
    : null;
  const reasoningEffort = userCfg.reasoningEffort || opts.reasoningEffort;
  const stream = opts.stream === true || userCfg.stream === true ? true : false;
  const isCompoundModel = modelName === GROQ_COMPOUND_MODEL || modelName === GROQ_COMPOUND_MINI_MODEL;

  const body = {
    model: modelName,
    messages,
    temperature,
    top_p: topP,
    max_completion_tokens: maxTokens,
    stream,
  };

  if (opts.compoundCustom && typeof opts.compoundCustom === "object") {
    body.compound_custom = opts.compoundCustom;
    if (!wantsJson) {
      body.response_format = body.response_format || { type: "text" };
    }
  }

  if (allowTools && groqTools && groqTools.length) {
    const hasBrowserSearch = groqTools.some(
      (tool) => tool && typeof tool === "object" && (tool.type === "browser_search" || tool.type === "web_search")
    );
    if (isCompoundModel && hasBrowserSearch) {
      body.compound_custom = {
        tools: {
          enabled_tools: ["web_search"],
        },
      };
      // Keep response as text unless JSON explicitly requested below.
      if (!wantsJson) {
        body.response_format = { type: "text" };
      }
    } else {
      body.tools = groqTools;
      body.tool_choice = opts.toolChoice || groqToolChoice || "auto";
    }
  }

  if (wantsJson) {
    body.response_format = { type: "json_object" };
  }
  if (stopSequences && stopSequences.length) {
    body.stop = stopSequences;
  }
  if (reasoningEffort) {
    body.reasoning_effort = reasoningEffort;
  }

  try {
    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${groqKey}`,
    };

    const resp = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });

    const rateLimit = parseRateLimitHeaders(resp);
    const respText = await resp.text();
    let respJson = null;
    try {
      respJson = JSON.parse(respText);
    } catch (err) {
      if (!resp.ok) {
        return { error: `Groq API error (Status ${resp.status}): ${respText}`, status: resp.status, rateLimit };
      }
      return { error: "Failed to parse Groq API response as JSON.", details: respText, status: resp.status, rateLimit };
    }

    if (!resp.ok) {
      const message = respJson?.error?.message || respText;
      return { error: `Groq API error: ${message}`, details: respJson, status: resp.status, rateLimit };
    }

    let statusCode = resp.status;
    const message = respJson?.choices?.[0]?.message || {};
    const content = message.content;
    let outputText = "";
    if (typeof content === "string" && content.trim()) {
      outputText = content.trim();
    } else if (Array.isArray(content)) {
      outputText = content
        .map((part) => (typeof part?.text === "string" ? part.text : ""))
        .filter(Boolean)
        .join(" ")
        .trim();
    }

    const toolCalls = Array.isArray(message.tool_calls) ? message.tool_calls : [];

    if (!outputText) {
      if (toolCalls.length) {
        return {
          error: "Model returned tool calls without textual content.",
          details: respJson || { tool_calls: toolCalls },
          status: statusCode,
          rateLimit,
        };
      }
      return { error: "Could not find text in Groq response.", details: respJson, status: statusCode, rateLimit };
    }

    return { ok: true, text: outputText, raw: respJson, status: statusCode, rateLimit };
  } catch (err) {
    return { error: `Network request failed: ${String(err)}` };
  }
}

async function callGroqWithRetry(promptText, opts = {}, providerSettings = {}) {
  const baseConfig = opts.generationConfig || {};
  const primaryModel =
    opts && opts.model
      ? coerceModelForProvider(LLMProvider.GROQ, opts.model)
      : coerceModelForProvider(LLMProvider.GROQ, providerSettings.model || DEFAULT_LLM_MODELS[LLMProvider.GROQ]);
  const secondaryModelOverride =
    opts && typeof opts.secondaryModel === "string" && opts.secondaryModel.trim()
      ? coerceModelForProvider(LLMProvider.GROQ, opts.secondaryModel)
      : null;
  const secondaryModel =
    secondaryModelOverride ||
    (primaryModel === GROQ_SECONDARY_MODEL ? DEFAULT_LLM_MODELS[LLMProvider.GROQ] : GROQ_SECONDARY_MODEL);

  const attempts = [];
  const boostedMaxTokens = Math.min(
    typeof baseConfig.maxOutputTokens === "number" ? Math.round(baseConfig.maxOutputTokens * 1.2) : 4800,
    16000
  );
  const maxAttempts = Math.max(1, Math.min(opts.maxAttempts || 3, 5));
  let currentModel = primaryModel;
  let attemptIndex = 0;
  let lastError = null;

  const computeWaitMs = (rateLimit, status) => {
    if (!rateLimit) return null;
    const MAX_WAIT_MS = 60000;

    let waitMs = null;
    if (status === 429 && Number.isFinite(rateLimit.retryAfterMs)) {
      waitMs = rateLimit.retryAfterMs;
    }

    if (waitMs === null) {
      const resets = []
        .concat(Number.isFinite(rateLimit.resetTokensMs) ? [rateLimit.resetTokensMs] : [])
        .concat(Number.isFinite(rateLimit.resetRequestsMs) ? [rateLimit.resetRequestsMs] : []);
      if (resets.length) {
        waitMs = Math.min(...resets);
      }
    }

    if (waitMs === null && status === 429) {
      waitMs = 300;
    }

    if (waitMs === null) {
      const lowTokens = Number.isFinite(rateLimit.remainingTokens) && rateLimit.remainingTokens <= 1;
      const lowRequests = Number.isFinite(rateLimit.remainingRequests) && rateLimit.remainingRequests <= 1;
      const proactiveResets = []
        .concat(lowTokens && Number.isFinite(rateLimit.resetTokensMs) ? [rateLimit.resetTokensMs] : [])
        .concat(lowRequests && Number.isFinite(rateLimit.resetRequestsMs) ? [rateLimit.resetRequestsMs] : []);
      if (proactiveResets.length) {
        waitMs = Math.min(...proactiveResets);
      }
    }

    if (waitMs === null || !Number.isFinite(waitMs)) return null;
    return Math.min(MAX_WAIT_MS, Math.max(0, waitMs));
  };

  while (attemptIndex < maxAttempts) {
    const useBoostedConfig = attemptIndex > 0;
    const generationConfig = useBoostedConfig
      ? {
          ...baseConfig,
          temperature: Math.max(0, Math.min(baseConfig.temperature ?? 0.3, 0.7)),
          maxOutputTokens: boostedMaxTokens,
        }
      : baseConfig;

    const resp = await callGroqDirect(
      promptText,
      {
        ...opts,
        model: currentModel,
        generationConfig,
      },
      { ...providerSettings, model: currentModel }
    );

    if (!resp.error) {
      return attemptIndex === 0 ? resp : { ...resp, attempts };
    }

    attempts.push({
      label:
        attemptIndex === 0
          ? "primary"
          : resp.status === 429 && currentModel !== primaryModel
            ? "retry_model_switch"
            : "retry",
      error: resp.error,
      details: resp.details || resp.raw || resp.rateLimit || null,
      status: resp.status,
      rateLimit: resp.rateLimit,
    });

    lastError = resp;
    attemptIndex += 1;
    if (attemptIndex >= maxAttempts) break;

    const waitMs = addJitter(computeWaitMs(resp.rateLimit, resp.status));
    if (Number.isFinite(waitMs) && waitMs > 0) {
      await new Promise((resolve) => setTimeout(resolve, waitMs));
    }

    const shouldSwitchModel = resp.status === 429 && secondaryModel && secondaryModel !== currentModel;
    currentModel = shouldSwitchModel ? secondaryModel : currentModel;
  }

  return { ...lastError, attempts };
}

async function callLlmWithRetry(promptText, opts = {}) {
  const settings = await loadLlmSettings();
  const providerOverride = opts && opts.providerOverride;
  const provider = normalizeLlmProvider(providerOverride, !!settings.geminiKey, !!settings.groqKey);
  const model = coerceModelForProvider(provider, opts && opts.model ? opts.model : settings.model);
  const providerSettings = { ...settings, model };

  // Force Groq format for all calls to avoid legacy Gemini transport.
  const sanitizedOpts =
    provider === LLMProvider.GROQ
      ? opts
      : {
          ...opts,
          providerOverride: LLMProvider.GROQ,
        };

  return callGroqWithRetry(promptText, sanitizedOpts, providerSettings);
}

function extractJsonFromText(s) {
  if (!s || typeof s !== "string") return null;
  
  const jsonMatch = s.match(/```json\s*([\s\S]*?)\s*```/);
  if (jsonMatch && jsonMatch[1]) {
    s = jsonMatch[1];
  }

  const first = s.indexOf('{');
  const last = s.lastIndexOf('}');
  if (first === -1 || last === -1 || last <= first) return null;
  
  const sub = s.substring(first, last + 1);
  try {
    return JSON.parse(sub);
  } catch (e) {
    const cleaned = sub.replace(/[\u2018\u2019\u201C\u201D]/g, '"').replace(/,\s*}/g, '}').replace(/,\s*]/g, ']');
    try { return JSON.parse(cleaned); } catch (e2) { return null; }
  }
}

function parseModelJsonResponse(resp) {
  const tryParseText = (text) => {
    if (!text) return null;
    const structured = extractJsonFromText(text);
    if (structured) return structured;
    try {
      return JSON.parse(text);
    } catch (err) {
      return null;
    }
  };

  let rawText = typeof resp?.text === "string" ? resp.text : "";
  let parsed = tryParseText(rawText);

  if (!parsed && resp?.raw?.candidates?.[0]?.content?.parts?.length) {
    const combined = resp.raw.candidates[0].content.parts
      .map((part) => {
        if (typeof part?.text === "string" && part.text.trim()) return part.text;
        if (part?.functionCall?.args) {
          try {
            return JSON.stringify(part.functionCall.args);
          } catch (err) {
            return null;
          }
        }
        if (part?.functionResponse?.response) {
          try {
            return JSON.stringify(part.functionResponse.response);
          } catch (err) {
            return null;
          }
        }
        if (part?.inlineData?.mimeType === "application/json" && part?.inlineData?.data) {
          const decoded = decodeBase64Text(part.inlineData.data);
          return decoded || null;
        }
        return null;
      })
      .filter(Boolean)
      .join("\n");
    if (combined) {
      rawText = rawText || combined;
      parsed = tryParseText(combined);
    }
  }

  if (!parsed && resp?.raw?.choices?.[0]?.message?.content) {
    const content = resp.raw.choices[0].message.content;
    if (typeof content === "string" && content.trim()) {
      rawText = rawText || content;
      parsed = tryParseText(content);
    }
  }

  return { parsed, rawText };
}

function stripMarkdownFences(text = "") {
  if (typeof text !== "string") return "";
  return text.replace(/```[\s\S]*?```/g, (block) => {
    const inner = block.replace(/^```\w*\s*/, "").replace(/```$/, "");
    return inner;
  });
}

function splitMarkdownLines(markdown = "") {
  if (typeof markdown !== "string") return [];
  return stripMarkdownFences(markdown)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length);
}

function parseKeyValueSegments(segment = "") {
  const out = {};
  if (!segment || typeof segment !== "string") return out;
  segment
    .split(/;\s*/)
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((part) => {
      const kv = part.split(/[:=]/);
      if (kv.length < 2) return;
      const key = kv.shift().trim().toLowerCase();
      const value = kv.join(":").trim();
      if (key) out[key] = value;
    });
  return out;
}

function parsePersonasFromMarkdown(markdown = "", fallbackCompany = "", productHint = "") {
  const lines = splitMarkdownLines(markdown);
  const personas = [];
  lines.forEach((raw, idx) => {
    const line = raw.replace(/^-+\s*/, "").trim();
    if (!line) return;
    const kv = parseKeyValueSegments(line);
    const name = kv.name || kv.persona || `Persona ${idx + 1}`;
    const designation = kv.title || kv.designation || "";
    const department = kv.department || kv.dept || "";
    const rawLinkedinKeywords =
      kv.linkedin_keywords || kv.linkedinsearch || kv.linkedinquery || kv.linkedin || kv.linkedinkeywords || "";
    // We ignore any provided search links that may contain personal names and regenerate using only company + position.
    const link =
      kv.searchlink ||
      kv.search_link ||
      kv.zoominfo ||
      kv.zoominfo_link ||
      kv.zoom ||
      kv.link ||
      "";
    let linkedinKeywords = normalizeLinkedInKeywords(
      rawLinkedinKeywords,
      fallbackCompany,
      designation || department || productHint || ""
    );
    if (!linkedinKeywords) {
      linkedinKeywords = buildLinkedInKeywordFallback({ name, designation, department }, fallbackCompany, productHint);
    }
    personas.push({
      name,
      designation,
      department,
      linkedin_keywords: linkedinKeywords || "",
      linkedin_search_url: toLinkedInPeopleSearchUrl(linkedinKeywords),
      zoominfo_link: buildZoomInfoSearchLink({ designation, department }, fallbackCompany) || link,
    });
  });
  return personas;
}

function parsePersonaEmailMarkdown(markdown = "", fallbackPersona = {}) {
  const lines = stripMarkdownFences(markdown).split(/\r?\n/);
  let personaLineIdx = -1;
  let subjectLineIdx = -1;
  let bodyLineIdx = -1;
  lines.forEach((raw, idx) => {
    const line = (raw || "").trim();
    if (personaLineIdx === -1 && /^persona\s*[:=]/i.test(line)) personaLineIdx = idx;
    if (subjectLineIdx === -1 && /^subject\s*[:=]/i.test(line)) subjectLineIdx = idx;
    if (bodyLineIdx === -1 && /^body\s*[:=]/i.test(line)) bodyLineIdx = idx;
  });

  const personaLine = personaLineIdx >= 0 ? lines[personaLineIdx].trim() : "";
  const personaLineClean = personaLine.replace(/^persona\s*[:=]\s*/i, "").trim();
  const personaKv = parseKeyValueSegments(personaLineClean);
  if (!personaKv.name && personaLineClean) {
    const firstSegment = personaLineClean.split(/;\s*/)[0];
    if (firstSegment) personaKv.name = firstSegment.trim();
  }
  const personaName =
    personaKv.name ||
    personaKv.persona ||
    fallbackPersona.name ||
    fallbackPersona.personaName ||
    fallbackPersona.persona_name ||
    "";
  const personaDesignation = personaKv.title || personaKv.designation || fallbackPersona.designation || "";
  const personaDepartment = personaKv.department || personaKv.dept || fallbackPersona.department || "";

  const subjectLine = subjectLineIdx >= 0 ? lines[subjectLineIdx].trim() : "";
  const subject = subjectLine.replace(/^subject\s*[:=]\s*/i, "").trim();

  let body = "";
  if (bodyLineIdx >= 0) {
    body = lines.slice(bodyLineIdx + 1).join("\n").trim();
  } else if (subjectLineIdx >= 0) {
    body = lines.slice(subjectLineIdx + 1).join("\n").trim();
  } else {
    body = lines.join("\n").trim();
  }

  return {
    personaName: personaName || fallbackPersona.name || "",
    personaDesignation,
    personaDepartment,
    subject,
    body,
  };
}

function parseTelephonicPitchMarkdown(markdown = "", fallbackPersona = {}) {
  const lines = stripMarkdownFences(markdown).split(/\r?\n/);
  let personaLineIdx = -1;
  let pitchLineIdx = -1;
  lines.forEach((raw, idx) => {
    const line = (raw || "").trim();
    if (personaLineIdx === -1 && /^persona\s*[:=]/i.test(line)) personaLineIdx = idx;
    if (pitchLineIdx === -1 && /^telephonic\s*pitch\s*[:=]?/i.test(line)) pitchLineIdx = idx;
  });

  const personaLine = personaLineIdx >= 0 ? lines[personaLineIdx].trim() : "";
  const personaLineClean = personaLine.replace(/^persona\s*[:=]\s*/i, "").trim();
  const personaKv = parseKeyValueSegments(personaLineClean);
  if (!personaKv.name && personaLineClean) {
    const firstSegment = personaLineClean.split(/;\s*/)[0];
    if (firstSegment) personaKv.name = firstSegment.trim();
  }
  const personaName =
    personaKv.name ||
    personaKv.persona ||
    fallbackPersona.name ||
    fallbackPersona.personaName ||
    fallbackPersona.persona_name ||
    "";
  const personaDesignation = personaKv.title || personaKv.designation || fallbackPersona.designation || "";
  const personaDepartment = personaKv.department || personaKv.dept || fallbackPersona.department || "";

  let script = "";
  if (pitchLineIdx >= 0) {
    script = lines.slice(pitchLineIdx + 1).join("\n").trim();
  } else if (personaLineIdx >= 0) {
    script = lines.slice(personaLineIdx + 1).join("\n").trim();
  } else {
    script = lines.join("\n").trim();
  }

  return {
    personaName: personaName || fallbackPersona.name || "",
    personaDesignation,
    personaDepartment,
    script,
  };
}

function parseHqLocationMarkdown(markdown = "") {
  const lines = splitMarkdownLines(markdown);
  for (const line of lines) {
    const match = line.match(/hq\s*location\s*[:=]\s*(.+)/i);
    if (match && match[1]) return match[1].trim();
  }
  return lines[0] || "";
}

function parseRevenueSectorMarkdown(markdown = "", fallbackCompany = "") {
  const lines = splitMarkdownLines(markdown);
  const result = {
    company_name: "",
    revenue_estimate: "",
    industry_sector: "",
  };
  lines.forEach((line) => {
    if (/^company\s*[:=]/i.test(line)) {
      result.company_name = line.replace(/^company\s*[:=]\s*/i, "").trim();
    } else if (/^revenue\s*[:=]/i.test(line)) {
      result.revenue_estimate = line.replace(/^revenue\s*[:=]\s*/i, "").trim();
    } else if (/^sector\s*[:=]/i.test(line) || /^industry\s*[:=]/i.test(line)) {
      result.industry_sector = line.replace(/^(sector|industry)\s*[:=]\s*/i, "").trim();
    }
  });
  if (!result.company_name) result.company_name = fallbackCompany || "";
  return result;
}

function parseTopNewsMarkdown(markdown = "") {
  const lines = stripMarkdownFences(markdown)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const items = [];
  let pending = null;

  const flushPending = () => {
    if (pending && pending.title) {
      items.push({ title: pending.title, summary: pending.summary || "" });
    }
    pending = null;
  };

  lines.forEach((raw) => {
    const line = raw.replace(/^[-*]\s*/, "").replace(/^\d+[\).\s]+/, "").trim();
    const kv = parseKeyValueSegments(line);
    const title = kv.title || kv.headline || "";
    const summary = kv.summary || kv.description || kv.desc || "";

    if (title && summary) {
      flushPending();
      items.push({ title, summary });
      return;
    }

    if (title) {
      flushPending();
      pending = { title, summary: "" };
      return;
    }

    if (summary && pending && !pending.summary) {
      pending.summary = summary;
      flushPending();
      return;
    }

    const titleMatch = line.match(/^title\s*[:=]\s*(.+)$/i);
    const summaryMatch = line.match(/^(summary|description)\s*[:=]\s*(.+)$/i);
    if (titleMatch) {
      flushPending();
      pending = { title: titleMatch[1].trim(), summary: "" };
      return;
    }
    if (summaryMatch && pending && !pending.summary) {
      pending.summary = summaryMatch[2].trim();
      flushPending();
      return;
    }
  });

  flushPending();
  return items;
}

function parseTargetCompaniesMarkdown(markdown = "") {
  const lines = stripMarkdownFences(markdown)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const companies = [];
  lines.forEach((raw) => {
    const line = raw.replace(/^[-*]\s*/, "").replace(/^\d+[\).\s]+/, "").trim();
    if (!line) return;
    const kv = parseKeyValueSegments(line);
    const name = kv.name || kv.company || "";
    const website = kv.website || kv.url || kv.site || "";
    const revenue = kv.revenue || kv.revenue_estimate || "";
    const notes = kv.notes || kv.note || kv.summary || "";
    if (name) {
      companies.push({
        name,
        website,
        revenue,
        notes,
      });
    }
  });
  return companies;
}

function parseMarkdownTable(markdown = "") {
  const clean = stripMarkdownFences(markdown);
  const lines = clean
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (lines.length < 2) return { headers: [], rows: [] };
  const headerIdx = lines.findIndex((line) => /^\|/.test(line) && line.includes("|"));
  if (headerIdx === -1 || headerIdx + 1 >= lines.length) return { headers: [], rows: [] };
  const headerLine = lines[headerIdx];
  const separatorLine = lines[headerIdx + 1];
  if (!/^-*:?[-| ]+$/.test(separatorLine.replace(/\|/g, ""))) {
    return { headers: [], rows: [] };
  }
  const headers = headerLine
    .split("|")
    .map((h) => h.trim())
    .filter(Boolean);
  if (!headers.length) return { headers: [], rows: [] };
  const rowLines = lines.slice(headerIdx + 2).filter((l) => l.includes("|"));
  const rows = rowLines.map((line) => {
    const cells = line
      .split("|")
      .map((c) => c.trim())
      .filter((_, idx) => idx > 0 && idx <= headers.length);
    const row = {};
    headers.forEach((h, idx) => {
      row[h] = cells[idx] || "";
    });
    return row;
  });
  return { headers, rows };
}

function escapeRegex(str = "") {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function sanitizeLinkedInKeywordString(raw = "") {
  return String(raw || "")
    .replace(/[“”"]/g, "")
    .replace(/\b(?:OR|AND)\b/gi, " ")
    .replace(/[|/]/g, " ")
    .replace(/[,;]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeLinkedInKeywords(rawKeywords = "", companyName = "", positionHint = "") {
  const company = sanitizeLinkedInKeywordString(companyName);
  const parts = String(rawKeywords || "")
    .split(/(?:\bOR\b|\bAND\b|[,/|;])/i)
    .map(sanitizeLinkedInKeywordString)
    .filter(Boolean);

  let primary = parts.length ? parts[0] : sanitizeLinkedInKeywordString(positionHint);

  if (company && primary) {
    const companyRe = new RegExp(`\\b${escapeRegex(company)}\\b`, "ig");
    primary = sanitizeLinkedInKeywordString(primary.replace(companyRe, ""));
  }

  if (!primary) primary = sanitizeLinkedInKeywordString(positionHint);

  const keywords = [company, primary].filter(Boolean).join(" ").trim();
  return sanitizeLinkedInKeywordString(keywords);
}

function toLinkedInPeopleSearchUrl(searchString) {
  const q = sanitizeLinkedInKeywordString(searchString ?? "");
  if (!q) {
    return "https://www.linkedin.com/search/results/people/?keywords=&origin=SWITCH_SEARCH_VERTICAL";
  }

  // LinkedIn uses a URL-encoded "keywords" param (spaces become %20).
  const keywords = encodeURIComponent(q);

  return `https://www.linkedin.com/search/results/people/?keywords=${keywords}&origin=SWITCH_SEARCH_VERTICAL`;
}

function buildLinkedInKeywordFallback(persona = {}, companyName = "", productName = "") {
  // Keep LinkedIn people search keywords impersonal: use role/department, never the individual's name.
  const positionHint = persona.designation || persona.department || productName || "";
  return normalizeLinkedInKeywords(positionHint, companyName, positionHint);
}

function buildZoomInfoSearchLink(persona = {}, companyName = "") {
  // Build a Google query that only contains the company name and the persona's role/department (no personal names).
  const parts = [];
  const role = typeof persona.designation === "string" ? persona.designation.trim() : "";
  const dept = typeof persona.department === "string" ? persona.department.trim() : "";
  const company = typeof companyName === "string" ? companyName.trim() : "";

  if (company) parts.push(`"${company}"`);
  if (role) parts.push(`"${role}"`);
  else if (dept) parts.push(`"${dept}"`);

  const scope = `site:linkedin.com/in`;

  const query = `${scope} ${parts.join(" ")}`.trim();

  return `https://www.google.com/search?q=${encodeURIComponent(query)}&num=20`;
}

function decodeBase64Text(b64) {
  if (!b64) return "";
  try {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const decoder = new TextDecoder('utf-8', { fatal: false });
    return decoder.decode(bytes);
  } catch (e) {
    return "";
  }
}

function normalizePitchingCompany(raw) {
  if (!raw || typeof raw !== "string") return "";
  return raw.trim();
}

async function loadPitchingCompany() {
  try {
    const data = await chrome.storage.local.get([PITCH_FROM_COMPANY_KEY]);
    const val = data && data[PITCH_FROM_COMPANY_KEY];
    return normalizePitchingCompany(val);
  } catch (err) {
    console.warn("Failed to load pitching company setting", err);
    return "";
  }
}

function getValueAtPath(obj, path) {
  if (!path) return undefined;
  const segments = path.replace(/\[(\d+)\]/g, ".$1").split(".");
  return segments.reduce((acc, key) => {
    if (acc === undefined || acc === null) return undefined;
    return acc[key];
  }, obj);
}

function truncateText(str, max = 3000) {
  if (typeof str !== "string") return "";
  if (str.length <= max) return str;
  return str.substring(0, max) + "...";
}

let telemetryFlushInFlight = false;

async function sendTelemetryBatch(metrics, endpointOverride) {
  if (!Array.isArray(metrics) || !metrics.length) return false;
  const endpoint = endpointOverride || (await getTelemetryEndpoint());
  if (!endpoint) {
    console.warn("Telemetry endpoint missing; cannot send metrics");
    return false;
  }

  try {
    const resp = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ metrics }),
      credentials: "omit",
    });

    if (!resp.ok) {
      console.warn(`Telemetry send failed`, { status: resp.status, statusText: resp.statusText });
      return false;
    }
    return true;
  } catch (err) {
    console.warn("Telemetry send error", err);
    return false;
  }
}

function buildResearchCycleMetric({ startTimeMs, endTimeMs, request = {}, runId = "", success = true, error = "" }) {
  const safeStart = typeof startTimeMs === "number" ? startTimeMs : Date.now();
  const safeEnd = typeof endTimeMs === "number" ? endTimeMs : Date.now();
  const durationMs = Math.max(0, safeEnd - safeStart);
  const requestSummary = {
    companyLength: typeof request.company === "string" ? request.company.length : 0,
    locationLength: typeof request.location === "string" ? request.location.length : 0,
    productLength: typeof request.product === "string" ? request.product.length : 0,
    docCount: Array.isArray(request.docs) ? request.docs.length : 0,
  };

  return {
    kind: "research_cycle_time",
    durationMs,
    startedAt: new Date(safeStart).toISOString(),
    finishedAt: new Date(safeEnd).toISOString(),
    success: !!success,
    runId: runId || "",
    request: requestSummary,
    error: success ? "" : truncateText(error || "", 200),
  };
}

function buildGenerationErrorMetric({ feature = "unknown", error = "", request = {}, runId = "" }) {
  const normalizedFeature = feature === "targets" ? "targets" : "briefs";
  return {
    kind: "generation_error",
    feature: normalizedFeature,
    occurredAt: new Date().toISOString(),
    runId: runId || "",
    error: truncateText(error || "", 300),
    request: {
      companyLength: typeof request.company === "string" ? request.company.length : 0,
      locationLength: typeof request.location === "string" ? request.location.length : 0,
      productLength: typeof request.product === "string" ? request.product.length : 0,
      docCount: Array.isArray(request.docs) ? request.docs.length : 0,
      sectorCount: Array.isArray(request.sectors) ? request.sectors.filter(Boolean).length : 0,
    },
  };
}

function buildSearchToExportConversionMetric({
  exportId,
  occurredAtMs,
  selectionType = "all",
  selectionCount = 0,
  format = "",
  templateId = "",
  researchEntryIds = [],
  resultRowCount = 0,
}) {
  const timestamp = typeof occurredAtMs === "number" ? occurredAtMs : Date.now();
  const normalizedSelectionType = ["all", "date", "custom"].includes(selectionType) ? selectionType : "all";
  const ids = Array.isArray(researchEntryIds)
    ? researchEntryIds.map((id) => (id == null ? "" : String(id))).filter(Boolean)
    : [];

  return {
    kind: "search_to_export_conversion",
    occurredAt: new Date(timestamp).toISOString(),
    export: {
      id: exportId ? String(exportId) : `export-${timestamp}`,
      format: format === "md" ? "md" : "xlsx",
      selectionType: normalizedSelectionType,
      selectionCount: Math.max(0, Number(selectionCount) || 0),
      templateId: templateId ? String(templateId) : "",
      resultRowCount: Math.max(0, Number(resultRowCount) || 0),
    },
    searchContext: {
      researchEntryIds: ids,
    },
  };
}

async function getTelemetryEndpoint() {
  const stored = await chrome.storage.local.get([TELEMETRY_ENDPOINT_KEY]);
  const storedEndpoint =
    stored && typeof stored[TELEMETRY_ENDPOINT_KEY] === "string"
      ? stored[TELEMETRY_ENDPOINT_KEY].trim()
      : "";
  return storedEndpoint || TELEMETRY_DEFAULT_ENDPOINT;
}

function generateInstallId() {
  try {
    if (typeof crypto !== "undefined") {
      if (typeof crypto.randomUUID === "function") {
        return crypto.randomUUID();
      }
      if (typeof crypto.getRandomValues === "function") {
        const buf = new Uint8Array(16);
        crypto.getRandomValues(buf);
        return `inst-${Array.from(buf)
          .map((b) => b.toString(16).padStart(2, "0"))
          .join("")}`;
      }
    }
  } catch (err) {
    // Fall through to time/random based id
  }
  return `inst-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

async function getOrCreateInstallationId() {
  const stored = await chrome.storage.local.get([INSTALLATION_ID_KEY]);
  const existingId = typeof stored[INSTALLATION_ID_KEY] === "string" ? stored[INSTALLATION_ID_KEY].trim() : "";
  if (existingId) return existingId;

  const newId = generateInstallId();
  await chrome.storage.local.set({ [INSTALLATION_ID_KEY]: newId });
  return newId;
}

function buildWeeklyActiveUserMetric({ installationId, occurredAtMs, trigger = "unknown" }) {
  const ts = typeof occurredAtMs === "number" ? occurredAtMs : Date.now();
  const version = chrome.runtime?.getManifest?.().version || "unknown";
  return {
    kind: "weekly_active_user",
    occurredAt: new Date(ts).toISOString(),
    installationId: installationId || "unknown",
    trigger,
    appVersion: version,
  };
}

async function markWeeklyActiveUser(trigger = "unknown") {
  try {
    const todayKey = new Date().toISOString().slice(0, 10);
    const stored = await chrome.storage.local.get([LAST_ACTIVE_PING_DATE_KEY]);
    if (stored && stored[LAST_ACTIVE_PING_DATE_KEY] === todayKey) return;

    const installationId = await getOrCreateInstallationId();
    const metric = buildWeeklyActiveUserMetric({
      installationId,
      occurredAtMs: Date.now(),
      trigger,
    });
    await chrome.storage.local.set({ [LAST_ACTIVE_PING_DATE_KEY]: todayKey });
    enqueueTelemetryMetric(metric).catch(() => {});
    flushTelemetryQueue().catch(() => {});
  } catch (err) {
    console.warn("Failed to record weekly active user metric", err);
  }
}

async function enqueueTelemetryMetric(metric) {
  if (!metric || typeof metric !== "object") return;
  try {
    const stored = await chrome.storage.local.get([TELEMETRY_QUEUE_KEY]);
    const queue = Array.isArray(stored[TELEMETRY_QUEUE_KEY]) ? stored[TELEMETRY_QUEUE_KEY] : [];
    queue.push(metric);
    const trimmed = queue.slice(-TELEMETRY_QUEUE_LIMIT);
    await chrome.storage.local.set({ [TELEMETRY_QUEUE_KEY]: trimmed });
    if (trimmed.length >= TELEMETRY_BATCH_SIZE) {
      flushTelemetryQueue().catch(() => {});
    }
  } catch (err) {
    console.warn("Failed to enqueue telemetry metric", err);
  }
}

async function flushTelemetryQueue() {
  if (telemetryFlushInFlight) return;
  telemetryFlushInFlight = true;
  try {
    const stored = await chrome.storage.local.get([TELEMETRY_QUEUE_KEY]);
    const queue = Array.isArray(stored[TELEMETRY_QUEUE_KEY]) ? stored[TELEMETRY_QUEUE_KEY] : [];
    if (!queue.length) return;

    const batch = queue.slice(0, TELEMETRY_BATCH_SIZE);
    const endpoint = await getTelemetryEndpoint();
    const ok = await sendTelemetryBatch(batch, endpoint);
    if (ok) {
      const remaining = queue.slice(batch.length);
      await chrome.storage.local.set({ [TELEMETRY_QUEUE_KEY]: remaining });
    } else {
      console.warn("Telemetry flush failed; will retry on next trigger");
    }
  } catch (err) {
    console.warn("Telemetry flush failed", err);
  } finally {
    telemetryFlushInFlight = false;
  }
}

function ensureTelemetryAlarm() {
  try {
    chrome.alarms.create(TELEMETRY_FLUSH_ALARM, {
      delayInMinutes: 1,
      periodInMinutes: TELEMETRY_FLUSH_INTERVAL_MINUTES,
    });
  } catch (err) {
    console.warn("Failed to schedule telemetry flush alarm", err);
  }
}

chrome.alarms?.onAlarm.addListener((alarm) => {
  if (alarm && alarm.name === TELEMETRY_FLUSH_ALARM) {
    flushTelemetryQueue().catch(() => {});
  }
});

chrome.runtime?.onInstalled.addListener(() => {
  ensureTelemetryAlarm();
  flushTelemetryQueue().catch(() => {});
});

chrome.runtime?.onStartup?.addListener(() => {
  ensureTelemetryAlarm();
  flushTelemetryQueue().catch(() => {});
});

ensureTelemetryAlarm();
flushTelemetryQueue().catch(() => {});

function normalizeTargetCompanies(rawList = []) {
  if (!Array.isArray(rawList)) return [];
  const normalized = [];
  const seen = new Set();

  for (const item of rawList) {
    if (!item || typeof item !== "object") continue;
    const name = item.name ? String(item.name).trim() : "";
    const website = item.website ? String(item.website).trim() : "";
    const revenue = item.revenue ? String(item.revenue).trim() : "";
    const notes =
      item.notes?.toString().trim() ||
      item.rationale?.toString().trim() ||
      item.summary?.toString().trim() ||
      item.reason?.toString().trim() ||
      "";
    if (!name) continue;

    const canonicalName = name.toLowerCase();
    const normalizedWebsite = website
      ? website.replace(/^https?:\/\//i, "").replace(/\/+$/, "").toLowerCase()
      : "";
    const dedupeKey = `${canonicalName}|${normalizedWebsite}`;
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    normalized.push({ name, website, revenue, notes });
    if (normalized.length >= TARGET_COMPANY_GOAL) break;
  }

  return normalized;
}

async function generateTargets({ product, location, sectors, docs, docName, docText, docBase64 }) {
  const trimmedProduct = typeof product === "string" ? product.trim() : "";
  const trimmedLocation = typeof location === "string" ? location.trim() : "";
  const normalizedSectors = Array.isArray(sectors)
    ? sectors
        .map((sector) => (typeof sector === "string" ? sector.trim() : ""))
        .filter((sector) => !!sector)
    : [];

  if (!trimmedProduct) {
    return { error: "Product name is required to generate targets." };
  }

  const normalizedDocs = Array.isArray(docs) ? docs.filter(Boolean) : [];
  let docSection = "No supporting document provided.";

  if (normalizedDocs.length) {
    const docSnippets = normalizedDocs
      .map((doc) => {
        const name = doc?.name || "uploaded document";
        const rawContent = doc?.content_b64 || doc?.content || "";
        const content = decodeBase64Text(rawContent) || (typeof rawContent === "string" ? rawContent : "");
        if (!content) return "";
        const snippet = truncateText(content, 4000);
        return `--- ${name} ---\n${snippet}`;
      })
      .filter(Boolean);

    if (docSnippets.length) {
      docSection = `Supporting documents provided (first 4000 characters each):
${docSnippets.join("\n\n")}`;
    } else {
      docSection = "Supporting documents referenced but no readable content was found.";
    }
  } else {
    let documentText = "";
    if (typeof docText === "string" && docText.trim()) {
      documentText = docText.trim();
    } else if (docBase64) {
      documentText = decodeBase64Text(docBase64);
    }

    const truncatedDoc = truncateText(documentText, 4000);
    docSection = truncatedDoc
      ? `Supporting document (${docName || "uploaded document"}) excerpt (first 4000 characters):
${truncatedDoc}`
      : "No supporting document provided.";
  }

  const sectorSection = normalizedSectors.length
    ? `Target sectors of the product (use these as guardrails): ${normalizedSectors.join(", ")}`
    : "Target sectors of the product were not provided. Infer the most relevant industries based on the context and product value.";

  const mullInstruction = normalizedSectors.length
    ? `- Spend several internal reasoning steps comparing "${trimmedProduct}" against each target sector (${normalizedSectors.join(
        ", "
      )}) before listing companies. Keep this reasoning private but let it guide your short list so every recommendation clearly fits one of those sectors.`
    : `- Spend several internal reasoning steps inferring which sectors benefit most from "${trimmedProduct}" before listing companies. Keep this reasoning private but let it guide your short list.`;

  let prompt = "";
  try {
    const buildTargetsPrompt = requirePromptBuilder("buildTargetsPrompt");
    prompt = buildTargetsPrompt({
      product: trimmedProduct,
      location: trimmedLocation,
      sectorSection,
      docSection,
      mullInstruction,
      targetCompanyGoal: TARGET_COMPANY_GOAL,
    });
  } catch (err) {
    return { error: err?.message || String(err) };
  }

  const resp = await callLlmWithRetry(prompt, {
    generationConfig: {
      temperature: 0.4,
      maxOutputTokens: 100000,
    },
  });

  if (resp.error) {
    return { error: resp.error, details: resp.details };
  }

  const rawText = typeof resp.text === "string" ? resp.text : "";
  let companies = parseTargetCompaniesMarkdown(rawText);

  if (!companies.length) {
    const { parsed } = parseModelJsonResponse(resp);
    if (parsed && Array.isArray(parsed.companies)) {
      companies = parsed.companies;
    }
  }

  if (!companies.length) {
    return { error: "Model did not return a company list.", details: rawText || null };
  }

  companies = normalizeTargetCompanies(companies);
  return { ok: true, companies };
}

function prepareDatasetForPrompt(entries = []) {
  return entries.map((entry) => {
    const briefHtml = truncateText(entry?.result?.brief_html || "", 2000);
    return {
      id: entry.id,
      createdAt: entry.createdAt,
      request: entry.request || {},
      result: {
        brief_html: briefHtml,
        personas: Array.isArray(entry?.result?.personas) ? entry.result.personas : [],
        personaEmails: Array.isArray(entry?.result?.personaEmails) ? entry.result.personaEmails : [],
        email: entry?.result?.email || {},
      },
    };
  });
}

function composeExportPrompt(columns, entries, format) {
  const dataset = prepareDatasetForPrompt(entries);
  if (promptBuilders && typeof promptBuilders.buildExportPrompt === "function") {
    return promptBuilders.buildExportPrompt({ columns, dataset, format });
  }
  return "";
}

function ensureRowValues(row, columns) {
  const normalized = {};
  columns.forEach((col) => {
    const header = col.header;
    const value = row && Object.prototype.hasOwnProperty.call(row, header) ? row[header] : "";
    if (value === undefined || value === null) {
      normalized[header] = "";
    } else if (typeof value === "string") {
      normalized[header] = value;
    } else if (typeof value === "number" || typeof value === "boolean") {
      normalized[header] = String(value);
    } else {
      try {
        normalized[header] = JSON.stringify(value);
      } catch (err) {
        normalized[header] = String(value);
      }
    }
  });
  return normalized;
}

function filterHistoryEntries(entries, selection = {}) {
  const { type = "all" } = selection;
  if (type === "all") {
    return [...entries];
  }

  if (type === "custom") {
    const ids = Array.isArray(selection.selectedIds) ? new Set(selection.selectedIds) : new Set();
    if (!ids.size) return [];
    return entries.filter((entry) => ids.has(entry.id));
  }

  if (type === "date") {
    const path = selection.dateFieldPath || "createdAt";
    const fromTime = selection.from ? new Date(selection.from).getTime() : null;
    const toTime = selection.to ? new Date(selection.to).getTime() : null;

    return entries.filter((entry) => {
      const value = getValueAtPath(entry, path);
      if (!value) return false;
      const ts = new Date(value).getTime();
      if (Number.isNaN(ts)) return false;
      if (fromTime !== null && ts < fromTime) return false;
      if (toTime !== null && ts > toTime) return false;
      return true;
    });
  }

  return [...entries];
}

function normalizeLocationString(candidate) {
  if (!candidate) return "";
  if (typeof candidate === "string") {
    return candidate.trim();
  }
  if (Array.isArray(candidate)) {
    const parts = candidate.map((item) => normalizeLocationString(item)).filter(Boolean);
    return parts.join(", ");
  }
  if (typeof candidate === "object") {
    const prioritizedKeys = [
      "formatted_address",
      "address",
      "address_text",
      "address_line",
      "description",
      "text",
      "value",
      "name",
      "full_address",
    ];
    for (const key of prioritizedKeys) {
      if (typeof candidate[key] === "string" && candidate[key].trim()) {
        return candidate[key].trim();
      }
    }

    const cityParts = [];
    if (candidate.city) cityParts.push(candidate.city);
    if (candidate.state || candidate.region || candidate.province) {
      cityParts.push(candidate.state || candidate.region || candidate.province);
    }
    if (candidate.country) cityParts.push(candidate.country);
    if (cityParts.length) {
      return cityParts
        .map((part) => (typeof part === "string" ? part.trim() : ""))
        .filter(Boolean)
        .join(", ");
    }
  }
  return "";
}

function resolveLocationFromStructured(structured) {
  if (!structured || typeof structured !== "object") return "";
  const candidates = [
    structured.hq_location,
    structured.headquarters,
    structured.location,
    structured.answer,
    structured.result,
  ];

  if (Array.isArray(structured.supporting_places)) {
    structured.supporting_places.forEach((place) => {
      candidates.push(place);
      if (place && typeof place === "object") {
        candidates.push(place.formatted_address);
      }
    });
  }

  for (const candidate of candidates) {
    const normalized = normalizeLocationString(candidate);
    if (normalized) {
      return normalized;
    }
  }
  return "";
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function columnNumberToName(n) {
  let name = "";
  let num = n;
  while (num > 0) {
    const remainder = (num - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    num = Math.floor((num - 1) / 26);
  }
  return name;
}

function buildWorksheetXml(headers, rows) {
  const totalRows = rows.length + 1;
  const lastColumn = columnNumberToName(headers.length);
  const dimensionRef = `A1:${lastColumn}${Math.max(totalRows, 1)}`;

  const makeCell = (rowIndex, colIndex, value, isHeader = false) => {
    const ref = `${columnNumberToName(colIndex)}${rowIndex}`;
    const escaped = escapeXml(value);
    const style = isHeader ? ` s="1"` : "";
    return `<c r="${ref}" t="inlineStr"${style}><is><t>${escaped}</t></is></c>`;
  };

  const rowsXml = [];
  const headerCells = headers
    .map((header, idx) => makeCell(1, idx + 1, header, true))
    .join("");
  rowsXml.push(`<row r="1">${headerCells}</row>`);

  rows.forEach((row, rowIdx) => {
    const cells = headers
      .map((header, colIdx) => makeCell(rowIdx + 2, colIdx + 1, row[header] || "", false))
      .join("");
    rowsXml.push(`<row r="${rowIdx + 2}">${cells}</row>`);
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${dimensionRef}"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    ${rowsXml.join("\n    ")}
  </sheetData>
</worksheet>`;
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
}

function buildContentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function buildRootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function buildWorkbookRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function buildWorkbookXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="Calc"/>
  <sheets>
    <sheet name="Export" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
}

function buildDocPropsCoreXml(timestamp) {
  const iso = timestamp.toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>AccountIQ Export</dc:creator>
  <cp:lastModifiedBy>AccountIQ Export</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified>
</cp:coreProperties>`;
}

function buildDocPropsAppXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>AccountIQ Export</Application>
</Properties>`;
}

function stringToUint8(str) {
  return new TextEncoder().encode(str);
}

function crc32(buf) {
  const table = crc32.table || (crc32.table = (() => {
    let c;
    const table = [];
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[n] = c >>> 0;
    }
    return table;
  })());

  let crc = 0 ^ (-1);
  for (let i = 0; i < buf.length; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xff];
  }
  return (crc ^ (-1)) >>> 0;
}

function dateToDosParts(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = Math.floor(date.getSeconds() / 2);

  const dosDate = ((year - 1980) << 9) | (month << 5) | day;
  const dosTime = (hours << 11) | (minutes << 5) | seconds;
  return {
    dosDate: year < 1980 ? 0 : dosDate,
    dosTime: year < 1980 ? 0 : dosTime,
  };
}

function assembleZip(files) {
  let totalSize = 0;
  const fileEntries = [];
  const centralEntries = [];
  let offset = 0;

  const encoder = new TextEncoder();
  const now = new Date();
  const { dosDate, dosTime } = dateToDosParts(now);

  files.forEach((file) => {
    const nameBytes = encoder.encode(file.path);
    let dataBytes = file.data instanceof Uint8Array ? file.data : encoder.encode(file.data);
    const crc = crc32(dataBytes);
    const compressedSize = dataBytes.length;
    const uncompressedSize = dataBytes.length;
    const localHeaderSize = 30 + nameBytes.length;

    const localHeader = new Uint8Array(localHeaderSize);
    const lhView = new DataView(localHeader.buffer);
    lhView.setUint32(0, 0x04034b50, true);
    lhView.setUint16(4, 20, true);
    lhView.setUint16(6, 0, true);
    lhView.setUint16(8, 0, true);
    lhView.setUint16(10, dosTime, true);
    lhView.setUint16(12, dosDate, true);
    lhView.setUint32(14, crc, true);
    lhView.setUint32(18, compressedSize, true);
    lhView.setUint32(22, uncompressedSize, true);
    lhView.setUint16(26, nameBytes.length, true);
    lhView.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);

    fileEntries.push(localHeader);
    fileEntries.push(dataBytes);

    const centralHeaderSize = 46 + nameBytes.length;
    const centralHeader = new Uint8Array(centralHeaderSize);
    const chView = new DataView(centralHeader.buffer);
    chView.setUint32(0, 0x02014b50, true);
    chView.setUint16(4, 20, true);
    chView.setUint16(6, 20, true);
    chView.setUint16(8, 0, true);
    chView.setUint16(10, 0, true);
    chView.setUint16(12, dosTime, true);
    chView.setUint16(14, dosDate, true);
    chView.setUint32(16, crc, true);
    chView.setUint32(20, compressedSize, true);
    chView.setUint32(24, uncompressedSize, true);
    chView.setUint16(28, nameBytes.length, true);
    chView.setUint16(30, 0, true);
    chView.setUint16(32, 0, true);
    chView.setUint16(34, 0, true);
    chView.setUint16(36, 0, true);
    chView.setUint32(38, 0, true);
    chView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);
    centralEntries.push(centralHeader);

    offset += localHeader.length + dataBytes.length;
  });

  const centralDirectorySize = centralEntries.reduce((acc, arr) => acc + arr.length, 0);
  const centralDirectoryOffset = offset;

  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(4, 0, true);
  eocdView.setUint16(6, 0, true);
  eocdView.setUint16(8, files.length, true);
  eocdView.setUint16(10, files.length, true);
  eocdView.setUint32(12, centralDirectorySize, true);
  eocdView.setUint32(16, centralDirectoryOffset, true);
  eocdView.setUint16(20, 0, true);

  totalSize =
    fileEntries.reduce((acc, arr) => acc + arr.length, 0) +
    centralEntries.reduce((acc, arr) => acc + arr.length, 0) +
    eocd.length;

  const output = new Uint8Array(totalSize);
  let pointer = 0;
  fileEntries.forEach((arr) => {
    output.set(arr, pointer);
    pointer += arr.length;
  });
  centralEntries.forEach((arr) => {
    output.set(arr, pointer);
    pointer += arr.length;
  });
  output.set(eocd, pointer);

  return output;
}

function buildXlsxFile(headers, rows) {
  const timestamp = new Date();
  const worksheetXml = buildWorksheetXml(headers, rows);
  const files = [
    { path: "[Content_Types].xml", data: buildContentTypesXml() },
    { path: "_rels/.rels", data: buildRootRelsXml() },
    { path: "xl/workbook.xml", data: buildWorkbookXml() },
    { path: "xl/_rels/workbook.xml.rels", data: buildWorkbookRelsXml() },
    { path: "xl/worksheets/sheet1.xml", data: worksheetXml },
    { path: "xl/styles.xml", data: buildStylesXml() },
    { path: "docProps/core.xml", data: buildDocPropsCoreXml(timestamp) },
    { path: "docProps/app.xml", data: buildDocPropsAppXml() },
  ];
  return assembleZip(files);
}

function uint8ToBase64(bytes) {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    const slice = bytes.subarray(i, i + chunk);
    binary += String.fromCharCode.apply(null, slice);
  }
  return btoa(binary);
}

function stringToBase64(str) {
  return uint8ToBase64(stringToUint8(str));
}

function generateMarkdownFromRows(headers, rows) {
  const headerLine = `| ${headers.join(" | ")} |`;
  const separator = `| ${headers.map(() => "---").join(" | ")} |`;
  const body = rows
    .map((row) => {
      const cells = headers.map((header) => {
        const value = row[header] || "";
        return value.replace(/\n/g, "<br>");
      });
      return `| ${cells.join(" | ")} |`;
    })
    .join("\n");
  return `${headerLine}\n${separator}${body ? `\n${body}` : ""}`;
}

async function persistHistoryEntry(storageKey, entry, limit = 25) {
  const existing = await chrome.storage.local.get([storageKey]);
  const history = Array.isArray(existing[storageKey]) ? existing[storageKey] : [];

  history.unshift(entry);
  const trimmed = history.slice(0, limit);

  await chrome.storage.local.set({ [storageKey]: trimmed });
  return entry;
}

async function saveResearchHistoryEntry(request, result) {
  try {
    const entry = {
      id: Date.now().toString(),
      createdAt: new Date().toISOString(),
      request: {
        company: request.company || "",
        location: request.location || "",
        product: request.product || "",
      },
      result: {
        brief_html: result.brief_html || "",
        modules: result.modules || {},
        company_name: result.company_name || request.company || "",
        revenue_estimate: result.revenue_estimate || "",
        top_5_news: Array.isArray(result.top_5_news) ? result.top_5_news : [],
        hq_location: result.hq_location || "",
        hq_lookup_error: result.hq_lookup_error || "",
        personas: Array.isArray(result.personas) ? result.personas : [],
        personaEmails: Array.isArray(result.personaEmails) ? result.personaEmails : [],
        personaEmailVersions: Array.isArray(result.personaEmailVersions)
          ? result.personaEmailVersions
          : (Array.isArray(result.personaEmails)
              ? result.personaEmails.map((draft) => ({ versions: [draft], activeIndex: 0 }))
              : []),
        industry_sector: typeof result.industry_sector === "string" ? result.industry_sector : "",
        telephonicPitches: Array.isArray(result.telephonicPitches) ? result.telephonicPitches : [],
        telephonicPitchVersions: Array.isArray(result.telephonicPitchVersions)
          ? result.telephonicPitchVersions
          : (Array.isArray(result.telephonicPitches)
              ? result.telephonicPitches.map((draft) => ({ versions: [draft], activeIndex: 0 }))
              : []),
        telephonicPitchError: typeof result.telephonicPitchError === "string" ? result.telephonicPitchError : "",
        telephonicPitchAttempts: Array.isArray(result.telephonicPitchAttempts) ? result.telephonicPitchAttempts : [],
        personaEmailVersionIndexes: Array.isArray(result.personaEmailVersionIndexes)
          ? result.personaEmailVersionIndexes
          : (Array.isArray(result.personaEmailVersions)
              ? result.personaEmailVersions.map((v) => v.activeIndex || 0)
              : []),
        telephonicPitchVersionIndexes: Array.isArray(result.telephonicPitchVersionIndexes)
          ? result.telephonicPitchVersionIndexes
          : (Array.isArray(result.telephonicPitchVersions)
              ? result.telephonicPitchVersions.map((v) => v.activeIndex || 0)
              : []),
        manualEdits: result.manualEdits || {},
        email: result.email || {},
        overview_error: result.overview_error || "",
        persona_error: result.persona_error || "",
      },
    };

    return await persistHistoryEntry(RESEARCH_HISTORY_KEY, entry);
  } catch (err) {
    console.warn("Failed to persist research history", err);
    return null;
  }
}

async function updateResearchHistoryEntry(id, resultUpdates = {}) {
  try {
    const data = await chrome.storage.local.get([RESEARCH_HISTORY_KEY]);
    const history = Array.isArray(data[RESEARCH_HISTORY_KEY]) ? data[RESEARCH_HISTORY_KEY] : [];
    const idx = history.findIndex((entry) => entry && entry.id === id);
    if (idx === -1) return { ok: false, error: "Entry not found" };
    const existing = history[idx];
    const mergedResult = { ...(existing.result || {}), ...(resultUpdates || {}) };
    history[idx] = { ...existing, result: mergedResult };
    await chrome.storage.local.set({ [RESEARCH_HISTORY_KEY]: history });
    return { ok: true, entry: history[idx] };
  } catch (err) {
    console.warn("Failed to update research history", err);
    return { ok: false, error: err?.message || String(err) };
  }
}

async function renameHistoryEntry(storageKey, id, title) {
  try {
    const data = await chrome.storage.local.get([storageKey]);
    const history = Array.isArray(data[storageKey]) ? data[storageKey] : [];
    const idx = history.findIndex((entry) => entry && entry.id === id);
    if (idx === -1) return { ok: false, error: "Entry not found" };
    const normalizedTitle = typeof title === "string" ? title.trim() : "";
    const updated = { ...(history[idx] || {}) };
    if (normalizedTitle) {
      updated.customTitle = normalizedTitle;
    } else {
      delete updated.customTitle;
    }
    history[idx] = updated;
    await chrome.storage.local.set({ [storageKey]: history });
    return { ok: true, entry: updated };
  } catch (err) {
    console.warn("Failed to rename history entry", err);
    return { ok: false, error: err?.message || String(err) };
  }
}

async function deleteHistoryEntry(storageKey, id) {
  try {
    const data = await chrome.storage.local.get([storageKey]);
    const history = Array.isArray(data[storageKey]) ? data[storageKey] : [];
    const nextHistory = history.filter((entry) => entry && entry.id !== id);
    if (nextHistory.length === history.length) {
      return { ok: false, error: "Entry not found" };
    }
    await chrome.storage.local.set({ [storageKey]: nextHistory });
    return { ok: true, deletedId: id };
  } catch (err) {
    console.warn("Failed to delete history entry", err);
    return { ok: false, error: err?.message || String(err) };
  }
}

async function saveTargetHistoryEntry(request, result) {
  try {
    const docRefs = Array.isArray(request.docs)
      ? request.docs
          .map((doc) => {
            const id = doc && doc.id ? String(doc.id) : "";
            const name = doc && doc.name ? String(doc.name) : "";
            if (!id && !name) return null;
            return { id, name };
          })
          .filter(Boolean)
      : [];

    const entry = {
      id: Date.now().toString(),
      createdAt: new Date().toISOString(),
      request: {
        product: request.product || "",
        location: request.location || "",
        docName: request.docName || "",
        docs: docRefs,
        sectors: Array.isArray(request.sectors)
          ? request.sectors
              .map((sector) => (typeof sector === "string" ? sector.trim() : ""))
              .filter((sector) => !!sector)
          : [],
      },
      result: {
        companies: Array.isArray(result?.companies) ? result.companies : [],
      },
    };

    return await persistHistoryEntry(TARGET_HISTORY_KEY, entry);
  } catch (err) {
    console.warn("Failed to persist target history", err);
    return null;
  }
}

function summarizePersonasForTelepitch(personas = []) {
  if (!Array.isArray(personas) || !personas.length) return "";
  return personas
    .map((persona, idx) => {
      const role = persona.designation || persona.department || `Persona`;
      const bits = [role, persona.department && role !== persona.department ? `Department: ${persona.department}` : null]
        .filter(Boolean)
        .join(" - ");
      return bits || `Persona`;
    })
    .join("\n");
}

function buildTelephonicPitchPrompt({ persona = {}, company, location, product, docsText, pitchFromCompany }) {
  const builder = requirePromptBuilder("buildTelephonicPitchPrompt");
  return builder({ persona, company, location, product, docsText, pitchFromCompany });
}

function buildPersonaEmailsPrompt({ persona = {}, company, location, product, docsText, pitchFromCompany }) {
  const builder = requirePromptBuilder("buildPersonaEmailsPrompt");
  return builder({ persona, company, location, product, docsText, pitchFromCompany });
}

function buildEmailRevisionPrompt({ persona = {}, company, location, product, baseEmail = {}, instructions = "", pitchingOrg }) {
  const builder = requirePromptBuilder("buildEmailRevisionPrompt");
  return builder({
    persona,
    company,
    location,
    product,
    baseEmail,
    instructions,
    pitchingOrg,
  });
}

function buildPitchRevisionPrompt({ persona = {}, company, location, product, basePitch = {}, instructions = "", pitchingOrg }) {
  const builder = requirePromptBuilder("buildPitchRevisionPrompt");
  return builder({
    persona,
    company,
    location,
    product,
    basePitch,
    instructions,
    pitchingOrg,
  });
}

function deriveTelephonicPitchArray(payload, depth = 0) {
  if (!payload || depth > 4) return [];
  if (Array.isArray(payload)) return payload;
  if (typeof payload !== "object") return [];

  const candidateKeys = [
    "telephonic_pitches",
    "telephonicPitches",
    "telephonicPitch",
    "pitches",
    "scripts",
    "phone_scripts",
    "phoneScripts",
  ];

  for (const key of candidateKeys) {
    const candidate = payload[key];
    if (Array.isArray(candidate) && candidate.length) return candidate;
    if (candidate && typeof candidate === "object") {
      return [candidate];
    }
    if (typeof candidate === "string" && candidate.trim()) {
      return [{ full_pitch: candidate.trim() }];
    }
  }

  // Fallback: if object looks like a single pitch
  if (payload.full_pitch || payload.fullPitch || payload.script || payload.call_script || payload.pitch) {
    return [payload];
  }

  if (payload.result && typeof payload.result === "object") {
    const nested = deriveTelephonicPitchArray(payload.result, depth + 1);
    if (nested.length) return nested;
  }

  if (payload.data && typeof payload.data === "object") {
    const nested = deriveTelephonicPitchArray(payload.data, depth + 1);
    if (nested.length) return nested;
  }

  for (const value of Object.values(payload)) {
    if (typeof value === "object" && value !== null) {
      const nested = deriveTelephonicPitchArray(value, depth + 1);
      if (nested.length) return nested;
    }
  }

  return [];
}

function normalizeTelephonicPitchEntry(pitch, idx, personas) {
  const fallbackPersona = (Array.isArray(personas) && personas[idx]) || {};
  const base = {
    personaName: fallbackPersona.name || `Persona ${idx + 1}`,
    personaDesignation: fallbackPersona.designation || "",
    personaDepartment: fallbackPersona.department || "",
    callGoal: "",
    opener: "",
    discoveryQuestion: "",
    valueStatement: "",
    proofPoint: "",
    cta: "",
    script: "",
  };

  if (!pitch) return base;
  if (typeof pitch === "string") {
    base.script = pitch;
    return base;
  }
  if (typeof pitch !== "object") {
    return base;
  }

  base.personaName = pitch.persona_name || pitch.name || base.personaName;
  base.personaDesignation = pitch.designation || pitch.persona_designation || base.personaDesignation;
  base.personaDepartment = pitch.department || pitch.persona_department || base.personaDepartment;
  base.callGoal = pitch.call_goal || pitch.callGoal || pitch.call_objective || pitch.goal || base.callGoal;
  base.opener = pitch.opener || pitch.opening || pitch.opening_hook || pitch.hook || base.opener;

  if (Array.isArray(pitch.discovery_questions) && pitch.discovery_questions.length) {
    base.discoveryQuestion = pitch.discovery_questions.join(" / ");
  } else {
    base.discoveryQuestion =
      pitch.discovery_question || pitch.discoveryQuestion || pitch.question || base.discoveryQuestion;
  }

  base.valueStatement =
    pitch.value_statement || pitch.valueStatement || pitch.value_pitch || pitch.value || base.valueStatement;
  base.proofPoint =
    pitch.proof_point ||
    pitch.proofPoint ||
    pitch.credibility_statement ||
    pitch.social_proof ||
    base.proofPoint;
  base.cta = pitch.cta || pitch.closing_prompt || pitch.next_step || pitch.closing || base.cta;
  base.script = pitch.full_pitch || pitch.fullPitch || pitch.pitch || pitch.call_script || pitch.script || base.script;

  if (!base.script) {
    const fallbackLines = [
      base.opener,
      base.discoveryQuestion ? `Discovery: ${base.discoveryQuestion}` : "",
      base.valueStatement,
      base.proofPoint,
      base.cta,
    ]
      .filter(Boolean)
      .join("\n");
    base.script = fallbackLines;
  }

  return base;
}

function summarizeTelephonicDebugInfo(payload) {
  if (!payload) return "";
  let str = "";
  if (typeof payload === "string") {
    str = payload;
  } else {
    try {
      str = JSON.stringify(payload);
    } catch (err) {
      str = "";
    }
  }
  if (!str) return "";
  const trimmed = str.trim();
  if (!trimmed) return "";
  return trimmed.length > 800 ? `${trimmed.slice(0, 800)}...` : trimmed;
}

function extractTelephonicPitchResponse(resp, personas) {
  let teleText = typeof resp.text === "string" ? resp.text : "";
  let teleParsed = extractJsonFromText(teleText);
  if (!teleParsed && teleText) {
    try {
      teleParsed = JSON.parse(teleText);
    } catch (err) {
      teleParsed = null;
    }
  }

  if (!teleParsed && resp.raw?.candidates?.[0]?.content?.parts?.length) {
    const combined = resp.raw.candidates[0].content.parts
      .map((part) => {
        if (typeof part?.text === "string" && part.text.trim()) return part.text;
        if (part?.functionCall?.args) {
          try {
            return JSON.stringify(part.functionCall.args);
          } catch (err) {
            return null;
          }
        }
        if (part?.functionResponse?.response) {
          try {
            return JSON.stringify(part.functionResponse.response);
          } catch (err) {
            return null;
          }
        }
        if (part?.inlineData?.mimeType === "application/json" && part?.inlineData?.data) {
          const decoded = decodeBase64Text(part.inlineData.data);
          return decoded || null;
        }
        return null;
      })
      .filter(Boolean)
      .join("\n");
    if (combined) {
      teleParsed = extractJsonFromText(combined);
      if (!teleParsed) {
        try {
          teleParsed = JSON.parse(combined);
        } catch (err) {
          teleParsed = null;
        }
      }
      if (!teleText) {
        teleText = combined;
      }
    }
  }

  const teleArray = deriveTelephonicPitchArray(teleParsed);
  if (!teleArray.length) {
    return {
      pitches: [],
      rawText: teleText || null,
      error: "Response did not include a telephonic_pitches array.",
    };
  }

  const normalized = teleArray
    .filter((entry) => entry && (typeof entry === "object" || typeof entry === "string"))
    .map((entry, idx) => normalizeTelephonicPitchEntry(entry, idx, personas))
    .filter(Boolean);

  if (!normalized.length) {
    return {
      pitches: [],
      rawText: teleText || null,
      error: "Unable to normalize telephonic pitch entries.",
    };
  }

  return { pitches: normalized };
}

async function generateTelephonicPitchScripts({ personas, company, location, product, docsText, pitchFromCompany }) {
  const pitchingCompany = (pitchFromCompany && pitchFromCompany.trim()) || (await loadPitchingCompany());
  const attempts = [];
  const pitches = [];

  const tasks = (personas || []).map(async (persona, idx) => {
    let prompt = "";
    try {
      prompt = buildTelephonicPitchPrompt({
        persona,
        company,
        location,
        product,
        docsText,
        pitchFromCompany: pitchingCompany,
      });
    } catch (err) {
      attempts.push({
        label: `persona_${idx + 1}`,
        error: err?.message || String(err),
      });
      pitches[idx] = normalizeTelephonicPitchEntry(null, idx, personas);
      return;
    }

    const generationConfig = {
      temperature: 0.2,
      maxOutputTokens: 4096,
    };

    try {
      const teleResp = await callLlmWithRetry(prompt, {
        model: LLAMA_33_MODEL,
        secondaryModel: LLAMA_33_MODEL,
        generationConfig,
      });

      if (teleResp.error) {
        attempts.push({
          label: `persona_${idx + 1}`,
          error: teleResp.error,
          details: summarizeTelephonicDebugInfo(teleResp.details || teleResp.text || teleResp.raw),
        });
        pitches[idx] = normalizeTelephonicPitchEntry(null, idx, personas);
        return;
      }

      const rawText = typeof teleResp.text === "string" ? teleResp.text : "";
      const parsed = parseTelephonicPitchMarkdown(rawText, persona);
      const script = parsed.script || rawText;
      const normalized = normalizeTelephonicPitchEntry(
        {
          persona_name: parsed.personaName,
          persona_designation: parsed.personaDesignation,
          persona_department: parsed.personaDepartment,
          full_pitch: script,
        },
        idx,
        personas
      );
      pitches[idx] = normalized;
    } catch (err) {
      attempts.push({
        label: `persona_${idx + 1}`,
        error: err?.message || String(err),
      });
      pitches[idx] = normalizeTelephonicPitchEntry(null, idx, personas);
    }
  });

  await Promise.all(tasks);

  const hasContent = pitches.some((p) => p && p.script);
  const error =
    hasContent || !attempts.length
      ? ""
      : attempts[attempts.length - 1]?.error || "Unable to generate telephonic pitches.";

  return { pitches, error, attempts };
}

async function generatePersonaEmails({ personas, company, location, product, docsText, pitchFromCompany }) {
  const pitchingCompany = (pitchFromCompany && pitchFromCompany.trim()) || (await loadPitchingCompany());

  const personaEmails = [];
  const attempts = [];

  const tasks = (personas || []).map(async (persona, idx) => {
    let prompt = "";
    try {
      prompt = buildPersonaEmailsPrompt({
        persona,
        company,
        location,
        product,
        docsText,
        pitchFromCompany: pitchingCompany,
      });
    } catch (err) {
      attempts.push({ label: `persona_${idx + 1}`, error: err?.message || String(err) });
      personaEmails[idx] = normalizePersonaEmails(personas, [])[idx] || {
        personaName: persona?.name || `Persona ${idx + 1}`,
        personaDesignation: persona?.designation || "",
        personaDepartment: persona?.department || "",
        subject: "",
        body: "",
      };
      return;
    }

    const resp = await callLlmWithRetry(prompt, {
      model: LLAMA_33_MODEL,
      secondaryModel: LLAMA_33_MODEL,
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 12000,
      },
    });

    if (resp.error) {
      attempts.push({ label: `persona_${idx + 1}`, error: resp.error, details: resp.details || resp.text || "" });
      personaEmails[idx] = normalizePersonaEmails(personas, [])[idx] || {
        personaName: persona?.name || `Persona ${idx + 1}`,
        personaDesignation: persona?.designation || "",
        personaDepartment: persona?.department || "",
        subject: "",
        body: "",
      };
      return;
    }

    const rawText = typeof resp.text === "string" ? resp.text : "";
    const parsed = parsePersonaEmailMarkdown(rawText, persona);
    personaEmails[idx] = {
      personaName: parsed.personaName || persona?.name || `Persona ${idx + 1}`,
      personaDesignation: parsed.personaDesignation || persona?.designation || "",
      personaDepartment: parsed.personaDepartment || persona?.department || "",
      subject: parsed.subject || "",
      body: parsed.body || "",
    };
  });

  await Promise.all(tasks);

  const error =
    personaEmails.some((e) => e && (e.subject || e.body)) || !attempts.length
      ? ""
      : attempts[attempts.length - 1]?.error || "Model did not return valid emails.";

  return { personaEmails, rawText: "", attempts, error };
}

function normalizePersonas(rawPersonas, companyName) {
  if (!Array.isArray(rawPersonas)) return [];
  return rawPersonas.map((p = {}) => {
    let linkedinKeywords = normalizeLinkedInKeywords(
      p.linkedin_keywords || p.linkedinKeywords || "",
      companyName || "",
      p.designation || p.department || p.product || ""
    );
    if (!linkedinKeywords) {
      linkedinKeywords = buildLinkedInKeywordFallback(p, companyName || "", p.product || "");
    }
    return {
      name: p.name || "",
      designation: p.designation || "",
      department: p.department || "",
      linkedin_keywords: linkedinKeywords,
      linkedin_search_url: toLinkedInPeopleSearchUrl(linkedinKeywords),
      zoominfo_link:
        p.zoominfo_link || p.zoomInfo || p.zoominfo || p.zoom || buildZoomInfoSearchLink(p, companyName || ""),
    };
  });
}

function normalizePersonaEmails(rawPersonas = [], personaEmailsArray = []) {
  const personaEmails = rawPersonas.map((p, idx) => {
    const personaName = p?.name || `Persona ${idx + 1}`;
    const emailData = p?.email || p?.persona_email || {};
    const subject =
      (emailData && typeof emailData === "object" ? emailData.subject : undefined) || p?.email_subject || "";
    const body = (emailData && typeof emailData === "object" ? emailData.body : undefined) || p?.email_body || "";
    return {
      personaName,
      personaDesignation: p?.designation || "",
      personaDepartment: p?.department || "",
      subject: subject || "",
      body: body || "",
    };
  });

  const emailsArray = Array.isArray(personaEmailsArray) ? personaEmailsArray : [];
  if (emailsArray.length) {
    personaEmails.forEach((entry, idx) => {
      const fallbackEmail =
        emailsArray[idx] ||
        emailsArray.find((pe) => {
          const candidateName = (pe.persona_name || pe.name || "").toLowerCase();
          return candidateName && candidateName === (entry.personaName || "").toLowerCase();
        });
      if (fallbackEmail) {
        const sub = fallbackEmail.subject || fallbackEmail.email_subject || "";
        const bod = fallbackEmail.body || fallbackEmail.email_body || "";
        if (!entry.subject) entry.subject = sub || "";
        if (!entry.body) entry.body = bod || "";
      }
    });
    if (!personaEmails.length) {
      emailsArray.forEach((pe, idx) => {
        personaEmails.push({
          personaName: pe.persona_name || pe.name || `Persona ${idx + 1}`,
          personaDesignation: pe.persona_designation || pe.designation || "",
          personaDepartment: pe.persona_department || pe.department || "",
          subject: pe.subject || pe.email_subject || "",
          body: pe.body || pe.email_body || "",
        });
      });
    }
  }

  return personaEmails;
}

function normalizeTelephonicPitches(parsed, personas) {
  const telephonicRaw = deriveTelephonicPitchArray(parsed);
  if (telephonicRaw.length) {
    const telephonicPitches = telephonicRaw
      .filter((entry) => entry && (typeof entry === "object" || typeof entry === "string"))
      .map((entry, idx) => normalizeTelephonicPitchEntry(entry, idx, personas))
      .filter(Boolean);
    if (!telephonicPitches.length) {
      return {
        telephonicPitches: [],
        telephonicPitchError: "Telephonic pitches present but could not be normalized.",
        telephonicPitchAttempts: [],
      };
    }
    return { telephonicPitches, telephonicPitchError: "", telephonicPitchAttempts: [] };
  }

  return {
    telephonicPitches: [],
    telephonicPitchError: "Model response missing telephonic_pitches array.",
    telephonicPitchAttempts: [],
  };
}

async function revisePersonaEmail({ persona, email, company, product, location, instructions, pitchingOrg }) {
  try {
    const prompt = buildEmailRevisionPrompt({
      persona: persona || {},
      company,
      location,
      product,
      baseEmail: email || {},
      instructions: instructions || "",
      pitchingOrg,
    });

    const resp = await callLlmWithRetry(prompt, {
      model: LLAMA_33_MODEL,
      secondaryModel: LLAMA_33_MODEL,
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 12000,
      },
    });

    if (resp.error) {
      return { error: resp.error, details: resp.details };
    }

    const rawText = typeof resp.text === "string" ? resp.text : "";
    const parsed = parsePersonaEmailMarkdown(rawText, persona);

    const draft = {
      personaName: persona?.name || persona?.personaName || persona?.persona_name || "",
      personaDesignation: persona?.designation || persona?.personaDesignation || persona?.persona_designation || "",
      personaDepartment: persona?.department || persona?.personaDepartment || persona?.persona_department || "",
      subject: parsed.subject || "",
      body: parsed.body || "",
    };

    return { draft, rawText };
  } catch (err) {
    return { error: err?.message || String(err) };
  }
}

async function reviseAllPersonaEmails({ personas = [], emails = [], company, product, location, instructions, pitchingOrg }) {
  const tasks = personas.map((persona, idx) =>
    revisePersonaEmail({
      persona,
      email: emails[idx] || {},
      company,
      product,
      location,
      instructions,
      pitchingOrg,
    })
      .then((res) => ({ ...res, personaIndex: idx }))
      .catch((err) => ({ error: err?.message || String(err), personaIndex: idx }))
  );
  const results = await Promise.all(tasks);
  return results;
}

async function revisePersonaPitch({ persona, pitch, company, product, location, instructions, pitchingOrg }) {
  try {
    const prompt = buildPitchRevisionPrompt({
      persona: persona || {},
      company,
      location,
      product,
      basePitch: pitch || {},
      instructions: instructions || "",
      pitchingOrg,
    });

    const resp = await callLlmWithRetry(prompt, {
      model: LLAMA_33_MODEL,
      secondaryModel: LLAMA_33_MODEL,
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 12000,
      },
    });

    if (resp.error) {
      return { error: resp.error, details: resp.details };
    }

    const rawText = typeof resp.text === "string" ? resp.text : "";
    const parsed = parseTelephonicPitchMarkdown(rawText, persona);
    const normalized = normalizeTelephonicPitchEntry(
      {
        persona_name: parsed.personaName,
        persona_designation: parsed.personaDesignation,
        persona_department: parsed.personaDepartment,
        full_pitch: parsed.script,
      },
      0,
      [persona || {}]
    );
    return {
      draft: normalized,
      rawText,
    };
  } catch (err) {
    return { error: err?.message || String(err) };
  }
}

async function reviseAllPersonaPitches({ personas = [], pitches = [], company, product, location, instructions, pitchingOrg }) {
  const tasks = personas.map((persona, idx) =>
    revisePersonaPitch({
      persona,
      pitch: pitches[idx] || {},
      company,
      product,
      location,
      instructions,
      pitchingOrg,
    })
      .then((res) => ({ ...res, personaIndex: idx }))
      .catch((err) => ({ error: err?.message || String(err), personaIndex: idx }))
  );
  const results = await Promise.all(tasks);
  return results;
}

function derivePrimaryEmail(parsed, personaEmails) {
  const emailObj = { subject: "", body: "" };
  if (Array.isArray(personaEmails) && personaEmails.length) {
    emailObj.subject = personaEmails[0]?.subject || "";
    emailObj.body = personaEmails[0]?.body || "";
  }

  if (!emailObj.subject && !emailObj.body) {
    if (typeof parsed?.personalized_email === "string") {
      emailObj.body = parsed.personalized_email;
    } else if (parsed?.personalized_email && typeof parsed.personalized_email === "object") {
      emailObj.subject = parsed.personalized_email.subject || "";
      emailObj.body = parsed.personalized_email.body || "";
    }
  }

  return emailObj;
}

function buildBriefHtmlFromOverview({ companyName, hqDisplay, revenue, industry, topNews = [] }) {
  let briefHtml = `<h4>${companyName || ""}</h4>`;
  briefHtml += `<p><strong>Headquarters:</strong> ${hqDisplay} &nbsp; <strong>Revenue:</strong> ${
    revenue || ""
  } &nbsp; <strong>Industry sector:</strong> ${industry || ""}</p>`;
  if (topNews && topNews.length) {
    briefHtml += `<h5>Top News</h5><ul>`;
    topNews.slice(0, 5).forEach((n) => {
      briefHtml += `<li><strong>${n.title || ""}</strong><div>${n.summary || ""}</div></li>`;
    });
    briefHtml += `</ul>`;
  } else {
    briefHtml += `<h5>Top News</h5><p>No recent headlines found.</p>`;
  }
  return briefHtml;
}

function normalizeTopNewsEntries(topNews = []) {
  if (!Array.isArray(topNews)) return [];
  return topNews
    .map((item) => {
      const title = item?.title || item?.headline || item?.name || "";
      const summary = item?.summary || item?.description || item?.details || "";
      return {
        title: title ? String(title) : "",
        summary: summary ? String(summary) : "",
      };
    })
    .filter((entry) => entry.title || entry.summary);
}

function parseHqRevenueSectorMarkdown(markdown = "", companyName = "") {
  const lines = splitMarkdownLines(markdown);
  const result = {
    company_name: companyName || "",
    hq_location: "",
    revenue_estimate: "",
    industry_sector: "",
  };

  lines.forEach((line) => {
    const hqMatch = line.match(/^hq location\s*:\s*(.+)$/i);
    if (hqMatch) {
      result.hq_location = hqMatch[1].trim();
      return;
    }
    const revMatch = line.match(/^revenue\s*:\s*(.+)$/i);
    if (revMatch) {
      result.revenue_estimate = revMatch[1].trim();
      return;
    }
    const secMatch = line.match(/^sector\s*:\s*(.+)$/i);
    if (secMatch) {
      result.industry_sector = secMatch[1].trim();
    }
  });

  return result;
}

async function fetchHqRevenueAndSector({ company, locationHint, product, docsText }) {
  if (!company) {
    return { error: "Missing company name", overview: { company_name: "", hq_location: "", revenue_estimate: "", industry_sector: "" }, rawText: "" };
  }

  const normalizedLocationHint = locationHint ? String(locationHint).trim() : "";
  let prompt = "";
  try {
    const buildHqRevenueSectorPrompt = requirePromptBuilder("buildHqRevenueSectorPrompt");
    prompt = buildHqRevenueSectorPrompt({
      company,
      locationHint: normalizedLocationHint,
      product,
      docsText,
    });
  } catch (err) {
    return { error: err?.message || String(err), overview: { company_name: company, hq_location: "", revenue_estimate: "", industry_sector: "" }, rawText: "" };
  }

  try {
    const resp = await callLlmWithRetry(prompt, {
      model: GROQ_COMPOUND_MINI_MODEL,
      secondaryModel: GROQ_COMPOUND_MINI_MODEL,
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 400,
      },
      compoundCustom: {
        tools: {
          enabled_tools: ["web_search"],
        },
      },
    });

    if (resp.error) {
      return { error: resp.error + (resp.details ? " Details: " + JSON.stringify(resp.details) : ""), rawText: "", overview: { company_name: company, hq_location: "", revenue_estimate: "", industry_sector: "" } };
    }

    const rawText = typeof resp.text === "string" ? resp.text : "";
    const parsed = parseHqRevenueSectorMarkdown(rawText, company);
    const missingAll = !parsed.hq_location && !parsed.revenue_estimate && !parsed.industry_sector;
    const error = missingAll ? "Model did not return HQ, revenue, or sector." : "";

    return { overview: parsed, rawText, error };
  } catch (err) {
    return { error: err?.message || String(err), rawText: "", overview: { company_name: company, hq_location: "", revenue_estimate: "", industry_sector: "" } };
  }
}

async function fetchRecentNewsEntries({ company, location, docsText }) {
  let prompt = "";
  try {
    const buildRecentNewsPrompt = requirePromptBuilder("buildRecentNewsPrompt");
    prompt = buildRecentNewsPrompt({ company, location, docsText });
  } catch (err) {
    return { error: err?.message || String(err), attempts: [] };
  }

  try {
    const resp = await callLlmWithRetry(prompt, {
      model: GROQ_COMPOUND_MINI_MODEL,
      secondaryModel: GROQ_COMPOUND_MINI_MODEL,
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 4000,
      },
    });

    if (resp.error) {
      return { error: resp.error + (resp.details ? " Details: " + JSON.stringify(resp.details) : ""), attempts: resp.details || [] };
    }

    const rawText = typeof resp.text === "string" ? resp.text : "";
    let topNews = parseTopNewsMarkdown(rawText);

    if (!topNews.length) {
      const { parsed, rawText: parsedRaw } = parseModelJsonResponse(resp);
      const fallbackParsed =
        (Array.isArray(parsed?.top_5_news) && parsed.top_5_news) ||
        (Array.isArray(parsed?.top_news) && parsed.top_news) ||
        (Array.isArray(parsed?.news) && parsed.news) ||
        [];
      topNews = fallbackParsed;
    }

    return { topNews: normalizeTopNewsEntries(topNews), rawText };
  } catch (err) {
    return { error: err?.message || String(err) };
  }
}

async function generateBriefOverview({ company, location, product, docsText }) {
  const [coreResult, newsResult] = await Promise.all([
    fetchHqRevenueAndSector({ company, locationHint: location, product, docsText }),
    fetchRecentNewsEntries({ company, location, docsText }),
  ]);

  const overview = {
    company_name: coreResult?.overview?.company_name || company || "",
    hq_location: coreResult?.overview?.hq_location || "",
    revenue_estimate: coreResult?.overview?.revenue_estimate || "",
    industry_sector: coreResult?.overview?.industry_sector || "",
    top_5_news: Array.isArray(newsResult?.topNews) ? newsResult.topNews : [],
  };

  const rawSegments = [];
  if (coreResult?.rawText) rawSegments.push(coreResult.rawText);
  if (newsResult?.rawText) rawSegments.push(newsResult.rawText);
  const rawText = rawSegments.join("\n\n---\n\n");

  const errors = [];
  if (coreResult?.error) errors.push(`overview: ${coreResult.error}`);
  if (newsResult?.error) errors.push(`news: ${newsResult.error}`);

  const response = { overview, rawText };
  if (errors.length) {
    response.error = errors.join(" | ");
    response.hq_error = coreResult?.error || "";
  }
  return response;
}

async function generatePersonaBrief({
  company,
  location,
  product,
  docsText,
  runId,
  includeEmails = true,
  includeTelephonicPitch = true,
  onPersonasDone,
  onEmailsDone,
  onTelephonicDone,
}) {
  const pitchFromCompany = await loadPitchingCompany();
  const pitchingOrg = pitchFromCompany || "your company";
  const prospectLabel = company || "the target company";
  let prompt = "";
  try {
    const buildPersonaBriefPrompt = requirePromptBuilder("buildPersonaBriefPrompt");
    prompt = buildPersonaBriefPrompt({
      prospectLabel,
      pitchingOrg,
      location,
      product,
      docsText,
      companyName: company,
    });
  } catch (err) {
    return { error: err?.message || String(err), attempts: [] };
  }

  const resp = await callLlmWithRetry(prompt, {
    model: LLAMA_33_MODEL,
    secondaryModel: LLAMA_33_MODEL,
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 6000,
    },
  });

  if (resp.error) {
    return { error: resp.error + (resp.details ? " Details: " + JSON.stringify(resp.details) : ""), attempts: resp.details || [] };
  }

  const rawText = typeof resp.text === "string" ? resp.text : "";
  const personas = parsePersonasFromMarkdown(rawText, company, product);

  if (!Array.isArray(personas) || !personas.length) {
    return { error: "Model did not return personas.", rawText };
  }

  if (runId) {
    emitBriefPartialUpdate(runId, {
      personas,
    });
  }
  if (typeof onPersonasDone === "function") {
    onPersonasDone({ personas });
  }

  const emailPromise = includeEmails
    ? generatePersonaEmails({
        personas,
        company,
        location,
        product,
        docsText,
        pitchFromCompany: pitchingOrg,
      })
    : Promise.resolve({ personaEmails: [], attempts: [], error: "" });
  const telephonicPromise = includeTelephonicPitch
    ? generateTelephonicPitchScripts({
        personas,
        company,
        location,
        product,
        docsText,
        pitchFromCompany: pitchingOrg,
      })
    : Promise.resolve({ pitches: [], attempts: [], error: "" });

  if (runId) {
    if (includeEmails) {
      emailPromise
        .then((res) => {
          const personaEmails = Array.isArray(res?.personaEmails) ? res.personaEmails : [];
          const email = derivePrimaryEmail({ persona_emails: personaEmails }, personaEmails);
          emitBriefPartialUpdate(runId, {
            personaEmails,
            email,
          });
        })
        .catch(() => {});
    }

    if (includeTelephonicPitch) {
      telephonicPromise
        .then((res) => {
          const telephonicPitches = Array.isArray(res?.pitches) ? res.pitches : [];
          emitBriefPartialUpdate(runId, {
            telephonicPitches,
            telephonicPitchError: res?.error || "",
            telephonicPitchAttempts: res?.attempts || [],
          });
        })
        .catch(() => {});
    }
  }

  if (includeEmails && typeof onEmailsDone === "function") {
    emailPromise.finally(() => onEmailsDone());
  }
  if (includeTelephonicPitch && typeof onTelephonicDone === "function") {
    telephonicPromise.finally(() => onTelephonicDone());
  }

  const [emailResult, telephonicResult] = await Promise.allSettled([emailPromise, telephonicPromise]);

  const resolvedEmailResult =
    emailResult.status === "fulfilled"
      ? emailResult.value
      : { personaEmails: [], error: emailResult.reason?.message || String(emailResult.reason) };

  const resolvedTelephonicResult =
    telephonicResult.status === "fulfilled"
      ? telephonicResult.value
      : { pitches: [], error: telephonicResult.reason?.message || String(telephonicResult.reason), attempts: [] };

  const personaEmails = Array.isArray(resolvedEmailResult.personaEmails) ? resolvedEmailResult.personaEmails : [];
  const telephonicPitches = Array.isArray(resolvedTelephonicResult.pitches)
    ? resolvedTelephonicResult.pitches
    : [];

  const telephonicPitchError = includeTelephonicPitch
    ? (resolvedTelephonicResult.error || (!telephonicPitches.length ? "No telephonic pitch generated." : "") || "")
    : "";
  const telephonicPitchAttempts = includeTelephonicPitch ? (resolvedTelephonicResult.attempts || []) : [];

  const email = includeEmails ? derivePrimaryEmail({ persona_emails: personaEmails }, personaEmails) : { subject: "", body: "" };

  return {
    personas,
    personaEmails: includeEmails ? personaEmails : [],
    telephonicPitches: includeTelephonicPitch ? telephonicPitches : [],
    telephonicPitchError,
    telephonicPitchAttempts,
    email,
    rawText,
  };
}

async function fetchProductContext({ product }) {
  const trimmedProduct = typeof product === "string" ? product.trim() : "";
  if (!trimmedProduct) return { context: "", rawText: "", error: "Product name is missing." };

  let prompt = "";
  try {
    const buildProductContextPrompt = requirePromptBuilder("buildProductContextPrompt");
    prompt = buildProductContextPrompt({ product: trimmedProduct });
  } catch (err) {
    return { context: "", rawText: "", error: err?.message || String(err) };
  }

  try {
    const resp = await callLlmWithRetry(prompt, {
      model: GROQ_COMPOUND_MINI_MODEL,
      secondaryModel: GROQ_COMPOUND_MINI_MODEL,
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 600,
      },
      compoundCustom: {
        tools: {
          enabled_tools: ["web_search"],
        },
      },
    });

    if (resp.error) {
      return { context: "", rawText: "", error: resp.error };
    }

    const rawText = typeof resp.text === "string" ? resp.text.trim() : "";
    return { context: rawText, rawText, error: "" };
  } catch (err) {
    return { context: "", rawText: "", error: err?.message || String(err) };
  }
}

async function generateBrief({ company, location, product, docs = [], runId, modules }) {
  try {
    const resolvedModules = resolveBriefModules(modules);
    const moduleSteps = getBriefProgressTotal(resolvedModules);
    if (!moduleSteps) {
      return { error: "No brief sections enabled.", modules: resolvedModules };
    }
    let totalSteps = moduleSteps;
    let completedSteps = 0;
    emitBriefProgress({
      runId,
      current: completedSteps,
      total: totalSteps,
      label: "Starting brief request",
    });

    let docsText = (docs || []).map(d => {
      const txt = decodeBase64Text(d.content_b64 || d.content || "");
      return `--- ${d.name || "doc"} ---\n${txt.substring(0, 4000)}`;
    }).join("\n\n");

    if (
      !docsText.trim() &&
      product &&
      (resolvedModules[BRIEF_MODULES.OVERVIEW] ||
        resolvedModules[BRIEF_MODULES.TOP_NEWS] ||
        resolvedModules[BRIEF_MODULES.PERSONAS] ||
        resolvedModules[BRIEF_MODULES.EMAILS] ||
        resolvedModules[BRIEF_MODULES.TELE_PITCH])
    ) {
      const productContextResult = await fetchProductContext({ product });
      if (productContextResult.context) {
        docsText = `--- Product context (web search) ---\n${productContextResult.context}`;
      }
    }

    const stepDone = (label) => {
      completedSteps = Math.min(totalSteps, completedSteps + 1);
      emitBriefProgress({
        runId,
        current: completedSteps,
        total: totalSteps,
        label,
      });
    };

    const overviewPromise = resolvedModules[BRIEF_MODULES.OVERVIEW]
      ? fetchHqRevenueAndSector({ company, locationHint: location, product, docsText })
        .then((res) => {
          stepDone(res.error ? "Overview generation failed" : "Overview generated");
          const overview = res.overview || {};
          emitBriefPartialUpdate(runId, {
            overview,
            hq_location: overview.hq_location || "",
            hq_lookup_error: res.hq_error || res.error || "",
          });
          if (res.error) {
            emitBriefPartialUpdate(runId, { overview_error: res.error });
          }
          return res;
        })
        .catch((err) => {
          const errorMsg = err?.message || String(err);
          stepDone("Overview generation failed");
          emitBriefPartialUpdate(runId, { overview_error: errorMsg, hq_lookup_error: errorMsg });
          return { error: errorMsg, overview: { company_name: company, hq_location: "", revenue_estimate: "", industry_sector: "" } };
        })
      : Promise.resolve({ overview: { company_name: company || "", hq_location: "", revenue_estimate: "", industry_sector: "" }, rawText: "", error: "" });

    const newsPromise = resolvedModules[BRIEF_MODULES.TOP_NEWS]
      ? fetchRecentNewsEntries({ company, location, docsText })
        .then((res) => {
          stepDone(res.error ? "Top news failed" : "Top news generated");
          const topNews = Array.isArray(res?.topNews) ? res.topNews : [];
          emitBriefPartialUpdate(runId, { top_5_news: topNews });
          if (res.error) {
            emitBriefPartialUpdate(runId, { overview_error: res.error });
          }
          return res;
        })
        .catch((err) => {
          const errorMsg = err?.message || String(err);
          stepDone("Top news failed");
          emitBriefPartialUpdate(runId, { overview_error: errorMsg });
          return { error: errorMsg, topNews: [], rawText: "" };
        })
      : Promise.resolve({ topNews: [], rawText: "", error: "" });

    let personasStepDone = false;
    let emailStepDone = false;
    let telephonicStepDone = false;
    const personaPromise = resolvedModules[BRIEF_MODULES.PERSONAS]
      ? generatePersonaBrief({
        company,
        location,
        product,
        docsText,
        runId,
        includeEmails: resolvedModules[BRIEF_MODULES.EMAILS],
        includeTelephonicPitch: resolvedModules[BRIEF_MODULES.TELE_PITCH],
        onPersonasDone: () => {
          if (personasStepDone) return;
          personasStepDone = true;
          stepDone("Personas generated");
        },
        onEmailsDone: resolvedModules[BRIEF_MODULES.EMAILS]
          ? () => {
              if (emailStepDone) return;
              emailStepDone = true;
              stepDone("Emails generated");
            }
          : null,
        onTelephonicDone: resolvedModules[BRIEF_MODULES.TELE_PITCH]
          ? () => {
              if (telephonicStepDone) return;
              telephonicStepDone = true;
              stepDone("Telephonic pitches generated");
            }
          : null,
      })
        .then((res) => {
          if (res.error && !personasStepDone) {
            personasStepDone = true;
            stepDone("Persona generation failed");
          }
          if (res.error) {
            emitBriefPartialUpdate(runId, { persona_error: res.error });
            if (resolvedModules[BRIEF_MODULES.EMAILS] && !emailStepDone) {
              emailStepDone = true;
              stepDone("Emails skipped");
            }
            if (resolvedModules[BRIEF_MODULES.TELE_PITCH] && !telephonicStepDone) {
              telephonicStepDone = true;
              stepDone("Telephonic pitches skipped");
            }
          }
          return res;
        })
        .catch((err) => {
          const errorMsg = err?.message || String(err);
          if (!personasStepDone) {
            personasStepDone = true;
            stepDone("Persona generation failed");
          }
          if (resolvedModules[BRIEF_MODULES.EMAILS] && !emailStepDone) {
            emailStepDone = true;
            stepDone("Emails skipped");
          }
          if (resolvedModules[BRIEF_MODULES.TELE_PITCH] && !telephonicStepDone) {
            telephonicStepDone = true;
            stepDone("Telephonic pitches skipped");
          }
          emitBriefPartialUpdate(runId, { persona_error: errorMsg });
          return {
            error: errorMsg,
            personas: [],
            personaEmails: [],
            telephonicPitches: [],
            telephonicPitchError: "",
            telephonicPitchAttempts: [],
            email: { subject: "", body: "" },
          };
        })
      : Promise.resolve({
        personas: [],
        personaEmails: [],
        telephonicPitches: [],
        telephonicPitchError: "",
        telephonicPitchAttempts: [],
        email: { subject: "", body: "" },
        error: "",
      });

    const [overviewResult, newsResult, personaResult] = await Promise.all([overviewPromise, newsPromise, personaPromise]);

    completedSteps = Math.max(completedSteps, totalSteps);
    emitBriefProgress({
      runId,
      current: totalSteps,
      total: totalSteps,
      label: "Brief ready",
    });

    const overview = overviewResult?.overview || { company_name: company || "", hq_location: "", revenue_estimate: "", industry_sector: "" };
    const topNews = resolvedModules[BRIEF_MODULES.TOP_NEWS] && Array.isArray(newsResult?.topNews) ? newsResult.topNews : [];
    const industrySector = resolvedModules[BRIEF_MODULES.OVERVIEW]
      ? (overview.industry_sector || overview.industrySector || overview.industry || "")
      : "";
    const revenueEstimate = resolvedModules[BRIEF_MODULES.OVERVIEW] ? (overview.revenue_estimate || "") : "";

    const resolvedHqLocation = resolvedModules[BRIEF_MODULES.OVERVIEW] ? (overview.hq_location || "") : "";
    const hqErrorMessage = resolvedModules[BRIEF_MODULES.OVERVIEW] && overviewResult?.error
      ? `HQ lookup failed: ${overviewResult.error}`
      : "";
    const displayedHq = resolvedHqLocation || hqErrorMessage || "Not found";
    const displayedIndustry = industrySector || "Not found";

    const personas = Array.isArray(personaResult?.personas) ? personaResult.personas : [];
    const personaEmails = Array.isArray(personaResult?.personaEmails) ? personaResult.personaEmails : [];
    const telephonicPitches = Array.isArray(personaResult?.telephonicPitches)
      ? personaResult.telephonicPitches
      : [];
    const telephonicPitchError = resolvedModules[BRIEF_MODULES.TELE_PITCH] ? (personaResult?.telephonicPitchError || "") : "";
    const telephonicPitchAttempts = resolvedModules[BRIEF_MODULES.TELE_PITCH] && Array.isArray(personaResult?.telephonicPitchAttempts)
      ? personaResult.telephonicPitchAttempts
      : [];
    const emailObj = personaResult?.email || { subject: "", body: "" };
    const personaEmailVersions = resolvedModules[BRIEF_MODULES.EMAILS]
      ? personaEmails.map((draft) => ({ versions: [draft], activeIndex: 0 }))
      : [];
    const telephonicPitchVersions = resolvedModules[BRIEF_MODULES.TELE_PITCH]
      ? telephonicPitches.map((draft) => ({ versions: [draft], activeIndex: 0 }))
      : [];

    const brief_html = resolvedModules[BRIEF_MODULES.OVERVIEW]
      ? buildBriefHtmlFromOverview({
          companyName: overview.company_name || company,
          hqDisplay: displayedHq,
          revenue: revenueEstimate,
          industry: displayedIndustry,
          topNews,
        })
      : "";

    const overviewErrors = [];
    if (resolvedModules[BRIEF_MODULES.OVERVIEW] && overviewResult?.error) {
      overviewErrors.push(`overview: ${overviewResult.error}`);
    }
    if (resolvedModules[BRIEF_MODULES.TOP_NEWS] && newsResult?.error) {
      overviewErrors.push(`news: ${newsResult.error}`);
    }
    const overviewError = overviewErrors.join(" | ");
    const personaError = resolvedModules[BRIEF_MODULES.PERSONAS] ? (personaResult?.error || "") : "";
    const combinedError = overviewError && personaError ? `${overviewError}; ${personaError}` : "";

    return {
      brief_html,
      modules: resolvedModules,
      company_name: overview.company_name || company,
      revenue_estimate: revenueEstimate,
      industry_sector: industrySector,
      top_5_news: topNews,
      personas,
      personaEmails,
      personaEmailVersions,
      telephonicPitches,
      telephonicPitchVersions,
      personaEmailVersionIndexes: personaEmailVersions.map((v) => v.activeIndex || 0),
      telephonicPitchVersionIndexes: telephonicPitchVersions.map((v) => v.activeIndex || 0),
      telephonicPitchError,
      telephonicPitchAttempts,
      email: emailObj,
      hq_location: resolvedHqLocation,
      hq_lookup_details: null,
      hq_lookup_error: resolvedModules[BRIEF_MODULES.OVERVIEW] ? (overviewResult?.error || "") : "",
      hq_lookup_raw: resolvedModules[BRIEF_MODULES.OVERVIEW] ? (overviewResult?.rawText || "") : "",
      raw_overview: resolvedModules[BRIEF_MODULES.OVERVIEW] ? (overviewResult?.rawText || "") : "",
      raw_personas: resolvedModules[BRIEF_MODULES.PERSONAS] ? (personaResult?.rawText || "") : "",
      overview_error: overviewError,
      persona_error: personaError,
      error: combinedError || undefined,
    };
  } catch (err) {
    return { error: String(err), modules: resolveBriefModules(modules) };
  }
}
chrome.action.onClicked.addListener((tab) => {
  chrome.tabs.create({
    url: chrome.runtime.getURL("popup.html")
  });
});

chrome.runtime.onMessage.addListener((req, sender, sendResponse) => {
  (async () => {
    try {
      if (req && req.action) {
        if (req.action === 'storeDoc') {
          const data = await chrome.storage.local.get(['docs']);
          const docs = data.docs || [];
          const id = Date.now().toString();
          docs.push({ id, name: req.name, content_b64: req.content_b64 });
          await chrome.storage.local.set({ docs });
          sendResponse({ ok: true, id });
          return;
        }
        if (req.action === 'listDocs') {
          const data = await chrome.storage.local.get(['docs']);
          const docs = data.docs || [];
          sendResponse({ docs });
          return;
        }
        if (req.action === 'getDocsForProduct') {
          const data = await chrome.storage.local.get(['docs']);
          const docs = data.docs || [];
          const product = (req.product || '').toLowerCase();
          const filtered = docs.filter(d => d.name && d.name.toLowerCase().includes(product));
          sendResponse({ docs: filtered });
          return;
        }
        if (req.action === 'generateTargets') {
          markWeeklyActiveUser("generateTargets").catch(() => {});
          const result = await generateTargets({
            product: req.product,
            location: req.location,
            sectors: req.sectors,
            docs: Array.isArray(req.docs) ? req.docs : [],
            docName: req.docName,
            docText: req.docText,
            docBase64: req.docBase64,
          });
          if (result && result.ok) {
            await saveTargetHistoryEntry(
              { product: req.product, location: req.location, docName: req.docName, docs: req.docs, sectors: req.sectors },
              { companies: result.companies }
            );
          } else if (result && result.error) {
            const errorMetric = buildGenerationErrorMetric({
              feature: "targets",
              error: [result.error, typeof result.details === "string" ? result.details : ""].filter(Boolean).join(" | "),
              request: {
                product: req.product,
                location: req.location,
                sectors: req.sectors,
                docs: req.docs,
              },
            });
            enqueueTelemetryMetric(errorMetric).catch(() => {});
            flushTelemetryQueue().catch(() => {});
          }
          sendResponse(result);
          return;
        }
        if (req.action === 'generateBrief') {
          markWeeklyActiveUser("generateBrief").catch(() => {});
          const payload = {
            company: req.company,
            location: req.location,
            product: req.product,
            docs: req.docs || [],
            runId: req.runId,
            modules: req.modules,
          };
          const startTs = Date.now();
          let result = null;
          let success = false;

          try {
            result = await generateBrief(payload);
            success = !result?.error;
          } catch (err) {
            result = { error: err?.message || String(err) };
          }

          if (result && !result.error) {
            await saveResearchHistoryEntry(payload, result);
          }

          const metric = buildResearchCycleMetric({
            startTimeMs: startTs,
            endTimeMs: Date.now(),
            request: payload,
            runId: payload.runId,
            success,
            error: result?.error || "",
          });
          const telemetryTasks = [enqueueTelemetryMetric(metric)];
          if (result?.error) {
            const errorMetric = buildGenerationErrorMetric({
              feature: "briefs",
              error: result.error,
              request: payload,
              runId: payload.runId,
            });
            telemetryTasks.push(enqueueTelemetryMetric(errorMetric));
          }
          Promise.all(telemetryTasks).catch(() => {});
          flushTelemetryQueue().catch(() => {});

          sendResponse(result);
          return;
        }
        if (req.action === "revisePersonaEmail") {
          const result = await revisePersonaEmail({
            persona: req.persona,
            email: req.email,
            company: req.company,
            product: req.product,
            location: req.location,
            instructions: req.instructions,
            pitchingOrg: req.pitchingOrg,
          });
          sendResponse(result);
          return;
        }
        if (req.action === "reviseAllPersonaEmails") {
          const result = await reviseAllPersonaEmails({
            personas: Array.isArray(req.personas) ? req.personas : [],
            emails: Array.isArray(req.emails) ? req.emails : [],
            company: req.company,
            product: req.product,
            location: req.location,
            instructions: req.instructions,
            pitchingOrg: req.pitchingOrg,
          });
          sendResponse({ results: result });
          return;
        }
        if (req.action === "revisePersonaPitch") {
          const result = await revisePersonaPitch({
            persona: req.persona,
            pitch: req.pitch,
            company: req.company,
            product: req.product,
            location: req.location,
            instructions: req.instructions,
            pitchingOrg: req.pitchingOrg,
          });
          sendResponse(result);
          return;
        }
        if (req.action === "reviseAllPersonaPitches") {
          const result = await reviseAllPersonaPitches({
            personas: Array.isArray(req.personas) ? req.personas : [],
            pitches: Array.isArray(req.pitches) ? req.pitches : [],
            company: req.company,
            product: req.product,
            location: req.location,
            instructions: req.instructions,
            pitchingOrg: req.pitchingOrg,
          });
          sendResponse({ results: result });
          return;
        }
        if (req.action === "updateResearchHistoryEntry") {
          const updateResult = await updateResearchHistoryEntry(req.id, req.result || {});
          sendResponse(updateResult);
          return;
        }
        if (req.action === "renameHistoryEntry") {
          const storageKey = req.historyType === "target" ? TARGET_HISTORY_KEY : RESEARCH_HISTORY_KEY;
          const renameResult = await renameHistoryEntry(storageKey, req.id, req.title || "");
          sendResponse(renameResult);
          return;
        }
        if (req.action === "deleteHistoryEntry") {
          const storageKey = req.historyType === "target" ? TARGET_HISTORY_KEY : RESEARCH_HISTORY_KEY;
          const deleteResult = await deleteHistoryEntry(storageKey, req.id);
          sendResponse(deleteResult);
          return;
        }
        if (req.action === 'getTargetHistory') {
          const data = await chrome.storage.local.get([TARGET_HISTORY_KEY]);
          const history = Array.isArray(data[TARGET_HISTORY_KEY]) ? data[TARGET_HISTORY_KEY] : [];
          sendResponse({ history });
          return;
        }
        if (req.action === 'exportResearch') {
          markWeeklyActiveUser("exportResearch").catch(() => {});
          const selection = req.selection || { type: "all" };
          let format = req.format === "md" ? "md" : "xlsx";

          let activeTemplate = req.template;
          const stored = await chrome.storage.local.get([EXPORT_TEMPLATES_KEY, EXPORT_TEMPLATE_KEY]);
          if (!activeTemplate || !Array.isArray(activeTemplate.columns) || !activeTemplate.columns.length) {
            const collection = stored && stored[EXPORT_TEMPLATES_KEY];
            if (collection && Array.isArray(collection.templates) && collection.templates.length) {
              const selectedId = req.templateId || collection.selectedTemplateId || collection.activeTemplateId;
              const fromCollection = collection.templates.find((tpl) => tpl.id === selectedId) || collection.templates[0];
              if (fromCollection && Array.isArray(fromCollection.columns) && fromCollection.columns.length) {
                activeTemplate = fromCollection;
              }
            }
          }
          if (!activeTemplate || !Array.isArray(activeTemplate.columns) || !activeTemplate.columns.length) {
            const storedTemplate = stored && stored[EXPORT_TEMPLATE_KEY];
            if (storedTemplate && Array.isArray(storedTemplate.columns) && storedTemplate.columns.length) {
              activeTemplate = storedTemplate;
            }
          }

          if (activeTemplate && (!req.format || req.format === "")) {
            if (activeTemplate.format === "md") {
              format = "md";
            } else if (activeTemplate.format === "xlsx") {
              format = "xlsx";
            }
          }

          if (!activeTemplate || !Array.isArray(activeTemplate.columns) || !activeTemplate.columns.length) {
            sendResponse({ error: "No export template found. Please add export columns in settings." });
            return;
          }

          const columns = activeTemplate.columns
            .map((col, idx) => {
              const header = (col && col.header ? String(col.header) : "").trim();
              const descriptionRaw = col && col.description ? String(col.description) : "";
              const description = descriptionRaw.trim() || `User defined column ${idx + 1}`;
              return header ? { header, description } : null;
            })
            .filter(Boolean);

          if (!columns.length) {
            sendResponse({ error: "Export template must include at least one column header." });
            return;
          }

          const storedData = await chrome.storage.local.get([RESEARCH_HISTORY_KEY]);
          const history = Array.isArray(storedData[RESEARCH_HISTORY_KEY]) ? storedData[RESEARCH_HISTORY_KEY] : [];

          if (!history.length) {
            sendResponse({ error: "No research history available to export." });
            return;
          }

          const filteredEntries = filterHistoryEntries(history, selection);
          if (!filteredEntries.length) {
            sendResponse({ error: "No research entries match the selected range." });
            return;
          }

  const prompt = composeExportPrompt(columns, filteredEntries, format);
  const llmResult = await callLlmWithRetry(prompt, {
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 4096,
    },
  });

          if (llmResult.error) {
            sendResponse({ error: llmResult.error, details: llmResult.details });
            return;
          }

          const rawText = typeof llmResult.text === "string" ? llmResult.text : "";
          let notes = "";
          let tableRows = [];

          const { headers: parsedHeaders, rows: parsedRows } = parseMarkdownTable(rawText);
          if (parsedRows && parsedRows.length) {
            tableRows = parsedRows;
          }
          const notesMatch = rawText.match(/Notes:\s*(.+)/i);
          if (notesMatch && notesMatch[1]) {
            notes = notesMatch[1].trim();
          }

          if (!tableRows.length) {
            const { parsed } = parseModelJsonResponse(llmResult);
            if (parsed && Array.isArray(parsed.rows)) {
              tableRows = parsed.rows;
              notes = parsed.notes || "";
            }
          }

          if (!tableRows.length) {
            sendResponse({ error: "Model did not return structured rows for export.", details: rawText || null });
            return;
          }

          const normalizedRows = tableRows.map((row) => ensureRowValues(row, columns));
          const headers = columns.map((col) => col.header);

          let markdownTable = "";
          if (format === "md") {
            if (parsedHeaders && parsedHeaders.length) {
              const body = normalizedRows
                .map((row) => `| ${headers.map((h) => row[h] || "").join(" | ")} |`)
                .join("\n");
              const separator = `| ${headers.map(() => "---").join(" | ")} |`;
              markdownTable = `| ${headers.join(" | ")} |\n${separator}${body ? `\n${body}` : ""}`;
            } else {
              markdownTable = generateMarkdownFromRows(headers, normalizedRows);
            }
          }

          let base64Data = "";
          let mimeType = "";
          let filename = "";

          if (format === "md") {
            const exportLines = [];
            exportLines.push(`# Research Export`);
            exportLines.push(`Generated: ${new Date().toISOString()}`);
            exportLines.push("");
            if (markdownTable) {
              exportLines.push(markdownTable);
            } else {
              exportLines.push(generateMarkdownFromRows(headers, normalizedRows));
            }
            if (notes) {
              exportLines.push("");
              exportLines.push(`> ${notes}`);
            }
            const markdownContent = exportLines.join("\n");
            base64Data = stringToBase64(markdownContent);
            mimeType = "text/markdown";
            filename = `research-export-${Date.now()}.md`;
          } else {
            const workbookBytes = buildXlsxFile(headers, normalizedRows);
            base64Data = uint8ToBase64(workbookBytes);
            mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            filename = `research-export-${Date.now()}.xlsx`;
          }

          sendResponse({
            ok: true,
            totalRows: normalizedRows.length,
            preview: {
              headers,
              rows: normalizedRows.slice(0, 10),
            },
            notes: notes || "",
            download: {
              format,
              mimeType,
              filename,
              base64: base64Data,
            },
          });

          const exportMetric = buildSearchToExportConversionMetric({
            exportId: filename,
            occurredAtMs: Date.now(),
            selectionType: selection.type,
            selectionCount: filteredEntries.length,
            format,
            templateId: req.templateId || activeTemplate?.id || "",
            researchEntryIds: filteredEntries.map((entry) => entry && entry.id).filter(Boolean),
            resultRowCount: normalizedRows.length,
          });
          enqueueTelemetryMetric(exportMetric).catch(() => {});
          flushTelemetryQueue().catch(() => {});
          return;
        }
        if (req.action === 'getResearchHistory') {
          const data = await chrome.storage.local.get([RESEARCH_HISTORY_KEY]);
          const history = Array.isArray(data[RESEARCH_HISTORY_KEY]) ? data[RESEARCH_HISTORY_KEY] : [];
          sendResponse({ history });
          return;
        }
      }
      sendResponse({ error: "Unknown action/type" });
    } catch (err) {
      sendResponse({ error: err.message || String(err) });
    }
  })();
  return true;
});

// Export selected LLM helpers for Node-based tests without affecting extension runtime.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    callGroqDirect,
    callGroqWithRetry,
    callGeminiDirect,
    callGeminiWithRetryInternal,
    callLlmWithRetry,
    parseModelJsonResponse,
    buildGroqMessages,
    translateToolsToGroq,
    parsePersonasFromMarkdown,
    parsePersonaEmailMarkdown,
    parseTelephonicPitchMarkdown,
    parseRevenueSectorMarkdown,
    parseHqLocationMarkdown,
    parseTopNewsMarkdown,
    parseTargetCompaniesMarkdown,
    parseMarkdownTable,
    toLinkedInPeopleSearchUrl,
    buildLinkedInKeywordFallback,
    buildZoomInfoSearchLink,
  };
}

})(); // end backgroundWebScope

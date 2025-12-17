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

async function callGeminiDirect(promptText, opts = {}) {
  const data = await chrome.storage.local.get("geminiKey");
  const geminiKey = data && data.geminiKey;
  if (!geminiKey) {
    return { error: "No Gemini API key found. Please add it in the popup." };
  }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiKey}`;

  const headers = {
    "Content-Type": "application/json",
  };

  // Merge caller config with safer, deterministic defaults and clamps
  const userCfg = opts.generationConfig || {};
  let generationConfig = {
    ...userCfg,
    // Keep temperature low for determinism; clamp to [0, 0.2]
    temperature: Math.max(0, Math.min(userCfg.temperature ?? 0.1, 0.2)),
    // Allow large outputs while bounded; default to 100000, clamp to 100000 max
    maxOutputTokens: Math.min(userCfg.maxOutputTokens ?? 100000, 100000),
    // Single candidate to reduce variance
    candidateCount: 1,
  };
    
  const customContents = Array.isArray(opts.contents) && opts.contents.length ? opts.contents : null;
  const requestedTools = Array.isArray(opts.tools) && opts.tools.length ? opts.tools.filter(Boolean) : null;
  const hasTools = !!(requestedTools && requestedTools.length);
  const allowDefaultTools = !opts.disableDefaultTools;

  // Tools + responseMimeType=json is unsupported; drop the mime hint if tools are requested.
  if (hasTools && generationConfig.responseMimeType === "application/json") {
    generationConfig = { ...generationConfig };
    delete generationConfig.responseMimeType;
  }

  const isStructured = generationConfig.responseMimeType === "application/json";

  const body = {
    contents: customContents || [{ role: "user", parts: [{ text: promptText || "" }] }],
    generationConfig: generationConfig,
  };

  if (hasTools) {
    body.tools = requestedTools;
  } else if (!isStructured && allowDefaultTools) {
    body.tools = [{ google_search: {} }];
  }

  if (opts.toolConfig && typeof opts.toolConfig === "object") {
    body.toolConfig = opts.toolConfig;
  }

  try {
    const resp = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });

    const respText = await resp.text();
    let respJson = null;

    try {
      respJson = JSON.parse(respText);
    } catch (e) {
      if (!resp.ok) {
        return { error: `Gemini API error (Status ${resp.status}): ${respText}` };
      }
      return { error: "Failed to parse Gemini API response as JSON.", details: respText };
    }

    if (!resp.ok) {
      const errorDetails = respJson?.error?.message || respText;
      return { error: `Gemini API error: ${errorDetails}`, details: respJson };
    }

    const candidate = respJson?.candidates?.[0];
    const parts = Array.isArray(candidate?.content?.parts) ? candidate.content.parts : [];
    const textParts = [];
    const structuredParts = [];
    parts.forEach((part) => {
      if (typeof part?.text === "string" && part.text.trim()) {
        textParts.push(part.text);
      } else if (part?.functionCall?.args) {
        try {
          structuredParts.push(JSON.stringify(part.functionCall.args));
        } catch (err) {
          // Ignore parse issues
        }
      } else if (part?.functionResponse?.response) {
        try {
          structuredParts.push(JSON.stringify(part.functionResponse.response));
        } catch (err) {
          // Ignore parse issues
        }
      } else if (part?.inlineData?.mimeType === "application/json" && part?.inlineData?.data) {
        const decoded = decodeBase64Text(part.inlineData.data);
        if (decoded) structuredParts.push(decoded);
      }
    });

    let outputText = textParts.join("\n").trim();
    if (!outputText && structuredParts.length) {
      outputText = structuredParts.join("\n");
    }

    if (!outputText) {
      if (candidate?.finishReason) {
        return { error: `Gemini generation stopped: ${candidate.finishReason}`, details: respJson };
      }
      return { error: "Could not find text in Gemini response.", details: respJson };
    }

    return { ok: true, text: outputText, raw: respJson };
  } catch (err) {
    return { error: `Network request failed: ${String(err)}` };
  }
}

async function callGeminiWithRetry(promptText, opts = {}) {
  const baseConfig = opts.generationConfig || {};
  const primary = await callGeminiDirect(promptText, opts);
  if (!primary.error) {
    return primary;
  }

  const attempts = [
    {
      label: "primary",
      error: primary.error,
      details: primary.details || primary.raw || null,
    },
  ];

  const boostedMaxTokens = Math.min(
    typeof baseConfig.maxOutputTokens === "number" ? Math.round(baseConfig.maxOutputTokens * 1.5) : 12000,
    100000
  );

  const retry = await callGeminiDirect(promptText, {
    ...opts,
    generationConfig: {
      ...baseConfig,
      maxOutputTokens: boostedMaxTokens,
    },
    tools: [],
    disableDefaultTools: true,
  });

  if (retry.error) {
    attempts.push({
      label: "retry_no_tools",
      error: retry.error,
      details: retry.details || retry.raw || null,
    });
    return { ...retry, attempts };
  }

  return { ...retry, attempts };
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

  return { parsed, rawText };
}

function buildZoomInfoSearchLink(persona, companyName) {
  const parts = [];

  if (persona.name) parts.push(`"${persona.name.trim()}"`);
  if (persona.designation) parts.push(`"${persona.designation.trim()}"`);
  if (persona.department) parts.push(`"${persona.department.trim()}"`);
  if (companyName) parts.push(`"${companyName.trim()}"`);

  const scope = `(site:zoominfo.com OR site:cognism.com OR site:linkedin.com/in)`;

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

  const prompt = `You are a B2B sales intelligence researcher who uses live web search to validate insights.
Identify companies located in the specified geography that would be high-priority targets for purchasing the product described below.
List only companies that plausibly operate in that location and have a clear fit with the product's value.

Product name: ${trimmedProduct}
Target location: ${trimmedLocation || "Not explicitly provided. Infer a sensible geography from context but prioritize the stated location if any."}
${sectorSection}

${docSection}

Guidelines:
- ${mullInstruction.replace(/^- /, "")}
- Use search to confirm the company's presence in the target geography, their core business, and the official website.
- Prefer mid-market or enterprise buyers whose needs align with the product.
- Aggressively remove duplicates or near-duplicates so each company appears only once.
- Provide up to ${TARGET_COMPANY_GOAL} distinct companies, aiming for ${TARGET_COMPANY_GOAL} whenever credible candidates exist. Only return fewer if the market truly lacks more qualified prospects.
- If revenue is unavailable, leave the revenue field as an empty string.
- Keep notes to one concise sentence explaining the fit.

Respond in STRICT JSON with this shape (no Markdown fences, no commentary):
{
  "companies": [
    {
      "name": "Company name",
      "website": "https://official.website",
      "revenue": "Most recent annual revenue or range, or empty string if unknown",
      "notes": "One sentence on why the company is a fit"
    }
  ]
}`;

  const resp = await callGeminiWithRetry(prompt, {
    generationConfig: {
      temperature: 0.4,
      maxOutputTokens: 100000,
    },
  });

  if (resp.error) {
    return { error: resp.error, details: resp.details };
  }

  const { parsed, rawText } = parseModelJsonResponse(resp);

  if (!parsed || !Array.isArray(parsed.companies)) {
    return { error: "Model did not return a structured company list.", details: rawText || null };
  }

  const companies = normalizeTargetCompanies(parsed.companies);
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
  const columnLines = columns
    .map((col, idx) => `${idx + 1}. ${col.header} - ${col.description}`)
    .join("\n");

  const dataset = prepareDatasetForPrompt(entries);
  const datasetJson = JSON.stringify(dataset, null, 2);

  const formatInstruction =
    format === "md"
      ? "Provide a Markdown table string in the field `markdownTable`."
      : "Ensure the JSON rows can be used to build an .xlsx file.";

  return `You are helping prepare research data for export.

Column specifications (respect the header text exactly):
${columnLines}

The research entries are provided as JSON below. Each entry may include nested details such as personas and generated content. Derive values for each column from the available data. If a value is missing, use an empty string. Do not invent data beyond reasonable inferences from the supplied content.

Research entries JSON:
${datasetJson}

Respond in strict JSON with this shape:
{
  "rows": [
    {
      "<Header 1>": "cell value",
      "<Header 2>": "cell value"
    }
  ],
  "notes": "optional short quality notes or considerations",
  "markdownTable": "optional markdown table representing all rows"
}

- The \`rows\` array must contain one object per research entry in the same order they were supplied.
- Each row object must include every header and only those headers.
- Use multiline strings where helpful (they will be preserved).
- ${formatInstruction}
`;
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

async function fetchHeadquartersLocation({ companyName, locationHint }) {
  if (!companyName) {
    return { location: "", metadata: null, rawText: "", error: "Missing company name" };
  }

  const normalizedLocationHint = locationHint ? String(locationHint).trim() : "";
  const focusInstruction = normalizedLocationHint
    ? `Return the primary headquarters (registered office) the company maintains in ${normalizedLocationHint}. Do not fall back to a different country's global HQ; if no headquarters exists in ${normalizedLocationHint}, return an empty hq_location.`
    : `Return the official global headquarters registered for the company.`;

  const prompt = `You are verifying the official headquarters for ${companyName}.
${focusInstruction}
Location hint to respect: ${normalizedLocationHint || "None provided (use global HQ)"}

Requirements:
- If a location hint is provided, interpret "headquarters" as the main registered office for that country/region (e.g., India HQ when the hint is India). Only return an address inside that geography.
- Use Google Search to gather the latest authoritative mentions of the headquarters address, favoring queries that include the location hint.
- Then call the Google Maps tool to pull the place details, coordinates, and formatted address so the answer is grounded in a real map listing.
- If there are conflicting sources, explain why the selected HQ is most accurate.

Respond strictly in JSON:
{
  "hq_location": "City, State/Region, Country"
}
`;

  try {
    const resp = await callGeminiWithRetry(prompt, {
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 9000,
        
      },
      tools: [{ google_search: {} }, { googleMaps: {} }],
    });

    if (resp.error) {
      console.warn("Failed to fetch headquarters location", resp.error, resp.details || resp);
      return { location: "", metadata: null, rawText: "", error: resp.error };
    }

    const { parsed, rawText } = parseModelJsonResponse(resp);

    const structured = parsed && typeof parsed === "object" ? parsed : null;
    if (!structured) {
      return {
        location: "",
        metadata: null,
        rawText,
        error: "Headquarters lookup returned no structured JSON.",
      };
    }

    const resolvedLocation = resolveLocationFromStructured(structured);

    if (!resolvedLocation) {
      return {
        location: "",
        metadata: structured,
        rawText,
        error: "Headquarters lookup did not include a location.",
      };
    }

    return {
      location: resolvedLocation,
      metadata: structured,
      rawText,
    };
  } catch (err) {
    console.warn("Failed to fetch headquarters location", err);
    return { location: "", metadata: null, rawText: "", error: err?.message || String(err) };
  }
}

function summarizePersonasForTelepitch(personas = []) {
  if (!Array.isArray(personas) || !personas.length) return "";
  return personas
    .map((persona, idx) => {
      const bits = [
        `Persona ${idx + 1}:`,
        `Name: ${persona.name || `Persona ${idx + 1}`}`,
        persona.designation ? `Designation: ${persona.designation}` : null,
        persona.department ? `Department: ${persona.department}` : null,
      ].filter(Boolean);
      return bits.join(" ");
    })
    .join("\n");
}

function buildTelephonicPitchPrompt({ personas, company, location, product, docsText, pitchFromCompany }) {
  const personaSummary = summarizePersonasForTelepitch(personas);
  const pitchingOrg = (pitchFromCompany && pitchFromCompany.trim()) || "your company";
  const prospectLabel = company || "the target company";
  const personaCount = Array.isArray(personas) ? personas.length : 0;
  return `You are a helpful assistant. Generate a concise telephonic sales pitch for the following:
  Perspective:
      - You represent ${pitchingOrg}. You are pitching ${prospectLabel}, who is the prospect.
      - Do not flip the roles: ${pitchingOrg} is the seller, ${prospectLabel} is the buyer.
  Instructions:
      - Research the product, company priorities, and each persona's responsibilities so the caller sounds well informed and relevant.
      - Build a 45-60 second phone script for every persona that starts with a personalized opener, adds one probing question, and ties ${product} to their KPIs.
      - Maintain a confident, consultative tone that feels natural for a live call. Avoid long email-like sentences.
      - Keep the telephonic script very crisp, short and sales friendly.
      - Here's an example of telephonic pitch: "Hi [Name], this is [Your Name] calling from [Your Company]. Am I catching you at a good time for quick minute?
      Great, thank you! I'll keep this really short. I work with companies in the [Target company sector]- helping them [your product feature verbs].
      You might have heard of [Your Product] - it's a solution that helps [Your Product Features]. 
      For a company like yours, where you're [Target Company Pain Point verbs], [Your Product] can [Your Product's Effects and Improvements].
      I'd love to set up a short 20 minute demo with our [Relevant Team from Your Company]. 
      They can walk you through how [Target Company] could use it for [Your Product's Effects and Improvements]. Would that work for you sometime this week?" 

  Prospect company: ${prospectLabel}
  Location: ${location || "N/A"}
  Product: ${product}
  Pitching organization (you): ${pitchingOrg}

  Known personas:
  ${personaSummary || "None provided. Infer from context."}

  Context docs (first 4000 chars each):
  ${docsText || "(no docs provided)"}

  Output JSON in this structure EXACTLY. Return one entry per persona in the same order provided. Include persona_name in every entry so it maps back cleanly. Do not include \`\`\`json markdown wrappers.
  Number of personas to cover: ${personaCount || "0"} (must match telephonic_pitches array length).
  {
    "telephonic_pitches":[
      {
        "persona_name":"",
        "full_pitch":""
      }
    ]
  }
  Example with two personas:
  {
    "telephonic_pitches":[
      {"persona_name":"<Persona 1 Name>","full_pitch":""},
      {"persona_name":"<Persona 2 Name>","full_pitch":""}
    ]
  }

  Keep each section crisp and action-oriented.`;
}

function buildPersonaEmailsPrompt({ personas, company, location, product, docsText, pitchFromCompany }) {
  const personaSummary = summarizePersonasForTelepitch(personas);
  const pitchingOrg = (pitchFromCompany && pitchFromCompany.trim()) || "your company";
  const prospectLabel = company || "the target company";
  return `You are a helpful assistant. Generate outbound email drafts for each persona.
Prospect company: ${prospectLabel}
Pitching organization (you): ${pitchingOrg}
Location: ${location || "N/A"}
Product: ${product}

Known personas:
${personaSummary || "None provided. Infer from context."}

Context docs (first 4000 chars each):
${docsText || "(no docs provided)"}

Output JSON exactly. Create one entry per persona in the same order provided:
{
  "persona_emails": [
    {"persona_name": "", "subject": "", "body": ""}
  ]
}

Rules:
- ${prospectLabel} is the prospect. Every email must be written from ${pitchingOrg}'s perspective pitching ${prospectLabel} on ${product}. Never reverse these roles.
- Create a separate subject and body for every persona listed; ensure persona_name matches the persona you are writing for.
- Emails must include a crisp subject and concise, sales-forward body tailored to that persona. Make the email very short, crisp and visually appealing with numbers and statistics from the documents uploaded. Here's an example of an email template:
    Subject: PVR INOX: Guaranteed 90% Faster Content Delivery with IBM Aspera.

      Dear [Name],

      PVR INOX's scale (over 1,700 screens) requires guaranteed, instantaneous content flow. Current transfer methods are slow, unreliable, and expensive.

      IBM Aspera changes the math:

      Speed: A 100GB DCP asset transfers in under 30 minutes, down from the standard 4-8 hours. This is a 90% time savings.

      Efficiency: We guarantee 95%+ network utilization, maximizing the ROI on your current bandwidth investment.

      Risk: Near-zero transfer failure, eliminating costly re-sends and critical release delays.

      Precedent: Major global studio 'X' leveraged Aspera to cut their distribution window by 50%.

      We eliminate content bottlenecks, securing your revenue and saving substantial operational costs.

      Are you available for a sharp 10-minute ROI discussion this week?

      Sincerely,

      [Your Name] [Your Title]`;
}

function buildEmailRevisionPrompt({ persona = {}, company, location, product, baseEmail = {}, instructions = "", pitchingOrg }) {
  const personaName = persona.name || persona.personaName || persona.persona_name || "the persona";
  const designation = persona.designation || persona.personaDesignation || persona.persona_designation || "";
  const department = persona.department || persona.personaDepartment || persona.persona_department || "";
  const subject = (baseEmail && baseEmail.subject) || "";
  const body = (baseEmail && baseEmail.body) || "";
  const hasExisting = subject || body;

  const personaLine = [
    `Persona: ${personaName}`,
    designation ? `Title: ${designation}` : null,
    department ? `Department: ${department}` : null,
  ]
    .filter(Boolean)
    .join(" | ");

  return `You are a helpful sales copywriter. Revise the outbound email for the given persona.
Prospect company: ${company || "N/A"}
Location: ${location || "N/A"}
Product: ${product || "N/A"}
Pitching organization (you): ${pitchingOrg || "your company"}
${personaLine}

${hasExisting ? "Existing email draft:" : "No existing draft. Create a fresh email:"}
Subject: ${subject || "(none)"}
Body:
${body || "(none)"}

User instructions to apply:
${instructions || "(none)"}

Rules:
- Keep it concise, sales-forward, and tailored to the persona.
- Keep the subject crisp; keep the body short and skimmable.
- Preserve correct roles: you represent ${pitchingOrg || "your company"}, pitching ${company || "the target company"} on ${product || "the product"}.

Return JSON ONLY:
{"subject": "", "body": ""}`;
}

function buildPitchRevisionPrompt({ persona = {}, company, location, product, basePitch = {}, instructions = "", pitchingOrg }) {
  const personaName = persona.name || persona.personaName || persona.persona_name || "the persona";
  const designation = persona.designation || persona.personaDesignation || persona.persona_designation || "";
  const department = persona.department || persona.personaDepartment || persona.persona_department || "";
  const hasExisting =
    basePitch.callGoal ||
    basePitch.call_goal ||
    basePitch.opener ||
    basePitch.discoveryQuestion ||
    basePitch.discovery_question ||
    basePitch.script ||
    basePitch.full_pitch;

  const personaLine = [
    `Persona: ${personaName}`,
    designation ? `Title: ${designation}` : null,
    department ? `Department: ${department}` : null,
  ]
    .filter(Boolean)
    .join(" | ");

  return `You are a helpful sales assistant. Revise the telephonic pitch for this persona.
Prospect company: ${company || "N/A"}
Location: ${location || "N/A"}
Product: ${product || "N/A"}
Pitching organization (you): ${pitchingOrg || "your company"}
${personaLine}

${hasExisting ? "Existing pitch sections (rewrite them):" : "No existing pitch. Draft a fresh one:"}
${JSON.stringify(basePitch || {}, null, 2)}

Instructions:
- Apply the user instructions (if any): ${instructions || "(none)"}.
- 45-60 seconds, confident and consultative.
- Personalized opener, one probing question, tie ${product || "the product"} to their KPIs.
- Keep it crisp and action-oriented; avoid email-like walls of text.
- Preserve correct roles: you represent ${pitchingOrg || "your company"}, pitching ${company || "the target company"}.
- Do not split the response into labeled sections (call goal, opener, CTA, etc.); keep everything woven into one tight script.

Return JSON ONLY:
{
  "script": ""
}`;
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
  const prompt = buildTelephonicPitchPrompt({
    personas,
    company,
    location,
    product,
    docsText,
    pitchFromCompany: pitchingCompany,
  });
  const attempts = [];

  const runAttempt = async (label) => {
    try {
      const generationConfig = {
        temperature: 0.2,
        maxOutputTokens: 4096,
      };
      const teleResp = await callGeminiWithRetry(prompt, {
        generationConfig,
        tools: [{ google_search: {} }],
      });

      if (teleResp.error) {
        attempts.push({
          label,
          error: teleResp.error,
          details: summarizeTelephonicDebugInfo(teleResp.details),
        });
        return null;
      }

      const parsed = extractTelephonicPitchResponse(teleResp, personas);
      if (parsed.pitches.length) {
        return parsed.pitches;
      }
      attempts.push({
        label,
        error: parsed.error || "Model response missing telephonic pitches.",
        details: summarizeTelephonicDebugInfo(parsed.rawText || teleResp.text || teleResp.raw),
      });
      return null;
    } catch (err) {
      attempts.push({
        label,
        error: err?.message || String(err),
      });
      return null;
    }
  };

  const structuredResult = await runAttempt("json-output");
  if (structuredResult && structuredResult.length) {
    return { pitches: structuredResult, attempts };
  }

  const fallbackResult = await runAttempt("text-output");
  if (fallbackResult && fallbackResult.length) {
    return { pitches: fallbackResult, attempts };
  }

  const lastError =
    attempts.length && attempts[attempts.length - 1].error
      ? attempts[attempts.length - 1].error
      : "Unable to generate telephonic pitches.";

  return { pitches: [], error: lastError, attempts };
}

async function generatePersonaEmails({ personas, company, location, product, docsText, pitchFromCompany }) {
  const pitchingCompany = (pitchFromCompany && pitchFromCompany.trim()) || (await loadPitchingCompany());
  const prompt = buildPersonaEmailsPrompt({
    personas,
    company,
    location,
    product,
    docsText,
    pitchFromCompany: pitchingCompany,
  });

  const resp = await callGeminiWithRetry(prompt, {
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 12000,
    },
    tools: [{ google_search: {} }],
  });

  if (resp.error) {
    return { personaEmails: [], error: resp.error, rawText: resp.details || resp.text || "" };
  }

  const { parsed, rawText } = parseModelJsonResponse(resp);
  if (!parsed) {
    return { personaEmails: [], error: "Model did not return valid JSON.", rawText };
  }

  const personaEmails = normalizePersonaEmails(personas, parsed.persona_emails || parsed.personaEmails || []);
  if (!personaEmails.length) {
    return { personaEmails: [], error: "Model response missing persona_emails array.", rawText };
  }

  return { personaEmails, rawText };
}

function normalizePersonas(rawPersonas, companyName) {
  if (!Array.isArray(rawPersonas)) return [];
  return rawPersonas.map((p = {}) => ({
    name: p.name || "",
    designation: p.designation || "",
    department: p.department || "",
    zoominfo_link:
      p.zoominfo_link || p.zoomInfo || p.zoominfo || p.zoom || buildZoomInfoSearchLink(p, companyName || ""),
  }));
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

    const resp = await callGeminiWithRetry(prompt, {
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 12000,
      },
    });

    if (resp.error) {
      return { error: resp.error, details: resp.details };
    }

    const { parsed, rawText } = parseModelJsonResponse(resp);
    if (!parsed) return { error: "Model did not return valid JSON.", rawText };

    const draft = {
      personaName: persona?.name || persona?.personaName || persona?.persona_name || "",
      personaDesignation: persona?.designation || persona?.personaDesignation || persona?.persona_designation || "",
      personaDepartment: persona?.department || persona?.personaDepartment || persona?.persona_department || "",
      subject: (parsed.subject && String(parsed.subject)) || "",
      body: (parsed.body && String(parsed.body)) || "",
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

    const resp = await callGeminiWithRetry(prompt, {
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 12000,
      },
    });

    if (resp.error) {
      return { error: resp.error, details: resp.details };
    }

    const { parsed, rawText } = parseModelJsonResponse(resp);
    if (!parsed) return { error: "Model did not return valid JSON.", rawText };

    const normalized = normalizeTelephonicPitchEntry(parsed, 0, [persona || {}]);
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

async function generateBriefOverview({ company, location, product, docsText }) {
  const prompt = `You are a helpful assistant. Research the target company and return concise commercial context.
Company: ${company}
Location: ${location || "N/A"}
Product: ${product}

Context docs (first 4000 chars each):
${docsText || "(no docs provided)"}

Provide JSON with: company_name, revenue_estimate (realistic string), industry_sector, and top_5_news (array of {"title","summary"}).
Do not include markdown fences.`;

  const resp = await callGeminiWithRetry(prompt, {
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 6000,
    },
    tools: [{ google_search: {} }],
  });

  if (resp.error) {
    return { error: resp.error + (resp.details ? " Details: " + JSON.stringify(resp.details) : ""), attempts: resp.details || [] };
  }

  const { parsed, rawText } = parseModelJsonResponse(resp);
  if (!parsed) {
    return { error: "Model did not return valid JSON.", rawText };
  }

  const topNews =
    (Array.isArray(parsed.top_5_news) && parsed.top_5_news) ||
    (Array.isArray(parsed.top_news) && parsed.top_news) ||
    [];

  const overview = {
    company_name: parsed.company_name || company || "",
    revenue_estimate: parsed.revenue_estimate || parsed.revenue || "",
    industry_sector: parsed.industry_sector || parsed.industrySector || parsed.industry || "",
    top_5_news: topNews,
  };

  return { overview, rawText };
}

async function generatePersonaBrief({ company, location, product, docsText, runId }) {
  const pitchFromCompany = await loadPitchingCompany();
  const pitchingOrg = pitchFromCompany || "your company";
  const prospectLabel = company || "the target company";
  const prompt = `You are a helpful assistant. Generate buying personas for the company and outreach-ready messaging.
Prospect company: ${prospectLabel}
Pitching organization (you): ${pitchingOrg}
Location: ${location || "N/A"}
Product: ${product}

Context docs (first 4000 chars each):
${docsText || "(no docs provided)"}

Output JSON exactly:
{
  "company_name": "",
  "key_personas": [
    {
      "name": "",
      "designation": "",
      "department": "",
      "zoominfo_link": ""
    }
  ],
  "telephonic_pitches": [{"full_pitch":""}]
  ,
  "persona_emails": [
    {"persona_name": "", "subject": "", "body": ""}
  ]
}

Rules:
- ${prospectLabel} is the prospect. Every email and telephonic script must be written from ${pitchingOrg}'s perspective pitching ${prospectLabel} on ${product}. Never reverse these roles.
- Only return personas relevant to ${product} purchase decisions.
- Emails must include a crisp subject and concise, sales-forward body tailored to that persona. Make the email very short, crisp and visually appealing with numbers and statistics from the documents uploaded. Here's an example of an email template:
    Subject: PVR INOX: Guaranteed 90% Faster Content Delivery with IBM Aspera.

      Dear [Name],

      PVR INOX's scale (over 1,700 screens) requires guaranteed, instantaneous content flow. Current transfer methods are slow, unreliable, and expensive.

      IBM Aspera changes the math:

      Speed: A 100GB DCP asset transfers in under 30 minutes, down from the standard 4-8 hours. This is a 90% time savings.

      Efficiency: We guarantee 95%+ network utilization, maximizing the ROI on your current bandwidth investment.

      Risk: Near-zero transfer failure, eliminating costly re-sends and critical release delays.

      Precedent: Major global studio 'X' leveraged Aspera to cut their distribution window by 50%.

      We eliminate content bottlenecks, securing your revenue and saving substantial operational costs.

      Are you available for a sharp 10-minute ROI discussion this week?

      Sincerely,

      [Your Name] [Your Title]

- Telephonic pitches should be 45-60 seconds. Here's an example of a telephonic pitch: "Hi [Name], this is [Your Name] calling from [Your Company]. Am I catching you at a good time for quick minute?
      Great, thank you! I'll keep this really short. I work with companies in the [Target company sector]- helping them [your product feature verbs].
      You might have heard of [Your Product] - it's a solution that helps [Your Product Features]. 
      For a company like yours, where you're [Target Company Pain Point verbs], [Your Product] can [Your Product's Effects and Improvements].
      I'd love to set up a short 20 minute demo with our [Relevant Team from Your Company]. 
      They can walk you through how [Target Company] could use it for [Your Product's Effects and Improvements]. Would that work for you sometime this week?" 
- Include a ZoomInfo or LinkedIn style Google search link for each persona (omit "google search:" text).`;

  const resp = await callGeminiWithRetry(prompt, {
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 12000,
    },
    tools: [{ google_search: {} }],
  });

  if (resp.error) {
    return { error: resp.error + (resp.details ? " Details: " + JSON.stringify(resp.details) : ""), attempts: resp.details || [] };
  }

  const { parsed, rawText } = parseModelJsonResponse(resp);
  if (!parsed) {
    return { error: "Model did not return valid JSON.", rawText };
  }

  const rawPersonas = Array.isArray(parsed.key_personas) ? parsed.key_personas : [];
  const personas = normalizePersonas(rawPersonas, parsed.company_name || company);
  const personaEmailsFromPersonaBrief = normalizePersonaEmails(
    rawPersonas,
    parsed.persona_emails || parsed.personaEmails || []
  );
  const telephonicFromPersonaBrief = normalizeTelephonicPitches(parsed, personas);
  const initialEmail = derivePrimaryEmail(parsed, personaEmailsFromPersonaBrief);

  if (runId) {
    emitBriefPartialUpdate(runId, {
      personas,
      personaEmails: personaEmailsFromPersonaBrief,
      telephonicPitches: telephonicFromPersonaBrief.telephonicPitches,
      telephonicPitchError: telephonicFromPersonaBrief.telephonicPitches.length
        ? telephonicFromPersonaBrief.telephonicPitchError
        : "",
      telephonicPitchAttempts: telephonicFromPersonaBrief.telephonicPitchAttempts,
      email: initialEmail,
    });
  }

  const emailPromise = generatePersonaEmails({
    personas,
    company,
    location,
    product,
    docsText,
    pitchFromCompany: pitchingOrg,
  });
  const telephonicPromise = generateTelephonicPitchScripts({
    personas,
    company,
    location,
    product,
    docsText,
    pitchFromCompany: pitchingOrg,
  });

  if (runId) {
    emailPromise
      .then((res) => {
        const personaEmails =
          (res?.personaEmails && res.personaEmails.length ? res.personaEmails : personaEmailsFromPersonaBrief) || [];
        const email = derivePrimaryEmail(parsed, personaEmails);
        emitBriefPartialUpdate(runId, {
          personaEmails,
          email,
        });
      })
      .catch(() => {
        emitBriefPartialUpdate(runId, {
          personaEmails: personaEmailsFromPersonaBrief,
          email: initialEmail,
        });
      });

    telephonicPromise
      .then((res) => {
        const telephonicPitches =
          (res?.pitches && res.pitches.length ? res.pitches : telephonicFromPersonaBrief.telephonicPitches) || [];
        const telephonicPitchError =
          res?.error ||
          (!telephonicPitches.length ? telephonicFromPersonaBrief.telephonicPitchError : "") ||
          "";
        const telephonicPitchAttempts = res?.attempts || telephonicFromPersonaBrief.telephonicPitchAttempts || [];
        emitBriefPartialUpdate(runId, {
          telephonicPitches,
          telephonicPitchError,
          telephonicPitchAttempts,
        });
      })
      .catch((err) => {
        const telephonicPitches = telephonicFromPersonaBrief.telephonicPitches || [];
        const fallbackError = telephonicFromPersonaBrief.telephonicPitchError || "";
        const errorMsg = err?.message || (err ? String(err) : "") || fallbackError;
        emitBriefPartialUpdate(runId, {
          telephonicPitches,
          telephonicPitchError: errorMsg || (!telephonicPitches.length ? fallbackError : ""),
          telephonicPitchAttempts: telephonicFromPersonaBrief.telephonicPitchAttempts || [],
        });
      });
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

  const personaEmails =
    (resolvedEmailResult.personaEmails && resolvedEmailResult.personaEmails.length
      ? resolvedEmailResult.personaEmails
      : personaEmailsFromPersonaBrief) || [];

  const telephonicPitches =
    (resolvedTelephonicResult.pitches && resolvedTelephonicResult.pitches.length
      ? resolvedTelephonicResult.pitches
      : telephonicFromPersonaBrief.telephonicPitches) || [];

  const telephonicPitchError =
    resolvedTelephonicResult.error ||
    (!telephonicPitches.length ? telephonicFromPersonaBrief.telephonicPitchError : "") ||
    "";

  const telephonicPitchAttempts =
    resolvedTelephonicResult.attempts || telephonicFromPersonaBrief.telephonicPitchAttempts || [];

  const email = derivePrimaryEmail(parsed, personaEmails);

  return {
    personas,
    personaEmails,
    telephonicPitches,
    telephonicPitchError,
    telephonicPitchAttempts,
    email,
    rawText,
  };
}

async function generateBrief({ company, location, product, docs = [], runId }) {
  try {
    let totalSteps = 3;
    let completedSteps = 0;
    emitBriefProgress({
      runId,
      current: completedSteps,
      total: totalSteps,
      label: "Starting brief request",
    });

    const docsText = (docs || []).map(d => {
      const txt = decodeBase64Text(d.content_b64 || d.content || "");
      return `--- ${d.name || "doc"} ---\n${txt.substring(0, 4000)}`;
    }).join("\n\n");

    const stepDone = (label) => {
      completedSteps = Math.min(totalSteps, completedSteps + 1);
      emitBriefProgress({
        runId,
        current: completedSteps,
        total: totalSteps,
        label,
      });
    };

    const hqPromise = fetchHeadquartersLocation({
      companyName: company,
      locationHint: location || "",
    })
      .then((res) => {
        stepDone("Headquarters located");
        emitBriefPartialUpdate(runId, {
          hq_location: res.location || "",
          hq_lookup_error: res.error || "",
          hq_lookup_details: res.metadata || null,
        });
        return res;
      })
      .catch((hqErr) => {
        console.warn("Headquarters lookup failed", hqErr);
        stepDone("Headquarters lookup failed");
        const fallback = { location: "", metadata: null, rawText: "", error: hqErr?.message || String(hqErr) };
        emitBriefPartialUpdate(runId, {
          hq_location: "",
          hq_lookup_error: fallback.error,
          hq_lookup_details: null,
        });
        return fallback;
      });

    const overviewPromise = generateBriefOverview({
      company,
      location,
      product,
      docsText,
    })
      .then((res) => {
        stepDone(res.error ? "Overview generation failed" : "Overview generated");
        if (res.error) {
          emitBriefPartialUpdate(runId, { overview_error: res.error });
        } else if (res.overview) {
          emitBriefPartialUpdate(runId, { overview: res.overview });
        }
        return res;
      })
      .catch((err) => {
        const errorMsg = err?.message || String(err);
        stepDone("Overview generation failed");
        emitBriefPartialUpdate(runId, { overview_error: errorMsg });
        return { error: errorMsg };
      });

    const personaPromise = generatePersonaBrief({
      company,
      location,
      product,
      docsText,
      runId,
    })
      .then((res) => {
        stepDone(res.error ? "Persona generation failed" : "Persona outreach generated");
        if (res.error) {
          emitBriefPartialUpdate(runId, { persona_error: res.error });
        } else {
          emitBriefPartialUpdate(runId, {
            personas: res.personas,
            personaEmails: res.personaEmails,
            telephonicPitches: res.telephonicPitches,
            telephonicPitchError: res.telephonicPitchError,
            telephonicPitchAttempts: res.telephonicPitchAttempts,
            email: res.email,
          });
        }
        return res;
      })
      .catch((err) => {
        const errorMsg = err?.message || String(err);
        stepDone("Persona generation failed");
        emitBriefPartialUpdate(runId, { persona_error: errorMsg });
        return {
          error: errorMsg,
          personas: [],
          personaEmails: [],
          telephonicPitches: [],
          telephonicPitchError: errorMsg,
          telephonicPitchAttempts: [],
          email: { subject: "", body: "" },
        };
      });

    const [hqLookupResult, overviewResult, personaResult] = await Promise.all([
      hqPromise,
      overviewPromise,
      personaPromise,
    ]);

    completedSteps = Math.max(completedSteps, totalSteps);
    emitBriefProgress({
      runId,
      current: totalSteps,
      total: totalSteps,
      label: "Brief ready",
    });

    const overview = overviewResult?.overview || {};
    const topNews = Array.isArray(overview.top_5_news) ? overview.top_5_news : [];
    const industrySector = overview.industry_sector || overview.industrySector || overview.industry || "";
    const revenueEstimate = overview.revenue_estimate || "";

    const resolvedHqLocation = hqLookupResult?.location || overview.hq_location || "";
    const hqErrorMessage = hqLookupResult?.error ? `HQ lookup failed: ${hqLookupResult.error}` : "";
    const displayedHq = resolvedHqLocation || hqErrorMessage || "Not found";
    const displayedIndustry = industrySector || "Not found";

    const personas = Array.isArray(personaResult?.personas) ? personaResult.personas : [];
    const personaEmails = Array.isArray(personaResult?.personaEmails) ? personaResult.personaEmails : [];
    const telephonicPitches = Array.isArray(personaResult?.telephonicPitches)
      ? personaResult.telephonicPitches
      : [];
    const telephonicPitchError = personaResult?.telephonicPitchError || "";
    const telephonicPitchAttempts = Array.isArray(personaResult?.telephonicPitchAttempts)
      ? personaResult.telephonicPitchAttempts
      : [];
    const emailObj = personaResult?.email || { subject: "", body: "" };
    const personaEmailVersions = personaEmails.map((draft) => ({ versions: [draft], activeIndex: 0 }));
    const telephonicPitchVersions = telephonicPitches.map((draft) => ({ versions: [draft], activeIndex: 0 }));

    const brief_html = buildBriefHtmlFromOverview({
      companyName: overview.company_name || company,
      hqDisplay: displayedHq,
      revenue: revenueEstimate,
      industry: displayedIndustry,
      topNews,
    });

    const overviewError = overviewResult?.error || "";
    const personaError = personaResult?.error || "";
    const combinedError = overviewError && personaError ? `${overviewError}; ${personaError}` : "";

    return {
      brief_html,
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
      hq_lookup_details: hqLookupResult?.metadata || null,
      hq_lookup_error: hqLookupResult?.error || "",
      hq_lookup_raw: hqLookupResult?.rawText || "",
      raw_overview: overviewResult?.rawText || "",
      raw_personas: personaResult?.rawText || "",
      overview_error: overviewError,
      persona_error: personaError,
      error: combinedError || undefined,
    };
  } catch (err) {
    return { error: String(err) };
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
          const payload = { company: req.company, location: req.location, product: req.product, docs: req.docs || [], runId: req.runId };
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
          const llmResult = await callGeminiWithRetry(prompt, {
            generationConfig: {
              temperature: 0.1,
              maxOutputTokens: 4096,
              responseMimeType: "application/json",
            },
          });

          if (llmResult.error) {
            sendResponse({ error: llmResult.error, details: llmResult.details });
            return;
          }

          const { parsed, rawText } = parseModelJsonResponse(llmResult);

          if (!parsed || !Array.isArray(parsed.rows)) {
            sendResponse({ error: "Model did not return structured rows for export.", details: rawText || null });
            return;
          }

          const normalizedRows = parsed.rows.map((row) => ensureRowValues(row, columns));
          const headers = columns.map((col) => col.header);

          let markdownTable = parsed.markdownTable || parsed.markdown || parsed.table;
          if (format === "md" && !markdownTable) {
            markdownTable = generateMarkdownFromRows(headers, normalizedRows);
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
            if (parsed.notes) {
              exportLines.push("");
              exportLines.push(`> ${parsed.notes}`);
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
            notes: parsed.notes || "",
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

})(); // end backgroundWebScope

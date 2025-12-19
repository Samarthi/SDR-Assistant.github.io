const assert = require("assert");
const path = require("path");

// Minimal globals so background-web.js can load in Node.
global.atob =
  global.atob ||
  function atobPolyfill(str) {
    return Buffer.from(str, "base64").toString("binary");
  };

const noop = () => {};
const chromeStub = {
  storage: { local: { get: async () => ({}), set: async () => ({}) } },
  runtime: {
    sendMessage: noop,
    onMessage: { addListener: noop },
    getURL: (p) => p,
    onInstalled: { addListener: noop },
    onStartup: { addListener: noop },
  },
  action: { onClicked: { addListener: noop } },
  tabs: { create: noop },
  alarms: { create: noop, clear: noop, onAlarm: { addListener: noop } },
};

global.chrome = chromeStub;

const {
  callGroqDirect,
  callGroqWithRetry,
  callGeminiDirect,
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
} = require(path.join(__dirname, "..", "background-web.js"));

const tests = [];
function test(name, fn) {
  tests.push({ name, fn });
}

function makeResponse({ ok = true, status = 200, body }) {
  return {
    ok,
    status,
    text: async () => body,
  };
}

function mockFetchSequence(sequence) {
  let idx = 0;
  global.fetch = async (url, init) => {
    const entry = sequence[Math.min(idx, sequence.length - 1)];
    idx += 1;
    return typeof entry === "function" ? entry(url, init, idx - 1) : entry;
  };
}

test("buildGroqMessages merges system, context, contents, and prompt", () => {
  const messages = buildGroqMessages({
    promptText: "fallback user prompt",
    contents: [
      { role: "system", parts: [{ text: "secondary sys" }] },
      { role: "user", parts: [{ text: "hello user" }] },
    ],
    systemInstruction: "primary system",
    contextText: "added context",
  });

  assert.strictEqual(messages[0].role, "system");
  assert.ok(messages[0].content.includes("primary system"));
  assert.ok(messages[0].content.includes("added context"));
  assert.deepStrictEqual(
    messages.filter((m) => m.role === "user").map((m) => m.content),
    ["hello user", "fallback user prompt"]
  );
});

test("translateToolsToGroq maps to single browser_search tool", () => {
  const { tools, choice } = translateToolsToGroq([{ google_search: {} }, { googleMaps: {} }], false);
  assert.strictEqual(choice, "auto");
  assert.ok(Array.isArray(tools));
  assert.strictEqual(tools.filter((t) => t.type === "browser_search").length, 1);
});

test("parseModelJsonResponse handles plain JSON text", () => {
  const { parsed, rawText } = parseModelJsonResponse({ text: '{"foo":123}' });
  assert.deepStrictEqual(parsed, { foo: 123 });
  assert.strictEqual(rawText, '{"foo":123}');
});

test("callGroqDirect sends tools and context", async () => {
  let capturedBody = null;
  mockFetchSequence([
    (url, init) => {
      capturedBody = JSON.parse(init.body);
      return makeResponse({
        body: JSON.stringify({ choices: [{ message: { content: "done" } }] }),
      });
    },
  ]);

  const resp = await callGroqDirect(
    "prompt",
    {
      systemInstruction: "system line",
      context: "context line",
      contents: [{ role: "user", parts: [{ text: "content user" }] }],
      tools: [{ google_search: {} }],
    },
    { groqKey: "test-key", model: "openai/gpt-oss-120b" }
  );

  assert.ok(resp.ok);
  assert.strictEqual(resp.text, "done");
  assert.ok(capturedBody);
  assert.ok(Array.isArray(capturedBody.tools));
  assert.strictEqual(capturedBody.tools.filter((t) => t.type === "browser_search").length, 1);
  const systemMsg = capturedBody.messages.find((m) => m.role === "system");
  assert.ok(systemMsg && systemMsg.content.includes("system line"));
  assert.ok(systemMsg && systemMsg.content.includes("context line"));
});

test("callGroqDirect surfaces tool-call-only responses as errors", async () => {
  mockFetchSequence([
    makeResponse({
      body: JSON.stringify({
        choices: [{ message: { content: null, tool_calls: [{ id: "tool-1" }] } }],
      }),
    }),
  ]);

  const resp = await callGroqDirect("prompt", {}, { groqKey: "test-key", model: "openai/gpt-oss-120b" });
  assert.ok(resp.error);
  assert.ok(String(resp.error).includes("tool calls"));
});

test("callGroqWithRetry falls back after 429", async () => {
  mockFetchSequence([
    makeResponse({
      ok: false,
      status: 429,
      body: JSON.stringify({ error: { message: "rate" } }),
    }),
    makeResponse({
      body: JSON.stringify({ choices: [{ message: { content: "recovered" } }] }),
    }),
  ]);

  const resp = await callGroqWithRetry(
    "retry prompt",
    { model: "openai/gpt-oss-120b" },
    { groqKey: "test-key", model: "openai/gpt-oss-120b" }
  );

  assert.ok(resp.ok);
  assert.strictEqual(resp.text, "recovered");
});

test("callGeminiDirect routes through Groq with Groq model", async () => {
  let capturedModel = null;
  mockFetchSequence([
    (url, init) => {
      const body = JSON.parse(init.body);
      capturedModel = body.model;
      return makeResponse({
        body: JSON.stringify({ choices: [{ message: { content: "gemini via groq" } }] }),
      });
    },
  ]);

  const resp = await callGeminiDirect("gemini prompt", { model: "gemini-2.0-pro" }, { groqKey: "test-key" });

  assert.ok(resp.ok);
  assert.ok(capturedModel.startsWith("openai/"));
});

test("callLlmWithRetry uses stored Groq settings", async () => {
  chromeStub.storage.local.get = async () => ({
    llmProvider: "groq",
    llmModel: "openai/gpt-oss-120b",
    geminiKey: "",
    groqKey: "stored-key",
  });

  mockFetchSequence([
    makeResponse({
      body: JSON.stringify({ choices: [{ message: { content: "llm done" } }] }),
    }),
  ]);

  const resp = await callLlmWithRetry("llm prompt");
  assert.ok(resp.ok);
  assert.strictEqual(resp.text, "llm done");
});

test("parsePersonasFromMarkdown extracts persona list", () => {
  const personas = parsePersonasFromMarkdown(`- Name=Alice Doe; Title=VP Marketing; Department=Marketing; SearchLink=https://example.com
- Name=Bob Roe; Title=CTO; Department=Engineering; SearchLink=https://example.org; LinkedInKeywords=CTO ACME`, "ACME");
  assert.strictEqual(personas.length, 2);
  assert.strictEqual(personas[0].name, "Alice Doe");
  assert.strictEqual(personas[0].designation, "VP Marketing");
  assert.strictEqual(personas[1].department, "Engineering");
  assert.strictEqual(personas[1].linkedin_keywords, "CTO ACME");
  assert.ok(personas[1].linkedin_search_url.includes("linkedin.com/search/results/people/"));
  assert.ok(personas[1].linkedin_search_url.includes("keywords=CTO%20ACME"));
});

test("parsePersonaEmailMarkdown extracts subject and body", () => {
  const parsed = parsePersonaEmailMarkdown(
    `Persona=Alice Doe; Title=VP Marketing; Department=Marketing
Subject: Hello Alice
Body:
Line one
Line two`,
    { name: "Fallback" }
  );
  assert.strictEqual(parsed.personaName, "Alice Doe");
  assert.strictEqual(parsed.subject, "Hello Alice");
  assert.ok(parsed.body.includes("Line two"));
});

test("parseTelephonicPitchMarkdown extracts script", () => {
  const parsed = parseTelephonicPitchMarkdown(
    `Persona=Bob Roe; Title=CTO; Department=Engineering
Telephonic Pitch:
Open with context.
Close with CTA.`,
    { name: "Fallback" }
  );
  assert.strictEqual(parsed.personaName, "Bob Roe");
  assert.ok(parsed.script.includes("CTA"));
});

test("toLinkedInPeopleSearchUrl builds encoded search URL", () => {
  const url = toLinkedInPeopleSearchUrl("CISO Adani");
  assert.ok(url.startsWith("https://www.linkedin.com/search/results/people/"));
  assert.ok(url.includes("CISO%20Adani"));
});

test("buildLinkedInKeywordFallback includes company and product context", () => {
  const keywords = buildLinkedInKeywordFallback(
    { name: "Zed", designation: "Head of IT" },
    "ACME Corp",
    "Cloud Security"
  );
  assert.ok(keywords.includes("Head of IT"));
  assert.ok(keywords.includes("ACME Corp"));
  assert.ok(keywords.includes("Cloud Security"));
});

test("parseRevenueSectorMarkdown reads labeled lines", () => {
  const parsed = parseRevenueSectorMarkdown(
    `Company: Sample Corp
Revenue: $10M
Sector: SaaS`
  );
  assert.strictEqual(parsed.company_name, "Sample Corp");
  assert.strictEqual(parsed.revenue_estimate, "$10M");
  assert.strictEqual(parsed.industry_sector, "SaaS");
});

test("parseHqLocationMarkdown reads HQ line", () => {
  const location = parseHqLocationMarkdown("HQ Location: Mumbai, Maharashtra, India");
  assert.strictEqual(location, "Mumbai, Maharashtra, India");
});

test("parseTopNewsMarkdown extracts headlines", () => {
  const items = parseTopNewsMarkdown(`- Title=OpenAI launches new model; Summary=Details about the launch
- Title=Partnership announced; Summary=OpenAI partners with X`);
  assert.strictEqual(items.length, 2);
  assert.strictEqual(items[0].title, "OpenAI launches new model");
  assert.ok(items[0].summary.includes("launch"));
});

test("parseTargetCompaniesMarkdown extracts companies", () => {
  const companies = parseTargetCompaniesMarkdown(`- Name=Acme Corp; Website=https://acme.com; Revenue=$50M; Notes=Strong fit
- Name=Beta Inc; Website=https://beta.com; Notes=Good prospect`);
  assert.strictEqual(companies.length, 2);
  assert.strictEqual(companies[0].name, "Acme Corp");
  assert.strictEqual(companies[1].website, "https://beta.com");
});

test("parseMarkdownTable parses simple table", () => {
  const md = `| A | B |
| --- | --- |
| a1 | b1 |
| a2 | b2 |`;
  const { headers, rows } = parseMarkdownTable(md);
  assert.deepStrictEqual(headers, ["A", "B"]);
  assert.strictEqual(rows.length, 2);
  assert.strictEqual(rows[1].B, "b2");
});

(async () => {
  let passed = 0;
  for (const { name, fn } of tests) {
    try {
      await fn();
      passed += 1;
    } catch (err) {
      console.error(`❌ ${name}:`, err);
      process.exit(1);
    }
  }
  console.log(`✅ ${passed} tests passed`);
})();

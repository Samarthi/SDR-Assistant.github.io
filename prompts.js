(function initPrompts(globalScope) {
  function buildTargetsPrompt({
    product = "",
    location = "",
    sectorSection = "",
    docSection = "",
    mullInstruction = "",
    targetCompanyGoal = 100,
  }) {
    const trimmedProduct = product || "N/A";
    const trimmedLocation =
      location || "Not explicitly provided. Infer a sensible geography from context but prioritize the stated location if any.";
    const mull = mullInstruction ? mullInstruction.replace(/^- /, "") : "";

    return `You are a B2B sales intelligence researcher who uses live web search to validate insights.
Identify companies located in the specified geography that would be high-priority targets for purchasing the product described below.
List only companies that plausibly operate in that location and have a clear fit with the product's value.

Product name: ${trimmedProduct}
Target location: ${trimmedLocation}
${sectorSection}

${docSection}

Guidelines:
- ${mull}
- Use search to confirm the company's presence in the target geography, their core business, and the official website.
- Prefer mid-market or enterprise buyers whose needs align with the product.
- Aggressively remove duplicates or near-duplicates so each company appears only once.
- Provide up to ${targetCompanyGoal} distinct companies, aiming for ${targetCompanyGoal} whenever credible candidates exist. Only return fewer if the market truly lacks more qualified prospects.
- If revenue is unavailable, leave the revenue field as an empty string.
- Keep notes to one concise sentence explaining the fit.

Return markdown only as bullet points (no code fences). Use this format for each company:
- Name=<Company>; Website=<official URL>; Revenue=<revenue or empty>; Notes=<one-line fit reason>`;
  }

  function buildTelephonicPitchPrompt({ persona = {}, company, location, product, docsText, pitchFromCompany }) {
    const personaLabel =
      persona.designation ||
      persona.personaDesignation ||
      persona.persona_designation ||
      persona.department ||
      "the persona";
    const designation = persona.designation || persona.personaDesignation || persona.persona_designation || "";
    const department = persona.department || persona.personaDepartment || persona.persona_department || "";
    const pitchingOrg = (pitchFromCompany && pitchFromCompany.trim()) || "your company";
    const prospectLabel = company || "the target company";
    return `You are a helpful assistant. Generate a concise telephonic sales pitch for one persona.
  Perspective:
      - You represent ${pitchingOrg}. You are pitching ${prospectLabel}, who is the prospect.
      - Do not flip the roles: ${pitchingOrg} is the seller, ${prospectLabel} is the buyer.
  Instructions:
      - Research the product, company priorities, and each persona's responsibilities so the caller sounds well informed and relevant.
      - Build a 45-60 second phone script for every persona that starts with a personalized opener, adds one probing question, and ties ${product} to their KPIs.
      - Maintain a confident, consultative tone that feels natural for a live call. Avoid long email-like sentences. Use the use cases and numbers from the context documents to build a solid pitch and customize it according to the prospect's context.
      - Keep the telephonic script very crisp, short and sales friendly.
      - Format the brief with sufficient newline characters, punctuations and spaces, like a formal brief would.
      - Include at least 4-5 of the most unique selling points of the product from the context docs.
      - Here's an example of a nice telephonic pitch, do not use this blindly: "Hi [Name], this is [Your Name] calling from [Your Company]. Am I catching you at a good time for quick minute?
      Great, thank you! I'll keep this really short. I work with companies in the [Target company sector]- helping them [your product feature verbs].
      You might have heard of [Your Product] - it's a solution that helps [Your Product Features]. 
      For a company like yours, where you're [Target Company Pain Point verbs], [Your Product] can [Your Product's Effects and Improvements].
      I'd love to set up a short 20 minute demo with our [Relevant Team from Your Company]. 
      They can walk you through how [Target Company] could use it for [Your Product's Effects and Improvements]. Would that work for you sometime this week?" 

  Prospect company: ${prospectLabel}
  Location: ${location || "N/A"}
  Product: ${product}
  Pitching organization (you): ${pitchingOrg}
  Persona: ${personaLabel}${designation ? ` | Title: ${designation}` : ""}${department ? ` | Department: ${department}` : ""}

  Context docs (first 4000 chars each):
  ${docsText || "(no docs provided)"}

  Return markdown only in this exact format (no headings, no code fences):
  Persona=${personaLabel}; Title=${designation || "N/A"}; Department=${department || "N/A"}
  Telephonic Pitch:
  <45-60 second script here>`;
  }

  function buildPersonaEmailsPrompt({ persona = {}, company, location, product, docsText, pitchFromCompany }) {
    const personaLabel =
      persona.designation ||
      persona.personaDesignation ||
      persona.persona_designation ||
      persona.department ||
      "the persona";
    const designation = persona.designation || persona.personaDesignation || persona.persona_designation || "";
    const department = persona.department || persona.personaDepartment || persona.persona_department || "";
    const pitchingOrg = (pitchFromCompany && pitchFromCompany.trim()) || "your company";
    const prospectLabel = company || "the target company";
    return `You are a helpful assistant. Generate outbound email drafts for each persona.
Prospect company: ${prospectLabel}
Pitching organization (you): ${pitchingOrg}
Location: ${location || "N/A"}
 Product: ${product}
 Persona: ${personaLabel}${designation ? ` | Title: ${designation}` : ""}${department ? ` | Department: ${department}` : ""}

Context docs (first 4000 chars each):
${docsText || "(no docs provided)"}

Rules:
- ${prospectLabel} is the prospect. Every email must be written from ${pitchingOrg}'s perspective pitching ${prospectLabel} on ${product}. Never reverse these roles.
- Create a separate subject and body for every persona listed; ensure persona_name matches the persona you are writing for.
- Emails must include a crisp subject and concise, sales-forward body tailored to that persona.  Make the email crisp (with bullets for each point) and visually appealing with numbers and statistics from the documents uploaded. 
- Including at least 4-5 most relevant unique selling points for the product from the context docs.
- Always address the recipient using the placeholder "[First Name]" in the salutation and anywhere you need their name. Do not use their title or any actual name.
- Format the brief with sufficient newline characters, punctuations and spaces, like a formal brief would.
- Here's an example of an email brief. Do not blindly use this. Build your own based on the data from context docs:
    Subject: PVR INOX: Guaranteed 90% Faster Content Delivery with IBM Aspera.

      Dear [First Name],

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

Return markdown only in this exact format (no headings, no code fences):
Persona=${personaLabel}; Title=${designation || "N/A"}; Department=${department || "N/A"}
Subject: <subject line>
Body:
<short email body>`;
  }

  function buildEmailRevisionPrompt({ persona = {}, company, location, product, baseEmail = {}, instructions = "", pitchingOrg }) {
    const personaLabel =
      persona.designation ||
      persona.personaDesignation ||
      persona.persona_designation ||
      persona.department ||
      "the persona";
    const designation = persona.designation || persona.personaDesignation || persona.persona_designation || "";
    const department = persona.department || persona.personaDepartment || persona.persona_department || "";
    const subject = baseEmail.subject || baseEmail.email_subject || "";
    const body = baseEmail.body || baseEmail.email_body || "";
    const prospectLabel = company || "the target company";
    const personaLine = [
      `Persona: ${personaLabel}`,
      designation ? `Title: ${designation}` : null,
      department ? `Department: ${department}` : null,
    ]
      .filter(Boolean)
      .join(" | ");

    const hasExisting = !!(subject || body);

    return `You are a helpful sales copywriter. Revise the outbound email for the given persona.
Prospect company: ${prospectLabel}
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

Return markdown only in this exact format (no headings, no code fences):
Persona=${personaLabel}; Title=${designation || "N/A"}; Department=${department || "N/A"}
Subject: <subject line>
Body:
<short email body>`;
  }

  function buildPitchRevisionPrompt({ persona = {}, company, location, product, basePitch = {}, instructions = "", pitchingOrg }) {
    const personaLabel =
      persona.designation ||
      persona.personaDesignation ||
      persona.persona_designation ||
      persona.department ||
      "the persona";
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
      `Persona: ${personaLabel}`,
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

Return markdown only in this exact format (no headings, no code fences):
Persona=${personaLabel}; Title=${designation || "N/A"}; Department=${department || "N/A"}
Telephonic Pitch:
<45-60 second script>`;
  }

  function buildHqRevenueSectorPrompt({ company, locationHint, product, docsText }) {
    const normalizedLocationHint = locationHint || "None provided (use global HQ)";
    return `You are a helpful assistant. Using live web search, return the official HQ location, realistic revenue, and primary industry sector for the company. The company HQ should be for the location hint provided and not necessarily the global HQ. Keep that in mind when searching.
Company: ${company}
Location hint: ${normalizedLocationHint}
Product: ${product || "N/A"}

Context docs (first 4000 chars each):
${docsText || "(no docs provided)"}

Return markdown only with exactly three lines (no headings, no code fences):
HQ Location: City/Town, State/Region, Country
Revenue: realistic revenue string or "Unknown"
Sector: primary industry sector`;
  }

  function buildRecentNewsPrompt({ company, location, docsText }) {
    return `You are a helpful assistant. Find the five most recent, relevant news headlines about the target company.
Company: ${company}
Location context: ${location || "N/A"}

Context docs (first 4000 chars each):
${docsText || "(no docs provided)"}

Return markdown only. Provide up to 5 bullet items in this format (no code fences):
- Title=<headline>; Summary=<one-sentence summary>`;
  }

  function buildPersonaBriefPrompt({ prospectLabel, pitchingOrg, location, product, docsText, companyName }) {
    return `You are a helpful assistant. Generate buying personas for ${prospectLabel}. Focus only on personas involved in purchasing ${product}.
Pitching organization (you): ${pitchingOrg}
Location: ${location || "N/A"}
Product: ${product}
Context docs (first 4000 chars each):
${docsText || "(no docs provided)"}

Do not include personal names in the personas. Refer to personas only by their title/role (e.g., CISO, VP Engineering).
Include a LinkedIn People search keyword string for each persona that would help find the right titles at ${companyName || "the company"} for ${product}. Do not return a LinkedIn URL - only the keyword string.
Return a simple keyword search string that always includes the corporation name at the front, followed by only the most common title for the persona's position. Do not include title variations, quotes, OR/AND operators, or other Boolean connectors (e.g., CompanyName VP Security).
The SearchLink should be a Google query that combines only the corporation name and the persona's position (no personal names).
Return markdown only. Provide one bullet per persona using this format (no headings, no code fences):
- Title=<Job title>; Department=<Department>; SearchLink=<ZoomInfo/LinkedIn style Google search link>; LinkedInKeywords=<keyword string for LinkedIn People search>`;
  }

  function buildProductContextPrompt({ product }) {
    return `You are a helpful assistant. Use browser search to quickly understand this product and summarize its top use cases.
Product: ${product}

Requirements:
- Rely on live web search to ground the summary; prefer recent, credible sources like the product's website.
- Capture what the product does, who uses it, and 3-5 core use cases or problems it solves.
- Keep it concise and factual.

Return markdown only as a short paragraph or up to 5 bullets (no headings, no code fences).`;
  }

  function buildExportPrompt({ columns = [], dataset = [], format }) {
    const columnLines = columns.map((col, idx) => `${idx + 1}. ${col.header} - ${col.description}`).join("\n");
    const datasetJson = JSON.stringify(dataset, null, 2);
    const formatInstruction =
      format === "md"
        ? "Return a markdown table with headers exactly matching the column headers below. Also include an optional line `Notes: ...` after the table if you want to add a note."
        : "Return a markdown table with headers exactly matching the column headers below. We will parse it into .xlsx.";

    return `You are helping prepare research data for export.

Column specifications (respect the header text exactly):
${columnLines}

The research entries are provided as JSON below. Each entry may include nested details such as personas and generated content. Derive values for each column from the available data. If a value is missing, use an empty string. Do not invent data beyond reasonable inferences from the supplied content.

Research entries JSON:
${datasetJson}

Return markdown only (no code fences) in this structure:
| <Header 1> | <Header 2> | ... |
| --- | --- | ... |
| row 1 col 1 | row 1 col 2 | ... |
| row 2 col 1 | row 2 col 2 | ... |
Notes: <optional short quality notes or considerations>

- Keep the row order aligned to the research entries provided.
- Every header must appear and only those headers.
- Use multiline cells where helpful (they will be preserved).
- ${formatInstruction}`;
  }

  const Prompts = {
    buildTargetsPrompt,
    buildTelephonicPitchPrompt,
    buildPersonaEmailsPrompt,
    buildEmailRevisionPrompt,
    buildPitchRevisionPrompt,
    buildHqRevenueSectorPrompt,
    buildRecentNewsPrompt,
    buildPersonaBriefPrompt,
    buildProductContextPrompt,
    buildExportPrompt,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = Prompts;
  }

  if (globalScope) {
    globalScope.Prompts = Prompts;
  }
})(typeof self !== "undefined" ? self : typeof global !== "undefined" ? global : this);

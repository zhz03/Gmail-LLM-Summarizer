/**
 * Gmail Summary Agent with Gemini (Apps Script) — Robust Classification
 * - JSON-forced LLM for classification
 * - Deterministic rules for Job Sites / Applications / Academic Review
 * - Per-site dynamic labels (e.g., LLM/Jobs/Sites/Glassdoor)
 * Author: Zhaoliang Zheng (zhz03@g.ucla.edu)
 */

const CONFIG = {
  LOOKBACK_DAYS: 1,
  BATCH_SIZE: 20,
  BODY_TRUNCATE: 2000,
  LABEL_ROOT: "LLM",
  GEMINI_MODEL: "gemini-2.5-flash",
  SUMMARY_RECIPIENT: Session.getActiveUser().getEmail(),
  SUMMARY_SUBJECT_PREFIX: "Daily Gmail Summary",
  generationConfig: { temperature: 0.2, topK: 32, topP: 0.9, maxOutputTokens: 4096 },
  VERTEX: { PROJECT_ID: "Gmail Summary", LOCATION: "us-central1", MODEL: "gemini-2.5-flash" },
  TIMEZONE: "America/Los_Angeles",
  DEBUG_LOG_LLM: true // set false to reduce logs
};

// --- Expanded job sites ---
const KNOWN_JOB_SITES = [
  { name: "LinkedIn", domains: ["linkedin.com","email.linkedin.com","notifications.linkedin.com"], keywords: ["linkedin","job alert","application was sent"] },
  { name: "Glassdoor", domains: ["glassdoor.com","email.glassdoor.com","glassdoormail.com"], keywords: ["glassdoor","job alert","jobs in"] },
  { name: "Indeed", domains: ["indeed.com","email.indeed.com"], keywords: ["indeed","job alert"] },
  { name: "ZipRecruiter", domains: ["ziprecruiter.com","mktg.ziprecruiter.com"], keywords: ["ziprecruiter","job alert"] },
  { name: "Monster", domains: ["monster.com","email.monster.com"], keywords: ["monster","job alert"] },
  { name: "Dice", domains: ["dice.com","email.dice.com"], keywords: ["dice","tech jobs","job alert"] },
  { name: "Hired", domains: ["hired.com","mail.hired.com"], keywords: ["hired","interview request"] },
  { name: "Wellfound (AngelList)", domains: ["wellfound.com","angel.co"], keywords: ["wellfound","angellist","startup jobs"] },
  { name: "Greenhouse", domains: ["greenhouse.io","mail.greenhouse.io"], keywords: ["greenhouse","application","interview"] },
  { name: "Lever", domains: ["lever.co","jobs.lever.co","hire.lever.co"], keywords: ["lever","application","interview"] },
  { name: "Workday", domains: ["myworkday.com","workday.com"], keywords: ["workday","candidate home","application"] },
  { name: "SmartRecruiters", domains: ["smartrecruiters.com","mail.smartrecruiters.com"], keywords: ["smartrecruiters","application"] },
  { name: "Ashby", domains: ["ashbyhq.com","jobs.ashbyhq.com"], keywords: ["ashby","application"] },
  { name: "Jobvite", domains: ["jobvite.com","talentcommunity.com"], keywords: ["jobvite","application"] },
  { name: "BambooHR", domains: ["bamboohr.com"], keywords: ["bamboohr","application"] },
  { name: "JazzHR", domains: ["jazzhr.com","app.jazz.co"], keywords: ["jazzhr","application"] },
  { name: "Recruitee", domains: ["recruitee.com","mail.recruitee.com"], keywords: ["recruitee","application"] },
  { name: "iCIMS", domains: ["icims.com","crm.icims.com"], keywords: ["icims","application"] },
  { name: "ADP Recruiting", domains: ["adp.com","recruiting.adp.com"], keywords: ["adp recruiting","application"] },
  { name: "Teamtailor", domains: ["teamtailor.com"], keywords: ["teamtailor","application"] },
  { name: "Jobcase", domains: ["jobcase.com"], keywords: ["jobcase","job alert"] },
  { name: "Seek", domains: ["seek.com","seek.com.au"], keywords: ["seek","job alert"] },
  { name: "StepStone", domains: ["stepstone.com","stepstone.de"], keywords: ["stepstone","job alert"] },
  { name: "Reed", domains: ["reed.co.uk"], keywords: ["reed","job alert"] },
  { name: "Michael Page", domains: ["michaelpage.com","pagepersonnel.com"], keywords: ["michael page","jobs"] },
  { name: "Handshake", domains: ["joinhandshake.com"], keywords: ["handshake","job alert"] },
  { name: "WayUp", domains: ["wayup.com"], keywords: ["wayup","job"] },
  { name: "ClearanceJobs", domains: ["clearancejobs.com"], keywords: ["clearancejobs"] },
  { name: "Levels.fyi", domains: ["levels.fyi","email.levels.fyi"], keywords: ["levels.fyi","job"] }
];

// --- Academic review terms/venues ---
const ACADEMIC_HINTS = {
  domains: [
    "ieee.org","ieeeaccess.ieee.org","manuscriptcentral.com","scholarone.com","editorialmanager.com",
    "springernature.com","elsevier.com","wiley.com","openreview.net","aaai.org","acm.org","neurips.cc","icml.cc","iclr.cc",
    "trb.org","annualmeeting.trb.org","webofscience.com","clarivate.com"
  ],
  keywords: [
    "reviewer","review complete","thank you for reviewing","agreeing to review","manuscript id",
    "reviewer center","submission","revise","decision letter","editorial manager","scholarone",
    "openreview","web of science","trb annual meeting","ieee access"
  ]
};

// ----- Labels -----
const LABELS = {
  ADS: "Ads",
  JOB_SITES: "Jobs/Sites",
  JOB_INTERVIEW: "Jobs/Interview",
  SCHOOL: "School",
  INTERNAL: "Internal",
  EXT_WORK: "External/Work",
  EXT_PERSONAL: "External/Personal",
  EXT_REVIEW: "External/AcademicReview",
  UNKNOWN: "Unknown"
};

// ===================== ENTRYPOINTS =====================

function runDaily() {
  const threads = fetchRecentThreads_(CONFIG.LOOKBACK_DAYS);
  if (!threads.length) {
    Logger.log("No threads found.");
    sendSummaryEmail_("No new emails today.");
    return;
  }

  const labelMap = ensureLabels_();
  const userDomain = (Session.getActiveUser().getEmail() || "").split("@").pop() || "";

  // flatten + deterministic hints
  const emailItems = flattenMessages_(threads).map(it => addHeuristics_(it, userDomain));

  // classify (JSON-forced) with fallback to heuristics
  const classified = classifyEmailsInBatches_(emailItems);

  applyLabels_(classified, labelMap);

  const html = buildDailySummary_(classified);
  const subject = `${CONFIG.SUMMARY_SUBJECT_PREFIX} - ${formatDate_(new Date())}`;
  MailApp.sendEmail(CONFIG.SUMMARY_RECIPIENT, subject, html.replace(/<[^>]+>/g, ""), { htmlBody: html });
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runDaily") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("runDaily")
    .timeBased()
    .atHour(8).nearMinute(30)
    .everyDays(1)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
  Logger.log("Daily trigger installed for ~08:30 LA time.");
}

// ===================== CORE LOGIC =====================

function fetchRecentThreads_(lookbackDays) {
  const query = `newer_than:${lookbackDays}d`;
  return GmailApp.search(query, 0, 500);
}

function flattenMessages_(threads) {
  const items = [];
  threads.forEach(thread => {
    const threadId = thread.getId();
    thread.getMessages().forEach(msg => {
      const from = msg.getFrom() || "";
      items.push({
        threadId,
        messageId: msg.getId(),
        from,
        to: msg.getTo() || "",
        cc: msg.getCc() || "",
        date: msg.getDate() ? msg.getDate().toISOString() : "",
        subject: msg.getSubject() || "",
        snippet: msg.getPlainBody().substring(0, CONFIG.BODY_TRUNCATE),
        fromDomain: extractDomain_(from)
      });
    });
  });
  return items;
}

// ---- Heuristics (deterministic rules) ----
function addHeuristics_(item, userDomain) {
  const subj = (item.subject || "").toLowerCase();
  const body = (item.snippet || "").toLowerCase();
  const dom = (item.fromDomain || "").toLowerCase();

  // Job site detection
  let jobSiteName = null;
  for (const s of KNOWN_JOB_SITES) {
    const hitDomain = (s.domains || []).some(d => dom.indexOf(d.toLowerCase()) !== -1);
    const hitKeyword = (s.keywords || []).some(k => subj.includes(k.toLowerCase()) || body.includes(k.toLowerCase()));
    if (hitDomain || hitKeyword) { jobSiteName = s.name; break; }
  }

  // Application / interview (company or ATS)
  const jobApp = /\b(thank you for applying|your application was sent|application received|interview|schedule|recruiter|hr)\b/i.test(item.subject + " " + item.snippet);

  // Academic review
  const acadDomain = ACADEMIC_HINTS.domains.some(d => dom.indexOf(d.toLowerCase()) !== -1);
  const acadKw = ACADEMIC_HINTS.keywords.some(k => subj.includes(k.toLowerCase()) || body.includes(k.toLowerCase()));
  const academic = acadDomain || acadKw;

  // School (simple)
  const school = /\b(university|college|campus|registrar|course|canvas notifications)\b/i.test(item.subject + " " + item.snippet);

  // Internal vs External
  const internal = userDomain && dom.endsWith(userDomain.toLowerCase());

  // Build heuristic categories
  const heurCats = new Set();
  if (jobSiteName) heurCats.add("Jobs/Sites");
  if (jobApp) heurCats.add("Jobs/Interview");
  if (academic) heurCats.add("External/AcademicReview");
  if (school) heurCats.add("School");
  if (internal) heurCats.add("Internal");

  return { ...item,
    heuristicJobSiteName: jobSiteName || null,
    heuristicCategories: Array.from(heurCats)
  };
}

function classifyEmailsInBatches_(items) {
  const batches = chunk_(items, CONFIG.BATCH_SIZE);
  const out = [];

  batches.forEach(batch => {
    const prompt = buildClassifierPrompt_(batch);
    const text = fetchGeminiJSON_(prompt); // JSON-forced

    let json = safeParseJSON_(text);
    if (!json || !Array.isArray(json.items)) {
      // Fallback to heuristics per item
      batch.forEach(it => {
        const cats = (it.heuristicCategories && it.heuristicCategories.length) ? it.heuristicCategories : ["Unknown"];
        const rec = {
          ...it,
          categories: cats,
          jobSiteName: it.heuristicJobSiteName || null,
          reasons: ["heuristic-fallback"]
        };
        out.push(rec);
      });
      return;
    }

    // Merge LLM + heuristics
    json.items.forEach(rec => {
      const base = batch.find(b => b.messageId === rec.messageId);
      if (!base) return;
      const merged = {
        ...base,
        ...rec
      };
      // if LLM didn't provide site name, use heuristic
      if ((!merged.jobSiteName || merged.jobSiteName === "null") && base.heuristicJobSiteName) {
        merged.jobSiteName = base.heuristicJobSiteName;
      }
      // if LLM gave Unknown but heuristics have categories, merge them
      const llmCats = new Set((merged.categories || []).map(x => String(x)));
      if ((llmCats.size === 0 || (llmCats.size === 1 && llmCats.has("Unknown")))
          && base.heuristicCategories && base.heuristicCategories.length) {
        base.heuristicCategories.forEach(c => llmCats.add(c));
        merged.categories = Array.from(llmCats);
        merged.reasons = (merged.reasons || []).concat(["heuristic-merge"]);
      }
      out.push(merged);
    });
  });

  return out;
}

function applyLabels_(classifiedItems, labelMap) {
  const byThread = new Map();
  for (const it of classifiedItems) {
    if (!byThread.has(it.threadId)) byThread.set(it.threadId, []);
    byThread.get(it.threadId).push(it);
  }

  byThread.forEach((arr, threadId) => {
    const thread = GmailApp.getThreadById(threadId);
    const cats = new Set();
    const siteNames = new Set();

    arr.forEach(it => {
      normalizeCategories_(it).forEach(c => cats.add(c));
      if (it.jobSiteName) siteNames.add(it.jobSiteName);
      else if (it.heuristicJobSiteName) siteNames.add(it.heuristicJobSiteName);
    });

    const labelsToAdd = [];
    if (cats.has("Ads")) labelsToAdd.push(labelMap.ADS);
    if (cats.has("Jobs/Sites")) labelsToAdd.push(labelMap.JOB_SITES);
    if (cats.has("Jobs/Interview")) labelsToAdd.push(labelMap.JOB_INTERVIEW);
    if (cats.has("School")) labelsToAdd.push(labelMap.SCHOOL);
    if (cats.has("Internal")) labelsToAdd.push(labelMap.INTERNAL);
    if (cats.has("External/Work")) labelsToAdd.push(labelMap.EXT_WORK);
    if (cats.has("External/Personal")) labelsToAdd.push(labelMap.EXT_PERSONAL);
    if (cats.has("External/AcademicReview")) labelsToAdd.push(labelMap.EXT_REVIEW);
    if (labelsToAdd.length === 0) labelsToAdd.push(labelMap.UNKNOWN);

    if (cats.has("Jobs/Sites") && siteNames.size > 0) {
      siteNames.forEach(name => {
        const dyn = ensureDynamicJobSiteLabel_(name);
        if (dyn) labelsToAdd.push(dyn);
      });
    }

    labelsToAdd.forEach(l => thread.addLabel(l));
  });
}

function buildDailySummary_(classified) {
  const stats = {
    total: classified.length,
    Ads: 0, JobSites: 0, Interview: 0, School: 0,
    Internal: 0, ExtWork: 0, ExtPersonal: 0, ExtReview: 0, Unknown: 0
  };
  const pick = (c) => classified.filter(x => normalizeCategories_(x).has(c));
  stats.Ads = pick("Ads").length;
  stats.JobSites = pick("Jobs/Sites").length;
  stats.Interview = pick("Jobs/Interview").length;
  stats.School = pick("School").length;
  stats.Internal = pick("Internal").length;
  stats.ExtWork = pick("External/Work").length;
  stats.ExtPersonal = pick("External/Personal").length;
  stats.ExtReview = pick("External/AcademicReview").length;
  stats.Unknown = pick("Unknown").length;

  const siteCount = {};
  classified.forEach(x => {
    const cats = normalizeCategories_(x);
    if (cats.has("Jobs/Sites")) {
      const name = x.jobSiteName || x.heuristicJobSiteName || "UnknownSite";
      siteCount[name] = (siteCount[name] || 0) + 1;
    }
  });
  const topSites = Object.entries(siteCount).sort((a,b)=>b[1]-a[1]).slice(0,10);

  const lite = classified.slice(0, 200).map(x => ({
    from: x.from, subject: x.subject,
    categories: Array.from(normalizeCategories_(x)).join(","),
    jobSiteName: x.jobSiteName || x.heuristicJobSiteName || ""
  }));

  const sumPrompt = `
You are an assistant generating a concise daily Gmail digest in English.
Group by: Ads, Jobs-Sites (note per-site), Job-Interview, School, Internal, External-Work, External-Personal, External-AcademicReview, Other.
Highlight urgent/interview/time-sensitive first, then actionable work items.
~150-250 words. End with 3-6 bullet action items.
Input JSON:
${JSON.stringify(lite, null, 2)}
  `.trim();

  const summary = fetchGeminiTEXT_(sumPrompt); // text mode

  const sitesHtml = topSites.length
    ? "<ul>" + topSites.map(([n,c]) => `<li>${n}: ${c}</li>`).join("") + "</ul>"
    : "<p>No job-site emails detected.</p>";

  return `
    <div style="font-family:Inter,Arial,sans-serif">
      <h2>📬 Gmail Daily Summary (${formatDate_(new Date())})</h2>
      <p><b>Total:</b> ${stats.total}</p>
      <ul>
        <li>Ads: ${stats.Ads}</li>
        <li>Jobs-Sites: ${stats.JobSites}</li>
        <li>Job-Interview: ${stats.Interview}</li>
        <li>School: ${stats.School}</li>
        <li>Internal: ${stats.Internal}</li>
        <li>External-Work: ${stats.ExtWork}</li>
        <li>External-Personal: ${stats.ExtPersonal}</li>
        <li>External-AcademicReview: ${stats.ExtReview}</li>
        <li>Unknown: ${stats.Unknown}</li>
      </ul>
      <h4>Top Job Sites</h4>
      ${sitesHtml}
      <hr/>
      <h3>LLM Summary</h3>
      <div>${nl2br_(escapeHtml_(summary))}</div>
    </div>
  `;
}

// ===================== LLM CALLS =====================

function buildClassifierPrompt_(batch) {
  const schema = {
    items: [{
      messageId: "string",
      categories: ["string"], // e.g. ["Jobs/Sites"], ["Jobs/Interview"], ["Ads"], ...
      reasons: ["string"],
      isExternal: "boolean|null",
      isInternal: "boolean|null",
      isJobSite: "boolean",
      isInterview: "boolean",
      jobSiteName: "string|null"
    }]
  };

  const siteGuide = KNOWN_JOB_SITES.map(s => ({ name: s.name, domains: s.domains || [], keywords: s.keywords || [] }));

  return `
You are an email classifier. Use fields: from, to, cc, subject, snippet (max ${CONFIG.BODY_TRUNCATE} chars), plus heuristic fields "heuristicJobSiteName" and "heuristicCategories".
Rules:
- Sender domain equals recipient primary domain → Internal; else External.
- Job platform notifications → "Jobs/Sites" (isJobSite=true). Provide jobSiteName when possible (prefer domain evidence).
- HR/recruiter/interview/application scheduling → "Jobs/Interview" (isInterview=true).
- University/course/admin → "School".
- External split: work-related → "External/Work"; personal → "External/Personal"; academic review (review invitations, decisions, reviewer center) → "External/AcademicReview".
- Ads/marketing/newsletters → "Ads".
- Otherwise → "Unknown".
Known job platforms (names, domains, keywords):
${JSON.stringify(siteGuide, null, 2)}
Return ONLY valid JSON (no markdown, no commentary) matching:
${JSON.stringify(schema, null, 2)}
Emails:
${JSON.stringify(batch, null, 2)}
  `.trim();
}

// Force JSON mode for classifier
function fetchGeminiJSON_(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("Please set GEMINI_API_KEY in Script Properties.");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }]}],
    generationConfig: { ...CONFIG.generationConfig, response_mime_type: "application/json" }
  };
  const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", muteHttpExceptions: true, payload: JSON.stringify(payload) });
  if (CONFIG.DEBUG_LOG_LLM) Logger.log("LLM(JSON) raw: " + res.getContentText().slice(0, 1000));
  const data = JSON.parse(res.getContentText());
  if (res.getResponseCode() >= 400) throw new Error(`Gemini error: ${res.getResponseCode()} ${res.getContentText()}`);
  const cand = data.candidates && data.candidates[0];
  const parts = cand && cand.content && cand.content.parts;
  return (parts && parts[0] && (parts[0].text || parts[0].inlineData?.data || "")) || "";
}

// Text mode for summary
function fetchGeminiTEXT_(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("Please set GEMINI_API_KEY in Script Properties.");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const payload = { contents: [{ role: "user", parts: [{ text: prompt }]}], generationConfig: CONFIG.generationConfig };
  const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", muteHttpExceptions: true, payload: JSON.stringify(payload) });
  if (res.getResponseCode() >= 400) throw new Error(`Gemini error: ${res.getResponseCode()} ${res.getContentText()}`);
  const data = JSON.parse(res.getContentText());
  const cand = data.candidates && data.candidates[0];
  const parts = cand && cand.content && cand.content.parts;
  return (parts && parts[0] && parts[0].text) || "";
}

// ===================== HELPERS =====================

function ensureLabels_() {
  const ensure = (name) => {
    const full = `${CONFIG.LABEL_ROOT}/${name}`;
    return GmailApp.getUserLabelByName(full) || GmailApp.createLabel(full);
  };
  return {
    ADS: ensure(LABELS.ADS),
    JOB_SITES: ensure(LABELS.JOB_SITES),
    JOB_INTERVIEW: ensure(LABELS.JOB_INTERVIEW),
    SCHOOL: ensure(LABELS.SCHOOL),
    INTERNAL: ensure(LABELS.INTERNAL),
    EXT_WORK: ensure(LABELS.EXT_WORK),
    EXT_PERSONAL: ensure(LABELS.EXT_PERSONAL),
    EXT_REVIEW: ensure(LABELS.EXT_REVIEW),
    UNKNOWN: ensure(LABELS.UNKNOWN)
  };
}

function ensureDynamicJobSiteLabel_(siteName) {
  if (!siteName) return null;
  const full = `${CONFIG.LABEL_ROOT}/${LABELS.JOB_SITES}/${sanitizeLabel_(siteName)}`;
  return GmailApp.getUserLabelByName(full) || GmailApp.createLabel(full);
}

function sanitizeLabel_(s) { return String(s).replace(/[^\w\-./ ]+/g, "").trim(); }

function normalizeCategories_(it) {
  const s = new Set();
  (it.categories || []).forEach(c => {
    const k = String(c || "").trim();
    switch (k) {
      case "Ads": s.add("Ads"); break;
      case "Jobs/Sites": case "Jobs/LinkedIn": s.add("Jobs/Sites"); break;
      case "Jobs/Interview": s.add("Jobs/Interview"); break;
      case "School": s.add("School"); break;
      case "Internal": s.add("Internal"); break;
      case "External/Work": s.add("External/Work"); break;
      case "External/Personal": s.add("External/Personal"); break;
      case "External/AcademicReview": s.add("External/AcademicReview"); break;
      default: s.add("Unknown");
    }
  });
  if (s.size === 0) s.add("Unknown");
  return s;
}

function chunk_(arr, size) { const out = []; for (let i=0;i<arr.length;i+=size) out.push(arr.slice(i,i+size)); return out; }

function safeParseJSON_(text) {
  if (!text) return null;
  let t = text.trim();
  // if model still wrapped with ```json fences, strip them
  t = t.replace(/^```json\s*/i, "").replace(/```$/i, "");
  const idx = t.indexOf("{");
  if (idx > 0) t = t.slice(idx);
  try { return JSON.parse(t); } catch { return null; }
}

function sendSummaryEmail_(body) {
  const subject = `${CONFIG.SUMMARY_SUBJECT_PREFIX} - ${formatDate_(new Date())}`;
  MailApp.sendEmail(CONFIG.SUMMARY_RECIPIENT, subject, body);
}

function nl2br_(s) { return s.replace(/\n/g, "<br>"); }
function escapeHtml_(s) { return s.replace(/[&<>"']/g, m => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[m])); }
function formatDate_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy-MM-dd"); }

function extractDomain_(fromField) {
  const match = String(fromField || "").match(/<([^>]+)>/);
  const email = (match ? match[1] : fromField || "").trim();
  const at = email.lastIndexOf("@");
  if (at === -1) return "";
  return email.slice(at + 1).toLowerCase();
}

// ============================================================
// JOB SEARCH AGENT v3 — Google Apps Script
// Full pipeline: profile-driven queries, job fetching,
// Gemini analysis, deduplication, profile refresh,
// dynamic category management, and resume catalog.
//
// Functions:
//   extractPostDate()      — regex helper to find dates in description text
//   getCategories()        — reads CATEGORIES tab, returns array of labels
//   addCategory()          — adds new category to CATEGORIES tab if not present
//   getRoleCategoryInstruction() — builds dynamic category prompt for Gemini
//   refreshProfile()       — reads new/updated resumes from Drive,
//                            updates PROFILE tab with fresh summary and queries
//   fetchJobs()            — reads queries from PROFILE tab, pulls jobs
//                            from JSearch, analyzes with Gemini
//   analyzeWithGemini()    — scores and categorizes a job description
//   analyzeExistingJobs()  — backfills analysis on rows missing it
//   backfillPostDates()    — fills actual_post_date for rows missing it
//   buildResumeCatalog()   — reads Resume folder Docs, uses Gemini to categorize,
//                            writes to RESUMES tab, auto-adds new categories
//   categorizeResume()     — calls Gemini to assign a category to one resume
//   getResumesByCategory() — lookup helper to filter resumes by category
// ============================================================


// ============================================================
// HELPER: EXTRACT POSTING DATE WITH REGEX
// ============================================================

function extractPostDate(text) {
  // Scans the full untruncated job description for common date patterns near
  // posting keywords — running on the full text before truncation means we never
  // miss a date that appears after the character cutoff sent to Gemini.
  if (!text) return "";

  var patterns = [
    /posted[:\s]+(\w+ \d{1,2},?\s*\d{4})/i,
    /posted[:\s]+(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i,
    /date posted[:\s]+(\w+ \d{1,2},?\s*\d{4})/i,
    /listing date[:\s]+(\w+ \d{1,2},?\s*\d{4})/i,
    /posted on[:\s]+(\w+ \d{1,2},?\s*\d{4})/i,
    /updated[:\s]+(\w+ \d{1,2},?\s*\d{4})/i,
    /listed[:\s]+(\w+ \d{1,2},?\s*\d{4})/i
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match) return match[1].trim();
  }
  return "";
}


// ============================================================
// HELPER: READ ALL EXISTING CATEGORIES FROM CATEGORIES TAB
// ============================================================

function getCategories() {
  // Opens the CATEGORIES tab and returns a flat array of all category label
  // strings currently in column A — this gets passed to Gemini so it can
  // pick a consistent existing label before creating something new.
  // Creates the tab with headers if it doesn't exist yet.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var catSheet = ss.getSheetByName("CATEGORIES");

  if (!catSheet) {
    catSheet = ss.insertSheet("CATEGORIES");
    catSheet.getRange(1, 1, 1, 3).setValues([["category", "date_added", "source"]]);
    Logger.log("Created new CATEGORIES tab.");
    return [];
  }

  const data = catSheet.getDataRange().getValues();
  var categories = [];

  data.forEach(function(row, i) {
    if (i === 0) return; // skip header row
    if (row[0]) categories.push(row[0]);
  });

  return categories;
}


// ============================================================
// HELPER: ADD NEW CATEGORY TO CATEGORIES TAB IF NOT PRESENT
// ============================================================

function addCategory(category, source) {
  // Checks whether the category already exists before adding — prevents
  // duplicates even if Gemini returns something already in the list.
  // Records the label, today's date, and whether it came from a job or resume.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catSheet = ss.getSheetByName("CATEGORIES");
  if (!catSheet) return;

  const data = catSheet.getDataRange().getValues();
  var exists = data.some(function(row) {
    return row[0].toString().toLowerCase() === category.toLowerCase();
  });

  if (!exists) {
    var today = new Date().toISOString().substring(0, 10);
    catSheet.appendRow([category, today, source]);
    Logger.log("New category added: " + category + " (source: " + source + ")");
  }
}


// ============================================================
// HELPER: BUILD DYNAMIC ROLE CATEGORY INSTRUCTION FOR GEMINI
// ============================================================

function getRoleCategoryInstruction() {
  // Reads current categories from the CATEGORIES tab and returns a dynamic
  // instruction string for the Gemini prompt — this keeps job categorization
  // consistent with resume categorization since both draw from the same list.
  var categories = getCategories();
  if (categories.length > 0) {
    return "pick the closest match from: " + categories.join(", ") + " — or create a concise new 2-4 word label if none fit, then it will be added to the category list automatically";
  } else {
    return "assign a concise 2-4 word category label that describes the target role";
  }
}


// ============================================================
// PROFILE REFRESH: READ NEW RESUMES FROM DRIVE, UPDATE PROFILE TAB
// ============================================================

function refreshProfile() {

  // --- CONFIG ---
  // RESUME_FOLDER_ID is the Google Drive folder ID for your Resume folder —
  // find it in the URL when you open the folder: drive.google.com/drive/folders/[ID]
  // PROFILE_SHEET_NAME must match your tab name exactly.
  const RESUME_FOLDER_ID = "1FKMiQkLfq2rcDFRnplIS7kUbtms92E0I";
  const PROFILE_SHEET_NAME = "PROFILE";
  const MAX_CHARS_PER_RESUME = 3000; // truncation per resume to manage token costs


  // --- GET PROFILE SHEET AND READ LAST UPDATED TIMESTAMP ---
  // The script reads the last_updated value from B2 to know which resumes to skip —
  // only Docs modified after this date get re-read, keeping repeat runs cheap.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName(PROFILE_SHEET_NAME);

  const lastUpdatedCell = profileSheet.getRange("B2").getValue();
  const lastUpdated = lastUpdatedCell ? new Date(lastUpdatedCell) : new Date(0);
  Logger.log("Last updated: " + lastUpdated);


  // --- FIND NEW OR UPDATED RESUME DOCS IN DRIVE ---
  // DriveApp opens the Resume folder and iterates through all files — we filter
  // to only Google Docs modified after the last_updated timestamp, so PDFs and
  // Word docs are automatically skipped without extra handling.
  const folder = DriveApp.getFolderById(RESUME_FOLDER_ID);
  const files = folder.getFiles();
  var resumeTexts = [];

  while (files.hasNext()) {
    var file = files.next();

    if (file.getMimeType() !== "application/vnd.google-apps.document") continue;

    if (file.getLastUpdated() <= lastUpdated) {
      Logger.log("Skipping unchanged: " + file.getName());
      continue;
    }

    Logger.log("Reading: " + file.getName());
    var doc = DocumentApp.openById(file.getId());
    var text = doc.getBody().getText();
    resumeTexts.push("=== " + file.getName() + " ===\n" + text.substring(0, MAX_CHARS_PER_RESUME));
  }

  if (resumeTexts.length === 0) {
    Logger.log("No new or updated resumes found. Profile is up to date.");
    return;
  }

  Logger.log("Found " + resumeTexts.length + " new/updated resumes. Sending to Gemini.");

  // --- READ EXISTING PROFILE SUMMARY ---
  // Passes the existing summary to Gemini alongside only the changed resumes —
  // Gemini merges new info into the existing summary rather than rewriting
  // everything, preserving detail built up over previous runs.
  const existingSummary = profileSheet.getRange("B3").getValue() || "";
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");

  const prompt = `You are updating a job seeker's career profile based on new or updated resume content.

Here is the existing profile summary:
${existingSummary}

Here are the new or updated resumes to incorporate:
${resumeTexts.join("\n\n")}

Based on all of this, return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "profile_summary": "updated 3-5 sentence profile summary incorporating any new experience, skills, or roles found in the updated resumes — preserve existing accurate information, add or update anything new",
  "queries": [
    "search query 1",
    "search query 2",
    "search query 3",
    "search query 4",
    "search query 5",
    "search query 6",
    "search query 7",
    "search query 8",
    "search query 9",
    "search query 10"
  ]
}

Query rules:
- Each query should target a different angle of the candidate's background
- Queries should be natural language job board search strings
- Include "remote" in most queries given the candidate's preference
- Cover: LLM evaluation, AI quality ops, AI enablement/training, localization, prompt engineering
- Avoid overly generic queries — be specific enough to surface relevant roles`;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  var result;
  try {
    const text = json.candidates[0].content.parts[0].text;
    const cleaned = text.replace(/```json|```/g, "").trim();
    result = JSON.parse(cleaned);
  } catch(e) {
    Logger.log("Gemini parse error in refreshProfile: " + e + " | Raw: " + response.getContentText());
    return;
  }

  // --- WRITE UPDATED PROFILE BACK TO SHEET ---
  // Overwrites the PROFILE tab with the new summary, updated queries, and a
  // fresh last_updated timestamp — the next fetchJobs() run will automatically
  // use the new queries without any other changes needed.
  profileSheet.getRange("B2").setValue(new Date().toISOString());
  profileSheet.getRange("B3").setValue(result.profile_summary);

  var queries = result.queries;
  for (var i = 0; i < queries.length; i++) {
    profileSheet.getRange(4 + i, 1).setValue("query_" + (i + 1));
    profileSheet.getRange(4 + i, 2).setValue(queries[i]);
  }
  for (var j = queries.length; j < 20; j++) {
    profileSheet.getRange(4 + j, 1).setValue("");
    profileSheet.getRange(4 + j, 2).setValue("");
  }

  Logger.log("Profile refreshed. " + queries.length + " queries written.");
}


// ============================================================
// MAIN: FETCH NEW JOBS USING PROFILE QUERIES, ANALYZE WITH GEMINI
// ============================================================

function fetchJobs() {

  // --- CONFIG ---
  // NUM_PAGES is per query — with 10 queries and 1 page each, that's up to
  // 100 jobs per run using 10 JSearch API requests from your 200/month budget.
  // DATE_POSTED filters to jobs Google indexed in the last 7 days.
  const SHEET_NAME = "RAW";
  const PROFILE_SHEET_NAME = "PROFILE";
  const NUM_PAGES = 1;
  const DATE_POSTED = "week";

  // --- GET JSEARCH API KEY ---
  // Retrieved from Script Properties — never written into the code itself.
  const jsearchKey = PropertiesService.getScriptProperties().getProperty("JSEARCH_KEY");
  if (!jsearchKey) {
    Logger.log("ERROR: No JSearch API key found. Add JSEARCH_KEY to Script Properties.");
    return;
  }

  // --- READ QUERIES AND PROFILE SUMMARY FROM PROFILE TAB ---
  // Reads all rows where column A starts with "query_" and collects column B
  // values as the search queries to run — adding or removing queries only
  // requires editing the PROFILE tab, not the code.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName(PROFILE_SHEET_NAME);
  const profileData = profileSheet.getDataRange().getValues();

  var queries = [];
  var profileSummary = "";

  profileData.forEach(function(row) {
    if (row[0] === "profile_summary") profileSummary = row[1];
    if (String(row[0]).startsWith("query_") && row[1]) queries.push(row[1]);
  });

  if (queries.length === 0) {
    Logger.log("ERROR: No queries found in PROFILE tab.");
    return;
  }

  Logger.log("Running " + queries.length + " queries from PROFILE tab.");

  // --- GET RAW SHEET AND HEADERS ---
  // Reads the header row into an array so every column can be found by name
  // using indexOf() regardless of its position in the sheet.
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // --- BUILD DEDUPLICATION SET ---
  // Reads all existing job_id values into a Set before any queries run —
  // the same job appearing in multiple query results only gets written once
  // and only analyzed once, saving both JSearch quota and Gemini tokens.
  const lastRow = sheet.getLastRow();
  var existingIds = new Set();
  if (lastRow > 1) {
    const jobIdCol = headers.indexOf("job_id") + 1;
    const existingIdValues = sheet.getRange(2, jobIdCol, lastRow - 1, 1).getValues();
    existingIdValues.forEach(function(row) {
      if (row[0]) existingIds.add(row[0]);
    });
  }

  var totalNewCount = 0;

  // --- LOOP THROUGH EACH QUERY ---
  // For each query from the PROFILE tab, builds the JSearch URL, fetches results,
  // and processes each job — deduplication Set prevents the same job appearing
  // twice even when multiple queries return overlapping results.
  queries.forEach(function(query, queryIndex) {
    Logger.log("Query " + (queryIndex + 1) + "/" + queries.length + ": " + query);

    const url = "https://jsearch.p.rapidapi.com/search-v2"
      + "?query=" + encodeURIComponent(query)
      + "&num_pages=" + NUM_PAGES
      + "&country=us"
      + "&date_posted=" + DATE_POSTED;

    const options = {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        "X-RapidAPI-Key": jsearchKey,
        "X-RapidAPI-Host": "jsearch.p.rapidapi.com"
      },
      muteHttpExceptions: true
    };

    var response, json;
    try {
      response = UrlFetchApp.fetch(url, options);
      json = JSON.parse(response.getContentText());
    } catch(e) {
      Logger.log("JSearch error on query '" + query + "': " + e);
      return;
    }

    if (json.status !== "OK" || !json.data || !json.data.jobs) {
      Logger.log("Bad response for query '" + query + "': " + response.getContentText().substring(0, 200));
      return;
    }

    var jobs = json.data.jobs;
    Logger.log("  → " + jobs.length + " results");
    var newCount = 0;

    jobs.forEach(function(job) {

      if (existingIds.has(job.job_id)) return;

      // --- MAP JSEARCH FIELDS TO COLUMN HEADERS ---
      // headers.map() goes through every column header and matches it to the
      // corresponding JSearch field — agent analysis columns are left blank
      // since Gemini fills them immediately after the row is written.
      const rowData = headers.map(function(header) {
        switch(header) {
          case "job_id":                    return job.job_id || "";
          case "job_title":                 return job.job_title || "";
          case "employer_name":             return job.employer_name || "";
          case "employer_website":          return job.employer_website || "";
          case "job_employment_type":       return job.job_employment_type || "";
          case "job_apply_link":            return job.job_apply_link || "";
          case "job_apply_is_direct":       return job.job_apply_is_direct || "";
          case "job_posted_at_datetime_utc":return job.job_posted_at_datetime_utc || "";
          case "job_location":              return job.job_location || "";
          case "job_city":                  return job.job_city || "";
          case "job_country":               return job.job_country || "";
          case "job_min_salary":            return job.job_min_salary || "";
          case "job_max_salary":            return job.job_max_salary || "";
          case "work_arrangement":          return job.work_arrangement || "";
          case "seniority_level":           return job.seniority_level || "";
          case "ai_ml_involved":            return job.ai_ml_involved || "";
          case "visa_sponsorship":          return job.visa_sponsorship || "";
          case "industry":                  return job.industry || "";
          case "job_description":           return job.job_description || "";
          case "role_category":             return "";
          case "fit_score":                 return "";
          case "rationale":                 return "";
          case "resume_recommended":        return "";
          case "resume_edits_suggested":    return "";
          case "actual_post_date":          return "";
          case "status":                    return "New";
          case "date_applied":              return "";
          case "where_applied":             return "JSearch";
          case "interview_stage":           return "";
          case "follow_up_needed":          return "No";
          case "notes":                     return "";
          default:                          return "";
        }
      });

      sheet.appendRow(rowData);
      existingIds.add(job.job_id);
      newCount++;
      totalNewCount++;

      // --- EXTRACT DATE AND RUN GEMINI ANALYSIS ---
      // extractPostDate runs on the FULL description before truncation —
      // Gemini then gets the smart-truncated version (first 1500 + last 500 chars)
      // to capture both intro context and requirements while skipping middle filler.
      // Profile summary from PROFILE tab grounds the fit scoring against real content.
      // New role_category values are auto-added to the CATEGORIES tab.
      const fullDesc = job.job_description || "";
      const regexDate = extractPostDate(fullDesc);

      Utilities.sleep(4000); // stay under Gemini free tier rate limit of ~15 req/min

      const newRowNumber = sheet.getLastRow();
      const shortDesc = fullDesc.substring(0, 1500) + "\n...\n" + fullDesc.slice(-500);
      const analysis = analyzeWithGemini(shortDesc, profileSummary);

      // Auto-add any new role_category to the CATEGORIES tab
      if (analysis.role_category) addCategory(analysis.role_category, "job");

      sheet.getRange(newRowNumber, headers.indexOf("role_category") + 1).setValue(analysis.role_category || "");
      sheet.getRange(newRowNumber, headers.indexOf("fit_score") + 1).setValue(analysis.fit_score || "");
      sheet.getRange(newRowNumber, headers.indexOf("rationale") + 1).setValue(analysis.rationale || "");
      sheet.getRange(newRowNumber, headers.indexOf("resume_recommended") + 1).setValue(analysis.resume_recommended || "");
      sheet.getRange(newRowNumber, headers.indexOf("resume_edits_suggested") + 1).setValue(analysis.resume_edits_suggested || "");
      sheet.getRange(newRowNumber, headers.indexOf("actual_post_date") + 1).setValue(regexDate || analysis.actual_post_date || "");

      Logger.log("  Analyzed: " + job.job_title + " at " + job.employer_name + " | Score: " + analysis.fit_score + " | Category: " + analysis.role_category);
    });

    Logger.log("  → " + newCount + " new jobs added from this query.");
    Utilities.sleep(1000); // brief pause between queries
  });

  Logger.log("Done. Total new jobs added: " + totalNewCount);
}


// ============================================================
// GEMINI ANALYSIS: SCORE AND CATEGORIZE A JOB DESCRIPTION
// ============================================================

function analyzeWithGemini(jobDescription, profileSummary) {

  // --- GET GEMINI API KEY ---
  // Retrieved from Script Properties — never written into the code itself.
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");

  // --- USE PROFILE SUMMARY IF PROVIDED, OTHERWISE USE FALLBACK ---
  // fetchJobs() passes the live profile summary from the PROFILE tab so Gemini
  // always scores against current resume content — analyzeExistingJobs() reads
  // it from the PROFILE tab too, with a hardcoded fallback if the tab is missing.
  const candidateProfile = profileSummary || `8+ years experience in AI quality operations,
LLM evaluation, annotation pipeline management, and training design. Skills include rubric
development, adversarial prompt testing, KPI dashboard design, Power Automate, Python.
Languages: C1 French, conversational German. Prefers remote, based in Kalamazoo MI,
open to EU relocation. Target roles: AI Eval, AI Ops, AI Enablement, Localization.`;

  // --- GET DYNAMIC CATEGORY INSTRUCTION ---
  // Built before the prompt so it can be interpolated cleanly as a variable —
  // reads the CATEGORIES tab to ensure job categorization stays consistent
  // with resume categorization since both draw from the same list.
  const categoryInstruction = getRoleCategoryInstruction();

  // --- BUILD THE PROMPT ---
  // The candidate profile grounds every analysis in real resume content pulled from
  // the PROFILE tab — Gemini returns a JSON object whose keys match sheet column
  // headers exactly so the script can write each value by column name with no
  // translation or mapping needed between what Gemini returns and where it goes.
  const prompt = `You are analyzing a job posting for a candidate with this background:
${candidateProfile}

Here is the job description:
${jobDescription}

Return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "role_category": "${categoryInstruction}",
  "fit_score": "integer 0-100",
  "rationale": "one sentence explaining the score",
  "resume_recommended": "a short label describing the best resume type for this role",
  "resume_edits_suggested": "one short sentence on what to tweak, or blank if none",
  "actual_post_date": "look for any date near words like posted, date, listed, or updated — return as YYYY-MM-DD or empty string"
}`;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  // --- PARSE GEMINI RESPONSE ---
  // Drills into the nested response structure, strips accidental markdown backticks,
  // and parses the JSON — returns a safe empty object on failure so the script
  // never crashes mid-run; failed rows can be retried with analyzeExistingJobs().
  try {
    const text = json.candidates[0].content.parts[0].text;
    const cleaned = text.replace(/```json|```/g, "").trim();
    return JSON.parse(cleaned);
  } catch(e) {
    Logger.log("Gemini parse error: " + e + " | Raw: " + response.getContentText());
    return {
      role_category: "",
      fit_score: "",
      rationale: "Analysis failed",
      resume_recommended: "",
      resume_edits_suggested: "",
      actual_post_date: ""
    };
  }
}


// ============================================================
// BACKFILL: ANALYZE EXISTING ROWS MISSING GEMINI OUTPUT
// ============================================================

function analyzeExistingJobs() {

  // --- GET SHEET, HEADERS, AND PROFILE SUMMARY ---
  // Reads the PROFILE tab to get the current profile summary for grounding Gemini —
  // falls back to the hardcoded string in analyzeWithGemini() if the tab is missing.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RAW");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();

  var profileSummary = "";
  try {
    const profileSheet = ss.getSheetByName("PROFILE");
    const profileData = profileSheet.getDataRange().getValues();
    profileData.forEach(function(row) {
      if (row[0] === "profile_summary") profileSummary = row[1];
    });
  } catch(e) {
    Logger.log("Could not read PROFILE tab — using fallback profile.");
  }

  if (lastRow < 2) {
    Logger.log("No data rows found.");
    return;
  }

  // --- FIND COLUMN POSITIONS BY NAME ---
  // indexOf() returns the zero-based array position — adding 1 converts to the
  // one-based column number that getRange() expects when writing to specific cells.
  const fitScoreCol = headers.indexOf("fit_score") + 1;
  const jobDescCol = headers.indexOf("job_description") + 1;
  const roleCatCol = headers.indexOf("role_category") + 1;
  const rationaleCol = headers.indexOf("rationale") + 1;
  const resumeRecCol = headers.indexOf("resume_recommended") + 1;
  const resumeEditsCol = headers.indexOf("resume_edits_suggested") + 1;
  const actualPostDateCol = headers.indexOf("actual_post_date") + 1;

  // Read all rows in one call — much faster than reading one row at a time
  // and avoids hitting Apps Script execution time limits on large sheets.
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var analyzed = 0;
  var skipped = 0;

  // --- LOOP THROUGH ROWS AND FILL BLANKS ---
  // Only processes rows where fit_score is empty AND job_description exists —
  // rows already analyzed or missing a description are skipped entirely.
  allData.forEach(function(row, i) {
    const fitScore = row[fitScoreCol - 1];
    const jobDesc = row[jobDescCol - 1];

    if (fitScore !== "" || !jobDesc) {
      skipped++;
      return;
    }

    Logger.log("Analyzing row " + (i + 2) + ": " + row[headers.indexOf("job_title")]);

    const regexDate = extractPostDate(jobDesc);
    Utilities.sleep(4000);

    const shortDesc = jobDesc.substring(0, 1500) + "\n...\n" + jobDesc.slice(-500);
    const analysis = analyzeWithGemini(shortDesc, profileSummary);

    if (analysis.role_category) addCategory(analysis.role_category, "job");

    const rowNumber = i + 2;
    sheet.getRange(rowNumber, roleCatCol).setValue(analysis.role_category || "");
    sheet.getRange(rowNumber, fitScoreCol).setValue(analysis.fit_score || "");
    sheet.getRange(rowNumber, rationaleCol).setValue(analysis.rationale || "");
    sheet.getRange(rowNumber, resumeRecCol).setValue(analysis.resume_recommended || "");
    sheet.getRange(rowNumber, resumeEditsCol).setValue(analysis.resume_edits_suggested || "");
    sheet.getRange(rowNumber, actualPostDateCol).setValue(regexDate || analysis.actual_post_date || "");

    analyzed++;
  });

  Logger.log("Done. Analyzed: " + analyzed + " | Skipped: " + skipped);
}


// ============================================================
// BACKFILL: FILL actual_post_date FOR ROWS MISSING IT (REGEX ONLY)
// ============================================================

function backfillPostDates() {

  // Targets rows that have a job_description but no actual_post_date — runs regex
  // on the full description text without calling Gemini, so this is fast, free,
  // and safe to run on large numbers of rows at any time.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RAW");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No data rows found.");
    return;
  }

  const jobDescCol = headers.indexOf("job_description") + 1;
  const actualPostDateCol = headers.indexOf("actual_post_date") + 1;
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var updated = 0;
  var skipped = 0;

  allData.forEach(function(row, i) {
    const jobDesc = row[jobDescCol - 1];
    const existingDate = row[actualPostDateCol - 1];

    if (!jobDesc || existingDate !== "") {
      skipped++;
      return;
    }

    const regexDate = extractPostDate(jobDesc);
    if (regexDate) {
      sheet.getRange(i + 2, actualPostDateCol).setValue(regexDate);
      Logger.log("Row " + (i + 2) + ": found date " + regexDate);
      updated++;
    } else {
      skipped++;
    }
  });

  Logger.log("Done. Updated: " + updated + " | Skipped: " + skipped);
}


// ============================================================
// RESUME CATALOG: BUILD AND MAINTAIN RESUMES TAB
// ============================================================

function buildResumeCatalog() {

  // --- CONFIG ---
  // CUTOFF_YEAR automatically excludes 2017-2020 era resumes without manual
  // filtering — anything last modified before this year is skipped entirely.
  // MAX_CHARS controls how much of each resume Gemini reads for categorization —
  // 500 characters is enough to identify the role focus without burning tokens.
  const RESUME_FOLDER_ID = "1FKMiQkLfq2rcDFRnplIS7kUbtms92E0I";
  const RESUMES_SHEET_NAME = "RESUMES";
  const CUTOFF_YEAR = 2023;
  const MAX_CHARS = 500;

  // --- FILES TO EXCLUDE ---
  // These are in the Resume folder but aren't resumes — job applications,
  // interview prep docs, and old translation files are all excluded.
  const EXCLUDE_NAMES = [
    "Application_Grames",
    "ANet Interview Question",
    "Employment Application",
    "CV_final",
    "Service+Translation",
    "service_resume",
    "resume_grames",
    "resume_updated",
    "Grant Grames  Resume",
    "Grant Grames Resume.docx"
  ];

  // --- GET OR CREATE RESUMES SHEET ---
  // Creates the RESUMES tab with headers if it doesn't exist yet —
  // on subsequent runs just opens the existing tab to append new rows.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var resumeSheet = ss.getSheetByName(RESUMES_SHEET_NAME);
  if (!resumeSheet) {
    resumeSheet = ss.insertSheet(RESUMES_SHEET_NAME);
    resumeSheet.getRange(1, 1, 1, 5).setValues([["filename", "doc_id", "category", "last_modified", "doc_url"]]);
    Logger.log("Created new RESUMES tab.");
  }

  // --- READ EXISTING CATALOG ---
  // Reads all doc_id values already in the catalog into a Set so we can skip
  // files already processed — only new files get a Gemini call, keeping repeat
  // runs fast and cheap regardless of how many resumes are in the folder.
  const existingData = resumeSheet.getDataRange().getValues();
  var existingIds = new Set();
  var existingRowNums = {};

  existingData.forEach(function(row, i) {
    if (i === 0) return;
    if (row[1]) {
      existingIds.add(row[1]);
      existingRowNums[row[1]] = i + 1;
    }
  });

  // --- READ EXISTING CATEGORIES ---
  // Gets the current category list before processing any files — passed to Gemini
  // for every resume so categorization stays consistent across the whole run.
  var categories = getCategories();
  Logger.log("Existing categories: " + (categories.length > 0 ? categories.join(", ") : "none yet"));

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");
  const folder = DriveApp.getFolderById(RESUME_FOLDER_ID);
  const files = folder.getFiles();
  var added = 0;
  var updated = 0;
  var skipped = 0;

  while (files.hasNext()) {
    var file = files.next();

    // Skip non-Google Docs — PDFs and Word docs can't be read as plain text
    if (file.getMimeType() !== "application/vnd.google-apps.document") {
      skipped++;
      continue;
    }

    var fileName = file.getName();

    // Skip excluded filenames
    var isExcluded = EXCLUDE_NAMES.some(function(excluded) {
      return fileName.toLowerCase().indexOf(excluded.toLowerCase()) !== -1;
    });
    if (isExcluded) {
      Logger.log("Excluding: " + fileName);
      skipped++;
      continue;
    }

    // Skip files last modified before the cutoff year
    var lastModified = file.getLastUpdated();
    if (lastModified.getFullYear() < CUTOFF_YEAR) {
      Logger.log("Skipping old file (" + lastModified.getFullYear() + "): " + fileName);
      skipped++;
      continue;
    }

    var docId = file.getId();
    var docUrl = file.getUrl();
    var lastModifiedStr = lastModified.toISOString().substring(0, 10);

    // If already cataloged, just update the timestamp and move on —
    // preserves any manual category corrections the user may have made.
    if (existingIds.has(docId)) {
      resumeSheet.getRange(existingRowNums[docId], 4).setValue(lastModifiedStr);
      updated++;
      Logger.log("Updated timestamp: " + fileName);
      continue;
    }

    // --- READ RESUME CONTENT AND CATEGORIZE ---
    // Opens the Doc and reads just the first MAX_CHARS characters — enough for
    // Gemini to identify the role focus without reading the whole resume.
    // The categories array is refreshed after each new addition so subsequent
    // calls within the same run see the updated list immediately.
    var doc = DocumentApp.openById(docId);
    var resumeText = doc.getBody().getText().substring(0, MAX_CHARS);
    var category = categorizeResume(resumeText, fileName, categories, apiKey);

    addCategory(category, "resume");

    if (categories.indexOf(category) === -1) categories.push(category);

    resumeSheet.appendRow([fileName, docId, category, lastModifiedStr, docUrl]);
    existingIds.add(docId);
    added++;

    Logger.log("Added: " + fileName + " → " + category);
    Utilities.sleep(2000); // brief pause to avoid Gemini rate limits
  }

  Logger.log("Done. Added: " + added + " | Updated: " + updated + " | Skipped: " + skipped);
  Logger.log("Review the RESUMES tab and manually correct any miscategorizations.");
}


// ============================================================
// HELPER: CATEGORIZE A RESUME USING GEMINI
// ============================================================

function categorizeResume(resumeText, fileName, existingCategories, apiKey) {
  // Sends the resume excerpt and existing category list to Gemini, asking it
  // to pick the closest matching label or invent a concise new one if nothing
  // fits — returns the category string for writing to the RESUMES tab.

  var categoriesStr = existingCategories.length > 0
    ? existingCategories.join(", ")
    : "none yet — create the first one";

  var prompt = `You are categorizing a resume into a role-based category for a job search system.

Existing categories (use one of these if it fits): ${categoriesStr}

Resume filename: ${fileName}
Resume excerpt (first 500 characters):
${resumeText}

Rules:
- If the resume clearly fits an existing category, return that exact label
- If no existing category fits well, create a concise new label (2-4 words max)
- Categories should describe the TARGET ROLE, not the candidate's background
- Be consistent — "AI Ops" not "AI Operations", "AI Eval" not "AI Evaluation"
- Never return more than one category

Return ONLY the category label, nothing else. No explanation, no punctuation, no quotes.`;

  var payload = { contents: [{ parts: [{ text: prompt }] }] };
  var options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    var text = json.candidates[0].content.parts[0].text.trim();
    text = text.replace(/["'.,]/g, "").trim();
    return text || "General";
  } catch(e) {
    Logger.log("Gemini categorization error for " + fileName + ": " + e);
    return "General";
  }
}


// ============================================================
// HELPER: GET RESUMES BY CATEGORY FOR JOB MATCHING
// ============================================================

function getResumesByCategory(category) {
  // Reads the RESUMES tab and returns all rows where the category matches
  // the given label — used when building the resume recommendation feature
  // so Gemini only reads relevant resume Docs rather than all 30+ files.
  // Always includes General resumes as a fallback option.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resumeSheet = ss.getSheetByName("RESUMES");
  if (!resumeSheet) {
    Logger.log("No RESUMES tab found — run buildResumeCatalog() first.");
    return [];
  }

  const data = resumeSheet.getDataRange().getValues();
  var matches = [];

  data.forEach(function(row, i) {
    if (i === 0) return;
    var rowCategory = row[2];
    if (rowCategory === category || rowCategory === "General") {
      matches.push({
        filename: row[0],
        doc_id: row[1],
        category: row[2],
        url: row[4]
      });
    }
  });

  Logger.log("Found " + matches.length + " resumes for category: " + category);
  return matches;
}

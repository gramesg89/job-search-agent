// ============================================================
// JOB SEARCH AGENT v2 — Google Apps Script
// Full pipeline: profile-driven queries, job fetching, 
// Gemini analysis, deduplication, and profile refresh.
//
// Functions:
//   refreshProfile()      — reads new/updated resumes from Drive,
//                           updates PROFILE tab with fresh summary and queries
//   fetchJobs()           — reads queries from PROFILE tab, pulls jobs
//                           from JSearch, analyzes with Gemini
//   analyzeWithGemini()   — scores and categorizes a job description
//   analyzeExistingJobs() — backfills analysis on rows missing it
//   extractPostDate()     — regex helper to find dates in description text
//   backfillPostDates()   — fills actual_post_date for rows missing it
// ============================================================


// ============================================================
// HELPER: EXTRACT POSTING DATE WITH REGEX
// ============================================================

function extractPostDate(text) {
  // Scans the full untruncated job description for common date patterns near
  // posting keywords — running on the full text before truncation means we never
  // miss a date that appears after the 2000 character cutoff sent to Gemini.
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
// PROFILE REFRESH: READ NEW RESUMES FROM DRIVE, UPDATE PROFILE TAB
// ============================================================

function refreshProfile() {

  // --- CONFIG ---
  // RESUME_FOLDER_ID is the Google Drive folder ID for your Resume folder —
  // find it in the URL when you open the folder: drive.google.com/drive/folders/[ID]
  // PROFILE_SHEET_NAME must match your tab name exactly.
  const RESUME_FOLDER_ID = "1FKMiQkLfq2rcDFRnplIS7kUbtms92E0I";
  const PROFILE_SHEET_NAME = "PROFILE";
  const MAX_CHARS_PER_RESUME = 3000; // truncation limit per resume to manage token costs


  // --- GET PROFILE SHEET AND READ LAST UPDATED TIMESTAMP ---
  // The script reads the last_updated value from B2 to know which resumes to skip —
  // only Docs modified after this date get re-read, so large resume libraries stay
  // cheap to process on every refresh rather than re-sending everything each time.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName(PROFILE_SHEET_NAME);

  const lastUpdatedCell = profileSheet.getRange("B2").getValue();
  const lastUpdated = lastUpdatedCell ? new Date(lastUpdatedCell) : new Date(0); // if blank, treat as epoch so all files are read
  Logger.log("Last updated: " + lastUpdated);


  // --- FIND NEW OR UPDATED RESUME DOCS IN DRIVE ---
  // DriveApp opens the Resume folder and iterates through all files — we filter
  // to only Google Docs (mimeType check) modified after the last_updated timestamp,
  // so PDFs and Word docs are automatically skipped without extra handling.
  const folder = DriveApp.getFolderById(RESUME_FOLDER_ID);
  const files = folder.getFiles();
  var resumeTexts = [];

  while (files.hasNext()) {
    var file = files.next();

    // Skip anything that isn't a Google Doc
    if (file.getMimeType() !== "application/vnd.google-apps.document") {
      continue;
    }

    // Skip files not modified since last refresh
    if (file.getLastUpdated() <= lastUpdated) {
      Logger.log("Skipping unchanged: " + file.getName());
      continue;
    }

    // Read the Doc content as plain text and truncate to save tokens
    Logger.log("Reading: " + file.getName());
    var doc = DocumentApp.openById(file.getId());
    var text = doc.getBody().getText();
    var truncated = text.substring(0, MAX_CHARS_PER_RESUME);
    resumeTexts.push("=== " + file.getName() + " ===\n" + truncated);
  }

  if (resumeTexts.length === 0) {
    Logger.log("No new or updated resumes found since last refresh. Profile is up to date.");
    return;
  }

  Logger.log("Found " + resumeTexts.length + " new/updated resumes. Sending to Gemini.");


  // --- READ EXISTING PROFILE SUMMARY FROM SHEET ---
  // Rather than regenerating from scratch, we pass the existing summary to Gemini
  // alongside only the changed resumes — Gemini merges the new information into
  // the existing summary rather than rewriting everything, keeping costs low.
  const existingSummary = profileSheet.getRange("B3").getValue() || "";


  // --- BUILD GEMINI PROMPT FOR PROFILE UPDATE ---
  // This prompt gives Gemini the existing profile and only the new/changed resumes,
  // asking it to return an updated profile summary and a fresh set of search queries
  // as a JSON object — the keys map directly to row labels in the PROFILE tab.
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


  // --- CALL GEMINI AND PARSE RESPONSE ---
  // Same API call pattern as analyzeWithGemini — POST to generateContent, drill into
  // the nested response structure, strip any accidental markdown, and parse as JSON.
  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

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
  // This overwrites the PROFILE tab with the new summary, updated queries, and a
  // fresh last_updated timestamp — the next fetchJobs() run will automatically
  // use the new queries without any other changes needed.
  profileSheet.getRange("B2").setValue(new Date().toISOString()); // update timestamp
  profileSheet.getRange("B3").setValue(result.profile_summary);   // update summary

  // Clear existing queries and write fresh ones
  // Queries start at row 4 (after last_updated in row 2 and profile_summary in row 3)
  var queries = result.queries;
  for (var i = 0; i < queries.length; i++) {
    profileSheet.getRange(4 + i, 1).setValue("query_" + (i + 1)); // column A label
    profileSheet.getRange(4 + i, 2).setValue(queries[i]);          // column B value
  }

  // Clear any leftover rows from a previous run that had more queries
  for (var j = queries.length; j < 20; j++) {
    profileSheet.getRange(4 + j, 1).setValue("");
    profileSheet.getRange(4 + j, 2).setValue("");
  }

  Logger.log("Profile refreshed successfully. " + queries.length + " queries written.");
}


// ============================================================
// MAIN: FETCH NEW JOBS USING PROFILE QUERIES, ANALYZE WITH GEMINI
// ============================================================

function fetchJobs() {

  // --- CONFIG ---
  // PROFILE_SHEET_NAME must match your tab name exactly — the script reads all
  // query_* rows from this tab instead of using a hardcoded query constant,
  // so changing queries only requires updating the PROFILE tab, not the code.
  const SHEET_NAME = "RAW";
  const PROFILE_SHEET_NAME = "PROFILE";
  const NUM_PAGES = 1; // per query — with 10 queries this gives up to 100 jobs per run
  const DATE_POSTED = "week";


  // --- GET API KEYS ---
  // Both keys are retrieved from Script Properties — never written into the code —
  // so the script is safe to share or push to GitHub without exposing credentials.
  const jsearchKey = PropertiesService.getScriptProperties().getProperty("JSEARCH_KEY");
  if (!jsearchKey) {
    Logger.log("ERROR: No JSearch API key found. Add JSEARCH_KEY to Script Properties.");
    return;
  }


  // --- READ QUERIES AND PROFILE SUMMARY FROM PROFILE TAB ---
  // The script reads all rows from the PROFILE tab where column A starts with "query_"
  // and collects the corresponding column B values as the search queries to run —
  // this means adding or removing queries only requires editing the sheet, not the code.
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
    Logger.log("ERROR: No queries found in PROFILE tab. Add query_1, query_2 etc. in column A.");
    return;
  }

  Logger.log("Running " + queries.length + " queries from PROFILE tab.");


  // --- GET RAW SHEET AND HEADERS ---
  // Same pattern as before — reads the header row into an array so every column
  // can be found by name using indexOf() regardless of its position in the sheet.
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];


  // --- BUILD DEDUPLICATION SET ---
  // Reads all existing job_id values into a Set before any queries run — this
  // deduplicates across all queries in one run, so the same job appearing in
  // multiple query results only gets written once and only analyzed once.
  const lastRow = sheet.getLastRow();
  var existingIds = new Set();
  if (lastRow > 1) {
    const jobIdCol = headers.indexOf("job_id") + 1;
    const existingIdValues = sheet.getRange(2, jobIdCol, lastRow - 1, 1).getValues();
    existingIdValues.forEach(function(row) {
      if (row[0]) existingIds.add(row[0]);
    });
  }


  // --- LOOP THROUGH EACH QUERY ---
  // For each query from the PROFILE tab, this builds the JSearch URL, fetches results,
  // and processes each job — the deduplication Set prevents the same job appearing
  // twice even when multiple queries return overlapping results.
  var totalNewCount = 0;

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

      // Skip duplicates — checks against both existing sheet rows AND jobs
      // already written earlier in this same run by a previous query
      if (existingIds.has(job.job_id)) {
        return;
      }


      // --- MAP JSEARCH FIELDS TO COLUMN HEADERS ---
      // Each column header is matched to its JSearch field name — custom agent
      // columns are left blank here since Gemini fills them immediately after.
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
      // Regex runs on the full description before any truncation, then Gemini
      // gets the smart-truncated version (start + end) to capture both the
      // intro context and the requirements section while skipping middle filler.
      // The profile summary from the PROFILE tab is passed as grounding context
      // so Gemini scores against your actual background, not a hardcoded string.
      const fullDesc = job.job_description || "";
      const regexDate = extractPostDate(fullDesc);

      Utilities.sleep(4000); // stay under Gemini free tier rate limit of ~15 req/min

      const newRowNumber = sheet.getLastRow();
      const descStart = fullDesc.substring(0, 1500);
      const descEnd = fullDesc.slice(-500);
      const shortDesc = descStart + "\n...\n" + descEnd;
      const analysis = analyzeWithGemini(shortDesc, profileSummary);

      sheet.getRange(newRowNumber, headers.indexOf("role_category") + 1).setValue(analysis.role_category || "");
      sheet.getRange(newRowNumber, headers.indexOf("fit_score") + 1).setValue(analysis.fit_score || "");
      sheet.getRange(newRowNumber, headers.indexOf("rationale") + 1).setValue(analysis.rationale || "");
      sheet.getRange(newRowNumber, headers.indexOf("resume_recommended") + 1).setValue(analysis.resume_recommended || "");
      sheet.getRange(newRowNumber, headers.indexOf("resume_edits_suggested") + 1).setValue(analysis.resume_edits_suggested || "");
      sheet.getRange(newRowNumber, headers.indexOf("actual_post_date") + 1).setValue(regexDate || analysis.actual_post_date || "");

      Logger.log("  Analyzed: " + job.job_title + " at " + job.employer_name + " | Score: " + analysis.fit_score);
    });

    Logger.log("  → " + newCount + " new jobs added from this query.");

    // Brief pause between queries to avoid hammering JSearch
    Utilities.sleep(1000);
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
  // always scores against current resume content — analyzeExistingJobs() may not
  // have the summary available so a hardcoded fallback ensures it still works.
  const candidateProfile = profileSummary || `8+ years experience in AI quality operations, 
LLM evaluation, annotation pipeline management, and training design. Skills include rubric 
development, adversarial prompt testing, KPI dashboard design, Power Automate, Python. 
Languages: C1 French, conversational German. Prefers remote, based in Kalamazoo MI, 
open to EU relocation. Target roles: AI Eval, AI Ops, AI Enablement, Localization.`;


  // --- BUILD THE PROMPT ---
  // The candidate profile grounds every analysis in real resume content pulled from
  // the PROFILE tab rather than a hardcoded string — Gemini returns a JSON object
  // whose keys match sheet column headers exactly so no mapping logic is needed.
  const prompt = `You are analyzing a job posting for a candidate with this background:
${candidateProfile}

Here is the job description:
${jobDescription}

Return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "role_category": "one of: AI Eval, AI Ops, AI Enablement, Localization, Data, PM, Other",
  "fit_score": "integer 0-100",
  "rationale": "one sentence explaining the score",
  "resume_recommended": "one of: AI Eval, AI Ops, Localization, General",
  "resume_edits_suggested": "one short sentence on what to tweak, or blank if none",
  "actual_post_date": "look for any date near words like posted, date, listed, or updated — return as YYYY-MM-DD or empty string"
}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

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
  // falls back to the hardcoded string in analyzeWithGemini() if not found.
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
  // indexOf() returns the zero-based array position of each column name —
  // adding 1 converts to the one-based column number getRange() expects.
  const fitScoreCol = headers.indexOf("fit_score") + 1;
  const jobDescCol = headers.indexOf("job_description") + 1;
  const roleCatCol = headers.indexOf("role_category") + 1;
  const rationaleCol = headers.indexOf("rationale") + 1;
  const resumeRecCol = headers.indexOf("resume_recommended") + 1;
  const resumeEditsCol = headers.indexOf("resume_edits_suggested") + 1;
  const actualPostDateCol = headers.indexOf("actual_post_date") + 1;

  // Read all rows in one call — much faster than reading one row at a time
  // on a large sheet and avoids hitting Apps Script execution time limits.
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var analyzed = 0;
  var skipped = 0;

  allData.forEach(function(row, i) {
    const fitScore = row[fitScoreCol - 1];
    const jobDesc = row[jobDescCol - 1];

    // Only process rows where fit_score is blank AND a job description exists —
    // rows already analyzed or without descriptions are skipped entirely.
    if (fitScore !== "" || !jobDesc) {
      skipped++;
      return;
    }

    Logger.log("Analyzing row " + (i + 2) + ": " + row[headers.indexOf("job_title")]);

    const regexDate = extractPostDate(jobDesc);

    Utilities.sleep(4000); // rate limit buffer between Gemini calls

    const descStart = jobDesc.substring(0, 1500);
    const descEnd = jobDesc.slice(-500);
    const shortDesc = descStart + "\n...\n" + descEnd;
    const analysis = analyzeWithGemini(shortDesc, profileSummary);

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
  // on the full description text and writes the result directly without calling
  // Gemini, so this is fast, free, and safe to run on large numbers of rows.
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

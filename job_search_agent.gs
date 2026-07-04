// ============================================================
// JOB SEARCH AGENT v4 — Google Apps Script
// Full pipeline: profile-driven queries, job fetching,
// Gemini analysis with real resume comparison, deduplication,
// profile refresh, dynamic category management, resume catalog.
//
// HOW TO USE:
//   Run runFullPipeline() daily — fetches new jobs then analyzes them.
//   If it times out during analysis, just run analyzeExistingJobs() again.
//   Run refreshProfile() when you update or add a resume.
//   Run buildResumeCatalog() when you add new resumes to Drive.
//
// Functions:
//   runFullPipeline()      — runs fetchJobsOnly() then analyzeExistingJobs()
//   fetchJobsOnly()        — pulls jobs from JSearch, writes raw rows, no Gemini
//   analyzeExistingJobs()  — analyzes unprocessed rows with Gemini + resume reading
//   analyzeWithGemini()    — scores and categorizes a job description
//   getResumeContent()     — reads a resume Doc from Drive by doc_id
//   refreshProfile()       — updates PROFILE tab from new/changed resume Docs
//   buildResumeCatalog()   — catalogs all resume Docs into RESUMES tab
//   categorizeResume()     — calls Gemini to assign a category to one resume
//   extractPostDate()      — regex helper to find dates in description text
//   backfillPostDates()    — fills actual_post_date for rows missing it
//   getCategories()        — reads CATEGORIES tab, returns array of labels
//   addCategory()          — adds new category to CATEGORIES tab if not present
//   getRoleCategoryInstruction() — builds dynamic category prompt for Gemini
//   getResumesByCategory() — returns resume rows matching a given category
// ============================================================


// ============================================================
// WRAPPER: RUN FULL PIPELINE
// ============================================================

function runFullPipeline() {
  // Runs fetchJobsOnly() first to pull new jobs from JSearch with no AI calls,
  // then immediately runs analyzeExistingJobs() to score and annotate each new
  // row — if analyzeExistingJobs() times out before finishing, just run it again
  // since it skips already-analyzed rows and picks up exactly where it left off.
  Logger.log("=== Starting full pipeline ===");
  fetchJobsOnly();
  Logger.log("=== Fetch complete. Starting analysis ===");
  analyzeExistingJobs();
  Logger.log("=== Pipeline complete ===");
}


// ============================================================
// STEP 1: FETCH JOBS FROM JSEARCH — NO GEMINI CALLS
// ============================================================

function fetchJobsOnly() {

  // --- CONFIG ---
  // NUM_PAGES is per query — with 10 queries at 1 page each that's up to 100
  // jobs per run using 10 JSearch API requests from your 200/month budget.
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

  // --- READ QUERIES FROM PROFILE TAB ---
  // Reads all rows where column A starts with "query_" and collects column B
  // values as the search queries to run — adding or removing queries only
  // requires editing the PROFILE tab, not the code.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName(PROFILE_SHEET_NAME);
  const profileData = profileSheet.getDataRange().getValues();

  var queries = [];
  profileData.forEach(function(row) {
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
  // Reads all existing job_id values into a Set before any queries run — the
  // same job appearing in multiple query results only gets written once, saving
  // JSearch quota and preventing duplicate rows in the sheet.
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
  // and writes raw rows to the sheet — NO Gemini calls here, keeping this fast
  // and well within the 6-minute Apps Script execution limit.
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
      // headers.map() iterates every column header and matches it to the
      // corresponding JSearch field — all agent analysis columns are left blank
      // since analyzeExistingJobs() fills them in the next step.
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
    });

    Logger.log("  → " + newCount + " new jobs written from this query.");
    Utilities.sleep(500); // brief pause between queries
  });

  Logger.log("fetchJobsOnly done. Total new jobs written: " + totalNewCount);
}


// ============================================================
// STEP 2: ANALYZE UNPROCESSED ROWS WITH GEMINI + RESUME READING
// ============================================================

function analyzeExistingJobs() {

  // --- GET SHEET, HEADERS, AND PROFILE SUMMARY ---
  // Reads the PROFILE tab for the current profile summary to ground Gemini —
  // falls back to the hardcoded string in analyzeWithGemini() if missing.
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
  const jobTitleCol = headers.indexOf("job_title") + 1;
  const roleCatCol = headers.indexOf("role_category") + 1;
  const rationaleCol = headers.indexOf("rationale") + 1;
  const resumeRecCol = headers.indexOf("resume_recommended") + 1;
  const resumeEditsCol = headers.indexOf("resume_edits_suggested") + 1;
  const actualPostDateCol = headers.indexOf("actual_post_date") + 1;

// Read all rows in one call then reverse so newest jobs (bottom of sheet)
// get analyzed first — if the function times out the most recent listings
// are already scored rather than only the oldest ones from the top.
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues().reverse;

  var analyzed = 0;
  var skipped = 0;

  // --- LOOP THROUGH ROWS AND FILL BLANKS ---
  // Only processes rows where fit_score is empty AND job_description exists —
  // rows already analyzed or missing a description are skipped entirely,
  // making this safe to run multiple times without duplicating any work.
  allData.forEach(function(row, i) {
    const fitScore = row[fitScoreCol - 1];
    const jobDesc = row[jobDescCol - 1];

    if (fitScore !== "" || !jobDesc) {
      skipped++;
      return;
    }

    Logger.log("Analyzing row " + (i + 2) + ": " + row[jobTitleCol - 1]);

    // --- EXTRACT DATE BEFORE TRUNCATION ---
    // Runs regex on the full description first so posting dates near the end
    // aren't lost when we truncate for Gemini below.
    const regexDate = extractPostDate(jobDesc);

    // --- STEP 1: SCORE AND CATEGORIZE THE JOB ---
    // Sends the smart-truncated description (start + end) to Gemini for scoring —
    // first 1500 chars captures intro and responsibilities, last 500 captures
    // requirements and salary that often appear at the bottom of postings.
    Utilities.sleep(4000);
    const shortDesc = jobDesc.substring(0, 1500) + "\n...\n" + jobDesc.slice(-500);
    const analysis = analyzeWithGemini(shortDesc, profileSummary);

    // Auto-add any new role_category to the CATEGORIES tab so it's available
    // for future job and resume categorization calls.
    if (analysis.role_category) addCategory(analysis.role_category, "job");

    // --- STEP 2: FIND BEST RESUME AND READ ITS CONTENT ---
    // Uses the role_category assigned by Gemini to look up matching resumes
    // in the RESUMES tab, then reads the actual Doc content from Drive —
    // this gives Gemini real resume text to compare against the job description
    // rather than guessing based on category labels alone.
    var resumeContent = "";
    var recommendedResumeName = analysis.resume_recommended || "";

    var matchingResumes = getResumesByCategory(analysis.role_category || "");

    if (matchingResumes.length > 0) {
      // Use the first matching resume — best candidate given category alignment
      var bestResume = matchingResumes[0];
      recommendedResumeName = bestResume.filename;
      resumeContent = getResumeContent(bestResume.doc_id);
      Logger.log("  Reading resume: " + recommendedResumeName);
    }

    // --- STEP 3: GET TARGETED EDIT SUGGESTIONS ---
    // Only runs if we successfully read a resume — sends both the job description
    // and the actual resume content to Gemini and asks for specific gap analysis,
    // returning edits grounded in what's actually missing rather than guesses.
    var resumeEdits = analysis.resume_edits_suggested || "";

    if (resumeContent) {
      Utilities.sleep(4000);
      resumeEdits = getResumeEditSuggestions(shortDesc, resumeContent, recommendedResumeName);
    }

    // --- WRITE ALL RESULTS TO THE ROW ---
    // Writes every analysis field back to the correct column in this row —
    // uses column position found by name so order in the sheet doesn't matter.
    const rowNumber = i + 2;
    sheet.getRange(rowNumber, roleCatCol).setValue(analysis.role_category || "");
    sheet.getRange(rowNumber, fitScoreCol).setValue(analysis.fit_score || "");
    sheet.getRange(rowNumber, rationaleCol).setValue(analysis.rationale || "");
    sheet.getRange(rowNumber, resumeRecCol).setValue(recommendedResumeName);
    sheet.getRange(rowNumber, resumeEditsCol).setValue(resumeEdits);
    sheet.getRange(rowNumber, actualPostDateCol).setValue(regexDate || analysis.actual_post_date || "");

    Logger.log("  Score: " + analysis.fit_score + " | Category: " + analysis.role_category + " | Resume: " + recommendedResumeName);

    analyzed++;
  });

  Logger.log("analyzeExistingJobs done. Analyzed: " + analyzed + " | Skipped: " + skipped);
}


// ============================================================
// GEMINI: SCORE AND CATEGORIZE A JOB DESCRIPTION
// ============================================================

function analyzeWithGemini(jobDescription, profileSummary) {

  // --- GET GEMINI API KEY ---
  // Retrieved from Script Properties — never written into the code itself.
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");

  // --- CANDIDATE PROFILE ---
  // fetchJobs() and analyzeExistingJobs() both pass the live profile summary
  // from the PROFILE tab — the hardcoded fallback ensures the function still
  // works even if the PROFILE tab is missing or empty.
  const candidateProfile = profileSummary || `8+ years experience in AI quality operations,
LLM evaluation, annotation pipeline management, and training design. Skills include rubric
development, adversarial prompt testing, KPI dashboard design, Power Automate, Python.
Languages: C1 French, conversational German. Prefers remote, based in Kalamazoo MI,
open to EU relocation. Target roles: AI Eval, AI Ops, AI Enablement, Localization.`;

  // --- DYNAMIC CATEGORY INSTRUCTION ---
  // Built before the prompt so it can be interpolated cleanly as a variable —
  // reads the CATEGORIES tab to keep job categorization consistent with resume
  // categorization since both draw from the same list.
  const categoryInstruction = getRoleCategoryInstruction();

  // --- BUILD THE PROMPT ---
  // Returns a JSON object whose keys match sheet column headers exactly —
  // no translation needed between what Gemini returns and where it gets written.
  const prompt = `You are analyzing a job posting for a candidate with this background:
${candidateProfile}

Here is the job description:
${jobDescription}

Return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "role_category": "${categoryInstruction}",
  "fit_score": "integer 0-100 based on how well the candidate's background matches this role",
  "rationale": "one sentence explaining the fit score",
  "resume_recommended": "a short label describing the best resume type for this role",
  "resume_edits_suggested": "leave blank — this will be filled separately after reading the actual resume",
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
  // never crashes mid-run; failed rows will be retried on the next run.
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
// GEMINI: GET TARGETED RESUME EDIT SUGGESTIONS
// ============================================================

function getResumeEditSuggestions(jobDescription, resumeContent, resumeName) {
  // Sends both the job description and the actual resume content to Gemini,
  // asking it to compare the two and return specific, actionable edit suggestions
  // based on real gaps — not guesses about what might be missing.
  // This is a separate Gemini call from analyzeWithGemini() so the two tasks
  // stay clean and the suggestions are always grounded in real resume content.

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");

  const prompt = `You are a career advisor comparing a job description against a candidate's resume to find specific gaps and improvement opportunities.

Resume name: ${resumeName}

Resume content:
${resumeContent.substring(0, 2000)}

Job description:
${jobDescription}

Based on comparing these two documents, provide specific and actionable edit suggestions.
Focus on:
- Skills or keywords in the job description that are missing from the resume
- Existing resume content that should be reframed to better match the job's language
- Specific bullets that should be added, removed, or reworded
- Anything the resume does well that directly matches this role

Return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "gaps": "comma-separated list of skills or keywords in the job description missing from the resume",
  "reframe": "one specific bullet or section to reframe and how",
  "add": "one specific thing to add to the resume for this role",
  "strength": "one thing already in the resume that directly matches this job",
  "summary": "one concise sentence summarizing the most important edit to make"
}`;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const text = json.candidates[0].content.parts[0].text;
    const cleaned = text.replace(/```json|```/g, "").trim();
    const result = JSON.parse(cleaned);
    // Return the summary as the cell value — full JSON stored for future use
    return result.summary || "";
  } catch(e) {
    Logger.log("Resume edit suggestion error: " + e);
    return "";
  }
}


// ============================================================
// HELPER: READ RESUME CONTENT FROM GOOGLE DRIVE
// ============================================================

function getResumeContent(docId) {
  // Opens a Google Doc by its ID and returns the full plain text content —
  // called by analyzeExistingJobs() to read the actual resume before sending
  // it to Gemini for gap analysis. Returns empty string on any error so the
  // script never crashes if a Doc is deleted or access is revoked.
  if (!docId) return "";
  try {
    var doc = DocumentApp.openById(docId);
    return doc.getBody().getText();
  } catch(e) {
    Logger.log("Could not read resume Doc " + docId + ": " + e);
    return "";
  }
}


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

  if (lastRow < 2) { Logger.log("No data rows found."); return; }

  const jobDescCol = headers.indexOf("job_description") + 1;
  const actualPostDateCol = headers.indexOf("actual_post_date") + 1;
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var updated = 0;
  var skipped = 0;

  allData.forEach(function(row, i) {
    const jobDesc = row[jobDescCol - 1];
    const existingDate = row[actualPostDateCol - 1];

    if (!jobDesc || existingDate !== "") { skipped++; return; }

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
// HELPER: READ ALL EXISTING CATEGORIES FROM CATEGORIES TAB
// ============================================================

function getCategories() {
  // Opens the CATEGORIES tab and returns a flat array of all category label
  // strings — passed to Gemini so it picks a consistent existing label before
  // creating something new. Creates the tab with headers if it doesn't exist.
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
    if (i === 0) return;
    if (row[0]) categories.push(row[0]);
  });
  return categories;
}


// ============================================================
// HELPER: ADD NEW CATEGORY TO CATEGORIES TAB IF NOT PRESENT
// ============================================================

function addCategory(category, source) {
  // Checks whether the category already exists before adding — prevents
  // duplicates even when Gemini returns something already in the list.
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
  // instruction string for the Gemini prompt — keeping job categorization
  // consistent with resume categorization since both draw from the same list.
  var categories = getCategories();
  if (categories.length > 0) {
    return "pick the closest match from: " + categories.join(", ") + " — or create a concise new 2-4 word label if none fit";
  } else {
    return "assign a concise 2-4 word category label that describes the target role";
  }
}


// ============================================================
// HELPER: GET RESUMES BY CATEGORY FOR JOB MATCHING
// ============================================================

function getResumesByCategory(category) {
  // Reads the RESUMES tab and returns all rows where the category matches
  // the given label — called by analyzeExistingJobs() so Gemini only reads
  // relevant resume Docs rather than all 30+ files.
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


// ============================================================
// PROFILE REFRESH: READ NEW RESUMES FROM DRIVE, UPDATE PROFILE TAB
// ============================================================

function refreshProfile() {

  // --- CONFIG ---
  // RESUME_FOLDER_ID: your Resume folder ID from the Drive URL
  // MAX_CHARS_PER_RESUME: truncation limit per resume to manage token costs —
  // only new/updated Docs are read so the total token cost stays low per run.
  const RESUME_FOLDER_ID = "1FKMiQkLfq2rcDFRnplIS7kUbtms92E0I";
  const PROFILE_SHEET_NAME = "PROFILE";
  const MAX_CHARS_PER_RESUME = 3000;

  // --- GET PROFILE SHEET AND READ LAST UPDATED TIMESTAMP ---
  // The last_updated value in B2 tells the script which resumes to skip —
  // only Docs modified after this date get re-read, keeping repeat runs cheap.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName(PROFILE_SHEET_NAME);
  const lastUpdatedCell = profileSheet.getRange("B2").getValue();
  const lastUpdated = lastUpdatedCell ? new Date(lastUpdatedCell) : new Date(0);
  Logger.log("Profile last updated: " + lastUpdated);

  // --- FIND NEW OR UPDATED RESUME DOCS IN DRIVE ---
  // Filters to only Google Docs modified after the last_updated timestamp —
  // PDFs and Word docs are automatically skipped without extra handling.
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
    resumeTexts.push("=== " + file.getName() + " ===\n" + doc.getBody().getText().substring(0, MAX_CHARS_PER_RESUME));
  }

  if (resumeTexts.length === 0) {
    Logger.log("No new or updated resumes found. Profile is up to date.");
    return;
  }

  Logger.log("Found " + resumeTexts.length + " new/updated resumes. Sending to Gemini.");

  // --- READ EXISTING PROFILE SUMMARY ---
  // Passes the existing summary alongside only the changed resumes — Gemini
  // merges new info rather than rewriting everything from scratch.
  const existingSummary = profileSheet.getRange("B3").getValue() || "";
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");

  const prompt = `You are updating a job seeker's career profile based on new or updated resume content.

Here is the existing profile summary:
${existingSummary}

Here are the new or updated resumes to incorporate:
${resumeTexts.join("\n\n")}

Based on all of this, return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "profile_summary": "updated 3-5 sentence profile summary — preserve existing accurate information, add or update anything new from the resumes",
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
- Each query targets a different angle of the candidate's background
- Natural language job board search strings
- Include remote in most queries
- Cover: LLM evaluation, AI quality ops, AI enablement/training, localization, prompt engineering
- Specific enough to surface relevant roles`;

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
    Logger.log("Gemini parse error in refreshProfile: " + e);
    return;
  }

  // --- WRITE UPDATED PROFILE BACK TO SHEET ---
  // Overwrites the PROFILE tab with new summary, queries, and timestamp —
  // the next fetchJobsOnly() run uses the new queries automatically.
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
// RESUME CATALOG: BUILD AND MAINTAIN RESUMES TAB
// ============================================================

function buildResumeCatalog() {

  // --- CONFIG ---
  // CUTOFF_YEAR automatically excludes pre-2023 resumes — anything last modified
  // before this year is skipped so old unrelated resumes don't pollute the catalog.
  // MAX_CHARS: how much of each resume Gemini reads for categorization —
  // 500 characters identifies the role focus without burning tokens.
  const RESUME_FOLDER_ID = "1FKMiQkLfq2rcDFRnplIS7kUbtms92E0I";
  const RESUMES_SHEET_NAME = "RESUMES";
  const CUTOFF_YEAR = 2023;
  const MAX_CHARS = 500;

  // --- FILES TO EXCLUDE ---
  // Non-resume files in the folder — job applications, interview prep docs,
  // and old translation files are all excluded by filename matching.
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
  // Reads all doc_id values already cataloged into a Set so only new files
  // get a Gemini call — preserves manual category corrections on existing rows.
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

    if (file.getMimeType() !== "application/vnd.google-apps.document") { skipped++; continue; }

    var fileName = file.getName();

    var isExcluded = EXCLUDE_NAMES.some(function(excluded) {
      return fileName.toLowerCase().indexOf(excluded.toLowerCase()) !== -1;
    });
    if (isExcluded) { Logger.log("Excluding: " + fileName); skipped++; continue; }

    var lastModified = file.getLastUpdated();
    if (lastModified.getFullYear() < CUTOFF_YEAR) {
      Logger.log("Skipping old file (" + lastModified.getFullYear() + "): " + fileName);
      skipped++;
      continue;
    }

    var docId = file.getId();
    var docUrl = file.getUrl();
    var lastModifiedStr = lastModified.toISOString().substring(0, 10);

    // If already cataloged, update timestamp only — preserve manual corrections
    if (existingIds.has(docId)) {
      resumeSheet.getRange(existingRowNums[docId], 4).setValue(lastModifiedStr);
      updated++;
      Logger.log("Updated timestamp: " + fileName);
      continue;
    }

    // --- READ AND CATEGORIZE NEW RESUME ---
    // Opens the Doc, reads the first MAX_CHARS, sends to Gemini for categorization,
    // auto-adds any new category to the CATEGORIES tab, then writes the catalog row.
    var doc = DocumentApp.openById(docId);
    var resumeText = doc.getBody().getText().substring(0, MAX_CHARS);
    var category = categorizeResume(resumeText, fileName, categories, apiKey);

    addCategory(category, "resume");
    if (categories.indexOf(category) === -1) categories.push(category);

    resumeSheet.appendRow([fileName, docId, category, lastModifiedStr, docUrl]);
    existingIds.add(docId);
    added++;

    Logger.log("Added: " + fileName + " → " + category);
    Utilities.sleep(2000);
  }

  Logger.log("buildResumeCatalog done. Added: " + added + " | Updated: " + updated + " | Skipped: " + skipped);
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

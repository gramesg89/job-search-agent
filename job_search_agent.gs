// ============================================================
// JOB SEARCH AGENT — Google Apps Script
// Fetches jobs from JSearch, analyzes them with Gemini,
// and writes structured results to a Google Sheet.
// ============================================================


function fetchJobs() {

  // --- CONFIG ---
  // These are the main settings you'll change regularly. QUERY controls what jobs
  // are searched; NUM_PAGES controls how many results come back (each page = 10 jobs
  // and costs 1 API request against your 200/month JSearch budget); DATE_POSTED
  // filters by recency so you're not seeing stale listings.
  const QUERY = "LLM evaluation AI quality operations remote";
  const SHEET_NAME = "RAW";
  const NUM_PAGES = 3;
  const DATE_POSTED = "week"; // options: "all", "today", "3days", "week", "month"


  // --- GET JSEARCH API KEY ---
  // PropertiesService is Apps Script's secure key-value store — it retrieves your
  // JSearch API key by the name you saved it under, so the actual key value never
  // appears anywhere in the code itself.
  const apiKey = PropertiesService.getScriptProperties().getProperty("JSEARCH_KEY");
  if (!apiKey) {
    Logger.log("ERROR: No API key found. Add JSEARCH_KEY to Script Properties.");
    return;
  }


  // --- BUILD AND SEND THE JSEARCH REQUEST ---
  // UrlFetchApp is Apps Script's built-in HTTP client — it constructs a GET request
  // to JSearch's search-v2 endpoint with your query and filters baked into the URL,
  // sends it, and returns the raw response text which we then parse into a JavaScript
  // object we can loop through.
  const url = "https://jsearch.p.rapidapi.com/search-v2"
    + "?query=" + encodeURIComponent(QUERY)
    + "&num_pages=" + NUM_PAGES
    + "&country=us"
    + "&date_posted=" + DATE_POSTED;

  const options = {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      "X-RapidAPI-Key": apiKey,
      "X-RapidAPI-Host": "jsearch.p.rapidapi.com"
    },
    muteHttpExceptions: true // prevents the script from crashing on API errors; we handle them manually below
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log("Raw response (first 500 chars): " + response.getContentText().substring(0, 500));
  const json = JSON.parse(response.getContentText());

  if (json.status !== "OK") {
    Logger.log("ERROR: Bad response from JSearch: " + response.getContentText());
    return;
  }


  // --- EXTRACT JOBS FROM RESPONSE ---
  // JSearch v2 nests the jobs array inside data.jobs rather than directly in data —
  // this pulls out just the array of job objects we actually need to loop through.
  const jobs = json.data.jobs;
  Logger.log("Jobs returned: " + jobs.length);


  // --- GET SHEET AND READ HEADER ROW ---
  // getSheetByName opens the specific tab by name, then getRange reads the entire
  // first row into a flat array — this lets the script find any column by its exact
  // header name later using indexOf(), so column order in the sheet never matters.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];


  // --- BUILD DEDUPLICATION SET ---
  // Before writing anything, this reads all existing job_id values from the sheet
  // into a Set (a collection of unique values) so we can instantly check whether
  // each incoming job already exists — if it does, we skip it entirely including
  // the Gemini API call, saving both quota and tokens.
  const lastRow = sheet.getLastRow();
  let existingIds = new Set();
  if (lastRow > 1) {
    const jobIdCol = headers.indexOf("job_id") + 1;
    const existingIdValues = sheet.getRange(2, jobIdCol, lastRow - 1, 1).getValues();
    existingIdValues.forEach(function(row) {
      if (row[0]) existingIds.add(row[0]);
    });
  }


  // --- LOOP THROUGH JOBS, WRITE TO SHEET, AND ANALYZE ---
  // For each job returned by JSearch, this checks for duplicates first, then maps
  // each JSearch field to its matching column header and appends the row, then
  // immediately calls Gemini to fill in the analysis columns before moving on —
  // so every new row is fully populated by the time the script finishes.
  let newCount = 0;

  jobs.forEach(function(job) {

    // Skip this job entirely if its ID already exists in the sheet
    if (existingIds.has(job.job_id)) {
      Logger.log("Skipping duplicate: " + job.job_title + " at " + job.employer_name);
      return;
    }


    // --- MAP JSEARCH FIELDS TO COLUMN HEADERS ---
    // headers.map() iterates through every column header and uses a switch statement
    // to look up the matching value from the JSearch job object — because column
    // names match JSearch field names exactly, there's no translation layer needed,
    // and adding a new column just means adding a new case to this switch.
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
        // Agent analysis columns — left blank here because Gemini fills them immediately after appendRow
        case "role_category":             return "";
        case "fit_score":                 return "";
        case "rationale":                 return "";
        case "resume_recommended":        return "";
        case "resume_edits_suggested":    return "";
        case "actual_post_date":          return ""; // Gemini extracts this from the description text below
        // Human-managed columns — sensible defaults set here, you update these manually
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


    // --- GEMINI ANALYSIS ---
    // After writing the raw JSearch row, this waits 4 seconds (to stay under Gemini's
    // free tier rate limit of ~15 requests/minute), then sends a truncated version of
    // the job description to Gemini — 2000 characters captures the most important
    // signals while using roughly a third of the tokens a full description would cost.
    // Gemini returns a JSON object whose keys match column headers exactly, so each
    // value gets written to the right cell using indexOf() to find the column by name.
    Utilities.sleep(4000);
    const newRowNumber = sheet.getLastRow();
    const shortDesc = (job.job_description || "").substring(0, 2000);
    const analysis = analyzeWithGemini(shortDesc);

    sheet.getRange(newRowNumber, headers.indexOf("role_category") + 1).setValue(analysis.role_category || "");
    sheet.getRange(newRowNumber, headers.indexOf("fit_score") + 1).setValue(analysis.fit_score || "");
    sheet.getRange(newRowNumber, headers.indexOf("rationale") + 1).setValue(analysis.rationale || "");
    sheet.getRange(newRowNumber, headers.indexOf("resume_recommended") + 1).setValue(analysis.resume_recommended || "");
    sheet.getRange(newRowNumber, headers.indexOf("resume_edits_suggested") + 1).setValue(analysis.resume_edits_suggested || "");
    sheet.getRange(newRowNumber, headers.indexOf("actual_post_date") + 1).setValue(analysis.actual_post_date || "");

    Logger.log("Analyzed: " + job.job_title + " | Score: " + analysis.fit_score + " | Posted: " + analysis.actual_post_date);
  });

  Logger.log("Done. " + newCount + " new jobs added.");
}


// ============================================================

function analyzeWithGemini(jobDescription) {

  // --- GET GEMINI API KEY ---
  // Same secure retrieval pattern as JSearch — the actual key value lives in Script
  // Properties and is never written into the code itself.
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");


  // --- BUILD THE PROMPT ---
  // This constructs the full instruction sent to Gemini on every call — the candidate
  // background at the top is permanent context that grounds every analysis in who you
  // actually are, while the job description is swapped in dynamically each time.
  // Gemini is instructed to return ONLY a raw JSON object with keys that match your
  // sheet column headers exactly, so the script can write each value by column name
  // with no translation or mapping needed.
  const prompt = `You are analyzing a job posting for a candidate with this background:
- 8+ years experience across AI quality operations, LLM evaluation, annotation
  pipeline management, and training design
- Built and managed rubric-based evaluation systems, adversarial prompt testing,
  and quality KPI dashboards reducing report time from 5-6 hours to under 2
- Experience: AI trainer/quality lead at Innodata, adversarial red teaming at
  Invisible Technologies, freelance QA at Gengo/Lionbridge, Interpretation PM
  at TransPerfect
- Skills: prompt engineering, rubric development, LLM evaluation, Python,
  Power Automate, SharePoint, training architecture, cross-functional QA
- Languages: C1 French, conversational German, readable Spanish, IPA/X-SAMPA
- Target roles: AI Eval, AI Ops, AI Quality, AI Enablement, Localization
- Strongly prefers remote; based in Kalamazoo MI; open to EU relocation
- Ideal salary: $70,000-$100,000 USD; open to international roles

Here is the job description:
${jobDescription}

Return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "role_category": "one of: AI Eval, AI Ops, AI Enablement, Localization, Data, PM, Other",
  "fit_score": "integer 0-100",
  "rationale": "one sentence explaining the score",
  "resume_recommended": "one of: AI Eval, AI Ops, Localization, General",
  "resume_edits_suggested": "one short sentence on what to tweak, or blank if none",
  "actual_post_date": "if the job description text explicitly states a posting date, return it as YYYY-MM-DD — otherwise return empty string"
}`;


  // --- BUILD AND SEND THE GEMINI REQUEST ---
  // The Gemini REST API expects a specific nested structure — contents contains an
  // array of message objects, each with a parts array containing the actual text.
  // UrlFetchApp sends it as a POST request with the payload serialized to JSON,
  // and muteHttpExceptions lets us handle API errors gracefully instead of crashing.
  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }]
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


  // --- PARSE GEMINI'S RESPONSE ---
  // Gemini returns its answer buried inside a nested structure (candidates → content
  // → parts → text), so this drills down to extract just the JSON string, strips any
  // accidental markdown backticks Gemini sometimes adds despite instructions, then
  // parses it into a JavaScript object we can read field by field. The try/catch
  // returns a safe empty object if anything goes wrong so the script never crashes
  // mid-run — the row stays in the sheet and can be retried with analyzeExistingJobs().
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
      actual_post_date: "" // empty fallback so the script doesn't crash on a failed analysis
    };
  }
}


// ============================================================

function analyzeExistingJobs() {

  // --- GET SHEET AND HEADERS ---
  // Opens the RAW tab and reads the entire header row into an array so every column
  // can be found by name throughout this function — same pattern as fetchJobs(),
  // meaning column order in the sheet never matters as long as header names match.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RAW");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No data rows found.");
    return;
  }


  // --- FIND RELEVANT COLUMN POSITIONS ---
  // indexOf() searches the headers array for each column name and returns its
  // zero-based position — adding 1 converts it to the one-based column number
  // that getRange() expects when writing back to specific cells.
  const fitScoreCol = headers.indexOf("fit_score") + 1;
  const jobDescCol = headers.indexOf("job_description") + 1;
  const roleCatCol = headers.indexOf("role_category") + 1;
  const rationaleCol = headers.indexOf("rationale") + 1;
  const resumeRecCol = headers.indexOf("resume_recommended") + 1;
  const resumeEditsCol = headers.indexOf("resume_edits_suggested") + 1;
  const actualPostDateCol = headers.indexOf("actual_post_date") + 1;


  // --- READ ALL ROWS AT ONCE ---
  // getRange with the full data dimensions pulls every row into memory in a single
  // API call — this is significantly faster than reading one row at a time, which
  // would hit Apps Script's execution limits quickly on a large sheet.
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let analyzed = 0;
  let skipped = 0;


  // --- LOOP THROUGH ROWS AND FILL BLANKS ---
  // For each row this checks two conditions: fit_score is empty (analysis hasn't run)
  // AND job_description exists (there's something to analyze) — if both are true it
  // sends the truncated description to Gemini and writes all returned fields back into
  // that specific row. Rows already analyzed or missing a description are skipped.
  allData.forEach(function(row, i) {
    const fitScore = row[fitScoreCol - 1];
    const jobDesc = row[jobDescCol - 1];

    if (fitScore !== "" || !jobDesc) {
      skipped++;
      return;
    }

    Logger.log("Analyzing row " + (i + 2) + ": " + row[headers.indexOf("job_title")]);

    // Wait 4 seconds between calls to stay under Gemini's free tier rate limit —
    // without this pause, back-to-back calls on a large sheet will hit the quota error.
    Utilities.sleep(4000);

    const shortDesc = jobDesc.substring(0, 2000);
    const analysis = analyzeWithGemini(shortDesc);

    const rowNumber = i + 2;
    sheet.getRange(rowNumber, roleCatCol).setValue(analysis.role_category || "");
    sheet.getRange(rowNumber, fitScoreCol).setValue(analysis.fit_score || "");
    sheet.getRange(rowNumber, rationaleCol).setValue(analysis.rationale || "");
    sheet.getRange(rowNumber, resumeRecCol).setValue(analysis.resume_recommended || "");
    sheet.getRange(rowNumber, resumeEditsCol).setValue(analysis.resume_edits_suggested || "");
    sheet.getRange(rowNumber, actualPostDateCol).setValue(analysis.actual_post_date || "");

    analyzed++;
  });

  Logger.log("Done. Analyzed: " + analyzed + " | Skipped (already done): " + skipped);
}

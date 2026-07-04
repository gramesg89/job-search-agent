function fetchJobs() {

  // --- CONFIG ---
  const QUERY = "LLM evaluation AI quality operations remote";
  const SHEET_NAME = "RAW"; // change if your tab has a different name
  const NUM_PAGES = 1;         // increase later for more results

  // --- GET API KEY ---
  const apiKey = PropertiesService.getScriptProperties().getProperty("JSEARCH_KEY");
  if (!apiKey) {
    Logger.log("ERROR: No API key found. Add JSEARCH_KEY to Script Properties.");
    return;
  }

  // --- CALL JSEARCH ---
  const url = "https://jsearch.p.rapidapi.com/search-v2"
    + "?query=" + encodeURIComponent(QUERY)
    + "&num_pages=" + NUM_PAGES
    + "&country=us"
    + "&date_posted=all";

  const options = {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      "X-RapidAPI-Key": apiKey,
      "X-RapidAPI-Host": "jsearch.p.rapidapi.com"
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log("Raw response: " + response.getContentText());
  const json = JSON.parse(response.getContentText());

  if (json.status !== "OK") {
    Logger.log("ERROR: Bad response from JSearch: " + response.getContentText());
    return;
  }

  const jobs = json.data.jobs;
  Logger.log("Jobs returned: " + jobs.length);

  // --- GET SHEET AND HEADERS ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // --- DEDUPLICATION: build set of existing job_ids ---
  const lastRow = sheet.getLastRow();
  let existingIds = new Set();
  if (lastRow > 1) {
    const jobIdCol = headers.indexOf("job_id") + 1;
    const existingIdValues = sheet.getRange(2, jobIdCol, lastRow - 1, 1).getValues();
    existingIdValues.forEach(function(row) {
      if (row[0]) existingIds.add(row[0]);
    });
  }

  // --- WRITE ROWS ---
  let newCount = 0;

  jobs.forEach(function(job) {

    // skip duplicates
    if (existingIds.has(job.job_id)) {
      Logger.log("Skipping duplicate: " + job.job_title + " at " + job.employer_name);
      return;
    }

    // map job fields to your column headers
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
        // custom fields — blank for now, AI fills these later
        case "role_category":             return "";
        case "fit_score":                 return "";
        case "rationale":                 return "";
        case "resume_recommended":        return "";
        case "resume_edits_suggested":    return "";
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
    const newRowNumber = sheet.getLastRow();
    const analysis = analyzeWithGemini(job.job_description || "");

    sheet.getRange(newRowNumber, headers.indexOf("role_category") + 1).setValue(analysis.role_category || "");
    sheet.getRange(newRowNumber, headers.indexOf("fit_score") + 1).setValue(analysis.fit_score || "");
    sheet.getRange(newRowNumber, headers.indexOf("rationale") + 1).setValue(analysis.rationale || "");
    sheet.getRange(newRowNumber, headers.indexOf("resume_recommended") + 1).setValue(analysis.resume_recommended || "");
    sheet.getRange(newRowNumber, headers.indexOf("resume_edits_suggested") + 1).setValue(analysis.resume_edits_suggested || "");

    Logger.log("Analyzed: " + job.job_title + " | Score: " + analysis.fit_score);
  });

  Logger.log("Done. " + newCount + " new jobs added.");
}

function analyzeWithGemini(jobDescription) {

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");
  
  const prompt = `You are analyzing a job posting for a candidate with this background:
- 8+ years experience in AI quality, LLM evaluation, annotation pipeline management
- Skills: prompt engineering, rubric development, adversarial testing, Python, training design
- Target roles: AI Eval, AI Ops, AI Quality, AI Enablement, Localization
- Prefers remote work; based in Michigan
- Strong linguistics background (French, German, IPA)
 with over 8 years of experience bridging linguistic precision, traditional process engineering, and Generative AI operations. His background includes scaling remote quality assurance teams, managing complex vendor capacities, and applying deep linguistic knowledge to AI model alignment and evaluation.
Core AI & Technical Skills: LLM Evaluation, Prompt Engineering & Optimization, Adversarial Prompting (Red-Teaming), RLHF, Model Alignment, Anomaly Detection, Python, Power Automate, and Microsoft Copilot Studio.
Quality & Operations Skills: Root-Cause Analysis, KPI Dashboard Design, SOP & Training Documentation, Calibration Reviews, Inter-rater Variance Reduction, and Agile Workflows.
Linguistics & Localization: Native English, C1 Professional French, IPA, X-SAMPA, translation quality control, and medical interpretation project management.

Here is the job description:
${jobDescription}

Return ONLY a JSON object with no markdown, no backticks, no explanation:
{
  "role_category": "one of: AI Eval, AI Ops, AI Enablement, Localization, Data, PM, Other",
  "fit_score": "integer 0-100",
  "rationale": "one sentence explaining the score",
  "resume_recommended": "one of: AI Eval, AI Ops, Localization, General",
  "resume_edits_suggested": "one short sentence on what to tweak, or leave blank if none"
}`;

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
      resume_edits_suggested: ""
    };
  }
}

function analyzeExistingJobs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RAW");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No data rows found.");
    return;
  }

  const fitScoreCol = headers.indexOf("fit_score") + 1;
  const jobDescCol = headers.indexOf("job_description") + 1;
  const roleCatCol = headers.indexOf("role_category") + 1;
  const rationaleCol = headers.indexOf("rationale") + 1;
  const resumeRecCol = headers.indexOf("resume_recommended") + 1;
  const resumeEditsCol = headers.indexOf("resume_edits_suggested") + 1;

  // Get all data at once
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let analyzed = 0;
  let skipped = 0;

  allData.forEach(function(row, i) {
    const fitScore = row[fitScoreCol - 1];
    const jobDesc = row[jobDescCol - 1];

    // Only analyze rows where fit_score is blank and job_description exists
    if (fitScore !== "" || !jobDesc) {
      skipped++;
      return;
    }

    Logger.log("Analyzing row " + (i + 2) + ": " + row[headers.indexOf("job_title")]);

    Utilities.sleep(4000);

    const shortDesc = jobDesc.substring(0, 2000);
    const analysis = analyzeWithGemini(shortDesc);

    const rowNumber = i + 2;
    sheet.getRange(rowNumber, roleCatCol).setValue(analysis.role_category || "");
    sheet.getRange(rowNumber, fitScoreCol).setValue(analysis.fit_score || "");
    sheet.getRange(rowNumber, rationaleCol).setValue(analysis.rationale || "");
    sheet.getRange(rowNumber, resumeRecCol).setValue(analysis.resume_recommended || "");
    sheet.getRange(rowNumber, resumeEditsCol).setValue(analysis.resume_edits_suggested || "");

    analyzed++;
  });

  Logger.log("Done. Analyzed: " + analyzed + " | Skipped (already done): " + skipped);
}

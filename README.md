# Job Search Agent
> An automated, AI-powered job search pipeline that finds, analyzes, and tracks roles matched to my background — without me having to manually search.

Built by [gramesg89](https://github.com/gramesg89)

---

## What This Is

Most job search tools make you do the searching. This one does it for you.

Instead of manually querying job boards and copy-pasting listings into a spreadsheet, this pipeline automatically pulls relevant jobs from across the web, scores them against my actual background using an LLM, recommends which resume variant to use, and logs everything to a structured Google Sheet — every day, without me touching it.

The goal is to replace the manual parts of job searching with an intelligent agent that knows who I am, what I'm looking for, and how to find it.

---

## Architecture

```
JSearch API (aggregates Google Jobs, LinkedIn, Indeed, etc.)
        ↓
Google Apps Script (orchestration layer)
        ↓
Gemini API (fit scoring, role classification, resume recommendation)
        ↓
Google Sheets (structured database / agent brain)
```

---

## What It Does Currently

- **Pulls jobs automatically** from JSearch using targeted queries, filtered to the last 7 days
- **Deduplicates results** — the same job is never written twice, even across multiple runs
- **Analyzes each job description** with Gemini, which returns:
  - `fit_score` — 0-100 match score against my background
  - `role_category` — AI Eval, AI Ops, AI Enablement, Localization, Data, PM, or Other
  - `rationale` — one sentence explaining the score
  - `resume_recommended` — which resume variant to use
  - `resume_edits_suggested` — what to tweak before applying
  - `actual_post_date` — extracted from the job description text if present, since Google Jobs indexing dates are unreliable
- **Writes structured rows** to a Google Sheet with columns mapped directly to JSearch field names for zero-translation data flow
- **Backfill function** — a separate `analyzeExistingJobs()` function re-runs Gemini analysis on any rows missing scores, useful when the API rate limits mid-run

---

## Sheet Schema

The Google Sheet is the agent's brain. Columns are split into three groups:

| Group | Columns | Filled by |
|---|---|---|
| JSearch fields | `job_id`, `job_title`, `employer_name`, `job_apply_link`, `work_arrangement`, `job_min_salary`, `job_max_salary`, `job_description`, and more | Apps Script (automatic) |
| Agent analysis | `role_category`, `fit_score`, `rationale`, `resume_recommended`, `resume_edits_suggested`, `actual_post_date` | Gemini API (automatic) |
| Human fields | `status`, `date_applied`, `interview_stage`, `follow_up_needed`, `notes` | Me (manual) |

Column names match JSearch JSON field names exactly — this means the script maps data to the sheet with no translation layer, and Gemini's JSON output keys match column headers directly.

---

## Stack

| Tool | Role |
|---|---|
| Google Apps Script | Orchestration, scheduling, API calls, sheet writes |
| JSearch API (RapidAPI) | Job data aggregation across Google Jobs, LinkedIn, Indeed |
| Gemini API | LLM analysis, fit scoring, resume recommendation |
| Google Sheets | Structured job database and agent memory |

---

## What's Planned Next

- [ ] **Resume-driven query generation** — instead of hardcoded search queries, Gemini reads my resume variants from Google Drive and generates optimal queries automatically based on my actual skills and experience
- [ ] **Scheduled trigger** — run `fetchJobs()` automatically every morning via Apps Script time-based trigger
- [ ] **Multiple query support** — run several targeted queries per session, deduplicate across all of them
- [ ] **Chat interface** — conversational layer on top of the sheet so I can ask questions like "show me every remote AI Eval role above 80% fit" or "which jobs should I follow up on this week"
- [ ] **Dashboard** — visual summary of pipeline activity, fit score distributions, role category breakdown, and application outcome trends
- [ ] **NotebookLM / Drive grounding** — ground Gemini's resume recommendations in actual resume content rather than label matching

---

## Background and Motivation

I've been applying to AI-adjacent roles — AI Evaluation, AI Ops, AI Enablement, Localization — for over two years. In that time I've tracked 140+ applications manually in a spreadsheet.

This project started as a practical tool to automate the repetitive parts of that process. It's also a deliberate portfolio artifact: it demonstrates API integration, LLM orchestration, structured data design, and agentic thinking in a domain I know deeply — AI quality and evaluation.

The longer-term vision is a specialized AI career advisor that knows my full application history, learns from my outcomes, and gets smarter about what to surface over time.

---

## Setup

> Note: This project runs entirely in Google Apps Script — no local Python environment or server required.

1. Create a copy of the Google Sheet with the schema described above
2. Open **Extensions → Apps Script** in the sheet
3. Paste the contents of `job_search_agent.gs` into the editor
4. Go to **Project Settings → Script Properties** and add:
   - `JSEARCH_KEY` — your RapidAPI key for JSearch
   - `GEMINI_KEY` — your Google AI Studio API key
5. Update the `SHEET_NAME` constant if your tab is named differently
6. Run `fetchJobs()` to test
7. Set a time-based trigger to run `fetchJobs()` daily

---

## Files

```
job-search-agent/
├── README.md
└── job_search_agent.gs    # Full Apps Script — fetch, analyze, backfill
```

---

## Notes

- API keys are stored in Apps Script Script Properties, never in the code
- JSearch free tier: 200 requests/month — `num_pages=3` at 3 requests/run keeps this sustainable
- Gemini free tier is sufficient with 2000-character description truncation and 4-second delays between calls
- `date_posted=week` filters JSearch results to the last 7 days, though Google's index dates are unreliable — `actual_post_date` is extracted from job description text as a more accurate fallback

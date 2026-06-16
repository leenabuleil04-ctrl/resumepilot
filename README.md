# ResumePilot — AI CV Tailoring App

> **Live:** https://resumepilot-app.streamlit.app  
> **Repo:** https://github.com/leenabuleil04-ctrl/resumepilot (branch: `main`)  
> **Main file:** `app_debug_v2.py` (2,973 lines)  
> **Stack:** Python · Streamlit · Anthropic Claude API · python-docx · PyPDF2

---

## What It Does

ResumePilot takes a job description and a candidate's CV, runs them through an 11-step AI pipeline powered by Claude, and outputs a professionally rewritten, ATS-optimized resume in English, Hebrew, or Arabic — without inventing any experience the candidate doesn't have.

---

## Architecture Overview

```
User Input (Job + CV)
        │
        ▼
┌─────────────────────────────────────────────┐
│           6-Step User Flow (Tabs)           │
│  Tab 0: Job  →  Tab 1: CV  →  Tab 2: Analysis  │
│  Tab 3: Refine  →  Tab 4: Preview  →  Tab 5: Export │
└─────────────────────────────────────────────┘
        │
        ▼
┌─────────────────────────────────────────────┐
│         11-Step AI Pipeline (Claude)        │
│  Steps 1+2: Deep CV + JD Analysis          │
│  Steps 3+4: Question Generation + Memory   │
│  Steps 5-10: Full CV Rewrite               │
│  Step 9: Multilingual Localization         │
│  Step 10: Improvement Log                  │
└─────────────────────────────────────────────┘
        │
        ▼
   DOCX / TXT Download
```

---

## 6-Step User Flow

### Tab 0 — The Job
- User enters **target job title** and pastes **full job description**
- Validation: minimum 30 words, minimum 10 unique words
- Demo job loader pre-fills a Junior Data Analyst JD for testing
- Session state: `job_role`, `job_desc`

### Tab 1 — CV Content
Three input methods:
- **Upload File** — PDF, DOCX, or TXT. Text extracted via PyPDF2 / python-docx. Auto-detects Hebrew/Arabic and warns user to use English version.
- **Paste Text** — Free-text paste with "Auto-Structure" button that runs section classifier
- **Manual Entry** — 9 labeled fields (personal details, summary, education, experience, projects, skills, courses, languages, volunteering)

All methods feed into `classify_sections_refined()` which parses raw text into structured section dict.

Session state: `cv_full_text`, `manual_cv_data`, `cv_uploaded`

### Tab 2 — Analysis (Steps 1+2)
Runs `call_claude_deep_analysis()` which sends CV + JD to Claude and returns:
- Match score (0–100, capped at 88 for realism)
- Match label: Weak / Moderate / Strong
- CV strengths and weaknesses (specific, referencing actual content)
- Hard skills, soft skills, domain knowledge required by JD
- ATS keywords present vs missing (exact phrases from JD)
- Bullets needing metrics
- Top 3 quick wins

Falls back to rule-based hybrid engine (`run_rule_based_analysis`) if Claude unavailable.

Session state: `ai_cv_analysis`, `analysis_results`, `analysis_done`

### Tab 3 — Refinement (Steps 3+4)
Runs `call_claude_generate_questions()` — generates 5–8 personalized questions targeting:
- Specific gaps identified in Step 2
- Missing ATS keywords
- Bullets that need numbers/outcomes
- Experience mentioned in CV that needs scope added

Questions reference actual CV content (project names, roles, tools). Falls back to generic gap-based questions if Claude unavailable.

User answers stored as structured memory dict:
```python
{
  "q_key": {
    "question": "...",
    "answer": "...",
    "section": "experience",
    "gap_addressed": "..."
  }
}
```

Session state: `dynamic_questions`, `follow_up_answers`

### Tab 4 — Preview (Steps 5–10)
Runs `call_claude_rewrite_cv()` — full CV reconstruction using:
- Original CV sections
- Job description
- Deep analysis results
- All user answers from Tab 3

Rewrite rules (enforced in prompt):
- Never add companies, degrees, or dates not in original CV
- Never invent metrics unless user explicitly stated them
- If user said "I don't have X" — X is not added anywhere
- All proper nouns stay exactly as written

Returns rewritten sections + `improvement_log` (list of what changed and why).

For Hebrew/Arabic: runs `call_claude_localize_cv()` which does native-level localization (not word-by-word translation). Falls back to `translate_cv_googletrans()` via deep-translator if Claude unavailable.

CV rendered as HTML preview using `render_cv_html()`.

Session state: `adjusted_cv_data`, `final_cv_data`, `improvement_log`, `cv_adjusted`, `active_lang`

### Tab 5 — Export
- **DOCX** — generated via `generate_docx()` using python-docx. Recruiter-ready formatting with section headings, bullet points, RTL support for Hebrew/Arabic.
- **TXT** — plain text with section headers for online job portal forms.

File named: `{FirstName}_{JobRole}_CV.docx`

---

## Design System

### Colors (CSS variables)
```css
--blue:       #2563EB   /* primary actions, links, active states */
--blue-hover: #1D4ED8
--blue-light: #EFF6FF   /* tip boxes, selected states */
--blue-mid:   #BFDBFE   /* borders on blue backgrounds */
--text:       #0F172A   /* primary text */
--muted:      #64748B   /* secondary text, hints */
--border:     #E5E7EB   /* card borders, dividers */
--bg:         #F8FAFC   /* page background */
--card:       #FFFFFF   /* card surfaces */
--success:    #16A34A
--warning:    #F59E0B
--error:      #DC2626
```

### Typography
- Font: **Inter** (Google Fonts), fallback to -apple-system
- Page badge: 11px, blue, uppercase, 0.6px letter-spacing
- Page title: 26px, 700 weight, -0.6px letter-spacing
- Page subtitle: 13px, muted color
- Body: 13px, regular weight
- Card content: 13px

### Components
- **Step bar** — horizontal progress tracker with dots (done/active/todo states) and connectors
- **Cards** — white bg, 1px border, 12px radius, 1px 4px shadow
- **Pills** — color-coded: green (strong), yellow (partial), red (missing)
- **Alert boxes** — green/orange/red backgrounds for status messages
- **Tip box** — blue-light background with blue border
- **Bento grid** — 3-column CSS grid on landing page, hover animation with blue top border reveal
- **Buttons** — primary (blue filled), secondary (blue outline), disabled (gray)

### Landing Page
Sections in order:
1. Hero with gradient radial background, H1 with blue `<span>`, subtitle, CTA button
2. Bento grid — 6 feature cards (semantic matching, gap detection, questions, rewriting, multilingual, DOCX)
3. "How it works" — 3-step section on gray background
4. Footer with brand + tagline

---

## AI Pipeline — Function Reference

| Function | Steps | Purpose | Max tokens |
|---|---|---|---|
| `call_claude_deep_analysis()` | 1+2 | CV+JD analysis, score, gaps, ATS keywords | 2000 |
| `call_claude_generate_questions()` | 3+4 | Personalized gap-targeting questions | 1500 |
| `call_claude_rewrite_cv()` | 5–10 | Full CV reconstruction | 4500 |
| `call_claude_localize_cv()` | 9 | Hebrew/Arabic professional localization | 4000 |

All calls go through `claude_api_call()` — central caller using `urllib.request` (no SDK dependency). Uses `claude-sonnet-4-6`. API key read from `st.secrets["ANTHROPIC_API_KEY"]`.

Error handling: exceptions stored in `st.session_state['_last_api_error']` and surfaced in UI warnings.

---

## Hybrid Matching Engine

Used as fallback when Claude is unavailable. Defined in `SKILL_GROUPS` — 24 skill groups covering:
- Programming languages (Python, SQL, R, JavaScript, etc.)
- Data tools (Tableau, Power BI, Excel, BigQuery)
- ML/AI (scikit-learn, TensorFlow, PyTorch)
- Cloud (AWS, Azure, GCP)
- Education (Bachelor's, Master's, CS degree, Stats degree)
- Languages (English, Hebrew, Arabic)
- Soft skills (communication, teamwork, problem solving, project management)

Each group has: JD trigger patterns (regex), strong keywords (full credit), partial keywords (50% credit), weight (1–3), and category.

Score formula:
```
raw_score = (earned_weight / total_weight) * 100
score = min(88, max(20, raw_score))
```

---

## CV Section Parser

`classify_sections_refined()` handles flat PDF text, structured DOCX text, and manual entry:
1. Pre-inserts newlines before known ALL-CAPS headers (English + Hebrew + Arabic)
2. Matches lines against `SECTION_REGEX` patterns (regex for each section in 3 languages)
3. Routes content into 9 section buckets: personal_details, summary, education, experience, projects, courses_training, volunteering, languages, skills

Sections detected: personal details, summary, education, experience, projects, courses/training, volunteering, languages, skills, additional information.

---

## Multilingual Support

| Language | Output | Direction | Font | Localization method |
|---|---|---|---|---|
| English | ✅ | LTR | Calibri | Native (rewrite) |
| Hebrew | ✅ | RTL | Arial | Claude (native Israeli CV conventions) |
| Arabic | ✅ | RTL | Arial | Claude (Modern Standard Arabic) |

Rules enforced during localization:
- All technical tool names stay in English (Python, SQL, GitHub, etc.)
- All institution/company names stay in original language
- All dates stay exactly as written
- All numbers, percentages, GPA values stay unchanged

Fallback: `translate_cv_googletrans()` using deep-translator with placeholder protection for proper nouns and dates.

DOCX export applies RTL paragraph formatting via `w:bidi` XML element when Hebrew/Arabic selected.

---

## Session State Reference

| Key | Type | Set in | Used in |
|---|---|---|---|
| `show_landing` | bool | init | App entry |
| `active_tab` | int | `go_to()` | All tabs |
| `job_role` | str | Tab 0 | Tab 2, 4, 5 |
| `job_desc` | str | Tab 0 | Tab 2, 3, 4 |
| `cv_full_text` | str | Tab 1 | Tab 2, 3 |
| `manual_cv_data` | dict | Tab 1 | Tab 2, 3, 4 |
| `cv_uploaded` | bool | Tab 1 | Tab 1 nav |
| `analysis_done` | bool | Tab 2 | Tab 3 |
| `ai_cv_analysis` | dict | Tab 2 | Tab 3, 4 |
| `analysis_results` | dict | Tab 2 | Tab 3, 5 |
| `dynamic_questions` | list | Tab 3 | Tab 3 |
| `follow_up_answers` | dict | Tab 3 | Tab 4 |
| `cv_adjusted` | bool | Tab 4 | Tab 4 |
| `adjusted_cv_data` | dict | Tab 4 | Tab 4, 5 |
| `final_cv_data` | dict | Tab 4 | Tab 5 |
| `improvement_log` | list | Tab 4 | Tab 4 |
| `active_lang` | str | Tab 4 | Tab 4, 5 |
| `_last_api_error` | str | API calls | Tab 3, 4 |

Navigation: `go_to(tab_index)` clears downstream state when going back (e.g., going back to Tab 1 wipes `ai_cv_analysis`, going back to Tab 2 wipes `dynamic_questions`).

---

## File Structure

```
cv_match/
├── app_debug_v2.py        # Main app — entire codebase
└── .streamlit/
    └── secrets.toml       # ANTHROPIC_API_KEY = "sk-ant-..."
```

No other files required. All dependencies installed via pip.

---

## Dependencies

```
streamlit
anthropic              # (not imported directly — using urllib.request)
python-docx
PyPDF2
deep-translator        # fallback translation
pandas                 # (imported, minor use)
```

---

## Deployment

Deployed on **Streamlit Community Cloud** connected to `leenabuleil04-ctrl/resumepilot` repo, `main` branch, entry point `app_debug_v2.py`.

API key set in Streamlit Cloud → App Settings → Secrets:
```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

To deploy changes: commit and push to `main` — Streamlit Cloud auto-redeploys within ~2 minutes.

---

## Known Issues & Roadmap

### Fixed (this version)
- ✅ 7 bugs fixed across pipeline
- ✅ 11-step AI pipeline fully rebuilt
- ✅ Full visual redesign (blue system, new homepage, bento grid)
- ✅ API errors now surfaced with real exception message
- ✅ Bold markdown (`**text**`) in question text now renders correctly as HTML

### In Progress
- 🔵 End-to-end AI pipeline test (upload CV + JD → full output verification)
- 🔵 "Land more interviews" header text rendering fix

### Roadmap
- **Phase 2:** Step flow redesign — replace tab navigation with a guided linear stepper UI
- **Phase 3:** Hebrew/Arabic DOCX fix — proper RTL text direction, font embedding, and bidirectional bullet points in exported Word documents

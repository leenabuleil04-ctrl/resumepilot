import streamlit as st
import pandas as pd
from io import BytesIO
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
import PyPDF2
import re
import json
import hashlib
from html import escape as _html_escape


def _esc(value):
    """Escape user or AI text before putting it inside HTML."""
    return _html_escape(str(value or ""), quote=False)
from collections import Counter

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="ResumePilot",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Hide Streamlit chrome — watermark, toolbar, deploy button
st.markdown(
    "<style>"
    "#MainMenu,footer,header{visibility:hidden!important;}"
    "[data-testid='stToolbar'],[data-testid='stDecoration'],"
    "[data-testid='stStatusWidget'],.stDeployButton{display:none!important;}"
    "</style>",
    unsafe_allow_html=True,
)

# --- INITIALIZE SESSION STATE ---
defaults = {
    'active_tab': 0,
    'job_role': "",
    'cv_full_text': "",
    'job_desc': "",
    'input_method': "Upload File",
    'manual_cv_data': {},
    'follow_up_answers': {},        # {key: {question, answer, section, gap_addressed}}
    'analysis_results': {"matches": [], "missing": [], "score": 0},
    'ai_cv_analysis': {},           # Step 1+2: deep CV+JD analysis from Claude
    'dynamic_questions': [],        # Step 3: personalized questions [{key, question, hint, section, gap_addressed}]
    'improvement_log': [],
    'ai_sections': {},              # AI-rewritten sections, waiting for the user's review          # Step 10: [{type, section, description}]
    'final_cv_data': {},
    'adjusted_cv_data': {},
    'career_recommendations': {},   # Low-match gap analysis / roadmap
    'active_lang': "English",
    'cv_uploaded': False,
    'analysis_done': False,
    'cv_adjusted': False,
    'show_landing': False,         # the website (docs/) is the homepage now
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --- CUSTOM CSS ---
# One design system shared with the website (docs/index.html):
# navy ink on cool paper, Bricolage Grotesque headings, Atkinson Hyperlegible body,
# and an editor's palette: red = removed, green = added, yellow = what the job asks for.
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700;12..96,800&family=Atkinson+Hyperlegible:ital,wght@0,400;0,700;1,400&display=swap');

    :root {
        --paper:   #F6F8FB;
        --sheet:   #FFFFFF;
        --ink:     #16213A;
        --muted:   #5A6478;
        --rule:    #DCE2EC;
        --strike:  #D6354A;
        --insert:  #0B7A5C;
        --mark:    #FFE45C;
        --action:  #2A3FD6;
        --display: "Bricolage Grotesque", system-ui, sans-serif;
        --body:    "Atkinson Hyperlegible", system-ui, sans-serif;
        /* old names kept so existing markup picks up the new look */
        --blue: var(--ink); --blue-hover: #0E1628; --blue-light: #EEF1FB; --blue-mid: #D5DCF5;
        --text: var(--ink); --border: var(--rule); --bg: var(--paper); --card: var(--sheet);
        --success: var(--insert); --warning: #B7791F; --error: var(--strike);
    }

    html, body, .stApp, [class*="css"], .stMarkdown, p, li, label, input, textarea, button, select {
        font-family: var(--body) !important;
    }
    .stApp { background: var(--paper) !important; color: var(--ink); }
    .block-container { padding: 0 24px 64px !important; max-width: 920px !important; }
    div[data-testid="stTabs"] { display: none; }
    a { color: var(--action); }
    h1, h2, h3, h4 { font-family: var(--display) !important; color: var(--ink) !important; letter-spacing: -0.02em; }
    :focus-visible { outline: 3px solid var(--action) !important; outline-offset: 2px; }

    /* ── HEADER ── */
    .rp-header { display: flex; align-items: center; gap: 10px; padding: 22px 0 18px; margin-bottom: 8px; }
    .rp-logo { display: flex; align-items: center; gap: 10px; text-decoration: none !important; }
    .rp-logo-mark { width: 26px; height: 32px; background: var(--sheet); border: 2px solid var(--ink); border-radius: 4px; position: relative; flex-shrink: 0; }
    .rp-logo-mark::before { content: ""; position: absolute; left: 4px; right: 4px; top: 8px; height: 3px; background: var(--strike); }
    .rp-logo-mark::after  { content: ""; position: absolute; left: 4px; right: 8px; top: 16px; height: 6px; background: var(--mark); }
    .rp-brand { font: 800 21px/1 var(--display); color: var(--ink); letter-spacing: -0.01em; }
    .rp-tagline { margin-left: auto; font-size: 15px; font-weight: 700; }
    .rp-tagline a { color: var(--ink); text-decoration: none; }
    .rp-tagline a:hover { text-decoration: underline; }

    /* ── STEP BAR: same numbered rule as the website's "how it works" ── */
    .rp-step-bar { display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; margin: 8px 0 40px; }
    .rp-step-item { border-top: 3px solid var(--rule); padding-top: 10px; }
    .rp-step-item.done   { border-top-color: var(--insert); }
    .rp-step-item.active { border-top-color: var(--ink); }
    .rp-step-num  { font: 800 20px/1 var(--display); color: #A3ACBD; }
    .rp-step-name { font-size: 14px; font-weight: 700; color: #8A93A6; margin-top: 4px; }
    .rp-step-item.done .rp-step-num, .rp-step-item.done .rp-step-name { color: var(--insert); }
    .rp-step-item.active .rp-step-num, .rp-step-item.active .rp-step-name { color: var(--ink); }
    @media (max-width: 640px) { .rp-step-name { display: none; } .rp-step-bar { gap: 6px; } }

    /* ── PAGE TITLES ── */
    .rp-page-badge { font-size: 15px; color: var(--muted); font-weight: 700; margin-bottom: 6px; }
    .rp-page-title { font: 800 clamp(34px, 5vw, 48px)/1.02 var(--display); letter-spacing: -0.03em; color: var(--ink); margin-bottom: 12px; }
    .rp-page-sub   { font-size: 19px; color: var(--muted); margin-bottom: 32px; line-height: 1.55; max-width: 36em; }

    /* ── CARDS ── */
    .rp-card, .rp-card-white, .content-card, .rp-checklist {
        background: var(--sheet); border: 1px solid var(--rule); border-radius: 10px;
        padding: 18px 22px; margin-bottom: 16px; box-shadow: none;
    }
    .rp-pill { font-size: 13px; font-weight: 700; padding: 2px 10px; border-radius: 9999px; }
    .rp-pill-strong  { background: #EAF6F1; color: var(--insert); }
    .rp-pill-partial { background: #FFF6C2; color: #6B5A00; }
    .rp-pill-missing { background: #FBE9EC; color: #B4233A; }
    .rp-tip { background: var(--sheet); border-left: 3px solid var(--mark); border-radius: 0 8px 8px 0; padding: 14px 16px; font-size: 16px; color: var(--ink); line-height: 1.55; margin-top: 14px; }
    .rp-q-num { display: inline-flex; align-items: center; justify-content: center; background: var(--ink); color: #fff; width: 24px; height: 24px; border-radius: 50%; font: 800 13px var(--display); margin-right: 8px; }
    .rp-q-text { font-size: 17px; font-weight: 700; color: var(--ink); margin-bottom: 8px; }
    .rp-dl-fmt  { font: 700 19px var(--display); margin-bottom: 3px; color: var(--ink); }
    .rp-dl-desc { font-size: 15px; color: var(--muted); margin-bottom: 14px; }
    .rp-checklist-title { font: 700 18px var(--display); margin-bottom: 10px; color: var(--ink); }
    .rp-check-item { font-size: 15px; color: var(--muted); padding: 7px 0; border-bottom: 1px solid var(--rule); }
    .rp-check-item:last-child { border-bottom: none; }
    .alert-box { padding: 12px 16px; border-radius: 8px; font-size: 15px; font-weight: 700; margin-bottom: 12px; }
    .alert-green  { background: #EAF6F1; color: var(--insert); }
    .alert-orange { background: #FFF6C2; color: #6B5A00; }
    .alert-red    { background: #FBE9EC; color: #B4233A; }

    /* ── BUTTONS: same as the website ── */
    .stButton > button, .stDownloadButton > button {
        border-radius: 12px !important; font: 700 16px/1 var(--body) !important;
        padding: 14px 20px !important; min-height: 50px; transition: transform .15s ease !important;
        border: 2px solid var(--ink) !important; background: transparent !important; color: var(--ink) !important;
        box-shadow: none !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover { transform: translateY(-2px); }
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
        background: var(--ink) !important; color: #fff !important;
    }
    .stButton > button p, .stDownloadButton > button p, .stButton > button div, .stDownloadButton > button div { font: 700 16px/1.2 var(--body) !important; color: inherit !important; }
    .stButton > button:disabled { background: #E7EBF2 !important; color: #8A93A6 !important; border-color: #E7EBF2 !important; transform: none; cursor: not-allowed !important; }

    /* ── INPUTS ── */
    .stTextInput input, .stTextArea textarea {
        border-radius: 10px !important; border: 1.5px solid var(--rule) !important; background: var(--sheet) !important;
        font-size: 17px !important; color: var(--ink) !important; padding: 12px 14px !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus { border-color: var(--action) !important; box-shadow: 0 0 0 3px rgba(42,63,214,.15) !important; }
    .stTextInput label p, .stTextArea label p, .stFileUploader label p, div[data-testid="stRadio"] > label p {
        font-size: 16px !important; font-weight: 700 !important; color: var(--ink) !important;
    }
    [data-testid="stFileUploaderDropzone"] { background: var(--sheet) !important; border: 2px dashed var(--rule) !important; border-radius: 12px !important; }

    /* ── RADIOS AS CHOICE CARDS ── */
    div[data-testid="stRadio"] > div[role="radiogroup"] { display: flex !important; flex-direction: row !important; gap: 10px !important; flex-wrap: wrap !important; }
    div[data-testid="stRadio"] > div[role="radiogroup"] > label {
        display: flex !important; align-items: center !important; gap: 8px !important;
        background: var(--sheet) !important; border: 2px solid var(--rule) !important; border-radius: 10px !important;
        padding: 10px 16px !important; cursor: pointer !important; min-width: 130px !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] > label p { font-size: 16px !important; font-weight: 700 !important; color: var(--ink) !important; }
    div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) { border-color: var(--ink) !important; }
    div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover { border-color: var(--ink) !important; }

    /* ── EXPANDERS / ALERTS ── */
    div[data-testid="stExpander"] details { border: 1px solid var(--rule) !important; border-radius: 10px !important; background: var(--sheet) !important; }
    div[data-testid="stExpander"] summary p { font: 700 17px var(--display) !important; color: var(--ink) !important; }
    div[data-testid="stAlert"], div[data-testid="stNotification"] { border-radius: 10px !important; }

    /* ── SCORE CARD ── */
    .rp-score-dark { background: var(--ink); color: #fff; border-radius: 14px; padding: 26px 30px; display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
    .rp-score-num   { font: 800 64px/1 var(--display); letter-spacing: -0.04em; }
    .rp-score-label { font-size: 15px; color: rgba(255,255,255,.7); margin-bottom: 4px; }
    .rp-score-desc  { font-size: 15px; color: rgba(255,255,255,.7); margin-top: 6px; }
    .rp-score-right { text-align: right; font-size: 15px; line-height: 2; }
    .rp-score-s { color: #8FE0C2; } .rp-score-p { color: var(--mark); } .rp-score-m { color: #F5A3AE; }
    .rp-analysis-item { display: flex; align-items: center; justify-content: space-between; padding: 10px 14px; border-radius: 8px; background: var(--sheet); border: 1px solid var(--rule); margin-bottom: 6px; font-size: 16px; color: var(--ink); }

    /* ── TRACKED CHANGES (the review step) ── */
    .rp-diff { background: var(--sheet); border: 1px solid var(--rule); border-radius: 8px; padding: 18px 20px; font-size: 16px; line-height: 1.7; white-space: pre-wrap; color: var(--ink); }
    .rp-del { color: var(--strike); text-decoration: line-through; text-decoration-thickness: 2px; }
    .rp-ins { color: var(--insert); font-weight: 700; }
    .rp-why { border-left: 3px solid var(--mark); padding: 2px 0 2px 12px; margin: 0 0 12px; font-size: 15px; }
    .rp-why b { display: block; font-size: 13px; color: var(--muted); }
    .rp-why .req { background: linear-gradient(transparent 55%, var(--mark) 55%); }
    .rp-legend { font-size: 15px; color: var(--muted); margin: -18px 0 18px; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# CONSTANTS & DICTIONARIES (unchanged from original)
# ============================================================

STOPWORDS = {
    "a","an","the","and","or","but","if","then","else","when","at","from","by",
    "for","with","about","against","between","into","through","during","before",
    "after","above","below","to","up","down","in","out","on","off","over","under",
    "again","further","once","here","there","all","any","both","each","few","more",
    "most","other","some","such","no","nor","not","only","own","same","so","than",
    "too","very","s","t","can","will","just","don","should","now","i","me","my",
    "myself","we","our","ours","ourselves","you","your","yours","yourself",
    "yourselves","he","him","his","himself","she","her","hers","herself","it","its",
    "itself","they","them","their","theirs","themselves","what","which","who","whom",
    "this","that","these","those","am","is","are","was","were","be","been","being",
    "have","has","had","having","do","does","did","doing"
}

SKILL_KEYWORDS = {
    "python","sql","java","javascript","react","node","html","css","git","github",
    "r","matlab","tableau","powerbi","power bi","excel","bigquery","pandas","numpy",
    "scikit-learn","tensorflow","pytorch","aws","azure","docker","kubernetes","unix",
    "linux","bash","spark","hadoop","c++","c#","php","ruby","swift","kotlin",
    "typescript","scala","go","rust","mongodb","postgresql","mysql","redis","kafka",
    "airflow","dbt","looker","databricks","snowflake","jira","confluence","jenkins",
}

# ============================================================
# SKILL GROUPS — Heart of the hybrid matching engine
# Each group has: trigger patterns (what to look for in JD),
# strong keywords (direct CV match), partial keywords (related),
# weight (importance), and a display label.
# ============================================================

SKILL_GROUPS = [
    # ── Programming Languages ──
    {
        "label": "Python",
        "weight": 3,
        "jd_patterns": [r'\bpython\b'],
        "strong": ["python", "pandas", "numpy", "scipy", "flask", "django", "scikit"],
        "partial": ["programming", "scripting", "automation", "jupyter"],
        "category": "Technical"
    },
    {
        "label": "SQL / Databases",
        "weight": 3,
        "jd_patterns": [r'\bsql\b', r'\bdatabase', r'\bqueries\b'],
        "strong": ["sql", "mysql", "postgresql", "sqlite", "bigquery", "database", "queries"],
        "partial": ["data manipulation", "data extraction", "spreadsheet", "excel"],
        "category": "Technical"
    },
    {
        "label": "R / Statistical Computing",
        "weight": 2,
        "jd_patterns": [r'\br\b(?:\s+programming|\s+language|\s+studio)?', r'\brstudio\b', r'\bggplot\b'],
        "strong": ["r", "rstudio", "ggplot", "tidyverse", "dplyr", "r programming"],
        "partial": ["statistics", "statistical analysis", "spss", "matlab", "python"],
        "category": "Technical"
    },
    {
        "label": "Data Visualization",
        "weight": 2,
        "jd_patterns": [r'\bvisuali', r'\btableau\b', r'\bpower\s*bi\b', r'\bdashboard'],
        "strong": ["tableau", "power bi", "powerbi", "matplotlib", "plotly", "seaborn", "ggplot", "dashboard", "visualization"],
        "partial": ["charts", "graphs", "reports", "excel", "presentations", "plots"],
        "category": "Technical"
    },
    {
        "label": "Machine Learning",
        "weight": 2,
        "jd_patterns": [r'\bmachine\s+learning\b', r'\bml\b', r'\bai\b', r'\bclassif', r'\bcluster'],
        "strong": ["machine learning", "ml", "scikit-learn", "sklearn", "tensorflow", "pytorch", "classification", "clustering", "neural"],
        "partial": ["predictive", "modeling", "statistical modeling", "regression", "deep learning", "ai"],
        "category": "Technical"
    },
    {
        "label": "Statistical Analysis",
        "weight": 2,
        "jd_patterns": [r'\bstatistic', r'\bregression\b', r'\bhypothesis', r'\bmodeling\b|\bmodelling\b'],
        "strong": ["statistics", "statistical", "regression", "hypothesis testing", "modeling", "quantitative", "spss", "anova"],
        "partial": ["data analysis", "analytics", "r", "python", "research", "math"],
        "category": "Technical"
    },
    {
        "label": "Git / Version Control",
        "weight": 1,
        "jd_patterns": [r'\bgit\b', r'\bversion\s+control\b', r'\bgithub\b'],
        "strong": ["git", "github", "gitlab", "version control", "bitbucket"],
        "partial": ["code", "repository", "collaboration", "open source"],
        "category": "Technical"
    },
    {
        "label": "Linux / Unix",
        "weight": 1,
        "jd_patterns": [r'\blinux\b', r'\bunix\b', r'\bbash\b', r'\bshell\b'],
        "strong": ["linux", "unix", "bash", "shell", "terminal", "command line"],
        "partial": ["scripting", "automation", "python", "programming"],
        "category": "Technical"
    },
    {
        "label": "Cloud (AWS / Azure / GCP)",
        "weight": 2,
        "jd_patterns": [r'\baws\b', r'\bazure\b', r'\bgcp\b', r'\bcloud\b'],
        "strong": ["aws", "azure", "gcp", "cloud", "s3", "lambda", "google cloud"],
        "partial": ["deployment", "docker", "kubernetes", "devops"],
        "category": "Technical"
    },
    {
        "label": "Excel / Spreadsheets",
        "weight": 1,
        "jd_patterns": [r'\bexcel\b', r'\bspreadsheet\b'],
        "strong": ["excel", "spreadsheet", "pivot", "vlookup", "google sheets"],
        "partial": ["data analysis", "reporting", "sql", "tables"],
        "category": "Technical"
    },
    {
        "label": "NLP / Text Analysis",
        "weight": 2,
        "jd_patterns": [r'\bnlp\b', r'\bnatural\s+language', r'\btext\s+(?:mining|analysis|processing)'],
        "strong": ["nlp", "natural language processing", "spacy", "nltk", "bert", "transformers", "text mining"],
        "partial": ["python", "machine learning", "classification", "text"],
        "category": "Technical"
    },
    {
        "label": "Network / Graph Analysis",
        "weight": 2,
        "jd_patterns": [r'\bnetwork\s+analysis\b', r'\bgraph\s+(?:theory|analysis)\b', r'\bnetworkx\b'],
        "strong": ["network analysis", "networkx", "graph", "nodes", "edges", "clustering coefficient"],
        "partial": ["python", "data analysis", "algorithms"],
        "category": "Technical"
    },
    {
        "label": "Agile / Scrum",
        "weight": 1,
        "jd_patterns": [r'\bagile\b', r'\bscrum\b', r'\bsprint\b', r'\bkanban\b'],
        "strong": ["agile", "scrum", "sprint", "kanban", "jira"],
        "partial": ["project management", "teamwork", "collaboration"],
        "category": "Technical"
    },
    # ── Education ──
    {
        "label": "Bachelor's Degree",
        "weight": 2,
        "jd_patterns": [r'\bbachelor', r'\bb\.?sc\b', r'\bundergraduate\b', r'\bdegree\b'],
        "strong": ["bachelor", "bsc", "b.sc", "undergraduate", "university", "college", "degree"],
        "partial": ["studying", "student", "final year", "graduate"],
        "category": "Education"
    },
    {
        "label": "Master's Degree",
        "weight": 2,
        "jd_patterns": [r'\bmaster', r'\bm\.?sc\b', r'\bmba\b', r'\bpostgraduate\b'],
        "strong": ["master", "msc", "m.sc", "mba", "postgraduate"],
        "partial": ["bachelor", "university", "advanced degree"],
        "category": "Education"
    },
    {
        "label": "Computer Science Degree",
        "weight": 2,
        "jd_patterns": [r'\bcomputer\s+science\b', r'\bcs\s+degree\b', r'\bsoftware\s+engineering\s+degree\b'],
        "strong": ["computer science", "cs degree", "software engineering", "information systems"],
        "partial": ["programming", "algorithms", "data structures", "university"],
        "category": "Education"
    },
    {
        "label": "Statistics / Math Degree",
        "weight": 2,
        "jd_patterns": [r'\bstatistics\s+degree\b', r'\bmathematic', r'\bquantitative\s+field\b'],
        "strong": ["statistics", "mathematics", "math", "quantitative", "data science degree"],
        "partial": ["engineering", "physics", "economics", "university"],
        "category": "Education"
    },
    # ── Languages ──
    {
        "label": "English",
        "weight": 1,
        "jd_patterns": [r'\benglish\b'],
        "strong": ["english"],
        "partial": ["fluent", "bilingual", "native"],
        "category": "Languages"
    },
    {
        "label": "Hebrew",
        "weight": 2,
        "jd_patterns": [r'\bhebrew\b'],
        "strong": ["hebrew"],
        "partial": ["fluent", "native", "israel"],
        "category": "Languages"
    },
    {
        "label": "Arabic",
        "weight": 1,
        "jd_patterns": [r'\barabic\b'],
        "strong": ["arabic"],
        "partial": ["fluent", "native"],
        "category": "Languages"
    },
    # ── Soft Skills ──
    {
        "label": "Communication Skills",
        "weight": 1,
        "jd_patterns": [r'\bcommunication\b', r'\bpresent', r'\bstakeholder'],
        "strong": ["communication", "presented", "presentation", "stakeholders", "written", "verbal", "public speaking"],
        "partial": ["tutoring", "teaching", "teamwork", "reporting", "documentation", "mentoring"],
        "category": "Soft Skills"
    },
    {
        "label": "Teamwork / Collaboration",
        "weight": 1,
        "jd_patterns": [r'\bteamwork\b', r'\bcollaborat', r'\bteam\s+player\b', r'\bcross.functional\b'],
        "strong": ["teamwork", "collaboration", "team player", "cross-functional", "worked with team"],
        "partial": ["project", "group", "together", "hackathon", "volunteer", "committee"],
        "category": "Soft Skills"
    },
    {
        "label": "Problem Solving",
        "weight": 1,
        "jd_patterns": [r'\bproblem.solv', r'\banalytical\b', r'\bcritical\s+think'],
        "strong": ["problem solving", "analytical", "critical thinking", "solutions"],
        "partial": ["research", "debugging", "algorithms", "optimized", "improved"],
        "category": "Soft Skills"
    },
    {
        "label": "Project Management",
        "weight": 1,
        "jd_patterns": [r'\bproject\s+management\b', r'\bproject\s+lead', r'\bcoordinat'],
        "strong": ["project management", "project lead", "coordinated", "managed project"],
        "partial": ["organized", "planned", "delivered", "timeline", "agile"],
        "category": "Soft Skills"
    },
]

# Words to NEVER treat as requirements — noise filter
NOISE_WORDS = {
    "junior","senior","analyst","engineer","developer","manager","associate","intern",
    "looking","seeking","hiring","join","growing","excited","dynamic","innovative",
    "responsibilities","requirements","qualifications","preferred","required","needed",
    "must","should","will","would","please","submit","apply","resume","cv",
    "company","team","organization","department","role","position","opportunity",
    "years","experience","background","foundation","ability","understanding",
    "familiarity","exposure","proficiency","minimum","maximum","ideally",
    "various","multiple","several","including","following","related","relevant",
    "excellent","strong","good","great","proven","demonstrated","solid",
    "work","working","use","using","used","provide","support","help","ensure",
    "maintain","create","build","deliver","drive","focus","based","given",
    "also","well","both","other","such","like","example","etc","ie","eg",
}


# ============================================================
# HYBRID ANALYSIS ENGINE
# ============================================================

def extract_active_skill_groups(job_text):
    """Find which skill groups are actually required by this job description."""
    job_lower = job_text.lower()
    active = []
    for group in SKILL_GROUPS:
        for pattern in group["jd_patterns"]:
            if re.search(pattern, job_lower):
                active.append(group)
                break
    return active


def match_group_against_cv(group, cv_text):
    """
    For one skill group, check the CV.
    Returns: "strong", "partial", or "missing"
    Also returns evidence (what was found).
    """
    cv_lower = cv_text.lower()
    # Check strong keywords
    found_strong = [kw for kw in group["strong"] if kw in cv_lower]
    if found_strong:
        return "strong", found_strong[0]
    # Check partial keywords
    found_partial = [kw for kw in group["partial"] if kw in cv_lower]
    if found_partial:
        return "partial", found_partial[0]
    return "missing", None


def run_rule_based_analysis(cv_text, job_desc):
    """
    Full rule-based hybrid analysis.
    Returns dict with strong, partial, missing lists + weighted score.
    """
    active_groups = extract_active_skill_groups(job_desc)
    if not active_groups:
        return {"strong": [], "partial": [], "missing": [], "score": 0, "groups": []}

    strong, partial, missing = [], [], []
    total_weight = 0
    earned_weight = 0

    for group in active_groups:
        result, evidence = match_group_against_cv(group, cv_text)
        w = group["weight"]
        total_weight += w
        entry = {"label": group["label"], "evidence": evidence, "category": group["category"], "weight": w}
        if result == "strong":
            strong.append(entry)
            earned_weight += w
        elif result == "partial":
            partial.append(entry)
            earned_weight += w * 0.5   # partial = 50% credit
        else:
            missing.append(entry)

    raw_score = (earned_weight / total_weight * 100) if total_weight > 0 else 0
    # Realistic cap: partial matches penalise slightly, cap at 88
    score = round(min(88, max(20, raw_score)))

    return {
        "strong": strong,
        "partial": partial,
        "missing": missing,
        "score": score,
        "total_reqs": len(active_groups),
        "groups": active_groups,
    }


def get_anthropic_api_key():
    """Return the Anthropic API key from st.secrets, or None if not configured."""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, AttributeError, FileNotFoundError):
        return None


# ============================================================
# CORE AI PIPELINE — 11-STEP CV TAILORING ENGINE
# ============================================================

# ============================================================
# USAGE LIMITS — protect the owner's API key from overuse
# ============================================================
# Every Claude call costs money. These limits cap how many calls one
# visitor, and the whole app, can make per day. A full CV run uses about
# 5 calls. Override the defaults in Streamlit secrets:
#   PER_USER_DAILY_CALLS = 15
#   GLOBAL_DAILY_CALLS   = 200
# Counts are kept in server memory and reset at midnight (Israel time)
# or when the app restarts.

DEFAULT_PER_USER_DAILY_CALLS = 15
DEFAULT_GLOBAL_DAILY_CALLS = 200
LIMIT_REACHED_USER_MSG = (
    "You've reached today's free limit for AI features. "
    "The app will keep working with basic (non-AI) analysis. Try again tomorrow."
)
LIMIT_REACHED_GLOBAL_MSG = (
    "ResumePilot has reached its daily AI limit for all users. "
    "The app will keep working with basic (non-AI) analysis. Try again tomorrow."
)


def _get_limit(name, default):
    try:
        return int(st.secrets[name])
    except Exception:
        return default


@st.cache_resource
def _usage_store():
    """One shared counter for all visitors (lives as long as the server process)."""
    import threading
    return {"lock": threading.Lock(), "day": None, "global": 0, "users": {}}


def _today():
    from datetime import datetime
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Asia/Jerusalem")).date().isoformat()
    except Exception:
        return datetime.utcnow().date().isoformat()


def _visitor_id():
    """Identify a visitor by IP address when available, else by browser session.

    IP-based counting stops people from resetting their limit by refreshing the page.
    """
    try:
        headers = st.context.headers
        fwd = headers.get("X-Forwarded-For") or headers.get("x-forwarded-for")
        if fwd:
            return "ip:" + fwd.split(",")[0].strip()
    except Exception:
        pass
    if "_visitor_id" not in st.session_state:
        import uuid
        st.session_state["_visitor_id"] = "session:" + uuid.uuid4().hex
    return st.session_state["_visitor_id"]


def check_and_count_usage():
    """Reserve one AI call. Returns (allowed, message_if_blocked)."""
    store = _usage_store()
    per_user = _get_limit("PER_USER_DAILY_CALLS", DEFAULT_PER_USER_DAILY_CALLS)
    global_cap = _get_limit("GLOBAL_DAILY_CALLS", DEFAULT_GLOBAL_DAILY_CALLS)
    vid = _visitor_id()
    with store["lock"]:
        today = _today()
        if store["day"] != today:
            store["day"], store["global"], store["users"] = today, 0, {}
        if store["global"] >= global_cap:
            return False, LIMIT_REACHED_GLOBAL_MSG
        used = store["users"].get(vid, 0)
        if used >= per_user:
            return False, LIMIT_REACHED_USER_MSG
        store["users"][vid] = used + 1
        store["global"] += 1
    return True, None


def claude_api_call(prompt, max_tokens=3000):
    """Central Claude API caller. Returns raw text or None on any failure."""
    api_key = get_anthropic_api_key()
    if not api_key:
        return None
    allowed, limit_msg = check_and_count_usage()
    if not allowed:
        st.session_state['_last_api_error'] = limit_msg
        if not st.session_state.get('_limit_warning_shown'):
            st.warning("" + limit_msg)
            st.session_state['_limit_warning_shown'] = True
        return None
    import urllib.request
    try:
        payload = json.dumps({
            "model": "claude-sonnet-4-6",
            "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": prompt}]
        }).encode()
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "anthropic-version": "2023-06-01",
                "x-api-key": api_key,
            }
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read())
            return data["content"][0]["text"]
    except Exception as e:
        st.session_state['_last_api_error'] = str(e)
        return None


def parse_json_response(text):
    if not text:
        return None
    # Strip code fences more aggressively
    text = re.sub(r'```json\s*', '', text)
    text = re.sub(r'```\s*', '', text)
    text = text.strip().strip('`').strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r'\{[\s\S]*\}', text)
        if m:
            try:
                return json.loads(m.group())
            except Exception:
                pass
        m = re.search(r'\[[\s\S]*\]', text)
        if m:
            try:
                arr = json.loads(m.group())
                if isinstance(arr, list):
                    return {"questions": arr}
            except Exception:
                pass
    return None


def build_sections_text(cv_sections):
    """Format CV sections as labeled text blocks for Claude prompts.

    Sections are sent in full. Cutting them here used to make the rewrite
    silently drop jobs and projects from long CVs.
    """
    out = ""
    order = ["personal_details", "summary", "education", "experience",
             "projects", "skills", "volunteering", "courses_training", "languages"]
    for key in order:
        val = cv_sections.get(key, "").strip()
        if val:
            label = key.upper().replace("_", " ")
            out += f"\n[{label}]\n{val}\n"
    return out


# ── STEP 1+2: Deep CV Analysis + Job Requirement Extraction ──
def call_claude_deep_analysis(cv_sections, job_role, job_desc):
    """
    Steps 1+2: Analyze CV structure, identify strengths/weaknesses,
    extract job requirements, compute ATS gaps and match score.
    Returns structured dict or None.
    """
    sections_text = build_sections_text(cv_sections)
    prompt = f"""You are a senior technical recruiter with 15 years of experience. Analyze this CV against the job description and return a precise, honest assessment.

TARGET ROLE: {job_role}

JOB DESCRIPTION:
{job_desc}

CV CONTENT:
{sections_text}

Return ONLY a JSON object (no markdown, no explanation):
{{
  "cv_strengths": ["specific strength referencing actual CV content", "another strength"],
  "cv_weaknesses": ["specific gap or weakness", "another gap"],
  "hard_skills_required": ["Python", "SQL", "exact hard skill from JD"],
  "soft_skills_required": ["communication", "exact soft skill from JD"],
  "domain_knowledge_required": ["data analysis", "specific domain expertise from JD"],
  "seniority_level": "junior",
  "ats_keywords_missing": ["exact phrase from JD absent from CV", "another keyword"],
  "ats_keywords_present": ["keyword found in both JD and CV"],
  "match_score": 68,
  "match_label": "Moderate Match",
  "score_rationale": "Honest 2-sentence explanation referencing actual CV content and JD requirements.",
  "metrics_missing": ["Experience bullet that needs a number or outcome added"],
  "quick_wins": ["Most impactful improvement #1", "Most impactful improvement #2", "Most impactful improvement #3"]
}}

STRICT RULES:
- match_score must be 0-100 integer — be realistic, not generous
- ats_keywords_missing: copy exact phrases/words from JD that are absent from the CV
- metrics_missing: paraphrase actual bullets from the CV that would benefit from quantification
- cv_strengths: cite only what is actually in this CV — no generic praise
- match_label must be one of: "Weak Match", "Moderate Match", "Strong Match" """

    raw = claude_api_call(prompt, max_tokens=2000)
    return parse_json_response(raw)


# ── STEP 3+4: Dynamic Question Generation & Structured Memory ──
def call_claude_generate_questions(cv_sections, job_role, job_desc, ai_analysis):
    """
    Step 3: Generate 5-8 specific, personalized questions targeting gaps
    identified in the deep analysis. NOT generic questions.
    Returns list of question dicts or [].
    """
    sections_text = build_sections_text(cv_sections)
    ats_missing    = (ai_analysis or {}).get("ats_keywords_missing", [])[:8]
    weaknesses     = (ai_analysis or {}).get("cv_weaknesses", [])[:5]
    metrics_gaps   = (ai_analysis or {}).get("metrics_missing", [])[:4]

    prompt = f"""You are a career coach helping a candidate strengthen their CV application. Generate 5-8 highly specific, personalized questions to uncover information that will make this CV noticeably stronger.

TARGET ROLE: {job_role}

CV GAPS IDENTIFIED BY ANALYSIS:
- Missing ATS keywords: {', '.join(ats_missing) or 'see CV below'}
- CV weaknesses: {', '.join(weaknesses) or 'see CV below'}
- Bullets needing metrics: {', '.join(metrics_gaps) or 'see CV below'}

CURRENT CV (abbreviated):
{sections_text}

QUESTION QUALITY BENCHMARK:
BAD: "Do you have SQL experience?" — too closed, not personalized
GOOD: "Your CV shows Python for data analysis — have you also used SQL to query databases, even in university coursework or side projects?"

BAD: "Tell us about your leadership skills"
GOOD: "You mention tutoring students — how many did you tutor, over what time period, and did you notice measurable improvement in their academic performance?"

BAD: "Describe a challenge you overcame"
GOOD: "Your network analysis project — what was the dataset size, and what specific finding or result did the analysis produce?"

Generate questions that:
1. Reference SPECIFIC content actually in this CV (name the project, role, or skill)
2. Target the identified gaps and missing keywords from the job description
3. Uncover quantifiable achievements (scope, numbers, outcomes)
4. Surface relevant experience with JD keywords the CV is missing

Return ONLY JSON:
{{
  "questions": [
    {{
      "key": "q_unique_snake_case_key",
      "question": "Full specific question text referencing their actual CV",
      "hint": "e.g., 'I tutored 15 students over 6 months, their grades improved by...'",
      "section": "experience",
      "gap_addressed": "Adds scope and outcome metrics to tutoring entry"
    }}
  ]
}}

Generate exactly 5-8 questions. Every question must be specific to THIS person's CV."""

    # NOTE: 1500 tokens was too tight for 5-8 questions (each with key,
    # question, hint, section, gap_addressed) plus the ```json fence —
    # Claude was hitting the token limit mid-response, which cut the JSON
    # off before it was valid (e.g. mid-string), causing "JSON parse
    # failed" and falling back to the generic questions. Bumped the
    # budget so full responses fit.
    raw = claude_api_call(prompt, max_tokens=3000)
    result = parse_json_response(raw)
    if result and isinstance(result.get("questions"), list):
        return result["questions"]

    # If it still failed (e.g. still truncated), try salvaging whatever
    # complete question objects are present in the partial JSON instead
    # of throwing everything away.
    if raw:
        salvaged = re.findall(r'\{\s*"key"\s*:\s*".*?\}\s*(?=,|\])', raw, flags=re.S)
        questions = []
        for chunk in salvaged:
            try:
                q = json.loads(chunk)
                if isinstance(q, dict) and q.get("key") and q.get("question"):
                    questions.append(q)
            except Exception:
                continue
        if questions:
            return questions
        st.session_state['_last_api_error'] = f"JSON parse failed. Preview: {raw[:120]}"
    return []


# ── STEPS 5-10: Full CV Rewrite — ATS, Summary, Bullets, Localization Prep ──
def call_claude_rewrite_cv(cv_sections, job_role, job_desc, ai_analysis, follow_up_answers, match_tier="high"):
    """
    Steps 5-10: True CV reconstruction using original CV + job description
    + all user answers. Produces professional-recruiter-quality output.
    Returns dict with rewritten sections + improvement_log, or None.
    """
    sections_text = build_sections_text(cv_sections)

    # Format answers preserving the structured memory (Step 4)
    answers_formatted = ""
    if follow_up_answers:
        for key, val in follow_up_answers.items():
            if isinstance(val, dict) and val.get("answer", "").strip():
                q = val.get("question", key)
                a = val["answer"].strip()
                answers_formatted += f"Q: {q}\nA: {a}\n\n"
            elif isinstance(val, str) and val.strip():
                answers_formatted += f"- {val.strip()}\n"

    ats_missing  = (ai_analysis or {}).get("ats_keywords_missing", [])[:12]
    strengths    = (ai_analysis or {}).get("cv_strengths", [])[:5]
    hard_skills  = (ai_analysis or {}).get("hard_skills_required", [])[:8]

    tier_instructions = ""
    if match_tier == "medium":
        tier_instructions = """
MATCH TIER: MEDIUM (35–69%)
This CV partially matches the job. Apply these honesty constraints:
- Only strengthen sections where the candidate's background genuinely overlaps with the job
- Do NOT force ATS keywords into bullets where they don't naturally apply
- If a required skill is completely absent from the CV (and not mentioned in user answers), leave it out entirely — do not invent or imply it
- Where there are real matches, make those shine clearly
- Write the summary to highlight what genuinely aligns, without overclaiming on the gaps
"""
    else:
        tier_instructions = """
MATCH TIER: HIGH (70%+)
Strong match. Rewrite the CV fully and confidently to target this role. Inject ATS keywords wherever they naturally apply.
"""

    prompt = f"""You are a professional CV writer with 15 years of experience placing candidates at top companies. Rewrite this CV to be a compelling, targeted application. Think like a human recruiter reading hundreds of CVs — make this one stand out.
{tier_instructions}

TARGET ROLE: {job_role}

JOB DESCRIPTION:
{job_desc}

ANALYSIS RESULTS:
- ATS keywords to inject naturally (only where genuinely applicable): {', '.join(ats_missing) or 'none'}
- Confirmed strengths to emphasize: {', '.join(strengths) or 'see CV'}
- Hard skills required by the role: {', '.join(hard_skills) or 'see JD'}

USER'S ANSWERS (CONFIRMED REAL INFORMATION — integrate these actively into the rewrite):
{answers_formatted if answers_formatted else "No additional answers provided — rewrite from CV content only."}

ORIGINAL CV:
{sections_text}

=== YOUR REWRITING INSTRUCTIONS ===

[PROFESSIONAL SUMMARY] — Write a completely new 3-4 sentence paragraph:
  • Open with the candidate's level + target role: "Final-year [field] student / [X]-year professional targeting a [role] position..."
  • Name their 2-3 strongest relevant technical skills with specific tools mentioned in the CV
  • Include their top achievement or differentiator (from CV or user answers)
  • Close with what value they bring — confident, active voice
  • AVOID: "hardworking", "passionate learner", "team player" as empty openers

[EXPERIENCE] — For EACH position, rewrite bullet points to:
  • Begin with a strong past-tense action verb: Developed, Engineered, Analyzed, Led, Designed, Optimized, Delivered, Streamlined, Reduced, Increased, Automated, Built
  • Add specifics from user answers (numbers of people, duration, outcome) if provided
  • Inject 1-2 JD keywords per bullet where they naturally apply — never force them
  • One bullet = one achievement, not a task description

[PROJECTS] — Rewrite each project entry to:
  • Lead with what was built and what tool: "Built a [X] using [Y] to [accomplish Z]"
  • Include metrics (dataset rows, accuracy, users, time saved) IF user mentioned them in answers
  • End with the impact or insight produced

[SKILLS] — Reorder: most JD-relevant skills first. Add any skills the user confirmed in answers.

[EDUCATION] — Keep all dates and institution names EXACTLY. Minor wording improvements only.

[VOLUNTEERING/ACTIVITIES] — Strengthen action verbs. Keep facts identical.

=== HARD RULES (violation = the output is wrong) ===
• NEVER add companies, degrees, job titles, or dates not in the original CV
• NEVER invent metrics (numbers, %, timeframes) unless the user explicitly stated them in their answers above
• If user said "I don't have X" — do NOT add X anywhere
• All proper nouns (university names, company names, tool names, city names) stay exactly as written
• Only include sections that exist in the original CV

Return ONLY valid JSON (no markdown fences, no preamble, no explanation):
{{
  "summary": "new professional summary paragraph",
  "experience": "full rewritten experience section — keep all entries, strengthen all bullets",
  "projects": "full rewritten projects section",
  "skills": "reordered, enriched skills list",
  "education": "education — minor improvements only, all names/dates identical",
  "volunteering": "volunteering section",
  "courses_training": "courses — unchanged",
  "languages": "languages — unchanged",
  "improvement_log": [
    {{"type": "Improved", "section": "summary", "description": "Rewrote to target {job_role} explicitly, led with Python/SQL as primary skills", "job_requirement": "3+ years with Python and SQL"}},
    {{"type": "Added", "section": "experience", "description": "Incorporated scope from user answer: tutored 40+ students over 2 semesters", "job_requirement": "Experience explaining data to non-technical people"}},
    {{"type": "Reorganized", "section": "skills", "description": "Moved Python and SQL to top — both are core requirements in JD", "job_requirement": "Strong SQL"}}
  ]
}}

improvement_log rules:
• "section" must be one of: summary, experience, projects, skills, education, volunteering, courses_training, languages
• One entry per meaningful change, in plain simple English (the user may not be a native speaker)
• "job_requirement" quotes or closely paraphrases the job-ad requirement this change serves; use "" if it serves none"""

    raw = claude_api_call(prompt, max_tokens=4500)
    return parse_json_response(raw)


# ── FACT CHECK: nothing lost, nothing invented ──
FACT_CHECK_SECTIONS = ["experience", "education", "projects", "volunteering", "courses_training"]
_YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")


def _years(text):
    return set(_YEAR_RE.findall(text or ""))


def fact_check_rewrite(original_sections, rewritten_sections, answers_text=""):
    """Compare the AI rewrite against the original CV, section by section.

    A section is rejected (the original is kept) if the rewrite:
      - drops a year that was in the original section (a job, degree or project went missing), or
      - adds a year that is not in the original CV or the user's answers (an invented date).
    Returns (safe_sections, notes) where notes explain every section that was kept as the original.
    """
    safe = dict(rewritten_sections or {})
    notes = []
    all_original = " ".join(str(v) for v in (original_sections or {}).values())
    allowed_years = _years(all_original) | _years(answers_text)
    for key in FACT_CHECK_SECTIONS:
        orig = (original_sections or {}).get(key, "") or ""
        new = safe.get(key)
        if not orig.strip() or not isinstance(new, str) or not new.strip():
            continue
        lost = sorted(_years(orig) - _years(new))
        invented = sorted(_years(new) - allowed_years)
        if lost or invented:
            safe[key] = orig
            label = key.replace("_", " ").title()
            reason = []
            if lost:
                reason.append("dropped " + ", ".join(lost))
            if invented:
                reason.append("added dates not in your CV (" + ", ".join(invented) + ")")
            notes.append(f"{label}: kept your original text because the AI rewrite " + " and ".join(reason) + ".")
    return safe, notes


def answers_to_text(follow_up_answers):
    parts = []
    for val in (follow_up_answers or {}).values():
        if isinstance(val, dict):
            parts.append(str(val.get("answer", "")))
        elif isinstance(val, str):
            parts.append(val)
    return " ".join(parts)


# ── CHANGE REVIEW: see every AI change, why it was made, and keep, undo or edit it ──
REVIEW_SECTIONS = [
    ("summary", "Summary"), ("experience", "Experience"), ("projects", "Projects"),
    ("skills", "Skills"), ("education", "Education"), ("volunteering", "Volunteering"),
    ("courses_training", "Courses & Training"), ("languages", "Languages"),
]
_SECTION_ALIASES = {label.lower(): key for key, label in REVIEW_SECTIONS}
_SECTION_ALIASES.update({key: key for key, _ in REVIEW_SECTIONS})
_SECTION_ALIASES.update({"courses": "courses_training", "courses training": "courses_training",
                         "training": "courses_training", "professional summary": "summary"})


def _section_key(name):
    name = str(name or "").strip().lower()
    return _SECTION_ALIASES.get(name) or _SECTION_ALIASES.get(name.replace("_", " "))


def reset_review_choices():
    for k in [k for k in st.session_state.keys() if str(k).startswith(("review_use_", "review_text_"))]:
        del st.session_state[k]


def changes_by_section(improvement_log):
    grouped = {}
    for item in improvement_log or []:
        if isinstance(item, dict):
            key = _section_key(item.get("section"))
            if key:
                grouped.setdefault(key, []).append(item)
    return grouped


def apply_review_choices(base_data, ai_sections, choices, edits):
    """Build the English CV from the user's choices.

    choices[key] is "ai" or "original"; edits[key] is the user's edited text (or None).
    Sections the user didn't review keep the AI version.
    """
    chosen = {}
    for key, _ in REVIEW_SECTIONS:
        original = (base_data.get(key) or "").strip()
        ai_text = (ai_sections.get(key) or "").strip()
        text = ai_text if (choices.get(key, "ai") == "ai" and ai_text) else original
        edited = edits.get(key)
        if isinstance(edited, str) and edited.strip():
            text = edited.strip()
        if text:
            chosen[key] = text
    return build_final_cv(base_data, chosen, None)


def render_match_report(ai, score):
    """The Match step as clean HTML blocks in the same style as the website."""
    color = "var(--insert)" if score >= 70 else ("#B7791F" if score >= 35 else "var(--strike)")
    label = _esc(ai.get("match_label", "") or ("Strong match" if score >= 70 else "Partial match" if score >= 35 else "Low match"))
    rationale = _esc(ai.get("score_rationale", ""))
    strengths = ai.get("cv_strengths", []) or []
    gaps = ai.get("cv_weaknesses", []) or ai.get("critical_gaps", []) or []
    present = (ai.get("ats_keywords_present", []) or [])[:14]
    missing = (ai.get("ats_keywords_missing", []) or [])[:14]
    metrics = ai.get("metrics_missing", []) or []
    wins = ai.get("quick_wins", []) or []

    def items(lst, cls, mark, empty):
        if not lst:
            return f'<p class="rp-empty">{empty}</p>'
        return "".join(f'<li class="{cls}"><span aria-hidden="true">{mark}</span>{_esc(x)}</li>' for x in lst)

    html = f"""<style>
.rp-score{{display:grid;grid-template-columns:auto 1fr;gap:28px;align-items:center;background:var(--sheet);border:1px solid var(--rule);border-radius:12px;padding:26px 30px;margin-bottom:20px}}
.rp-score .n{{font:800 88px/0.9 var(--display);letter-spacing:-.04em}}
.rp-score .l{{font:800 24px/1.2 var(--display);margin:0 0 6px;color:var(--ink)}}
.rp-score p{{margin:0;color:var(--muted);font-size:17px;line-height:1.55}}
.rp-two{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}}
.rp-box{{background:var(--sheet);border:1px solid var(--rule);border-radius:12px;padding:20px 22px}}
.rp-box h3{{font:700 20px/1.2 var(--display);margin:0 0 12px;color:var(--ink)}}
.rp-box ul{{list-style:none;margin:0;padding:0}}
.rp-box li{{display:flex;gap:10px;font-size:16px;line-height:1.5;margin:8px 0;color:var(--ink)}}
.rp-box li span{{font-weight:700;flex-shrink:0}}
.rp-box li.ok span{{color:var(--insert)}} .rp-box li.no span{{color:var(--strike)}}
.rp-empty{{color:var(--muted);margin:0;font-size:16px}}
.rp-chip{{display:inline-block;font-size:15px;font-weight:700;padding:4px 10px;border-radius:6px;margin:0 6px 8px 0}}
.rp-chip.ok{{background:#EAF6F1;color:var(--insert)}}
.rp-chip.no{{background:var(--mark);color:var(--ink)}}
.rp-box .sub{{font-size:14px;color:var(--muted);margin:4px 0 8px;font-weight:700}}
.rp-notes li{{border-left:3px solid var(--mark);padding-left:12px}}
@media (max-width:640px){{.rp-two{{grid-template-columns:1fr}}.rp-score{{grid-template-columns:1fr}}}}
</style>
<div class="rp-score"><div class="n" style="color:{color}">{score}%</div><div><p class="l">{label}</p><p>{rationale}</p></div></div>
<div class="rp-two">
<div class="rp-box"><h3>Already fits the job</h3><ul>{items(strengths, "ok", "✓", "Nothing clear yet. The questions in the next step can help.")}</ul></div>
<div class="rp-box"><h3>Missing or weak</h3><ul>{items(gaps, "no", "✗", "No major gaps found.")}</ul></div>
</div>"""
    if present or missing:
        html += '<div class="rp-box" style="margin-bottom:20px"><h3>Words from the job ad</h3>'
        if present:
            html += '<p class="sub">Already in your CV</p><div>' + "".join(f'<span class="rp-chip ok">{_esc(k)}</span>' for k in present) + "</div>"
        if missing:
            html += '<p class="sub">Not in your CV yet. We only add them if they are true for you.</p><div>' + "".join(f'<span class="rp-chip no">{_esc(k)}</span>' for k in missing) + "</div>"
        html += "</div>"
    if metrics or wins:
        html += '<div class="rp-two">'
        if wins:
            html += f'<div class="rp-box rp-notes"><h3>Quick wins</h3><ul>{items(wins, "", "", "")}</ul></div>'
        if metrics:
            html += f'<div class="rp-box rp-notes"><h3>Lines that need a number</h3><ul>{items(metrics, "", "", "")}</ul></div>'
        html += "</div>"
    return html


def diff_html(before, after):
    """Tracked changes, readable like an editor's markup.

    Lines are compared first. A line that changed a little gets word-level marks;
    a line that was rewritten shows the old line struck through, then the new line.
    """
    import difflib

    def word_diff(old, new):
        a, b = old.split(), new.split()
        lead = old[:len(old) - len(old.lstrip())]
        parts = []
        for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
            if op == "equal":
                parts.append(_esc(" ".join(a[i1:i2])))
                continue
            if i2 > i1:
                parts.append(f'<span class="rp-del">{_esc(" ".join(a[i1:i2]))}</span>')
            if j2 > j1:
                parts.append(f'<span class="rp-ins">{_esc(" ".join(b[j1:j2]))}</span>')
        return _esc(lead) + " ".join(parts)

    def line_change(old, new):
        if not old.strip():
            return f'<span class="rp-ins">{_esc(new)}</span>'
        if not new.strip():
            return f'<span class="rp-del">{_esc(old)}</span>'
        if difflib.SequenceMatcher(None, old, new).ratio() >= 0.6:
            return word_diff(old, new)
        return f'<span class="rp-del">{_esc(old.strip())}</span> <span class="rp-ins">{_esc(new.strip())}</span>'

    a_lines = (before or "").splitlines()
    b_lines = (after or "").splitlines()
    out = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a_lines, b_lines, autojunk=False).get_opcodes():
        if op == "equal":
            out.extend(_esc(l) for l in a_lines[i1:i2])
        elif op == "delete":
            out.extend(f'<span class="rp-del">{_esc(l)}</span>' for l in a_lines[i1:i2])
        elif op == "insert":
            out.extend(f'<span class="rp-ins">{_esc(l)}</span>' for l in b_lines[j1:j2])
        else:
            old, new = a_lines[i1:i2], b_lines[j1:j2]
            for k in range(max(len(old), len(new))):
                out.append(line_change(old[k] if k < len(old) else "", new[k] if k < len(new) else ""))
    return "\n".join(out)

def render_change_review(base_data, ai_sections, improvement_log):
    """Show each changed section: why it changed, before/after, keep or undo, and edit."""
    changed = [
        (key, label) for key, label in REVIEW_SECTIONS
        if (ai_sections.get(key) or "").strip()
        and (ai_sections.get(key) or "").strip() != (base_data.get(key) or "").strip()
    ]
    if not changed:
        return build_final_cv(base_data, ai_sections, None)

    grouped = changes_by_section(improvement_log)
    st.markdown("### Review the changes")
    st.markdown('<p class="rp-legend"><span class="rp-del">Red</span> was removed, '
                '<span class="rp-ins">green</span> was added, and the yellow notes say what the job asks for.</p>',
                unsafe_allow_html=True)

    choices, edits = {}, {}
    for key, label in changed:
        reasons = grouped.get(key, [])
        use_key, text_key = f"review_use_{key}", f"review_text_{key}"
        status = "AI version" if st.session_state.get(use_key, "ai") == "ai" else "your original"
        count = f"{len(reasons)} change{'s' if len(reasons) != 1 else ''}" if reasons else "changed"
        with st.expander(f"{label}: {count}, using {status}"):
            for item in reasons:
                req = str(item.get("job_requirement", "") or "").strip()
                note = f'<b>{_esc(item.get("type", "Changed"))}</b>{_esc(item.get("description", ""))}'
                if req:
                    note += f'<br>Job asks for: <span class="req">{_esc(req)}</span>'
                st.markdown(f'<div class="rp-why">{note}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="rp-diff">{diff_html(base_data.get(key, ""), ai_sections.get(key, ""))}</div>',
                        unsafe_allow_html=True)
            choice = st.radio(
                "Which version goes in your CV?", ["ai", "original"],
                format_func=lambda v: "Keep the AI version" if v == "ai" else "Go back to my original",
                horizontal=True, key=use_key,
            )
            # Refill the edit box whenever the user switches version
            if st.session_state.get(f"{text_key}_from") != choice:
                st.session_state[text_key] = ai_sections.get(key, "") if choice == "ai" else base_data.get(key, "")
                st.session_state[f"{text_key}_from"] = choice
            st.text_area("Edit the text if you want — this is exactly what goes in your CV:",
                         key=text_key, height=180)
            choices[key] = choice
            edits[key] = st.session_state.get(text_key)
    return apply_review_choices(base_data, ai_sections, choices, edits)


# ── MATCH TIER HELPER ──
def get_match_tier(score):
    """Classify a match score into high / medium / low tier."""
    if score >= 70:
        return "high"
    if score >= 35:
        return "medium"
    return "low"


# ── LOW-MATCH: Career Gap Analysis & Roadmap ──
def call_claude_career_recommendations(cv_sections, job_role, job_desc, ai_analysis):
    """
    For low-match CVs (< 35%): instead of rewriting, produce an honest
    gap analysis and actionable career development roadmap.
    Returns structured dict or None.
    """
    sections_text = build_sections_text(cv_sections)
    missing_skills = (ai_analysis or {}).get("ats_keywords_missing", [])[:10]
    weaknesses     = (ai_analysis or {}).get("cv_weaknesses", [])[:6]
    match_score    = (ai_analysis or {}).get("match_score", 0)

    prompt = f"""You are a career counselor. This candidate's CV scored {match_score}% against a job — a low match. Do NOT rewrite or spin the CV to fit. Give honest, specific, actionable career development advice.

TARGET ROLE: {job_role}

JOB DESCRIPTION:
{job_desc}

CANDIDATE'S CURRENT CV:
{sections_text}

IDENTIFIED GAPS:
- Missing skills: {', '.join(missing_skills) or 'see analysis'}
- CV weaknesses for this role: {', '.join(weaknesses) or 'see analysis'}

Return ONLY a JSON object (no markdown, no explanation):
{{
  "why_low_match": "2-3 sentence honest explanation of the core mismatch between this CV and the job requirements",
  "critical_missing_skills": ["Most important missing skill", "Second most important", "Third"],
  "nice_to_have_missing": ["Secondary missing skill", "another"],
  "recommended_projects": [
    {{
      "title": "Project name",
      "description": "What to build and how — 1-2 sentences, specific to the job's stack",
      "skills_gained": ["skill1", "skill2"],
      "difficulty": "Beginner",
      "estimated_time": "2-4 weeks"
    }}
  ],
  "recommended_courses": [
    {{
      "name": "Course or resource name",
      "platform": "Coursera / YouTube / Udemy / freeCodeCamp / etc",
      "skills_gained": ["skill1"],
      "free": true
    }}
  ],
  "recommended_certifications": [
    {{
      "name": "Certification name",
      "why_relevant": "How it directly addresses a specific job requirement"
    }}
  ],
  "what_is_strong": "1-2 sentences about what the candidate does have going for them — be genuine",
  "realistic_timeline": "e.g., 3-6 months of focused work to become competitive for this role",
  "better_fit_roles": ["Role that better matches the current CV", "Another suitable alternative role"]
}}

Be specific and reference the actual CV content and job requirements. No generic advice."""

    raw = claude_api_call(prompt, max_tokens=2000)
    return parse_json_response(raw)


# ── STEP 9: Professional Multilingual Localization via Claude ──
def call_claude_localize_cv(cv_data, target_lang, hebrew_form="neutral"):
    """
    Step 9: Localize the CV to Hebrew or Arabic using Claude.
    Native-level professional phrasing — not word-by-word translation.
    Returns None if Claude is unavailable (no third-party fallback: CVs hold personal data).
    """
    lang_style = {
        "Hebrew": (
            "professional Israeli Hebrew (עברית מקצועית לשוק העבודה הישראלי). "
            "Use formal Israeli CV conventions and natural Hebrew phrasing — "
            "this is LOCALIZATION by a native Israeli CV writer, not a word-by-word translation."
        ),
        "Arabic": (
            "professional Modern Standard Arabic suitable for the Israeli job market. "
            "Use formal Arabic CV conventions — "
            "this is LOCALIZATION by a native Arabic CV writer, not a word-by-word translation."
        ),
    }

    sections_to_translate = {}
    for key in ["summary", "education", "experience", "projects",
                "skills", "volunteering", "courses_training", "languages"]:
        val = cv_data.get(key, "").strip()
        if val:
            sections_to_translate[key] = val

    if not sections_to_translate:
        return cv_data

    sections_json = json.dumps(sections_to_translate, ensure_ascii=False)
    hebrew_form_rule = {
        "feminine": "write about the candidate in feminine forms (e.g. פיתחה, ניהלה, אחראית)",
        "masculine": "write about the candidate in masculine forms (e.g. פיתח, ניהל, אחראי)",
    }.get(hebrew_form, "avoid gendered verb forms; use nouns and neutral phrasing (e.g. פיתוח מערכת, ניהול צוות)")

    prompt = f"""Translate and localize this CV content to {lang_style[target_lang]}.

MANDATORY RULES:
- Keep ALL technical tool names in English: Python, SQL, Excel, GitHub, Tableau, TensorFlow, etc.
- Keep ALL institution and company names in their original language
- Keep ALL dates EXACTLY as written (2022–2025, Present, etc.)
- Keep ALL numbers, percentages, and GPA values exactly
- For Hebrew: {hebrew_form_rule}
- Do NOT translate proper nouns (universities, companies, cities, person names)

CV SECTIONS TO LOCALIZE (JSON):
{sections_json}

Return the same JSON keys with localized text. Return ONLY the JSON object — no explanation, no markdown."""

    raw = claude_api_call(prompt, max_tokens=4000)
    result = parse_json_response(raw)

    if result and isinstance(result, dict):
        final = cv_data.copy()
        for key, val in result.items():
            if val and isinstance(val, str) and val.strip():
                final[key] = val
        return final

    return None

GENERIC_IGNORE_WORDS = NOISE_WORDS  # alias for backward compat

ACTION_VERBS = {
    "assisted": "Engineered","helped": "Optimized","did": "Developed",
    "worked on": "Spearheaded","responsible for": "Directed","learned": "Mastered",
    "managed": "Architected","wrote": "Formulated"
}

HEADINGS_MAP = {
    "English": ["Personal Details","Professional Summary","Education","Professional Experience",
                "Projects","Courses & Training","Volunteering / Community","Languages","Skills","Additional Information"],
    "Hebrew":  ["פרטים אישיים","תמצית מקצועית","השכלה","ניסיון תעסוקתי","פרויקטים",
                "קורסים והכשרות","התנדבות / קהילה","שפות","כישורים","מידע נוסף"],
    "Arabic":  ["البيانات الشخصية","الملخص المهني","التعليم","الخبرة العملية","المشاريع",
                "الدورات والتدريب","التطوع / المجتمع","اللغات","المهارات","معلومات إضافية"]
}

SECTION_REGEX = {
    "personal_details": r"(personal details|contact|contact info|פרטים אישיים|البيانات الشخصية)",
    "summary": r"(summary|profile|about me|professional summary|תמצית|פרופיל|תמצית מקצועית|الملخص|ملخص مهني)",
    "education": r"(education|academic|degrees|השכלה|לימודים|التعليم|المؤهلات العلمية)",
    "experience": r"(experience|work|employment|history|professional experience|ניסיון|ניסיון תעסוקתי|الخبرة|الخبرة العملية)",
    "projects": r"(projects|key projects|academic projects|פרויקטים|פרוייקטים|المشاريع)",
    "courses_training": r"(courses|training|certification|certifications|קורסים|הכשרות|הסמכות|الدورات|التدريب)",
    "volunteering": r"(volunteering|community|extracurricular|volunteer|leadership|activities|התנדבות|פעילות קהילתית|التطوع)",
    "languages": r"(languages|שפות|اللغات)",
    "skills": r"(skills|competencies|technologies|technical skills|כישורים|מיומנויות|יכולות|المهارات)"
}


# ============================================================
# CORE FUNCTIONS
# ============================================================

def detect_language(text):
    if re.search(r'[\u0590-\u05FF]', text): return "Hebrew"
    if re.search(r'[\u0600-\u06FF]', text): return "Arabic"
    return "English"


def validate_job_description(job_text):
    words = job_text.strip().split()
    if len(words) < 30:
        return False, "Please paste a real job description (at least a few sentences with requirements)."
    unique_words = {w.lower() for w in words if len(w) > 2}
    if len(unique_words) < 10:
        return False, "Please paste a real job description with responsibilities and requirements."
    return True, ""


def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        return text
    except:
        return ""


def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except:
        return ""


def classify_sections_refined(text):
    sections = {
        "personal_details":"","summary":"","education":"","experience":"",
        "projects":"","courses_training":"","volunteering":"","languages":"",
        "skills":"","additional_information":""
    }
    # Insert newlines before known section headers (handles flat PDF)
    ALL_CAPS_HEADERS = [
        "PROFESSIONAL SUMMARY","PROFESSIONAL EXPERIENCE","WORK EXPERIENCE",
        "EDUCATION","PROJECTS","TECHNICAL SKILLS","SKILLS","LANGUAGES",
        "COURSES","VOLUNTEERING","LEADERSHIP & ACTIVITIES","LEADERSHIP",
        "ACTIVITIES","CERTIFICATIONS","TRAINING","ACADEMIC BACKGROUND",
        "תמצית מקצועית","השכלה","ניסיון תעסוקתי","פרויקטים","כישורים","שפות","התנדבות",
        "الملخص المهني","التعليم","الخبرة العملية","المشاريع","المهارات","اللغات","التطوع",
    ]
    for hw in sorted(ALL_CAPS_HEADERS, key=len, reverse=True):
        text = text.replace(hw, f"\n\n{hw}\n")
    current_section = "personal_details"
    lines = text.split('\n')
    for line in lines:
        clean_line = line.strip()
        if not clean_line: continue
        low_line = clean_line.lower()
        found_header = False
        if len(clean_line) < 60:
            for key, pattern in SECTION_REGEX.items():
                # Require the pattern to be the whole line (allow leading/trailing
                # decoration like dashes, colons, whitespace — but no other words).
                if re.search(r'^[\s\-=*•·_|]*' + pattern + r'[\s\-=*•·_|:]*$', low_line, re.IGNORECASE):
                    current_section = key
                    found_header = True
                    break
        if not found_header:
            sections[current_section] += clean_line + "\n"
    # Trim personal_details to just name + contact (max 5 lines)
    pd_lines = [l for l in sections["personal_details"].split('\n') if l.strip()]
    if len(pd_lines) > 5:
        sections["personal_details"] = "\n".join(pd_lines[:5])
    for k in sections:
        sections[k] = re.sub(r'\n{3,}', '\n\n', sections[k].strip())
    return sections


def format_contact_info(text):
    if not text: return "CANDIDATE", ""
    # Normalize separators
    text_clean = re.sub(r'\s*[|·•]\s*', '\n', text)
    lines = [l.strip() for l in text_clean.split('\n') if l.strip()]
    name, contact_parts = "", []
    for i, line in enumerate(lines):
        low = line.lower()
        is_contact = (any(k in low for k in ['email','phone','@','linkedin','github','israel','jerusalem','city','tel aviv'])
                      or re.search(r'\d{3}', line) or '@' in line)
        if i == 0 and not is_contact and len(line.split()) <= 6:
            name = line
        else:
            contact_parts.append(line)
    if not name and lines:
        m = re.match(r'^([A-Za-z\u0590-\u05FF\u0600-\u06FF ]{2,40}?)(?=\s+\S*[@\d])', text)
        name = m.group(1).strip() if m else lines[0]
        if not contact_parts: contact_parts = lines[1:]
    contact_str = " | ".join(contact_parts)
    contact_str = re.sub(r'\b(Email|Phone|Tel|Mobile|LinkedIn|Address)\s*:\s*', '', contact_str, flags=re.IGNORECASE)
    return name.strip(), contact_str.strip()


def group_skills_segregated(skills_text):
    if not skills_text: return {"PROFESSIONAL": [], "TECHNICAL": []}
    # Remove subheading labels that bleed in from PDF
    cleaned = re.sub(r'(Technical Skills|Business & Professional Skills|Professional Skills)\s*[,\n]?', '', skills_text, flags=re.IGNORECASE)
    skills = [s.strip() for s in re.split(r'[,\n•|]', cleaned) if s.strip() and len(s.strip()) > 1]
    tech_kw = set(list(SKILL_KEYWORDS) + ["sql","python","aws","docker","git","react","tableau","powerbi",
                                           "excel","data analysis","r","networkx","pandas","numpy","scipy",
                                           "matplotlib","unix","bash","bigquery","spss","statistical modeling",
                                           "machine learning","network analysis","rstudio","bigquery"])
    technical, professional = [], []
    for s in skills:
        sl = s.lower()
        if any(kw in sl for kw in tech_kw):
            technical.append(s)
        else:
            professional.append(s)
    return {"PROFESSIONAL": list(dict.fromkeys(professional)), "TECHNICAL": list(dict.fromkeys(technical))}


def build_final_cv(cv_sections, adjusted_sections, job_role):
    """
    Merges adjusted sections back into cv_sections.
    Falls back to original section if adjustment is missing or empty.
    Always preserves: personal_details exactly as-is.
    """
    final = cv_sections.copy()
    if adjusted_sections:
        for key, val in adjusted_sections.items():
            if val and val.strip() and key != "personal_details":
                final[key] = val.strip()
    # Always keep personal_details untouched
    final["personal_details"] = cv_sections.get("personal_details", "")
    return final


def split_cv_entries(text):
    """Split flat CV text into separate entries, handling inline • bullets."""
    if not text: return []
    # Expand inline bullets to newlines
    text = re.sub(r'\s*•\s*', '\n• ', text)
    # If double newlines exist, use them
    if '\n\n' in text:
        return [b.strip() for b in text.split('\n\n') if b.strip()]
    # Otherwise split: new entry starts when a non-bullet line follows bullet lines
    lines = text.split('\n')
    blocks, current = [], []
    for line in lines:
        stripped = line.strip()
        if not stripped: continue
        is_bullet = stripped.startswith(('•','-','*','·'))
        if current and not is_bullet and not current[-1].startswith('•'):
            blocks.append('\n'.join(current))
            current = [stripped]
        else:
            current.append(stripped)
    if current: blocks.append('\n'.join(current))
    return [b for b in blocks if b.strip()] or [text]


def render_entry(block, show_subtitle=True):
    """Render one CV entry block → HTML."""
    block = re.sub(r'\s*•\s*', '\n• ', block)
    lines = [l.strip() for l in block.split('\n') if l.strip()]
    if not lines: return ""
    title = lines[0].lstrip('•-* ')
    html_out = f'<div class="cv-entry"><div class="cv-entry-title">{_esc(title)}</div>'
    subtitle_done = False
    bullets = []
    for line in lines[1:]:
        is_bullet = line.startswith(('•','-','*'))
        clean = line.lstrip('•-* ').strip()
        if not clean: continue
        if is_bullet:
            bullets.append(clean)
        elif not subtitle_done and show_subtitle:
            html_out += f'<div class="cv-entry-sub">{_esc(clean)}</div>'
            subtitle_done = True
        else:
            bullets.append(clean)
    if bullets:
        html_out += '<ul class="cv-bullets">' + ''.join(f'<li>{_esc(b)}</li>' for b in bullets) + '</ul>'
    html_out += '</div>'
    return html_out


def extract_name_and_contact(personal_details_text):
    """
    Robustly extracts name and contact line from personal_details.
    Handles flat PDF text, structured text, and pipe-separated formats.
    Returns (name, contact_line) — name is NEVER repeated in contact_line.
    """
    text = personal_details_text.strip()
    if not text:
        return "CANDIDATE", ""

    CONTACT_KEYWORDS = ['email', 'phone', 'tel', 'mobile', '@', 'linkedin',
                        'github', 'israel', 'jerusalem', 'city', 'address',
                        'ירושלים', 'תל אביב', 'القدس', 'إسرائيل']

    # ── Strategy 1: newline-separated (clean PDF/manual) ──
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if len(lines) >= 2:
        first = lines[0]
        low = first.lower()
        has_contact = any(k in low for k in CONTACT_KEYWORDS) or '@' in first
        is_name_length = len(first.split()) <= 5 and len(first) <= 45
        if not has_contact and is_name_length:
            name = first
            # Contact = everything after first line, joined
            rest = " | ".join(lines[1:])
            rest = re.sub(r'\b(Email|Phone|Tel|Mobile|E-mail)\s*:\s*', '', rest, flags=re.IGNORECASE)
            # Remove name if it appears at start of contact line
            if rest.upper().startswith(name.upper()):
                rest = rest[len(name):].strip(' |·,')
            contact = rest.strip(' |·')
            return name.strip(), contact

    # ── Strategy 2: flat single line — extract name from beginning ──
    # Pattern: "FIRSTNAME LASTNAME City, Country Email: ... Phone: ..."
    flat = text.replace('\n', ' ')

    # Try: name = first all-caps word(s) before a city/country keyword
    m = re.match(
        r'^([A-Z][A-Z\s]{2,30}?)(?=\s+(?:Jerusalem|Israel|Tel Aviv|Haifa|Email|Phone|\d{3}|ירושלים|القدس))',
        flat
    )
    if m:
        name = m.group(1).strip()
        rest = flat[len(name):].strip()
        rest = re.sub(r'\b(Email|Phone|Tel|Mobile|E-mail)\s*:\s*', '', rest, flags=re.IGNORECASE)
        # Remove leading duplicate of name
        if rest.upper().startswith(name.upper()):
            rest = rest[len(name):].strip(' |·,')
        # Build clean contact from rest
        parts = re.split(r'\s*[|·•]\s*', rest)
        parts = [p.strip() for p in parts if p.strip() and p.strip().upper() != name.upper()]
        return name, " | ".join(parts)

    # ── Strategy 3: pipe-separated "Name | City | Email | Phone" ──
    if '|' in flat:
        parts = [p.strip() for p in flat.split('|')]
        if parts:
            name = parts[0]
            contact = " | ".join(parts[1:])
            contact = re.sub(r'\b(Email|Phone|Tel|Mobile)\s*:\s*', '', contact, flags=re.IGNORECASE)
            return name.strip(), contact.strip()

    # ── Strategy 4: last resort — first 2-3 words as name ──
    words = flat.split()
    name = ' '.join(words[:3]) if len(words) >= 3 else flat
    rest = ' '.join(words[3:]) if len(words) > 3 else ''
    rest = re.sub(r'\b(Email|Phone|Tel|Mobile)\s*:\s*', '', rest, flags=re.IGNORECASE)
    return name.strip(), rest.strip()


def render_cv_html(cv_data, lang="English"):
    """
    Renders CV as clean Israeli-professional HTML.
    Section order: Summary → Education → Technical Skills → Projects → Experience → Volunteering → Languages
    """
    is_rtl = lang in ["Hebrew", "Arabic"]
    dir_attr = 'dir="rtl"' if is_rtl else 'dir="ltr"'
    text_align = 'right' if is_rtl else 'left'
    h = HEADINGS_MAP[lang]

    name, contact = extract_name_and_contact(cv_data.get("personal_details", ""))
    if not name: name = "CANDIDATE"

    # ── CSS embedded in the CV page ──
    cv_css = f"""
    <style>
    .isr-cv {{
        background: white;
        max-width: 780px;
        margin: 0 auto;
        padding: 40px 50px;
        font-family: 'Calibri', 'Arial', sans-serif;
        font-size: 10.5pt;
        color: #111;
        line-height: 1.45;
        direction: {('rtl' if is_rtl else 'ltr')};
        text-align: {text_align};
        box-shadow: 0 4px 24px rgba(0,0,0,0.10);
    }}
    .isr-name {{
        font-size: 20pt;
        font-weight: 700;
        text-align: center;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
        color: #16213A;
        text-transform: uppercase;
    }}
    .isr-contact {{
        font-size: 10pt;
        text-align: center;
        color: #334155;
        margin-bottom: 10px;
    }}
    .isr-divider {{
        border: none;
        border-top: 1.5px solid #16213A;
        margin: 8px 0 14px 0;
    }}
    .isr-section-title {{
        font-size: 11.5pt;
        font-weight: 700;
        color: #16213A;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        border-bottom: 1px solid #94a3b8;
        padding-bottom: 2px;
        margin-top: 14px;
        margin-bottom: 6px;
    }}
    .isr-entry {{ margin-bottom: 8px; }}
    .isr-entry-header {{
        display: flex;
        justify-content: space-between;
        align-items: baseline;
        flex-direction: {'row-reverse' if is_rtl else 'row'};
    }}
    .isr-entry-title {{ font-weight: 700; font-size: 10.5pt; color: #16213A; }}
    .isr-entry-date  {{ font-size: 9.5pt; color: #5A6478; font-weight: 600; white-space: nowrap; }}
    .isr-entry-sub   {{ font-size: 10pt; color: #5A6478; font-style: italic; margin-bottom: 3px; }}
    .isr-bullets     {{ margin: 3px 0; {'padding-right:16px;padding-left:0;' if is_rtl else 'padding-left:16px;'} }}
    .isr-bullets li  {{ font-size: 10pt; color: #16213A; margin-bottom: 2px; }}
    .isr-skills-row  {{ margin-bottom: 4px; font-size: 10pt; }}
    .isr-skills-label{{ font-weight: 700; color: #16213A; }}
    </style>
    """

    def sec(title):
        return f'<div class="isr-section-title">{_esc(title)}</div>'

    def parse_date_from_title(title):
        """Extract trailing date from title like 'Company Name | 2025–Present'"""
        m = re.search(r'[\|｜]\s*(\d{4}[–\-]\w+|\d{4})\s*$', title)
        if m:
            date = m.group(1)
            clean_title = title[:m.start()].strip()
            return clean_title, date
        return title, ""

    def render_entries_html(content, show_date=True):
        out = ""
        for block in split_cv_entries(content):
            block = re.sub(r'\s*•\s*', '\n• ', block)
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            if not lines: continue

            raw_title = lines[0].lstrip('•-* ')
            title, date = parse_date_from_title(raw_title)

            out += '<div class="isr-entry">'
            out += '<div class="isr-entry-header">'
            out += f'<span class="isr-entry-title">{_esc(title)}</span>'
            if date:
                out += f'<span class="isr-entry-date">{_esc(date)}</span>'
            out += '</div>'

            subtitle_done = False
            bullets = []
            for line in lines[1:]:
                is_bullet = line.startswith(('•','-','*'))
                clean = line.lstrip('•-* ').strip()
                if not clean: continue
                if is_bullet:
                    bullets.append(clean)
                elif not subtitle_done:
                    out += f'<div class="isr-entry-sub">{_esc(clean)}</div>'
                    subtitle_done = True
                else:
                    bullets.append(clean)

            if bullets:
                out += '<ul class="isr-bullets">' + ''.join(f'<li>{_esc(b)}</li>' for b in bullets) + '</ul>'
            out += '</div>'
        return out

    # ── BUILD HTML ──
    html = f'{cv_css}<div class="isr-cv" {dir_attr}>'

    # Header
    html += f'<div class="isr-name">{_esc(name)}</div>'
    if contact:
        html += f'<div class="isr-contact">{_esc(contact)}</div>'
    html += '<hr class="isr-divider">'

    # 1. SUMMARY
    if cv_data.get("summary"):
        html += sec(h[1])
        summary = cv_data["summary"].replace('\n', ' ').strip()
        html += f'<div style="font-size:10.5pt;color:#16213A;line-height:1.5;margin-bottom:8px;">{_esc(summary)}</div>'

    # 2. EDUCATION
    if cv_data.get("education"):
        html += sec(h[2])
        edu = re.sub(r'\s*•\s*', '\n• ', cv_data["education"])
        # Split on lines that start with a capital and look like institution names
        edu_blocks = re.split(r'\n(?=[A-Z\u0590-\u05FF\u0600-\u06FF])', edu)
        for block in edu_blocks:
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            if not lines: continue
            raw_title = lines[0].lstrip('•-* ')
            title, date = parse_date_from_title(raw_title)
            html += '<div class="isr-entry">'
            html += '<div class="isr-entry-header">'
            html += f'<span class="isr-entry-title">{_esc(title)}</span>'
            if date: html += f'<span class="isr-entry-date">{_esc(date)}</span>'
            html += '</div>'
            for l in lines[1:]:
                clean = l.lstrip('•-* ').strip()
                if clean:
                    html += f'<div class="isr-entry-sub">{_esc(clean)}</div>'
            html += '</div>'

    # 3. EXPERIENCE
    if cv_data.get("experience"):
        html += sec(h[3])
        html += render_entries_html(cv_data["experience"])

    # 4. PROJECTS
    if cv_data.get("projects"):
        html += sec(h[4])
        html += render_entries_html(cv_data["projects"])

    # 5. SKILLS
    raw_skills = cv_data.get("skills", "")
    if raw_skills:
        raw_skills_clean = re.sub(
            r'(Technical Skills|Business & Professional Skills|Professional Skills|'
            r'כישורים טכניים|כישורים מקצועיים|المهارات التقنية|المهارات المهنية)\s*[,\n]?',
            '', raw_skills, flags=re.IGNORECASE)
        skills_data = group_skills_segregated(raw_skills_clean)

        # Skills section title
        skills_title = h[8] if len(h) > 8 else ("כישורים" if lang == "Hebrew" else "المهارات" if lang == "Arabic" else "Skills")
        html += sec(skills_title)

        if skills_data["TECHNICAL"]:
            tech_label = ("תכנות וכלים" if lang == "Hebrew" else "البرمجة والأدوات" if lang == "Arabic" else "Technical")
            html += f'<div class="isr-skills-row"><span class="isr-skills-label">{tech_label}:</span> {" · ".join(_esc(x) for x in skills_data["TECHNICAL"])}</div>'
        if skills_data["PROFESSIONAL"]:
            prof_label = ("כישורים מקצועיים" if lang == "Hebrew" else "المهارات المهنية" if lang == "Arabic" else "Professional")
            html += f'<div class="isr-skills-row"><span class="isr-skills-label">{prof_label}:</span> {" · ".join(_esc(x) for x in skills_data["PROFESSIONAL"])}</div>'

    # 6. VOLUNTEERING / AWARDS / LEADERSHIP
    if cv_data.get("volunteering"):
        html += sec(h[6])
        vol = re.sub(r'^\s*[&\*]\s*$', '', cv_data["volunteering"], flags=re.MULTILINE)
        html += render_entries_html(vol)

    # 7. LANGUAGES
    if cv_data.get("languages"):
        html += sec(h[7])
        langs = re.sub(r'\s*•\s*', ' · ', cv_data["languages"]).strip(' ·')
        html += f'<div style="font-size:10pt;color:#16213A;">{_esc(langs)}</div>'

    # 8. COURSES (if exists)
    if cv_data.get("courses_training"):
        html += sec(h[5])
        items = [c.strip() for c in re.split(r'[,\n]', cv_data["courses_training"]) if c.strip() and len(c.strip()) > 3]
        html += '<ul class="isr-bullets">' + ''.join(f'<li>{i}</li>' for i in items) + '</ul>'

    html += '</div>'

    # IMPORTANT: strip leading whitespace from every line before returning.
    # Streamlit's markdown renderer treats any line starting with 4+ spaces
    # as an indented code block (standard CommonMark behaviour), which is
    # exactly what was happening here — the CSS/HTML was being shown as
    # raw text instead of being rendered, because it started with 4 spaces
    # of indentation. Stripping leading whitespace is harmless for
    # HTML/CSS (browsers ignore it outside <pre> tags) and fixes the bug.
    html = re.sub(r'(?m)^[ \t]+', '', html)
    return html


def generate_docx(cv_data, output_lang):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()
    for sec in doc.sections:
        sec.top_margin  = docx.shared.Cm(1.8)
        sec.bottom_margin = docx.shared.Cm(1.8)
        sec.left_margin  = docx.shared.Cm(2.0)
        sec.right_margin = docx.shared.Cm(2.0)

    is_rtl = output_lang in ["Hebrew","Arabic"]
    h = HEADINGS_MAP[output_lang]
    font_name = 'Arial' if is_rtl else 'Calibri'

    style = doc.styles['Normal']
    style.font.name = font_name
    style.font.size = Pt(10.5)

    def set_para_format(para, rtl=False):
        if rtl:
            pPr = para._p.get_or_add_pPr()
            bidi = OxmlElement('w:bidi')
            pPr.append(bidi)
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'right')
            pPr.append(jc)

    def add_para(text, bold=False, italic=False, size=10.5, align="left", color_rgb=None, space_before=0, space_after=2):
        p = doc.add_paragraph()
        if align == "center": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right" or is_rtl: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_para_format(p, is_rtl)
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after  = Pt(space_after)
        run = p.add_run(str(text))
        run.bold = bold; run.italic = italic
        run.font.name = font_name
        run.font.size = Pt(size)
        if color_rgb:
            run.font.color.rgb = RGBColor(*color_rgb)
        return p

    def add_section_heading(title):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT
        set_para_format(p, is_rtl)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after  = Pt(2)
        run = p.add_run(title.upper())
        run.bold = True; run.font.size = Pt(11)
        run.font.name = font_name
        run.font.color.rgb = RGBColor(30, 41, 59)
        # Bottom border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '4')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '94a3b8')
        pBdr.append(bottom)
        pPr.append(pBdr)

    def add_bullet_item(text):
        p = doc.add_paragraph(style='List Bullet')
        set_para_format(p, is_rtl)
        if is_rtl: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(text.lstrip('•-* ').strip())
        run.font.size = Pt(10); run.font.name = font_name
        p.paragraph_format.space_after = Pt(1)

    def parse_date_from_title(title):
        m = re.search(r'[\|｜]\s*(\d{4}[–\-]\w+|\d{4})\s*$', title)
        if m:
            return title[:m.start()].strip(), m.group(1)
        return title, ""

    # ── NAME ──
    name, contact = extract_name_and_contact(cv_data.get("personal_details",""))
    if not name: name = "CANDIDATE"
    name_p = add_para(name.upper(), bold=True, size=18, align="center", space_after=2)
    if contact:
        contact_clean = re.sub(r'\b(Email|Phone|Tel)\s*:\s*', '', contact, flags=re.IGNORECASE)
        add_para(contact_clean, size=10, align="center", space_after=4)

    # Divider line (using a paragraph with border)
    div_p = doc.add_paragraph()
    div_p.paragraph_format.space_before = Pt(2)
    div_p.paragraph_format.space_after  = Pt(4)
    pPr = div_p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'12')
    bot.set(qn('w:space'),'1'); bot.set(qn('w:color'),'0f172a')
    pBdr.append(bot); pPr.append(pBdr)

    # ── SUMMARY ──
    if cv_data.get("summary"):
        add_section_heading(h[1])
        add_para(cv_data["summary"].replace('\n',' ').strip(), size=10.5)

    # ── EDUCATION ──
    if cv_data.get("education"):
        add_section_heading(h[2])
        edu = re.sub(r'\s*•\s*', '\n• ', cv_data["education"])
        edu_blocks = re.split(r'\n(?=[A-Z\u0590-\u05FF\u0600-\u06FF])', edu)
        for block in edu_blocks:
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            if not lines: continue
            raw_title = lines[0].lstrip('•-* ')
            title, date = parse_date_from_title(raw_title)
            entry_text = f"{title}   {date}" if date else title
            add_para(entry_text, bold=True, size=10.5, space_before=4, space_after=1)
            for l in lines[1:]:
                clean = l.lstrip('•-* ').strip()
                if clean: add_para(clean, italic=True, size=10, space_after=1)

    # ── SKILLS ──
    raw_skills = cv_data.get("skills","")
    if raw_skills:
        raw_skills_clean = re.sub(
            r'(Technical Skills|Business & Professional Skills|Professional Skills|'
            r'כישורים טכניים|כישורים מקצועיים|المهارات التقنية|المهارات المهنية)\s*[,\n]?',
            '', raw_skills, flags=re.IGNORECASE)
        skills_data = group_skills_segregated(raw_skills_clean)
        skills_title = h[8] if len(h) > 8 else ("כישורים" if output_lang=="Hebrew" else "المهارات" if output_lang=="Arabic" else "Skills")
        add_section_heading(skills_title)
        if skills_data["TECHNICAL"]:
            tech_label = "תכנות וכלים" if output_lang=="Hebrew" else "البرمجة والأدوات" if output_lang=="Arabic" else "Technical"
            p = doc.add_paragraph()
            set_para_format(p, is_rtl)
            r1 = p.add_run(f"{tech_label}: "); r1.bold=True; r1.font.size=Pt(10); r1.font.name=font_name
            r2 = p.add_run(" · ".join(skills_data["TECHNICAL"])); r2.font.size=Pt(10); r2.font.name=font_name
        if skills_data["PROFESSIONAL"]:
            prof_label = "כישורים מקצועיים" if output_lang=="Hebrew" else "المهارات المهنية" if output_lang=="Arabic" else "Professional"
            p = doc.add_paragraph()
            set_para_format(p, is_rtl)
            r1 = p.add_run(f"{prof_label}: "); r1.bold=True; r1.font.size=Pt(10); r1.font.name=font_name
            r2 = p.add_run(" · ".join(skills_data["PROFESSIONAL"])); r2.font.size=Pt(10); r2.font.name=font_name

    # ── PROJECTS ──
    if cv_data.get("projects"):
        add_section_heading(h[4])
        for block in split_cv_entries(cv_data["projects"]):
            block = re.sub(r'\s*•\s*', '\n• ', block)
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            if not lines: continue
            raw_title = lines[0].lstrip('•-* ')
            title, date = parse_date_from_title(raw_title)
            entry_text = f"{title}   {date}" if date else title
            add_para(entry_text, bold=True, size=10.5, space_before=4, space_after=1)
            for line in lines[1:]:
                clean = line.lstrip('•-* ').strip()
                if not clean: continue
                if line.startswith(('•','-','*')): add_bullet_item(clean)
                else: add_para(clean, italic=True, size=10, space_after=1)

    # ── EXPERIENCE ──
    if cv_data.get("experience"):
        add_section_heading(h[3])
        for block in split_cv_entries(cv_data["experience"]):
            block = re.sub(r'\s*•\s*', '\n• ', block)
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            if not lines: continue
            raw_title = lines[0].lstrip('•-* ')
            title, date = parse_date_from_title(raw_title)
            entry_text = f"{title}   {date}" if date else title
            add_para(entry_text, bold=True, size=10.5, space_before=4, space_after=1)
            sub_done = False
            for line in lines[1:]:
                clean = line.lstrip('•-* ').strip()
                if not clean: continue
                if line.startswith(('•','-','*')): add_bullet_item(clean)
                elif not sub_done:
                    add_para(clean, italic=True, size=10, space_after=1); sub_done=True
                else: add_bullet_item(clean)

    # ── VOLUNTEERING ──
    if cv_data.get("volunteering"):
        add_section_heading(h[6])
        vol = re.sub(r'^\s*[&\*]\s*$','', cv_data["volunteering"], flags=re.MULTILINE)
        for block in split_cv_entries(vol):
            block = re.sub(r'\s*•\s*', '\n• ', block)
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            if not lines: continue
            raw_title = lines[0].lstrip('•-* ')
            title, date = parse_date_from_title(raw_title)
            add_para(f"{title}   {date}" if date else title, bold=True, size=10.5, space_before=4, space_after=1)
            for line in lines[1:]:
                clean = line.lstrip('•-* ').strip()
                if clean and line.startswith(('•','-','*')): add_bullet_item(clean)
                elif clean: add_para(clean, size=10, space_after=1)

    # ── LANGUAGES ──
    if cv_data.get("languages"):
        add_section_heading(h[7])
        langs_clean = re.sub(r'\s*•\s*', ' · ', cv_data["languages"]).strip(' ·')
        add_para(langs_clean, size=10)

    # ── COURSES ──
    if cv_data.get("courses_training"):
        add_section_heading(h[5])
        items = [c.strip() for c in re.split(r'[,\n]', cv_data["courses_training"]) if c.strip() and len(c.strip()) > 3]
        for item in items: add_bullet_item(item)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

    is_rtl = output_lang in ["Hebrew", "Arabic"]
    h = HEADINGS_MAP[output_lang]

    # Set default font
    style = doc.styles['Normal']
    style.font.name = 'Arial' if is_rtl else 'Calibri'
    style.font.size = Pt(11)

    def set_rtl_para(para):
        """Apply RTL formatting to a paragraph."""
        if is_rtl:
            pPr = para._p.get_or_add_pPr()
            bidi = OxmlElement('w:bidi')
            pPr.append(bidi)
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'right')
            pPr.append(jc)

    def add_para(text, bold=False, italic=False, size=11, align="left", color=None):
        p = doc.add_paragraph()
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right" or is_rtl:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_rtl_para(p)
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor(*color)
        return p

    def add_section_heading(title):
        p = add_para(title.upper(), bold=True, size=10, color=(30, 41, 59))
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(2)
        # Add bottom border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '4')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '94a3b8')
        pBdr.append(bottom)
        pPr.append(pBdr)
        return p

    def add_bullet(text):
        p = doc.add_paragraph(style='List Bullet')
        set_rtl_para(p)
        if is_rtl:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(text.lstrip('•-* ').strip())
        run.font.size = Pt(10)
        return p

    # ── NAME & CONTACT ──
    name, contact = format_contact_info(cv_data.get("personal_details",""))
    if not name: name = "CANDIDATE"
    name_p = add_para(name.upper(), bold=True, size=20, align="center")
    name_p.paragraph_format.space_after = Pt(2)
    if contact:
        contact_clean = re.sub(r'(Email|Phone|Tel)\s*:\s*', '', contact, flags=re.IGNORECASE)
        add_para(contact_clean, size=10, align="center")

    # ── SECTIONS ──
    section_order = [
        ("summary",          h[1]),
        ("education",        h[2]),
        ("experience",       h[3]),
        ("projects",         h[4]),
        ("courses_training", h[5]),
        ("volunteering",     h[6]),
        ("languages",        h[7]),
    ]

    for key, heading in section_order:
        content = cv_data.get(key, "").strip()
        if not content:
            continue

        add_section_heading(heading)

        if key in ["experience", "projects", "volunteering"]:
            for block in split_cv_entries(content):
                block = re.sub(r'\s*•\s*', '\n• ', block)
                lines = [l.strip() for l in block.split('\n') if l.strip()]
                for i, line in enumerate(lines):
                    clean = line.lstrip('•-* ').strip()
                    if not clean:
                        continue
                    if i == 0:
                        add_para(clean, bold=True, size=10.5)
                    elif line.startswith(('•', '-', '*')):
                        add_bullet(clean)
                    else:
                        add_para(clean, italic=True, size=10)

        elif key == "education":
            # Each school on its own
            edu_entries = re.split(r'\n(?=[A-Z])', re.sub(r'\s*•\s*', '\n• ', content))
            for entry in edu_entries:
                lines = [l.strip() for l in entry.split('\n') if l.strip()]
                if not lines:
                    continue
                add_para(lines[0].lstrip('•-* '), bold=True, size=10.5)
                for l in lines[1:]:
                    add_para(l.lstrip('•-* '), italic=True, size=10)

        elif key == "courses_training":
            items = [c.strip() for c in re.split(r'[,\n]', content) if c.strip() and len(c.strip()) > 3]
            for item in items:
                add_bullet(item)

        elif key == "languages":
            langs_clean = re.sub(r'\s*•\s*', ' · ', content).strip(' ·')
            add_para(langs_clean, size=10)

        elif key == "summary":
            summary = content.replace('\n', ' ').strip()
            add_para(summary, size=10.5)

        else:
            add_para(content, size=10)

    # ── SKILLS ──
    raw_skills = cv_data.get("skills", "")
    if raw_skills:
        raw_skills = re.sub(r'(Technical Skills|Business & Professional Skills|Professional Skills)\s*[,\n]?', '', raw_skills, flags=re.IGNORECASE)
        skills_data = group_skills_segregated(raw_skills)

        if skills_data["TECHNICAL"]:
            add_section_heading(h[8] if len(h) > 8 else "Technical Skills")
            add_para(" · ".join(skills_data["TECHNICAL"]), size=10)
        if skills_data["PROFESSIONAL"]:
            add_section_heading("Professional Skills")
            add_para(" · ".join(skills_data["PROFESSIONAL"]), size=10)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# NAVIGATION HELPERS
# ============================================================

TAB_NAMES = ["The job", "Your CV", "Match", "Questions", "Review", "Download"]
TAB_ICONS = ["", "", "", "✏️", "👁️", "⬇️"]

def go_to(tab_index):
    st.session_state.active_tab = tab_index
    if tab_index < 2:
        # Back to Job or CV Content — wipe deep analysis so it re-runs
        st.session_state.ai_cv_analysis = {}
        st.session_state.analysis_done = False
    if tab_index < 3:
        # Back before Refinement — wipe dynamic questions
        st.session_state.dynamic_questions = []
    if tab_index < 4:
        # Back before Preview — wipe rewrite output and career recs
        st.session_state.cv_adjusted = False
        st.session_state.adjusted_cv_data = {}
        st.session_state.final_cv_data = {}
        st.session_state.improvement_log = []
        st.session_state.career_recommendations = {}
    st.rerun()

LOGO_SVG = """<svg width="32" height="32" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect x="5" y="2" width="18" height="24" rx="2" stroke="#16213A" stroke-width="1.8" fill="white"/>
  <line x1="9" y1="9" x2="19" y2="9" stroke="#DCE2EC" stroke-width="1.2"/>
  <line x1="9" y1="13" x2="19" y2="13" stroke="#DCE2EC" stroke-width="1.2"/>
  <line x1="9" y1="17" x2="15" y2="17" stroke="#DCE2EC" stroke-width="1.2"/>
  <circle cx="22" cy="22" r="7" fill="white" stroke="#16213A" stroke-width="1.8"/>
  <line x1="22" y1="17" x2="22" y2="22" stroke="#16213A" stroke-width="2" stroke-linecap="round"/>
  <line x1="22" y1="22" x2="25.5" y2="25.5" stroke="#16213A" stroke-width="2" stroke-linecap="round"/>
  <circle cx="22" cy="22" r="1.2" fill="#16213A"/>
</svg>"""

LOGO_SVG_SM = """<svg width="24" height="24" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect x="5" y="2" width="18" height="24" rx="2" stroke="#16213A" stroke-width="1.8" fill="white"/>
  <circle cx="22" cy="22" r="7" fill="white" stroke="#16213A" stroke-width="1.8"/>
  <line x1="22" y1="17" x2="22" y2="22" stroke="#16213A" stroke-width="2" stroke-linecap="round"/>
  <line x1="22" y1="22" x2="25.5" y2="25.5" stroke="#16213A" stroke-width="2" stroke-linecap="round"/>
  <circle cx="22" cy="22" r="1.2" fill="#16213A"/>
</svg>"""


SITE_URL = "https://leenabuleil04-ctrl.github.io/resumepilot/"


def render_header():
    st.markdown(f"""
    <div class="rp-header">
        <a class="rp-logo" href="{SITE_URL}" target="_self"><span class="rp-logo-mark"></span><span class="rp-brand">ResumePilot</span></a>
        <span class="rp-tagline"><a href="{SITE_URL}#how" target="_self">How it works</a></span>
    </div>
    """, unsafe_allow_html=True)


def render_step_bar(current):
    html = '<nav class="rp-step-bar" aria-label="Progress">'
    for i, name in enumerate(TAB_NAMES):
        state = "done" if i < current else ("active" if i == current else "todo")
        num = "✓" if i < current else str(i + 1)
        aria = ' aria-current="step"' if i == current else ""
        html += f'<div class="rp-step-item {state}"{aria}><div class="rp-step-num">{num}</div><div class="rp-step-name">{_esc(name)}</div></div>'
    html += '</nav>'
    st.markdown(html, unsafe_allow_html=True)


def nav_buttons(current, can_proceed=True, proceed_label="Continue", back_label="Back"):
    st.markdown("<div style='margin-top:36px; padding-top:12px; border-top:1px solid #DCE2EC;'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.2, 2.6, 2.2])
    with col1:
        if current > 0:
            if st.button(back_label, key=f"back_{current}", use_container_width=True):
                go_to(current - 1)
    with col3:
        if current < len(TAB_NAMES) - 1:
            btn = st.button(
                proceed_label,
                key=f"next_{current}",
                use_container_width=True,
                disabled=not can_proceed,
                type="primary"
            )
            if btn:
                go_to(current + 1)


# ============================================================
# APP HEADER + LANDING PAGE
# ============================================================

# Otherwise show the app
render_header()
current_tab = st.session_state.active_tab
render_step_bar(current_tab)
st.markdown("<div style='margin-bottom:24px;'></div>", unsafe_allow_html=True)


# ============================================================
# TAB 0 — THE JOB
# ============================================================

if current_tab == 0:
    st.markdown('<div class="rp-page-badge">Step 1 of 6</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-title">Which job are you applying for?</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-sub">Paste the job ad. Everything that follows is tailored to it.</div>', unsafe_allow_html=True)

    job_role_input = st.text_input(
        "Job title",
        value=st.session_state.job_role,
        placeholder="e.g., Data Analyst Intern, Software Engineer, Marketing Associate",
        key="job_role_input"
    )
    st.session_state.job_role = job_role_input

    st.session_state.job_desc = st.text_area(
        "Job ad",
        value=st.session_state.job_desc,
        height=280,
        placeholder="Paste the full job description here — include responsibilities, requirements, and qualifications..."
    )

    col_demo, _ = st.columns([1, 3])
    with col_demo:
        if st.button("Load demo job", use_container_width=True):
            st.session_state.job_role = "Junior Data Analyst"
            st.session_state.job_desc = """We are looking for a Junior Data Analyst to join our growing team.

Responsibilities:
- Analyze large datasets using Python and SQL
- Build dashboards and reports using Tableau or Power BI
- Apply statistical methods including regression and modeling
- Collaborate with cross-functional teams to deliver insights
- Present findings to stakeholders

Requirements:
- Bachelor's degree in Statistics, Mathematics, Computer Science or related field
- Proficiency in Python and SQL
- Experience with data visualization tools (Tableau, Power BI, or Matplotlib)
- Strong analytical and problem-solving skills
- Knowledge of machine learning concepts
- Git version control experience
- Excellent communication and teamwork skills
- Hebrew and English required"""
            st.rerun()

    # Validation feedback
    job_ready = bool(st.session_state.job_role.strip() and st.session_state.job_desc.strip())
    if st.session_state.job_desc.strip():
        is_valid, err = validate_job_description(st.session_state.job_desc)
        if not is_valid:
            st.error(f"{err}")
            job_ready = False
        else:
            st.markdown('<div class="alert-box alert-green">✓ Job description looks good.</div>', unsafe_allow_html=True)

    nav_buttons(0, can_proceed=job_ready, proceed_label="Continue")


# ============================================================
# TAB 1 — CV CONTENT
# ============================================================

elif current_tab == 1:
    st.markdown('<div class="rp-page-badge">Step 2 of 6</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-title">Add your CV</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-sub">Upload a PDF or Word file, or paste the text. English, Hebrew or Arabic is fine.</div>', unsafe_allow_html=True)

    # English-only notice

    method = st.radio(
        "How do you want to add it?",
        ["Upload File (PDF/DOCX/TXT)", "Paste Text", "Manual Entry"],
        format_func=lambda v: {"Upload File (PDF/DOCX/TXT)": "Upload a file", "Paste Text": "Paste the text",
                               "Manual Entry": "Type it in"}.get(v, v),
        horizontal=True,
        key="cv_method_radio"
    )
    st.session_state.input_method = method

    st.markdown("<div style='margin-bottom:24px;'></div>", unsafe_allow_html=True)

    # ---- UPLOAD ----
    if method == "Upload File (PDF/DOCX/TXT)":
        uploaded_file = st.file_uploader("Upload your CV", type=["pdf","docx","txt"], label_visibility="collapsed")
        if uploaded_file:
            with st.spinner("Reading your CV..."):
                if uploaded_file.type == "application/pdf":
                    text = extract_text_from_pdf(uploaded_file)
                elif "wordprocessingml" in uploaded_file.type:
                    text = extract_text_from_docx(uploaded_file)
                else:
                    text = str(uploaded_file.read(), "utf-8")

            if text.strip():
                # Detect if uploaded CV is in Hebrew or Arabic
                detected_lang = detect_language(text)
                if detected_lang in ["Hebrew", "Arabic"]:
                    st.info(f"Your CV is in **{detected_lang}** — that's fine. We'll read it as is. You can choose English, Hebrew or Arabic for the final CV in the Preview step.")

                st.session_state.cv_full_text = text
                st.session_state.manual_cv_data = classify_sections_refined(text)
                st.session_state.cv_uploaded = True
                st.success(f"CV uploaded and parsed! ({len(text.split())} words detected)")

                with st.expander("Preview detected sections"):
                    for k, v in st.session_state.manual_cv_data.items():
                        if v.strip():
                            st.markdown(f"**{k.replace('_',' ').title()}**")
                            st.text(v[:300] + ("..." if len(v) > 300 else ""))
            else:
                st.error("Could not extract text from this file. Try copy-pasting instead.")

    # ---- PASTE TEXT ----
    elif method == "Paste Text":
        pasted = st.text_area(
            "Paste your CV text here (in English):",
            value=st.session_state.cv_full_text,
            height=320,
            placeholder="Copy and paste your full CV content here in English..."
        )
        st.session_state.cv_full_text = pasted

        if pasted.strip():
            detected_lang = detect_language(pasted)
            if detected_lang in ["Hebrew", "Arabic"]:
                st.info(f"Your CV is in **{detected_lang}** — that's fine. You can choose English, Hebrew or Arabic for the final CV in the Preview step.")

        if st.button("Split into sections", type="primary"):
            if pasted.strip():
                st.session_state.manual_cv_data = classify_sections_refined(pasted)
                st.session_state.cv_uploaded = True
                st.success("Text structured into sections!")
                with st.expander("Preview detected sections"):
                    for k, v in st.session_state.manual_cv_data.items():
                        if v.strip():
                            st.markdown(f"**{k.replace('_',' ').title()}**")
                            st.text(v[:300] + ("..." if len(v) > 300 else ""))
            else:
                st.warning("Please paste your CV text first.")

        if st.session_state.cv_full_text.strip():
            st.session_state.cv_uploaded = True

    # ---- MANUAL ENTRY ----
    elif method == "Manual Entry":
        st.markdown("**Fill in each section in English below:**")
        fields = [
            ("personal_details", "👤 Full Name & Contact Info",
             "Your Full Name\nyour.email@example.com | +000-000-0000 | City, Country | linkedin.com/in/yourprofile"),
            ("summary",          "Professional Summary",
             "e.g., Final-year Computer Science student with experience in data analysis and software development. Passionate about turning data into actionable insights."),
            ("education",        "🎓 Education",
             "e.g., University Name\nBSc in Your Major\n2022–2025 | GPA: 3.8"),
            ("experience",       "Work Experience",
             "e.g., Job Title – Company Name\nMonth Year – Month Year\n- What you did and achieved\n- Another responsibility or result"),
            ("projects",         "Projects",
             "e.g., Project Name – Brief description\n- Tools used: Python, SQL, etc.\n- What problem it solved or result it produced"),
            ("skills",           "🛠️ Skills",
             "e.g., Python, SQL, Excel, Git, Data Analysis, Communication, Problem Solving, Teamwork"),
            ("courses_training", "📚 Courses & Certifications",
             "e.g., Google Data Analytics Certificate (Coursera, 2024)\nAdvanced Excel for Data Analysis (LinkedIn Learning)"),
            ("languages",        "Languages",
             "e.g., English (Native), Arabic (Fluent), French (Intermediate)"),
            ("volunteering",     "🤝 Volunteering & Activities",
             "e.g., Volunteer Tutor – Local NGO | 2023–Present\n- Helped 20+ students with math and science\nUniversity Student Council – Events Coordinator"),
        ]
        c1, c2 = st.columns(2)
        any_filled = False
        for i, (key, label, placeholder) in enumerate(fields):
            col = c1 if i % 2 == 0 else c2
            with col:
                st.markdown(f"**{label}**")
                val = st.text_area(
                    label, value=st.session_state.manual_cv_data.get(key,""),
                    placeholder=placeholder, key=f"manual_{key}",
                    label_visibility="collapsed", height=110
                )
                st.session_state.manual_cv_data[key] = val
                if val.strip(): any_filled = True

        if any_filled:
            st.session_state.cv_full_text = "\n".join(st.session_state.manual_cv_data.values())
            st.session_state.cv_uploaded = True

    if method == "Paste Text":
        pasted_text = st.session_state.cv_full_text.strip()
        if not pasted_text:
            st.error("Please paste your CV text before continuing.")
            cv_ready = False
        elif len(pasted_text) < 50:
            st.error(f"CV text is too short ({len(pasted_text)} characters). Please paste your full resume — minimum 50 characters.")
            cv_ready = False
        else:
            cv_ready = True
    else:
        cv_ready = bool(
            st.session_state.cv_full_text.strip() or
            any(v.strip() for v in st.session_state.manual_cv_data.values())
        )
    nav_buttons(1, can_proceed=cv_ready, proceed_label="Continue")


# ============================================================
# TAB 2 — ANALYSIS
# ============================================================

elif current_tab == 2:
    st.markdown('<div class="rp-page-badge">Step 3 of 6</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-title">How well you match</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-sub">What already fits the job, what is missing, and what a recruiter will look for.</div>', unsafe_allow_html=True)

    cv_text = st.session_state.cv_full_text or "\n".join(st.session_state.manual_cv_data.values())
    job_ok  = bool(st.session_state.job_desc.strip())
    cv_ok   = bool(cv_text.strip())

    if not job_ok or not cv_ok:
        st.warning("Please complete The Job and CV Content steps first.")
        nav_buttons(2, can_proceed=False)
    else:
        is_valid, err = validate_job_description(st.session_state.job_desc)
        if not is_valid:
            st.error(f"{err}")
            nav_buttons(2, can_proceed=False)
        else:
            # ── Parse CV into sections ──
            cv_sections = st.session_state.manual_cv_data.copy()
            if not any(v.strip() for v in cv_sections.values()):
                cv_sections = classify_sections_refined(cv_text)

            # ── Run deep Claude analysis (Step 1+2) ──
            if not st.session_state.ai_cv_analysis:
                with st.spinner("🤖 Analyzing your CV against the job requirements... (~15 seconds)"):
                    result = call_claude_deep_analysis(
                        cv_sections,
                        st.session_state.job_role,
                        st.session_state.job_desc,
                    )
                if result:
                    st.session_state.ai_cv_analysis = result
                else:
                    # Fallback: run rule-based and build a compatible structure
                    rule = run_rule_based_analysis(cv_text, st.session_state.job_desc)
                    st.session_state.ai_cv_analysis = {
                        "cv_strengths":           [g["label"] for g in rule["strong"]],
                        "cv_weaknesses":          [g["label"] for g in rule["missing"]],
                        "hard_skills_required":   [g["label"] for g in rule["missing"]],
                        "soft_skills_required":   [],
                        "domain_knowledge_required": [],
                        "ats_keywords_missing":   [g["label"] for g in rule["missing"]],
                        "ats_keywords_present":   [g["label"] for g in rule["strong"]],
                        "match_score":            rule["score"],
                        "match_label":            ("Strong Match" if rule["score"] >= 70
                                                   else "Moderate Match" if rule["score"] >= 45
                                                   else "Weak Match"),
                        "score_rationale":        "Based on rule-based keyword matching (AI unavailable).",
                        "metrics_missing":        [],
                        "quick_wins":             [f"Add {g['label']} to your CV" for g in rule["missing"][:3]],
                    }
                    if get_anthropic_api_key() is None:
                        st.error("**ANTHROPIC_API_KEY not found.** Add it to `.streamlit/secrets.toml` to enable AI analysis. Showing rule-based results.")
                    else:
                        st.warning("AI analysis failed (API error). Showing rule-based results.")

            # Persist to analysis_results for downstream tabs
            ai = st.session_state.ai_cv_analysis
            st.session_state.analysis_results = {
                "matches": ai.get("ats_keywords_present", []),
                "missing": ai.get("ats_keywords_missing", []),
                "score":   ai.get("match_score", 0),
            }
            st.session_state.analysis_done = True

            score = ai.get("match_score", 0)
            st.markdown(render_match_report(ai, score), unsafe_allow_html=True)
            nav_buttons(2, can_proceed=True, proceed_label="Continue to questions")


# ============================================================
# TAB 3 — REFINEMENT
# ============================================================

elif current_tab == 3:
    st.markdown('<div class="rp-page-badge">Step 4 of 6</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-title">A few quick questions</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-sub">Your answers add the real numbers and results your CV is missing. Skip any you don\'t know.</div>', unsafe_allow_html=True)

    if not st.session_state.analysis_done:
        st.info("ℹ️ Please complete the Analysis step first.")
        nav_buttons(3, can_proceed=False)
    else:
        score = st.session_state.analysis_results.get("score", 0)
        tier  = get_match_tier(score)

        if tier == "high":
            st.success(f"Your CV scores **{score}%** — strong match. Answer below to make it even sharper.")
        elif tier == "medium":
            st.warning(f"Your CV scores **{score}%** — partial match. Answer honestly below. Only real experience will be used in the rewrite.")
        else:
            # Low match — skip the refinement questions entirely
            st.markdown(f"""
            <div style="background:#FBE9EC;border:1px solid #F5C6CE;border-radius:12px;padding:1.4rem 1.6rem;margin-bottom:1.2rem;">
                <div style="font-size:1.05rem;font-weight:700;color:#B4233A;margin-bottom:8px;">Low Match — {score}%</div>
                <div style="font-size:0.92rem;color:#B4233A;line-height:1.65;">
                    Your CV doesn't overlap enough with this role for honest tailoring.<br>
                    Rewriting it to "fit" would mean exaggerating or fabricating experience — which ResumePilot won't do.<br><br>
                    <strong>Instead, the next step will show you a career development roadmap:</strong> what skills are missing,
                    which projects or courses would close the gap, and which roles you're better positioned for right now.
                </div>
            </div>
            """, unsafe_allow_html=True)

            ats_missing = st.session_state.ai_cv_analysis.get("ats_keywords_missing", [])
            better_roles = st.session_state.ai_cv_analysis.get("quick_wins", [])
            if ats_missing:
                chips = " ".join(
                    f'<span style="display:inline-block;background:#FBE9EC;color:#B4233A;padding:3px 10px;'
                    f'border-radius:9999px;font-size:0.8rem;margin:2px;">{_esc(k)}</span>'
                    for k in ats_missing[:10]
                )
                st.markdown(f'<div style="margin-bottom:12px;"><strong style="font-size:0.85rem;color:#5A6478;">Skills gap at a glance:</strong><br>{chips}</div>', unsafe_allow_html=True)

            nav_buttons(3, can_proceed=True, proceed_label="See your plan")
            st.stop()

        st.markdown("<div style='margin-bottom:24px;'></div>", unsafe_allow_html=True)

        # ── Generate dynamic personalized questions (Step 3) ──
        if not st.session_state.dynamic_questions:
            cv_sections = st.session_state.manual_cv_data.copy()
            if not any(v.strip() for v in cv_sections.values()):
                cv_sections = classify_sections_refined(
                    st.session_state.cv_full_text or ""
                )
            with st.spinner("🤖 Generating personalized questions for your CV... (~10 seconds)"):
                questions = call_claude_generate_questions(
                    cv_sections,
                    st.session_state.job_role,
                    st.session_state.job_desc,
                    st.session_state.ai_cv_analysis,
                )
            if questions:
                st.session_state.dynamic_questions = questions
            else:
                # Fallback: generic gap-based questions
                missing = st.session_state.analysis_results.get("missing", [])
                fallback_qs = []
                for i, gap in enumerate(missing[:5]):
                    gap_label = gap if isinstance(gap, str) else str(gap)
                    fallback_qs.append({
                        "key": f"q_gap_{i}",
                        "question": f"The role requires **{gap_label}**. Do you have any experience, coursework, or projects involving this? Describe briefly.",
                        "hint": f"e.g., I used {gap_label} in a university project / I studied it in [course] / I don't have this yet",
                        "section": "skills",
                        "gap_addressed": f"Clarifies exposure to {gap_label}",
                    })
                if not fallback_qs:
                    fallback_qs = [
                        {"key": "q_achievement", "question": "What is your biggest achievement relevant to this role? Include numbers or outcomes if possible.", "hint": "e.g., Built a model that improved accuracy by 15% / Led a team of 5 students for 6 months", "section": "experience", "gap_addressed": "Adds quantifiable achievement to rewrite"},
                        {"key": "q_tools", "question": "Are there tools, platforms, or technologies you use regularly that aren't listed on your CV?", "hint": "e.g., I use Tableau weekly but forgot to add it / I've been learning AWS on the side", "section": "skills", "gap_addressed": "Surfaces unlisted relevant skills"},
                    ]
                st.session_state.dynamic_questions = fallback_qs
                if get_anthropic_api_key() is None:
                    st.error("**ANTHROPIC_API_KEY not found.** Showing fallback questions. Add your key to enable personalized AI questions.")
                else:
                    err_detail = st.session_state.get('_last_api_error', 'unknown error')
                    st.warning(f"Could not generate personalized questions ({err_detail}). Showing standard questions.")

        # ── Render questions (Step 4: store all answers in structured memory) ──
        questions = st.session_state.dynamic_questions
        st.markdown(f'<p style="font-size:17px;color:#5A6478;margin:0 0 20px">{len(questions)} questions. Your answers are saved as you type.</p>', unsafe_allow_html=True)

        for q in questions:
            key    = q.get("key", "q_unknown")
            text   = q.get("question", "")
            hint   = q.get("hint", "")
            section= q.get("section", "")
            gap    = q.get("gap_addressed", "")

            # Show section badge if available
            badge = ""
            text_html = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', _esc(text))
            st.markdown(f'<div style="font:700 19px/1.35 var(--display);color:#16213A;margin:8px 0 6px;">{text_html}{badge}</div>', unsafe_allow_html=True)
            if gap:
                st.markdown(f'<div class="rp-why" style="margin-bottom:10px"><b>Why we ask</b>{_esc(gap)}</div>', unsafe_allow_html=True)

            # Retrieve prior answer (supports both old string format and new dict format)
            prior = st.session_state.follow_up_answers.get(key, "")
            if isinstance(prior, dict):
                prior = prior.get("answer", "")

            answer = st.text_area(
                text,
                value=prior,
                placeholder=hint,
                key=f"refine_{key}",
                label_visibility="collapsed",
                height=90,
            )

            # Store as structured memory (Step 4)
            st.session_state.follow_up_answers[key] = {
                "question":      text,
                "answer":        answer,
                "section":       section,
                "gap_addressed": gap,
            }

            st.markdown("")

        st.markdown("<div style='margin-bottom:24px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="rp-tip">Leave a question blank if it doesn\'t apply. Your new CV only uses what you write here and what is already in your CV.</div>', unsafe_allow_html=True)

        st.markdown("")
        nav_buttons(3, can_proceed=True, proceed_label="Write my new CV")


# ============================================================
# TAB 4 — PREVIEW
# ============================================================

elif current_tab == 4:
    st.markdown('<div class="rp-page-badge">Step 5 of 6</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-title">Review your new CV</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-sub">Every change is marked. Keep it, undo it, or edit it before you download.</div>', unsafe_allow_html=True)

    lang = st.radio("Output language:", ["English","Hebrew","Arabic"], horizontal=True, key="lang_select")
    if lang == "Hebrew":
        form_labels = {"neutral": "Neutral wording", "feminine": "Feminine (פיתחה)", "masculine": "Masculine (פיתח)"}
        st.radio(
            "How should your Hebrew CV refer to you?",
            list(form_labels.keys()),
            format_func=lambda k: form_labels[k],
            horizontal=True,
            key="hebrew_form",
        )
    if lang != st.session_state.active_lang:
        st.session_state.active_lang = lang
        st.markdown(
            "<script>window.parent.document.querySelector('.main').scrollTo(0,0);</script>",
            unsafe_allow_html=True,
        )
        st.rerun()
    st.session_state.active_lang = lang

    # ── Get base CV sections ──
    base_data = st.session_state.manual_cv_data.copy()
    if not any(v.strip() for v in base_data.values()):
        base_data = classify_sections_refined(st.session_state.cv_full_text or "")

    has_content = any(v.strip() for v in base_data.values())

    if not has_content:
        st.warning("No CV content found. Please go back to **CV Content** and add your information.")
        nav_buttons(4, can_proceed=False)
    else:
        score = st.session_state.analysis_results.get("score", 0)
        tier  = get_match_tier(score)

        # ══════════════════════════════════════════
        # LOW MATCH — Career Roadmap, no CV rewrite
        # ══════════════════════════════════════════
        if tier == "low":
            st.markdown(f"""
            <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:12px;padding:1.2rem 1.6rem;margin-bottom:1.4rem;">
                <div style="font-size:1rem;font-weight:700;color:#c2410c;margin-bottom:6px;">Match Score: {score}% — Career Gap Analysis Mode</div>
                <div style="font-size:0.9rem;color:#7c2d12;line-height:1.6;">
                    This CV doesn't match the role well enough for tailoring without exaggerating your background.
                    Below is an honest breakdown of what's missing and a concrete roadmap to close the gap.
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Generate career recommendations (once per session)
            if not st.session_state.career_recommendations:
                with st.spinner("Analyzing gaps and building your career roadmap... (~15 seconds)"):
                    recs = call_claude_career_recommendations(
                        base_data,
                        st.session_state.job_role,
                        st.session_state.job_desc,
                        st.session_state.ai_cv_analysis,
                    )
                if recs:
                    st.session_state.career_recommendations = recs
                else:
                    st.session_state.career_recommendations = {"_fallback": True}
                    st.warning("Could not generate AI recommendations. Showing gap analysis from the Analysis tab.")

            recs = st.session_state.career_recommendations

            if recs and not recs.get("_fallback"):
                # ── Why it's a low match ──
                if recs.get("why_low_match"):
                    st.markdown(f"""
                    <div style="background:#fef2f2;border-left:4px solid #ef4444;border-radius:0 10px 10px 0;
                        padding:12px 16px;margin-bottom:1rem;font-size:0.9rem;color:#B4233A;line-height:1.6;">
                        <strong>Why the low match:</strong> {recs['why_low_match']}
                    </div>""", unsafe_allow_html=True)

                # ── What is strong ──
                if recs.get("what_is_strong"):
                    st.markdown(f"""
                    <div style="background:#EAF6F1;border-left:4px solid #22c55e;border-radius:0 10px 10px 0;
                        padding:12px 16px;margin-bottom:1rem;font-size:0.9rem;color:#0B7A5C;line-height:1.6;">
                        <strong>What you do have:</strong> {recs['what_is_strong']}
                    </div>""", unsafe_allow_html=True)

                col_crit, col_nice = st.columns(2)

                # ── Missing skills ──
                with col_crit:
                    critical = recs.get("critical_missing_skills", [])
                    if critical:
                        st.markdown("**🔴 Critical missing skills**")
                        for s in critical:
                            st.markdown(f'<div style="background:#fee2e2;border-radius:6px;padding:6px 12px;'
                                        f'margin-bottom:4px;font-size:0.87rem;color:#991b1b;">✗ {_esc(s)}</div>',
                                        unsafe_allow_html=True)

                with col_nice:
                    nice = recs.get("nice_to_have_missing", [])
                    if nice:
                        st.markdown("**🟡 Nice-to-have gaps**")
                        for s in nice:
                            st.markdown(f'<div style="background:#fef9c3;border-radius:6px;padding:6px 12px;'
                                        f'margin-bottom:4px;font-size:0.87rem;color:#713f12;">◎ {_esc(s)}</div>',
                                        unsafe_allow_html=True)

                st.markdown("<div style='margin-bottom:1.2rem;'></div>", unsafe_allow_html=True)

                # ── Recommended projects ──
                projects = recs.get("recommended_projects", [])
                if projects:
                    st.markdown("**🛠️ Projects to build**")
                    for p in projects:
                        skills_str = " · ".join(p.get("skills_gained", []))
                        st.markdown(f"""
                        <div style="background:#EEF1FB;border:1px solid #D5DCF5;border-radius:10px;
                            padding:12px 16px;margin-bottom:8px;">
                            <div style="font-weight:700;font-size:0.92rem;color:#16213A;">{p.get('title','')}</div>
                            <div style="font-size:0.85rem;color:#2A3FD6;margin:4px 0;">{p.get('description','')}</div>
                            <div style="font-size:0.78rem;color:#5A6478;">
                                Skills: {skills_str} &nbsp;·&nbsp; {p.get('difficulty','')} &nbsp;·&nbsp; {p.get('estimated_time','')}
                            </div>
                        </div>""", unsafe_allow_html=True)

                # ── Recommended courses ──
                courses = recs.get("recommended_courses", [])
                if courses:
                    st.markdown("**📚 Courses to take**")
                    for c in courses:
                        free_badge = '<span style="background:#EAF6F1;color:#0B7A5C;padding:1px 7px;border-radius:9999px;font-size:0.75rem;font-weight:600;margin-left:6px;">FREE</span>' if c.get("free") else ""
                        skills_str = " · ".join(c.get("skills_gained", []))
                        st.markdown(f"""
                        <div style="background:#F6F8FB;border:1px solid #DCE2EC;border-radius:8px;
                            padding:10px 14px;margin-bottom:6px;font-size:0.87rem;">
                            <strong>{c.get('name','')}</strong>{free_badge}
                            <span style="color:#5A6478;margin-left:8px;">{c.get('platform','')}</span>
                            <div style="color:#5A6478;font-size:0.8rem;margin-top:2px;">Covers: {skills_str}</div>
                        </div>""", unsafe_allow_html=True)

                # ── Certifications ──
                certs = recs.get("recommended_certifications", [])
                if certs:
                    st.markdown("**🏅 Certifications worth getting**")
                    for cert in certs:
                        st.markdown(f"""
                        <div style="background:#F6F8FB;border:1px solid #DCE2EC;border-radius:8px;
                            padding:10px 14px;margin-bottom:6px;font-size:0.87rem;">
                            <strong>{cert.get('name','')}</strong>
                            <div style="color:#5A6478;font-size:0.8rem;margin-top:2px;">{cert.get('why_relevant','')}</div>
                        </div>""", unsafe_allow_html=True)

                # ── Timeline + better fit roles ──
                col_t, col_r = st.columns(2)
                with col_t:
                    if recs.get("realistic_timeline"):
                        st.markdown(f"""
                        <div style="background:#F6F8FB;border-radius:10px;padding:12px 16px;font-size:0.88rem;">
                            <strong>⏱️ Realistic timeline</strong><br>
                            <span style="color:#5A6478;">{recs['realistic_timeline']}</span>
                        </div>""", unsafe_allow_html=True)
                with col_r:
                    better_roles = recs.get("better_fit_roles", [])
                    if better_roles:
                        roles_html = "".join(
                            f'<div style="background:#eff6ff;color:#1d4ed8;padding:4px 12px;'
                            f'border-radius:9999px;font-size:0.82rem;margin-bottom:4px;display:inline-block;margin-right:4px;">{_esc(r)}</div>'
                            for r in better_roles
                        )
                        st.markdown(f"""
                        <div style="background:#F6F8FB;border-radius:10px;padding:12px 16px;font-size:0.88rem;">
                            <strong>Better-fit roles right now</strong><br><br>{roles_html}
                        </div>""", unsafe_allow_html=True)

            st.markdown("<div style='margin-bottom:1.4rem;'></div>", unsafe_allow_html=True)

            # ── Re-run button ──
            col_regen, _ = st.columns([2, 3])
            with col_regen:
                if st.button("Write the plan again"):
                    st.session_state.career_recommendations = {}
                    st.rerun()

            # ── Still show the original CV (formatted, not rewritten) ──
            st.markdown("---")
            st.markdown("**Your original CV (formatted, not rewritten):**")
            st.caption("We didn't change the content — only applied formatting so you can still download it.")
            st.session_state.final_cv_data = base_data   # original, no rewrite

            if lang in ["Hebrew", "Arabic"]:
                content_id = hashlib.md5(json.dumps(base_data, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:10]
                cache_key = f"translated_{lang}_{st.session_state.get('hebrew_form', 'neutral') if lang == 'Hebrew' else ''}_{content_id}"
                cached = st.session_state.get(cache_key, {})
                if not cached:
                    with st.spinner(f"Formatting CV in {'Hebrew' if lang == 'Hebrew' else 'Arabic'}..."):
                        localized = call_claude_localize_cv(base_data, lang, st.session_state.get("hebrew_form", "neutral"))
                    if localized:
                        st.session_state[cache_key] = localized
                        display_cv = localized
                    else:
                        st.warning(f"Couldn't translate to {lang} right now. Showing the English version — try again in a minute.")
                        display_cv = base_data
                else:
                    display_cv = cached
            else:
                display_cv = base_data

            st.session_state.final_cv_data = display_cv

            st.markdown('<div style="background:#F6F8FB;padding:30px 15px;border-radius:12px;margin-top:12px;">', unsafe_allow_html=True)
            st.markdown(render_cv_html(display_cv, lang), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            nav_buttons(4, can_proceed=True, proceed_label="Continue to download")

        # ══════════════════════════════════════════════════════════
        # MEDIUM or HIGH MATCH — CV Rewrite (honest for medium)
        # ══════════════════════════════════════════════════════════
        else:
            # ── Steps 5-10: Full CV Rewrite (English, run once per session) ──
            if not st.session_state.cv_adjusted:
                spinner_msg = (
                    "Rewriting your CV for this role... (~20 seconds)"
                    if tier == "high"
                    else "Tailoring the matching parts of your CV honestly... (~20 seconds)"
                )
                with st.spinner(spinner_msg):
                    rewrite_result = call_claude_rewrite_cv(
                        base_data,
                        st.session_state.job_role,
                        st.session_state.job_desc,
                        st.session_state.ai_cv_analysis,
                        st.session_state.follow_up_answers,
                        match_tier=tier,
                    )

                if rewrite_result:
                    improvement_log = rewrite_result.pop("improvement_log", [])
                    st.session_state.improvement_log = improvement_log if isinstance(improvement_log, list) else []
                    safe_result, fact_notes = fact_check_rewrite(
                        base_data, rewrite_result, answers_to_text(st.session_state.follow_up_answers)
                    )
                    st.session_state.fact_check_notes = fact_notes
                    st.session_state.ai_sections = {
                        k: v.strip() for k, v in safe_result.items()
                        if isinstance(v, str) and v.strip() and k != "personal_details"
                    }
                    reset_review_choices()
                    st.session_state.adjusted_cv_data = build_final_cv(base_data, safe_result, st.session_state.job_role)
                    st.session_state.cv_adjusted = True
                else:
                    # Show the original CV unchanged; never add text the user didn't write.
                    st.session_state.adjusted_cv_data = base_data.copy()
                    st.session_state.fact_check_notes = []
                    st.session_state.ai_sections = {}
                    reset_review_choices()
                    st.session_state.cv_adjusted = True
                    st.session_state.improvement_log = []
                    if get_anthropic_api_key() is None:
                        st.error(
                            "**ANTHROPIC_API_KEY not found.** "
                            "Add it to `.streamlit/secrets.toml` as `ANTHROPIC_API_KEY = \"sk-ant-...\"` "
                            "to enable AI rewriting. Showing your original CV for now."
                        )
                    else:
                        st.warning("AI rewrite failed (API error). Showing your original CV.")

            for note in st.session_state.get("fact_check_notes", []):
                st.info(note)

            english_cv = render_change_review(base_data, st.session_state.get("ai_sections", {}),
                                              st.session_state.improvement_log)
            st.session_state.adjusted_cv_data = english_cv

            # ── Medium match: show missing skills warning ──
            if tier == "medium":
                missing_skills = st.session_state.ai_cv_analysis.get("ats_keywords_missing", [])
                if missing_skills:
                    chips = " ".join(
                        f'<span style="display:inline-block;background:#FBE9EC;color:#B4233A;padding:3px 10px;'
                        f'border-radius:9999px;font-size:0.78rem;margin:2px;">{_esc(k)}</span>'
                        for k in missing_skills[:10]
                    )
                    st.markdown(f"""
                    <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;
                        padding:12px 16px;margin-bottom:1rem;">
                        <div style="font-weight:700;font-size:0.9rem;color:#c2410c;margin-bottom:6px;">
                            Partial match ({score}%) — only real overlaps were tailored
                        </div>
                        <div style="font-size:0.82rem;color:#78350f;margin-bottom:8px;">
                            These skills are required by the job but absent from your CV.
                            They were <strong>not added</strong> to avoid dishonesty — consider building them before applying:
                        </div>
                        {chips}
                    </div>""", unsafe_allow_html=True)

            # ── Step 9: Localization for Hebrew / Arabic ──
            if lang in ["Hebrew", "Arabic"]:
                content_id = hashlib.md5(json.dumps(english_cv, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:10]
                cache_key = f"translated_{lang}_{st.session_state.get('hebrew_form', 'neutral') if lang == 'Hebrew' else ''}_{content_id}"
                cached = st.session_state.get(cache_key, {})
                if not cached:
                    with st.spinner(f"Localizing CV to {'Hebrew' if lang == 'Hebrew' else 'Arabic'}... (~15 seconds)"):
                        localized = call_claude_localize_cv(english_cv, lang, st.session_state.get("hebrew_form", "neutral"))
                    if localized:
                        st.session_state[cache_key] = localized
                        final_cv = localized
                    else:
                        st.warning(f"Couldn't translate to {lang} right now. Showing the English version — try again in a minute.")
                        final_cv = english_cv
                else:
                    final_cv = cached
            else:
                final_cv = english_cv

            st.session_state.final_cv_data = final_cv

            # ── Regenerate button ──
            col_regen, _ = st.columns([2, 3])
            with col_regen:
                if st.button("Rewrite my CV again"):
                    st.session_state.cv_adjusted = False
                    st.session_state.adjusted_cv_data = {}
                    st.session_state.final_cv_data = {}
                    st.session_state.improvement_log = []
                    st.session_state.ai_sections = {}
                    reset_review_choices()
                    for k in [k for k in st.session_state.keys() if str(k).startswith("translated_")]:
                        del st.session_state[k]
                    st.rerun()

            st.markdown("<div style='margin-bottom:24px;'></div>", unsafe_allow_html=True)

            # ── CV Preview ──
            st.markdown('<div style="background:#F6F8FB;padding:30px 15px;border-radius:12px;">', unsafe_allow_html=True)
            st.markdown(render_cv_html(final_cv, lang), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            ready_msg = "**Your tailored CV is ready!**" if tier == "high" else "**Your CV is ready** — only genuine overlaps with the job were strengthened."
            st.markdown(f'<div style="margin-top:1rem;padding:1rem;background:#EAF6F1;border-radius:10px;border:1px solid #BFE3D5;">{ready_msg} Download it in the next tab.</div>', unsafe_allow_html=True)

            nav_buttons(4, can_proceed=True, proceed_label="Continue to download")


# ============================================================
# TAB 5 — EXPORT
# ============================================================

elif current_tab == 5:
    st.markdown('<div class="rp-page-badge">Step 6 of 6</div>', unsafe_allow_html=True)
    st.markdown('<div class="rp-page-title">Download your CV</div>', unsafe_allow_html=True)
    _export_score = st.session_state.analysis_results.get("score", 0)
    _export_sub = (
        "Your tailored CV is ready to send."
        if get_match_tier(_export_score) == "high"
        else "Your partially tailored CV is ready — review carefully before sending."
        if get_match_tier(_export_score) == "medium"
        else "Your original CV with clean formatting — read the notes below before applying."
    )
    st.markdown(f'<div class="rp-page-sub">{_export_sub}</div>', unsafe_allow_html=True)

    cv_f = st.session_state.final_cv_data
    has_content = bool(cv_f and any(v.strip() for v in cv_f.values()))

    if not has_content:
        st.warning("Please complete the Preview step first.")
        nav_buttons(5, can_proceed=False)
    else:
        active_lang = st.session_state.active_lang
        score = st.session_state.analysis_results.get("score", 0)
        tier  = get_match_tier(score)

        # Score badge + tier-specific message
        if score:
            score_class = "alert-green" if tier == "high" else "alert-orange" if tier == "medium" else "alert-red"
            st.markdown(f'<div class="alert-box {score_class}">Match Score: {score}% for {_esc(st.session_state.job_role)}</div>', unsafe_allow_html=True)

        if tier == "low":
            st.markdown("""
            <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;
                padding:14px 18px;margin-bottom:16px;font-size:0.9rem;color:#7c2d12;line-height:1.65;">
                <strong>Heads up:</strong> Your CV didn't match this role well enough for tailoring.
                What you're downloading is your <strong>original CV with clean formatting applied</strong> — the content was not rewritten to fit this job.<br><br>
                Go back to the <strong>Review your new CV</strong> tab to see your career roadmap and the skills you need to build before applying to this role.
            </div>
            """, unsafe_allow_html=True)
        elif tier == "medium":
            missing = st.session_state.ai_cv_analysis.get("ats_keywords_missing", [])
            missing_str = ", ".join(missing[:5]) if missing else "some required skills"
            st.markdown(f"""
            <div style="background:#fffbeb;border:1px solid #fde68a;border-radius:10px;
                padding:14px 18px;margin-bottom:16px;font-size:0.9rem;color:#78350f;line-height:1.65;">
                ℹ️ <strong>Partial match:</strong> Only the sections where your background genuinely overlaps with this role were strengthened.
                Skills you don't have ({missing_str}) were <strong>not added</strong> — don't claim them in interviews.
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="background:#EAF6F1;border:1px solid #BFE3D5;border-radius:10px;
                padding:14px 18px;margin-bottom:16px;font-size:0.9rem;color:#0B7A5C;line-height:1.65;">
                <strong>Strong match:</strong> Your CV has been fully tailored for this role. Review every line before sending.
            </div>
            """, unsafe_allow_html=True)

        # File name
        pd_raw = cv_f.get("personal_details","Candidate")
        name_part, _ = format_contact_info(pd_raw)
        first_name = name_part.split()[0] if name_part else "Candidate"
        job_slug = re.sub(r'[^a-zA-Z0-9]', '_', st.session_state.job_role) if st.session_state.job_role else "CV"
        f_name = f"{first_name}_{job_slug}_CV"

        st.markdown("<div style='margin-bottom:24px;'></div>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)

        with c1:
            st.markdown('<div class="rp-dl-fmt">DOCX</div><div class="rp-dl-desc">Recommended — send directly to employers</div>', unsafe_allow_html=True)
            docx_buf = generate_docx(cv_f, active_lang)
            st.download_button(
                "Download DOCX",
                docx_buf,
                file_name=f"{f_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )

        with c2:
            st.markdown('<div class="rp-dl-fmt">TXT</div><div class="rp-dl-desc">For online forms and portals</div>', unsafe_allow_html=True)
            h_map = HEADINGS_MAP[active_lang]
            section_order = [
                ("personal_details", h_map[0]),
                ("summary",          h_map[1]),
                ("education",        h_map[2]),
                ("experience",       h_map[3]),
                ("projects",         h_map[4]),
                ("courses_training", h_map[5]),
                ("volunteering",     h_map[6]),
                ("languages",        h_map[7]),
                ("skills",           h_map[8]),
            ]
            txt_out = f"{st.session_state.job_role} — CV\n{'='*50}\n\n"
            for key, title in section_order:
                v = cv_f.get(key,"").strip()
                if v:
                    if key == "skills":
                        v = re.sub(r'(Technical Skills|Business & Professional Skills|Professional Skills)\s*[,\n]?', '', v, flags=re.IGNORECASE)
                    v = re.sub(r'\s*•\s*', '\n  • ', v)
                    txt_out += f"{title.upper()}\n{'-'*len(title)}\n{v.strip()}\n\n"
            st.download_button(
                "Download TXT",
                txt_out,
                file_name=f"{f_name}.txt",
                mime="text/plain",
                use_container_width=True,
                type="secondary"
            )

        st.markdown("""
        <div class="rp-checklist" style="margin-top:20px;">
            <div class="rp-checklist-title">Before you send</div>
            <div class="rp-check-item">Review every line — make sure everything is accurate.</div>
            <div class="rp-check-item">Never include anything you cannot back up in an interview.</div>
            <div class="rp-check-item">Customize the summary for each job you apply to.</div>
            <div class="rp-check-item">Good luck. You've got this. 🍀</div>
        </div>
        """, unsafe_allow_html=True)
        nav_buttons(5, can_proceed=False)

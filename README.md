# Canvas Essay Grader

Automatically grades student essays submitted via Canvas LMS. Essays are fetched from Google Docs, graded by Claude AI against an IELTS-style rubric, and scores are submitted back to Canvas. Feedback can also be posted as comments directly onto students' Google Docs.

---

## Features

- Fetches all submissions from a Canvas assignment
- Reads student essays from Google Docs (shared links)
- Grades with Claude AI across 4 IELTS criteria (Task Response, Coherence & Cohesion, Lexical Resource, Grammar Range & Accuracy)
- Submits rubric scores + feedback to Canvas
- Posts per-criterion comments directly onto students' Google Docs
- Live progress stream — grades students one by one as results arrive
- Per-student or bulk submit/comment actions

---

## Requirements

- Python 3.10+
- A Canvas LMS account with API access
- An Anthropic API key
- A Google Cloud service account with Drive API enabled

---

## Setup

### 1. Clone and create a virtual environment

```bash
cd canvas_grader
python3 -m venv venv
source venv/bin/activate        # macOS/Linux
# venv\Scripts\activate         # Windows
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure environment variables

Create a `.env` file in the project root:

```env
CANVAS_API_TOKEN=your_canvas_token_here
ANTHROPIC_API_KEY=your_anthropic_key_here
GOOGLE_SERVICE_ACCOUNT_JSON=/absolute/path/to/service-account.json
```

| Variable | Where to get it |
|---|---|
| `CANVAS_API_TOKEN` | Canvas → Account → Settings → Approved Integrations → New Access Token |
| `ANTHROPIC_API_KEY` | https://console.anthropic.com → API Keys |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | See Google Cloud setup below |

### 4. Google Cloud setup (for Google Docs commenting)

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a project → **APIs & Services → Enable APIs** → enable **Google Drive API**
3. **APIs & Services → Credentials → Create Credentials → Service Account**
4. Open the service account → **Keys → Add Key → Create new key → JSON**
5. Save the downloaded `.json` file and set its path in `.env` as `GOOGLE_SERVICE_ACCOUNT_JSON`

> **Note:** Students' Google Docs must be shared as **"Anyone with the link can edit"** — no further sharing with the service account is needed.

---

## Running the app

```bash
source venv/bin/activate        # if not already active
python app.py
```

Then open your browser at:

```
http://localhost:5000
```

---

## Usage

1. **Fill in the configuration panel**
   - Enter your Canvas **Course ID** and **Assignment ID** (found in the Canvas URL)
   - Set the **Total Points** (default: 36)
   - Optionally edit the grading rubric

2. **Click "Fetch & Grade All Submissions"**
   - The app fetches each submission, reads the Google Doc, and grades it with Claude
   - Results stream in live — scores, per-criterion breakdown, and feedback appear as each essay is graded

3. **Review and edit**
   - Adjust scores or feedback directly in the table before submitting

4. **Submit to Canvas**
   - Click **"Submit Grades to Canvas"** to post all rubric scores at once
   - Or use the **"↑ Submit"** button on any row to submit a single student

5. **Post Google Doc comments** *(optional)*
   - Click **"💬 Post Google Doc Comments"** to post AI feedback as margin comments on all docs
   - Or use the **"💬 Comment"** button per row for a single student

6. **Export**
   - Click **"↓ Export CSV"** to download all results

---

## Finding Canvas Course ID and Assignment ID

The IDs are in the Canvas URL when you open an assignment:

```
https://canvas.instructure.com/courses/COURSE_ID/assignments/ASSIGNMENT_ID
```

---

## Project structure

```
canvas_grader/
├── app.py                          # Flask app — all backend logic
├── templates/
│   └── index.html                  # Frontend UI
├── requirements.txt
├── .env                            # Secrets (never commit this)
├── .gitignore
└── README.md
```

---

## Environment file reference (`.env`)

```env
# Canvas LMS
CANVAS_API_TOKEN=7~xxxxxxxxxxxx

# Anthropic Claude
ANTHROPIC_API_KEY=sk-ant-api03-xxxxxxxxxxxx

# Google Drive (path to service account JSON key file)
GOOGLE_SERVICE_ACCOUNT_JSON=/Users/yourname/keys/service-account.json
```

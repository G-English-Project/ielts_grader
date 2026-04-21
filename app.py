from flask import Flask, render_template, jsonify, request, Response, send_from_directory
import requests
import re
import json
import os
import uuid
import base64
import anthropic
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

load_dotenv(override=True)

app = Flask(__name__)

# ─────────────────────────────────────────────
#  CREDENTIALS  (loaded from .env)
# ─────────────────────────────────────────────
CANVAS_API_TOKEN              = os.getenv("CANVAS_API_TOKEN")
CANVAS_DOMAIN                 = "canvas.instructure.com"
ANTHROPIC_API_KEY             = os.getenv("ANTHROPIC_API_KEY")
GOOGLE_SERVICE_ACCOUNT_JSON   = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

# ─────────────────────────────────────────────
#  RUBRIC  (edit this to change grading criteria)
# ─────────────────────────────────────────────
RUBRIC = """
Task response
Band 9 - 9pts
The prompt is appropriately addressed and explored in depth.
A clear and fully developed position is presented which directly answers the question/s.
Ideas are relevant, fully extended and well supported.
Any lapses in content or support are extremely rare.
Band 8  - 8pts
The prompt is appropriately and sufficiently addressed.
A clear and well-developed position is presented in response to the question/s.
Ideas are relevant, well extended and supported.
There may be occasional omissions or lapses in content.
Band 7  - 7pts
The main parts of the prompt are appropriately addressed.
A clear and developed position is presented.
Main ideas are extended and supported but there may be a tendency to over-generalise or there may be a lack of focus and precision in supporting ideas/ material.
Band 6  - 6pts
The main parts of the prompt are addressed (though some may be more fully covered than others). An appropriate format is used.
A position is presented that is directly relevant to the prompt, although the conclusions drawn may be unclear, unjustified or repetitive.
Main ideas are relevant, but some may be insufficiently developed or may lack clarity, while some supporting arguments and evidence may be less relevant or inadequate.
Band 5  - 5pts
The main parts of the prompt are incompletely addressed. The format may be inappropriate in places.
The writer expresses a position, but the development is not always clear.
Some main ideas are put forward, but they are limited and are not sufficiently developed and/or there may be irrelevant detail.
There may be some repetition.
Band 4  - 4pts
The prompt is tackled in a minimal way, or the answer is tangential, possibly due to some misunderstanding of the prompt. The format may be inappropriate.
A position is discernible, but the reader has to read carefully to find it.
Main ideas are difficult to identify and such ideas that are identifiable may lack relevance, clarity and/or support.
Large parts of the response may be repetitive.


Coherence and Cohesion
Band 9
The message can be followed effortlessly.
Cohesion is used in such a way that it very rarely attracts attention.
Any lapses in coherence or cohesion are minimal.
Paragraphing is skilfully managed.
Band 8
The message can be followed with ease.
Information and ideas are logically sequenced, and cohesion is well managed.
Occasional lapses in coherence and cohesion may occur.
Paragraphing is used sufficiently and appropriately.
Band 7
Information and ideas are logically organised, and there is a clear progression throughout the response. (A few lapses may occur, but these are minor.)
A range of cohesive devices including reference and substitution is used flexibly but with some inaccuracies or some over/under use.
Paragraphing is generally used effectively to support overall coherence, and the sequencing of ideas within a paragraph is generally logical.
Band 6
Information and ideas are generally arranged coherently and there is a clear overall progression.
Cohesive devices are used to some good effect but cohesion within and/or between sentences may be faulty or mechanical due to misuse, overuse or omission.
The use of reference and substitution may lack flexibility or clarity and result in some repetition or error.
Paragraphing may not always be logical and/or the central topic may not always be clear.
Band 5
Organisation is evident but is not wholly logical and there may be a lack of overall progression.
Nevertheless, there is a sense of underlying coherence to the response.
The relationship of ideas can be followed but the sentences are not fluently linked to each other.
There may be limited/overuse of cohesive devices with some inaccuracy.
The writing may be repetitive due to inadequate and/or inaccurate use of reference and substitution.
Paragraphing may be inadequate or missing.
Band 4
Information and ideas are evident but not arranged coherently and there is no clear progression within the response.
Relationships between ideas can be unclear and/or inadequately marked. There is some use of basic cohesive devices, which may be inaccurate or repetitive.
There is inaccurate use or a lack of substitution or referencing.
There may be no paragraphing and/or no clear main topic within paragraphs.


Lexical Resource
Band 9
Full flexibility and precise use are widely evident.
A wide range of vocabulary is used accurately and appropriately with very natural and sophisticated control of lexical features.
Minor errors in spelling and word formation are extremely rare and have minimal impact on communication.
Band 8
A wide resource is fluently and flexibly used to convey precise meanings.
There is skilful use of uncommon and/or idiomatic items when appropriate, despite occasional inaccuracies in word choice and collocation.
Occasional errors in spelling and/or word formation may occur, but have minimal impact on communication.
Band 7
The resource is sufficient to allow some flexibility and precision.
There is some ability to use less common and/or idiomatic items.
An awareness of style and collocation is evident, though inappropriacies occur.
There are only a few errors in spelling and/or word formation and they do not detract from overall clarity.
Band 6
The resource is generally adequate and appropriate for the task.
The meaning is generally clear in spite of a rather restricted range or a lack of precision in word choice.
If the writer is a risk-taker, there will be a wider range of vocabulary used but higher degrees of inaccuracy or inappropriacy.
There are some errors in spelling and/or word formation, but these do not impede communication.
Band 5
The resource is limited but minimally adequate for the task.
Simple vocabulary may be used accurately but the range does not permit much variation in expression.
There may be frequent lapses in the appropriacy of word choice and a lack of flexibility is apparent in frequent simplifications and/or repetitions.
Errors in spelling and/or word formation may be noticeable and may cause some difficulty for the reader.
Band 4
The resource is limited and inadequate for or unrelated to the task. Vocabulary is basic and may be used repetitively.
There may be inappropriate use of lexical chunks (e.g. memorised phrases, formulaic language and/or language from the input material).
Inappropriate word choice and/or errors in word formation and/or in spelling may impede meaning.


Grammar Range and Accuracy
Band 9
A wide range of structures is used with full flexibility and control.
Punctuation and grammar are used appropriately throughout.
Minor errors are extremely rare and have minimal impact on communication.
Band 8
A wide range of structures is flexibly and accurately used.
The majority of sentences are error-free, and punctuation is well managed.
Occasional, non-systematic errors and inappropriacies occur, but have minimal impact on communication.
Band 7
A variety of complex structures is used with some flexibility and accuracy.
Grammar and punctuation are generally well controlled, and error-free sentences are frequent.
A few errors in grammar may persist, but these do not impede communication.
Band 6
A mix of simple and complex sentence forms is used but flexibility is limited.
Examples of more complex structures are not marked by the same level of accuracy as in simple structures.
Errors in grammar and punctuation occur, but rarely impede communication.
Band 5
The range of structures is limited and rather repetitive.
Although complex sentences are attempted, they tend to be faulty, and the greatest accuracy is achieved on simple sentences.
Grammatical errors may be frequent and cause some difficulty for the reader.
Punctuation may be faulty.
Band 4
A very limited range of structures is used.
Subordinate clauses are rare and simple sentences predominate.
Some structures are produced accurately but grammatical errors are frequent and may impede meaning.
Punctuation is often faulty or inadequate.

Total: 36 points
"""

anthropic_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# ─────────────────────────────────────────────
#  GOOGLE DRIVE + DOCS CLIENT
# ─────────────────────────────────────────────
try:
    _gdocs_creds = service_account.Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_JSON,
        scopes=[
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/documents",
        ]
    )
    docs_service = build("docs", "v1", credentials=_gdocs_creds, cache_discovery=False)
except Exception as _e:
    docs_service = None
    print(f"[WARN] Google Docs not configured: {_e}")

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────

def get_canvas_sections(course_id, assignment_id=None):
    """Fetch all sections for a course from Canvas.

    If assignment_id is provided, also counts how many students in each
    section have a submitted (non-unsubmitted) submission for that assignment,
    so the UI can warn when a section has no submissions.
    """
    headers = {"Authorization": f"Bearer {CANVAS_API_TOKEN}"}

    # 1. Fetch all sections
    url = f"https://{CANVAS_DOMAIN}/api/v1/courses/{course_id}/sections"
    all_sections = []
    params = {"include[]": "total_students", "per_page": 100}
    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=20)
        resp.raise_for_status()
        all_sections.extend(resp.json())
        link = resp.headers.get("Link", "")
        next_url = None
        for part in link.split(","):
            if 'rel="next"' in part:
                next_url = part.split(";")[0].strip().strip("<>")
        url = next_url
        params = {}

    sections = [
        {"id": s["id"], "name": s["name"], "total_students": s.get("total_students", 0)}
        for s in all_sections
    ]

    if not assignment_id:
        return sections

    # 2. Fetch all course submissions once and build a lookup
    sub_url = f"https://{CANVAS_DOMAIN}/api/v1/courses/{course_id}/assignments/{assignment_id}/submissions"
    all_subs = []
    sub_params = {"per_page": 100}
    while sub_url:
        resp = requests.get(sub_url, headers=headers, params=sub_params, timeout=20)
        resp.raise_for_status()
        all_subs.extend(resp.json())
        link = resp.headers.get("Link", "")
        next_url = None
        for part in link.split(","):
            if 'rel="next"' in part:
                next_url = part.split(";")[0].strip().strip("<>")
        sub_url = next_url
        sub_params = {}

    submitted_ids = {
        s["user_id"] for s in all_subs
        if s.get("workflow_state") not in ("unsubmitted", "deleted")
        and (s.get("url") or s.get("body"))
    }

    # 3. Per section: count enrolled students who have a submission
    for section in sections:
        sid = section["id"]
        enroll_resp = requests.get(
            f"https://{CANVAS_DOMAIN}/api/v1/sections/{sid}/enrollments",
            headers=headers,
            params={"type[]": "StudentEnrollment", "state[]": "active", "per_page": 100},
            timeout=20,
        )
        enrolled_ids = {e["user_id"] for e in enroll_resp.json()}
        section["submission_count"] = len(enrolled_ids & submitted_ids)

    return sections


def _get_section_student_ids(section_id):
    """Return the list of active student user_ids enrolled in a section."""
    url = f"https://{CANVAS_DOMAIN}/api/v1/sections/{section_id}/enrollments"
    headers = {"Authorization": f"Bearer {CANVAS_API_TOKEN}"}
    student_ids = []
    params = {"type[]": "StudentEnrollment", "state[]": "active", "per_page": 100}

    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=20)
        resp.raise_for_status()
        student_ids.extend(e["user_id"] for e in resp.json())
        link = resp.headers.get("Link", "")
        next_url = None
        for part in link.split(","):
            if 'rel="next"' in part:
                next_url = part.split(";")[0].strip().strip("<>")
        url = next_url
        params = {}

    return student_ids


def get_canvas_submissions(course_id, assignment_id, section_id=None):
    """Fetch submissions (with user info) from Canvas.

    If section_id is provided, fetches all course submissions then filters
    to only students enrolled in that section. (Canvas does not reliably
    support student_ids[] filtering or the section submissions endpoint.)
    """
    headers = {"Authorization": f"Bearer {CANVAS_API_TOKEN}"}
    url = f"https://{CANVAS_DOMAIN}/api/v1/courses/{course_id}/assignments/{assignment_id}/submissions"
    params = {"include[]": "user", "per_page": 100}
    all_subs = []

    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=20)
        resp.raise_for_status()
        all_subs.extend(resp.json())
        link = resp.headers.get("Link", "")
        next_url = None
        for part in link.split(","):
            if 'rel="next"' in part:
                next_url = part.split(";")[0].strip().strip("<>")
        url = next_url
        params = {}

    if section_id:
        allowed_ids = set(_get_section_student_ids(section_id))
        all_subs = [s for s in all_subs if s.get("user_id") in allowed_ids]

    return all_subs


def extract_gdoc_id(text):
    """Pull a Google Doc ID out of any string (URL, HTML body, etc.)."""
    if not text:
        return None
    match = re.search(r"docs\.google\.com/document/d/([a-zA-Z0-9_-]+)", text)
    return match.group(1) if match else None


def fetch_gdoc_text(doc_id):
    """Export a Google Doc as plain text, using the service account when available."""
    # Try authenticated export via Drive API first (works for private docs)
    if docs_service:
        try:
            drive = build("drive", "v3", credentials=_gdocs_creds, cache_discovery=False)
            resp = drive.files().export(fileId=doc_id, mimeType="text/plain").execute()
            text = resp.decode("utf-8") if isinstance(resp, bytes) else resp
            return text.strip() if text else None
        except Exception:
            pass
    # Fall back to unauthenticated public export
    try:
        url = f"https://docs.google.com/document/d/{doc_id}/export?format=txt"
        resp = requests.get(url, allow_redirects=True, timeout=20)
        if resp.status_code == 200:
            return resp.text.strip()
    except Exception:
        pass
    return None


# Maps our breakdown keys → keyword to match in Canvas criterion descriptions
CRITERION_KEY_MAP = {
    "task_response":    "task response",
    "coherence":        "coherence",
    "lexical_resource": "lexical",
    "grammar":          "grammar",
}

def get_rubric_criteria(course_id, assignment_id):
    """Fetch assignment rubric from Canvas.

    Returns a dict keyed by our breakdown keys:
    {
      "task_response": {
          "criterion_id": "_6718",
          "ratings": {9: "blank", 8: "_7435", 7: "_6849", ...}
      },
      ...
    }
    """
    url = f"https://{CANVAS_DOMAIN}/api/v1/courses/{course_id}/assignments/{assignment_id}"
    headers = {"Authorization": f"Bearer {CANVAS_API_TOKEN}"}
    resp = requests.get(url, headers=headers, timeout=20)
    resp.raise_for_status()
    rubric = resp.json().get("rubric", [])

    mapping = {}
    for criterion in rubric:
        description = criterion.get("description", "").lower()
        cid = criterion["id"]
        # Build points → rating_id lookup for this criterion
        ratings_map = {
            int(r["points"]): r["id"]
            for r in criterion.get("ratings", [])
        }
        for key, keyword in CRITERION_KEY_MAP.items():
            if keyword in description:
                mapping[key] = {"criterion_id": cid, "ratings": ratings_map}
                break
    return mapping


CRITERION_LABELS = {
    "task_response":    "Task Response",
    "coherence":        "Coherence and Cohesion",
    "lexical_resource": "Lexical Resource",
    "grammar":          "Grammar Range and Accuracy",
}

def _gdocs_len(s):
    """String length in Google Docs index units (UTF-16 code units).
    Characters outside the BMP (e.g. emoji) count as 2 units each.
    """
    return sum(2 if ord(c) > 0xFFFF else 1 for c in s)


# Per-criterion label colors (RGB 0-1 scale)
_CRITERION_RGB = {
    "task_response":    {"red": 0.10, "green": 0.45, "blue": 0.91},  # blue
    "coherence":        {"red": 0.48, "green": 0.18, "blue": 0.55},  # purple
    "lexical_resource": {"red": 0.88, "green": 0.40, "blue": 0.00},  # orange
    "grammar":          {"red": 0.11, "green": 0.53, "blue": 0.31},  # green
}


def append_gdoc_feedback(doc_id, overall_feedback, breakdown, criterion_comments, criterion_inline=None, task_type="task2"):
    """Append AI feedback as colored text at the end of a Google Doc.

    Uses the Docs API batchUpdate to insert formatted text in one atomic call.
    Each criterion header is colored; quotes are italic-gray; issues are red;
    suggestions are green.
    """
    if not docs_service:
        raise RuntimeError("Google Docs service not configured.")

    if criterion_inline is None:
        criterion_inline = {}

    # ── Find document end index ──────────────────────────────────────────────
    doc = docs_service.documents().get(documentId=doc_id).execute()
    body_content = doc.get("body", {}).get("content", [])
    # endIndex of the last element includes the trailing sentinel \n; insert just before it
    end_index = body_content[-1].get("endIndex", 1) - 1

    # ── Build text + style segments ──────────────────────────────────────────
    parts = []          # text strings
    styles = []         # (start_offset, end_offset, textStyle_dict, fields_str)
    offset = 0          # running offset in Docs index units

    def emit(text, style=None, fields=None):
        nonlocal offset
        parts.append(text)
        ln = _gdocs_len(text)
        if style and fields:
            styles.append((offset, offset + ln, style, fields))
        offset += ln

    CRITERION_ORDER = [
        ("task_response",    "Task Achievement" if task_type == "task1" else "Task Response"),
        ("coherence",        "Coherence & Cohesion"),
        ("lexical_resource", "Lexical Resource"),
        ("grammar",          "Grammar Range & Accuracy"),
    ]

    # Separator header
    sep = "-" * 44
    emit(f"\n\n{sep}\n")
    emit(
        "  AI FEEDBACK\n",
        {"bold": True, "fontSize": {"magnitude": 13, "unit": "PT"},
         "foregroundColor": {"color": {"rgbColor": {"red": 0.13, "green": 0.13, "blue": 0.13}}}},
        "bold,fontSize,foregroundColor",
    )
    emit(f"{sep}\n\n")

    for key, label in CRITERION_ORDER:
        score   = breakdown.get(key, "?")
        summary = criterion_comments.get(key, "").strip()
        inlines = criterion_inline.get(key, [])
        rgb     = _CRITERION_RGB.get(key, {"red": 0.3, "green": 0.3, "blue": 0.3})

        # Criterion header — colored + bold
        emit(
            f"[{label}]  Band {score}/9\n",
            {"bold": True,
             "fontSize": {"magnitude": 11, "unit": "PT"},
             "foregroundColor": {"color": {"rgbColor": rgb}}},
            "bold,fontSize,foregroundColor",
        )

        # Summary text — dark gray
        if summary:
            emit(
                summary + "\n\n",
                {"foregroundColor": {"color": {"rgbColor": {"red": 0.15, "green": 0.15, "blue": 0.15}}}},
                "foregroundColor",
            )
        else:
            emit("\n")

        # Inline tips
        for item in inlines:
            quote      = item.get("quote", "").strip()
            issue      = item.get("issue", "").strip()
            suggestion = item.get("suggestion", "").strip()
            if not issue:
                continue
            if quote:
                emit(
                    f'  "{quote}"\n',
                    {"italic": True,
                     "foregroundColor": {"color": {"rgbColor": {"red": 0.45, "green": 0.45, "blue": 0.45}}}},
                    "italic,foregroundColor",
                )
            emit(
                f"  ! {issue}\n",
                {"bold": True,
                 "foregroundColor": {"color": {"rgbColor": {"red": 0.78, "green": 0.07, "blue": 0.12}}}},
                "bold,foregroundColor",
            )
            if suggestion:
                emit(
                    f"  -> {suggestion}\n",
                    {"foregroundColor": {"color": {"rgbColor": {"red": 0.11, "green": 0.53, "blue": 0.31}}}},
                    "foregroundColor",
                )

        emit("\n")

    # Overall feedback
    if overall_feedback:
        emit(
            "[Overall Feedback]\n",
            {"bold": True,
             "fontSize": {"magnitude": 11, "unit": "PT"},
             "foregroundColor": {"color": {"rgbColor": {"red": 0.20, "green": 0.20, "blue": 0.20}}}},
            "bold,fontSize,foregroundColor",
        )
        emit(overall_feedback + "\n")

    full_text = "".join(parts)

    # ── Build batchUpdate request list ───────────────────────────────────────
    reqs = [
        {"insertText": {"location": {"index": end_index}, "text": full_text}}
    ]
    for (start, end, style, fields) in styles:
        reqs.append({
            "updateTextStyle": {
                "range": {
                    "startIndex": end_index + start,
                    "endIndex":   end_index + end,
                },
                "textStyle": style,
                "fields":    fields,
            }
        })

    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": reqs},
    ).execute()

    return True, None


_task_images: dict = {}


def grade_with_claude(essay_text, total_points, rubric_text,
                      task_type="task2", essay_topic="",
                      image_data=None, image_media_type=None):
    """Send essay to Claude and get back structured grades. Rubric is cached."""
    task_label = (
        "IELTS Writing Task 1 (Academic — describe a visual: graph, chart, diagram, or map)"
        if task_type == "task1"
        else "IELTS Writing Task 2 (Extended essay — argument, opinion, or discussion)"
    )
    topic_block = f"\nTASK PROMPT:\n{essay_topic}" if essay_topic.strip() else ""

    user_content: list = []
    if task_type == "task1" and image_data and image_media_type:
        user_content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": image_media_type, "data": image_data},
        })
    user_content.append({
        "type": "text",
        "text": f"{topic_block}\n\nESSAY:\n{essay_text[:5000]}".lstrip(),
    })

    response = anthropic_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2000,
        temperature=0.1,
        system=[
            {
                "type": "text",
                "text": f"""You are a strict, experienced IELTS writing examiner. Your default assumption is that student work has weaknesses — your job is to find them and explain them clearly.

SCORING PHILOSOPHY — be strict:
- Award the band the response EARNS, not the band that would encourage the student. Most student responses fall in Band 5–6.
- A Band 7 requires consistent control with only minor lapses. Band 8+ is rare and requires near-native fluency.
- Deduct a full band for each significant recurring error type (e.g. systematic article errors, repeated off-topic paragraphs, no overview in Task 1).
- Do NOT round up. If a response is between two bands, assign the lower one.

INLINE COMMENTS — concise and targeted:
- Pick the 2–4 most impactful errors per criterion only.
- Quote the EXACT phrase (3–8 words max).
- Problem statement: one short clause — no padding.
- Suggestion: one short rewrite or fix — no explanations.
- Prioritise errors that most lower the band score.

Good inline comment examples:
- quote: "students become outdated", issue: "subject is wrong — knowledge becomes outdated, not students", suggestion: "their knowledge becomes outdated"
- quote: "Firstly", issue: "'Firstly' implies a numbered list but no second point follows in the paragraph", suggestion: "One reason is that"
- quote: "jobs", issue: "imprecise — use a formal, specific word", suggestion: "occupation / employment"
- quote: "the graph shows an increase", issue: "vague — no values or time period cited", suggestion: "the proportion rose sharply from 20% in 2000 to 35% by 2020"
- quote: "However, universities", issue: "'However' signals contrast but the previous sentence is not a contrasting idea", suggestion: "Furthermore, / In addition,"
- quote: "In my opinion", issue: "personal opinion is inappropriate in Task 1 — report data only", suggestion: "Overall, it is evident that"
- quote: "make a research", issue: "wrong collocation", suggestion: "conduct research / carry out a study"

RUBRIC:
{rubric_text}""",
                "cache_control": {"type": "ephemeral"},
            },
            {
                "type": "text",
                "text": f"""TASK TYPE: {task_label}

{"TASK 1 — key criteria:" if task_type == "task1" else "TASK 2 — key criteria:"}
{"""Task Achievement: overview present, all key features covered, data accurate, no opinion, more than 150 words.
Coherence: intro paraphrases task, body groups trends logically, cohesive devices varied.
Lexical: data-description verbs (peaked, fluctuated, rose sharply), no verb repetition.
Grammar: passives for reporting, comparatives correct, subject-verb agreement.
Flag: missing overview, data misread, listing every figure, copied prompt, "In my opinion".
""" if task_type == "task1" else """Task Response: all prompt parts addressed, clear position maintained, ideas developed with specific support, more than 250 words.
Coherence: 4-paragraph structure, topic sentence per paragraph, cohesive devices varied.
Lexical: topic vocabulary accurate, correct collocations, no informal language.
Grammar: mixed sentence types, correct tense/articles/agreement.
Flag: ignoring one part of the prompt, contradicting own position, weak examples, over-generalising.
"""}
Return ONLY valid JSON — no extra text. Use this exact structure:
{{
  "task_response_score":      <integer 0–9, scoring {'Task Achievement' if task_type == 'task1' else 'Task Response'}>,
  "coherence_score":          <integer 0–9>,
  "lexical_resource_score":   <integer 0–9>,
  "grammar_score":            <integer 0–9>,
  "total_score":              <integer 0–{total_points}>,
  "task_response_comment":    "<1–2 sentence summary of {'Task Achievement' if task_type == 'task1' else 'Task Response'} performance>",
  "task_response_inline": [
    {{"quote": "<exact phrase>", "issue": "<problem>", "suggestion": "<fix>"}},
    ...2–4 items max...
  ],
  "coherence_comment":        "<1–2 sentence summary of Coherence and Cohesion>",
  "coherence_inline": [
    {{"quote": "<exact phrase>", "issue": "<problem>", "suggestion": "<fix>"}},
    ...2–4 items max...
  ],
  "lexical_resource_comment": "<1–2 sentence summary of Lexical Resource>",
  "lexical_resource_inline": [
    {{"quote": "<exact phrase>", "issue": "<problem>", "suggestion": "<fix>"}},
    ...2–4 items max...
  ],
  "grammar_comment":          "<1–2 sentence summary of Grammar Range and Accuracy>",
  "grammar_inline": [
    {{"quote": "<exact phrase>", "issue": "<problem>", "suggestion": "<fix>"}},
    ...2–4 items max...
  ],
  "feedback": "<2–3 sentence summary: what was done well + top 2 priorities to improve>"
}}""",
            },
        ],
        messages=[
            {"role": "user", "content": user_content},
            {"role": "assistant", "content": "{"},
        ],
    )
    raw = "{" + response.content[0].text
    # Strip markdown code fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw.strip())
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


# ─────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("templates", "index.html")


_rubric_store: dict = {}


@app.route("/api/store-rubric", methods=["POST"])
def store_rubric():
    """Store a custom rubric server-side and return a short key."""
    rubric = (request.json or {}).get("rubric", "").strip()
    if not rubric:
        return jsonify({"error": "rubric is required"}), 400
    key = str(uuid.uuid4())
    _rubric_store[key] = rubric
    return jsonify({"key": key})


@app.route("/api/upload-task-image", methods=["POST"])
def upload_task_image():
    file = request.files.get("image")
    if not file:
        return jsonify({"error": "No image provided"}), 400
    data = base64.b64encode(file.read()).decode("utf-8")
    raw_type = (file.content_type or "").split(";")[0].strip().lower()
    ext = (file.filename or "").rsplit(".", 1)[-1].lower()
    _ext_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                "gif": "image/gif", "webp": "image/webp"}
    _allowed = {"image/jpeg", "image/png", "image/gif", "image/webp"}
    media_type = raw_type if raw_type in _allowed else _ext_map.get(ext, "image/jpeg")
    key = str(uuid.uuid4())
    _task_images[key] = {"data": data, "media_type": media_type}
    return jsonify({"key": key})


@app.route("/api/grade-stream")
def grade_stream():
    """Server-Sent Events: grades each student one-by-one.
    Query params: course_id, assignment_id, total_points, task_type, essay_topic, task_image_key"""
    course_id       = request.args.get("course_id", "").strip()
    assignment_id   = request.args.get("assignment_id", "").strip()
    total_points    = int(request.args.get("total_points", 36))
    rubric_key      = request.args.get("rubric_key", "").strip()
    rubric_text     = _rubric_store.get(rubric_key) or request.args.get("rubric", "").strip() or RUBRIC
    section_id      = request.args.get("section_id", "").strip() or None
    task_type       = request.args.get("task_type", "task2").strip() or "task2"
    essay_topic     = request.args.get("essay_topic", "").strip()
    task_image_key  = request.args.get("task_image_key", "").strip() or None

    img = _task_images.get(task_image_key) if task_image_key else None
    image_data       = img["data"] if img else None
    image_media_type = img["media_type"] if img else None

    def generate():
        if not course_id or not assignment_id:
            yield f"data: {json.dumps({'type': 'error', 'message': 'Course ID and Assignment ID are required.'})}\n\n"
            return

        try:
            yield f"data: {json.dumps({'type': 'progress', 'message': 'Fetching submissions from Canvas…'})}\n\n"

            submissions = get_canvas_submissions(course_id, assignment_id, section_id)
            active = [
                s for s in submissions
                if s.get("workflow_state") not in ("unsubmitted", "deleted")
                and s.get("user_id")
            ]

            if len(active) == 0 and section_id:
                yield f"data: {json.dumps({'type': 'error', 'message': f'No submissions found for this section with Assignment ID {assignment_id}. This section may use a different Assignment ID — check the Canvas URL when viewing this section.'})}\n\n"
                return

            yield f"data: {json.dumps({'type': 'progress', 'message': f'Found {len(active)} submission(s). Starting grading…', 'current': 0, 'total': len(active)})}\n\n"

            for i, sub in enumerate(active):
                student_name = sub.get("user", {}).get("name", f"Student {sub['user_id']}")
                student_id   = sub["user_id"]
                raw          = sub.get("url") or sub.get("body") or ""
                doc_id       = extract_gdoc_id(raw)

                yield f"data: {json.dumps({'type': 'progress', 'message': f'Grading {student_name}…', 'current': i, 'total': len(active)})}\n\n"

                google_doc_url = f"https://docs.google.com/document/d/{doc_id}/edit" if doc_id else raw
                base = {
                    "student":        student_name,
                    "student_id":     student_id,
                    "google_doc_url": google_doc_url,
                    "doc_id":         doc_id,
                    "max_score":      total_points,
                }

                def _err(msg):
                    return json.dumps({
                        "type": "result", "current": i + 1, "total": len(active),
                        "data": {**base, "error": msg},
                    })

                if not doc_id:
                    yield f"data: {_err('No Google Doc link found in submission.')}\n\n"
                    continue

                essay_text = fetch_gdoc_text(doc_id)
                if not essay_text:
                    yield f"data: {_err('Could not read the Google Doc (check sharing settings).')}\n\n"
                    continue

                try:
                    result = grade_with_claude(
                        essay_text, total_points, rubric_text,
                        task_type, essay_topic, image_data, image_media_type,
                    )

                    def _inline_issues(key):
                        items = []
                        for it in _inline_raw(key):
                            quote = it.get("quote", "")
                            issue = it.get("issue", "")
                            suggestion = it.get("suggestion", "")
                            line = f'"{quote}" — {issue}'
                            if suggestion:
                                line += f" → {suggestion}"
                            items.append(line)
                        return items

                    KEYS = ("task_response", "coherence", "lexical_resource", "grammar")
                    # For Task 1 Claude may return task_achievement_* instead of task_response_*
                    def _score(key):
                        v = result.get(f"{key}_score")
                        if v is None and key == "task_response":
                            v = result.get("task_achievement_score")
                        return v or 0
                    def _comment(key):
                        v = result.get(f"{key}_comment", "")
                        if not v and key == "task_response":
                            v = result.get("task_achievement_comment", "")
                        return v
                    def _inline_raw(key):
                        v = result.get(f"{key}_inline")
                        if not v and key == "task_response":
                            v = result.get("task_achievement_inline")
                        return v or []
                    criterion_scores   = {k: _score(k) for k in KEYS}
                    criterion_comments = {k: _comment(k) for k in KEYS}
                    criterion_inline   = {k: _inline_raw(k) for k in KEYS}

                    breakdown = {
                        k: {
                            "score":         criterion_scores[k],
                            "max":           9,
                            "justification": criterion_comments[k],
                            "issues":        _inline_issues(k),
                        }
                        for k in KEYS
                    }

                    yield f"data: {json.dumps({'type': 'result', 'current': i + 1, 'total': len(active), 'data': {**base, 'score': result['total_score'], 'overall_feedback': result['feedback'], 'breakdown': breakdown, 'criterion_scores': criterion_scores, 'criterion_comments': criterion_comments, 'criterion_inline': criterion_inline, 'task_type': task_type}})}\n\n"

                except Exception as e:
                    yield f"data: {_err(f'Grading error: {e}')}\n\n"

        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"

        yield f"data: {json.dumps({'type': 'done'})}\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/default-rubric")
def default_rubric():
    """Return the server-side default rubric text."""
    return jsonify({"rubric": RUBRIC.strip()})


@app.route("/api/canvas-assignments")
def canvas_assignments():
    """Return all assignments for a Canvas course."""
    course_id = request.args.get("course_id", "").strip()
    if not course_id:
        return jsonify({"error": "course_id is required"}), 400
    try:
        url = f"https://{CANVAS_DOMAIN}/api/v1/courses/{course_id}/assignments"
        headers = {"Authorization": f"Bearer {CANVAS_API_TOKEN}"}
        all_assignments = []
        params = {"per_page": 100, "order_by": "due_at"}
        while url:
            resp = requests.get(url, headers=headers, params=params, timeout=20)
            resp.raise_for_status()
            all_assignments.extend(resp.json())
            link = resp.headers.get("Link", "")
            next_url = None
            for part in link.split(","):
                if 'rel="next"' in part:
                    next_url = part.split(";")[0].strip().strip("<>")
            url = next_url
            params = {}
        result = [
            {
                "id":             a["id"],
                "name":           a.get("name", ""),
                "points_possible": a.get("points_possible") or 0,
                "due_at":         a.get("due_at", ""),
            }
            for a in all_assignments
        ]
        return jsonify({"assignments": result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/canvas-sections")
def canvas_sections():
    """Return all sections for a Canvas course, with submission counts per section."""
    course_id     = request.args.get("course_id", "").strip()
    assignment_id = request.args.get("assignment_id", "").strip() or None
    if not course_id:
        return jsonify({"error": "course_id is required"}), 400
    try:
        sections = get_canvas_sections(course_id, assignment_id)
        return jsonify({"sections": sections})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/rubric-criteria")
def rubric_criteria():
    """Return the Canvas rubric criterion IDs + rating IDs for an assignment."""
    course_id     = request.args.get("course_id", "").strip()
    assignment_id = request.args.get("assignment_id", "").strip()
    if not course_id or not assignment_id:
        return jsonify({"error": "course_id and assignment_id are required"}), 400
    try:
        return jsonify(get_rubric_criteria(course_id, assignment_id))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/submit-grades", methods=["POST"])
def submit_grades():
    """Post rubric assessment (per-criterion scores + rating_id + comments) to Canvas."""
    body          = request.json
    grades        = body.get("grades", [])
    course_id     = body.get("course_id", "")
    assignment_id = body.get("assignment_id", "")

    headers = {"Authorization": f"Bearer {CANVAS_API_TOKEN}"}
    results = []

    # Fetch criterion IDs + rating maps once for the whole batch
    try:
        criteria = get_rubric_criteria(course_id, assignment_id)
    except Exception as e:
        return jsonify({"error": f"Could not fetch rubric criteria: {e}"}), 500

    for g in grades:
        if g.get("score") is None:
            results.append({"student_id": g["student_id"], "success": False, "reason": "No score"})
            continue

        url = (
            f"https://{CANVAS_DOMAIN}/api/v1/courses/{course_id}"
            f"/assignments/{assignment_id}/submissions/{g['student_id']}"
        )

        breakdown = g.get("breakdown", {})
        criterion_comments = g.get("criterion_comments", {})
        payload = {"comment[text_comment]": g.get("feedback", "")}

        # Submit score + matching rating_id + comment per rubric criterion
        for key, info in criteria.items():
            cid      = info["criterion_id"]
            pts      = int(breakdown.get(key, 0))
            rating   = info["ratings"].get(pts, "")
            comment  = criterion_comments.get(key, "")
            payload[f"rubric_assessment[{cid}][points]"]    = pts
            payload[f"rubric_assessment[{cid}][rating_id]"] = rating
            payload[f"rubric_assessment[{cid}][comments]"]  = comment

        resp = requests.put(url, headers=headers, data=payload, timeout=20)
        results.append({
            "student_id":  g["student_id"],
            "success":     resp.status_code == 200,
            "status_code": resp.status_code,
        })

    submitted = sum(1 for r in results if r["success"])
    return jsonify({"submitted": submitted, "total": len(results), "results": results})


@app.route("/api/post-gdoc-comments", methods=["POST"])
def post_gdoc_comments_route():
    """Append AI feedback as colored text at the end of each student's Google Doc."""
    body   = request.json
    grades = body.get("grades", [])

    results = []
    for g in grades:
        doc_id = g.get("doc_id")
        if not doc_id:
            results.append({"student_id": g["student_id"], "success": False, "reason": "No doc_id"})
            continue
        try:
            ok, err = append_gdoc_feedback(
                doc_id,
                g.get("feedback", ""),
                g.get("breakdown", {}),
                g.get("criterion_comments", {}),
                g.get("criterion_inline", {}),
                task_type=g.get("task_type", "task2"),
            )
            results.append({
                "student_id": g["student_id"],
                "success":    ok,
                **({"reason": err} if err else {}),
            })
        except Exception as e:
            results.append({"student_id": g["student_id"], "success": False, "reason": str(e)})

    posted = sum(1 for r in results if r["success"])
    return jsonify({"posted": posted, "total": len(results), "results": results})



if __name__ == "__main__":
    print("🎓  Canvas Grader running at http://localhost:5000")
    app.run(debug=True, port=5000)

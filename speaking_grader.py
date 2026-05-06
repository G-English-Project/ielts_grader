"""
Speaking grader — Google S2T transcription + Claude IELTS band scoring.

Flow:
  1. Download audio from Canvas embedded video or Google Drive link
  2. Convert to mono 16 kHz FLAC via ffmpeg (in-memory temp files)
  3. Transcribe with Google Cloud Speech-to-Text (long_running_recognize)
  4. Compute fluency / pronunciation metrics from word-level S2T output
  5. Grade all 4 IELTS criteria with Claude
"""

import json
import os
import re
import subprocess
import uuid
from datetime import datetime, timezone

import anthropic
import requests
from collections import namedtuple

from google.cloud import storage
from google.cloud.speech_v2 import SpeechClient
from google.cloud.speech_v2.types import cloud_speech
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CANVAS_DOMAIN = "canvas.instructure.com"

FILLERS = {"uh", "um", "uhm", "huh", "ah", "er", "hmm", "oh"}

# Map our internal keys → substring to match in Canvas rubric descriptions
SPEAKING_CRITERION_KEY_MAP = {
    "fluency":          "fluency",
    "lexical_resource": "lexical",
    "grammar":          "grammar",
    "pronunciation":    "pronunciation",
}

SPEAKING_CRITERION_LABELS = {
    "fluency":          "Fluency & Coherence",
    "lexical_resource": "Lexical Resource",
    "grammar":          "Grammar Range and Accuracy",
    "pronunciation":    "Pronunciation",
}

SPEAKING_RUBRIC = """\
Fluency & Coherence (9 points — whole band scores 4–9):
Band 9: Fluent with only very occasional repetition/self-correction. Any hesitation is content-related. Speech cohesive; topic fully coherent and extended.
Band 8: Fluent. Hesitation occasionally for words/grammar but mostly content-related. Coherent and relevant.
Band 7: Keeps going without noticeable effort. Some hesitation/repetition/self-correction mid-sentence but coherence unaffected. Flexible discourse markers.
Band 6: Able to keep going; some willingness to produce long turns. Coherence may be lost due to hesitation. Uses discourse markers though not always appropriately.
Band 5: Usually keeps going but relies on repetition/self-correction/slow speech. Hesitations often for basic lexis/grammar. Overuse of certain discourse markers.
Band 4: Cannot keep going without noticeable pauses. Slow, frequent repetition. Simple sentences linked but with repetitious connectives. Some coherence breakdowns.

Lexical Resource (9 points):
Band 9: Total flexibility and precise use in all contexts. Sustained accurate and idiomatic language.
Band 8: Wide resource, flexibly used. Skilful use of less common/idiomatic items despite occasional inaccuracies. Effective paraphrase.
Band 7: Flexibly used for variety of topics. Some less common/idiomatic items; style/collocation awareness evident though inappropriacies occur. Effective paraphrase.
Band 6: Sufficient to discuss topics at length. Vocabulary may be inappropriate but meaning clear. Generally able to paraphrase.
Band 5: Sufficient for familiar and unfamiliar topics but limited flexibility. Attempts paraphrase but not always with success.
Band 4: Sufficient for familiar topics only. Basic meaning conveyed on unfamiliar topics. Frequent inappropriacies and word-choice errors. Rarely attempts paraphrase.

Grammar Range and Accuracy (9 points):
Band 9: Structures precise and accurate at all times (apart from native-speaker-type slips).
Band 8: Wide range of structures, flexibly used. Majority error-free. Occasional non-systematic errors.
Band 7: Range of structures flexibly used. Error-free sentences frequent. Simple and complex structures used effectively despite some errors.
Band 6: Mix of short and complex forms, limited flexibility. Errors frequent in complex structures but rarely impede communication.
Band 5: Basic forms fairly well controlled. Complex structures attempted but limited in range and often contain errors.
Band 4: Basic sentence forms; some short utterances error-free. Structures repetitive; errors frequent; short turns.

Pronunciation (9 points):
Band 9: Full range of phonological features to convey meaning. Effortlessly understood; accent has no effect on intelligibility.
Band 8: Wide range of phonological features. Sustained rhythm, flexible stress/intonation. Easily understood; accent has minimal effect.
Band 7: Displays all positive features of band 6 and some of band 8.
Band 6: Range of phonological features but variable control. Chunking generally appropriate but rhythm may be affected. Some effective intonation/stress but not sustained. Individual words/phonemes mispronounced occasionally. Generally understood without much effort.
Band 5: Displays all positive features of band 4 and some of band 6.
Band 4: Some acceptable phonological features but limited range. Frequent rhythm lapses. Limited intonation/stress control. Words frequently mispronounced, causing lack of clarity. Understanding requires effort.
"""


# ---------------------------------------------------------------------------
# Google S2T client
# ---------------------------------------------------------------------------

_CANVAS_UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

def _canvas_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "User-Agent": _CANVAS_UA,
        "Accept": "application/json",
    }


def _s2t_client(region="us-central1"):
    sa_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_path:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON env var not set")
    creds = service_account.Credentials.from_service_account_file(
        sa_path,
        scopes=["https://www.googleapis.com/auth/cloud-platform"],
    )
    from google.api_core.client_options import ClientOptions
    return SpeechClient(
        credentials=creds,
        client_options=ClientOptions(api_endpoint=f"{region}-speech.googleapis.com"),
    )


# ---------------------------------------------------------------------------
# Step 1 — Resolve source URL (don't download the file into Python memory)
# ---------------------------------------------------------------------------

def resolve_audio_urls(submission, canvas_token):
    """Return list of (url, ext) for ALL audio/video sources found in a submission body.

    Supports multiple Canvas iframes and/or multiple Google Drive links in one submission.
    Each URL is already resolved to a direct CDN/usercontent URL (no auth headers needed).
    """
    body = submission.get("body") or ""
    sources = []

    # All Canvas media attachment iframes — each has its own attachment ID + verifier
    for m in re.finditer(r'/media_attachments_iframe/(\d+)[^"\']*[?&]verifier=([A-Za-z0-9]+)', body):
        att_id, verifier = m.group(1), m.group(2)
        dl_url = (
            f"https://{CANVAS_DOMAIN}/files/{att_id}/download"
            f"?download_frd=1&verifier={verifier}"
        )
        try:
            resp = requests.get(dl_url, headers=_canvas_headers(canvas_token),
                                timeout=30, allow_redirects=True, stream=True)
            resp.raise_for_status()
            resp.close()
            sources.append((resp.url, "mp4"))
        except Exception as e:
            print(f"[speaking] warning: could not resolve Canvas attachment {att_id}: {e}")

    # All Google Drive file links
    for m in re.finditer(r'drive\.google\.com/file/d/([A-Za-z0-9_-]+)', body):
        file_id = m.group(1)
        # Avoid duplicates
        if any(file_id in url for url, _ in sources):
            continue
        dl_url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download"
        try:
            resp = requests.get(dl_url, timeout=30, allow_redirects=True, stream=True)
            resp.raise_for_status()
            ctype = resp.headers.get("Content-Type", "")
            ext = "m4a" if "audio/mp4" in ctype or "m4a" in ctype else "mp4"
            resp.close()
            sources.append((resp.url, ext))
        except Exception as e:
            print(f"[speaking] warning: could not resolve Drive file {file_id}: {e}")

    if not sources:
        raise ValueError("No supported audio source found in submission (Canvas media or Google Drive link required)")

    return sources


# ---------------------------------------------------------------------------
# Step 2 — Extract & concatenate audio from one or more URLs via ffmpeg
# ---------------------------------------------------------------------------

def extract_audio_ogg(sources):
    """Extract audio-only from one or more (url, ext) sources and return merged OGG_OPUS bytes.

    Multiple sources are concatenated in order (e.g. Part 1 + Part 2 of an exam).
    Key flags: -vn skips video decode entirely (much faster), output goes to stdout.
    """
    cmd = ["ffmpeg", "-y"]

    for url, _ in sources:
        cmd += ["-i", url]

    if len(sources) > 1:
        # Concat filter: merge all audio streams sequentially
        inputs = "".join(f"[{i}:a]" for i in range(len(sources)))
        cmd += [
            "-filter_complex", f"{inputs}concat=n={len(sources)}:v=0:a=1[outa]",
            "-map", "[outa]",
        ]
    else:
        cmd += ["-vn"]   # single source — simpler path

    cmd += [
        "-ac", "1",
        "-ar", "16000",
        "-c:a", "libopus",
        "-b:a", "64k",       # raised from 32k — better clarity for accented speech
        "-f", "ogg",
        "pipe:1",
    ]

    proc = subprocess.run(cmd, capture_output=True, timeout=600)
    if proc.returncode != 0:
        raise RuntimeError(f"ffmpeg failed: {proc.stderr.decode()[:600]}")
    if not proc.stdout:
        raise RuntimeError("ffmpeg produced no output — check URL or audio track")
    return proc.stdout


# ---------------------------------------------------------------------------
# Step 3 — Transcribe via GCS (handles any duration)
# ---------------------------------------------------------------------------

def _gcs_client():
    sa_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    creds = service_account.Credentials.from_service_account_file(
        sa_path,
        scopes=["https://www.googleapis.com/auth/cloud-platform"],
    )
    return storage.Client(credentials=creds)


def transcribe_audio(ogg_bytes):
    """Upload OGG to GCS, transcribe via URI (no inline size/duration limit), delete after.

    Requires GCS_BUCKET env var pointing to a bucket where the service account
    has Storage Object Admin role.
    """
    bucket_name = os.getenv("GCS_BUCKET", "").strip()
    if not bucket_name:
        raise ValueError(
            "GCS_BUCKET env var not set. Add GCS_BUCKET=<your-bucket-name> to .env"
        )

    # Upload to GCS with a unique temp key
    gcs  = _gcs_client()
    blob_name = f"s2t-temp/{uuid.uuid4().hex}.ogg"
    bucket = gcs.bucket(bucket_name)
    blob   = bucket.blob(blob_name)

    try:
        blob.upload_from_string(ogg_bytes, content_type="audio/ogg")
        gcs_uri = f"gs://{bucket_name}/{blob_name}"
        print(f"[speaking] uploaded {len(ogg_bytes)//1024} KB to {gcs_uri}")

        # Transcribe from GCS URI using Speech-to-Text v2 + chirp_2 model
        sa_path    = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
        with open(sa_path) as _f:
            project_id = json.load(_f)["project_id"]
        recognizer = f"projects/{project_id}/locations/us-central1/recognizers/_"

        s2t    = _s2t_client(region="us-central1")
        config = cloud_speech.RecognitionConfig(
            auto_decoding_config=cloud_speech.AutoDetectDecodingConfig(),
            language_codes=["en-US"],
            model="chirp_2",
            features=cloud_speech.RecognitionFeatures(
                enable_word_time_offsets=True,
                enable_word_confidence=True,
                enable_automatic_punctuation=True,
            ),
            adaptation=cloud_speech.SpeechAdaptation(
                phrase_sets=[
                    cloud_speech.SpeechAdaptation.AdaptationPhraseSet(
                        inline_phrase_set=cloud_speech.PhraseSet(
                            phrases=[
                                cloud_speech.PhraseSet.Phrase(value=p, boost=15.0)
                                for p in [
                                    "furthermore", "in addition", "on the other hand",
                                    "however", "consequently", "therefore", "moreover",
                                    "in conclusion", "to summarize", "for instance",
                                    "for example", "as a result", "in contrast",
                                    "nevertheless", "in my opinion", "I believe",
                                    "to begin with", "first of all", "on balance",
                                ]
                            ],
                        )
                    )
                ]
            ),
        )
        request = cloud_speech.BatchRecognizeRequest(
            recognizer=recognizer,
            config=config,
            files=[cloud_speech.BatchRecognizeFileMetadata(uri=gcs_uri)],
            recognition_output_config=cloud_speech.RecognitionOutputConfig(
                inline_response_config=cloud_speech.InlineOutputConfig(),
            ),
        )
        operation = s2t.batch_recognize(request=request)
        response  = operation.result(timeout=600)

    finally:
        # Always delete the temp file from GCS
        try:
            blob.delete()
            print(f"[speaking] deleted temp GCS object {blob_name}")
        except Exception as e:
            print(f"[speaking] warning: could not delete GCS object {blob_name}: {e}")

    # v2 batch_recognize returns {gcs_uri: FileResult}; we always have one file
    _Word = namedtuple("_Word", ["word", "start_time", "end_time", "confidence"])
    parts, words = [], []
    for file_result in response.results.values():
        for result in file_result.transcript.results:
            alt = result.alternatives[0]
            parts.append(alt.transcript)
            for w in alt.words:
                words.append(_Word(
                    word=w.word,
                    start_time=w.start_offset,
                    end_time=w.end_offset,
                    confidence=w.confidence,
                ))
    return " ".join(parts), words


# ---------------------------------------------------------------------------
# Step 4 — Compute metrics
# ---------------------------------------------------------------------------

def compute_metrics(transcript, words):
    """Return dict of fluency and pronunciation metrics derived from S2T word list."""
    if not words:
        return {"error": "No words returned by S2T"}

    def _secs(t):
        return t.total_seconds()

    # Duration
    last = words[-1]
    duration = _secs(last.end_time)

    word_count = len(words)
    wpm = round(word_count / duration * 60) if duration > 0 else 0

    # Fillers
    filler_list = [w.word for w in words if w.word.lower().strip(".,?!") in FILLERS]
    filler_count = len(filler_list)

    # Pauses (gap > 0.4 s between consecutive words)
    pauses = []
    for i in range(1, len(words)):
        prev_end  = _secs(words[i-1].end_time)
        cur_start = _secs(words[i].start_time)
        gap = cur_start - prev_end
        if gap > 0.4:
            pauses.append(round(gap, 2))

    # Low-confidence words — threshold 0.5 avoids flooding the list with
    # accent-related ASR uncertainty that isn't true mispronunciation.
    low_conf = [
        {"word": w.word, "confidence": round(w.confidence, 2)}
        for w in words if w.confidence < 0.5
    ]
    high_conf_count = sum(1 for w in words if w.confidence >= 0.5)
    intelligibility_pct = round(high_conf_count / word_count * 100) if word_count else 0

    return {
        "duration_secs":      round(duration),
        "word_count":         word_count,
        "wpm":                wpm,
        "filler_count":       filler_count,
        "filler_rate_pct":    round(filler_count / word_count * 100, 1) if word_count else 0,
        "filler_examples":    filler_list[:15],
        "pause_count":        len(pauses),
        "avg_pause_secs":     round(sum(pauses) / len(pauses), 2) if pauses else 0,
        "max_pause_secs":     round(max(pauses), 2) if pauses else 0,
        "low_conf_words":     low_conf[:20],
        "intelligibility_pct": intelligibility_pct,
    }


# ---------------------------------------------------------------------------
# Step 5 — Claude grading
# ---------------------------------------------------------------------------

def grade_with_claude(transcript, metrics, topic):
    """Return grading dict with band + comments for each of the 4 criteria."""
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    # Only surface very low confidence words (< 0.5) — higher threshold causes
    # false positives on accented speech that S2T simply isn't calibrated for.
    very_low_conf = [w for w in metrics.get("low_conf_words", []) if w["confidence"] < 0.5]
    low_conf_str = (
        ", ".join(f'"{w["word"]}"' for w in very_low_conf[:10])
        or "none — pronunciation was generally intelligible to the ASR engine"
    )

    wpm   = metrics.get("wpm", 0)
    frate = metrics.get("filler_rate_pct", 0)

    # Translate raw numbers into a plain-English description so Claude
    # doesn't over-penalise based on a number without context.
    wpm_note = (
        "slightly slow (may reflect careful pacing, not necessarily disfluency)" if wpm < 100
        else "natural pace for EFL speaker" if wpm < 140
        else "fast — check for clarity"
    )
    filler_note = (
        "minimal fillers — strong fluency signal" if frate < 3
        else "moderate fillers — typical of B2 EFL speakers" if frate < 8
        else "frequent fillers — noticeable impact on fluency"
    )

    prompt = f"""You are an experienced IELTS Speaking examiner grading a video submission from a Vietnamese university EFL student.

IMPORTANT CONTEXT — read before grading:
• The transcript below was produced by automatic speech recognition (ASR/Google S2T).
  ASR makes errors on accented speech: words may be mis-spelled, split, or dropped.
  Do NOT penalise grammar or vocabulary for what look like transcription artifacts
  (e.g. a missing article or odd word could be an ASR error, not the student's error).
  Grade based on the overall pattern, not isolated surface forms.
• S2T word-confidence scores are biased against non-native accents — a low score
  means the ASR was uncertain, NOT that the student definitely mispronounced the word.
  Use the low-confidence list only as a soft hint, not as a definitive error list.
• These are university students who have studied English for several years.
  A realistic score distribution for this cohort is Band 5–7. Reserve Band 4 for
  speech that is very difficult to follow; reserve Band 8+ for near-native output.
• Grade the COMMUNICATION SUCCESS — could a native English speaker understand
  the student without significant effort? That should anchor your judgment.

ASSIGNMENT TOPIC: {topic or "IELTS Speaking test"}

TRANSCRIPT (auto-generated — may contain ASR errors):
{transcript}

FLUENCY SIGNALS (from audio timing analysis):
- Duration: {metrics.get('duration_secs', '?')} s | Words: {metrics.get('word_count', '?')} | WPM: {wpm} ({wpm_note})
- Fillers (uh/um/er): {metrics.get('filler_count', 0)} ({frate}% of words) — {filler_note}
- Pauses >0.4 s: {metrics.get('pause_count', 0)}, avg {metrics.get('avg_pause_secs', 0)} s, max {metrics.get('max_pause_secs', 0)} s

PRONUNCIATION SIGNAL (ASR confidence — interpret cautiously for accented speech):
- Words the ASR engine was most uncertain about: {low_conf_str}

IELTS SPEAKING RUBRIC:
{SPEAKING_RUBRIC}

GRADING INSTRUCTIONS:
1. Grade holistically. Read the full transcript first to understand what the student is communicating.
2. Assign a whole-number band (4–9) for each criterion. Most students will fall in 5–7.
3. For each criterion write 2–3 specific, constructive sentences:
   - Start with what the student did WELL (give concrete examples from the transcript).
   - Then note 1–2 specific areas to improve with actionable advice.
4. For pronunciation: base your judgment on the overall intelligibility impression,
   not just the ASR low-confidence list. Mention specific sounds or words only if
   you see a clear pattern (e.g. missing final consonants, /θ/ errors).
5. For grammar: focus on structural range and frequency of errors that impede meaning —
   ignore minor slips that don't affect comprehension.
6. In "annotations", mark specific phrases from the transcript for colour-highlighting in the report:
   - "good_vocab": up to 5 impressive or sophisticated vocabulary items / phrases used correctly
   - "vocab_error": up to 5 wrong word choices or collocation errors
   - "grammar_error": up to 5 clear grammatical mistakes (include 2–5 words of context)
   - "filler": up to 10 hesitation words exactly as they appear (uh, um, er, oh, ah, hmm)
   - "pronunciation_error": up to 3 words from the low-confidence list that are likely mispronounced
   CRITICAL: "text" must be an EXACT substring copied character-for-character from the transcript.
   Keep each annotation short (1–5 words). Only mark high-confidence cases.

Respond with ONLY this JSON (no markdown fences, no extra text):
{{
  "fluency": {{
    "band": <integer 4–9>,
    "comments": "<2–3 sentences>"
  }},
  "lexical_resource": {{
    "band": <integer 4–9>,
    "comments": "<2–3 sentences>"
  }},
  "grammar": {{
    "band": <integer 4–9>,
    "comments": "<2–3 sentences>"
  }},
  "pronunciation": {{
    "band": <integer 4–9>,
    "comments": "<2–3 sentences>"
  }},
  "overall_comments": "<3–4 sentences summarising overall performance: estimated band average, key strengths, top priority to improve>",
  "annotations": [
    {{"text": "<exact phrase from transcript>", "type": "<good_vocab|vocab_error|grammar_error|filler|pronunciation_error>"}}
  ]
}}"""

    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()
    print(f"[speaking] Claude raw response:\n{raw[:800]}")

    match = re.search(r'\{.*\}', raw, re.DOTALL)
    if not match:
        raise ValueError(f"Claude returned no JSON object. Raw response: {raw[:300]}")
    cleaned = re.sub(r',\s*([}\]])', r'\1', match.group())
    result = json.loads(cleaned)
    print(f"[speaking] Bands — fluency:{result.get('fluency',{}).get('band')} "
          f"lex:{result.get('lexical_resource',{}).get('band')} "
          f"gram:{result.get('grammar',{}).get('band')} "
          f"pron:{result.get('pronunciation',{}).get('band')}")
    return result


# ---------------------------------------------------------------------------
# Rubric helpers (reused by app.py route)
# ---------------------------------------------------------------------------

def get_speaking_rubric_criteria(course_id, assignment_id, canvas_token):
    """Return criterion_id + ratings_map for the 4 speaking criteria.

    Structure mirrors get_rubric_criteria() in app.py:
    {
      "fluency": {"criterion_id": "_3746", "ratings": {9: "blank", 8: "_6396", ...}},
      ...
    }
    """
    url = f"https://{CANVAS_DOMAIN}/api/v1/courses/{course_id}/assignments/{assignment_id}"
    resp = requests.get(url, headers=_canvas_headers(canvas_token), timeout=20)
    resp.raise_for_status()
    rubric = resp.json().get("rubric", [])

    mapping = {}
    for criterion in rubric:
        desc = criterion.get("description", "").lower()
        cid = criterion["id"]
        ratings_map = {int(r["points"]): r["id"] for r in criterion.get("ratings", [])}
        for key, keyword in SPEAKING_CRITERION_KEY_MAP.items():
            if keyword in desc:
                mapping[key] = {"criterion_id": cid, "ratings": ratings_map}
                break
    return mapping


# ---------------------------------------------------------------------------
# Google Doc report
# ---------------------------------------------------------------------------

REPORT_FOLDER_ID = "1uuWhN29u6x0ETZCH6ikO-bq_Ws4vaS5q"
REPORT_DOC_ID    = "1h6w6WsK8zEc59eoryzceJUv7E99y_7rOdsRtHr-OR58"

CRITERIA_ORDER = ["fluency", "lexical_resource", "grammar", "pronunciation"]
_CRIT_LABEL = {
    "fluency":          "Fluency & Coherence",
    "lexical_resource": "Lexical Resource",
    "grammar":          "Grammar Range & Accuracy",
    "pronunciation":    "Pronunciation",
}

# Highlight colours (RGB 0–1) for transcript annotation types
_ANN = {
    "good_vocab":         {"bg": {"red": 0.56, "green": 0.93, "blue": 0.56}},
    "vocab_error":        {"bg": {"red": 1.0,  "green": 0.4,  "blue": 0.4}},
    "grammar_error":      {"bg": {"red": 1.0,  "green": 1.0,  "blue": 0.4}},
    "filler":             {"fg": {"red": 0.85, "green": 0.0,  "blue": 0.0}},
    "pronunciation_error":{"bg": {"red": 1.0,  "green": 0.73, "blue": 0.3}},
}
_LEGEND = [
    (" Good vocab ",             "good_vocab"),
    (" Vocab/collocation error ","vocab_error"),
    (" Grammar error ",          "grammar_error"),
    (" Filler / hesitation ",    "filler"),
    (" Pronunciation error ",    "pronunciation_error"),
]

# Foreground (text) colours for the four IELTS criteria in the scores section
_CRIT_FG = {
    "fluency":          {"red": 0.1,  "green": 0.3,  "blue": 0.8},
    "lexical_resource": {"red": 0.07, "green": 0.53, "blue": 0.07},
    "grammar":          {"red": 0.75, "green": 0.35, "blue": 0.0},
    "pronunciation":    {"red": 0.5,  "green": 0.0,  "blue": 0.6},
}


def _docs_clients():
    """Return (docs_service, drive_service) using the service account."""
    sa_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    creds = service_account.Credentials.from_service_account_file(
        sa_path,
        scopes=[
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/documents",
        ],
    )
    docs  = build("docs",  "v1", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    return docs, drive


def create_speaking_report(students_data, title, existing_doc_id=None, tab_title=None):
    """Write a speaking report into a Google Doc tab.

    If existing_doc_id is given, appends into a tab named tab_title (creates
    the tab if it doesn't exist yet).  All students for the same assignment
    are written into the same tab, so re-running appends new students without
    duplicating existing ones.

    students_data: list of dicts with keys:
        name, transcript, bands, comments, overall_comments, submitted_at,
        submission_type, topic

    Returns (doc_url, doc_id).
    """
    docs_svc, drive_svc = _docs_clients()

    if existing_doc_id:
        doc_id  = existing_doc_id
        doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"
        # Find current end of document to append after existing content
        doc          = docs_svc.documents().get(documentId=doc_id).execute()
        body_content = doc.get("body", {}).get("content", [])
        end_idx      = body_content[-1].get("endIndex", 2)
        insert_at       = end_idx - 1   # just before final sentinel \n
        doc_has_content = insert_at > 1
    else:
        file_meta = {
            "name":     title,
            "mimeType": "application/vnd.google-apps.document",
            "parents":  [REPORT_FOLDER_ID],
        }
        created         = drive_svc.files().create(body=file_meta, fields="id,webViewLink").execute()
        doc_id          = created["id"]
        doc_url         = created["webViewLink"]
        insert_at       = 1
        doc_has_content = False

    # ── Build segments and inline colour spans ───────────────────────────────
    segments     = []   # (text, para_style, bold)
    inline_spans = []   # (start_in_fulltext, end_in_fulltext, textStyle_dict, fields_str)
    _pos         = [0]  # running character count within full_text

    def seg(text, style=None, bold=False):
        segments.append((text, style, bold))
        _pos[0] += len(text)

    def seg_colored(text, ann_type, style=None, bold=False):
        """Write text highlighted with an annotation type colour."""
        start = _pos[0]
        seg(text, style=style, bold=bold)
        end = _pos[0]
        colors = _ANN.get(ann_type, {})
        if "bg" in colors:
            inline_spans.append((start, end,
                {"backgroundColor": {"color": {"rgbColor": colors["bg"]}}},
                "backgroundColor"))
        if "fg" in colors:
            inline_spans.append((start, end,
                {"foregroundColor": {"color": {"rgbColor": colors["fg"]}}},
                "foregroundColor"))

    def seg_fg(text, fg, style=None, bold=False):
        """Write text with an explicit foreground (text) colour."""
        start = _pos[0]
        seg(text, style=style, bold=bold)
        end = _pos[0]
        inline_spans.append((start, end,
            {"foregroundColor": {"color": {"rgbColor": fg}}},
            "foregroundColor"))

    # ── Legend (written once, at the top of each new section) ────────────────
    def _write_legend():
        seg("Legend: ", bold=True)
        for label, ann_type in _LEGEND:
            seg_colored(label, ann_type)
            seg("  ")
        seg("\n\n")

    # ── Assignment section header ─────────────────────────────────────────────
    if tab_title:
        if doc_has_content:
            seg("\n\n")
        seg(f"{tab_title}\n", "HEADING_1")
        seg("\n")
        _write_legend()
        section_open = True
    else:
        section_open = False
        if not doc_has_content:
            _write_legend()

    for i, s in enumerate(students_data):
        name             = s.get("name") or f"Student {i+1}"
        transcript       = s.get("transcript", "").strip()
        bands            = s.get("bands", {})
        comments         = s.get("comments", {})
        overall_comments = s.get("overall_comments", "").strip()
        metrics          = s.get("metrics", {})
        submitted_at     = s.get("submitted_at", "")
        submission_type  = s.get("submission_type", "")
        topic            = s.get("topic", "")
        annotations      = s.get("annotations", [])
        total            = sum(int(bands.get(k, 0)) for k in CRITERIA_ORDER)

        # Separator between students
        if i > 0 or (doc_has_content and not section_open):
            seg("\n" + "─" * 60 + "\n\n")

        # ── Student name — Heading 2 ──
        seg(f"{name}\n", "HEADING_2")

        # ── Metadata block ──
        meta_lines = []
        if submitted_at:
            try:
                dt = datetime.fromisoformat(submitted_at.replace("Z", "+00:00"))
                submitted_fmt = dt.strftime("%d %B %Y, %H:%M").lstrip("0")
            except Exception:
                submitted_fmt = submitted_at[:16]
            meta_lines.append(f"Submitted:    {submitted_fmt}")
        if submission_type:
            meta_lines.append(f"Type:         {submission_type.replace('_', ' ').title()}")
        if topic:
            meta_lines.append(f"Topic:        {topic}")
        if meta_lines:
            seg("\n".join(meta_lines) + "\n")

        # ── Speaking stats line ──
        if metrics:
            dur   = metrics.get("duration_secs", 0)
            mins, secs = divmod(int(dur), 60)
            dur_fmt = f"{mins}m {secs:02d}s" if mins else f"{secs}s"
            stats = (
                f"WPM: {metrics.get('wpm', '—')}  ·  "
                f"Words: {metrics.get('word_count', '—')}  ·  "
                f"Duration: {dur_fmt}  ·  "
                f"Fillers: {metrics.get('filler_count', '—')}  ·  "
                f"Long pauses: {metrics.get('pause_count', '—')}  ·  "
                f"Avg pause: {metrics.get('avg_pause_secs', '—')}s  ·  "
                f"Intelligibility: {metrics.get('intelligibility_pct', '—')}%"
            )
            seg(stats + "\n\n")
        else:
            seg("\n")

        # ── Transcript — Heading 3 ──
        seg("Transcript\n", "HEADING_3")
        trans_start = _pos[0]          # remember where transcript text begins
        seg(f"{transcript}\n\n")

        # Apply annotation highlights within this transcript
        print(f"[report] {name}: {len(annotations)} annotations from Claude")
        for ann in annotations:
            phrase   = ann.get("text", "").strip()
            ann_type = ann.get("type", "")
            colors   = _ANN.get(ann_type)
            if not phrase or not colors:
                continue
            for m in re.finditer(re.escape(phrase), transcript, re.IGNORECASE):
                span_s = trans_start + m.start()
                span_e = trans_start + m.end()
                if "bg" in colors:
                    inline_spans.append((span_s, span_e,
                        {"backgroundColor": {"color": {"rgbColor": colors["bg"]}}},
                        "backgroundColor"))
                if "fg" in colors:
                    inline_spans.append((span_s, span_e,
                        {"foregroundColor": {"color": {"rgbColor": colors["fg"]}}},
                        "foregroundColor"))

        # Auto-highlight fillers from metrics (guarantees red text even if Claude
        # did not annotate them explicitly)
        filler_words = set(w.lower() for w in metrics.get("filler_examples", []))
        filler_colors = _ANN["filler"]
        for word in filler_words:
            for m in re.finditer(r'\b' + re.escape(word) + r'\b', transcript, re.IGNORECASE):
                inline_spans.append((trans_start + m.start(), trans_start + m.end(),
                    {"foregroundColor": {"color": {"rgbColor": filler_colors["fg"]}}},
                    "foregroundColor"))

        # Auto-highlight low-confidence words from S2T as orange pronunciation errors
        pron_bg = _ANN["pronunciation_error"]["bg"]
        low_conf_words = set(w["word"].lower() for w in metrics.get("low_conf_words", []))
        for word in low_conf_words:
            for m in re.finditer(r'\b' + re.escape(word) + r'\b', transcript, re.IGNORECASE):
                inline_spans.append((trans_start + m.start(), trans_start + m.end(),
                    {"backgroundColor": {"color": {"rgbColor": pron_bg}}},
                    "backgroundColor"))

        # ── Scores — Heading 3 ──
        seg("Scores & Feedback\n", "HEADING_3")
        for key in CRITERIA_ORDER:
            band    = bands.get(key, "?")
            label   = _CRIT_LABEL[key]
            comment = comments.get(key, "")
            fg      = _CRIT_FG[key]
            seg_fg(f"{label}: Band {band}\n", fg, bold=True)
            seg_fg(f"{comment}\n\n",           fg)
        seg(f"Total: {total} / 36\n", bold=True)

        # ── Overall Comments — Heading 3 ──
        if overall_comments:
            seg("\nOverall Comments\n", "HEADING_3")
            seg(f"{overall_comments}\n")

    # ── Execute: insert full text then apply all styles ──────────────────────
    full_text = "".join(text for text, _, _ in segments)
    if not full_text:
        return doc_url, doc_id

    requests_list = [
        {"insertText": {"location": {"index": insert_at}, "text": full_text}}
    ]

    # Paragraph styles and bold
    cursor = insert_at
    for text, style, bold in segments:
        start = cursor
        end   = cursor + len(text)
        if style:
            requests_list.append({
                "updateParagraphStyle": {
                    "range":          {"startIndex": start, "endIndex": end},
                    "paragraphStyle": {"namedStyleType": style},
                    "fields":         "namedStyleType",
                }
            })
        if bold:
            requests_list.append({
                "updateTextStyle": {
                    "range":     {"startIndex": start, "endIndex": end - 1},
                    "textStyle": {"bold": True},
                    "fields":    "bold",
                }
            })
        cursor = end

    # Inline colour spans (highlights + red text)
    for span_s, span_e, text_style, fields in inline_spans:
        doc_s = insert_at + span_s
        doc_e = insert_at + span_e
        if doc_s >= doc_e:
            continue
        requests_list.append({
            "updateTextStyle": {
                "range":     {"startIndex": doc_s, "endIndex": doc_e},
                "textStyle": text_style,
                "fields":    fields,
            }
        })

    docs_svc.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests_list},
    ).execute()

    return doc_url, doc_id

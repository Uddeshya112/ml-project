"""
Web API + Frontend Server
==========================
Serves the HTML frontend (frontend/index.html) and exposes the existing
from-scratch NLP pipeline over a small JSON API.

NOTHING in src/ or app.py is modified — this file only *imports* the same
modules the Streamlit dashboard uses:

    preprocessing -> vectorization (TF / IDF / TF-IDF) -> cosine similarity
    -> skill extraction -> ranking

Run:
    python server.py
    # then open http://localhost:5000

Endpoints:
    GET  /               The frontend
    GET  /api/health     Liveness check
    GET  /api/samples    Sample JDs and resumes from data/
    POST /api/parse      Extract text from uploaded .pdf/.docx/.txt/.md/.csv
    POST /api/analyze    Rank resumes against a job description
"""

import os
import sys

from flask import Flask, jsonify, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.preprocessing.tokenizer import tokenize
from src.preprocessing.stopwords import remove_stopwords
from src.preprocessing.stemmer import stem_tokens
from src.vectorization.tf import compute_tf
from src.vectorization.idf import compute_idf
from src.vectorization.tfidf import compute_tfidf, get_top_terms
from src.similarity.ranker import rank_resumes, extract_skills, load_skills_database
from src.parsers.pdf_parser import extract_text_from_pdf
from src.parsers.docx_parser import extract_text_from_docx

FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")
SAMPLE_RESUMES_DIR = os.path.join(BASE_DIR, "data", "sample_resumes")
SAMPLE_JDS_DIR = os.path.join(BASE_DIR, "data", "sample_job_descriptions")

MAX_UPLOAD_MB = 25
ALLOWED_EXTENSIONS = (".pdf", ".docx", ".txt", ".md", ".csv")

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def _read_text_file(path):
    """Read a text file, falling back to latin-1 if it is not valid UTF-8."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        with open(path, "r", encoding="latin-1") as f:
            return f.read()


def _decode_bytes(raw):
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def _prettify(filename, strip_suffix=""):
    stem = os.path.splitext(filename)[0]
    if strip_suffix and stem.endswith(strip_suffix):
        stem = stem[: -len(strip_suffix)]
    return stem.replace("_", " ").title()


def _prep_steps(text):
    """Same three steps as ranker.preprocess(), but keep each stage's output."""
    tokens = tokenize(text)
    no_stop = remove_stopwords(tokens)
    stems = stem_tokens(no_stop)
    return tokens, no_stop, stems


def _skills_payload(skills):
    """Convert extract_skills() output into the shape the frontend uses."""
    return {"found": skills["all_skills"], "byCategory": skills["by_category"]}


def _error(message, status=400):
    return jsonify({"error": message}), status


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
@app.get("/")
def index():
    return send_from_directory(FRONTEND_DIR, "index.html")


@app.get("/api/health")
def health():
    return jsonify({"status": "ok"})


@app.get("/api/samples")
def samples():
    jds, resumes = [], []

    if os.path.isdir(SAMPLE_JDS_DIR):
        for fname in sorted(os.listdir(SAMPLE_JDS_DIR)):
            if fname.lower().endswith(".txt"):
                jds.append({
                    "id": os.path.splitext(fname)[0],
                    "label": _prettify(fname, "_jd"),
                    "text": _read_text_file(os.path.join(SAMPLE_JDS_DIR, fname)),
                })

    if os.path.isdir(SAMPLE_RESUMES_DIR):
        for fname in sorted(os.listdir(SAMPLE_RESUMES_DIR)):
            if fname.lower().endswith(".txt"):
                resumes.append({
                    "id": "sample_" + os.path.splitext(fname)[0],
                    "name": fname,
                    "text": _read_text_file(os.path.join(SAMPLE_RESUMES_DIR, fname)),
                })

    return jsonify({"jds": jds, "resumes": resumes})


@app.post("/api/parse")
def parse_files():
    """Extract plain text from uploaded resume files (same parsers as app.py)."""
    files = request.files.getlist("files")
    if not files:
        return _error("No files were uploaded.")

    out = []
    for f in files:
        name = f.filename or "resume"
        lower = name.lower()
        entry = {"name": name, "text": "", "error": None}
        try:
            if lower.endswith(".pdf"):
                entry["text"] = extract_text_from_pdf(f)
            elif lower.endswith(".docx"):
                entry["text"] = extract_text_from_docx(f)
            elif lower.endswith((".txt", ".md", ".csv")):
                entry["text"] = _decode_bytes(f.read())
            else:
                entry["error"] = "Unsupported file type (use " + ", ".join(ALLOWED_EXTENSIONS) + ")"
        except Exception as exc:  # parser errors are reported per-file
            entry["error"] = str(exc)

        if not entry["error"] and not (entry["text"] or "").strip():
            entry["error"] = "No readable text found in this file."
        out.append(entry)

    return jsonify({"files": out})


@app.post("/api/analyze")
def analyze():
    """
    Body: {"jd": "<job description text>",
           "resumes": [{"id": "...", "name": "...", "text": "..."}, ...]}
    """
    data = request.get_json(silent=True) or {}
    jd_text = (data.get("jd") or "").strip()
    resumes_in = data.get("resumes") or []

    if not jd_text:
        return _error("Job description is empty.")

    # Keep only resumes that actually contain text; key by a unique id so
    # duplicate file names cannot overwrite each other.
    by_id, names = {}, {}
    for i, r in enumerate(resumes_in):
        text = (r.get("text") or "").strip()
        if not text:
            continue
        rid = str(r.get("id") or f"resume_{i}")
        if rid in by_id:
            rid = f"{rid}_{i}"
        by_id[rid] = text
        names[rid] = r.get("name") or rid

    if not by_id:
        return _error("Provide at least one resume with text.")

    # ---- Core ranking: the exact same function the Streamlit app calls ----
    ranked = rank_resumes(jd_text, by_id, detailed=True)

    # ---- Extra data for the math walkthrough (uses the same modules) ----
    jd_tokens, jd_no_stop, jd_stems = _prep_steps(jd_text)
    prep = {rid: _prep_steps(text) for rid, text in by_id.items()}
    idf = compute_idf([jd_stems] + [p[2] for p in prep.values()])
    jd_skills = extract_skills(jd_text, load_skills_database())

    candidates = []
    for r in ranked:
        rid = r["name"]  # rank_resumes() names results by the dict key (our id)
        tokens, no_stop, stems = prep[rid]
        tf = compute_tf(stems)
        tfidf = compute_tfidf(stems, idf)
        d = r["details"] or {}

        candidates.append({
            "id": rid,
            "name": names[rid],
            "cosinePct": r["score_pct"],
            "skillMatchPct": r["skill_match_rate"],
            "matchedSkills": r["matched_skills"],
            "missingSkills": r["missing_skills"],
            "extraSkills": r["extra_skills"],
            "prep": {"tokens": tokens, "noStop": no_stop, "stems": stems},
            "cosDetail": {
                "similarity": d.get("similarity", r["score"]),
                "dot": d.get("dot_product", 0.0),
                "magA": d.get("magnitude_a", 0.0),
                "magB": d.get("magnitude_b", 0.0),
            },
            "topTerms": [
                {"term": term, "tf": tf.get(term, 0.0), "idf": idf.get(term, 0.0), "tfidf": weight}
                for term, weight in get_top_terms(tfidf, n=10)
            ],
        })

    return jsonify({
        "jdSkills": _skills_payload(jd_skills),
        "jdPrep": {"tokens": jd_tokens, "noStop": jd_no_stop, "stems": jd_stems},
        "candidates": candidates,
    })


@app.errorhandler(413)
def too_large(_e):
    return _error(f"Upload too large (limit {MAX_UPLOAD_MB} MB).", 413)


@app.errorhandler(500)
def server_error(_e):
    return _error("Internal server error while processing the request.", 500)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

"""
DOCX Episode Splitter — FastAPI Backend
Detects episodes via configurable regex patterns, splits DOCX files,
and returns a ZIP archive for download.
"""

from __future__ import annotations

import os
import re
import uuid
import shutil
import zipfile
import tempfile
from pathlib import Path
from typing import Optional

import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# ─────────────────────────────────────────────────────────────
# App setup
# ─────────────────────────────────────────────────────────────
app = FastAPI(
    title="DOCX Episode Splitter",
    description="Upload a multi-episode DOCX → detect → split → download ZIP",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = Path(tempfile.gettempdir()) / "ep_splitter"
UPLOAD_DIR.mkdir(exist_ok=True)

# In-memory job store  {job_id → dict}
JOBS: dict[str, dict] = {}


# ─────────────────────────────────────────────────────────────
# Episode detection — add new patterns here, order = priority
# ─────────────────────────────────────────────────────────────
EPISODE_PATTERNS: list[re.Pattern] = [
    re.compile(r"^Ep\s+(\d+)(?:\s*[:\-–—]\s*(.*))?$", re.IGNORECASE),
    re.compile(r"^Episode\s+(\d+)(?:\s*[:\-–—]\s*(.*))?$", re.IGNORECASE),
    re.compile(r"^Chapter\s+(\d+)(?:\s*[:\-–—]\s*(.*))?$", re.IGNORECASE),
    re.compile(r"^Part\s+(\d+)(?:\s*[:\-–—]\s*(.*))?$", re.IGNORECASE),
]
HEADING_EPISODE_PATTERN = re.compile(
    r"^(?:Ep|Episode|Chapter|Part)\s+(\d+)(?:\s*[:\-–—]\s*|\s+)?(.*)$",
    re.IGNORECASE,
)


def detect_episode(text: str) -> Optional[dict]:
    """Return {number, title, heading} if text matches an episode heading."""
    text = text.strip()
    for pat in EPISODE_PATTERNS:
        m = pat.match(text)
        if m:
            title = m.group(2) if len(m.groups()) > 1 and m.group(2) else ""
            return {
                "number": int(m.group(1)),
                "title": title.strip(),
                "heading": text,
            }
    return None


def detect_heading_episode(text: str) -> Optional[dict]:
    """Detect a numbered episode in an explicit Word heading paragraph."""
    text = text.strip()
    match = HEADING_EPISODE_PATTERN.match(text)
    if not match:
        return None
    return {
        "number": int(match.group(1)),
        "title": (match.group(2) or "").strip(),
        "heading": text,
    }


def is_heading_style(style_id: str) -> bool:
    """Return True for Word's built-in numbered heading styles."""
    return bool(re.fullmatch(r"Heading[1-6]", style_id or "", re.IGNORECASE))


def body_child_text(child) -> str:
    """Extract visible text from one direct document-body child."""
    return "".join(node.text or "" for node in child.iter(qn("w:t")))


# ─────────────────────────────────────────────────────────────
# DOCX analysis
# ─────────────────────────────────────────────────────────────
def analyse_docx(doc_path: str) -> dict:
    doc = Document(doc_path)
    body = doc.element.body
    children = list(body.iterchildren())
    candidates: list[dict] = []

    # Only direct body paragraphs can be chapter boundaries. Paragraphs nested
    # inside tables are document content, not structural chapter markers.
    for body_idx, child in enumerate(children):
        if child.tag != qn("w:p"):
            continue
        para = Paragraph(child, doc)
        style_id = para.style.style_id if para.style is not None else ""
        ep = detect_heading_episode(para.text) if is_heading_style(style_id) else detect_episode(para.text)
        if not ep:
            continue
        candidates.append({**ep, "boundary_idx": body_idx, "style_id": style_id})

    # Prefer explicit Word headings when the document supplies them. Otherwise
    # use strict, punctuation-aware text matches. Keep the first occurrence of
    # each number so repeated running titles cannot split a chapter again.
    heading_candidates = [c for c in candidates if is_heading_style(c["style_id"])]
    selected = heading_candidates or candidates
    boundaries: list[dict] = []
    seen_numbers: set[int] = set()
    for candidate in selected:
        if candidate["number"] in seen_numbers:
            continue
        seen_numbers.add(candidate["number"])
        boundaries.append(candidate)

    episodes: list[dict] = []
    content_indices = [i for i, child in enumerate(children) if child.tag != qn("w:sectPr")]
    last_content_idx = content_indices[-1] if content_indices else -1

    for i, boundary in enumerate(boundaries):
        # Preserve front matter in the first output and trailing material in the
        # last output. Across a complete split, no document-body element is lost.
        start_idx = 0 if i == 0 else boundary["boundary_idx"]
        end_idx = (
            boundaries[i + 1]["boundary_idx"] - 1
            if i + 1 < len(boundaries)
            else last_content_idx
        )
        word_count = sum(
            len(body_child_text(children[j]).split())
            for j in range(start_idx, end_idx + 1)
            if children[j].tag != qn("w:sectPr")
        )
        episodes.append(
            {
                "number": boundary["number"],
                "title": boundary["title"],
                "heading": boundary["heading"],
                "start_idx": start_idx,
                "end_idx": end_idx,
                "word_count": word_count,
            }
        )

    total_words = sum(e["word_count"] for e in episodes)
    stats = {
        "total_episodes": len(episodes),
        "total_words": total_words,
        "avg_words": round(total_words / len(episodes)) if episodes else 0,
        "largest": max(episodes, key=lambda e: e["word_count"], default=None),
        "smallest": min(episodes, key=lambda e: e["word_count"], default=None),
    }
    return {"episodes": episodes, "stats": stats}


# ─────────────────────────────────────────────────────────────
# DOCX writing — trim a clone of the original package
# ─────────────────────────────────────────────────────────────
def build_docx(src_path: str, episodes: list[dict], ep_indices: list[int], out_path: str):
    new_doc = Document(src_path)
    body = new_doc.element.body
    children = list(body.iterchildren())
    first = episodes[ep_indices[0]]["start_idx"]
    last = episodes[ep_indices[-1]]["end_idx"]

    # Remove only whole body elements outside the selected range. Included XML
    # nodes and all related package parts remain untouched, preserving tables,
    # drawings, styles, hyperlinks, page breaks, headers, and footers.
    for i, child in enumerate(children):
        if child.tag == qn("w:sectPr"):
            continue
        if i < first or i > last:
            body.remove(child)
    new_doc.save(out_path)


# ─────────────────────────────────────────────────────────────
# Pydantic models
# ─────────────────────────────────────────────────────────────
class SplitRequest(BaseModel):
    job_id: str
    mode: str                                  # "n_files" | "per_file" | "custom"
    n_files: Optional[int] = None
    per_file: Optional[int] = None
    custom_ranges: Optional[list[str]] = None  # e.g. ["1-20", "21-40"]


# ─────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok", "version": "1.0.0"}


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    if not (file.filename or "").endswith(".docx"):
        raise HTTPException(400, "Only .docx files are accepted.")

    job_id = str(uuid.uuid4())
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir()
    doc_path = str(job_dir / "source.docx")

    with open(doc_path, "wb") as f:
        f.write(await file.read())

    try:
        result = analyse_docx(doc_path)
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(422, f"Could not parse document: {exc}")

    if not result["episodes"]:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(
            422,
            "No episodes found. Headings must match patterns like  'Ep 1 - Title'  or  'Chapter 1 - Title'.",
        )

    JOBS[job_id] = {
        "doc_path": doc_path,
        "job_dir": str(job_dir),
        "filename": file.filename,
        **result,
    }

    return {"job_id": job_id, "filename": file.filename, **result}


@app.post("/split")
def split(req: SplitRequest):
    job = JOBS.get(req.job_id)
    if not job:
        raise HTTPException(404, "Job not found — please upload again.")

    episodes = job["episodes"]
    n = len(episodes)
    doc_path = job["doc_path"]
    job_dir = Path(job["job_dir"])
    out_dir = job_dir / "output"
    out_dir.mkdir(exist_ok=True)

    # ── Compute groups ──────────────────────────────────────
    groups: list[tuple[int, int]] = []

    if req.mode == "n_files":
        k = max(1, req.n_files or 1)
        base, extra = divmod(n, k)
        idx = 0
        for i in range(k):
            size = base + (1 if i < extra else 0)
            if size:
                groups.append((idx, idx + size - 1))
                idx += size

    elif req.mode == "per_file":
        p = max(1, req.per_file or 1)
        for start in range(0, n, p):
            groups.append((start, min(start + p - 1, n - 1)))

    elif req.mode == "custom":
        if not req.custom_ranges:
            raise HTTPException(400, "custom_ranges required for custom mode.")
        ep_map = {e["number"]: i for i, e in enumerate(episodes)}
        for r in req.custom_ranges:
            parts = r.strip().split("-")
            if len(parts) != 2:
                raise HTTPException(400, f"Invalid range '{r}'. Use format '1-20'.")
            a, b = int(parts[0]), int(parts[1])
            if a not in ep_map or b not in ep_map:
                raise HTTPException(400, f"Episode numbers {a}–{b} not in document.")
            groups.append((ep_map[a], ep_map[b]))
    else:
        raise HTTPException(400, f"Unknown mode '{req.mode}'.")

    # ── Build DOCX files ─────────────────────────────────────
    output_files = []
    for g_start, g_end in groups:
        ep_s = episodes[g_start]["number"]
        ep_e = episodes[g_end]["number"]
        fname = f"Episodes_{ep_s:03d}_{ep_e:03d}.docx"
        out_path = str(out_dir / fname)
        build_docx(doc_path, episodes, list(range(g_start, g_end + 1)), out_path)
        output_files.append({"filename": fname, "ep_start": ep_s, "ep_end": ep_e, "path": out_path})

    # ── Pack ZIP ─────────────────────────────────────────────
    zip_path = str(job_dir / "Split_Document.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in output_files:
            zf.write(f["path"], f["filename"])

    JOBS[req.job_id]["output_files"] = output_files
    JOBS[req.job_id]["zip_path"] = zip_path

    return {
        "job_id": req.job_id,
        "files": [{"filename": f["filename"], "ep_start": f["ep_start"], "ep_end": f["ep_end"]} for f in output_files],
        "zip_filename": "Split_Document.zip",
    }


@app.get("/download/{job_id}/zip")
def download_zip(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    zp = job.get("zip_path")
    if not zp or not os.path.exists(zp):
        raise HTTPException(404, "ZIP not ready — call /split first.")
    return FileResponse(zp, media_type="application/zip", filename="Split_Document.zip")


@app.get("/download/{job_id}/file/{filename}")
def download_file(job_id: str, filename: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    match = next((f for f in job.get("output_files", []) if f["filename"] == filename), None)
    if not match or not os.path.exists(match["path"]):
        raise HTTPException(404, "File not found.")
    return FileResponse(
        match["path"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8765, reload=True, log_level="info")

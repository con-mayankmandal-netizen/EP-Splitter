# ✂ EpSplit — DOCX Episode Splitter

> Upload a multi-episode Word document → auto-detect all episodes → split into batches → download as ZIP.

![EpSplit UI](https://img.shields.io/badge/built%20with-FastAPI%20%2B%20Vanilla%20JS-7c6fff?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)
![Python](https://img.shields.io/badge/python-3.10%2B-blue?style=flat-square)

---

## What it does

You have a `.docx` file containing 100 episodes like this:

```
Ep 1 - A Stunning Recovery
[content...]

Ep 2 - Mistaken Identity
[content...]

Ep 3 - An Unhappy Reunion
[content...]
```

EpSplit automatically:
1. **Detects** all episode boundaries using configurable regex patterns
2. **Analyses** episode count, word counts, first/last episodes
3. **Splits** the document into multiple `.docx` files using your chosen batch size
4. **Downloads** everything as a single `Split_Document.zip`

---

## Quick start

### Option A — Browser only (no server needed)

Just open `frontend/index.html` in your browser. Processing happens entirely client-side using JSZip.

```bash
# Clone the repo
git clone https://github.com/yourusername/ep-splitter.git
cd ep-splitter

# Open in browser
open frontend/index.html
```

### Option B — With FastAPI backend (for large files / production)

```bash
cd backend
pip install -r requirements.txt
python main.py
```

Then open `http://localhost:8765` or point the frontend at your API:
```javascript
// In frontend/index.html, set:
const API_BASE = 'http://localhost:8765';
const USE_API = true;
```

---

## Detection patterns

EpSplit detects episode boundaries using these heading patterns (case-insensitive):

| Pattern | Example |
|---------|---------|
| `Ep N - Title` | `Ep 1 - A Stunning Recovery` |
| `Episode N - Title` | `Episode 25 - The Return` |
| `Chapter N - Title` | `Chapter 3 - Into the Dark` |
| `Part N - Title` | `Part 1 - Beginnings` |

### Adding custom patterns

In `backend/main.py`, extend `EPISODE_PATTERNS`:
```python
EPISODE_PATTERNS = [
    re.compile(r"^Ep\s+(\d+)\s*[-–—]\s*(.+)$", re.IGNORECASE),
    re.compile(r"^MyCustom\s+(\d+)\s*[-–—]?\s*(.*)$", re.IGNORECASE),  # ← add here
]
```

In `frontend/index.html`, extend `PATTERNS`:
```javascript
const PATTERNS = [
  /^Ep\s+(\d+)\s*[-–—]\s*(.+)$/i,
  /^MyCustom\s+(\d+)\s*[-–—]?\s*(.*)$/i,  // ← add here
];
```

---

## Splitting modes

### Mode A — Split into N files
Input: `5` → splits 100 episodes into **5 files of 20 each**

```
Episodes_001_020.docx  (Ep 1–20)
Episodes_021_040.docx  (Ep 21–40)
Episodes_041_060.docx  (Ep 41–60)
Episodes_061_080.docx  (Ep 61–80)
Episodes_081_100.docx  (Ep 81–100)
```

### Mode B — Episodes per file
Input: `20` → one file per 20 episodes (same result as above)

### Mode C — Custom ranges
Input:
```
1-15
16-30
31-50
51-100
```
→ 4 files with exactly those episode ranges

---

## Output naming

All output files use zero-padded episode numbers:
```
Episodes_001_020.docx
Episodes_021_040.docx
Split_Document.zip   ← contains all files
```

---

## API reference

When running the FastAPI backend:

### `GET /health`
Health check.

### `POST /upload`
```
Content-Type: multipart/form-data
Body: file=<docx>
```
Returns `{job_id, filename, episodes[], stats{}}`.

### `POST /split`
```json
{
  "job_id": "uuid",
  "mode": "n_files",
  "n_files": 5,
  "per_file": null,
  "custom_ranges": null
}
```
Returns `{job_id, files[], zip_filename}`.

### `GET /download/{jobId}/zip`
Download the ZIP archive.

### `GET /download/{jobId}/file/{filename}`
Download a single split DOCX.

---

## Folder structure

```
ep-splitter/
├── backend/
│   ├── main.py              # FastAPI server
│   └── requirements.txt
├── frontend/
│   ├── index.html           # Standalone single-file app
│   └── package.json         # Optional if using Vite
├── .github/
│   └── workflows/
│       └── deploy.yml       # GitHub Pages deploy
└── README.md
```

---

## Deploy to GitHub Pages

The frontend is a single HTML file — deploy it anywhere static.

Push to `main` and GitHub Actions will deploy automatically:

```yaml
# .github/workflows/deploy.yml (already included)
on:
  push:
    branches: [main]
```

Then enable GitHub Pages in your repo settings → Pages → Deploy from `gh-pages` branch.

---

## What formatting is preserved

| Feature | Status |
|---------|--------|
| Text content | ✅ Preserved |
| Paragraph structure | ✅ Preserved |
| Headings | ✅ Preserved |
| Bold / Italic / Underline | ✅ Preserved |
| Font size & name | ✅ Preserved |
| Text alignment | ✅ Preserved |
| Images (client-side) | ✅ Preserved (via XML clone) |

---

## License

MIT — free to use, modify, and distribute.

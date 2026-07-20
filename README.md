# doc2md

A headless CLI that converts **PowerPoint**, **Word** and **PDF** documents to Markdown using a vision LLM (any OpenAI-compatible endpoint).

## How it works

```
pptx / docx ──(MS Office COM)──► PDF ──(PyMuPDF)──► HD image per page ──(vision LLM)──► Markdown per page ──► single .md file
pdf ───────────────────────────►┘                                                             │
                                                     embedded figures ──► <name>_images/ ◄────┘
```

Each document page is rendered as a high-resolution image (300 DPI by default), then transcribed one page at a time by a vision model using a purpose-built prompt:

- **PowerPoint** — slide transcription prompt (HTML tables, charts, Gantt charts, diagrams, KPI cards).
- **PDF & Word** — document transcription prompt (HTML tables, LaTeX equations, headers/footers ignored, figure extraction).

## Requirements

- Windows with **Python 3.10+**
- **Microsoft Office** installed — required for `.ppt/.pptx/.doc/.docx` (COM automation). PDFs do **not** require Office.
- Access to an OpenAI-compatible LLM endpoint serving a **vision** model (e.g. OpenAI, LM Studio, Ollama, vLLM, or an internal gateway).

## Installation

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -e .
```

## Configuration

```powershell
.venv\Scripts\doc2md init-config      # generates config.yaml
```

Then edit `config.yaml`:

```yaml
llm:
  base_url: "http://localhost:1234/v1"   # OpenAI-compatible endpoint
  api_key: "${OPENAI_API_KEY}"           # or paste a key directly; may be empty
  model: "gpt-4o"                        # vision model
  api_mode: chat_completion              # chat_completion (default) | response
  # temperature: 0.1                     # optional — not sent at all when unset
  ssl_verify: true                       # default true; false for self-signed
  # headers:                             # optional custom headers
  #   X-Custom-Header: "value"
  timeout: 120
  max_retries: 3

render:
  dpi: 300
  image_format: png                      # png | jpeg

processing:
  concurrency: 3                         # pages transcribed concurrently
  skip_hidden_slides: true               # exclude hidden PowerPoint slides

output:
  dir: ./output
  page_marker: true                      # insert # Page N heading + <!-- page N --> between pages
  extract_images: true                   # extract figures to <name>_images/
```

`api_key` supports environment variable references: `"${VARIABLE_NAME}"`.

## Usage

```powershell
# Convert using ./config.yaml
doc2md convert slides.pptx

# Different config + selected pages
doc2md convert report.pdf -c prod.yaml --pages 1-3,7

# Override model / dpi / output without editing the config
doc2md convert doc.docx --model qwen2-vl-72b --dpi 200 -o out\doc.md
```

Output: one `.md` file per document in `output.dir`. Each page starts with a `# Page N` heading followed by a `<!-- page N -->` marker, making per-page content easy to extract.

## Figures and image references

When the transcription references a figure, doc2md extracts the embedded image from the (converted) PDF into a `<hash>_images/` folder next to the `.md` file and rewrites the reference to point at the real file. The folder name is the first 12 hex characters of the SHA-256 of the document name:

```
output/
├── paper.md
└── 3f4a9c1b72e0_images/
    ├── page-003-fig-1.png
    └── page-004-fig-1.jpg
```

- The vision model labels figures `page-N-fig-K` in reading order; extraction follows the same order, so references and files line up.
- Only embedded raster images above a small-size threshold are extracted (decorative icons and rules are skipped). Purely vector-drawn charts are not extracted yet.
- If the number of references and extracted figures on a page differ, a warning is printed so you can check the mapping.
- Disable with `extract_images: false` in the config.

### PowerPoint: embedded slide renders

Slides are mostly visual (charts, tables, shapes), so for `.ppt/.pptx` doc2md embeds the **full rendered slide image** directly above each slide's transcription instead of extracting individual figures:

```markdown
# Page 3
<!-- page 3 -->

![Slide 3](3f4a9c1b72e0_images/page-0003.png)

*... slide transcription ...*
```

This way the original slide stays visible alongside the transcribed Markdown.

## Behaviour notes

- **Hidden slides (PowerPoint)** — slides marked as hidden are skipped entirely: not rendered, not sent to the LLM, and not included in the Markdown. The console lists the skipped slides. Page markers keep the original slide numbers, so `# Page 5` / `<!-- page 5 -->` still means slide 5 even when earlier slides were skipped. If every requested slide is hidden, the conversion stops with an error (exit code `1`). Disable with `skip_hidden_slides: false`.
- **Retries** — failed requests (429/5xx/timeouts/network errors) are retried with exponential backoff up to `max_retries`.
- **Failed pages** — after retries are exhausted, a `[Page N transcription failed: ...]` placeholder is inserted into the output and remaining pages continue. Exit code `1` is returned so automation can detect partial failures.
- **Optional temperature** — when absent from the config, the `temperature` field is omitted from the payload entirely (the server uses its own default).
- **SSL verification** — controlled via `ssl_verify`; disable only for trusted internal endpoints.

## Prompts

Prompts live as plain text files in `doc2md/prompts/` (`ppt.txt`, `document.txt`) and can be edited to suit your needs. An editable install (`pip install -e .`) makes prompt changes take effect immediately.

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | Full success |
| 1 | Conversion/rendering failed, or one or more pages failed transcription |
| 2 | Config or CLI argument problem |

## Known limitations

- Windows only for Office conversion (COM automation). PDF conversion works on any platform supported by PyMuPDF.
- Output quality depends on the vision model's capability and the render resolution.
- Only embedded raster figures are extracted; vector-drawn charts/diagrams are skipped.
- Password-protected documents are not supported.

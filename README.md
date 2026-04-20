# Resume Rebuild Assistant

Extract structured data from any resume (PDF, Word, or scanned image), edit it in a guided form, and render polished output documents from templates.

Built for workforce development programs — AJCCs, INVEST, Hilton, and similar.

## Prerequisites

- **Python 3.11+**
- **Poppler** on PATH (for scanned PDF → image conversion)
- **Tesseract** on PATH (available but not directly used — Claude vision handles OCR)
- **ANTHROPIC_API_KEY** environment variable set
- **Microsoft Word** installed (only needed for PDF export via docx2pdf)

## Install

```bash
cd C:\Users\Brandon\projects\resume-rebuild-assistant
uv venv
uv pip install -r requirements.txt
```

## Setup

Generate the default Word template (one-time):

```bash
uv run python create_template.py
```

This creates `templates/default.docx` with Jinja2 tokens for docxtpl rendering.

## Run

```bash
uv run streamlit run app.py
```

Opens at http://localhost:8501.

## Usage Walkthrough

1. **Upload tab** — Drop a resume file (PDF, DOCX, PNG, JPG). Click "Extract Resume Data". Claude analyzes the content and returns structured JSON.
2. **Edit tab** — Review and modify every field: contact info, summary, skills (comma-separated tag editor), work history (add/remove/reorder jobs and bullets), education, certifications.
3. **Generate tab** — Pick a template from the dropdown, click "Generate Word Document" or "Generate PDF". Download the output.
4. **History tab** — All processed resumes are saved in SQLite. Click "Load" to reload any previous resume for re-editing.

## Creating a New Template

1. Copy `templates/default.docx` to `templates/my_template.docx`
2. Edit in Word — the template uses [docxtpl](https://docxtpl.readthedocs.io/) Jinja2 syntax:
   - `{{contact.name}}` — simple variable
   - `{{contact_line}}` — pre-joined email | phone | location
   - `{{skills_text}}` — skills joined by bullet dots
   - `{%p for job in work_history %}...{%p endfor %}` — paragraph-level loops
   - `{%p if has_summary %}...{%p endif %}` — conditional sections
3. Available context variables:
   - `contact.name`, `contact.email`, `contact.phone`, `contact.location`, `contact.linkedin`
   - `summary`, `skills_text`
   - `work_history[].role`, `.employer`, `.location`, `.start_date`, `.end_date`, `.bullets[]`
   - `education[].institution`, `.degree`, `.field`, `.graduation_year`, `.honors`
   - `certifications[].name`, `.issuer`, `.date`
   - Boolean flags: `has_summary`, `has_skills`, `has_work`, `has_education`, `has_certs`
4. Save and it appears in the Generate tab dropdown automatically.

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Claude API error" | Check ANTHROPIC_API_KEY is set. Verify at https://console.anthropic.com. |
| "Poppler not found" | Install Poppler and add its `bin/` to system PATH. Restart terminal. |
| Scanned PDF returns garbled text | The pypdf text threshold (50 chars) may need tuning — edit `ingestion.py` line with `len(full_text) > 50`. |
| PDF generation fails | Requires Microsoft Word on Windows. Use "Generate Word" instead and export from Word manually. |
| Template rendering errors | Check template syntax — `{%p %}` tags must be the sole content in their paragraph. |

## Project Structure

```
resume-rebuild-assistant/
├── app.py              — Streamlit UI (tabs: upload, edit, generate, history)
├── ingestion.py        — File loading: DOCX text, PDF text/scan, images
├── extraction.py       — Claude API call for structured JSON extraction
├── rendering.py        — docxtpl Word rendering + docx2pdf PDF conversion
├── db.py               — SQLite persistence (auto-creates resume_rebuild.db)
├── config.py           — All tuneable settings in one place
├── create_template.py  — Generates the default.docx template
├── prompts/
│   └── extraction.txt  — Claude extraction prompt (editable without code changes)
├── templates/
│   └── default.docx    — Clean single-column ATS-friendly template
├── output/             — Generated documents land here
├── logs/
│   └── api_calls.log   — Every Claude call logged with cost estimate
└── requirements.txt
```

## API Cost

Each extraction uses one Claude API call (two if the first response isn't valid JSON). Estimated cost per resume: ~$0.01–0.05 for text, ~$0.05–0.15 for scanned images.

Costs are logged to `logs/api_calls.log` for tracking.

---
name: cv-to-form
description: >
  Convert a candidate CV (PDF) into a formatted Evotek company form (.docx).
  Use this skill whenever you receive a CV PDF and need to produce a standardized
  Evotek intake form — even if the user doesn't mention "Evotek" explicitly.
  Triggers on: "convert CV", "CV to form", "fill in Evotek form", "candidate form",
  "resume to docx", or any request to transform a resume PDF into a company template.
---

# Evotek CV → Form Converter

Convert a candidate's CV (PDF) into the standardized Evotek intake form (.docx).

## Prerequisites

The agent needs Python 3 and `pdfplumber` + `python-docx` installed:

```bash
pip install pdfplumber python-docx
```

## Skill Location

All bundled files live relative to this SKILL.md:

| File | Purpose |
|------|---------|
| `scripts/extract_cv_text.py` | Extracts raw text from a CV PDF |
| `scripts/generate_form.py` | Fills the Evotek template with candidate data |
| `assets/Evotek_cv_form_template.docx` | The blank form template |

Resolve these paths from the skill directory — never assume they're in the CWD.

---

## Workflow

### Step 1 — Ask for output language

Ask the user:

> *"What language should the output CV form be written in? (e.g. Vietnamese, English, or both)"*

Use this language throughout the extracted fields.

### Step 2 — Extract text from the CV

```bash
python <skill-dir>/scripts/extract_cv_text.py <path-to-pdf>
```

Read the output carefully. This is the raw data source for all fields.

### Step 3 — Edit the DATA SECTION in `generate_form.py`

Open `<skill-dir>/scripts/generate_form.py` and fill in every variable in the **DATA SECTION** at the top. The fields are:

| Variable | Format | Notes |
|----------|--------|-------|
| `CANDIDATE_SLUG` | PascalCase | e.g. `PhanTienDat` |
| `OUTPUT_LANG` | string | The language chosen in Step 1 |
| `NAME` | string | Name only — no phone, email, address, or DOB |
| `ROLE` | string | Job title |
| `ACHI` | `- ` bullets | Overview / key achievements, one per line starting with `- ` |
| `EDU1_YEAR` / `EDU1_DETAIL` | strings | First education entry |
| `EDU2_YEAR` / `EDU2_DETAIL` | strings | Second entry, or `"No info"` |
| `ENG_MARK` | string | English score, or `"No info"` |
| `TECH_OS`, `TECH_DB`, `TECH_PROG`, `TECH_DEVTOOL`, `TECH_METHOD` | strings | Skills per category, or `"No info"` |
| `PROJECTS` | list of dicts | One dict per project with keys: `name`, `duration`, `pos`, `size`, `des`, `res`, `tech` |

**Rules for filling data:**

- **Every project** from the CV must appear in `PROJECTS`. Add as many entries as needed.
- Use `- ` (hyphen + space) for all list bullets — never use `•` or other symbols.
- Set missing/unextractable fields to `"No info"`.
- Write descriptions and responsibilities **in detail** — do not summarize or abbreviate.
- Write all content in the output language chosen in Step 1.

**PROJECTS dict format:**

```python
{
    'name': 'Project Name',
    'duration': '01/2023 - 12/2023',
    'pos': 'Your Role',
    'size': 'Team size or "No info"',
    'des': 'One or two sentence description of the project.',
    'res': '- Responsibility 1\n- Responsibility 2\n- Responsibility 3',
    'tech': 'Technology 1, Technology 2',
}
```

### Step 4 — Run the form generator

```bash
python <skill-dir>/scripts/generate_form.py <path-to-pdf>
```

The `.docx` output is saved **in the same directory** as the input PDF, named `Evotek_cv_form_<SLUG>.docx`.

### Step 5 — Verify the output

Run a quick verification check:

```python
from docx import Document
import re

doc = Document('<output-path>')
remaining = []
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for m in re.finditer(r'\[\[.*?\]\]', cell.text):
                remaining.append(m.group())

print(f'Tables: {len(doc.tables)}')
if remaining:
    print(f'UNREPLACED: {set(remaining)}')
else:
    print('All placeholders replaced.')
```

Confirm all of the following:

- No `[[...]]` placeholders remain in the document
- All list items use `- ` hyphens
- The Name cell contains name only (no contact info)
- Missing fields show `"No info"`
- Content is detailed, not summarized
- All projects from the CV are present as separate tables
- The output file is in the same folder as the input PDF

---

## Common Pitfalls

- **Forgetting projects**: The template has 3 built-in project tables. If the CV has more than 3 projects, the script automatically clones additional tables. Always include every project.
- **Bullet symbols**: Using `•`, `●`, or `–` instead of `- ` will break the template formatting. The script auto-converts some, but be consistent from the start.
- **Template path**: The script resolves the template relative to its own directory (`assets/`), so it works from any CWD. Do not hardcode an absolute path.
- **Encoding**: Project descriptions with special characters (Vietnamese diacritics, etc.) work fine — `python-docx` handles Unicode natively.

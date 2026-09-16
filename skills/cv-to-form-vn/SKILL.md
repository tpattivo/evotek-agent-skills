---
name: cv-to-form-vn
description: >
  Convert a candidate CV (PDF) into a formatted Evotek company form (.docx)
  with Vietnamese headings (TÊN, VỊ TRÍ, TỔNG QUAN, DỰ ÁN, etc.).
  Use this skill whenever you receive a CV PDF and need to produce a
  standardized Evotek intake form with Vietnamese section labels — even if
  the user doesn't mention "Evotek" explicitly.
  Triggers on: "convert CV Vietnamese", "CV to form Vietnamese", "fill in
  Evotek form Vietnamese", "简历转表单", "cv-to-form-vn", or any request to
  transform a resume PDF into the Vietnamese version of the Evotek template.
---

# Evotek CV → Form Converter (Vietnamese Headings)

Convert a candidate's CV (PDF) into the standardized Evotek intake form (.docx)
with **Vietnamese section headings**.

## Template headings

| English | Vietnamese |
|---------|-----------|
| NAME | TÊN |
| ROLE | VỊ TRÍ |
| OVERVIEW | TỔNG QUAN |
| EDUCATION AND TRAINING | GIÁO DỤC VÀ ĐÀO TẠO |
| LANGUAGES | NGÔN NGỮ |
| TECHNICAL SKILLS | CHUYÊN MÔN KỸ THUẬT |
| RECENT ASSIGNMENTS | DỰ ÁN |

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
| `scripts/generate_form.py` | Fills the Vietnamese Evotek template with candidate data |
| `assets/Evotek_cv_form_template.docx` | The blank Vietnamese form template |

Resolve these paths from the skill directory — never assume they're in the CWD.

---

## Workflow

### Step 1 — Ask for output language

Ask the user:

> *"What language should the output CV form content be written in? (e.g. Vietnamese, English, or both)"*

Note: The template headings are always Vietnamese. This step determines the
language for the **content** (candidate data, descriptions, responsibilities).

### Step 2 — Extract text from the CV

```bash
python <skill-dir>/scripts/extract_cv_text.py <path-to-pdf>
```

Read the output carefully. This is the raw data source for all fields.

### Step 3 — Edit the DATA SECTION in `generate_form.py`

Open `<skill-dir>/scripts/generate_form.py` and fill in every variable in the
**DATA SECTION** at the top. The fields are:

| Variable | Format | Notes |
|----------|--------|-------|
| `CANDIDATE_SLUG` | PascalCase | e.g. `PhanTienDat` |
| `OUTPUT_LANG` | string | The language chosen in Step 1 |
| `NAME` | string | Name only — no phone, email, address, or DOB |
| `ROLE` | string | Job title |
| `ACHI` | `- ` bullets | Tổng quan — overview / key achievements, one per line starting with `- ` |
| `EDU1_YEAR` / `EDU1_DETAIL` | strings | First education entry |
| `EDU2_YEAR` / `EDU2_DETAIL` | strings | Second entry, or `"No info"` |
| `ENG_MARK` | string | English score, or `"No info"` |
| `TECH_OS`, `TECH_DB`, `TECH_PROG`, `TECH_DEVTOOL`, `TECH_METHOD` | strings | Chuyên môn kỹ thuật — skills per category, or `"No info"` |
| `PROJECTS` | list of dicts | Dự án — one dict per project with keys: `name`, `duration`, `pos`, `size`, `des`, `res`, `tech` |

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

> **Batch note:** when processing multiple CVs in parallel, skip this step —
> write a JSON data file instead (see **Batch Processing** below).

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
- The TÊN cell contains name only (no contact info)
- Missing fields show `"No info"`
- Content is detailed, not summarized
- All projects from the CV are present as separate tables
- Section headings are in Vietnamese (TÊN, VỊ TRÍ, TỔNG QUAN, DỰ ÁN, etc.)
- The output file is in the same folder as the input PDF

---

## Batch Processing — Multiple CVs in Parallel (Sub-Agents)

To process several CVs at once, fan out one sub-agent per CV. Parallel is safe
because each worker writes its **own** JSON data file instead of editing the
shared `generate_form.py`.

### Output language

Ask the Step 1 question a single time, then pass the chosen language to every
sub-agent prompt.

### Launching the sub-agents

Each prompt must be self-contained (sub-agents don't inherit this
conversation). Fill in absolute paths:

```text
Convert the CV at <abs-path-to-cv.pdf> into the Evotek intake form (Vietnamese headings).
Output language: <chosen-language>.

1. Extract the CV text:
   python <skill-dir>/scripts/extract_cv_text.py <abs-path-to-cv.pdf>
2. Write the candidate data to <abs-path-to-data.json> (schema below).
   Never edit generate_form.py.
3. Generate the form:
   python <skill-dir>/scripts/generate_form.py <abs-path-to-cv.pdf> <abs-path-to-data.json>
4. Verify the output .docx with the Step 5 checklist.

Report the output file path and the verification result.
```

### JSON data file

Keys are the DATA SECTION variable names. Omit any key to fall back to the
defaults in `generate_form.py`. The Step 3 filling rules still apply (every
project in `PROJECTS`, `- ` bullets, `"No info"` for missing fields, detailed
descriptions, all content in the chosen output language).

```json
{
  "CANDIDATE_SLUG": "PhanTienDat",
  "OUTPUT_LANG": "Vietnamese",
  "NAME": "Phan Tien Dat",
  "ROLE": "Fullstack Developer",
  "ACHI": "- Achievement 1\n- Achievement 2\n- Achievement 3",
  "EDU1_YEAR": "2020 - 2024",
  "EDU1_DETAIL": "B.Eng, HCMUT\nGPA: 3.2",
  "EDU2_YEAR": "No info",
  "EDU2_DETAIL": "No info",
  "ENG_MARK": "IELTS 7.0",
  "TECH_OS": "Windows, Linux",
  "TECH_DB": "MySQL, PostgreSQL",
  "TECH_PROG": "Python, JavaScript, C#",
  "TECH_DEVTOOL": "Git, Docker, VS Code",
  "TECH_METHOD": "Agile/Scrum",
  "PROJECTS": [
    {
      "name": "Project Name",
      "duration": "01/2023 - 12/2023",
      "pos": "Role",
      "size": "No info",
      "des": "Project description.",
      "res": "- Responsibility 1\n- Responsibility 2",
      "tech": "Tech 1, Tech 2"
    }
  ]
}
```

### Why parallel is safe

| Piece | Shared state? |
|-------|---------------|
| `extract_cv_text.py` | no — read-only |
| `<slug>.json` per CV | no — one file per sub-agent |
| output `.docx` | no — slug-named, written next to its PDF |
| verification | no — read-only |

### Batch pitfalls

- **Never** let two sub-agents edit `generate_form.py` — the DATA SECTION is a
  single shared file; concurrent edits race and the last writer wins.
- Keep slugs unique — two CVs with the same slug overwrite each other's output.
- The JSON file is never modified by the script, so re-running the generator
  is idempotent.

## Common Pitfalls

- **Forgetting projects**: The template has 3 built-in project tables. If the CV has more than 3 projects, the script automatically clones additional tables. Always include every project.
- **Bullet symbols**: Using `•`, `●`, or `–` instead of `- ` will break the template formatting. The script auto-converts some, but be consistent from the start.
- **Template path**: The script resolves the template relative to its own directory (`assets/`), so it works from any CWD. Do not hardcode an absolute path.
- **Encoding**: Project descriptions with special characters (Vietnamese diacritics, etc.) work fine — `python-docx` handles Unicode natively.
- **Template difference**: This skill uses the Vietnamese-heading template. Do not mix with the English `cv-to-form` skill's template.

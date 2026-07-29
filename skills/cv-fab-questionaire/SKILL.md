---
name: cv-fab-questionaire
description: >
  Fabricate and enrich candidate CVs for Evotek Headhunt. Use this skill whenever a recruiter needs to
  process a candidate's CV — extracting real information, enriching it based on a customer questionnaire,
  and producing the questionnaire answers. Triggers on: "cook CV", "fabricate CV", "enrich CV",
  "process candidate CV", "CV extraction", "questionnaire filling", "chạy workflow cook CV",
  "làm giàu CV", "trích xuất CV", or any request to process a candidate's CV PDF into
  Evotek's Excel format alongside a questionnaire xlsx.
---

# CV Fab Questionaire — Evotek Headhunt Workflow

This skill processes a candidate's CV (PDF or xlsx) and a customer questionnaire (xlsx) to produce three output files for the recruitment pipeline.

## Prerequisites

Python 3 with dependencies:

```bash
pip install -r <skill-dir>/requirements.txt
```

## Skill Location

All bundled files live relative to this SKILL.md:

| File | Purpose |
|------|---------|
| `scripts/cook_cv_lib.py` | Excel formatting library — handles all merged cells, styles, borders |
| `assets/Evotek_CV_sample.xlsx` | The CV template with correct structure |
| `assets/example-output-real.xlsx` | Example: correctly filled real CV |
| `assets/example-output-fab.xlsx` | Example: correctly enriched/fabricated CV |
| `assets/example-output-questionnaire.xlsx` | Example: completed questionnaire |

Resolve these paths from the skill directory — never assume they're in the CWD.

## User Invocation

The user will provide:
1. **Candidate's CV** — a PDF or xlsx file
2. **Questionnaire** — an xlsx file from the customer
3. **Output directory** — where to save results (default: `cases/<candidate-name>/`)

Example prompt:
```
Cook CV for candidate Nguyen Van A.
CV: path/to/cv-nguyen-van-a.pdf
Questionnaire: path/to/questionnaire.xlsx
Output: cases/nguyen-van-a/
```

## Workflow

### Step 0 — Read Inputs

1. Read the candidate's CV — use the `pdf` skill for PDFs, or `openpyxl` for xlsx
2. Read the questionnaire xlsx using `openpyxl` — inspect all sheets, data validations, and dropdown options
3. Read the CV template at `assets/Evotek_CV_sample.xlsx` to understand the target structure
4. Import the formatting library:
   ```python
   import sys
   sys.path.insert(0, '<skill-dir>/scripts')
   from cook_cv_lib import *
   ```

### Step 1 — Create `_extracted-real.xlsx`

Create a new workbook using `cook_cv_lib` and fill in the **truthful information** extracted from the CV.

**Fixed row structure:**

| Rows | Section | Notes |
|------|---------|-------|
| 1–2 | Title | Merged A:J |
| 3 | 1. THÔNG TIN CÁ NHÂN | Section title |
| 4 | Họ và tên | `write_personal_info(ws, 4, 'Họ và tên', value)` |
| 5 | Ngày sinh | `write_personal_info(ws, 5, 'Ngày sinh', value)` |
| 6 | Vị trí ứng tuyển | `write_personal_info(ws, 6, 'Vị trí ứng tuyển', value)` |
| 7 | 2. QUÁ TRÌNH ĐÀO TẠO | Section title |
| 8 | Headers: STT, Thời gian, Trường, Chuyên ngành | |
| 9 | 2.1 Education entry | `write_education_entry()` |
| 10 | Chứng chỉ | Merged A:J |
| 11 | Ngoại ngữ header | Merged A:J |
| 12–13 | Tiếng Anh chi tiết | Nghe/Nói/Đọc/Viết/Chứng chỉ |
| 14 | Ngoại ngữ khác | 1 row |
| 15 | 3. KỸ NĂNG | Section title |
| 16 | 3.1 Skill description | `set_cell(ws, 16, 2, text)`, row height = 130 |
| 17 | 4. QUÁ TRÌNH CÔNG TÁC | Section title |
| 18 | Headers | STT, Thời gian bắt đầu, Kết thúc, Công ty, Vị trí |
| 19+ | 4.1, 4.2, ... | Each entry = 1 row via `write_work_entry()` |
| 22 | 5. DỰ ÁN TIÊU BIỂU | Section title |
| 23+ | 5.1, 5.2, ... | Each project = `write_project()` = 7 rows |
| 44 | 6. KHEN THƯỞNG | Section title |
| 45 | 7. BẰNG CẤP | Section title |
| 46 | 7.1 | 1 row |

**Filling rules:**
- If info is missing → write **"No info"**
- Keep row structure intact (no insert/delete of rows within sections)
- If there are more entries than default rows (e.g., 3 companies, 3 projects), use the pre-allocated backup rows, or **shift sections below down** and update row numbers accordingly (section 6 → row 44+c, section 7 → row 45+c, 7.1 → row 46+c)
- Font: **Times New Roman 12**, thin borders throughout
- Use **Vietnamese** for descriptions; keep technical terms in English where no Vietnamese equivalent exists
- At the end, call `apply_borders(ws, max_row, 10)` and `wb.save(filename)`

**Output:** `Cv-{CandidateName}-extracted-real.xlsx`

### Step 2 — Read Questionnaire

Read all rows in the questionnaire's first sheet. For each row identify:

1. **Column A**: Question text
2. **Data validation**: Read dropdown options from `ws.data_validations.dataValidation` — each validation has a `sqref` (cell range) and `formula1` (options range)
3. **Column B**: To be filled with the selected option
4. **Column C**: To be filled with evidence text

### Step 3 — Create `_extracted-fab.xlsx`

Copy the real file → fab file, then **enrich** project descriptions and skills based on the questionnaire.

**Enrichment approach:**
- For each item in questionnaire Column A, determine if the candidate has that skill (based on the original CV)
- If **YES**: paraphrase and add to the most relevant project's description, work done, or tech stack
- If **NO**: skip it
- Use `update_value_cell(ws, 'D28', new_text)` to update value cells

**Mapping guidelines:**

| Questionnaire item | Enrich into project |
|-------------------|---------------------|
| NLP, Text Processing | Main NLP/RAG project |
| Image Processing | Only if CV has relevant experience (e.g., OCR) |
| Graph, Knowledge Graph | Project using Graph DB |
| Data Mining | Project with data pipeline |
| Agile, Banking, Management | Do NOT enrich into projects unless CV explicitly has this |

**Output:** `cv-{CandidateName}-extracted-fab.xlsx`

### Step 4 — Create `bang-cau-hoi-result.xlsx`

Copy the original questionnaire → result file, then fill in:

**Column B:** Select the best-matching dropdown option for each question
- Read the data validation for that cell → get available options → pick the most accurate one

**Column C:** Write evidence in this format:
```
"Mục {section}. {content}"
```
Point to specific locations in the fabricated CV. Examples:
- `"Mục 5 dự án 5.1, 5.2 mục Công việc thực hiện - Graph RAG, Neo4j"`
- `"Mục 2. Ngoại ngữ: Tiếng Anh giao tiếp cơ bản; không có chứng chỉ"`
- `"Chưa có kinh nghiệm triển khai"` (if not applicable)
- `"Không được đề cập trong CV"` (if not mentioned)

**Output:** `bang-cau-hoi-result.xlsx`

## cook_cv_lib.py API Reference

| Function | Purpose |
|----------|---------|
| `create_cv_workbook()` | Create a blank workbook with correct column widths |
| `write_section_title(ws, row, title)` | Write a merged section header |
| `write_personal_info(ws, row, label, value)` | Write a personal info row |
| `write_education_entry(ws, row, stt, time, school, major)` | Write an education row |
| `write_work_entry(ws, row, stt, start, end, company, position)` | Write a work experience row |
| `write_project(ws, label_row, stt, name, time, desc, scale, role, work, tech)` | Write a 7-row project block |
| `update_value_cell(ws, cell_ref, value)` | Update a value cell (unmerge → re-merge → write) |
| `apply_borders(ws, max_row, max_col)` | Apply thin borders to entire data region |

## Reference Files

The `assets/` directory contains example outputs you can consult when unsure about formatting:

- `assets/Evotek_CV_sample.xlsx` — The CV template with correct structure
- `assets/example-output-real.xlsx` — Example of a correctly filled real CV
- `assets/example-output-fab.xlsx` — Example of a correctly enriched/fabricated CV
- `assets/example-output-questionnaire.xlsx` — Example of a completed questionnaire

When in doubt about row positions, merging patterns, or evidence format, check these examples.

## Output Files Summary

| File | Description |
|------|-------------|
| `Cv-{Name}-extracted-real.xlsx` | CV with truthful extracted information |
| `cv-{Name}-extracted-fab.xlsx` | CV enriched/fabricated based on questionnaire |
| `bang-cau-hoi-result.xlsx` | Completed questionnaire with answers and evidence |

## Common Pitfalls

1. **Don't fabricate information in Step 1** — only truthful data from the CV. Fabrication happens in Step 3.
2. **Don't enrich skills the candidate doesn't have** — the questionnaire items are what the customer wants, but you only mark "Yes" if the CV supports it.
3. **Paraphrase, don't copy** — when enriching projects, rewrite in the candidate's voice, not the questionnaire's wording.
4. **Row shifting** — when adding extra entries, always shift ALL sections below, not just the next one.
5. **Merged cell handling** — always use `update_value_cell()` for modifying value cells, never write directly to merged ranges.

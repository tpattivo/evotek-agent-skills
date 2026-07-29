# Evotek Agent Skills

A collection of agent skills for automating Evotek workflows.

## Skills

| Skill | Description |
|-------|-------------|
| [`cv-to-form`](skills/cv-to-form/) | Convert a candidate CV (PDF) into the Evotek intake form (.docx) |

## Installation

Each skill is self-contained under `skills/<name>/`. To use a skill:

1. Copy the skill folder (e.g. `skills/cv-to-form/`) into your agent's skills directory
2. Install Python dependencies: `pip install -r <skill>/requirements.txt`
3. The skill auto-triggers when relevant — see each skill's `SKILL.md` for trigger phrases

## Adding a new skill

Create a folder under `skills/` following this structure:

```
skills/
  my-skill/
    ├── SKILL.md          # Instructions + YAML frontmatter (name, description)
    ├── requirements.txt  # Python deps (if any)
    ├── scripts/          # Executable helpers (optional)
    └── assets/           # Templates, fonts, etc. (optional)
```

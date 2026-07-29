# Evotek Agent Skills

This repo contains reusable agent skills for Evotek workflows.

## Available Skills

| Skill | Location | What it does |
|-------|----------|-------------|
| `cv-to-form` | `skills/cv-to-form/` | Converts a candidate CV (PDF) into the Evotek intake form (.docx) |
| `cv-fab-questionaire` | `skills/cv-fab-questionaire/` | Extracts real CV info, enriches based on questionnaire, fills questionnaire answers |

## How to use

Each skill is a self-contained folder under `skills/`. To install a skill in another repo or agent:

1. Copy the skill folder into your agent's skills directory
2. Run `pip install -r <skill-folder>/requirements.txt`
3. The agent will auto-trigger the skill when relevant

## Adding new skills

1. Create a new folder under `skills/<kebab-case-name>/`
2. Add a `SKILL.md` with YAML frontmatter (`name`, `description`)
3. Add `scripts/`, `assets/`, `requirements.txt` as needed
4. Update this file and `README.md`

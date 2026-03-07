# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This project is a collection of custom Claude Code skills (slash commands) and configurations. There is no build system, test runner, or package manager — the repository is pure Markdown.

## Skills System

Skills live in `.claude/skills/`. Each `.md` file becomes a `/<filename>` slash command, invoked via the Skill tool.

**Skill file format:**

```markdown
---
name: skill-name
description: "Trigger description — Claude uses this to auto-detect when to run the skill"
---

Skill prompt body here.
```

The `description` frontmatter controls when Claude auto-triggers the skill. Keep it specific so it fires on the right user phrases.

## Existing Skills

### `/weekly-task-allocation`
Processes a Microsoft Planner Excel/CSV export and generates a styled HTML email for weekly team task allocation. Auto-triggers on phrases like "do the weekly email" or "process the planner file".

- Filters rows to `Urgent` priority or `ThisWeek` label only
- Organises tasks per team member (Raymond, Jasmine, Victor, Edward, Mark, Guy) using label matching
- Outputs a self-contained `.html` file ready to paste into Outlook
- Sign-off is always from **Guy Baranyay**; year is **2026**

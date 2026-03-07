# Skills

Custom slash commands for this project. Each `.md` file defines a skill invocable via `/<skill-name>`.

## Structure

```
.claude/skills/
  <skill-name>.md   # Defines the /<skill-name> slash command
```

## Creating a Skill

Each skill file should contain a prompt that Claude will execute when the command is invoked.
Optionally include frontmatter to configure behavior.

Example: `.claude/skills/my-skill.md`

```markdown
---
description: "Short description of what this skill does"
---

Your skill prompt here. Describe what Claude should do when this command is run.
```


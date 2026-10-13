---
title: Coding agents and llms.txt
description: Give Claude Code, Cursor or any other coding agent what it needs to embed the editor and write schemas and hooks.
sidebar:
  order: 1
---

These docs come in forms coding agents read well, and with a skill that teaches an agent how to use the editor.

## llms.txt

| File | What it is |
| --- | --- |
| [`/llms.txt`](/llms.txt) | An index of every page, each with a one-line summary, following the [llms.txt](https://llmstxt.org) convention. |
| [`/llms-full.txt`](/llms-full.txt) | Every page in one Markdown file, the API reference included. |
| Any page with `.md` | That page as Markdown: [`/schemas/async-hooks.md`](/schemas/async-hooks.md). |

Point an agent at `llms-full.txt` when it needs the whole picture, or at single pages to keep its context small.

## The skill

The skill is a short, dense guide to embedding the editor and writing schemas, hooks and save handlers, with the details agents most often get wrong: hooks aren't called for empty cells, percentages are fractions, the editor needs a height, and hooks don't survive JSON.

### Claude Code

Add it to a project, where everyone working on it gets it:

```sh
mkdir -p .claude/skills/workbook-editor
curl -fsSL https://workbook-editor-docs.pages.dev/agents/workbook-editor/SKILL.md -o .claude/skills/workbook-editor/SKILL.md
```

Or to `~/.claude/skills/workbook-editor/` for every project of yours. Claude Code uses it whenever a task involves the editor.

### Cursor

```sh
mkdir -p .cursor/rules
curl -fsSL https://workbook-editor-docs.pages.dev/agents/workbook-editor.mdc -o .cursor/rules/workbook-editor.mdc
```

### Other agents

Most agents read an `AGENTS.md` at the project's root. Add:

```md
## Workbook Editor
This project embeds the Workbook Editor (@sitepen/spreadsheet). Before changing its schemas, hooks or options,
read https://workbook-editor-docs.pages.dev/agents/workbook-editor/SKILL.md, and https://workbook-editor-docs.pages.dev/llms-full.txt for detail.
```

---
title: Coding agents and llms.txt
description: Give Claude Code, Cursor or any other coding agent what it needs to embed the editor, write schemas and hooks, and write plugins.
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

Point an agent at `llms-full.txt` when it needs the whole picture, or at single pages to keep its context small. For a plugin, [`/plugins/writing.md`](/plugins/writing.md) has the whole contract and a complete plugin to start from.

## The skill

The skill is a short, dense guide to embedding the editor and writing schemas, hooks, save handlers and plugins, with the details agents most often get wrong: hooks aren't called for empty cells, percentages are fractions, the editor needs a height, hooks don't survive JSON, and a plugin must open its popups inside the click and import only types from `@ruled/sheet/plugin`.

### Claude Code

Add it to a project, where everyone working on it gets it:

```sh
mkdir -p .claude/skills/ruled
curl -fsSL https://ruled.page/agents/ruled/SKILL.md -o .claude/skills/ruled/SKILL.md
```

Or to `~/.claude/skills/ruled/` for every project of yours. Claude Code uses it whenever a task involves the editor.

### Cursor

```sh
mkdir -p .cursor/rules
curl -fsSL https://ruled.page/agents/ruled.mdc -o .cursor/rules/ruled.mdc
```

### Other agents

Most agents read an `AGENTS.md` at the project's root. Add:

```md
## Ruled
This project embeds Ruled (@ruled/sheet), a spreadsheet editor. Before changing its schemas, hooks, options or plugins,
read https://ruled.page/agents/ruled/SKILL.md, and https://ruled.page/llms-full.txt for detail.
```

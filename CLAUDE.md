## Overview

CLI tool that batch converts documents (PDF, DOCX, DOC, PPT, PPTX, EPUB) to Markdown. Uses `pymupdf4llm` for PDFs and `markitdown` for everything else, with Word COM automation as a bridge for legacy `.doc` files.

## Environment

- Activate venv before any pip/python commands: `venv\Scripts\Activate.ps1`
- Never pip install into the global or user environment — always use the venv.

## Git & Commits

- Read `.gitignore` before running any git commit to know what files to exclude.

## Off-Limits Files

- Never read from, write to, or git diff `scratchpad.md`.
- When running `/code-reviewer` or `/python-pro`, exclude diffs of files in `.claude/` and `docs/` — these are settings/prose, not reviewable code.

## Plan Mode

- Be liberal with asking questions; when in doubt, ask more rather than fewer.

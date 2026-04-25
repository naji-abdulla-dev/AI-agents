# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

General-purpose research and task execution agent. The task is defined in `WORK.md`; the role and approach are defined in `INSTRUCTIONS.md`.

## Workflow

1. Read `WORK.md` to understand the specific task.
2. Read `INSTRUCTIONS.md` for the role, approach, and output format.
3. Research using WebSearch and WebFetch for current information.
4. Produce the output in the format specified in `WORK.md` (default: markdown).
5. Save the output file to `output/`.
6. Read back the saved file and verify it matches the requirements before finishing.

## Output Conventions

- Final output files → `output/` folder
- Intermediate Python scripts → `tmp/` folder
- Default output format is Markdown (`.md`) unless `WORK.md` specifies otherwise (e.g. `.docx`, `.xlsx`, `.pdf`)
- After saving, always read the file back to verify correctness

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

System design document generator. Given a system or problem in `WORK.md`, produce a professional `.docx` document covering the full design from requirements through deep-dive component specs.

## Workflow

1. Read `WORK.md` for the system to design (e.g. "Design Twitter Feed", "Design a Chat System").
2. Research real-world implementations and reference architectures using WebSearch/WebFetch.
3. Write a Python script using `python-docx` to generate the Word document.
4. Save the script to `tmp/generate_<system>_design.py`, run it, output goes to `output/`.
5. Read back the `.docx` and verify all sections, diagrams, and tables are clean and readable.

## Output Structure (per document)

Each `.docx` should cover, in order:
1. Problem statement — framed in an interesting and engaging way
2. Functional and non-functional requirements
3. Core components — high-level list with responsibility of each
4. Data model — schema, key entities, relationships (table format where applicable)
5. Data flow — request lifecycle, event flow
6. High-level architecture — overall diagram (ASCII or described), component interactions
7. Deep dives — detailed design specs for each major component; use diagrams and data sheets
8. Real-world examples — how companies like Google, Meta, Netflix solve similar problems
9. References — blog posts, YouTube videos, papers

## Output Conventions

- Final `.docx` → `output/`
- Intermediate Python scripts → `tmp/`
- Diagrams: ASCII art embedded in the document, or clearly described with labels
- Tables must have wrapped text and readable column widths
- After saving, read back the file and iterate until content is professional and clear

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

Coding interview study guide generator. Given a data structures & algorithms topic in `WORK.md`, produce a comprehensive `.docx` study guide covering concepts, algorithms, patterns, and interview-ready code.

## Workflow

1. Read `WORK.md` for the topic (e.g. "Graph Algorithms", "Dynamic Programming").
2. Research the topic thoroughly — algorithms, complexity, common interview patterns.
3. Write a Python script using `python-docx` to generate the Word document.
4. Save the script to `tmp/<topic>_guide.py`, run it, output goes to `output/`.
5. Read back the generated `.docx` and verify all sections are present and readable.

## Output Structure (per guide)

Each `.docx` should cover, in order:
1. Concept overview — what it is and why it matters in interviews
2. Key algorithms — with time/space complexity for each
3. Common patterns — when to recognise and apply each
4. Typical interview questions and strategies
5. Code implementations — Python, clean and commented, both brute-force and optimised
6. Diagrams (ASCII or described) where they aid understanding
7. Comparisons to related structures where applicable

## Output Conventions

- Final `.docx` → `output/`
- Intermediate Python scripts → `tmp/`
- Default language: Python, unless `WORK.md` specifies otherwise
- After saving, read back the file and iterate to fix any readability issues

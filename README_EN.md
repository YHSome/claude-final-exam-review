# Final Exam AI Tutor — Claude Code Skill

> 📖 [中文版 (Chinese)](README.md)

An AI tutor skill for Claude Code that helps students master course material from scratch and ace final exams. Think of it as a private tutor that reads your lecture slides and explains every concept in detail — in the browser, with beautifully rendered LaTeX formulas.

## Quick Start

```bash
git clone https://github.com/YHSome/claude-final-exam-review.git
cd claude-final-exam-review

# Drop your lecture files here (PPT/PDF/DOCX/images/HTML/Markdown — any format)
mkdir Calculus && # add your slides

# Optional: install document extraction dependencies
pip install unstructured

# Launch
claude
# Then type: /final-exam-review
```

## How It Works

- **Document Extraction** (`extract_document.py`) — Uses the `unstructured` library to parse PPT/PDF/DOCX/images (with OCR) into structured text, preserving titles, tables, and lists
- **Formula Completion** — AI fills in missing LaTeX formulas from extracted skeletons (PPT extractions lose embedded equation objects)
- **HTML Rendering** — All teaching content written to `_extracted/<subject>/<chapter>/xxx.讲义.html` with embedded MathJax, opened via local HTTP server `localhost:8888` (never `file://` — browsers block CDN loading)
- **Dual-Screen Mode** — Browser shows rendered content (formulas, diagrams, exercises); terminal handles brief interactions only ("got it", "next section", "give me exercises")
- **Hot Updates** — Q&A, exercises, and corrections appended directly to HTML files using the Edit tool (avoids bash/Python LaTeX escaping issues)

## Skill Structure

```
.claude/skills/
├── final-exam-review.md    # Skill definition (frontmatter + full teaching pipeline)
└── extract_document.py     # Multi-format document extractor
```

## Features

- 🧑‍🏫 **Tutor-first design** — Not a summarizer. Explains every concept with intuition first, then formal definition, then worked examples
- 📐 **STEM-optimized** — LaTeX formula completion, step-by-step derivations, "why" before "what"
- 📝 **Progressive exercises** — Level 1 (concept check) → Level 2 (comprehensive) → Level 3 (mock exam)
- 🔄 **Interactive Q&A** — Follow-up questions appended to lecture HTML, browser auto-refreshes
- 🌐 **Browser-native** — All content in the browser with MathJax rendering. Terminal = remote control, browser = blackboard.

## Technical Docs

| Resource | Link |
|:---|:---|
| Claude Code Docs | https://docs.anthropic.com/en/docs/claude-code |
| Skill Development Guide | https://docs.anthropic.com/en/docs/claude-code/skills |
| Skill Best Practices | https://docs.anthropic.com/en/docs/claude-code/skills#skill-best-practices |
| Claude API / Anthropic SDK | https://docs.anthropic.com/en/api |
| MathJax 3 Docs | https://docs.mathjax.org/en/latest/ |
| Unstructured Library | https://docs.unstructured.io/ |

## Disclaimer

This skill is for educational assistance only. Lecture materials belong to their original authors. Do not upload copyrighted course materials to public repositories.

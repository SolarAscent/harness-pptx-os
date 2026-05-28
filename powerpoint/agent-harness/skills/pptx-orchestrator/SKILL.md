---
name: pptx-orchestrator
description: Multi-agent orchestrator for autonomous PowerPoint generation. Coordinates specialized sub-agents to research content, design slide narratives, build slides via harness CLI, and QA the final deck. Use when asked to create a complete presentation from a brief, paper, or topic description.
---

# pptx-orchestrator — Multi-Agent PPT Generation

This skill coordinates a team of Claude Code sub-agents to autonomously create high-quality PowerPoint presentations from scratch. The orchestrator (you) manages the process end-to-end, delegating work to specialized agents at each phase.

## When to Use

Use this skill when the user asks to:
- "Create a presentation about X"
- "Make slides from this paper/notes/brief"
- "Generate a PPT for my talk on Y"
- Any request to produce a complete `.pptx` file

## Architecture

```
Orchestrator (You)
  ├── Phase 1: Research Agent — understands the topic
  ├── Phase 2: Narrative Agent — designs slide outline
  ├── Phase 3: Build Agent(s) — execute harness CLI commands
  └── Phase 4: QA Agent — reviews and fixes issues
```

## Phase 1: Research

**Goal**: Extract the core argument, key evidence, and structural outline from the source material.

Spawn a **Research Agent** (subagent_type="general-purpose") with this prompt template:

```
You are a research analyst. Extract the following from the source material:

SOURCE MATERIAL:
[user's brief / paper abstract / notes]

Extract and return:
1. TOPIC: One sentence describing the subject
2. AUDIENCE: Who this presentation is for (academic, business, general)
3. KEY CLAIM: The central thesis or main argument
4. EVIDENCE (3-5 items): Key findings, data points, or supporting arguments
5. STRUCTURE: Proposed slide structure (5 sections, each with 2-4 subsections)
6. KEY TERMS: Important technical terms with brief definitions

Return as structured text. No commentary, just the extracted information.
```

The Research Agent should NOT write slide content — only the raw material.

## Phase 2: Narrative Design

**Goal**: Transform research into a concrete slide-by-slide plan.

Spawn a **Narrative Agent** (subagent_type="general-purpose") with this prompt template:

```
You are a presentation designer. Using the research below, design a complete slide-by-slide outline.

RESEARCH:
[paste Research Agent output]

DESIGN SYSTEM:
- Canvas: 960x540 (16:9 widescreen)
- Colors: Purple primary (#6F2F9F), dark text (#3B3838), gray muted (#AFABAB), white bg
- Fonts: Calibri, 30pt section headings, 22pt body, 10-11pt page numbers
- Each slide gets a page number at (880, 505)

SLIDE ARCHETYPES available:
- cover: Title + presenter + date, two-line title at (80,140-205), 40pt bold
- contents: "CONTENTS" header + 5 numbered purple ovals, active section highlighted
- section-tagged: Purple oval section tag at (12,8,42x42) + "PART" label + title
- formula: Section-tagged content with LaTeX equations and gray annotations
- evidence: Section-tagged content with colored callout boxes
- closing: Centered "THANKS!" at (240,200), 64pt bold

OUTPUT FORMAT — For each slide, provide:
1. Slide number and archetype
2. SLIDE TITLE (conclusion-style, states the point not the topic)
3. 3-4 BULLETS (Chinese or English as appropriate, concise)
4. VISUAL TYPE (text-only, formula, table, diagram, evidence-boxes)
5. Any FORMULAS that need LaTeX rendering
6. Any TABLE structure (headers + rows)

The outline should be 10-16 slides. Default structure:
1. Cover
2. Contents
3-N. Content slides organized in 4-5 sections, each prefaced by a contents slide
Last. Closing

Return ONLY the slide outline, numbered, ready for building.
```

The Narrative Agent output is the build plan.

## Phase 3: Build

**Goal**: Execute harness CLI commands to construct each slide.

**CRITICAL**: Activate the harness first:
```bash
source /Users/tianyangsong/Code/Agent/test/pptx/.venv-office-harness/bin/activate
```

### Build Strategy

For a deck with N slides, the build script follows this pattern:

```bash
#!/bin/bash
PPT="cli-anything-powerpoint"
source /Users/tianyangsong/Code/Agent/test/pptx/.venv-office-harness/bin/activate

# Helper: section tag
section_tag() {
  local sl=$1; local num=$2; local title="$3"; local pg=$4; local prefix="$5"
  [ -z "$prefix" ] && prefix="$num"
  $PPT add-oval $sl 12 8 42 42 --fill-color "111,47,159" --line-color "111,47,159"
  $PPT add-text $sl 12 8 42 42 "$prefix" --font-size 18 --font-color "255,255,255" --bold --align center
  $PPT add-text $sl 10 50 44 18 "PART" --font-size 12 --font-color "111,47,159" --bold --align center
  $PPT add-text $sl 70 12 800 42 "$title" --font-size 30 --font-color "0,0,0" --bold
  $PPT add-text $sl 880 505 50 20 "$pg" --font-size 11 --font-color "137,137,137" --align right
}

# Helper: contents slide
contents_slide() {
  local sl=$1; local active=$2
  $PPT add-slide
  $PPT slide-bg $sl "255,255,255"
  $PPT add-text $sl 570 30 300 50 "CONTENTS" --font-size 44 --font-color "59,56,56" --bold
  local titles=("$3" "$4" "$5" "$6" "$7")
  local y=130
  for i in 1 2 3 4 5; do
    $PPT add-oval $sl 390 $y 40 40 --fill-color "111,47,159" --line-color "111,47,159"
    $PPT add-text $sl 390 $y 40 40 "$i" --font-size 24 --font-color "255,255,255" --bold --align center
    local color="175,171,171"
    [ "$i" -eq "$active" ] && color="59,56,56"
    $PPT add-text $sl 450 $((y+2)) 400 36 "${titles[$((i-1))]}" --font-size 24 --font-color "$color" --bold
    y=$((y+65))
  done
  $PPT add-text $sl 880 505 50 20 "$sl" --font-size 11 --font-color "137,137,137" --align right
}

# Start building
OUTPUT_DIR="output"
mkdir -p "$OUTPUT_DIR"

$PPT new
$PPT delete-slide 1

# === Build each slide here ===
# Slide 1: Cover
$PPT add-slide
$PPT slide-bg 1 "255,255,255"
# ... cover elements ...

# Continue with all slides from the Narrative Agent plan...

# === Save ===
$PPT save-as "$OUTPUT_DIR/presentation.pptx"
$PPT export-pdf "$OUTPUT_DIR/presentation.pdf" --pages all
$PPT close
```

### Formula Handling

When the Narrative Agent identifies a formula, use `add-latex`:

```bash
$PPT add-latex $slide "\\frac{a}{b} = \\sum_{i=1}^{n} x_i" \
  --left 80 --top 200 --font-size 18 --font-color "0,0,0" --bg-color "255,255,255"
```

### Table Handling

When the Narrative Agent identifies a table, build it with rects + text:

```bash
y=118; row_h=26
# Header
$PPT add-rect $slide 50 $y 860 $row_h --fill-color "111,47,159"
$PPT add-text $slide 55 $((y+2)) 200 22 "Col 1" --font-size 14 --font-color "255,255,255" --bold
$PPT add-text $slide 260 $((y+2)) 200 22 "Col 2" --font-size 14 --font-color "255,255,255" --bold
y=$((y+row_h))
# Data rows (alternating bg: 245,245,248 / 255,255,255)
for row in data; do
  # ... add rect + text for each cell
  y=$((y+row_h))
done
```

### Build Rules

1. **One slide at a time**: Never pre-create blank slides. Build slide N completely before slide N+1.
2. **Verify each slide**: After building, the agent should confirm the slide index matches expectations.
3. **Contents slides inline**: Place each contents slide at the correct position within the slide sequence.
4. **Page numbers**: Every slide gets a page number at (880, 505), font-size 11, color "137,137,137".
5. **Background first**: Always set `slide-bg` immediately after `add-slide`.

## Phase 4: QA and Repair

**Goal**: Review the built PPTX for visual issues and fix them.

### Quick QA (always run)

Export to PDF and check:
```bash
$PPT export-pdf "$OUTPUT_DIR/presentation.pdf" --pages all
```

Run these checks manually (or spawn a QA agent):
1. **Slide count**: Does the deck have the expected number of slides?
2. **Text overflow**: Are any text boxes too small for their content?
3. **Element overlap**: Do any shapes or text boxes intersect?
4. **Font size**: Are any fonts below 10pt (unreadable when projected)?
5. **Color consistency**: Do all elements use the design system colors?
6. **Formula rendering**: Do LaTeX formulas display correctly?

### Programmatic QA (optional, if harness_pptx is importable)

```python
from harness_pptx.qa.engine import QAEngine
from harness_pptx.models.scene_graph import SceneGraph

# If you built via SceneGraph, run:
engine = QAEngine(theme=theme)
report = engine.run(scene_graph)
for issue in report.issues:
    print(f"[{issue.severity.value}] {issue.category}: {issue.message}")
```

### Repair Loop

For each issue found:
1. Identify the affected element
2. Adjust position, size, font, or color
3. Re-render and verify the fix

Common fixes:
- **Text overflow**: Increase text box height or decrease font size by 2pt
- **Overlap**: Shift one element up/down by 10-20pt
- **Low contrast**: Darken text color or lighten background
- **Missing page number**: Add page number at (880, 505)

## Quality Bar

A "done" presentation must have:
- [ ] Correct slide count (10-16 for standard deck)
- [ ] Consistent purple design system on all slides
- [ ] Page numbers on every slide
- [ ] No text overflow (text fits within its box)
- [ ] No element overlap
- [ ] All formulas rendered via `add-latex` (not plain text)
- [ ] Tables have visible headers and alternating row colors
- [ ] Cover slide has title + presenter info
- [ ] Closing slide says "THANKS!"
- [ ] Contents slides highlight the correct active section
- [ ] PDF exports cleanly without rendering artifacts

## Common Failure Patterns

| Symptom | Likely Cause | Fix |
|---------|-------------|-----|
| Slide indices off by 1 | Default title slide not deleted | Always `delete-slide 1` after `new` |
| AppleScript error -1728 | Targeting a slide that doesn't exist yet | Build sequentially, verify slide count |
| Formula shows raw LaTeX | Used `add-text` instead of `add-latex` | Use `add-latex` for all formulas |
| Black rectangles in PDF | Transparent formula image on dark bg | Set `--bg-color "255,255,255"` |
| Contents slide wrong section highlighted | Wrong `active` parameter | Verify active section number |
| Text truncated | Text box too small for font size | Increase box height or shrink font |

## Output

The final deliverable is:
- `output/presentation.pptx` — editable PowerPoint file
- `output/presentation.pdf` — PDF export for review

Report to the user:
- Slide count and archetypes used
- Any formulas rendered via LaTeX
- Any tables constructed
- Known issues or limitations

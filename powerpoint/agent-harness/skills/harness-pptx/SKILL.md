---
name: harness-pptx
description: PowerPoint automation harness for creating, editing, and exporting PPTX files. Uses AppleScript on macOS and python-pptx on Windows/Linux for cross-platform support. Use for any PPT creation/editing task — from simple slide decks to complex academic presentations with LaTeX formulas.
---

# harness-pptx — PowerPoint Automation Skill

Use the CLI-Anything PowerPoint harness to create, edit, and export PowerPoint presentations.

**Platform support:**
- **macOS**: `cli-anything-powerpoint` drives Microsoft PowerPoint via AppleScript (full fidelity, all native features)
- **Windows / Linux**: `harness_pptx` pipeline uses python-pptx for cross-platform OOXML manipulation (no PowerPoint installation needed)

## Activation

**macOS:**
```bash
source /Users/tianyangsong/Code/Agent/test/pptx/.venv-office-harness/bin/activate
```

**Windows / Linux:**
```bash
pip install python-pptx
# Then use the harness_pptx Python pipeline:
python -c "from harness_pptx.pipeline import DeckPipeline; ..."
```

## Core Commands

| Command | Purpose |
|---------|---------|
| `cli-anything-powerpoint new` | Create a new presentation |
| `cli-anything-powerpoint open <file>` | Open an existing .pptx file |
| `cli-anything-powerpoint add-slide` | Add a blank slide |
| `cli-anything-powerpoint delete-slide <n>` | Delete slide at index n |
| `cli-anything-powerpoint add-text <slide> <left> <top> <w> <h> <"text"> [--options]` | Add a text box |
| `cli-anything-powerpoint add-oval <slide> <left> <top> <w> <h> [--options]` | Add an oval shape |
| `cli-anything-powerpoint add-rect <slide> <left> <top> <w> <h> [--options]` | Add a rectangle shape |
| `cli-anything-powerpoint add-image <slide> <left> <top> <w> <h> <path> [--options]` | Insert an image |
| `cli-anything-powerpoint add-line <slide> <x1> <y1> <x2> <y2> [--options]` | Draw a line |
| `cli-anything-powerpoint add-latex <slide> <"latex"> [--options]` | Render LaTeX formula as image |
| `cli-anything-powerpoint add-table <slide> <rows> <cols> <"json_data"> [--options]` | Add a table (text-box grid) |
| `cli-anything-powerpoint slide-bg <slide> <"r,g,b">` | Set slide background color |
| `cli-anything-powerpoint save-as <file>` | Save presentation to file |
| `cli-anything-powerpoint export-pdf <file> [--pages n]` | Export slides to PDF |
| `cli-anything-powerpoint close` | Close presentation without saving |

All commands support `--json` for machine-readable output.

## Critical Rules

### 1. Build slides one at a time
**Never pre-create multiple blank slides.** Build each slide completely before moving to the next. Pre-creating blank slides causes index confusion and AppleScript errors when targeting slides that don't yet exist.

Always start with:
```bash
cli-anything-powerpoint new
cli-anything-powerpoint delete-slide 1  # Remove default title slide
```

Then for each slide: `add-slide` → set background → add all elements → move to next.

### 2. Canvas is always 960×540 (16:9 widescreen)
All coordinates are in points. The bottom-right corner is at (960, 540). Keep a 20pt safe margin from canvas edges.

### 3. Bash backslash escaping
When calling `add-latex` from bash, double-escape backslashes:
```bash
cli-anything-powerpoint add-latex 1 "\\frac{a}{b} = \\sum_{i=1}^{n} x_i"
```

### 4. Use helper functions for repeated components
Define bash functions for reusable patterns (section tags, contents slides, etc.) to reduce code duplication and ensure consistency.

## Design System

### Color Palette (Purple Academic)

| Token | RGB | Hex | Usage |
|-------|-----|-----|-------|
| Purple (primary) | `111,47,159` | `#6F2F9F` | Section tags, accent ovals, highlights |
| Dark text | `59,56,56` | `#3B3838` | Headers, active section text |
| Gray (inactive) | `175,171,171` | `#AFABAB` | Inactive section text |
| Light gray | `137,137,137` | `#898989` | Page numbers, annotations |
| White | `255,255,255` | `#FFFFFF` | Slide backgrounds, text on dark shapes |
| Black | `0,0,0` | `#000000` | Body text |
| Red | `255,0,0` | `#FF0000` | Emphasis, warnings |
| Light purple | `233,235,245` | `#E9EBF5` | Evidence boxes, table headers |
| Table alt row | `245,245,248` | `#F5F5F8` | Alternating table row backgrounds |
| Table border | `200,200,210` | `#C8C8D2` | Table grid lines |

### Positioning Conventions

| Element | Position (left, top, w, h) | Notes |
|---------|---------------------------|-------|
| Section tag oval | `(12, 8, 42, 42)` | Purple, contains section number |
| Section "PART" label | `(10, 50, 44, 18)` | Purple text, below oval |
| Section title text | `(70, 12, 800, 42)` | Black bold, to right of oval |
| Page number | `(880, 505, 50, 20)` | Light gray |
| Contents header | `(570, 30, 300, 50)` | "CONTENTS", dark, 44pt |
| Contents oval | `(390, y, 40, 40)` | y starts at 130, +65 per item |
| Contents label | `(450, y+2, 400, 36)` | To right of oval |
| Body content start | y=85-120 | Varies by slide complexity |
| Cover title | `(80, 140-205, 800, 55)` | Two-line, 40pt bold |

### Typography Scale

| Role | Font Size | Bold? | Color |
|------|-----------|-------|-------|
| Cover title | 40 | Yes | `0,0,0` |
| Section heading | 30 | Yes | `0,0,0` |
| CONTENTS header | 44 | Yes | `59,56,56` |
| Contents item (active) | 24 | Yes | `59,56,56` |
| Contents item (inactive) | 24 | No | `175,171,171` |
| Oval number | 18-24 | Yes | `255,255,255` |
| Body heading | 22 | Yes | `0,0,0` |
| Body text | 22 | No | `0,0,0` |
| Sub-bullet | 14-16 | No | `0,0,0` |
| Table header | 14 | Yes | `255,255,255` |
| Table cell | 11-12 | No | `0,0,0` |
| Formula annotation | 13 | No | `137,137,137` |
| Page number | 10-11 | No | `137,137,137` |
| Citation | 12 | No | `0,0,0` |
| THANKS! | 64 | Yes | `0,0,0` |

### Slide Archetypes

**Cover Slide**: White background, two-line title at (80, 140-205), 40pt bold, presenter/date below, page number at (880, 505).

**Contents Slide**: White background, "CONTENTS" header, 5 numbered purple ovals in vertical list (y=130, stride=65), active section in dark bold, inactive in gray.

**Section-Tagged Content**: White background, section_tag header (purple oval + PART label + title), then body content below. This is the most common archetype.

**Formula-Dense Slide**: Section-tagged content with `add-latex` for equations. Use gray annotations (13pt, `137,137,137`) next to each formula. Use light purple background panels (`245,245,248`) for grouped formulas.

**Evidence Slide**: Section-tagged content with colored callout boxes. Use light purple (`233,235,245`), light green (`197,223,180`), or light orange (`250,228,213`) for evidence boxes.

**Closing Slide**: White background, large centered "THANKS!" at (240, 200, 480, 80), 64pt bold, page number.

## LaTeX Formula Detection Heuristic

When scanning content for formulas, use this decision tree:

1. **Greek letters** (α, β, λ, μ, Σ, etc.) → use `add-latex`
2. **Fractions** (`\frac{}{}`) → use `add-latex`
3. **Sums/Integrals** (`\sum`, `\int`, `\prod`, `\lim`) → use `add-latex`
4. **Nested subscripts** (`x_{i,j}`) or combined super/subscripts (`x_i^2`) → use `add-latex`
5. **Multi-line or display equations** → use `add-latex`
6. **Simple inline** variables like `x_i` or `a^2` → `add-text` with Unicode subscripts is OK
7. **Plain text with no math** → use `add-text`

### add-latex Options

```bash
cli-anything-powerpoint add-latex <slide> '<latex>' \
  --left 80 --top 100 \
  --font-size 18 \
  --font-color "0,0,0" \
  --bg-color "255,255,255" \
  --dpi 200
```

- `--width` / `--height`: rendered image size (0 = auto-size from formula)
- `--font-size`: mathtext font size (default 18)
- `--font-color`: `"r,g,b"` (default `"0,0,0"`)
- `--bg-color`: `"r,g,b"` (default `"255,255,255"` — match slide background)
- `--dpi`: rendering resolution (default 200)

## Table Construction Pattern

Since the harness uses text-box grids (not native tables), build tables like this:

```bash
# Table header row
cli-anything-powerpoint add-rect $slide 50 $y 860 26 --fill-color "111,47,159"
cli-anything-powerpoint add-text $slide 55 $((y+2)) 200 22 "Header 1" --font-size 14 --font-color "255,255,255" --bold
cli-anything-powerpoint add-text $slide 260 $((y+2)) 200 22 "Header 2" --font-size 14 --font-color "255,255,255" --bold

# Table data rows (alternating backgrounds)
for row in data; do
  y=$((y+26))
  bg=$([ $((row_num % 2)) -eq 0 ] && echo "245,245,248" || echo "255,255,255")
  cli-anything-powerpoint add-rect $slide 50 $y 860 26 --fill-color "$bg"
  # Add cell text...
done
```

## Bash Function Templates

### section_tag

```bash
section_tag() {
  local sl=$1; local num=$2; local title="$3"; local pg=$4; local prefix="$5"
  [ -z "$prefix" ] && prefix="$num"
  cli-anything-powerpoint add-oval $sl 12 8 42 42 --fill-color "111,47,159" --line-color "111,47,159"
  cli-anything-powerpoint add-text $sl 12 8 42 42 "$prefix" --font-size 18 --font-color "255,255,255" --bold --align center
  cli-anything-powerpoint add-text $sl 10 50 44 18 "PART" --font-size 12 --font-color "111,47,159" --bold --align center
  cli-anything-powerpoint add-text $sl 70 12 800 42 "$title" --font-size 30 --font-color "0,0,0" --bold
  cli-anything-powerpoint add-text $sl 880 505 50 20 "$pg" --font-size 11 --font-color "137,137,137" --align right
}
```

### contents_slide

```bash
contents_slide() {
  local sl=$1; local active=$2
  cli-anything-powerpoint add-slide
  cli-anything-powerpoint slide-bg $sl "255,255,255"
  cli-anything-powerpoint add-text $sl 570 30 300 50 "CONTENTS" --font-size 44 --font-color "59,56,56" --bold
  local sections=("Introduction" "Problem Formalization" "Framework" "Parameter Estimation" "Experiments and Conclusion")
  local y=130
  for i in 1 2 3 4 5; do
    cli-anything-powerpoint add-oval $sl 390 $y 40 40 --fill-color "111,47,159" --line-color "111,47,159"
    cli-anything-powerpoint add-text $sl 390 $y 40 40 "$i" --font-size 24 --font-color "255,255,255" --bold --align center
    local color="175,171,171"
    [ "$i" -eq "$active" ] && color="59,56,56"
    cli-anything-powerpoint add-text $sl 450 $((y+2)) 400 36 "${sections[$((i-1))]}" --font-size 24 --font-color "$color" --bold
    y=$((y+65))
  done
  cli-anything-powerpoint add-text $sl 880 505 50 20 "$sl" --font-size 11 --font-color "137,137,137" --align right
}
```

## Build Order Pattern

```bash
PPT="cli-anything-powerpoint"
$PPT new
$PPT delete-slide 1

# Slide 1: Cover
$PPT add-slide
$PPT slide-bg 1 "255,255,255"
# ... cover elements ...

# Slide 2: Contents (active=1)
contents_slide 2 1

# Slide 3: Section 1 content
$PPT add-slide
$PPT slide-bg 3 "255,255,255"
section_tag 3 1 "Introduction" 3
# ... body content ...

# ... continue for all slides ...

$PPT save-as "output/presentation.pptx"
```

## Output Path

Always save to the project's `output/` directory:
```bash
OUTPUT_DIR="/path/to/project/output"
mkdir -p "$OUTPUT_DIR"
$PPT save-as "$OUTPUT_DIR/presentation.pptx"
```

# PPT Harness Agent Guide

You do not directly place PowerPoint shapes unless asked to debug rendering.
Your job is to convert user intent into a valid DeckSpec.

## ⚠️ MANDATORY: LLM Content Generation Rules

**READ THIS SECTION FIRST. These rules are enforced by the pipeline.**
**Violating them produces broken slides. Follow them exactly.**

### The Core Principle

This harness has 25 professionally-designed slide type templates with
predefined layouts, automatic alignment, theme tokens, and QA/repair.
Each template expects specific structured fields. Your ONLY job is to
produce RICH, DENSE content that fills these fields.

**Sparse content = wasted template = broken presentation.**

### Slide Type Field Reference (COMPLETE)

Every slide type requires specific fields. When generating a SlideIntent,
populate `extra` with ALL required fields for that type.

| Type | Required Fields | Optional Fields | Content Shape |
|------|----------------|-----------------|---------------|
| `cover` | title | subtitle, author, date | Single text + metadata |
| `agenda` | title, **items** (list) | numbered | 4-8 bullet items |
| `section-divider` | title | section_number, subtitle, background_color | Centered title on color block |
| `executive-summary` | title, **key_points** (list) | subtitle, bottom_line | 3-5 key points + optional bottom line |
| `problem` | title, **problem_statement** | pain_points (list), impact, context | Callout box + 2-4 pain points |
| `solution` | title, **solution_summary** | key_features (list), how_it_works, benefits | Callout box + 2-5 features |
| `timeline` | title, **milestones** (list of {date, event}) | — | Horizontal timeline, 3-8 dots |
| `process` | title, **steps** (list of {label, name, description}) | orientation | Numbered flow, 3-6 steps |
| `framework` | title, **framework_name** | components (list of {title, description}), description | Conceptual boxes, 3-5 components |
| `comparison` | title, **left** (list), **right** (list) | left_label, right_label | Two-column, 2-5 per side |
| `before-after` | title, **before_points** (list), **after_points** (list) | before_label, after_label | Two-column contrast |
| `data-insight` | title, **insight** | chart_data, supporting_points (list), source | Chart area + 1-3 points |
| `chart` | title, **chart_data** (dict) | chart_type, caption, source | Native chart |
| `table` | title, **headers** (list), **rows** (list of lists) | caption, source | Structured table |
| `case-study` | title | company, challenge, solution, results, quote | Narrative layout |
| `quote` | **quote_text** | attribution, role, context | Large pull-quote |
| `team` | title, **members** (list of {name, role, bio}) | layout | People grid |
| `roadmap` | title, **phases** (list of {name, timeline, description}) | timeline_label | Phase cards |
| `architecture` | title | layers (list of {name, description}), components, description | Layer diagram |
| `workflow` | title, **steps** (list of {name, description}) | direction, roles | Pipeline flow |
| `risk` | title, **risks** (list of {name, level, description}) | mitigation, risk_matrix | Risk cards |
| `recommendation` | title, **recommendations** (list) | rationale, next_steps, priority | Numbered recs + CTA |
| `conclusion` | title, **key_takeaways** (list) | call_to_action, subtitle | 3-6 takeaways + CTA button |
| `thank-you` | (none required) | message, contact, email, website, subtitle | Centered message + contact |
| `appendix` | title, **content** (list) | type, reference, section_label | Supplementary |

### Content Density Rules

1. **MINIMUM DENSITY**: Every body slide MUST have at least 3 content items
   (bullets, steps, milestones, key_points, etc.). A slide with only a title
   and 1-2 items WILL be flagged as broken by QA.
2. **MAXIMUM DENSITY**: No more than 6 items per list. No more than 15 total
   elements per slide. Split dense slides rather than cramming.
3. **RICH TEXT**: `key_message` must be 50-150 characters of SUBSTANTIVE text
   — never a one-word label, never the same as the title, never empty.
4. **TITLE AS MESSAGE**: Every title must state the KEY MESSAGE of the slide.
   "30% Readmission Reduction via Predictive Analytics" → good.
   "Results" → bad. The audience should understand the slide from the title alone.

### Structured Data Requirements

List fields (items, steps, milestones, members, risks, etc.) MUST contain
properly structured objects, NOT flat strings. Examples:

```json
// CORRECT: milestones as array of objects
"extra": {
  "milestones": [
    {"date": "Q1 2025", "event": "Discovery & Research Phase"},
    {"date": "Q2 2025", "event": "MVP Development & Testing"},
    {"date": "Q3 2025", "event": "Beta Launch with 5 Hospitals"},
    {"date": "Q4 2025", "event": "Full Market Rollout"}
  ]
}

// WRONG: milestones as flat strings
"extra": {
  "milestones": ["Q1", "Q2", "Q3", "Q4"]
}
```

### SlideIntent Output Format

Always populate the `extra` dict with ALL required fields for the chosen
slide type. The pipeline's `_populate_slide_content` function reads from
`extra` first, then `key_message`, then generates sensible defaults.

```json
{
  "slide_id": "slide-3",
  "seq": 3,
  "slide_type": "problem",
  "title": "Rising Readmission Rates Threaten Hospital Viability",
  "key_message": "Hospitals face $2.4M annual penalties from CMS readmission penalties, with 30-day readmission rates averaging 15% across the industry.",
  "extra": {
    "problem_statement": "The healthcare industry loses $26B annually due to preventable readmissions, with no scalable solution in market.",
    "pain_points": [
      "Fragmented patient data across EMR systems prevents holistic risk assessment",
      "Reactive intervention models identify at-risk patients only after deterioration",
      "Manual discharge planning misses 40% of high-risk patients due to workflow gaps"
    ],
    "impact": "This costs the average 300-bed hospital $2.4M per year in penalties alone."
  }
}
```

## Workflow

1. Read the user's raw request.
2. Identify audience, goal, tone, language, duration, and required format.
3. Produce a deck outline before producing slides.
4. Choose slide types from the 25-type library using the reference table above.
5. **Populate ALL required fields for each slide type in extra.**
6. **Ensure minimum 3 content items per body slide.**
7. Keep each slide to one main idea.
8. Prefer structured content over prose.
9. Use theme tokens, never raw colors unless required by brand.
10. Do not use absolute coordinates — the layout engine handles positioning.
11. Add speaker notes when useful.
12. Run build → preview → QA → repair until QA passes.

## Visual Design Standards

These rules prevent the most common design issues. The templates follow these
rules automatically. You do NOT need to specify coordinates.

### Canvas

- Default slide size: 960pt × 540pt (16:9 widescreen)
- Safe content area: x=80..880, y=40..500 (80pt horizontal margins, 40pt top, 40pt bottom above accent bar)

### Font Sizes (template-managed)

- Title: 26–36pt (28pt recommended)
- Subtitle: 16–20pt (18pt recommended)
- Body text: 14–16pt (15pt recommended)
- Callout/highlight: 16–18pt
- Kicker/label: 13–14pt
- Contact info: 16pt minimum
- Never use font sizes below 12pt — they are unreadable on projected slides

### Alignment (template-managed)

- Title x must match the left edge of content elements below it (typically x=80)
- Content elements within the same visual group must share the same x offset
- Centered text must use `alignment: "center"` with equal left/right margins
- Vertically center content blocks on the slide (optical center ≈ y=250 for the content block)

### Spacing & Vertical Distribution (template-managed)

- Distribute content vertically across the full slide — avoid concentrating all elements in the top half
- Minimum 40pt between distinct content groups
- Minimum 32pt between items within a group
- The bottom 28pt of every slide has an accent bar at (0, 512, 960, 28)
- Body slides without accent bars feel unfinished and bottom-heavy

### Visual Anchoring (template-managed)

- Every slide type includes a full-width accent bar at the bottom edge
- Cover and section-divider slides may use full-bleed backgrounds instead
- Accent bars (4px wide colored rectangles) appear left of key bullet points for emphasis

### Density & Content Limits

- Max 6 bullets per slide
- Max 5 key points on executive summary
- Max 8 agenda items
- Max 15 total elements per slide (QA check will warn above this)
- **Min 3 content items per body slide** (QA will flag slides with fewer)
- If content exceeds limits, split into multiple slides

### Color & Contrast (template-managed)

- Always resolve semantic color tokens via Theme (never hardcode hex values)
- `font_color="background"` only on dark-filled shapes (primary, accent backgrounds)
- `font_color="muted"` should be used sparingly — it reduces contrast
- Ensure text-on-background has sufficient contrast for projection

## Output Contract

Always produce:

- deck.spec.json
- theme.json
- asset manifest
- final pptx
- preview images
- qa report

## Slide Rules

- One idea per slide.
- Title must state the message, not just the topic.
- No slide should contain more than 6 bullets.
- **No slide should contain fewer than 3 content items.**
- If content is dense, split it.
- Use diagrams for relationships.
- Use tables only for comparison.
- Use charts only when there is numerical data.
- Use callouts to explain why a visual matters.

## Narrative Flow

```
Cover → (Agenda if >7 slides) → Executive Summary → Problem → Solution
→ How It Works (Process/Architecture) → Evidence (Timeline/Data/Chart/Comparison)
→ Team/Risk/Roadmap (as relevant) → Recommendations → Conclusion → Thank You
```

## Component Selection

Use:
- `cover` for first slide
- `agenda` for decks longer than 7 slides
- `section-divider` for major transitions
- `executive-summary` for high-level overviews
- `problem` for challenge/pain-point slides
- `solution` for solution overviews
- `process` for step-by-step flows
- `timeline` for schedules and milestones
- `data-insight` for charts with commentary
- `comparison` for side-by-side alternatives
- `framework` for conceptual models
- `architecture` for system diagrams
- `conclusion` for final recommendation
- `thank-you` for closing slide

## Do Not

- Do not insert full-slide screenshots as a shortcut.
- Do not manually guess pixel positions.
- Do not create slides with fewer than 3 content items.
- Do not leave `key_message` empty or with a one-word placeholder.
- Do not leave `extra` empty when the slide type requires specific fields.
- Do not create unreadable dense slides (>15 elements, >6 bullets).
- Do not ignore QA errors or warnings.
- Do not use shape index as stable identity.
- Do not use font sizes below 12pt.
- Do not leave large empty areas at the bottom of slides.
- Do not misalign title x with content x.

## Architecture

```
User Text → Brief Parser → Story Planner → Slide Outline → Slide Intent Spec
→ Theme + Layout Engine → Editable Deck Scene Graph → PowerPoint Renderer
→ Preview Export → QA + Repair Loop → Final PPTX
```

## Key Modules

| Module | Path | Purpose |
|--------|------|---------|
| Models | `harness_pptx/models/` | Pydantic data contracts |
| Themes | `harness_pptx/themes/` | Design tokens & presets |
| Slide Types | `harness_pptx/slide_types/` | 25 template functions |
| Layout | `harness_pptx/layout/` | vstack/hstack/grid/split |
| Backends | `harness_pptx/backends/` | AppleScript, PPTX XML |
| Renderer | `harness_pptx/renderer/` | SceneGraph → PowerPoint |
| Content | `harness_pptx/content/` | LLM-driven text understanding |
| Narrative | `harness_pptx/narrative/` | SCQA, Pyramid, Three-Act |
| QA | `harness_pptx/qa/` | Automated quality checks |
| Repair | `harness_pptx/repair/` | Auto-fix pipeline |
| CLI | `harness_pptx/cli/` | Commands & workspace |
| Pipeline | `harness_pptx/pipeline.py` | End-to-end orchestrator |

## Quick Start

```python
from harness_pptx.pipeline import DeckPipeline, PipelineConfig

config = PipelineConfig(theme="corporate", target_slides=12)
pipeline = DeckPipeline(config)
result = pipeline.run("We need a pitch deck for our AI startup...")
print(f"Done: {result.output_path}")
```

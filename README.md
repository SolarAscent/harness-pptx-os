# 🎬 Harness PPTX — The PowerPoint Generation Operating System

<p align="center">
  <b>Text → Brief → Story → Outline → SceneGraph → Theme + Layout → PPTX → QA → Repair</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-0.2.0-blue" alt="version">
  <img src="https://img.shields.io/badge/python-3.10%2B-green" alt="python">
  <img src="https://img.shields.io/badge/platform-macOS/Windows-lightgrey" alt="platform">
  <img src="https://img.shields.io/badge/slide_types-25-orange" alt="slide types">
  <img src="https://img.shields.io/badge/themes-7-purple" alt="themes">
</p>

---

**Harness PPTX** transforms raw text ideas into professional, editable PowerPoint
presentations through a 13-step AI-powered pipeline. It's not a PowerPoint
remote-control — it's a **presentation operating system** that understands
content, plans narratives, designs layouts, renders slides, checks quality, and
repairs defects automatically.

> Built for AI agents. Drives real Microsoft PowerPoint via AppleScript.
> macOS native. 25 slide types. 7 themes. Vision-powered QA.

---

## ✨ Features

| Category | Capability |
|----------|-----------|
| 🧠 **Content Understanding** | LLM-powered brief parsing, story planning, outline building, intent classification |
| 🎨 **Design System** | 7 preset themes (corporate, academic, startup, consulting, technical, minimal, dark) with semantic color tokens, font stacks, spacing scales |
| 📐 **Layout Engine** | Declarative primitives: `vstack`, `hstack`, `grid`, `split`, `columns` — no manual coordinates |
| 🧩 **25 Slide Types** | Cover, Agenda, Problem, Solution, Process, Timeline, Comparison, Framework, Architecture, Data Insight, Chart, Table, Team, Roadmap, Risk, Recommendation, Conclusion, Thank You, and more |
| ⚡ **PowerPoint Rendering** | Drives real Microsoft PowerPoint via AppleScript — native shapes, text boxes, charts, tables |
| 🔍 **QA System** | 7 automated checks: text overflow, overlap, font size, margins, density, contrast, color theme |
| 👁 **Vision Review** | AI-powered visual inspection of rendered slides via vision API (doubao-seed-2-0-pro) |
| 🔧 **Auto-Repair** | Detected issues trigger automatic repair loop — re-render until QA passes |
| 🎯 **Agent-Native** | All AI agents produce structured intents, not manual shape commands |
| 🌐 **Backend Abstraction** | AppleScript (macOS) + PPTX XML skeleton for cross-platform |

## 🚀 Quick Start

### Prerequisites

- macOS with Microsoft PowerPoint installed
- Python 3.10+
- (Optional) Vision API key for AI-powered visual review

### Installation

```bash
# Clone and install
git clone https://github.com/SolarAscent/harness-pptx-os.git
cd harness-pptx-os
pip install -e powerpoint/agent-harness

# macOS: use the bundled virtual environment
source powerpoint/.venv-office-harness/bin/activate
```

### 5-Second Demo

```python
from harness_pptx.pipeline import DeckPipeline, PipelineConfig

config = PipelineConfig(
    theme="corporate",
    target_slides=12,
    output_dir="./output",
    enable_visual_review=False,  # Set True for AI visual review
)

pipeline = DeckPipeline(config)
result = pipeline.run("We need an investor pitch deck for our AI healthcare startup.")

print(f"Done! → {result.output_path}")
print(f"QA passed: {result.qa_report.passed}")
print(f"Issues: {len(result.qa_report.issues)}")
```

**What happens under the hood:**

1. **Brief Parser** extracts audience, goals, tone, key messages from your text
2. **Story Planner** designs a narrative arc: Cover → Agenda → Problem → Solution → Evidence → Conclusion
3. **Outline Builder** determines slide count, sections, and page structure
4. **Intent Classifier** assigns the best slide type to each page from the 25-type library
5. **Slide Types** build structured SceneGraph nodes with proper layouts and theme tokens
6. **PowerPoint Renderer** drives real PowerPoint via AppleScript to create native shapes and text
7. **QA Engine** runs 7 automated checks on the rendered output
8. **Repair Engine** fixes any detected issues in an automatic loop

---

## 📖 Tutorial

### 1. Basic Usage — Generate a Deck from Text

```python
from harness_pptx.pipeline import DeckPipeline, PipelineConfig

# The simplest possible invocation
config = PipelineConfig(theme="corporate", target_slides=10)
pipeline = DeckPipeline(config)
result = pipeline.run("""
We need a product launch presentation for our new SaaS analytics platform.
Target audience: CTOs and data leaders at mid-market companies.
Key message: democratizing data science with no-code ML pipelines.
Style: modern, data-driven, inspiring.
""")

print(f"Output: {result.output_path}")
print(f"Slides: {len(result.intents)}")
print(f"Brief: audience={result.brief.audience}, tone={result.brief.tone}")
```

### 2. Choosing a Theme

```python
from harness_pptx.themes.base import ThemeRegistry

themes = ThemeRegistry()
print("Available themes:", themes.list())
# → ['corporate', 'academic', 'startup', 'consulting', 'technical', 'minimal', 'dark']

# Each theme has its own color palette, fonts, and spacing
corp = themes.get("corporate")
print(corp.colors)
# → primary: #003366, accent: #0078D7, background: #FFFFFF, ...

# Use any theme in the pipeline
config = PipelineConfig(theme="dark", target_slides=12)
```

### 3. Programmatic Slide Creation

```python
from harness_pptx.slide_types.base import get_slide_type_registry
from harness_pptx.themes.base import ThemeRegistry

reg = get_slide_type_registry()
theme = ThemeRegistry().get("corporate")

# Build individual slides with complete content
content = {
    "slide_id": "problem-1",
    "seq": 2,
    "title": "Rising Readmission Rates Threaten Hospital Margins",
    "problem_statement": "Preventable readmissions cost the US healthcare system $26B annually.",
    "pain_points": [
        "Fragmented patient data across 5+ EMR systems per hospital",
        "Reactive intervention models identify risk only after deterioration",
        "Manual discharge planning misses 40% of high-risk patients",
    ],
    "impact": "Average 300-bed hospital loses $2.4M/year in CMS penalties.",
}

slide_type = reg.get("problem")
slide_node = slide_type.build(content, theme, registry={})
print(f"Slide: {slide_node.id} → {len(slide_node.all_element_ids())} elements")
```

### 4. Working with the SceneGraph

```python
from harness_pptx.models.scene_graph import SceneGraph

# The SceneGraph is the universal intermediate representation
sg = result.scene_graph

# Inspect every slide and element
for slide in sg.slide_order():
    print(f"\n=== {slide.id} [{slide.slide_type}] ===")
    for eid in slide.all_element_ids():
        el = sg.get_element(eid)
        if el:
            print(f"  {el.role.value}: {getattr(el, 'text', el.id)[:60]}")

# The element registry maps stable IDs → element objects
print(f"\nTotal elements: {len(sg.element_registry)}")
```

### 5. Running QA & Visual Review

```python
from harness_pptx.qa.engine import QAEngine

# Model-based QA — instant, no API needed
qa = QAEngine(theme=theme)
report = qa.run(sg)

for issue in report.issues:
    print(f"[{issue.severity.value}] {issue.category.value}: {issue.message}")

# AI-powered visual review — requires ARK_API_KEY
config = PipelineConfig(enable_visual_review=True)
pipeline = DeckPipeline(config)
result = pipeline.run("Your presentation brief...")
# Visual review runs automatically, results in result.qa_report
```

### 6. Inspect a Rendered PPTX

```python
from harness_pptx.backends.registry import get_backend

backend = get_backend()
pres = backend.open_presentation("./output/my-deck.pptx")

# Count slides
count = backend.slide_count(pres)
print(f"Slides: {count}")

# Export as PNGs for inspection
pngs = backend.export_all_pngs(pres, "./previews/")
print(f"Exported {len(pngs)} preview images")

backend.close(pres)
```

---

## 🧩 Slide Type Gallery

All 25 types with their required content fields:

| # | Type | Required Fields | Visual |
|---|------|----------------|--------|
| 1 | `cover` | title | Centered title + subtitle + author + date |
| 2 | `agenda` | title, **items** (list) | Numbered list, 4–8 items |
| 3 | `section-divider` | title | Full-bleed colored background |
| 4 | `executive-summary` | title, **key_points** (list) | Bulleted key points + bottom line |
| 5 | `problem` | title, **problem_statement** | Callout box + pain points + impact |
| 6 | `solution` | title, **solution_summary** | Callout box + features + how-it-works |
| 7 | `timeline` | title, **milestones** ({date, event}) | Horizontal timeline dots |
| 8 | `process` | title, **steps** ({label, name, desc}) | Numbered flow, 3–6 steps |
| 9 | `framework` | title, **framework_name** | Conceptual boxes, 3–5 components |
| 10 | `comparison` | title, **left** (list), **right** (list) | Two-column with labels |
| 11 | `before-after` | title, **before_points**, **after_points** | Transformation contrast |
| 12 | `data-insight` | title, **insight** | Chart area + supporting points |
| 13 | `chart` | title, **chart_data** (dict) | Native PowerPoint chart |
| 14 | `table` | title, **headers** (list), **rows** (list) | Structured data table |
| 15 | `case-study` | title | Narrative: company, challenge, solution, results |
| 16 | `quote` | **quote_text** | Large pull-quote with attribution |
| 17 | `team` | title, **members** ({name, role, bio}) | People grid |
| 18 | `roadmap` | title, **phases** ({name, timeline, desc}) | Phase cards |
| 19 | `architecture` | title | Layer diagram |
| 20 | `workflow` | title, **steps** ({name, description}) | Pipeline flow |
| 21 | `risk` | title, **risks** ({name, level, desc}) | Risk assessment cards |
| 22 | `recommendation` | title, **recommendations** (list) | Numbered actions + CTA |
| 23 | `conclusion` | title, **key_takeaways** (list) | Takeaways + CTA button |
| 24 | `thank-you` | (none required) | Centered message + contact info |
| 25 | `appendix` | title, **content** (list) | Supplementary material |

---

## 🎨 Theme System

```
themes/presets/
├── corporate.json      # Deep blue, professional — for business & enterprise
├── academic.json       # Clean serif, muted tones — for research & education
├── startup.json        # Bold accent, modern feel — for pitch decks
├── consulting.json     # Refined grays, precision — for strategy decks
├── technical.json      # Monospace accents, cool palette — for engineering
├── minimal.json        # Bare essentials, black & white — for content focus
└── dark.json           # Dark background, high contrast — for screens & night
```

Each theme defines:

| Token | Example (corporate) |
|-------|-------------------|
| `colors.primary` | `#003366` |
| `colors.accent` | `#0078D7` |
| `colors.background` | `#FFFFFF` |
| `colors.text` | `#1A1A1A` |
| `colors.muted` | `#767676` |
| `fonts.title` | `Calibri` |
| `spacing.lg` | `24pt` |
| `radii.md` | `6pt` |

---

## 🔍 QA Checks

The QA engine runs 7 automated model-based checks on every slide:

| Check | Description | Severity |
|-------|-------------|----------|
| Text Overflow | Text estimated to exceed bounding box | Warning |
| Element Overlap | Two elements occupying the same space | Warning |
| Font Size | Text below 9pt minimum | Warning |
| Margin | Element within 20pt of canvas edge | Warning |
| Slide Density | More than 15 elements on one slide | Warning |
| Contrast | Text/background contrast analysis | Info |
| Color Theme | Color compliance with theme tokens | Info |

### AI-Powered Visual Review

For the highest quality, enable vision-based review. This renders each slide to
PNG and sends it to a vision-capable LLM for professional design inspection
across 8 dimensions: text overflow, alignment, font sizing, overlap, contrast,
margin, information density, and overall aesthetics.

```bash
# Set up vision API key (doubao-seed-2-0-pro via Volcengine ARK)
export ARK_API_KEY="your-api-key"

# Enable in pipeline config
config = PipelineConfig(enable_visual_review=True)
```

---

## 📁 Project Structure

```
harness-pptx-os/
├── README.md                  # This file (repo root)
│
├── powerpoint/
│   ├── agent-harness/         # Core harness_pptx package
│   │   ├── setup.py
│   │   ├── .gitignore
│   │   ├── harness_pptx/
│   │   │   ├── __init__.py        # v0.2.0
│   │   │   ├── pipeline.py        # DeckPipeline orchestrator (13 steps)
│   │   │   ├── AGENT.md           # AI agent usage guide
│   │   │   │
│   │   │   ├── models/            # Pydantic data contracts
│   │   │   │   ├── content.py     # Brief, Outline, SlideIntent
│   │   │   │   ├── element.py     # TextElement, ShapeElement, ChartElement...
│   │   │   │   ├── scene_graph.py # SceneGraph, SlideNode, LayerNode
│   │   │   │   ├── theme.py       # Theme color/font/spacing tokens
│   │   │   │   ├── layout.py      # BBox, LayoutSpec, alignment primitives
│   │   │   │   └── qa.py         # QAIssue, QAReport, Severity enum
│   │   │   │
│   │   │   ├── content/           # LLM-driven text understanding
│   │   │   │   ├── brief_parser.py
│   │   │   │   ├── story_planner.py
│   │   │   │   ├── intent_classifier.py
│   │   │   │   └── prompts.py     # LLM prompt templates
│   │   │   │
│   │   │   ├── slide_types/       # 25 template functions
│   │   │   │   ├── cover.py, agenda.py, problem.py, solution.py...
│   │   │   │   └── base.py        # SlideType protocol + registry
│   │   │   │
│   │   │   ├── themes/            # Design system
│   │   │   │   ├── base.py        # ThemeRegistry
│   │   │   │   └── presets/       # 7 theme JSON files
│   │   │   │
│   │   │   ├── backends/          # PowerPoint rendering
│   │   │   │   ├── interface.py   # RendererInterface (ABC)
│   │   │   │   └── applescript_backend.py  # macOS native (AppleScript)
│   │   │   │
│   │   │   ├── renderer/          # SceneGraph → PowerPoint
│   │   │   ├── qa/                # QA engine + visual review
│   │   │   ├── repair/            # Auto-repair system
│   │   │   ├── vision/            # Vision API client
│   │   │   ├── layout/            # Layout engine
│   │   │   ├── cli/               # CLI commands & workspace
│   │   │   ├── narrative/         # Story frameworks
│   │   │   ├── stability/         # Transaction/retry/logging
│   │   │   └── memory/            # Agent memory
│   │   │
│   │   └── cli_anything/          # Legacy CLI remote-control
│   │
│   └── .venv-office-harness/      # macOS virtual environment
│
├── word/                          # Word harness (sibling project)
└── excel/                         # Excel harness (sibling project)
```

---

## 🔧 CLI Commands

```bash
# Generate a complete deck from text
ppt-harness create-from-text input.md --audience "investors" --style "data-driven" --slides 12

# Plan a deck outline
ppt-harness plan input.md --out outline.json

# Build from a spec file
ppt-harness build spec.json --theme corporate --out deck.pptx

# Run QA on an existing deck
ppt-harness qa deck.pptx --out report.json

# Inspect a deck as SceneGraph JSON
ppt-harness inspect deck.pptx --out scene.json

# Auto-repair a deck
ppt-harness repair deck.pptx --ref spec.json
```

---

## 🧪 Running Tests

```bash
# Install dev dependencies
pip install pytest pytest-cov

# Run the test suite
pytest harness_pptx/ -v

# Run with coverage
pytest harness_pptx/ --cov=harness_pptx --cov-report=html
```

---

## 🤖 For AI Agents

This project includes a comprehensive [AGENT.md](harness_pptx/AGENT.md) guide
specifically for AI agents. It contains:

- **Mandatory content generation rules** — minimum density, structured data formats
- **Complete slide type field reference** — all 25 types with required/optional fields
- **Visual design standards** — font sizes, alignment, spacing, anchoring
- **Output contract** — what files to produce

If you're an AI agent working with this codebase, **read AGENT.md first**.
The rules in it are enforced by the pipeline and violating them produces
broken slides.

### Key Rules for AI Agents

1. ⚠️ Every body slide must have ≥3 content items
2. ⚠️ `key_message` must be 50-150 chars of substantive text
3. ⚠️ `extra` dict must contain ALL required fields for the chosen slide type
4. ⚠️ Structured fields (milestones, steps, members) must be arrays of objects, not flat strings
5. ⚠️ Titles must state the message, not label the topic
6. ⚠️ Never leave `extra` empty when the slide type has required fields

---

## 📄 License

MIT


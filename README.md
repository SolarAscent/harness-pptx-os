# Harness PPTX

PowerPoint automation for AI agents on macOS. Drives Microsoft PowerPoint via
AppleScript to create, edit, and export `.pptx` files from structured slide
specifications.

## Installation

Requires macOS with Microsoft PowerPoint installed, and Python 3.10+.

```bash
git clone https://github.com/SolarAscent/harness-pptx-os.git
cd harness-pptx-os
pip install -e powerpoint/agent-harness
```

On macOS, activate the bundled virtual environment:

```bash
source powerpoint/.venv-office-harness/bin/activate
```

## Quick start

```python
from harness_pptx.pipeline import DeckPipeline, PipelineConfig

config = PipelineConfig(theme="corporate", target_slides=10, output_dir="./output")
pipeline = DeckPipeline(config)
result = pipeline.run(
    "We need a product launch presentation for our SaaS analytics platform. "
    "Target audience: CTOs and data leaders. Style: modern, data-driven."
)
print(f"Done: {result.output_path}")
```

## Pipeline

The pipeline processes freeform text into a rendered `.pptx` through six stages:

1. **Brief Parser** — extracts audience, tone, and key messages
2. **Story Planner** — designs a narrative arc for the deck
3. **Outline Builder** — determines slide count, sections, and page structure
4. **Intent Classifier** — assigns a slide type to each page from 25 templates
5. **Slide Types + Layout** — builds a SceneGraph with resolved coordinates using
   declarative primitives (`vstack`, `hstack`, `grid`, `split`)
6. **PowerPoint Renderer** — AppleScript backend drives native shapes and text

Optional QA and repair loop runs automated checks (text overflow, element overlap,
font size, margin, density, contrast) and re-renders until issues are resolved.

## Slide types

cover, agenda, section-divider, executive-summary, problem, solution, timeline,
process, framework, comparison, before-after, data-insight, chart, table,
case-study, quote, team, roadmap, architecture, workflow, risk, recommendation,
conclusion, thank-you, appendix

Each type expects specific structured fields (milestones as `{date, event}`
objects, team members as `{name, role, bio}`, etc.). See
[AGENT.md](powerpoint/agent-harness/harness_pptx/AGENT.md) for the complete
field reference.

## Themes

Seven built-in themes selectable by name: `corporate`, `academic`, `startup`,
`consulting`, `technical`, `minimal`, `dark`.

Each theme defines semantic color tokens (`primary`, `accent`, `background`,
`text`, `muted`), font stacks, and spacing scales. Slide templates consume
tokens, not raw hex values.

```python
from harness_pptx.themes.base import ThemeRegistry

themes = ThemeRegistry()
theme = themes.get("dark")
print(theme.colors)   # primary, accent, background, text, muted
```

## CLI

The `cli-anything-powerpoint` entry point gives direct control over PowerPoint:

```bash
cli-anything-powerpoint new --output deck.pptx
cli-anything-powerpoint open deck.pptx --add-slide --add-text "Hello"
cli-anything-powerpoint save-as deck.pptx --format pdf
```

## Project structure

```
harness-pptx-os/
├── README.md
├── powerpoint/
│   └── agent-harness/
│       ├── setup.py
│       ├── skills/                    # Claude Code skills
│       ├── harness_pptx/              # Core package
│       │   ├── pipeline.py            # DeckPipeline orchestrator
│       │   ├── AGENT.md               # AI agent instructions
│       │   ├── models/                # Pydantic data contracts
│       │   ├── content/               # LLM-driven text understanding
│       │   ├── slide_types/           # 25 slide templates
│       │   ├── themes/                # Design tokens and presets
│       │   ├── backends/              # AppleScript + PPTX XML backends
│       │   ├── layout/                # Declarative layout engine
│       │   ├── renderer/              # SceneGraph → PowerPoint
│       │   ├── qa/                    # Automated quality checks
│       │   ├── repair/                # Auto-fix pipeline
│       │   ├── cli/                   # CLI commands and workspace
│       │   ├── narrative/             # Story frameworks
│       │   ├── stability/             # Transaction, retry, logging
│       │   ├── vision/                # Vision API client
│       │   └── memory/                # Agent memory persistence
│       └── cli_anything/              # macOS CLI remote-control
│           └── powerpoint/
```

## For AI agents

This project is designed for AI agents to generate presentations programmatically.
Read [AGENT.md](powerpoint/agent-harness/harness_pptx/AGENT.md) before using the
pipeline. Key constraints:

- Every body slide needs at least 3 content items
- Titles must state the message, not label the topic ("30% Readmission Reduction
  via Predictive Analytics", not "Results")
- Structured fields (milestones, steps, members) must be arrays of objects, not
  flat strings
- Font sizes below 12pt are rejected by QA

## License

MIT

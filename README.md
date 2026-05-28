# Harness PPTX

PowerPoint automation for AI agents — cross-platform (macOS, Windows, Linux).
Drives Microsoft PowerPoint via AppleScript on macOS, and uses python-pptx for
native OOXML manipulation on all platforms. Create, edit, and export `.pptx`
files from structured slide specifications.

## Installation

Python 3.10+ is required. On macOS, Microsoft PowerPoint provides the highest
fidelity; on Windows and Linux, no Office installation is needed.

```bash
git clone https://github.com/SolarAscent/harness-pptx-os.git
cd harness-pptx-os
pip install -e powerpoint/agent-harness
```

**macOS** — activate the bundled virtual environment for AppleScript support:

```bash
source powerpoint/.venv-office-harness/bin/activate
```

**Windows** — install pywin32 for COM automation (native PowerPoint control):

```bash
pip install pywin32
```

**Linux** — or cross-platform without Office — install python-pptx:

```bash
pip install python-pptx
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
6. **PowerPoint Renderer** — AppleScript backend (macOS, native shapes + text) or
   python-pptx backend (cross-platform, native tables + charts)

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

## Backends

The harness auto-detects the platform and selects the best backend:

| Platform | Backend | Requirements |
|----------|---------|--------------|
| **macOS** | AppleScript | Microsoft PowerPoint |
| **Windows** | COM (pywin32) | Microsoft PowerPoint + `pip install pywin32` |
| **Windows (no Office)** | python-pptx | `pip install python-pptx` |
| **Linux** | python-pptx | `pip install python-pptx` |

To select a backend explicitly:

```python
from harness_pptx.backends.registry import get_backend

backend = get_backend("com")        # Windows COM (native PowerPoint)
backend = get_backend("applescript") # macOS AppleScript
backend = get_backend("pptx-xml")   # cross-platform (no Office needed)
```

The COM backend provides native tables, native charts, shape grouping, and slide-level PNG export — same fidelity as the macOS AppleScript backend.

## Themes

Eight built-in themes: `corporate`, `academic`, `academic_purple`, `startup`,
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
│       │   ├── backends/              # AppleScript + COM + python-pptx backends
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

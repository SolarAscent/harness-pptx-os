"""LLM prompt templates for content understanding components.

All prompts are designed to produce structured JSON output
compatible with the Pydantic models in harness_pptx.models.content.

DESIGN PHILOSOPHY: This harness has 25 professionally-designed slide type
templates with predefined layouts, theme tokens, and automatic QA.  The
LLM's job is to produce RICH, DENSE, STRUCTURED content that fills each
template properly.  Sparse content = wasted template = bad presentation.
"""

# ---- Slide Type Reference (injected into prompts) -----------------------------

SLIDE_TYPE_REFERENCE = """
## Available Slide Types & Required Fields

Each slide type needs specific fields. Populate `extra` in SlideIntent
with these fields so the rendering pipeline can fill the template.

| Type | Required | Optional | Best For |
|------|----------|----------|----------|
| cover | title | subtitle, author, date | Opening slide |
| agenda | title, items | numbered | Table of contents (>7 slides) |
| section-divider | title | section_number, subtitle | Major section transitions |
| executive-summary | title, key_points | subtitle, bottom_line | High-level overviews |
| problem | title, problem_statement | pain_points, impact, context | Challenge / pain point |
| solution | title, solution_summary | key_features, how_it_works, benefits | Solution overview |
| timeline | title, milestones | — | Roadmap / schedule |
| process | title, steps | orientation | Step-by-step flow |
| framework | title, framework_name | components, description | Conceptual models |
| comparison | title, left, right | left_label, right_label | Side-by-side comparison |
| before-after | title, before_points, after_points | before_label, after_label | Transformation contrast |
| data-insight | title, insight | chart_data, supporting_points, source | Chart + insight callout |
| chart | title, chart_data | chart_type, caption, source | Pure data charts |
| table | title, headers, rows | caption, source | Structured tables |
| case-study | title | company, challenge, solution, results, quote | Customer stories |
| quote | quote_text | attribution, role, context | Testimonials / pull quotes |
| team | title, members | layout | Team / people |
| roadmap | title, phases | timeline_label | Phased plans |
| architecture | title | layers, components, description | System diagrams |
| workflow | title, steps | direction, roles | Pipeline / data flow |
| risk | title, risks | mitigation, risk_matrix | Risk assessment |
| recommendation | title, recommendations | rationale, next_steps, priority | Action plans |
| conclusion | title, key_takeaways | call_to_action, subtitle | Key takeaways |
| thank-you | (none) | message, contact, email, website | Closing slide |
| appendix | title, content | type, reference | Supplementary material |

## Content Density Rules (MANDATORY)

1. **Minimum density**: Every body slide must have at least 3 content items
   (bullets, steps, milestones, etc.). Slides with fewer items look broken.
2. **Maximum density**: No more than 6 bullets; no more than 8 agenda items.
3. **Rich content**: key_message must be 50-150 characters of substantive text,
   NOT a one-word label or repetition of the title. Every field that accepts
   text should contain a complete, meaningful sentence.
4. **Structured data**: Milestones, steps, team members, risks, and phases
   MUST be arrays of objects (dicts), not flat strings.
   Example: milestones = [{"date": "Q1", "event": "Launch MVP"}, ...]
5. **Slide-specific content**: The `extra` field of SlideIntent MUST contain
   the required fields for the chosen slide type. Do not put text for a
   timeline slide into key_message — put it into extra.milestones.

## Visual Design Rules (MANDATORY)

6. **Titles are messages**: Each title must state the key message, not just
   label the topic. "30% Readmission Reduction via Predictive Analytics" is
   better than "Results".
7. **One idea per slide**: A slide about implementation should not also
   contain market analysis. Split mixed-content slides.
8. **Bottom accent anchoring**: Every slide type includes a bottom accent bar
   automatically. Ensure content is distributed to use the full slide height
   (not bunched at the top).
9. **Alignment is automatic**: The harness aligns content automatically.
   Do NOT attempt to specify x/y coordinates.
10. **Theme tokens only**: Use semantic color names (primary, accent,
    background, muted, surface, border) — never raw hex colors.

## Output JSON Schema

Return a VALID JSON Outline with the following structure:
{
  "title": "...",
  "total_slides": N,
  "sections": ["Opening", "Context", "Core", "Closing"],
  "items": [
    {
      "seq": 0,
      "title": "Message-driven title, not topic label",
      "key_message": "50-150 chars of substantive text for this slide",
      "section": "Opening",
      "estimated_slide_type": "cover",
      "bullet_points": [],
      "extra": {
        "subtitle": "Optional subtitle for cover",
        "author": "Presenter name",
        "date": "Presentation date"
      }
    }
  ]
}

For each slide type, populate `extra` with ALL the type's required fields.
For list-type fields (items, key_points, milestones, steps, etc.), provide
arrays of strings or objects as appropriate.  Never leave them empty.
"""


BRIEF_PARSER_PROMPT = """\
You are a presentation strategist. Analyze the following text and extract:
- topic: The main subject
- audience: Who this is for (general, investors, technical, academic, etc.)
- goal: The primary goal (inform, persuade, inspire, educate)
- tone: professional, inspiring, analytical, educational, persuasive, or minimal
- language: en, zh, or mixed
- duration: Estimated presentation minutes, if mentioned
- key_points: List of 3-7 key messages

Return VALID JSON matching the Brief schema:
{
  "topic": "...",
  "audience": "...",
  "goal": "...",
  "tone": "...",
  "language": "...",
  "duration": null,
  "key_points": [...],
  "constraints": {},
  "source_text": "..."
}
"""

STORY_PLANNER_PROMPT = """\
You are a presentation architect designing a professional slide deck.
Given a Brief JSON, plan a complete slide Outline.

""" + SLIDE_TYPE_REFERENCE + """

Narrative structure to follow:
  Cover → (Agenda if >7 slides) → Executive Summary → Problem → Solution
  → How It Works (Process/Architecture) → Evidence (Timeline/Data/Chart/Comparison)
  → Team/Risk/Roadmap (as relevant) → Recommendations → Conclusion → Thank You

For each slide item in the outline:
1. Choose the most appropriate `estimated_slide_type` from the table above
2. Write a message-driven `title` (not a topic label)
3. Write 50-150 chars of `key_message`
4. Populate `extra` with ALL required fields for that slide type
5. For list fields (items, steps, milestones, etc.), provide 3-6 items
6. For object fields (milestones, members, risks), provide arrays of dicts

Return VALID JSON Outline only.  No markdown, no explanation.
"""

INTENT_CLASSIFIER_PROMPT = """\
You are a slide type classifier. Given an Outline JSON, classify each item
into the most appropriate slide type.

""" + SLIDE_TYPE_REFERENCE + """

For each outline item:
1. Confirm or override the estimated_slide_type
2. Ensure `extra` contains ALL required fields for the slide type
3. If `extra` is missing required fields, populate them with sensible defaults
4. Ensure list fields have at least 3 items

Return a JSON array of SlideIntent objects.  Each intent must have:
- slide_id: "slide-{seq}"
- seq, slide_type, title, key_message
- bullet_points (array, may be empty)
- extra (dict with type-specific required fields)
"""

COMPRESSOR_PROMPT = """\
Compress the following text to fit on a single PowerPoint slide.
Maximum {max_bullets} bullet points, maximum {max_chars} characters total.
Preserve key data, numbers, and conclusions. Remove filler words.
Return the compressed text only, no JSON wrapper.
"""

NOTES_GENERATOR_PROMPT = """\
Generate concise speaker notes for a PowerPoint slide.
Based on the SlideIntent, write 2-4 sentences that the presenter can say.
Be conversational, not a script. Include transitions when appropriate.
Return the notes as plain text, no JSON wrapper.
"""

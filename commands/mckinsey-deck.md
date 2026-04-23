---
description: Build a McKinsey-style PPTX deck from a brief. Delegates to the mckinsey-slide-agent which picks the right template for each slide, explains its choice, and produces a real .pptx file.
argument-hint: <one-paragraph brief describing the deck you want>
---

You have been asked to build a McKinsey-style deck.

**Brief from the user:**

$ARGUMENTS

Delegate this to the `mckinsey-slide-agent` subagent. Pass the brief verbatim
and let the agent:

1. Plan the slide arc
2. Pick a template per slide and defend each choice
3. Generate the Python build script under `output/`
4. Build the `.pptx`
5. Render PNG previews if `soffice` / `pdftoppm` are available
6. Report back with the output path, per-slide rationales, and any caveats

If the brief is in Korean, the agent will respond in Korean and use the
Apple SD Gothic Neo theme.

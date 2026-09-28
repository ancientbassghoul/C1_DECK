# Raycast Challenge v3 — Deck-Building Instructions

## Scope and authority

- Build the simplified Raycast Challenge presentation described in `raycast_challenge_deck_for_generator_v3.md`.
- The v3 blueprint is authoritative for slide order, titles, on-slide text, visuals, and speaker notes.
- `raycast_challenge_deck_for_generator_v2.md` is historical context only. Do not restore its act structure, jokes, mascot, speech bubbles, or omitted deep dives unless v3 explicitly requests them.
- The presentation contains Slides 1–19, followed by Appendix A1, followed by the simple closing Slide 20.
- Do not merge slides. Do not silently omit, summarize, or rewrite blueprint text.

## Critical file-safety rule

- `presentation.pptx` is read-only reference material. Do not modify, overwrite, rename, delete, or append slides to it.
- Any future builder must write to a different output file, preferably `presentation_v3.pptx`.
- Any future builder code must live in a new file, preferably `build_raycast_deck_v3.py`; do not overwrite the existing builder merely to create v3.
- Before and after a build, verify that `presentation.pptx` is unchanged.

## Presentation direction

This is a concise technical case study, not the earlier engineering-war-story version. Its narrative is:

**Task → difficulty → brief failed approaches → manual diagnostic baseline → architectural solution → key debugging discoveries → visual proof → numerical proof → appendix → clean close.**

Keep the tone direct, technical, and evidence-led.

- No jokes.
- No mascot.
- No speech bubbles or mascot animations.
- No persistent act/subway rail.
- No act-opening slides.
- Do not add decorative UI, fake dashboards, or ornamental icons.

## Design system to retain

### Palette

| Role | Color | HEX | Usage |
|---|---|---:|---|
| Primary | Midnight Telemetry | `#101827` | Main dark background, headings, technical structure |
| Secondary | Target Green | `#27E38D` | Working, trusted, retained, successful |
| Accent | Failure Coral | `#FF5B5B` | Failed, regressed, warning, known defect |

White, warm off-white, and neutral gray may support readability but carry no narrative meaning. Green and coral must never be the only distinction; reinforce meaning with labels, line treatment, checks, crosses, or other visible structure.

### Typography

- Headers: Bahnschrift SemiBold
- Body: Aptos
- Code, filenames, flags, and log text: Cascadia Mono
- Deck title: 54–60 pt
- Slide title: 36–42 pt
- Major result/callout: 24–32 pt
- Body: 18–22 pt
- Code: 17–22 pt

### Layout language

Use the smallest, clearest visual structure for each claim:

- Cinematic Evidence Frame for footage, screenshots, point clouds, and proof artifacts.
- Split Diagnosis for before/after or approach/failure comparisons.
- Pipeline Spine for multi-step technical flows.
- A simple two-column comparison for Slide 4.
- A dominant heatmap with restrained numeric callouts for Slide 19.

Prefer one primary claim and approximately 3–5 visual units per slide. Avoid dense card grids, excessive pills/badges, and dashboard styling. A sequence of 3–6 actions should normally become a pipeline. A comparison should read spatially before it is read as text.

## Text integrity

- Include every title, kicker/subtitle, body item, visual instruction, and speaker note specified in v3.
- Every dash under `Body:` is a distinct text item unless the blueprint explicitly defines a table or another structure.
- Do not invent additional technical claims or performance claims.
- Do not promote speaker-note material into on-slide text unless the blueprint says to do so.
- Minor mechanical corrections such as filename escaping or typographic punctuation are allowed only when meaning is unchanged.

## Visual and media handling

- Use the visual specified for each slide when it exists.
- Preserve the visual's aspect ratio; crop intentionally rather than stretching.
- For video slides, use the specified poster frame and video asset. Keep the poster visible if playback is unavailable.
- Slide 3 references source images outside `visuals`; if they are unavailable at build time, use a clearly labeled placeholder and report the missing paths.
- Slide 18 must use `results_slide/How_we_measure_Visual.png` as its measurement-method evidence.
- Slide 19 must use `results_slide/score_pairs_matrix.png` as the dominant artifact.
- Do not substitute the obsolete rough `6 / 6 / 1` result summary for the measured 650-cross-check evaluation in `results_slide`.
- If a specified asset is missing, insert a clearly labeled placeholder rather than silently changing the intended content.

## Results-slide integrity

The authoritative measured results are in:

- `results_slide/score_pairs_summary.txt`
- `results_slide/score_summary.txt`
- `results_slide/score_pairs_matrix.png`

`results_slide/Results_summary_ELI5.md` is optional explanatory context only. Do not depend on its prose or copy it into the deck; the two score summaries and matrix are the authoritative sources.

For the main result, preserve these distinctions:

- 7 ground features, 62 manual marks, and 650 directed source/target cross-checks.
- 649 checks project into-frame; one projects out-of-frame; there are no ray misses.
- Median view error across the pairwise evaluation is 4.05 m.
- Median image error is 38.9 px; 48/649 in-frame checks, approximately 7%, are within 10 px.
- The first and fresh picking sessions produced approximately 4.2 m and 4.0 m respectively, showing repeatability.
- Best source frames are `04709` at 2.67 m and `12035` at 2.77 m; worst is `04569` at 8.08 m.
- Frame `05934` has a known physically invalid camera solution below the ground with flipped orientation.
- Pixel error is range-dependent and favors distant, zoomed-out frames. The metre-based view error is the headline comparison metric.

Do not claim that the original ≤10 px target was achieved across the dataset. The intended claim is that the system produces a stable, measurable cross-frame solution that generally finds the correct area despite the hostile imagery, with known error and a clearly identified failure mode.

## Build validation for future implementation

Before delivery, verify:

- Slides 1–19, Appendix A1, and closing Slide 20 are present and ordered exactly as v3.
- All specified text is present.
- All supplied images, posters, and videos are used on their intended slides.
- Slide 19's heatmap remains legible at presentation scale.
- The output opens successfully in PowerPoint.
- `presentation.pptx` remains unchanged.

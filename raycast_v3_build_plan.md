# Raycast Challenge v3 — Production Plan

## 1. Source of Truth

Create `raycast_v3_build_plan.md` as the durable production contract containing this plan.

Authority order:

1. `AGENTS.md` — safety, design, and validation rules.
2. `raycast_v3_build_plan.md` — batching and production workflow.
3. `raycast_challenge_deck_for_generator_v3.md` — exact slide content, order, assets, and notes.
4. `raycast_challenge_deck_for_generator_v2.md` — historical context only.

The finished deck contains Slides 1–19, Appendix A1, then closing Slide 20: 21 physical slides total.

`presentation.pptx` remains read-only. All work uses `build_raycast_deck_v3.py` and separate v3 output files.

## 2. Locked Visual Language

- Palette: Midnight `#101827`, Target Green `#27E38D`, Failure Coral `#FF5B5B`, supported by warm off-white and neutral gray.
- Typography: Bahnschrift SemiBold headings, Aptos body, Cascadia Mono technical labels.
- Light slides use warm off-white, dark headings, and one continuous neutral divider.
- Dark slides use Midnight backgrounds with restrained green/coral signals.
- No act rail, mascot, speech bubbles, jokes, decorative icons, fake dashboards, or short green bars resembling progress indicators.
- Use green/coral only for semantic meaning; reinforce them with labels, crosses, checks, or structure.
- Favor one dominant artifact, approximately 3–5 visual units, generous margins, and minimal containers.
- Preserve image aspect ratios. Use deliberate cover crops for cinematic imagery and contain fitting for evidence that must remain complete.
- Videos are embedded with their supplied poster frames and configured for click-to-play.
- Every blueprint body dash remains a distinct text item. Speaker-note content stays in notes.

Locked proof references:

- Slide 1: full-bleed hero image with the previous deck’s top title panel; no old rail or mascot.
- Slide 4: light split-diagnosis comparison, neutral divider, clean rounded columns, coral crosses, no overlapping-shape seam.
- Slide 19: heatmap on the left, one Midnight evidence strip on the right, bottom takeaway, and known-failure annotation.

Layout direction by batch:

- Slides 1–4: cinematic opening; visual-plus-four-row assignment; sharp/blurred split challenge; locked failure comparison.
- Slides 5–8: dominant picker video with four compact feature rows; scorer evidence with one explanatory claim; cinematic seven-frame proof; trusted-core visual for the failed extension.
- Slides 9–13: MASt3R pivot; GPU decision evidence; cloud-to-local pipeline; drift warning; Qwen anchor-discovery flow.
- Slides 14–17: Blender debug evidence; hard-black versus feathered-gray split; terrain-aware ray-cast evidence; nearly full-bleed final demonstration video.
- Slides 18–20 and A1: measurement UI screenshot with compact methodology rail; locked heatmap result; appendix experiment; minimal Midnight closing slide.

Slide 18 specifically uses `results_slide/How_we_measure_Visual.png`. Slide 19 uses `results_slide/score_pairs_matrix.png`.

## 3. Stable Batch Contract

| Batch | Slides | Narrative purpose |
|---|---|---|
| Batch 0 — Design Proof | 1, 4, 19 | Establish and lock cinematic, comparison, and evidence layouts |
| Batch 1 — Challenge | 1–4 | Task, assignment, hostile imagery, failed approaches |
| Batch 2 — Manual Baseline | 5–8 | Picker, scorer, seven-frame proof, unsuccessful extension |
| Batch 3 — New Architecture | 9–13 | MASt3R, GPU/cloud workflow, drift, Qwen anchor |
| Batch 4 — Debugging and Demo | 14–17 | Visual debugging, attention magnets, terrain, working demo |
| Batch 5 — Evidence and Close | 18, 19, A1, 20 | Measurement method, quantified result, appendix, clean close |

When asked to “build Batch N”:

1. Implement only that batch’s new slides and any explicitly requested corrections.
2. Regenerate a cumulative deck containing every approved slide through that batch in final order.
3. Render all new or changed slides at 1600×900.
4. Produce a cumulative contact sheet to check rhythm and consistency.
5. Stop for review; do not begin the next batch until explicitly requested.

Corrections remain part of the active batch. Systemic feedback updates shared helpers and all previously completed slides. Local feedback changes only the named slides.

Review outputs:

- `presentation_v3_working.pptx` — latest cumulative deck.
- `review/batch_0N/presentation_v3_batch_0N.pptx` — stable batch snapshot.
- `review/batch_0N/Slide_XX.png` — rendered new or corrected slides.
- `review/batch_0N/contact_sheet.png` — cumulative visual review.

After Batch 5 approval, produce `presentation_v3.pptx`.

## 4. Builder Implementation

Evolve `build_raycast_deck_v3.py` into one deterministic builder:

- Shared constants and helpers own palette, typography, titles, dividers, image fitting, video embedding, notes, and evidence layouts.
- Each final slide has a dedicated builder function.
- An ordered registry represents `1–19, A1, 20`.
- A batch map controls cumulative generation.
- CLI contract:
  - `--proof` builds Slides 1, 4, and 19.
  - `--through-batch N` builds the cumulative deck through Batch N.
  - `--output PATH` selects the non-original output file.
- The builder always creates a fresh presentation; it never loads or appends to `presentation.pptx`.
- Missing assets generate labeled placeholders and a reported warning rather than silent substitution.
- The measured-result values come only from the score summaries and matrix, never the obsolete `6 / 6 / 1` estimate.

## 5. Validation and Acceptance

For every batch:

- Compile the builder and reopen the generated file with `python-pptx`.
- Open and render the deck through desktop PowerPoint.
- Confirm cumulative slide counts: 4, 8, 13, 17, then 21.
- Verify all expected titles, body items, kickers, notes, images, posters, and videos.
- Check text overflow, cropping, font substitution, seams, accidental shadows, and minimum readable type.
- Confirm video poster frames remain visible and embedded videos are click-to-play.
- Inspect new slides individually and the cumulative contact sheet.
- Confirm Slide 18’s screenshot remains readable and Slide 19’s heatmap labels and values remain legible.
- Verify the hash and modification timestamp of `presentation.pptx` remain unchanged.

Assumptions:

- Batch 0’s corrected proof is the visual baseline.
- Existing supplied visuals are authoritative and are not regenerated.
- The original ≤10 px requirement remains part of the assignment framing, while the result is presented honestly as a stable 4.05 m median view error.
- A batch is considered approved only after explicit user confirmation or an explicit request to begin the next batch.

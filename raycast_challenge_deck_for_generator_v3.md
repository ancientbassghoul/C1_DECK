# Raycast Challenge — Simplified Deck Blueprint (v3)

Read `raycast_challenge_deck_for_generator_v2.md` for historical context before building. This v3 file is the authoritative blueprint.

After an informal interview, the presentation was deliberately simplified around five questions:

1. What is the task?
2. Why is it hard?
3. Briefly, what was tested and failed?
4. How was it solved?
5. What visual and numerical evidence demonstrates the result?

No jokes, mascots, speech bubbles, act structure, subway map, or unnecessary deep dives.

## Critical output rule

**Do not touch `presentation.pptx`.** It is read-only reference material. Any future code must write a different file, preferably `presentation_v3.pptx`.

The evidence-led main sequence contains Slides 1–19. Appendix A1 follows Slide 19, and a simple closing Slide 20 follows the appendix.

---

## Slide 1

**Title:** The Raycast Challenge

**Kicker/subtitle:** Pick a pixel in one drone frame. Find the same point in twelve others in ≤10 pixels accuracy.

**Visual:** `visuals\Raycast_Slide_1_Visual.png`

## Slide 2

**Title:** The Assignment

**Body:**

- A sequence of drone FPV frames, one stationary van, one agricultural field
- Several frames badly motion/defocus-blurred
- Deliverable: source code, a methodology writeup, and *proof* — not just a claim — of ≤10px reprojection accuracy
- Complete architectural freedom. "Ingenuity and a tinkerer mindset," officially encouraged.

**Visual:** `visuals\Raycast_Slide_2_Visual.png`

## Slide 3

**Title:** The Challenges

**Body:**

- The field is highly repetitive — one patch of dirt looks like every other patch of dirt
- Several frames are so blurred that even a human struggles to place a feature confidently
- GPS telemetry on the drones is off by multiple metres — you can't just trust the numbers you're handed

**Visual:** Side-by-side crop of a sharp frame and one of the worst-blurred frames, showing the same rough patch of ground.

- Sharp frame: `D:\EXTEND\C1_IMGS\2026-02-15_16-25-03_12035.png`
- Worst-blurred frame: `D:\EXTEND\C1_IMGS\2026-02-15_16-25-03_10305.png`

## Slide 4

**Title:** What Had Failed

**Body:** Present this as a two-column comparison: **Solution | Why it failed**.

| Solution | Why it failed |
|---|---|
| Homography | Lack of good joint-plane corner features; distracting van features |
| GeoCalib → PyCeres | Lack of good distinguishable features |

**Visual:** A clean split-diagnosis layout using the two-column comparison above. The failed solutions should be structurally marked in Failure Coral; do not add a decorative image.

## Slide 5

**Title:** The Manual Correspondence Picker

**Body:**

- Multi-frame grid, independent zoom/pan, adjustable marker size
- Four feature types, cycled with a key: `ground`, `roof`, `wheel_axis`, `roof_edge` — the last two are *pairs* with known real-world distances (wheelbase, van width) baked in as hard constraints
- Right-click removes a single point; `d` deletes a whole correspondence or just one frame's mark, depending on context; `[`/`]` step through saved correspondences filtered by type
- A quit-safety net: unsaved changes require a second keypress to actually discard

**Visual:** Video: `visuals\Raycast_Slide_8_Visual_RETIMED.mp4`

**Poster frame:** `visuals\Raycast_Slide_8_Poster.jpg`

**Speaker notes:** The justification of this tool is to test how PyCeres performs when handed good feature matching.

## Slide 6

**Title:** The Scorer

**Body:** The scorer was built to assess manual feature-matching quality. It compares crops with CLIP and acts as a warning signal, not geometric proof: a low score can mean either a wrong pick or a correct feature seen through blur or a very different angle.

**Visual:** `visuals\Raycast_Slide_9_Visual.png`

## Slide 7

**Title:** "Manual Feature Matching Works Great on 7 Frames"

**Body:** Known van geometry (wheel axis, roof edges) as hard anchors + hand-picked ground correspondences → a genuinely good calibration on the 7 clearest frames. So PyCeres is good, and our problem is the feature matching.

**Visual:** Video: `visuals\Raycast_Slide_10_Visual.mp4`

**Poster frame:** `visuals\Raycast_Slide_10_Poster.jpg`

## Slide 8

**Title:** The Other Six

**Body:** Remembering past days of a base manual calibration – optimized by automatic processes, I was hoping to be able to treat the 7 manually solved frames as core, then add other frames to be solved automatically. It didn't succeed.

**Visual:** `visuals\Raycast_Slide_11_Visual.png`

## Slide 9

**Title:** The Solution

**Body:** The old stack just couldn't crack it, and something smarter was needed. MASt3R's dense reconstruction could handle the low-texture, repetitive ground the old feature matchers choked on.

**Visual:** `visuals\Raycast_Slide_20_NEW_Visual.png`

**Speaker notes:** Mention that this magic requires better hardware.

## Slide 10

**Title:** Step Zero: Get a GPU

**Body:** RunPod, chosen over Google Colab specifically because Colab's runtimes wipe on timeout — meaning MASt3R's CUDA/C++ kernels would need recompiling from scratch every session, versus RunPod's persistent volume storage compiling once and staying ready.

**Visual:** `visuals\Raycast_Slide_21_Visual.png`

## Slide 11

**Title:** Compute Headless in the Cloud, Test Locally

**Body:** RunPod has no display; the interactive viewer needs one. `--export-solve` runs the full pipeline remotely and serializes camera poses + terrain point clouds to disk. Those files travel to the local machine, where `--import-solve` skips OCR/MASt3R/Ceres entirely and opens the viewer straight against the already-solved cameras.

**Visual:** `visuals\Raycast_Slide_23_Visual.png`

## Slide 12

**Title:** MASt3R's Achilles' Heel

**Body:** MASt3R has one weakness — it drifts. Qwen finds a shared anchor in the images to constrain that drift.

**Visual:** `visuals\Raycast_Slide_23.5_Visual.png`

## Slide 13

**Title:** Qwen's Job: Find the One Thing Every Frame Shares

**Body:** Per frame: list every man-made object, bbox + label. Across all frames: consolidate labels, rank by how many frames contain each one. Among the top-coverage candidates, a *second* Qwen call — reasoning over label text alone — picks the single best fixed 3D reference point. Output: that object's pixel centroid in every frame it appears in, weighted by CLIP confidence.

**Visual:** `visuals\Raycast_Slide_24_Visual.png`

## Slide 14

**Title:** `--export-mesh`: Visual Debugging

**Body:** Even with the right tools, the reconstruction wasn't good. So it was time to see things with our eyes: a debug tool that dumps the raw MASt3R reconstruction + solved cameras (and, once solved, the Ceres-refined + Sim(3)-aligned version) as real `.glb` scenes, importable straight into Blender.

**Visual:** `visuals\Raycast_Slide_28_Poster.jpg`

## Slide 15

**Title:** Attention Magnets

**Body:** Hard black HUD fills acted as infinite-contrast "attention magnets" for MASt3R's vision transformer, warping its initial pose estimates. Fixed with a feathered gray fill — Gaussian-blurred, no hard edge.

**Visual:** Two images side by side.

- **Hard Black Fill:** `visuals\Raycast_Slide_29_01_Visual.png`
- **Feathered Grey Fill:** `visuals\Raycast_Slide_29_02_Visual.png`

## Slide 16

**Title:** Flat Earth

**Body:** Replaced the flat `z = 0` ground-plane assumption with an interpolated height field built from the solved 3D points themselves — ray-casting now iterates to find where a ray actually meets the real, bumpy terrain, falling back to flat only if that doesn't converge.

**Visual:** `visuals\Raycast_Slide_32_Visual.png`

## Slide 17

**Title:** See It In Action

**Visual:** Video: `visuals\Raycast_Slide_35_Visual.mp4`

**Poster frame:** `visuals\Raycast_Slide_35_Poster.jpg`

## Slide 18

**Title:** How the Result Was Measured

**Body:**

- 7 ground features, manually marked in up to 13 frames: 62 judged marks in total
- Every marked frame was tested as the starting view against every other marked frame
- 650 directed source-to-target cross-checks: 649 projected into-frame, 1 projected out-of-frame, 0 ray misses
- The score is the distance between the system's projected point and the manually marked location
- Metres are the headline metric because pixel error favors distant, zoomed-out frames

**Visual:** `results_slide\How_we_measure_Visual.png` as the dominant evidence artifact. It shows the evaluation UI across the frame grid: the manually marked truth, the system's projected point, and the pixel/view/ground error readout. Preserve the full screenshot so the repeated cross-frame measurement is immediately visible; crop only the outer application chrome if necessary for legibility.

**Speaker notes:** The manual marks are imperfect, especially in blurred distant frames. Errors below roughly one metre can be difficult to distinguish from click uncertainty. The result should therefore be presented as a measured engineering evaluation, not false precision.

## Slide 19

**Title:** Measured Result: About One Van Length

**Kicker/subtitle:** 650 cross-checks across all 13 frames

**Body:**

- **4.05 m median view error** — roughly one van length
- **38.9 px median image error** — 48 of 649 in-frame checks, about 7%, landed within 10 px
- **Repeatable:** approximately 4.2 m in the first picking session and 4.0 m in the fresh session
- **Best source:** `04709` at 2.67 m; **worst source:** `04569` at 8.08 m

**Visual:** `results_slide\score_pairs_matrix.png` as the dominant evidence artifact, occupying roughly the left 70–75% of the canvas. Keep the row/column labels, values, median margins, and color scale legible. Place the four body results in one restrained vertical evidence strip on the right, not as separate dashboard cards.

**Bottom takeaway:** The system consistently finds the correct area of the field across extremely difficult imagery, with a stable measured error and a visible frame-by-frame failure pattern.

**Known failure annotation:** Frame `05934` has a physically invalid camera solution below the ground with flipped orientation. Its anchor is usable, but its camera axes are not.

**Speaker notes:** Pixel error varies strongly with distance and framing, so the metre-based view error is the fairer headline comparison. `12035` is one of the best frames to start from but one of the hardest to land in because its camera looks down much more steeply than the others. The result is not the original ≤10 px target across the dataset; it is a stable, quantified solution over all 13 hostile frames.

---

# Appendix

## Appendix A1

**Title:** Better Is the Enemy of Good

**Body:** An incremental two-stage Ceres solve (lock down the good frames first, register the blurry ones motion-only against that structure). Then, one level deeper: splitting MASt3R's *own* reconstruction the same way — a clean backbone built only from good frames, with each bad frame individually resectioned against it so it can never contaminate another bad frame or the backbone itself. Unfortunately, **neither produced better real-world output.**

**Visual:** `visuals\Raycast_Slide_33_Visual.png`

**Speaker notes:** Worth mentioning that this whole effort was across half a day. As feedback was already good, we decided to move on to the next challenge, knowing we can continue fighting it another day.

---

## Slide 20

**Title:** Thank You. Questions?

**Visual:** A simple closing slide using the established Midnight Telemetry background and restrained Target Green accent. No body text, metrics, appendix material, mascot, or decorative illustration.

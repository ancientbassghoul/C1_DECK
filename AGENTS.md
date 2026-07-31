# We are building a 40 slides deck.

# This is the design system:

## Presentation structure

The deck is an engineering war story with eleven acts:

| Act                          | Slides | Narrative job                                                                  |
| ---------------------------- | -----: | ------------------------------------------------------------------------------ |
| 0 — The Setup                |    1–3 | Establish the challenge, evidence bar, and hostile dataset                     |
| I — Bring-Up                 |    4–6 | Show the surprisingly mature first pipeline—and its first wrong assumptions    |
| II — By Hand, For Now        |   7–11 | Build a trusted manual baseline and discover the recurring “good core” pattern |
| III — Asking for Directions  |  12–18 | Pause, report honestly, seek guidance, and reject an attractive diversion      |
| IV — Starting Over           |  19–23 | Perform the architectural rewrite and establish cloud/local development        |
| V — How the New Brain Thinks |  24–25 | Explain the new open-world anchor-discovery pipeline                           |
| VI — The Bug Hunt            |  26–29 | Reveal model-integration and dataset-calibration failures                      |
| VII — Silent Failures        |  30–32 | Expose bugs that produced plausible but incorrect output                       |
| VIII — Tuning the Machine    |  33–36 | Replace brittle thresholds and assumptions with adaptive, observable methods   |
| IX — The Twist               |  37–38 | Sound reasoning produces worse empirical results                               |
| X — Where It Landed          |  39–40 | Present the autonomous final pipeline and working outcome                      |

The narrative progression is:

**Hard challenge → manual truth → external guidance → architectural reset → invisible failures → better observability → empirical humility → autonomous result.**

The central recurring idea is:

> **Trust a good core, then extend outward carefully.**

It appears first in manual matching, returns in confidence weighting, and culminates in the two-stage MASt3R strategy. Slide 39 should visually bookend Slide 4’s pipeline diagram.

## Core themes

* Telemetry, model confidence, and “reasonable” defaults are all uncertain sensor readings.
* Debugging tools are part of the architecture—not secondary utilities.
* A silent fallback is often more dangerous than a crash.
* General technical advice must be recalibrated against the actual dataset.
* Elegant reasoning does not outrank empirical output.
* Engineering judgment includes deleting work, rejecting pivots, and knowing when to stop.

## Confirmed Design System

### Three-color palette

| Role      | Color              |       HEX | Usage                                                         |
| --------- | ------------------ | --------: | ------------------------------------------------------------- |
| Primary   | Midnight Telemetry | `#101827` | Main dark background, headings, technical structure           |
| Secondary | Target Green       | `#27E38D` | Working, trusted, retained, successful, current position      |
| Accent    | Failure Coral      | `#FF5B5B` | Removed, broken, regressed, warnings, crossed-out assumptions |

White, warm off-white, and neutral gray may support readability, but they carry no narrative meaning.

Green and red will never be the only distinction: success also uses solid lines/checks; failure uses strike-throughs, breaks, crosses, or dashed treatment.

### Typography

* **Headers:** Bahnschrift SemiBold
  Technical, condensed, and visually related to telemetry/HUD typography.

* **Body:** Aptos
  Highly readable for explanations, annotations, and speaker-facing technical material.

* **Code exception:** Cascadia Mono
  Reserved for filenames, flags, code, configuration values, and log output.

Recommended hierarchy:

* Deck title: 54–60 pt
* Slide title: 36–42 pt
* Callout: 24–28 pt
* Body: 18–22 pt
* Code: 17–22 pt

### Four reusable layout formats

1. **Cinematic Evidence Frame**
   Full-bleed drone footage, screenshot, Blender view, or point cloud. One large claim with restrained annotations, crosshairs, rays, or bounding boxes.

2. **Split Diagnosis**
   Before/after, expected/observed, old/new, deleted/added, or working/failing. The comparison should be visible spatially before the audience reads it.

3. **Pipeline Spine**
   A strong left-to-right technical flow with the current transformation emphasized. Slide 4 establishes the visual grammar; Slide 39 returns to it with the completed architecture.

4. **Trusted Core + Perimeter**
   A solid central backbone with uncertain frames, sensors, or hypotheses extending around it. This becomes the visual signature for the recurring “trust a good core” principle.

Code excerpts, charts, and number lines use the **Cinematic Evidence** format: one dominant artifact rather than a dashboard of small panels.

## Rules for converting text into visuals

* A sequence of 3–6 actions becomes a pipeline, not bullets.
* Two opposing states become a split composition.
* A threshold or distribution becomes a number line or compact chart.
* “Kept versus deleted” becomes solid green versus struck-out red.
* A central trusted set with uncertain additions becomes the core-and-perimeter format.
* A bug is shown through its evidence: code, logs, incorrect bbox, matrix fingerprint, or before/after output.
* Icons must encode concrete meaning. No decorative icon beside every bullet.
* Technical icons use one consistent thin-outline style and retain short labels.
* Avoid dense card grids, pills, badges, and dashboard-like UI styling.
* Limit the audience to one primary claim and approximately 3–5 visual units per slide.
* Use progressive reveals for complex pipelines and comparisons.

## Persistent act navigation

The left rail remains visible throughout:

* All eleven rail labels are always present.
* Only the current act displays its slide-station dots.
* The current station uses Target Green.
* Completed stations are solid but muted.
* Upcoming stations are outlined.
* The full current-act name appears at the upper-left.
* The rail remains structurally identical in dark and light slide modes.

## Visual rhythm

The alternation between dark and bright slides will be deliberate:

* **Dark mode:** act openings, architectural pivots, severe bugs, the twist, and major conclusions.
* **Light mode:** technical explanation, evidence comparison, tooling, and calibration.
* **Full-image mode:** moments where real footage or generated imagery should carry the emotional weight.

The source footage is naturally subdued—blue-gray sky, brown terrain, dark green fields, and extreme white sun—so the green/red signals will remain legible without competing with it.

Speech bubbles and the mascot will always be final click-triggered overlays. The underlying slide will not reserve space or distort its composition for them.

This Design System is confirmed as the visual contract for the deck.

# CRITICAL SCRIPT INTEGRITY RULE:
1. NON-DESTRUCTIVE MODIFICATION: Do NOT rewrite or delete any layout code, text, or image insertion lines (`add_picture`) built for previous slides.
2. FILE EDIT METHOD: Load the existing `presentation.pptx` file using `prs = Presentation('presentation.pptx')` and append the new Act slides to the existing presentation object.
3. PRESERVE ALL ASSETS: Verify that all previously inserted images (e.g., `slide_01_hero.png`) remain untouched in their respective slides before saving the updated `.pptx` file.

# STRICT EXECUTION REQUIREMENT: You are a Python pptx builder.
Your blueprint for the deck is raycast_challenge_deck_for_generator.md
Do NOT summarize, rephrase, omit, merge, or shorten ANY text from the raycast_challenge_deck_for_generator.md file. 
The text is ALREADY slide-ready.

KEEP THESE RULS AT ALL TIMES:

1. VERBATIM TEXT: Every dash (`-`) under `Body:` MUST be its own distinct text item/card.
2. LAYOUTS: Use 2-column grids for slides with 4+ items.
3. SPEECH BUBBLES: Format all "Speech bubble:" lines as styled accent cards in the bottom corner.
4. SUBWAY MAP: Maintain the Act rail widget on the left edge (highlighting Act [X] and current slide dots).
5. KICKERS: Include all "Kicker/subtitle:" text under slide titles.

# VISUALS
Visuals for the slides will be in the "visuals" folder. When creating a slide, if you see a visual for that slide in the visuals folder - use it. If not - put a place holder for the visual in the presentation.


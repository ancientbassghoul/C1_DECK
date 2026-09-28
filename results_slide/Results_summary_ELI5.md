Reprojection accuracy — ELI5 summary

The game. I picked a spot on the ground in one photo. The system worked out where that spot sits in the real world and drew a dot where it should appear in every other photo. Then I marked where the spot really is in each photo. How far apart the dot and my mark are is the error.

What I measured. 7 ground features, each marked in up to 13 photos, 62 marks in all. Every photo was then tested as the starting photo against every other photo. That gives 650 checks.

Headline result: the system is close, but not accurate.
- A typical miss is about 4 m on the ground, roughly one van length (the van is about 5 m long and 1.9 m wide).
- In image terms, a typical miss is about 39 pixels. The goal was 10 px or less, and only 7% of checks met it.
- Only 2% of checks were within half a metre.
- So the system reliably finds the right area of the field, but not the exact spot.

The result repeats. The first picking session gave 4.2 m, and this fresh session gave 4.0 m, with different clicks. So this is a stable property of the solve, not luck or sloppy clicking.

It matters which photo you start from.
- Best starting photos: 04709 (2.7 m) and 12035 (2.8 m).
- Worst starting photo, by far: 04569 (8.1 m). Starting from it sends the dot 13–18 m off in some frames.
- Most other photos land in the 3–5 m range.

Some photos are hard to land in.
- 10671 (6.1 m), 12035 (5.5 m) and 09763 (5.1 m) are where dots land worst.
- 12035 is interesting: it's one of the best photos to start from but one of the worst to land in. Its camera looks down more steeply (about 46–51°, versus 10–18° for the rest). So its own view projects cleanly onto the ground, but points from other views arrive there badly.

Why measure in metres, not only pixels. In a far-away photo, a 4 m miss is only a few pixels. In a close-up, it's dozens of pixels. So pixels make the blurry, distant photos look best, which is backwards. Metres treat every photo fairly.

Why the "ground error" number is bigger (about 6.7 m). The drone sees the ground at a very shallow angle, like looking aistake up or down in the picture becomes a big mistake along the ground. That's also why 04569, the flattest view atabout 10°, shows the huge 17 m ground error.

One clear bug. Photo 05934's camera ends up about 8 m below the ground, looking up at the field. That's physically impossible, and it shows the solve got that camera's height wrong.

Honest caveats.
- I placed my "true" marks while the system's dot was on screen, so my marks may lean toward the system's answer. If anye slightly larger.
- On blurry, far-away frames, my own clicks are uncertain by a metre or more. So errors below about 1 m can't really be told apart from click noise.
One sentence for the slide. Across 650 cross-checks, the solved cameras put a ground point about 4 m (roughly one van length) from where it really is. That's consistent between sessions, best when picking from the steeper, closer views (04709, 12035), and worst from the flattest one (04569). That's well short of the 10-pixel target, which only about 7% of checks met.

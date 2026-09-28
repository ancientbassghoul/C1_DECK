# The Raycast Challenge Deck

This project generates the PowerPoint presentation `presentation.pptx` from
`build_raycast_deck.py`. Slide visuals live in `visuals`, source-frame images
live in `dataset`, and the mascot image lives in `assets`.

The final click-triggered mascot and speech-bubble animations are applied by
`apply_speech_animations.ps1`. This step requires Windows and the desktop
version of Microsoft PowerPoint.

## Rebuild the virtual environment

Run these commands in PowerShell from the project folder:

```powershell
# Close PowerPoint before rebuilding or generating the deck.
Remove-Item -LiteralPath .\venv -Recurse -Force
py -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install python-pptx
```

If PowerShell blocks environment activation, activation is optional. Use the
virtual environment's Python executable directly:

```powershell
.\venv\Scripts\python.exe -m pip install --upgrade pip
.\venv\Scripts\python.exe -m pip install python-pptx
```

## Generate the presentation

With the virtual environment activated:

```powershell
python .\build_raycast_deck.py
```

Without activation:

```powershell
.\venv\Scripts\python.exe .\build_raycast_deck.py
```

The default output is `presentation.pptx`. To write elsewhere:

```powershell
python .\build_raycast_deck.py --output .\output\raycast_deck.pptx
```

For a build without PowerPoint animations:

```powershell
python .\build_raycast_deck.py --skip-animations
```

## Raycast Challenge v3

The approved embedded v3 deck is `presentation_v3.pptx`. The repository version,
`presentation_v3_linked.pptx`, links to these Git LFS-managed videos instead of
duplicating them inside the PowerPoint file:

- `visuals/Raycast_Slide_8_Visual_RETIMED.mp4`
- `visuals/Raycast_Slide_10_Visual.mp4`
- `visuals/Raycast_Slide_35_Visual.mp4`

Keep the linked presentation at the repository root and preserve the `visuals`
folder structure. After cloning, run `git lfs pull` before opening the deck.
PowerPoint playback still requires codecs supported by the presentation machine.

To rebuild the full embedded v3 deck:

```powershell
.\venv\Scripts\python.exe .\build_raycast_deck_v3.py --through-batch 5 --output .\presentation_v3.pptx
```

To regenerate the linked copy without modifying the embedded deck:

```powershell
.\venv\Scripts\python.exe .\create_linked_v3_deck.py
```

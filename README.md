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

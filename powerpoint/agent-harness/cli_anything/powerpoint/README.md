# CLI-Anything PowerPoint

Install:

```bash
python3 -m pip install -e .
```

Usage:

```bash
cli-anything-powerpoint
cli-anything-powerpoint info --json
cli-anything-powerpoint new
cli-anything-powerpoint open ~/Desktop/input.pptx
cli-anything-powerpoint add-title-slide "Title" --subtitle "Subtitle"
cli-anything-powerpoint save-as ~/Desktop/output.pptx
cli-anything-powerpoint export-pdf ~/Desktop/output.pdf
```

Inspection and editing:

```bash
cli-anything-powerpoint list-slides --json
cli-anything-powerpoint list-shapes 1 --json
cli-anything-powerpoint shape-info 1 3 --json
cli-anything-powerpoint set-text 1 3 "Revised copy" --font-size 24 --bold
cli-anything-powerpoint move-shape 1 3 --left 120 --top 90 --width 420 --height 80
cli-anything-powerpoint set-fill 1 3 "#F8F0E8"
cli-anything-powerpoint set-line 1 3 "#AA0000" --weight 1.5
cli-anything-powerpoint z-order 1 3 front
cli-anything-powerpoint delete-shape 1 3
```

The macOS backend uses Microsoft PowerPoint's AppleScript interface. Before each operation it primes Office's sandbox file access by touching the startup disk inside PowerPoint's scripting context, which avoids the recurring Grant File Access prompt for normal local paths after macOS has allowed automation. Windows COM support is represented in the backend selection layer and can be filled in with `win32com.client` when run on Windows.

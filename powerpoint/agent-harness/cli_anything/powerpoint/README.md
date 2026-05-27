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

The macOS backend uses Microsoft PowerPoint's AppleScript interface. Before each operation it primes Office's sandbox file access by touching the startup disk inside PowerPoint's scripting context, which avoids the recurring Grant File Access prompt for normal local paths after macOS has allowed automation. Windows COM support is represented in the backend selection layer and can be filled in with `win32com.client` when run on Windows.

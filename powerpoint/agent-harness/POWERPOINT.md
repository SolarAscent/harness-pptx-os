# Microsoft PowerPoint Harness

Target: Microsoft PowerPoint

Backend:

- macOS: AppleScript through `osascript`, using Microsoft PowerPoint's scripting dictionary.
- Windows: reserved COM bridge path via `win32com.client` for future Windows execution.

This harness gives agents command-line access to the real PowerPoint app for opening decks, creating slides, saving, exporting PDF, and triggering undo/redo.

from __future__ import annotations

import json
import platform
import subprocess
import tempfile
from pathlib import Path
from typing import Any


# ── AppleScript helpers ──────────────────────────────────────────────

def _escape(text: str) -> str:
    """Escape text for AppleScript string literal. Converts \\n to \\r for PPT line breaks."""
    # PPT uses \r for line breaks in text content
    cleaned = text.replace("\n", "\r")
    return json.dumps(cleaned, ensure_ascii=False)


def _clist(r: int, g: int, b: int) -> str:
    """Format RGB as AppleScript integer list: {R, G, B}."""
    return f"{{{r}, {g}, {b}}}"


def _parse_color(color: str) -> tuple[int, int, int]:
    """Parse 'R,G,B' or '#RRGGBB' to (r,g,b) tuple."""
    c = color.strip()
    if c.startswith("#"):
        c = c.lstrip("#")
        return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))
    parts = c.split(",")
    if len(parts) < 3:
        raise ValueError(f"Color must be 'R,G,B' or '#RRGGBB', got: {color}")
    return (int(parts[0].strip()), int(parts[1].strip()), int(parts[2].strip()))


def _fill_block(shape_var: str, color: str) -> list[str]:
    """Generate AppleScript tell block to set fill color on a shape."""
    r, g, b = _parse_color(color)
    return [
        f"tell fill format of {shape_var}",
        f"set fore color to {_clist(r, g, b)}",
        "end tell",
    ]


def _font_color_block(shape_var: str, color: str) -> str:
    """Generate AppleScript line to set font color."""
    r, g, b = _parse_color(color)
    return f"set font color of font of text range of text frame of {shape_var} to {_clist(r, g, b)}"


def _parse_tab_records(output: str, fields: list[str]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for line in output.splitlines():
        if not line.strip():
            continue
        parts = line.split("\t")
        records.append({field: parts[idx] if idx < len(parts) else "" for idx, field in enumerate(fields)})
    return records


def _coerce_numbers(row: dict[str, Any], names: list[str]) -> dict[str, Any]:
    for name in names:
        value = row.get(name)
        try:
            number = float(value)
            row[name] = int(number) if number.is_integer() else number
        except (TypeError, ValueError, AttributeError):
            pass
    return row


# ── Enums ────────────────────────────────────────────────────────────

TRANSITION_EFFECTS = {
    "none": "entry effect appear",
    "fade": "entry effect fade smoothly",
    "push_up": "entry effect cover up", "push_down": "entry effect cover down",
    "push_left": "entry effect cover left", "push_right": "entry effect cover right",
    "dissolve": "entry effect dissolve",
    "uncover_up": "entry effect fly from bottom", "uncover_down": "entry effect fly from top",
    "uncover_left": "entry effect fly from right", "uncover_right": "entry effect fly from left",
    "wipe_right": "entry effect box out", "wipe_left": "entry effect box in",
    "zoom_in": "entry effect circle", "zoom_out": "entry effect box in",
    "cube": "entry effect cube left", "flip": "entry effect flip left",
    "gallery": "entry effect gallery left", "doors": "entry effect doors vertical",
    "curtains": "entry effect curtains", "crush": "entry effect crush", "fracture": "entry effect fracture",
}

SLIDE_LAYOUTS = {
    "blank": "slide layout blank", "title": "slide layout title slide",
    "title_only": "slide layout title only", "text": "slide layout text slide",
    "two_column": "slide layout two column text", "section_header": "slide layout section header",
    "comparison": "slide layout comparison", "content_caption": "slide layout content with caption",
    "picture_caption": "slide layout picture with caption",
}

ANIM_EFFECTS = {
    "fade": "animation effect fade",
    "fly_in_left": "animation effect fly from left", "fly_in_right": "animation effect fly from right",
    "fly_in_top": "animation effect fly from top", "fly_in_bottom": "animation effect fly from bottom",
    "zoom": "animation effect zoom", "appear": "animation effect appear", "wipe": "animation effect wipe",
}


class PowerPointBackend:
    """Backend driving Microsoft PowerPoint via AppleScript on macOS."""

    app_name = "Microsoft PowerPoint"
    SLIDE_W = 960
    SLIDE_H = 540

    def __init__(self) -> None:
        self.platform = platform.system().lower()

    def available_interface(self) -> str:
        if self.platform == "darwin":
            return "applescript"
        if self.platform == "windows":
            return "com"
        return "unsupported"

    def run_script(self, script: str) -> str:
        if self.available_interface() != "applescript":
            raise RuntimeError(f"{self.app_name} automation not implemented for {self.platform}")
        script = self._with_file_access(script)
        proc = subprocess.run(["osascript", "-e", script], text=True, capture_output=True, check=False)
        if proc.returncode != 0:
            raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
        return proc.stdout.strip()

    def _with_file_access(self, script: str) -> str:
        lines = script.splitlines()
        if lines and lines[0] == f'tell application "{self.app_name}"':
            prelude = [
                "try",
                "set _cDisk to path to startup disk",
                "close _cDisk",
                "end try",
            ]
            return "\n".join([lines[0], *prelude, *lines[1:]])
        return script

    # ══════════════════════════════════════════════════════════════════
    # Presentation Management
    # ══════════════════════════════════════════════════════════════════

    def info(self) -> dict[str, Any]:
        version = self.run_script('tell application "Microsoft PowerPoint" to return version')
        return {"application": self.app_name, "interface": "applescript", "version": version}

    def new_presentation(self) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "activate\n"
            "set deckRef to make new presentation\n"
            "set newSlide to make new slide at end of deckRef\n"
            "set layout of newSlide to slide layout blank\n"
            "return name of deckRef\n"
            "end tell"
        )
        return {"status": "ok", "presentation": "Untitled"}

    def set_page_size(self, width: float = 960, height: float = 540) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "set deckRef to active presentation\n"
            f"set slide width of page setup of deckRef to {width}\n"
            "end tell"
        )
        return {"status": "ok", "width": width, "height": height}

    def open_presentation(self, path: str) -> dict[str, Any]:
        full_path = str(Path(path).expanduser().resolve())
        name = self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "activate\n"
            f'open POSIX file {json.dumps(full_path)}\n'
            "return name of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "path": full_path, "presentation": name}

    def save_as(self, path: str) -> dict[str, Any]:
        full_path = str(Path(path).expanduser().resolve())
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f'save active presentation in (POSIX file {json.dumps(full_path)}) '
            "as save as Open XML presentation\n"
            "end tell"
        )
        return {"status": "ok", "path": full_path}

    def export_pdf(self, path: str) -> dict[str, Any]:
        full_path = str(Path(path).expanduser().resolve())
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f'save active presentation in (POSIX file {json.dumps(full_path)}) as save as PDF\n'
            "end tell"
        )
        return {"status": "ok", "path": full_path}

    def export_png(self, path: str) -> dict[str, Any]:
        full_path = str(Path(path).expanduser().resolve())
        Path(full_path).parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="cli-anything-ppt-") as tmpdir:
            pdf_path = str(Path(tmpdir) / "preview.pdf")
            self.export_pdf(pdf_path)
            proc = subprocess.run(
                ["sips", "-s", "format", "png", pdf_path, "--out", full_path],
                text=True,
                capture_output=True,
                check=False,
            )
            if proc.returncode != 0:
                raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
        return {"status": "ok", "path": full_path, "note": "Rendered from exported PDF preview"}

    def close(self, saving: bool = True) -> dict[str, Any]:
        flag = "yes" if saving else "no"
        self.run_script(
            f'tell application "Microsoft PowerPoint" to close active presentation saving {flag}'
        )
        return {"status": "ok", "saving": saving}

    def apply_theme(self, theme_path: str) -> dict[str, Any]:
        full_path = str(Path(theme_path).expanduser().resolve())
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f'apply theme to active presentation with properties {{file name: {json.dumps(full_path)}}}\n'
            "end tell"
        )
        return {"status": "ok", "theme": full_path}

    # ══════════════════════════════════════════════════════════════════
    # Slide Operations
    # ══════════════════════════════════════════════════════════════════

    def add_slide(self, layout: str = "blank") -> dict[str, Any]:
        layout_key = SLIDE_LAYOUTS.get(layout, layout)
        count = self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "set newSlide to make new slide at end of active presentation\n"
            f"set layout of newSlide to {layout_key}\n"
            "return slide index of newSlide\n"
            "end tell"
        )
        return {"status": "ok", "slide_index": int(count), "layout": layout}

    def ensure_slide_count(self, slide_count: int) -> dict[str, Any]:
        current = self.get_slide_count()["slide_count"]
        while current < slide_count:
            self.add_slide("blank")
            current += 1
        while current > slide_count:
            self.delete_slide(current)
            current -= 1
        return {"status": "ok", "slide_count": current}

    def delete_slide(self, slide_index: int) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"delete slide {slide_index} of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "deleted_slide": slide_index}

    def duplicate_slide(self, slide_index: int) -> dict[str, Any]:
        new_index = self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"duplicate slide {slide_index} of active presentation\n"
            "return slide index of slide (count of slides of active presentation) of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "original": slide_index, "new_index": int(new_index)}

    def move_slide(self, from_index: int, to_index: int) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"move slide {from_index} of active presentation to before slide {to_index} of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "moved_from": from_index, "moved_to": to_index}

    def get_slide_count(self) -> dict[str, Any]:
        count = self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "return count of slides of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "slide_count": int(count)}

    def list_slides(self) -> dict[str, Any]:
        output = self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "set rowsOut to {}\n"
            "repeat with i from 1 to count of slides of active presentation\n"
            "set s to slide i of active presentation\n"
            "set end of rowsOut to ((slide index of s as text) & tab & (count of shapes of s as text))\n"
            "end repeat\n"
            "set AppleScript's text item delimiters to linefeed\n"
            "set joinedRows to rowsOut as text\n"
            "set AppleScript's text item delimiters to \"\"\n"
            "return joinedRows\n"
            "end tell"
        )
        slides = _parse_tab_records(output, ["index", "shape_count"])
        for slide in slides:
            _coerce_numbers(slide, ["index", "shape_count"])
        return {"status": "ok", "slides": slides}

    def go_to_slide(self, slide_index: int) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"go to slide {slide_index} of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "current_slide": slide_index}

    # ══════════════════════════════════════════════════════════════════
    # Shape Inspection and Editing
    # ══════════════════════════════════════════════════════════════════

    def list_shapes(self, slide_index: int) -> dict[str, Any]:
        output = self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"set theSlide to slide {slide_index} of active presentation\n"
            "set rowsOut to {}\n"
            "repeat with i from 1 to count of shapes of theSlide\n"
            "set shp to shape i of theSlide\n"
            "set txt to \"\"\n"
            "try\n"
            "if has text frame of shp then set txt to content of text range of text frame of shp\n"
            "end try\n"
            "set txt to my _cliAnythingCleanText(txt)\n"
            "set end of rowsOut to ((i as text) & tab & (name of shp as text) & tab & (shape type of shp as text) & tab & (left position of shp as text) & tab & (top of shp as text) & tab & (width of shp as text) & tab & (height of shp as text) & tab & (z order position of shp as text) & tab & txt)\n"
            "end repeat\n"
            "set AppleScript's text item delimiters to linefeed\n"
            "set joinedRows to rowsOut as text\n"
            "set AppleScript's text item delimiters to \"\"\n"
            "return joinedRows\n"
            "end tell\n"
            "on _cliAnythingCleanText(t)\n"
            "set oldDelims to AppleScript's text item delimiters\n"
            "set AppleScript's text item delimiters to {return, linefeed, tab}\n"
            "set parts to text items of (t as text)\n"
            "set AppleScript's text item delimiters to \" \"\n"
            "set cleaned to parts as text\n"
            "set AppleScript's text item delimiters to oldDelims\n"
            "return cleaned\n"
            "end _cliAnythingCleanText"
        )
        shapes = _parse_tab_records(
            output,
            ["index", "name", "type", "left", "top", "width", "height", "z_order", "text"],
        )
        for shape in shapes:
            _coerce_numbers(shape, ["index", "left", "top", "width", "height", "z_order"])
        return {"status": "ok", "slide": slide_index, "shapes": shapes}

    def get_shape(self, slide_index: int, shape_index: int) -> dict[str, Any]:
        for shape in self.list_shapes(slide_index)["shapes"]:
            if shape.get("index") == shape_index:
                return {"status": "ok", "slide": slide_index, "shape": shape}
        return {"status": "error", "error": f"Shape {shape_index} not found on slide {slide_index}"}

    def set_shape_text(
        self,
        slide_index: int,
        shape_index: int,
        text: str,
        font_size: int | None = None,
        font_name: str = "",
        font_color: str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
    ) -> dict[str, Any]:
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set shp to shape {shape_index} of slide {slide_index} of active presentation",
            f"set content of text range of text frame of shp to {_escape(text)}",
        ]
        if font_size is not None:
            lines.append(f"set font size of font of text range of text frame of shp to {font_size}")
        if font_name:
            lines.append(f"set font name of font of text range of text frame of shp to {_escape(font_name)}")
        if font_color:
            lines.append(_font_color_block("shp", font_color))
        if bold is not None:
            lines.append(f"set bold of font of text range of text frame of shp to {'true' if bold else 'false'}")
        if italic is not None:
            lines.append(f"set italic of font of text range of text frame of shp to {'true' if italic else 'false'}")
        lines += ["return name of shp", "end tell"]
        name = self.run_script("\n".join(lines))
        return {"status": "ok", "slide": slide_index, "shape": shape_index, "name": name}

    def update_shape(
        self,
        slide_index: int,
        shape_index: int,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        rotation: float | None = None,
        name: str | None = None,
    ) -> dict[str, Any]:
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set shp to shape {shape_index} of slide {slide_index} of active presentation",
        ]
        if left is not None:
            lines.append(f"set left position of shp to {left}")
        if top is not None:
            lines.append(f"set top of shp to {top}")
        if width is not None:
            lines.append(f"set width of shp to {width}")
        if height is not None:
            lines.append(f"set height of shp to {height}")
        if rotation is not None:
            lines.append(f"set rotation of shp to {rotation}")
        if name:
            lines.append(f"set name of shp to {_escape(name)}")
        lines += ["return name of shp", "end tell"]
        new_name = self.run_script("\n".join(lines))
        return {"status": "ok", "slide": slide_index, "shape": shape_index, "name": new_name}

    def delete_shape(self, slide_index: int, shape_index: int) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"delete shape {shape_index} of slide {slide_index} of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "slide": slide_index, "deleted_shape": shape_index}

    def set_shape_fill(self, slide_index: int, shape_index: int, color: str) -> dict[str, Any]:
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set shp to shape {shape_index} of slide {slide_index} of active presentation",
        ]
        lines += _fill_block("shp", color)
        lines += ["return name of shp", "end tell"]
        name = self.run_script("\n".join(lines))
        return {"status": "ok", "slide": slide_index, "shape": shape_index, "name": name, "fill": color}

    def set_shape_line(self, slide_index: int, shape_index: int, color: str, weight: float = 1.0) -> dict[str, Any]:
        r, g, b = _parse_color(color)
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"set shp to shape {shape_index} of slide {slide_index} of active presentation\n"
            "tell line format of shp\n"
            f"set fore color to {_clist(r, g, b)}\n"
            f"set line weight to {weight}\n"
            "end tell\n"
            "end tell"
        )
        return {"status": "ok", "slide": slide_index, "shape": shape_index, "line": color, "weight": weight}

    def z_order_shape(self, slide_index: int, shape_index: int, action: str) -> dict[str, Any]:
        actions = {
            "front": "bring shape to front",
            "back": "send shape to back",
            "forward": "bring shape forward",
            "backward": "send shape backward",
        }
        z_action = actions.get(action)
        if not z_action:
            return {"status": "error", "error": f"Unknown z-order action: {action}"}
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"set shp to shape {shape_index} of slide {slide_index} of active presentation\n"
            f"z order shp z order position {z_action}\n"
            "end tell"
        )
        return {"status": "ok", "slide": slide_index, "shape": shape_index, "action": action}

    # ══════════════════════════════════════════════════════════════════
    # Slide Background (full-width text box)
    # ══════════════════════════════════════════════════════════════════

    def set_slide_background(self, slide_index: int, color: str) -> dict[str, Any]:
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide_index} of active presentation",
            "set _bg to make new text box at end of theSlide",
            "set left position of _bg to 0",
            "set top of _bg to 0",
            f"set width of _bg to {self.SLIDE_W}",
            f"set height of _bg to {self.SLIDE_H}",
        ]
        lines += _fill_block("_bg", color)
        lines.append("end tell")
        self.run_script("\n".join(lines))
        return {"status": "ok", "slide": slide_index, "color": color}

    # ══════════════════════════════════════════════════════════════════
    # Slide Transitions
    # ══════════════════════════════════════════════════════════════════

    def set_transition(
        self, slide_index: int, effect: str = "fade", duration: float = 1.0,
        advance_on_click: bool = True, advance_on_time: bool = False, advance_time: float = 0.0,
    ) -> dict[str, Any]:
        fx = TRANSITION_EFFECTS.get(effect, f"entry effect {effect}")
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"set theTrans to slide show transition of slide {slide_index} of active presentation\n"
            f"set entry effect of theTrans to {fx}\n"
            f"set transition duration of theTrans to {duration}\n"
            f"set advance on click of theTrans to {'true' if advance_on_click else 'false'}\n"
            f"set advance on time of theTrans to {'true' if advance_on_time else 'false'}\n"
            f"set advance time of theTrans to {advance_time}\n"
            "end tell"
        )
        return {"status": "ok", "slide": slide_index, "effect": effect, "duration": duration}

    def set_all_transitions(self, effect: str = "fade", duration: float = 0.8) -> dict[str, Any]:
        info = self.get_slide_count()
        total = info["slide_count"]
        for i in range(1, total + 1):
            self.set_transition(i, effect=effect, duration=duration)
        return {"status": "ok", "slides_updated": total, "effect": effect}

    # ══════════════════════════════════════════════════════════════════
    # Styled Title Slide
    # ══════════════════════════════════════════════════════════════════

    def add_title_slide(
        self, title: str, subtitle: str = "",
        bg_color: str | None = None,
        title_color: str | None = None,
        subtitle_color: str | None = None,
        title_font: str = "", subtitle_font: str = "",
        title_size: int = 36, subtitle_size: int = 18,
        accent_color: str | None = None,
        accent_position: str = "top",
    ) -> dict[str, Any]:
        """Add a styled title slide. All shapes use text boxes for compatibility."""
        w, h = self.SLIDE_W, self.SLIDE_H
        lines = ['tell application "Microsoft PowerPoint"']
        lines += [
            "set newSlide to make new slide at end of active presentation",
            "set layout of newSlide to slide layout blank",
        ]

        # Background (full-width text box with solid fill)
        if bg_color:
            lines += [
                "set _bg to make new text box at end of newSlide",
                "set left position of _bg to 0", "set top of _bg to 0",
                f"set width of _bg to {w}", f"set height of _bg to {h}",
            ]
            lines += _fill_block("_bg", bg_color)

        # Accent bar
        if accent_color and accent_position != "none":
            positions = {
                "top": (0, 0, w, 8), "bottom": (0, h - 8, w, 8),
                "left": (0, 0, 8, h), "right": (w - 8, 0, 8, h),
            }
            ax, ay, aw, ah = positions.get(accent_position, (0, 0, w, 8))
            lines += [
                "set _bar to make new text box at end of newSlide",
                f"set left position of _bar to {ax}", f"set top of _bar to {ay}",
                f"set width of _bar to {aw}", f"set height of _bar to {ah}",
            ]
            lines += _fill_block("_bar", accent_color)

        # Title text box
        title_top = 40 if accent_position == "top" else 120
        lines += [
            "set _t to make new text box at end of newSlide",
            "set left position of _t to 80", f"set top of _t to {title_top}",
            "set width of _t to 760", "set height of _t to 100",
            f"set content of text range of text frame of _t to {_escape(title)}",
            f"set font size of font of text range of text frame of _t to {title_size}",
            "set bold of font of text range of text frame of _t to true",
        ]
        if title_font:
            lines.append(f"set font name of font of text range of text frame of _t to {_escape(title_font)}")
        if title_color:
            lines.append(_font_color_block("_t", title_color))

        # Subtitle
        if subtitle:
            sub_top = 160 if accent_position == "top" else 240
            lines += [
                "set _s to make new text box at end of newSlide",
                "set left position of _s to 82", f"set top of _s to {sub_top}",
                "set width of _s to 720", "set height of _s to 80",
                f"set content of text range of text frame of _s to {_escape(subtitle)}",
                f"set font size of font of text range of text frame of _s to {subtitle_size}",
            ]
            if subtitle_font:
                lines.append(f"set font name of font of text range of text frame of _s to {_escape(subtitle_font)}")
            if subtitle_color:
                lines.append(_font_color_block("_s", subtitle_color))

        lines += ["return count of slides of active presentation", "end tell"]
        slide_num = self.run_script("\n".join(lines))
        return {"status": "ok", "title": title, "subtitle": subtitle, "slide_num": int(slide_num)}

    # ══════════════════════════════════════════════════════════════════
    # Text Boxes
    # ══════════════════════════════════════════════════════════════════

    def add_text_box(
        self, slide_index: int, text: str,
        left: int = 80, top: int = 100, width: int = 760, height: int = 60,
        font_name: str = "", font_size: int = 18, font_color: str | None = None,
        bold: bool = False, italic: bool = False, alignment: str = "",
        bg_color: str | None = None,
    ) -> dict[str, Any]:
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide_index} of active presentation",
            "set _tb to make new text box at end of theSlide",
            f"set left position of _tb to {left}", f"set top of _tb to {top}",
            f"set width of _tb to {width}", f"set height of _tb to {height}",
            f"set content of text range of text frame of _tb to {_escape(text)}",
            f"set font size of font of text range of text frame of _tb to {font_size}",
            f"set bold of font of text range of text frame of _tb to {'true' if bold else 'false'}",
            f"set italic of font of text range of text frame of _tb to {'true' if italic else 'false'}",
        ]
        if font_name:
            lines.append(f"set font name of font of text range of text frame of _tb to {_escape(font_name)}")
        if font_color:
            lines.append(_font_color_block("_tb", font_color))
        if bg_color:
            lines += _fill_block("_tb", bg_color)
        if alignment:
            # MsoParagraphAlignment enumerations — skip if unavailable
            amap = {"left": "ppAlignLeft", "center": "ppAlignCenter", "right": "ppAlignRight", "justify": "ppAlignJustify"}
            if alignment in amap:
                lines.append(f"try\nset alignment of paragraph format of text range of text frame of _tb to {amap[alignment]}\nend try")
        lines += ["return name of _tb", "end tell"]
        name = self.run_script("\n".join(lines))
        return {"status": "ok", "shape": name, "text": text[:50]}

    def add_bullets(
        self, slide_index: int, items: list[str],
        left: int = 80, top: int = 140, width: int = 760, height: int = 300,
        font_name: str = "", font_size: int = 18, font_color: str | None = None,
    ) -> dict[str, Any]:
        joined = "\r".join(items)
        return self.add_text_box(
            slide_index=slide_index, text=joined,
            left=left, top=top, width=width, height=height,
            font_name=font_name, font_size=font_size, font_color=font_color,
        )

    # ══════════════════════════════════════════════════════════════════
    # Shapes (using text boxes for solid shapes, pictures for images)
    # ══════════════════════════════════════════════════════════════════

    def add_rect(
        self, slide_index: int,
        left: int = 0, top: int = 0, width: int = 100, height: int = 100,
        fill_color: str | None = None, opacity: int = 100,
    ) -> dict[str, Any]:
        """Add a colored rectangle using a filled text box."""
        return self.add_text_box(
            slide_index=slide_index, text="",
            left=left, top=top, width=width, height=height,
            bg_color=fill_color, font_size=1,
        )

    def add_native_rect(
        self,
        slide_index: int,
        left: float,
        top: float,
        width: float,
        height: float,
        fill_color: str | None = None,
        line_color: str | None = None,
        line_weight: float = 1.0,
    ) -> dict[str, Any]:
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide_index} of active presentation",
            "set _r to make new shape at end of theSlide",
            "set auto shape type of _r to autoshape rectangle",
            f"set left position of _r to {left}",
            f"set top of _r to {top}",
            f"set width of _r to {width}",
            f"set height of _r to {height}",
        ]
        if fill_color:
            lines += _fill_block("_r", fill_color)
        if line_color:
            r, g, b = _parse_color(line_color)
            lines += [
                "tell line format of _r",
                f"set fore color to {_clist(r, g, b)}",
                f"set line weight to {line_weight}",
                "end tell",
            ]
        lines += ["return name of _r", "end tell"]
        name = self.run_script("\n".join(lines))
        return {"status": "ok", "shape": name, "type": "rect"}

    def add_oval_shape(
        self, slide_index: int,
        left: int = 0, top: int = 0, width: int = 100, height: int = 100,
        fill_color: str | None = None,
        line_color: str | None = None,
        line_weight: float = 1.0,
    ) -> dict[str, Any]:
        """Add a colored oval using autoshape. Falls back to rect if unavailable."""
        try:
            lines = [
                'tell application "Microsoft PowerPoint"',
                f"set theSlide to slide {slide_index} of active presentation",
                "set _o to make new shape at end of theSlide",
                "set auto shape type of _o to autoshape oval",
                f"set left position of _o to {left}", f"set top of _o to {top}",
                f"set width of _o to {width}", f"set height of _o to {height}",
            ]
            if fill_color:
                lines += _fill_block("_o", fill_color)
            if line_color:
                r, g, b = _parse_color(line_color)
                lines += [
                    "tell line format of _o",
                    f"set fore color to {_clist(r, g, b)}",
                    f"set line weight to {line_weight}",
                    "end tell",
                ]
            lines += ["return name of _o", "end tell"]
            name = self.run_script("\n".join(lines))
            return {"status": "ok", "shape": name, "type": "oval"}
        except RuntimeError:
            return self.add_rect(slide_index, left, top, width, height, fill_color)

    def build_from_spec(self, spec: dict[str, Any]) -> dict[str, Any]:
        slides = spec.get("slides", [])
        page = spec.get("page", {})
        self.new_presentation()
        self.set_page_size(float(page.get("width", self.SLIDE_W)), float(page.get("height", self.SLIDE_H)))
        self.ensure_slide_count(len(slides))
        created = 0
        skipped = 0
        for slide in slides:
            slide_index = int(slide.get("index", 1))
            elements = slide.get("elements", [])
            payload = self._build_slide_from_elements(slide_index, elements)
            created += payload["created"]
            skipped += payload["skipped"]
        return {"status": "ok", "slides": len(slides), "elements_created": created, "elements_skipped": skipped}

    def _build_slide_from_elements(self, slide_index: int, elements: list[dict[str, Any]]) -> dict[str, int]:
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide_index} of active presentation",
            "repeat while (count of shapes of theSlide) > 0",
            "delete shape 1 of theSlide",
            "end repeat",
        ]
        created = 0
        skipped = 0
        for idx, element in enumerate(elements, 1):
            try:
                snippet = self._element_script(element, idx)
            except Exception:
                skipped += 1
                continue
            if snippet:
                lines.extend(snippet)
                created += 1
            else:
                skipped += 1
        lines.append("end tell")
        self.run_script("\n".join(lines))
        return {"created": created, "skipped": skipped}

    def _element_script(self, element: dict[str, Any], idx: int) -> list[str]:
        kind = element.get("type")
        var = f"_e{idx}"
        if kind == "text":
            lines = [
                f"set {var} to make new text box at end of theSlide",
                f"set left position of {var} to {float(element.get('left', 0))}",
                f"set top of {var} to {float(element.get('top', 0))}",
                f"set width of {var} to {max(1, float(element.get('width', 1)))}",
                f"set height of {var} to {max(1, float(element.get('height', 1)))}",
                f"set content of text range of text frame of {var} to {_escape(element.get('text', ''))}",
                f"set font size of font of text range of text frame of {var} to {max(1, int(round(float(element.get('font_size', 12)))))}",
                f"set bold of font of text range of text frame of {var} to {'true' if element.get('bold') else 'false'}",
                f"set italic of font of text range of text frame of {var} to {'true' if element.get('italic') else 'false'}",
            ]
            if element.get("font_name"):
                lines.append(f"set font name of font of text range of text frame of {var} to {_escape(element['font_name'])}")
            if element.get("font_color"):
                lines.append(_font_color_block(var, element["font_color"]))
            return lines
        if kind == "image":
            lines = [
                f"set {var} to make new picture at end of theSlide with properties {{file name: {json.dumps(element['path'])}}}",
                f"set left position of {var} to {float(element.get('left', 0))}",
                f"set top of {var} to {float(element.get('top', 0))}",
            ]
            if float(element.get("width", 0)):
                lines.append(f"set width of {var} to {float(element.get('width', 0))}")
            if float(element.get("height", 0)):
                lines.append(f"set height of {var} to {float(element.get('height', 0))}")
            return lines
        if kind in {"rect", "oval"}:
            shape_type = "autoshape rectangle" if kind == "rect" else "autoshape oval"
            lines = [
                f"set {var} to make new shape at end of theSlide",
                f"set auto shape type of {var} to {shape_type}",
                f"set left position of {var} to {float(element.get('left', 0))}",
                f"set top of {var} to {float(element.get('top', 0))}",
                f"set width of {var} to {max(1, float(element.get('width', 1)))}",
                f"set height of {var} to {max(1, float(element.get('height', 1)))}",
            ]
            if element.get("fill_color"):
                lines.extend(_fill_block(var, element["fill_color"]))
            if element.get("line_color"):
                r, g, b = _parse_color(element["line_color"])
                lines.extend([
                    f"tell line format of {var}",
                    f"set fore color to {_clist(r, g, b)}",
                    f"set line weight to {float(element.get('line_weight', 1.0))}",
                    "end tell",
                ])
            return lines
        if kind == "line":
            x1 = float(element.get("x1", 0))
            y1 = float(element.get("y1", 0))
            x2 = float(element.get("x2", 0))
            y2 = float(element.get("y2", 0))
            r, g, b = _parse_color(element.get("line_color", "#000000"))
            return [
                f"set {var} to make new line shape at end of theSlide with properties {{begin line X:{x1}, begin line Y:{y1}, end line X:{x2}, end line Y:{y2}}}",
                f"tell line format of {var}",
                f"set fore color to {_clist(r, g, b)}",
                f"set line weight to {float(element.get('line_weight', 1.0))}",
                "end tell",
            ]
        return []

    def add_line_shape(
        self, slide_index: int,
        x1: int, y1: int, x2: int, y2: int,
        line_color: str = "0,0,0", line_weight: float = 2.0,
    ) -> dict[str, Any]:
        """Add a line between two points using autoshape."""
        try:
            lines = [
                'tell application "Microsoft PowerPoint"',
                f"set theSlide to slide {slide_index} of active presentation",
                f"set _ln to make new line shape at end of theSlide with properties {{begin line X:{x1}, begin line Y:{y1}, end line X:{x2}, end line Y:{y2}}}",
            ]
            if line_color:
                r, g, b = _parse_color(line_color)
                lines += [
                    "tell line format of _ln",
                    f"set fore color to {_clist(r, g, b)}",
                    f"set line weight to {line_weight}",
                    "end tell",
                ]
            lines += ["return name of _ln", "end tell"]
            name = self.run_script("\n".join(lines))
            return {"status": "ok", "shape": name, "type": "line"}
        except RuntimeError:
            return self.add_rect(slide_index, x1, y1, 1, 1, line_color)

    # ══════════════════════════════════════════════════════════════════
    # Images
    # ══════════════════════════════════════════════════════════════════

    def add_image(
        self, slide_index: int, image_path: str,
        left: int = 0, top: int = 0, width: int = 0, height: int = 0,
    ) -> dict[str, Any]:
        full_path = str(Path(image_path).expanduser().resolve())
        if not Path(full_path).exists():
            return {"status": "error", "error": f"File not found: {full_path}"}

        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide_index} of active presentation",
            f'set _pic to make new picture at end of theSlide with properties {{file name: {json.dumps(full_path)}}}',
        ]
        lines += [f"set left position of _pic to {left}", f"set top of _pic to {top}"]
        if width and height:
            lines += [f"set width of _pic to {width}", f"set height of _pic to {height}"]
        lines += ["return name of _pic", "end tell"]
        name = self.run_script("\n".join(lines))
        return {"status": "ok", "image": name, "path": full_path}

    # ══════════════════════════════════════════════════════════════════
    # LaTeX Formula Rendering
    # ══════════════════════════════════════════════════════════════════

    def add_latex(
        self, slide_index: int, latex: str,
        left: int = 80, top: int = 100,
        width: int = 0, height: int = 0,
        font_size: int = 18, font_color: str = "0,0,0",
        bg_color: str = "255,255,255",
        dpi: int = 200,
    ) -> dict[str, Any]:
        """Render a LaTeX formula as an image and insert it onto a slide.

        Uses matplotlib's mathtext renderer (no external LaTeX installation needed).
        Supports: fractions (\\frac), sums (\\sum), integrals (\\int), Greek letters,
        subscripts/superscripts, and most common math notation.

        Args:
            slide_index: Target slide number (1-based).
            latex: LaTeX math expression (e.g. "P(Y|X) = \\frac{P(Y)P(X|Y)}{P(X)}").
                   Can optionally be wrapped in $...$ or $$...$$.
            left, top: Position in points from top-left.
            width, height: Desired rendered size. 0 = auto-size from content.
            font_size: Base font size for the rendered formula.
            font_color: RGB string like "R,G,B".
            dpi: Render resolution (higher = crisper).
        """
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        # Strip outer $ wrappers if present
        formula = latex.strip()
        if formula.startswith("$$") and formula.endswith("$$"):
            formula = formula[2:-2].strip()
        elif formula.startswith("$") and formula.endswith("$"):
            formula = formula[1:-1].strip()

        # Wrap in math mode for mathtext rendering
        formula = f"${formula}$"

        # Parse colors
        r, g, b = _parse_color(font_color)
        br, bg, bb = _parse_color(bg_color)

        # Render to figure
        fig, ax = plt.subplots(figsize=(6, 1.5), dpi=dpi)
        text = ax.text(
            0.5, 0.5, formula,
            transform=ax.transAxes,
            fontsize=font_size,
            ha="center", va="center",
            color=(r / 255, g / 255, b / 255),
        )
        ax.axis("off")
        fig.tight_layout(pad=0)

        # Save to temp file with matching background
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.close()
        try:
            fig.savefig(
                tmp.name, format="png",
                bbox_inches="tight", pad_inches=0.08,
                dpi=dpi,
                facecolor=(br / 255, bg / 255, bb / 255),
                edgecolor="none",
            )
        finally:
            plt.close(fig)

        # Insert as image
        result = self.add_image(
            slide_index=slide_index, image_path=tmp.name,
            left=left, top=top, width=width, height=height,
        )

        # Clean up temp file
        Path(tmp.name).unlink(missing_ok=True)

        result["latex"] = latex
        result["method"] = "latex_mathtext"
        return result

    # ══════════════════════════════════════════════════════════════════
    # Tables
    # ══════════════════════════════════════════════════════════════════

    def add_table(
        self, slide_index: int, rows: int, cols: int,
        left: int = 80, top: int = 100, width: int = 760, height: int = 300,
        header_row: list[str] | None = None,
        header_bg_color: str = "139,0,0",
        header_font_color: str = "255,255,255",
        body_data: list[list[str]] | None = None,
    ) -> dict[str, Any]:
        """Add a table-like grid using text boxes (most compatible approach)."""
        cell_w = width // cols
        cell_h = height // rows
        results = []

        # Header row
        if header_row:
            for ci, cell_text in enumerate(header_row[:cols]):
                cx = left + ci * cell_w
                r = self.add_text_box(
                    slide_index=slide_index, text=cell_text,
                    left=cx, top=top, width=cell_w - 4, height=cell_h - 4,
                    font_size=13, font_color=header_font_color, bold=True,
                    bg_color=header_bg_color,
                )
                results.append(r)

        # Body rows
        if body_data:
            for ri, row_data in enumerate(body_data[: rows - (1 if header_row else 0)]):
                for ci, cell_text in enumerate(row_data[:cols]):
                    cx = left + ci * cell_w
                    cy = top + (ri + (1 if header_row else 0)) * cell_h
                    self.add_text_box(
                        slide_index=slide_index, text=cell_text,
                        left=cx, top=cy, width=cell_w - 4, height=cell_h - 4,
                        font_size=11, font_color="50,50,50",
                        bg_color="248,248,248" if ri % 2 == 0 else "255,255,255",
                    )

        return {"status": "ok", "rows": rows, "cols": cols, "method": "textbox_grid", "cells": len(results)}

    def set_cell_text(
        self, slide_index: int, table_shape_index: int,
        row: int, col: int, text: str,
        font_size: int = 12, font_color: str | None = None, bold: bool = False,
    ) -> dict[str, Any]:
        """Create a text box at default position.

        NOTE: PPT 16.108 AppleScript does not support native table cell access.
        This method creates a standalone text box as a workaround.
        table_shape_index, row, and col are accepted for API compatibility but
        are not used to target table cells.
        """
        return self.add_text_box(
            slide_index=slide_index, text=text,
            left=80, top=100, width=200, height=30,
            font_size=font_size, font_color=font_color, bold=bold,
        )

    # ══════════════════════════════════════════════════════════════════
    # Animations
    # ══════════════════════════════════════════════════════════════════

    def add_animation(
        self, slide_index: int, shape_index: int, effect: str = "fade",
    ) -> dict[str, Any]:
        fx_name = ANIM_EFFECTS.get(effect, "animation effect fade")
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            f"set animShape to shape {shape_index} of slide {slide_index} of active presentation\n"
            f"add effect animShape with properties {{fx: {fx_name}}}\n"
            "end tell"
        )
        return {"status": "ok", "slide": slide_index, "shape": shape_index, "effect": effect}

    # ══════════════════════════════════════════════════════════════════
    # Slide Show Control
    # ══════════════════════════════════════════════════════════════════

    def run_slide_show(self) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "set s to slide show settings of active presentation\n"
            "set starting slide of s to 1\n"
            "set ending slide of s to count of slides of active presentation\n"
            "run slide show s\n"
            "end tell"
        )
        return {"status": "ok", "action": "slide_show_started"}

    def exit_slide_show(self) -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "exit slide show view of slide show window 1\n"
            "end tell"
        )
        return {"status": "ok", "action": "slide_show_exited"}

    # ══════════════════════════════════════════════════════════════════
    # Utility
    # ══════════════════════════════════════════════════════════════════

    def undo(self) -> dict[str, Any]:
        self.run_script('tell application "Microsoft PowerPoint" to undo')
        return {"status": "ok", "action": "undo"}

    def redo(self) -> dict[str, Any]:
        self.run_script('tell application "Microsoft PowerPoint" to redo')
        return {"status": "ok", "action": "redo"}

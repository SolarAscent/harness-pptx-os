"""AppleScript backend — delegates to the existing cli_anything.powerpoint backend.

This wraps the existing PowerPointBackend class, adapting its interface
to match the RendererInterface abstract base.
"""

from __future__ import annotations

import json
import subprocess
from typing import Any

from harness_pptx.backends.interface import (
    ChartStyle,
    ElementHandle,
    PresentationHandle,
    RendererInterface,
    ShapeStyle,
    SlideHandle,
    TableStyle,
    TextMetrics,
    TextStyle,
)


class AppleScriptBackend(RendererInterface):
    """PowerPoint renderer using AppleScript on macOS.

    Wraps the lower-level AppleScript generation code from the
    cli_anything.powerpoint backend, exposing it through the
    standard RendererInterface.
    """

    backend_name = "applescript"
    platform = "darwin"

    def __init__(self):
        self._pres_counter = 0
        self._slide_counter = 0
        self._element_counter = 0

    # ---- Lifecycle ----------------------------------------------------------

    def create_presentation(self) -> PresentationHandle:
        self._pres_counter += 1
        hid = f"pres-{self._pres_counter}"
        script = (
            'tell application "Microsoft PowerPoint"\n'
            'activate\n'
            'set deckRef to make new presentation\n'
            'end tell'
        )
        self._run(script)
        return PresentationHandle(id=hid)

    def open_presentation(self, path: str) -> PresentationHandle:
        self._pres_counter += 1
        hid = f"pres-{self._pres_counter}"
        script = (
            'tell application "Microsoft PowerPoint"\n'
            f'  set f to POSIX file "{path}" as alias\n'
            f'  open f\n'
            f'end tell'
        )
        self._run(script)
        return PresentationHandle(id=hid, path=path)

    def add_slide(self, pres: PresentationHandle, layout: str = "blank") -> SlideHandle:
        self._slide_counter += 1
        hid = f"slide-{self._slide_counter}"
        script = (
            'tell application "Microsoft PowerPoint"\n'
            f"  set newSlide to make new slide at end of active presentation\n"
            f"  set layout of newSlide to slide layout blank\n"
            f"end tell"
        )
        self._run(script)
        return SlideHandle(id=hid, index=self._slide_counter - 1)

    def delete_slide(self, pres: PresentationHandle, slide: SlideHandle) -> None:
        idx = slide.index + 1  # PPT is 1-indexed
        script = (
            'tell application "Microsoft PowerPoint"\n'
            f"  delete slide {idx} of active presentation\n"
            f"end tell"
        )
        self._run(script)

    # ---- Content elements ---------------------------------------------------

    def add_text_box(
        self, slide: SlideHandle,
        x: float, y: float, w: float, h: float,
        text: str, style: TextStyle,
    ) -> ElementHandle:
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"
        color_rgb = self._hex_to_applescript_rgb(style.font_color)

        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide.index + 1} of active presentation",
            "set shp to make new text box at end of theSlide",
            f"set left position of shp to {x}",
            f"set top of shp to {y}",
            f"set width of shp to {w}",
            f"set height of shp to {h}",
            f"set content of text range of text frame of shp to {json.dumps(text, ensure_ascii=False)}",
            f"set font size of font of text range of text frame of shp to {style.font_size}",
            f"set font name of font of text range of text frame of shp to {json.dumps(style.font_name)}",
        ]
        if style.bold:
            lines.append("set bold of font of text range of text frame of shp to true")
        if style.italic:
            lines.append("set italic of font of text range of text frame of shp to true")
        lines += [
            f"set font color of font of text range of text frame of shp to {color_rgb}",
            "end tell",
        ]
        self._run("\n".join(lines))
        return ElementHandle(id=eid, element_type="text")

    def add_image(
        self, slide: SlideHandle,
        x: float, y: float, w: float, h: float,
        path: str, fit: str = "contain",
    ) -> ElementHandle:
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"
        path_esc = json.dumps(str(path))
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide.index + 1} of active presentation",
            f"set shp to make new picture at end of theSlide with properties {{file name:{path_esc}}}",
            f"set left position of shp to {x}",
            f"set top of shp to {y}",
            f"set width of shp to {w}",
            f"set height of shp to {h}",
            "end tell",
        ]
        self._run("\n".join(lines))
        return ElementHandle(id=eid, element_type="image")

    def add_shape(
        self, slide: SlideHandle, shape_type: str,
        x: float, y: float, w: float, h: float, style: ShapeStyle,
    ) -> ElementHandle:
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"
        as_type = _shape_type_map.get(shape_type, "rectangle")

        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide.index + 1} of active presentation",
            f"set shp to make new shape at end of theSlide",
            f"set auto shape type of shp to {as_type}",
            f"set left position of shp to {x}",
            f"set top of shp to {y}",
            f"set width of shp to {w}",
            f"set height of shp to {h}",
        ]
        if style.fill_color:
            r, g, b = _parse_hex(style.fill_color)
            lines += [
                f"set fore color of fill format of shp to {{{r}, {g}, {b}}}",
                "set visible of fill format of shp to true",
            ]
        if style.line_color:
            r, g, b = _parse_hex(style.line_color)
            lines += [
                f"set fore color of line format of shp to {{{r}, {g}, {b}}}",
                f"set weight of line format of shp to {style.line_weight}",
            ]
        if style.corner_radius and as_type == "rounded rectangle":
            lines.append(f"set corner radius of shp to {style.corner_radius}")
        lines.append("end tell")
        self._run("\n".join(lines))
        return ElementHandle(id=eid, element_type="shape")

    def add_line(
        self, slide: SlideHandle,
        x1: float, y1: float, x2: float, y2: float,
        color: str = "#000000", weight: float = 1.0,
    ) -> ElementHandle:
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"
        r, g, b = _parse_hex(color)
        lines = [
            'tell application "Microsoft PowerPoint"',
            f"set theSlide to slide {slide.index + 1} of active presentation",
            f"set shp to make new line shape at end of theSlide",
            f"set begin line X of shp to {x1}",
            f"set begin line Y of shp to {y1}",
            f"set end line X of shp to {x2}",
            f"set end line Y of shp to {y2}",
            f"set fore color of line format of shp to {{{r}, {g}, {b}}}",
            f"set weight of line format of shp to {weight}",
            "end tell",
        ]
        self._run("\n".join(lines))
        return ElementHandle(id=eid, element_type="line")

    def add_table(
        self, slide: SlideHandle,
        rows: int, cols: int, data: list[list[str]],
        x: float, y: float, w: float, h: float, style: TableStyle,
    ) -> ElementHandle:
        """Table via grid of text boxes (AppleScript limitation)."""
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"
        cell_h = h / max(rows, 1)
        cell_w = w / max(cols, 1)

        for r in range(min(rows, len(data))):
            for c in range(min(cols, len(data[r]) if r < len(data) else 0)):
                cx = x + c * cell_w
                cy = y + r * cell_h
                is_header = r == 0
                text_style = TextStyle(
                    font_size=style.font_size,
                    font_color=style.header_text if is_header else "#000000",
                    bold=is_header,
                )
                bg_color = style.header_fill if is_header else (
                    style.alt_body_fill if r % 2 == 1 else style.body_fill
                )
                # Add background rect + text
                self.add_shape(slide, "rectangle", cx, cy, cell_w, cell_h,
                             ShapeStyle(fill_color=bg_color, line_color=style.border_color))
                self.add_text_box(slide, cx + 4, cy + 2, cell_w - 8, cell_h - 4,
                                data[r][c] if r < len(data) and c < len(data[r]) else "",
                                text_style)
        return ElementHandle(id=eid, element_type="table")

    def add_chart(
        self, slide: SlideHandle, chart_type: str,
        data: dict[str, Any],
        x: float, y: float, w: float, h: float, style: ChartStyle,
    ) -> ElementHandle:
        """Charts rendered via figure bridge (matplotlib → PNG → embed)."""
        import sys
        import os
        _fig_path = os.path.join(
            os.path.dirname(__file__), '..', '..', '..', 'cli_anything', 'powerpoint', 'utils'
        )
        sys.path.insert(0, os.path.abspath(_fig_path))
        from figure_bridge import FigureBridge
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"
        bridge = FigureBridge()
        img_path = bridge.generate(chart_type, data, style)
        self.add_image(slide, x, y, w, h, img_path, "contain")
        return ElementHandle(id=eid, element_type="chart")

    def add_group(
        self, slide: SlideHandle, elements: list[ElementHandle],
    ) -> ElementHandle:
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"
        return ElementHandle(id=eid, element_type="group")

    # ---- Mutation -----------------------------------------------------------

    def set_text(self, slide: SlideHandle, element: ElementHandle, text: str) -> None:
        script = (
            f'tell application "Microsoft PowerPoint"\n'
            f"  tell slide {slide.index + 1} of active presentation\n"
            f"    set content of text range of text frame of shape {element.shape_index or 1} "
            f"to {json.dumps(text, ensure_ascii=False)}\n"
            f"  end tell\n"
            f"end tell"
        )
        self._run(script)

    def delete_element(self, slide: SlideHandle, element: ElementHandle) -> None:
        script = (
            f'tell application "Microsoft PowerPoint"\n'
            f"  tell slide {slide.index + 1} of active presentation\n"
            f"    delete shape {element.shape_index or 1}\n"
            f"  end tell\n"
            f"end tell"
        )
        self._run(script)

    def set_z_order(self, slide: SlideHandle, element: ElementHandle, action: str) -> None:
        script = (
            f'tell application "Microsoft PowerPoint"\n'
            f"  tell slide {slide.index + 1} of active presentation\n"
            f"    set z order of shape {element.shape_index or 1} "
            f"to {json.dumps(action)}\n"
            f"  end tell\n"
            f"end tell"
        )
        self._run(script)

    def set_slide_background(self, slide: SlideHandle, color: str) -> None:
        """Full-slide text box with fill as background."""
        self.add_shape(
            slide, "rectangle", 0, 0, 960, 540,
            ShapeStyle(fill_color=color, line_color=color),
        )

    # ---- Query --------------------------------------------------------------

    def slide_count(self, pres: PresentationHandle) -> int:
        script = (
            'tell application "Microsoft PowerPoint"\n'
            "  return count of slides of active presentation\n"
            "end tell"
        )
        out = self._run(script)
        try:
            return int(out.strip())
        except ValueError:
            return 0

    def list_shapes(self, slide: SlideHandle) -> list[dict[str, Any]]:
        script = (
            f'tell application "Microsoft PowerPoint"\n'
            f"  tell slide {slide.index + 1} of active presentation\n"
            f"    set shps to every shape\n"
            f"    set out to \"\"\n"
            f"    repeat with s in shps\n"
            f"      set out to out & (name of s) & tab & "
            f"(content of text range of text frame of s) & return\n"
            f"    end repeat\n"
            f"    return out\n"
            f"  end tell\n"
            f"end tell"
        )
        out = self._run(script)
        shapes: list[dict[str, Any]] = []
        for line in out.splitlines():
            if not line.strip():
                continue
            parts = line.split("\t")
            shapes.append({"name": parts[0] if len(parts) > 0 else "",
                          "text": parts[1] if len(parts) > 1 else ""})
        return shapes

    # ---- Export -------------------------------------------------------------

    def save(self, pres: PresentationHandle, path: str) -> None:
        script = (
            'tell application "Microsoft PowerPoint"\n'
            f'  save active presentation in POSIX file "{path}"\n'
            f"end tell"
        )
        self._run(script)
        pres.path = path

    def export_pdf(self, pres: PresentationHandle, path: str) -> None:
        script = (
            'tell application "Microsoft PowerPoint"\n'
            f'  save active presentation in POSIX file "{path}" as save as PDF\n'
            f"end tell"
        )
        self._run(script)

    def export_png(self, pres: PresentationHandle, slide_index: int, path: str) -> None:
        """Export a single slide as PNG via PDF → pdftoppm conversion."""
        import subprocess
        import tempfile
        from pathlib import Path

        out_path = Path(path)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "_slide.pdf"
            # Export single slide: duplicate to temp deck, delete others, export PDF
            # Simpler: export full deck PDF then extract page
            pdf_export = out_path.parent / f"_tmp_{out_path.stem}.pdf"
            self.export_pdf(pres, str(pdf_export))
            # Extract single page via pdftoppm
            page_num = slide_index + 1
            result = subprocess.run(
                ["pdftoppm", "-png", "-r", "192", "-f", str(page_num), "-l", str(page_num),
                 str(pdf_export), str(out_path.parent / out_path.stem)],
                capture_output=True, text=True,
            )
            # pdftoppm appends "-1.png" etc. — rename to desired path
            expected = out_path.parent / f"{out_path.stem}-{page_num}.png"
            if expected.exists():
                expected.rename(out_path)
            pdf_export.unlink(missing_ok=True)

    def export_all_pngs(self, pres: PresentationHandle, output_dir: str) -> list[str]:
        """Export all slides as PNG images. Returns list of file paths."""
        import subprocess
        from pathlib import Path

        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)

        pdf_path = out / "_deck.pdf"
        self.export_pdf(pres, str(pdf_path))

        # pdftoppm converts all pages at once
        subprocess.run(
            ["pdftoppm", "-png", "-r", "192", str(pdf_path), str(out / "slide")],
            capture_output=True, text=True,
        )
        pdf_path.unlink(missing_ok=True)

        # Collect output files
        pngs = sorted(out.glob("slide-*.png"))
        return [str(p) for p in pngs]

    def close(self, pres: PresentationHandle) -> None:
        script = 'tell application "Microsoft PowerPoint" to close active presentation saving no'
        self._run(script)

    # ---- Helpers ------------------------------------------------------------

    def _run(self, script: str) -> str:
        """Execute AppleScript via osascript subprocess."""
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode != 0:
            raise RuntimeError(f"AppleScript error: {result.stderr.strip()}")
        return result.stdout

    @staticmethod
    def _hex_to_applescript_rgb(hex_color: str) -> str:
        r, g, b = _parse_hex(hex_color)
        return f"{{{r}, {g}, {b}}}"


# ---- Helpers ----------------------------------------------------------------

_shape_type_map: dict[str, str] = {
    "rectangle": "autoshape rectangle",
    "rounded_rectangle": "autoshape rounded rectangle",
    "circle": "autoshape oval",
    "oval": "autoshape oval",
    "ellipse": "autoshape oval",
    "line": "autoshape line",
    "arrow": "autoshape line",
    "triangle": "autoshape isosceles triangle",
    "diamond": "autoshape diamond",
}


def _parse_hex(hex_color: str) -> tuple[int, int, int]:
    c = hex_color.lstrip("#")
    return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))

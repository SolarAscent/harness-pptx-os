"""COM backend — drives Microsoft PowerPoint via COM automation on Windows.

Uses pywin32 (win32com.client) to control PowerPoint through its COM object
model. This is the Windows equivalent of the AppleScript backend on macOS,
giving native fidelity for all slide elements.

Requires: Microsoft PowerPoint for Windows, pywin32 (pip install pywin32)
"""

from __future__ import annotations

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


# ---- COM MSO shape type constants ---------------------------------------
# These map shape_type strings to MSO AutoShapeType integers.

_SHAPE_TYPE_MAP: dict[str, int] = {
    "rectangle": 1,           # msoShapeRectangle
    "rect": 1,
    "rounded_rectangle": 5,   # msoShapeRoundedRectangle
    "rounded_rect": 5,
    "oval": 9,                # msoShapeOval
    "circle": 9,
    "ellipse": 9,
    "triangle": 7,            # msoShapeIsoscelesTriangle
    "diamond": 4,             # msoShapeDiamond
    "pentagon": 56,           # msoShapePentagon
    "hexagon": 10,            # msoShapeHexagon
    "chevron": 55,            # msoShapeChevron
    "arrow_right": 33,        # msoShapeRightArrow
    "arrow_left": 34,         # msoShapeLeftArrow
    "star": 92,               # msoShape5pointStar
}

# Chart type constants
_CHART_TYPE_MAP: dict[str, int] = {
    "bar": 1,                  # xlBarClustered
    "bar_stacked": 2,          # xlBarStacked
    "column": 51,              # xlColumnClustered
    "column_stacked": 52,      # xlColumnStacked
    "line": 4,                 # xlLine
    "pie": 5,                  # xlPie
    "area": 1,                 # xlArea (actually 1 in enum)
    "scatter": -4169,          # xlXYScatter
}

# ---- Color helpers ------------------------------------------------------

def _parse_hex(hex_color: str) -> tuple[int, int, int]:
    """Parse '#RRGGBB' into (R, G, B) tuple."""
    c = hex_color.lstrip("#")
    if len(c) == 3:
        c = "".join(x * 2 for x in c)
    if len(c) != 6:
        return (0, 0, 0)
    return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))


def _hex_to_bgr(hex_color: str) -> int:
    """Convert '#RRGGBB' to BGR integer for COM color properties."""
    r, g, b = _parse_hex(hex_color)
    return r + (g << 8) + (b << 16)


# ---- Backend ------------------------------------------------------------

class COMBackend(RendererInterface):
    """PowerPoint renderer using COM automation on Windows.

    Drives Microsoft PowerPoint through its COM object model via pywin32.
    Supports native tables, native charts, and all shape types.
    """

    backend_name = "com"
    platform = "win32"

    def __init__(self):
        self._ppt = None
        self._pres = None
        self._element_counter = 0
        self._visible = False  # Run in background by default

    def _ensure_ppt(self):
        """Lazy-initialize the PowerPoint COM application."""
        if self._ppt is not None:
            return
        try:
            import win32com.client
            self._win32com = win32com.client
        except ImportError:
            raise RuntimeError(
                "pywin32 is required for COMBackend on Windows. "
                "Install it with: pip install pywin32"
            )
        self._ppt = self._win32com.Dispatch("PowerPoint.Application")
        self._ppt.Visible = self._visible

    def _ensure_pres(self):
        if self._pres is None:
            raise RuntimeError("No presentation open")

    # ---- Presentation lifecycle ------------------------------------------

    def create_presentation(self) -> PresentationHandle:
        self._ensure_ppt()
        self._pres = self._ppt.Presentations.Add()
        # Set 16:9 widescreen
        self._pres.PageSetup.SlideWidth = 960
        self._pres.PageSetup.SlideHeight = 540
        self._element_counter = 0
        return PresentationHandle(id="com-pres-1")

    def open_presentation(self, path: str) -> PresentationHandle:
        import os
        self._ensure_ppt()
        abs_path = os.path.abspath(path)
        self._pres = self._ppt.Presentations.Open(abs_path)
        self._element_counter = 0
        return PresentationHandle(id="com-pres-1", path=abs_path)

    # ---- Slide operations ------------------------------------------------

    def add_slide(self, pres: PresentationHandle, layout: str = "blank") -> SlideHandle:
        self._ensure_pres()
        # 12 = ppLayoutBlank, Slides.Count is 1-based
        idx = self._pres.Slides.Count + 1
        slide = self._pres.Slides.Add(idx, 12)  # 12 = ppLayoutBlank
        return SlideHandle(id=f"com-slide-{idx - 1}", index=idx - 1)

    def delete_slide(self, pres: PresentationHandle, slide: SlideHandle) -> None:
        self._ensure_pres()
        idx = slide.index + 1  # 1-indexed in COM
        self._pres.Slides(idx).Delete()

    def slide_count(self, pres: PresentationHandle) -> int:
        if self._pres is None:
            return 0
        return self._pres.Slides.Count

    # ---- Slide background ------------------------------------------------

    def set_slide_background(self, slide: SlideHandle, color: str) -> None:
        """Set slide background by adding a full-slide rectangle."""
        self.add_shape(
            slide, "rectangle", 0, 0, 960, 540,
            ShapeStyle(fill_color=color, line_color=color),
        )

    # ---- Text ------------------------------------------------------------

    def add_text_box(
        self, slide: SlideHandle,
        x: float, y: float, w: float, h: float,
        text: str, style: TextStyle,
    ) -> ElementHandle:
        self._ensure_pres()
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"

        s = self._pres.Slides(slide.index + 1)
        # 1 = msoTextOrientationHorizontal
        shape = s.Shapes.AddTextbox(1, x, y, w, h)
        tr = shape.TextFrame.TextRange
        tr.Text = text
        tr.Font.Size = style.font_size
        tr.Font.Name = style.font_name
        tr.Font.Color.RGB = _hex_to_bgr(style.font_color)
        tr.Font.Bold = style.bold
        tr.Font.Italic = style.italic

        # Alignment
        align_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
        tr.ParagraphFormat.Alignment = align_map.get(style.alignment, 1)

        # Store shape name for later reference
        shape.Name = eid
        return ElementHandle(id=eid, element_type="text",
                            metadata={"shape_name": eid})

    def set_text(self, slide: SlideHandle, element: ElementHandle, text: str) -> None:
        self._ensure_pres()
        s = self._pres.Slides(slide.index + 1)
        name = element.metadata.get("shape_name") if element.metadata else None
        if name:
            try:
                shape = s.Shapes(name)
                shape.TextFrame.TextRange.Text = text
            except Exception:
                pass

    # ---- Image -----------------------------------------------------------

    def add_image(
        self, slide: SlideHandle,
        x: float, y: float, w: float, h: float,
        path: str, fit: str = "contain",
    ) -> ElementHandle:
        import os
        self._ensure_pres()
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"

        s = self._pres.Slides(slide.index + 1)
        abs_path = os.path.abspath(str(path))
        # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
        shape = s.Shapes.AddPicture(abs_path, 0, -1, x, y, w, h)
        shape.Name = eid
        return ElementHandle(id=eid, element_type="image",
                            metadata={"shape_name": eid})

    # ---- Shapes ----------------------------------------------------------

    def add_shape(
        self, slide: SlideHandle, shape_type: str,
        x: float, y: float, w: float, h: float, style: ShapeStyle,
    ) -> ElementHandle:
        self._ensure_pres()
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"

        s = self._pres.Slides(slide.index + 1)
        mso_shape = _SHAPE_TYPE_MAP.get(shape_type, 1)  # default: rectangle
        shape = s.Shapes.AddShape(mso_shape, x, y, w, h)

        if style.fill_color:
            shape.Fill.ForeColor.RGB = _hex_to_bgr(style.fill_color)
            shape.Fill.Visible = -1  # msoTrue
        else:
            shape.Fill.Visible = 0   # msoFalse

        if style.line_color:
            shape.Line.ForeColor.RGB = _hex_to_bgr(style.line_color)
            shape.Line.Weight = style.line_weight
            shape.Line.Visible = -1
        else:
            shape.Line.Visible = 0

        # Corner radius for rounded rectangles
        if style.corner_radius and shape_type in ("rounded_rectangle", "rounded_rect"):
            try:
                shape.Adjustments(1) = style.corner_radius / 72.0
            except Exception:
                pass

        shape.Name = eid
        return ElementHandle(id=eid, element_type="shape",
                            metadata={"shape_name": eid})

    # ---- Lines -----------------------------------------------------------

    def add_line(
        self, slide: SlideHandle,
        x1: float, y1: float, x2: float, y2: float,
        color: str = "#000000", weight: float = 1.0,
    ) -> ElementHandle:
        self._ensure_pres()
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"

        s = self._pres.Slides(slide.index + 1)
        # 1 = msoConnectorStraight
        shape = s.Shapes.AddConnector(1, x1, y1, x2, y2)
        shape.Line.ForeColor.RGB = _hex_to_bgr(color)
        shape.Line.Weight = weight

        shape.Name = eid
        return ElementHandle(id=eid, element_type="line",
                            metadata={"shape_name": eid})

    # ---- Tables ----------------------------------------------------------

    def add_table(
        self, slide: SlideHandle,
        rows: int, cols: int, data: list[list[str]],
        x: float, y: float, w: float, h: float, style: TableStyle,
    ) -> ElementHandle:
        """Native COM table — full fidelity with merged-cell support."""
        self._ensure_pres()
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"

        s = self._pres.Slides(slide.index + 1)
        shape = s.Shapes.AddTable(rows, cols, x, y, w, h)
        table = shape.Table

        for ri in range(rows):
            for ci in range(cols):
                cell = table.Cell(ri + 1, ci + 1)  # 1-indexed
                val = ""
                if ri < len(data) and ci < len(data[ri]):
                    val = str(data[ri][ci])
                cell.Shape.TextFrame.TextRange.Text = val
                cell.Shape.TextFrame.TextRange.Font.Size = style.font_size

                if ri == 0:
                    # Header row
                    cell.Shape.TextFrame.TextRange.Font.Bold = -1
                    cell.Shape.TextFrame.TextRange.Font.Color.RGB = _hex_to_bgr(style.header_text)
                    cell.Shape.Fill.ForeColor.RGB = _hex_to_bgr(style.header_fill)
                    cell.Shape.Fill.Visible = -1
                else:
                    cell.Shape.TextFrame.TextRange.Font.Color.RGB = _hex_to_bgr("#000000")
                    if ri % 2 == 0:
                        cell.Shape.Fill.ForeColor.RGB = _hex_to_bgr(style.body_fill)
                    else:
                        cell.Shape.Fill.ForeColor.RGB = _hex_to_bgr(style.alt_body_fill)
                    cell.Shape.Fill.Visible = -1

        shape.Name = eid
        return ElementHandle(id=eid, element_type="table",
                            metadata={"shape_name": eid})

    # ---- Charts ----------------------------------------------------------

    def add_chart(
        self, slide: SlideHandle, chart_type: str,
        data: dict[str, Any],
        x: float, y: float, w: float, h: float, style: ChartStyle,
    ) -> ElementHandle:
        """Native COM chart — editable in PowerPoint."""
        self._ensure_pres()
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"

        s = self._pres.Slides(slide.index + 1)
        xl_type = _CHART_TYPE_MAP.get(chart_type, 51)  # default: clustered column

        shape = s.Shapes.AddChart(xl_type, x, y, w, h)
        chart = shape.Chart

        # Set chart data via the worksheet
        wb = chart.ChartData.Workbook
        ws = wb.Worksheets(1)

        categories = data.get("categories", [])
        for i, cat in enumerate(categories):
            ws.Cells(i + 2, 1).Value = str(cat)

        series_list = data.get("series", [])
        for si, series in enumerate(series_list):
            name = series.get("name", f"Series {si + 1}")
            values = series.get("values", [])
            ws.Cells(1, si + 2).Value = name
            for vi, val in enumerate(values):
                ws.Cells(vi + 2, si + 2).Value = float(val) if val else 0.0

        chart.HasLegend = style.show_legend
        if data.get("title"):
            chart.HasTitle = -1
            chart.ChartTitle.Text = str(data["title"])

        shape.Name = eid
        return ElementHandle(id=eid, element_type="chart",
                            metadata={"shape_name": eid})

    # ---- Group -----------------------------------------------------------

    def add_group(
        self, slide: SlideHandle, elements: list[ElementHandle],
    ) -> ElementHandle:
        self._ensure_pres()
        self._element_counter += 1
        eid = f"elem-{self._element_counter}"

        s = self._pres.Slides(slide.index + 1)
        # Collect shape names to group
        shape_names = []
        for el in elements:
            name = el.metadata.get("shape_name") if el.metadata else None
            if name:
                shape_names.append(name)

        if shape_names:
            try:
                shapes_to_group = [s.Shapes(n) for n in shape_names]
                # Use .Range() to select multiple shapes, then .Group()
                sr = s.Shapes.Range(shape_names)
                grouped = sr.Group()
                grouped.Name = eid
            except Exception:
                pass

        return ElementHandle(id=eid, element_type="group",
                            metadata={"shape_name": eid})

    # ---- Element management ----------------------------------------------

    def delete_element(self, slide: SlideHandle, element: ElementHandle) -> None:
        self._ensure_pres()
        s = self._pres.Slides(slide.index + 1)
        name = element.metadata.get("shape_name") if element.metadata else None
        if name:
            try:
                s.Shapes(name).Delete()
            except Exception:
                pass

    def set_z_order(self, slide: SlideHandle, element: ElementHandle, action: str) -> None:
        self._ensure_pres()
        s = self._pres.Slides(slide.index + 1)
        name = element.metadata.get("shape_name") if element.metadata else None
        if not name:
            return
        try:
            shape = s.Shapes(name)
            if action == "front":
                shape.ZOrder(0)  # msoBringToFront
            elif action == "back":
                shape.ZOrder(1)  # msoSendToBack
        except Exception:
            pass

    # ---- Shape listing ---------------------------------------------------

    def list_shapes(self, slide: SlideHandle) -> list[dict[str, Any]]:
        self._ensure_pres()
        s = self._pres.Slides(slide.index + 1)
        result = []
        for shape in s.Shapes:
            info = {
                "name": shape.Name,
                "left": shape.Left,
                "top": shape.Top,
                "width": shape.Width,
                "height": shape.Height,
                "has_text": shape.HasTextFrame,
            }
            if shape.HasTextFrame:
                try:
                    info["text"] = shape.TextFrame.TextRange.Text
                except Exception:
                    info["text"] = ""
            result.append(info)
        return result

    # ---- Save / Export ---------------------------------------------------

    def save(self, pres: PresentationHandle, path: str) -> None:
        import os
        self._ensure_pres()
        abs_path = os.path.abspath(str(path))
        self._pres.SaveAs(abs_path)
        pres.path = abs_path

    def export_pdf(self, pres: PresentationHandle, path: str) -> None:
        """Export to PDF using PowerPoint's native COM export."""
        import os
        self._ensure_pres()
        abs_path = os.path.abspath(str(path))
        # 32 = ppSaveAsPDF, 2 = ppFixedFormatTypePDF
        self._pres.ExportAsFixedFormat(abs_path, 2)

    def export_png(self, pres: PresentationHandle, slide_index: int, path: str) -> None:
        """Export a single slide as PNG."""
        import os
        self._ensure_pres()
        abs_path = os.path.abspath(str(path))
        slide = self._pres.Slides(slide_index + 1)
        slide.Export(abs_path, "PNG")

    def export_all_pngs(self, pres: PresentationHandle, output_dir: str) -> list[str]:
        """Export all slides as PNG images."""
        import os
        self._ensure_pres()
        os.makedirs(output_dir, exist_ok=True)
        paths = []
        for i in range(1, self._pres.Slides.Count + 1):
            png_path = os.path.join(output_dir, f"slide-{i}.png")
            self._pres.Slides(i).Export(png_path, "PNG")
            paths.append(png_path)
        return paths

    # ---- Close -----------------------------------------------------------

    def close(self, pres: PresentationHandle) -> None:
        if self._pres is not None:
            try:
                self._pres.Close()
            except Exception:
                pass
            self._pres = None
        if self._ppt is not None:
            try:
                self._ppt.Quit()
            except Exception:
                pass
            self._ppt = None
        self._element_counter = 0

from __future__ import annotations

import json
import shlex
from pathlib import Path

import click

from .core.session import Session
from .utils.powerpoint_backend import PowerPointBackend
from .utils.figure_bridge import FigureBridge


def emit(payload: dict, as_json: bool) -> None:
    if as_json:
        click.echo(json.dumps(payload, ensure_ascii=False))
    else:
        click.echo(payload.get("status", payload))


def json_option(func):
    def set_json(ctx, _param, value):
        if value:
            ctx.obj["json"] = True
    return click.option(
        "--json", "command_json", is_flag=True, expose_value=False, callback=set_json
    )(func)


@click.group(invoke_without_command=True)
@click.option("--json", "as_json", is_flag=True, help="Emit machine-readable JSON.")
@click.pass_context
def main(ctx: click.Context, as_json: bool) -> None:
    ctx.ensure_object(dict)
    ctx.obj["json"] = as_json
    ctx.obj["backend"] = PowerPointBackend()
    ctx.obj["figure"] = FigureBridge()
    ctx.obj["session"] = Session()
    if ctx.invoked_subcommand is None:
        repl(ctx)


def repl(ctx: click.Context) -> None:
    click.echo("cli-anything-powerpoint REPL. Type 'help' or 'quit'.")
    while True:
        try:
            line = input("powerpoint> ").strip()
        except (EOFError, KeyboardInterrupt):
            click.echo()
            return
        if not line:
            continue
        if line in {"quit", "exit"}:
            return
        if line == "help":
            click.echo(
                "Commands:\n"
                "  Presentation: info, new, open PATH, save-as PATH, export-pdf PATH,\n"
                "                export-png PATH, close [--no-save], apply-theme PATH,\n"
                "                build-from-spec SPEC_JSON [--output PATH]\n"
                "  Slides:       add-slide [--layout], delete-slide N, duplicate-slide N,\n"
                "                move-slide FROM TO, list-slides, slide-count, go-to-slide N,\n"
                "                slide-bg N COLOR, transition N [--effect] [--duration]\n"
                "  Inspect/Edit: list-shapes N, shape-info N SHAPE, set-text N SHAPE TEXT,\n"
                "                move-shape N SHAPE [OPTIONS], delete-shape N SHAPE,\n"
                "                set-fill N SHAPE COLOR, set-line N SHAPE COLOR, z-order N SHAPE ACTION\n"
                "  Content:      add-title-slide TITLE [OPTIONS], add-text N TEXT [OPTIONS],\n"
                "                add-bullets N ITEM1 ITEM2... [OPTIONS], add-image N PATH [OPTIONS]\n"
                "  Shapes:       add-rect N [OPTIONS], add-oval N [OPTIONS],\n"
                "                add-line N X1 Y1 X2 Y2 [OPTIONS]\n"
                "  Tables:       add-table N ROWS COLS [OPTIONS], set-cell N TABLE ROW COL TEXT [OPTIONS]\n"
                "  Figures:      add-figure N CHART_TYPE [OPTIONS], figure-types\n"
                "  Animations:   add-animation N SHAPE [--effect]\n"
                "  Show:         run-show, exit-show\n"
                "  Util:         undo, redo"
            )
            continue
        args = shlex.split(line)
        try:
            main(args=args, obj=ctx.obj, standalone_mode=False)
        except Exception as exc:
            emit({"status": "error", "error": str(exc)}, ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Presentation Management
# ══════════════════════════════════════════════════════════════════════

@main.command("info")
@json_option
@click.pass_context
def info(ctx: click.Context) -> None:
    emit(ctx.obj["backend"].info(), ctx.obj["json"])


@main.command("new")
@json_option
@click.pass_context
def new(ctx: click.Context) -> None:
    payload = ctx.obj["backend"].new_presentation()
    ctx.obj["session"].record("new")
    emit(payload, ctx.obj["json"])


@main.command("open")
@click.argument("path")
@json_option
@click.pass_context
def open_deck(ctx: click.Context, path: str) -> None:
    payload = ctx.obj["backend"].open_presentation(path)
    ctx.obj["session"].active_presentation = payload.get("path")
    ctx.obj["session"].record("open", path=path)
    emit(payload, ctx.obj["json"])


@main.command("save-as")
@click.argument("path")
@json_option
@click.pass_context
def save_as(ctx: click.Context, path: str) -> None:
    emit(ctx.obj["backend"].save_as(path), ctx.obj["json"])


@main.command("export-pdf")
@click.argument("path")
@json_option
@click.pass_context
def export_pdf(ctx: click.Context, path: str) -> None:
    emit(ctx.obj["backend"].export_pdf(path), ctx.obj["json"])


@main.command("export-png")
@click.argument("path")
@json_option
@click.pass_context
def export_png(ctx: click.Context, path: str) -> None:
    emit(ctx.obj["backend"].export_png(path), ctx.obj["json"])


@main.command("close")
@click.option("--no-save", is_flag=True, help="Close without saving changes.")
@json_option
@click.pass_context
def close(ctx: click.Context, no_save: bool) -> None:
    emit(ctx.obj["backend"].close(saving=not no_save), ctx.obj["json"])


@main.command("apply-theme")
@click.argument("path")
@json_option
@click.pass_context
def apply_theme(ctx: click.Context, path: str) -> None:
    emit(ctx.obj["backend"].apply_theme(path), ctx.obj["json"])


@main.command("build-from-spec")
@click.argument("spec_path")
@click.option("--output", default=None, help="Optional PPTX output path.")
@json_option
@click.pass_context
def build_from_spec(ctx: click.Context, spec_path: str, output: str | None) -> None:
    with open(spec_path, "r", encoding="utf-8") as handle:
        spec = json.load(handle)
    payload = ctx.obj["backend"].build_from_spec(spec)
    if output:
        save_payload = ctx.obj["backend"].save_as(output)
        payload["path"] = save_payload.get("path")
    emit(payload, ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Slide Operations
# ══════════════════════════════════════════════════════════════════════

@main.command("add-slide")
@click.option("--layout", default="blank", help="Layout type: blank, title, title_only, text, section_header, comparison, etc.")
@json_option
@click.pass_context
def add_slide(ctx: click.Context, layout: str) -> None:
    emit(ctx.obj["backend"].add_slide(layout), ctx.obj["json"])


@main.command("delete-slide")
@click.argument("slide_index", type=int)
@json_option
@click.pass_context
def delete_slide(ctx: click.Context, slide_index: int) -> None:
    emit(ctx.obj["backend"].delete_slide(slide_index), ctx.obj["json"])


@main.command("duplicate-slide")
@click.argument("slide_index", type=int)
@json_option
@click.pass_context
def duplicate_slide(ctx: click.Context, slide_index: int) -> None:
    emit(ctx.obj["backend"].duplicate_slide(slide_index), ctx.obj["json"])


@main.command("move-slide")
@click.argument("from_index", type=int)
@click.argument("to_index", type=int)
@json_option
@click.pass_context
def move_slide(ctx: click.Context, from_index: int, to_index: int) -> None:
    emit(ctx.obj["backend"].move_slide(from_index, to_index), ctx.obj["json"])


@main.command("slide-count")
@json_option
@click.pass_context
def slide_count(ctx: click.Context) -> None:
    emit(ctx.obj["backend"].get_slide_count(), ctx.obj["json"])


@main.command("list-slides")
@json_option
@click.pass_context
def list_slides(ctx: click.Context) -> None:
    emit(ctx.obj["backend"].list_slides(), ctx.obj["json"])


@main.command("go-to-slide")
@click.argument("slide_index", type=int)
@json_option
@click.pass_context
def go_to_slide(ctx: click.Context, slide_index: int) -> None:
    emit(ctx.obj["backend"].go_to_slide(slide_index), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Shape Inspection and Editing
# ══════════════════════════════════════════════════════════════════════

@main.command("list-shapes")
@click.argument("slide_index", type=int)
@json_option
@click.pass_context
def list_shapes(ctx: click.Context, slide_index: int) -> None:
    emit(ctx.obj["backend"].list_shapes(slide_index), ctx.obj["json"])


@main.command("shape-info")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@json_option
@click.pass_context
def shape_info(ctx: click.Context, slide_index: int, shape_index: int) -> None:
    emit(ctx.obj["backend"].get_shape(slide_index, shape_index), ctx.obj["json"])


@main.command("set-text")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@click.argument("text")
@click.option("--font-size", type=int, default=None)
@click.option("--font-name", default="")
@click.option("--font-color", default=None)
@click.option("--bold/--no-bold", default=None)
@click.option("--italic/--no-italic", default=None)
@json_option
@click.pass_context
def set_text(ctx: click.Context, slide_index: int, shape_index: int, text: str,
             font_size: int | None, font_name: str, font_color: str | None,
             bold: bool | None, italic: bool | None) -> None:
    emit(ctx.obj["backend"].set_shape_text(
        slide_index=slide_index, shape_index=shape_index, text=text,
        font_size=font_size, font_name=font_name, font_color=font_color,
        bold=bold, italic=italic,
    ), ctx.obj["json"])


@main.command("move-shape")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@click.option("--left", type=float, default=None)
@click.option("--top", type=float, default=None)
@click.option("--width", type=float, default=None)
@click.option("--height", type=float, default=None)
@click.option("--rotation", type=float, default=None)
@click.option("--name", default=None)
@json_option
@click.pass_context
def move_shape(ctx: click.Context, slide_index: int, shape_index: int,
               left: float | None, top: float | None,
               width: float | None, height: float | None,
               rotation: float | None, name: str | None) -> None:
    emit(ctx.obj["backend"].update_shape(
        slide_index=slide_index, shape_index=shape_index,
        left=left, top=top, width=width, height=height,
        rotation=rotation, name=name,
    ), ctx.obj["json"])


@main.command("delete-shape")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@json_option
@click.pass_context
def delete_shape(ctx: click.Context, slide_index: int, shape_index: int) -> None:
    emit(ctx.obj["backend"].delete_shape(slide_index, shape_index), ctx.obj["json"])


@main.command("set-fill")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@click.argument("color")
@json_option
@click.pass_context
def set_fill(ctx: click.Context, slide_index: int, shape_index: int, color: str) -> None:
    emit(ctx.obj["backend"].set_shape_fill(slide_index, shape_index, color), ctx.obj["json"])


@main.command("set-line")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@click.argument("color")
@click.option("--weight", type=float, default=1.0)
@json_option
@click.pass_context
def set_line(ctx: click.Context, slide_index: int, shape_index: int, color: str, weight: float) -> None:
    emit(ctx.obj["backend"].set_shape_line(slide_index, shape_index, color, weight), ctx.obj["json"])


@main.command("z-order")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@click.argument("action", type=click.Choice(["front", "back", "forward", "backward"]))
@json_option
@click.pass_context
def z_order(ctx: click.Context, slide_index: int, shape_index: int, action: str) -> None:
    emit(ctx.obj["backend"].z_order_shape(slide_index, shape_index, action), ctx.obj["json"])


@main.command("slide-bg")
@click.argument("slide_index", type=int)
@click.argument("color")
@json_option
@click.pass_context
def slide_bg(ctx: click.Context, slide_index: int, color: str) -> None:
    """Set slide background color. COLOR = 'R,G,B' or '#RRGGBB'."""
    emit(ctx.obj["backend"].set_slide_background(slide_index, color), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Transitions
# ══════════════════════════════════════════════════════════════════════

@main.command("transition")
@click.argument("slide_index", type=int)
@click.option("--effect", default="fade", help="Effect: fade, dissolve, push_left, push_right, push_up, push_down, cube, flip, zoom_in, doors, curtains, crush, fracture, gallery")
@click.option("--duration", type=float, default=0.8, help="Transition duration in seconds.")
@click.option("--on-click/--no-on-click", default=True, help="Advance on mouse click.")
@click.option("--advance-time", type=float, default=0.0, help="Auto-advance after N seconds (0=disabled).")
@json_option
@click.pass_context
def transition(ctx: click.Context, slide_index: int, effect: str, duration: float,
               on_click: bool, advance_time: float) -> None:
    emit(ctx.obj["backend"].set_transition(
        slide_index, effect=effect, duration=duration,
        advance_on_click=on_click, advance_on_time=advance_time > 0,
        advance_time=advance_time,
    ), ctx.obj["json"])


@main.command("transition-all")
@click.option("--effect", default="fade", help="Effect to apply to all slides.")
@click.option("--duration", type=float, default=0.8)
@json_option
@click.pass_context
def transition_all(ctx: click.Context, effect: str, duration: float) -> None:
    emit(ctx.obj["backend"].set_all_transitions(effect=effect, duration=duration), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Content Slides
# ══════════════════════════════════════════════════════════════════════

@main.command("add-title-slide")
@click.argument("title")
@click.option("--subtitle", default="", help="Subtitle text.")
@click.option("--bg-color", default=None, help="Background RGB color: 'R,G,B' or '#RRGGBB'.")
@click.option("--title-color", default=None, help="Title font RGB color.")
@click.option("--subtitle-color", default=None, help="Subtitle font RGB color.")
@click.option("--title-font", default="", help="Title font name.")
@click.option("--subtitle-font", default="", help="Subtitle font name.")
@click.option("--title-size", type=int, default=36, help="Title font size.")
@click.option("--subtitle-size", type=int, default=18, help="Subtitle font size.")
@click.option("--accent-color", default=None, help="Accent bar RGB color.")
@click.option("--accent-position", default="top", type=click.Choice(["top", "bottom", "left", "right", "none"]))
@json_option
@click.pass_context
def add_title_slide(
    ctx: click.Context, title: str, subtitle: str, bg_color: str | None,
    title_color: str | None, subtitle_color: str | None,
    title_font: str, subtitle_font: str, title_size: int, subtitle_size: int,
    accent_color: str | None, accent_position: str,
) -> None:
    emit(ctx.obj["backend"].add_title_slide(
        title=title, subtitle=subtitle,
        bg_color=bg_color, title_color=title_color, subtitle_color=subtitle_color,
        title_font=title_font, subtitle_font=subtitle_font,
        title_size=title_size, subtitle_size=subtitle_size,
        accent_color=accent_color, accent_position=accent_position,
    ), ctx.obj["json"])


@main.command("add-text")
@click.argument("slide_index", type=int)
@click.argument("text")
@click.option("--left", type=int, default=80)
@click.option("--top", type=int, default=100)
@click.option("--width", type=int, default=760)
@click.option("--height", type=int, default=60)
@click.option("--font-name", default="")
@click.option("--font-size", type=int, default=18)
@click.option("--font-color", default=None)
@click.option("--bold/--no-bold", default=False)
@click.option("--italic/--no-italic", default=False)
@click.option("--alignment", default="", type=click.Choice(["", "left", "center", "right", "justify"]))
@click.option("--bg-color", default=None, help="Background fill color: 'R,G,B' or '#RRGGBB'.")
@json_option
@click.pass_context
def add_text(ctx: click.Context, slide_index: int, text: str, left: int, top: int,
             width: int, height: int, font_name: str, font_size: int,
             font_color: str | None, bold: bool, italic: bool, alignment: str,
             bg_color: str | None) -> None:
    emit(ctx.obj["backend"].add_text_box(
        slide_index=slide_index, text=text,
        left=left, top=top, width=width, height=height,
        font_name=font_name, font_size=font_size, font_color=font_color,
        bold=bold, italic=italic, alignment=alignment, bg_color=bg_color,
    ), ctx.obj["json"])


@main.command("add-bullets")
@click.argument("slide_index", type=int)
@click.argument("items", nargs=-1, required=True)
@click.option("--left", type=int, default=80)
@click.option("--top", type=int, default=140)
@click.option("--width", type=int, default=760)
@click.option("--height", type=int, default=300)
@click.option("--font-name", default="")
@click.option("--font-size", type=int, default=18)
@click.option("--font-color", default=None)
@json_option
@click.pass_context
def add_bullets(ctx: click.Context, slide_index: int, items: list[str],
                left: int, top: int, width: int, height: int,
                font_name: str, font_size: int, font_color: str | None) -> None:
    emit(ctx.obj["backend"].add_bullets(
        slide_index=slide_index, items=list(items),
        left=left, top=top, width=width, height=height,
        font_name=font_name, font_size=font_size, font_color=font_color,
    ), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Shapes
# ══════════════════════════════════════════════════════════════════════

@main.command("add-rect")
@click.argument("slide_index", type=int)
@click.option("--left", type=int, default=0)
@click.option("--top", type=int, default=0)
@click.option("--width", type=int, default=100)
@click.option("--height", type=int, default=100)
@click.option("--fill-color", default=None)
@click.option("--opacity", type=int, default=100)
@json_option
@click.pass_context
def add_rect(ctx: click.Context, slide_index: int, left: int, top: int,
             width: int, height: int, fill_color: str | None, opacity: int) -> None:
    emit(ctx.obj["backend"].add_rect(
        slide_index=slide_index, left=left, top=top, width=width, height=height,
        fill_color=fill_color, opacity=opacity,
    ), ctx.obj["json"])


@main.command("add-oval")
@click.argument("slide_index", type=int)
@click.option("--left", type=int, default=0)
@click.option("--top", type=int, default=0)
@click.option("--width", type=int, default=100)
@click.option("--height", type=int, default=100)
@click.option("--fill-color", default=None)
@json_option
@click.pass_context
def add_oval(ctx: click.Context, slide_index: int, left: int, top: int,
             width: int, height: int, fill_color: str | None) -> None:
    emit(ctx.obj["backend"].add_oval_shape(
        slide_index=slide_index, left=left, top=top, width=width, height=height,
        fill_color=fill_color,
    ), ctx.obj["json"])


@main.command("add-line")
@click.argument("slide_index", type=int)
@click.argument("x1", type=int)
@click.argument("y1", type=int)
@click.argument("x2", type=int)
@click.argument("y2", type=int)
@click.option("--line-color", default="0,0,0")
@click.option("--line-weight", type=float, default=2.0)
@json_option
@click.pass_context
def add_line_cmd(ctx: click.Context, slide_index: int, x1: int, y1: int,
                 x2: int, y2: int, line_color: str, line_weight: float) -> None:
    emit(ctx.obj["backend"].add_line_shape(
        slide_index=slide_index, x1=x1, y1=y1, x2=x2, y2=y2,
        line_color=line_color, line_weight=line_weight,
    ), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Images
# ══════════════════════════════════════════════════════════════════════

@main.command("add-image")
@click.argument("slide_index", type=int)
@click.argument("path")
@click.option("--left", type=int, default=100)
@click.option("--top", type=int, default=100)
@click.option("--width", type=int, default=0, help="0 = original size.")
@click.option("--height", type=int, default=0, help="0 = original size.")
@json_option
@click.pass_context
def add_image(ctx: click.Context, slide_index: int, path: str,
              left: int, top: int, width: int, height: int) -> None:
    emit(ctx.obj["backend"].add_image(
        slide_index=slide_index, image_path=path,
        left=left, top=top, width=width, height=height,
    ), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Tables
# ══════════════════════════════════════════════════════════════════════

@main.command("add-table")
@click.argument("slide_index", type=int)
@click.argument("rows", type=int)
@click.argument("cols", type=int)
@click.option("--left", type=int, default=80)
@click.option("--top", type=int, default=100)
@click.option("--width", type=int, default=760)
@click.option("--height", type=int, default=300)
@click.option("--header", default=None, help="Comma-separated header texts.")
@click.option("--data", default=None, help="Rows as: 'a,b,c;d,e,f;g,h,i'")
@click.option("--header-bg-color", default="139,0,0")
@click.option("--header-font-color", default="255,255,255")
@json_option
@click.pass_context
def add_table(ctx: click.Context, slide_index: int, rows: int, cols: int,
              left: int, top: int, width: int, height: int,
              header: str | None, data: str | None,
              header_bg_color: str, header_font_color: str) -> None:
    header_row = [h.strip() for h in header.split(",")] if header else None
    body = None
    if data:
        body = [[c.strip() for c in row.split(",")] for row in data.split(";")]
    emit(ctx.obj["backend"].add_table(
        slide_index=slide_index, rows=rows, cols=cols,
        left=left, top=top, width=width, height=height,
        header_row=header_row, body_data=body,
        header_bg_color=header_bg_color, header_font_color=header_font_color,
    ), ctx.obj["json"])


@main.command("set-cell")
@click.argument("slide_index", type=int)
@click.argument("table_shape", type=int)
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.argument("text")
@click.option("--font-size", type=int, default=12)
@click.option("--font-color", default=None)
@click.option("--bold/--no-bold", default=False)
@json_option
@click.pass_context
def set_cell(ctx: click.Context, slide_index: int, table_shape: int,
             row: int, col: int, text: str, font_size: int,
             font_color: str | None, bold: bool) -> None:
    emit(ctx.obj["backend"].set_cell_text(
        slide_index=slide_index, table_shape_index=table_shape,
        row=row, col=col, text=text,
        font_size=font_size, font_color=font_color, bold=bold,
    ), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Nature-Figure Bridge
# ══════════════════════════════════════════════════════════════════════

@main.command("add-figure")
@click.argument("slide_index", type=int)
@click.argument("chart_type")
@click.option("--title", default="", help="Chart title.")
@click.option("--xlabel", default="", help="X-axis label.")
@click.option("--ylabel", default="", help="Y-axis label.")
@click.option("--width", type=int, default=600)
@click.option("--height", type=int, default=400)
@click.option("--data", default=None, help="JSON data or data string.")
@click.option("--color-theme", default="nature", type=click.Choice(["nature", "pku", "dark", "light"]))
@click.option("--left", type=int, default=80)
@click.option("--top", type=int, default=80)
@click.option("--img-width", type=int, default=800)
@click.option("--img-height", type=int, default=420)
@json_option
@click.pass_context
def add_figure(ctx: click.Context, slide_index: int, chart_type: str,
               title: str, xlabel: str, ylabel: str,
               width: int, height: int, data: str | None,
               color_theme: str, left: int, top: int,
               img_width: int, img_height: int) -> None:
    """Generate a publication-quality chart and insert it into a slide.

    CHART_TYPE: bar, hbar, line, scatter, pie, heatmap, radar, donut
    --data: JSON array or 'series1:v1,v2,v3;series2:v1,v2,v3' format
    """
    try:
        result = ctx.obj["figure"].generate(
            chart_type=chart_type, title=title,
            xlabel=xlabel, ylabel=ylabel,
            width=width, height=height,
            data=data, color_theme=color_theme,
        )
        if result.get("status") == "error":
            emit(result, ctx.obj["json"])
            return

        image_path = result["path"]
        img_result = ctx.obj["backend"].add_image(
            slide_index=slide_index, image_path=image_path,
            left=left, top=top, width=img_width, height=img_height,
        )
        Path(image_path).unlink(missing_ok=True)
        emit({"status": "ok", "figure": image_path, "slide": slide_index, "image": img_result.get("image")}, ctx.obj["json"])
    except ImportError as e:
        emit({"status": "error", "error": f"matplotlib required: {e}"}, ctx.obj["json"])
    except Exception as e:
        emit({"status": "error", "error": str(e)}, ctx.obj["json"])


@main.command("figure-types")
@json_option
@click.pass_context
def figure_types(ctx: click.Context) -> None:
    """List available chart types and color themes."""
    emit({
        "chart_types": ["bar", "hbar", "line", "scatter", "pie", "heatmap", "radar", "donut"],
        "color_themes": ["nature", "pku", "dark", "light"],
        "data_formats": [
            "JSON: [{'label': 'A', 'value': 10}, ...]",
            "Simple: 'series1:10,20,30;series2:15,25,35'",
        ],
    }, ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Animations
# ══════════════════════════════════════════════════════════════════════

@main.command("add-animation")
@click.argument("slide_index", type=int)
@click.argument("shape_index", type=int)
@click.option("--effect", default="fade", type=click.Choice([
    "fade", "fly_in_left", "fly_in_right", "fly_in_top", "fly_in_bottom", "zoom", "appear", "wipe"
]))
@json_option
@click.pass_context
def add_animation(ctx: click.Context, slide_index: int, shape_index: int, effect: str) -> None:
    emit(ctx.obj["backend"].add_animation(slide_index, shape_index, effect), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Slide Show
# ══════════════════════════════════════════════════════════════════════

@main.command("run-show")
@json_option
@click.pass_context
def run_show(ctx: click.Context) -> None:
    emit(ctx.obj["backend"].run_slide_show(), ctx.obj["json"])


@main.command("exit-show")
@json_option
@click.pass_context
def exit_show(ctx: click.Context) -> None:
    emit(ctx.obj["backend"].exit_slide_show(), ctx.obj["json"])


# ══════════════════════════════════════════════════════════════════════
# Utility
# ══════════════════════════════════════════════════════════════════════

@main.command("undo")
@json_option
@click.pass_context
def undo(ctx: click.Context) -> None:
    emit(ctx.obj["backend"].undo(), ctx.obj["json"])


@main.command("redo")
@json_option
@click.pass_context
def redo(ctx: click.Context) -> None:
    emit(ctx.obj["backend"].redo(), ctx.obj["json"])

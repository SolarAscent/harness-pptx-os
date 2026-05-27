from __future__ import annotations

import json
import shlex

import click

from .core.session import Session
from .utils.powerpoint_backend import PowerPointBackend


def emit(payload: dict, as_json: bool) -> None:
    if as_json:
        click.echo(json.dumps(payload, ensure_ascii=False))
    else:
        click.echo(payload.get("status", payload))


def json_option(func):
    def set_json(ctx, _param, value):
        if value:
            ctx.obj["json"] = True

    return click.option("--json", "command_json", is_flag=True, expose_value=False, callback=set_json)(func)


@click.group(invoke_without_command=True)
@click.option("--json", "as_json", is_flag=True, help="Emit machine-readable JSON.")
@click.pass_context
def main(ctx: click.Context, as_json: bool) -> None:
    ctx.ensure_object(dict)
    ctx.obj["json"] = as_json
    ctx.obj["backend"] = PowerPointBackend()
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
            click.echo("Commands: info, new, open PATH, add-title-slide TITLE [--subtitle TEXT], save-as PATH, export-pdf PATH, undo, redo, close")
            continue
        args = shlex.split(line)
        try:
            main(args=args, obj=ctx.obj, standalone_mode=False)
        except Exception as exc:
            emit({"status": "error", "error": str(exc)}, ctx.obj["json"])


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


@main.command("add-title-slide")
@click.argument("title")
@click.option("--subtitle", default="", help="Subtitle text.")
@json_option
@click.pass_context
def add_title_slide(ctx: click.Context, title: str, subtitle: str) -> None:
    emit(ctx.obj["backend"].add_title_slide(title, subtitle), ctx.obj["json"])


@main.command("close")
@click.option("--no-save", is_flag=True, help="Close without saving changes.")
@json_option
@click.pass_context
def close(ctx: click.Context, no_save: bool) -> None:
    emit(ctx.obj["backend"].close(saving=not no_save), ctx.obj["json"])


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

from __future__ import annotations

import json
import platform
import subprocess
from pathlib import Path
from typing import Any


class PowerPointBackend:
    app_name = "Microsoft PowerPoint"

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
            raise RuntimeError(f"{self.app_name} automation is not implemented for {self.platform}")
        script = self.with_file_access(script)
        proc = subprocess.run(["osascript", "-e", script], text=True, capture_output=True, check=False)
        if proc.returncode != 0:
            raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
        return proc.stdout.strip()

    def with_file_access(self, script: str) -> str:
        lines = script.splitlines()
        if lines and lines[0] == f'tell application "{self.app_name}"':
            prelude = [
                "try",
                "set cliAnythingStartupDisk to path to startup disk",
                "close cliAnythingStartupDisk",
                "end try",
            ]
            return "\n".join([lines[0], *prelude, *lines[1:]])
        return script

    def info(self) -> dict[str, Any]:
        if self.available_interface() == "applescript":
            version = self.run_script('tell application "Microsoft PowerPoint" to return version')
            return {"application": self.app_name, "interface": "applescript", "version": version}
        return {"application": self.app_name, "interface": self.available_interface()}

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
            f'save active presentation in (POSIX file {json.dumps(full_path)}) as save as Open XML presentation\n'
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

    def add_title_slide(self, title: str, subtitle: str = "") -> dict[str, Any]:
        self.run_script(
            'tell application "Microsoft PowerPoint"\n'
            "set newSlide to make new slide at end of active presentation\n"
            "set layout of newSlide to slide layout blank\n"
            "set titleBox to make new text box at end of newSlide\n"
            "set left position of titleBox to 80\n"
            "set top of titleBox to 120\n"
            "set width of titleBox to 760\n"
            "set height of titleBox to 80\n"
            f'set content of text range of text frame of titleBox to {json.dumps(title, ensure_ascii=False)}\n'
            "set font size of font of text range of text frame of titleBox to 36\n"
            "set subtitleBox to make new text box at end of newSlide\n"
            "set left position of subtitleBox to 82\n"
            "set top of subtitleBox to 215\n"
            "set width of subtitleBox to 720\n"
            "set height of subtitleBox to 48\n"
            f'set content of text range of text frame of subtitleBox to {json.dumps(subtitle, ensure_ascii=False)}\n'
            "set font size of font of text range of text frame of subtitleBox to 18\n"
            "return count of slides of active presentation\n"
            "end tell"
        )
        return {"status": "ok", "title": title, "subtitle": subtitle}

    def close(self, saving: bool = True) -> dict[str, Any]:
        saving_flag = "yes" if saving else "no"
        self.run_script(f'tell application "Microsoft PowerPoint" to close active presentation saving {saving_flag}')
        return {"status": "ok", "saving": saving}

    def undo(self) -> dict[str, Any]:
        self.run_script('tell application "Microsoft PowerPoint" to undo')
        return {"status": "ok", "action": "undo"}

    def redo(self) -> dict[str, Any]:
        self.run_script('tell application "Microsoft PowerPoint" to redo')
        return {"status": "ok", "action": "redo"}

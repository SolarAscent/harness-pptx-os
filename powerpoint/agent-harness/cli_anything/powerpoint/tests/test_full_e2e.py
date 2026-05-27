from click.testing import CliRunner

from cli_anything.powerpoint.powerpoint_cli import main


def test_info_json(monkeypatch):
    class FakeBackend:
        def info(self):
            return {"application": "Microsoft PowerPoint", "interface": "applescript"}

    monkeypatch.setattr("cli_anything.powerpoint.powerpoint_cli.PowerPointBackend", FakeBackend)
    result = CliRunner().invoke(main, ["--json", "info"])
    assert result.exit_code == 0
    assert '"application": "Microsoft PowerPoint"' in result.output

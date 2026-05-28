from cli_anything.powerpoint.core.session import Session
from cli_anything.powerpoint.utils.powerpoint_backend import _parse_tab_records, _coerce_numbers


def test_session_records_actions():
    session = Session()
    session.record("open", path="/tmp/a.pptx")
    assert session.history == [{"action": "open", "path": "/tmp/a.pptx"}]


def test_parse_shape_records():
    rows = _parse_tab_records("1\tTitle\tshape type text box\t10\t20\n", ["index", "name", "type", "left", "top"])
    _coerce_numbers(rows[0], ["index", "left", "top"])
    assert rows[0]["index"] == 1
    assert rows[0]["name"] == "Title"
    assert rows[0]["left"] == 10

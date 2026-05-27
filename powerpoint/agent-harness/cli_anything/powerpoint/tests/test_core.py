from cli_anything.powerpoint.core.session import Session


def test_session_records_actions():
    session = Session()
    session.record("open", path="/tmp/a.pptx")
    assert session.history == [{"action": "open", "path": "/tmp/a.pptx"}]

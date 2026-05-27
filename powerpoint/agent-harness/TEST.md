# Test Plan

1. Unit test session state.
2. Test the Click command surface with backend calls mocked.
3. Optionally run a live smoke test on machines with PowerPoint automation access.

Commands:

```bash
python3 -m pip install -e .
python3 -m pytest cli_anything/powerpoint/tests
cli-anything-powerpoint info --json
```

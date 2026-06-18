import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

try:
    import app
except Exception as exc:
    print("Could not import the app.")
    print("Install dependencies first with: pip install -r requirements.txt")
    print(f"Error: {exc}")
    raise SystemExit(1)


def main():
    status = app.setup_status()
    print(json.dumps(status, indent=2))

    client = app.app.test_client()
    health = client.get("/healthz")
    setup = client.get("/setup-status")

    print("")
    print(f"/healthz returned HTTP {health.status_code}")
    print(f"/setup-status returned HTTP {setup.status_code}")

    if status["ok"]:
        print("")
        print("Setup check passed.")
        return 0

    print("")
    print("Setup check needs attention:")
    for error in status["errors"]:
        print(f"- {error}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

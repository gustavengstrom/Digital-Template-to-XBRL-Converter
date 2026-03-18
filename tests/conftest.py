"""Shared pytest configuration and hooks for the test suite."""

import json
from pathlib import Path


def pytest_addoption(parser):
    parser.addoption(
        "--snapshot-update",
        action="store_true",
        default=False,
        help="Update the survey snapshot file after running tests.",
    )


def pytest_sessionfinish(session, exitstatus):
    """After the test session, update the snapshot if --snapshot-update was passed."""
    if not session.config.getoption("--snapshot-update", default=False):
        return

    combined_path = (
        Path(__file__).parent.parent / "output" / "surveys" / "survey_data_all.json"
    )
    snapshot_path = Path(__file__).parent / "data" / "survey_snapshot.json"

    if not combined_path.exists():
        print(f"\nWARNING: {combined_path} not found — snapshot not updated.")
        return

    with open(combined_path, encoding="utf-8") as f:
        combined = json.load(f)

    snapshot: dict[str, list[dict]] = {}
    for sec in combined:
        name = sec.get("name", "")
        qs = sec.get("survey_data_proxy", [])
        snapshot[name] = [
            {
                "id": q["id"],
                "name": q["name"],
                "answer_type": q["answer_type"],
            }
            for q in qs
        ]

    snapshot_path.parent.mkdir(parents=True, exist_ok=True)
    with open(snapshot_path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, indent=2, ensure_ascii=False)
        f.write("\n")

    total = sum(len(v) for v in snapshot.values())
    print(
        f"\n✅ Snapshot updated: {snapshot_path} — "
        f"{len(snapshot)} sections, {total} questions."
    )

import json
import os


STATE_FILE = "data/widget_state.json"


def load_completed_items():
    """Return the set of Gmail message IDs completed in the widget."""

    if not os.path.exists(STATE_FILE):
        return set()

    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        return set(data.get("completed", []))

    except (json.JSONDecodeError, OSError):
        return set()


def save_completed_items(completed):
    """Save completed Gmail message IDs."""

    os.makedirs("data", exist_ok=True)

    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(
            {"completed": list(completed)},
            f,
            indent=4
        )


def mark_completed(message_id):
    completed = load_completed_items()

    completed.add(message_id)

    save_completed_items(completed)


def unmark_completed(message_id):
    completed = load_completed_items()

    completed.discard(message_id)

    save_completed_items(completed)
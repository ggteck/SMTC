"""Local preferences for the Open Order email page."""

import pickle
from datetime import date, timedelta
from pathlib import Path


PREFERENCES_PATH = Path(__file__).with_name("email_management_preferences.pkl")


def defaults():
    return {
        "folders": [],
        "destination": "",
        "since": date.today() - timedelta(days=1),
    }


def load_preferences(path=PREFERENCES_PATH):
    preferences = defaults()
    try:
        with open(path, "rb") as file:
            saved = pickle.load(file)
        if isinstance(saved, dict):
            preferences.update({key: saved[key] for key in preferences if key in saved})
    except (FileNotFoundError, EOFError, pickle.UnpicklingError):
        pass
    return preferences


def save_preferences(preferences, path=PREFERENCES_PATH):
    path = Path(path)
    temporary = path.with_suffix(".tmp")
    with open(temporary, "wb") as file:
        pickle.dump(preferences, file)
    temporary.replace(path)

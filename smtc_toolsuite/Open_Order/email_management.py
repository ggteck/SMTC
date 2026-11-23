"""Streamlit page for downloading attachments from selected Outlook folders."""

import sys
import re
from pathlib import Path

import streamlit as st

from Open_Order.outlook_worker_client import run_outlook
from Open_Order.preferences import load_preferences, save_preferences


def choose_destination(current):
    from tkinter import Tk, filedialog

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        initial = current if current and Path(current).is_dir() else str(Path.home())
        return filedialog.askdirectory(parent=root, initialdir=initial)
    finally:
        root.destroy()


st.title("Email management")
st.write("Choose Outlook mail folders and a local folder for their attachments.")

if sys.platform != "win32":
    st.info("This page requires Windows with the Outlook desktop app installed.")
    st.stop()

preferences = load_preferences()
try:
    available = run_outlook({"action": "list"})["folders"]
except Exception as exc:
    st.error(f"Could not connect to Outlook: {exc}")
    st.stop()

if not available:
    st.warning("No Outlook mail folders are available.")
    st.stop()

selected = st.multiselect(
    "Mail folders",
    options=list(available),
    default=[key for key in preferences["folders"] if key in available],
    format_func=lambda key: available[key],
)
destination = preferences["destination"]
if st.button("Choose attachment destination folder"):
    try:
        destination = choose_destination(destination) or destination
    except Exception as exc:
        st.error(f"Could not open the folder picker: {exc}")
st.write(f"Destination: {destination or 'No folder selected'}")
since = st.date_input("Messages received since", value=preferences["since"])
filename_regex = st.text_input(
    "Attachment filename regex (optional)",
    value=preferences["filename_regex"],
    help="Matches the original attachment filename. Leave blank to download all attachments.",
)
try:
    re.compile(filename_regex)
    regex_error = None
except re.error as exc:
    regex_error = str(exc)
    st.error(f"Invalid filename regex: {regex_error}")

updated_preferences = {
    "folders": selected,
    "destination": destination,
    "since": since,
    "filename_regex": filename_regex,
}
preference_error = False
if updated_preferences != preferences:
    try:
        save_preferences(updated_preferences)
        preferences = updated_preferences
    except OSError as exc:
        preference_error = True
        st.error(f"Could not save preferences: {exc}")

if st.button("Save attachments", type="primary", disabled=regex_error is not None or preference_error):
    target = Path(preferences["destination"]).expanduser() if preferences["destination"] else None
    if not preferences["folders"]:
        st.error("Select at least one mail folder and save preferences first.")
    elif target is None or not target.is_dir():
        st.error("Select an existing attachment destination folder and save preferences first.")
    else:
        try:
            st.write("Reading Outlook messages...")
            result = run_outlook({
                "action": "save",
                "folders": preferences["folders"],
                "destination": str(target),
                "since": preferences["since"].isoformat(),
                "filename_regex": preferences["filename_regex"],
            })
            saved, skipped, errors = result["saved"], result["skipped"], result["errors"]
            st.success(f"Saved {saved} attachments to {target}. Skipped {skipped} already saved files.")
            for error in errors:
                st.warning(error)
        except Exception as exc:
            st.error(f"Could not read Outlook messages: {exc}")

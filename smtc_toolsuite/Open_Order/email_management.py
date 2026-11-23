"""Streamlit page for downloading attachments from selected Outlook folders."""

import sys
from datetime import date
from pathlib import Path

import streamlit as st

from Open_Order.outlook_mail import list_folders, outlook_namespace, save_attachments
from Open_Order.preferences import load_preferences, save_preferences


st.title("Email management")
st.write("Choose Outlook mail folders and a local folder for their attachments.")

if sys.platform != "win32":
    st.info("This page requires Windows with the Outlook desktop app installed.")
    st.stop()

preferences = load_preferences()
try:
    with outlook_namespace() as outlook:
        available = list_folders(outlook)
except Exception as exc:
    st.error(f"Could not connect to Outlook: {exc}")
    st.stop()

if not available:
    st.warning("No Outlook mail folders are available.")
    st.stop()

with st.form("email_preferences"):
    selected = st.multiselect(
        "Mail folders",
        options=list(available),
        default=[key for key in preferences["folders"] if key in available],
        format_func=lambda key: available[key],
    )
    destination = st.text_input("Attachment destination folder", value=preferences["destination"])
    since = st.date_input("Messages received since", value=preferences["since"])
    submitted = st.form_submit_button("Save preferences")

if submitted:
    preferences = {"folders": selected, "destination": destination.strip(), "since": since}
    try:
        save_preferences(preferences)
        st.success("Preferences saved.")
    except OSError as exc:
        st.error(f"Could not save preferences: {exc}")

if st.button("Save attachments", type="primary"):
    target = Path(preferences["destination"]).expanduser() if preferences["destination"] else None
    if not preferences["folders"]:
        st.error("Select at least one mail folder and save preferences first.")
    elif target is None or not target.is_dir():
        st.error("Select an existing attachment destination folder and save preferences first.")
    else:
        try:
            st.write("Reading Outlook messages...")
            with outlook_namespace() as outlook:
                saved, skipped, errors = save_attachments(
                    outlook, preferences["folders"], target, preferences["since"]
                )
            st.success(f"Saved {saved} attachments to {target}. Skipped {skipped} already saved files.")
            for error in errors:
                st.warning(error)
        except Exception as exc:
            st.error(f"Could not read Outlook messages: {exc}")

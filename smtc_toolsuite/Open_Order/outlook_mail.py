"""Read Outlook folders and save attachments through the installed desktop client."""

import hashlib
import ntpath
import re
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path


@contextmanager
def outlook_namespace():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        yield win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    finally:
        pythoncom.CoUninitialize()


def folder_key(folder):
    return f"{folder.StoreID}|{folder.EntryID}"


def list_folders(namespace):
    """Return visible folders across all Outlook stores, including subfolders."""
    result = {}

    def visit(folder, parts):
        try:
            key = folder_key(folder)
            result[key] = " / ".join((*parts, folder.Name))
            children = folder.Folders
        except Exception:
            return
        for child in children:
            visit(child, (*parts, folder.Name))

    for root in namespace.Folders:
        visit(root, ())
    return dict(sorted(result.items(), key=lambda item: item[1].casefold()))


def safe_filename(filename):
    name = ntpath.basename(str(filename)).strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).rstrip(" .")
    return name or "attachment"


def save_attachments(namespace, folder_keys, destination, since, filename_regex=""):
    """Save attachments from direct messages in selected folders; return counts and errors."""
    destination = Path(destination)
    cutoff = datetime.combine(since, datetime.min.time())
    filename_pattern = re.compile(filename_regex) if filename_regex else None
    saved = skipped = 0
    errors = []

    def report_error(message):
        if len(errors) < 10:
            errors.append(message)
        elif len(errors) == 10:
            errors.append("Additional Outlook errors were omitted from this page.")

    for key in folder_keys:
        try:
            store_id, entry_id = key.split("|", 1)
            folder = namespace.GetFolderFromID(entry_id, store_id)
            items = folder.Items
        except Exception as exc:
            report_error(f"Could not open a selected folder: {exc}")
            continue

        # Sorting is handled by Outlook. For a large mailbox we can stop as soon
        # as the first older message is reached instead of loading its history.
        sorted_by_date = False
        try:
            items.Sort("[ReceivedTime]", True)
            sorted_by_date = True
        except (AttributeError, TypeError):
            pass  # Also supports simple iterable mail collections in tests.
        except Exception as exc:
            report_error(f"Could not sort {folder.Name} by received date: {exc}")

        try:
            message_count = items.Count
        except AttributeError:
            messages = iter(items)
        else:
            def get_messages():
                for index in range(1, message_count + 1):
                    try:
                        yield items.Item(index)
                    except Exception as exc:
                        report_error(f"Could not read message {index} in {folder.Name}: {exc}")

            messages = get_messages()

        for message in messages:
            try:
                received = message.ReceivedTime
                received = datetime.fromtimestamp(received.timestamp())
                if received < cutoff:
                    if sorted_by_date:
                        break
                    continue
                attachments = message.Attachments
                message_id = message.EntryID
                count = attachments.Count
            except Exception:
                continue  # Non-mail items need not have a received date or attachments.

            for index in range(1, count + 1):
                try:
                    attachment = attachments.Item(index)
                    raw_name = ntpath.basename(str(attachment.FileName))
                    if filename_pattern and not filename_pattern.search(raw_name):
                        continue
                    original = safe_filename(raw_name)
                    stem, suffix = ntpath.splitext(original)
                    token = hashlib.sha256(f"{store_id}|{message_id}|{index}".encode()).hexdigest()[:12]
                    filename = f"{stem}_{received:%Y-%m-%d-%H%M%S}_{token}{suffix}"
                    target = destination / filename
                    if target.exists():
                        skipped += 1
                        continue
                    attachment.SaveAsFile(str(target))
                    saved += 1
                except Exception as exc:
                    report_error(f"Could not save an attachment in {folder.Name}: {exc}")

    return saved, skipped, errors

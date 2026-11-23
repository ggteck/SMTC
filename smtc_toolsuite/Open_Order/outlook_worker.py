"""Isolated Outlook COM process used by the Email management page."""

import json
import sys
import traceback
from datetime import date

from outlook_mail import list_folders, outlook_namespace, save_attachments


def handle_request(request):
    action = request["action"]
    with outlook_namespace() as outlook:
        if action == "list":
            return {"folders": list_folders(outlook)}
        if action == "save":
            saved, skipped, errors = save_attachments(
                outlook,
                request["folders"],
                request["destination"],
                date.fromisoformat(request["since"]),
            )
            return {"saved": saved, "skipped": skipped, "errors": errors}
    raise ValueError(f"Unknown Outlook action: {action}")


def main():
    try:
        request = json.load(sys.stdin)
        result = handle_request(request)
        print(json.dumps(result), flush=True)
    except Exception:
        traceback.print_exc(file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

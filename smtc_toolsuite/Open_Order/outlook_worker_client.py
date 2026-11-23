"""Run Outlook COM work separately from the Streamlit server process."""

import json
import subprocess
import sys
from pathlib import Path


WORKER_PATH = Path(__file__).with_name("outlook_worker.py")


def run_outlook(request, worker_path=WORKER_PATH, timeout=900):
    try:
        result = subprocess.run(
            [sys.executable, str(worker_path)],
            input=json.dumps(request),
            text=True,
            capture_output=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("Outlook did not respond before the 15-minute limit.") from exc

    if result.returncode:
        exit_code = f"0x{result.returncode & 0xFFFFFFFF:08X}"
        details = result.stderr.strip()[-2000:]
        raise RuntimeError(f"Outlook worker exited with code {exit_code}. {details}".strip())
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError("Outlook worker returned an invalid result.") from exc

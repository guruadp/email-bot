#!/usr/bin/env python3
"""Runs both bots in one Render Background Worker.

- email_bot.py: long-running subprocess, restarted automatically if it exits.
- lead_automation.py: one cycle every LEAD_POLL_MINUTES (default 15).

Set the Render start command to:  python3 start.py
"""

import os
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
LEAD_INTERVAL = int(os.getenv("LEAD_POLL_MINUTES", "15")) * 60
# Generous timeout so a first-run device-code sign-in (done by watching the
# logs) has time to complete; normal cycles finish in seconds.
LEAD_TIMEOUT = int(os.getenv("LEAD_RUN_TIMEOUT_SECONDS", "1500"))


def log(msg):
    print(f"[start] {msg}", flush=True)


def main():
    email_bot = None
    next_lead_run = 0.0
    while True:
        if email_bot is None or email_bot.poll() is not None:
            if email_bot is not None:
                log(f"email_bot exited with code {email_bot.returncode}; restarting in 5s")
                time.sleep(5)
            email_bot = subprocess.Popen([sys.executable, str(BASE / "email_bot.py")])
            log(f"email_bot started (pid {email_bot.pid})")

        if time.time() >= next_lead_run:
            log("running lead_automation cycle")
            try:
                result = subprocess.run(
                    [sys.executable, str(BASE / "lead_automation.py")],
                    timeout=LEAD_TIMEOUT,
                )
                if result.returncode != 0:
                    log(f"lead_automation exited with code {result.returncode}")
            except subprocess.TimeoutExpired:
                log(f"lead_automation timed out after {LEAD_TIMEOUT}s")
            next_lead_run = time.time() + LEAD_INTERVAL

        time.sleep(15)


if __name__ == "__main__":
    main()

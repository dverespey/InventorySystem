"""Shared helpers for the Perspective E2E harness (Playwright, Python).

No third-party deps beyond playwright. Reads optional gateway creds from
environment or a gitignored scripts/e2e/.env (KEY=VALUE lines).
"""
import os, re, subprocess, time

BASE = os.environ.get("GW_BASE", "http://localhost:8088")
WRAPPER_LOG = os.environ.get("GW_LOG", "/usr/local/ignition/logs/wrapper.log")
ARTIFACTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "artifacts")


def load_env():
    """Populate os.environ from scripts/e2e/.env if present (no override of real env)."""
    envp = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(envp):
        for line in open(envp):
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


def creds():
    load_env()
    return os.environ.get("GATEWAY_USER"), os.environ.get("GATEWAY_PASS")


def log_marker():
    """Byte offset of the end of wrapper.log — capture before an action, then
    read_log_since(offset) to get only lines the action produced."""
    try:
        return os.path.getsize(WRAPPER_LOG)
    except OSError:
        return 0


def read_log_since(offset):
    try:
        with open(WRAPPER_LOG, "r", errors="replace") as f:
            f.seek(offset)
            return f.read()
    except OSError:
        return ""


def grep_spike_since(offset, needle="SPIKE"):
    return [l for l in read_log_since(offset).splitlines() if needle in l]


class Report:
    """Accumulates per-check PASS/FAIL/SKIP and prints a summary + exit code."""
    def __init__(self):
        self.rows = []

    def check(self, name, ok, detail=""):
        status = "PASS" if ok else "FAIL"
        self.rows.append((status, name, detail))
        print("  [%s] %s%s" % (status, name, ("  — " + detail) if detail else ""))
        return ok

    def skip(self, name, why):
        self.rows.append(("SKIP", name, why))
        print("  [SKIP] %s  — %s" % (name, why))

    def summary_exit(self):
        p = sum(1 for r in self.rows if r[0] == "PASS")
        f = sum(1 for r in self.rows if r[0] == "FAIL")
        s = sum(1 for r in self.rows if r[0] == "SKIP")
        print("\n===== E2E SUMMARY: %d PASS / %d FAIL / %d SKIP =====" % (p, f, s))
        return 1 if f else 0

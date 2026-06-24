#!/usr/bin/env python3
"""_ignenv.py — ONE env-parametrized home for the Ignition gateway-path + DB-connection
constants shared by the CPython tooling (generators + e2e harness).

Why this exists (repo-split-plan.md §4.C/§4.D + adversary B2/N1): the gateway path and the
`Inventory_Spike` DB-connection name were hardcoded across 8 generators (no env fallback), 2
already-env generators, and the e2e harness. Centralizing them here makes the spike->prod
rename a SINGLE edit point (an env var), and keeps every tool consistent.

BEHAVIOR-PRESERVING: the defaults below ARE the current spike values, so nothing changes at
runtime unless an env var is set. The documented prod override (David's cut-time decision) is:

    IGN_DB_CONN=Inventory        # prod DB-connection name (replaces Inventory_Spike)

The project name stays `InventorySystem`; the repo will become `inventory` at the cut.

Pure stdlib; importable from both scripts/gen_*.py and scripts/e2e/*.py.
"""
import os

# --- gateway / project ---------------------------------------------------------------------
# GATEWAY_ROOT: the Ignition install root on the dev/CI box (8.1.52 here).
GATEWAY_ROOT = os.environ.get("IGN_GATEWAY_ROOT", "/usr/local/ignition")

# PROJECT: the Ignition Perspective project name. Stays `InventorySystem` through the split
# (David's decision); overridable for a future rename. (e2e/lib.py also reads GW_PROJECT for
# the client-route segment; both default to the same value.)
PROJECT = os.environ.get("IGN_PROJECT", os.environ.get("GW_PROJECT", "InventorySystem"))

# GATEWAY_PROJECT_DIR: the on-disk project-as-code root, derived from PROJECT + GATEWAY_ROOT,
# but fully overridable in one shot via IGN_PROJECT_DIR (mirrors the legacy PROJ_DIR env the
# 2 already-env generators used).
GATEWAY_PROJECT_DIR = os.environ.get(
    "IGN_PROJECT_DIR",
    os.environ.get("PROJ_DIR", os.path.join(GATEWAY_ROOT, "data", "projects", PROJECT)),
)

# Convenience: the Perspective resource root under the project dir (where view.json lives).
PERSPECTIVE_DIR = os.path.join(GATEWAY_PROJECT_DIR, "com.inductiveautomation.perspective")

# The gateway wrapper.log (e2e log-tail asserts). Mirrors lib.py's GW_LOG default.
GATEWAY_LOG = os.environ.get("GW_LOG", os.path.join(GATEWAY_ROOT, "logs", "wrapper.log"))

# --- DB connection -------------------------------------------------------------------------
# DB_CONN: the Ignition DB-connection (datasource) name the Jython app code + generated views
# bind through. SINGLE prod-rename edit point on the CPython side:
#     IGN_DB_CONN=Inventory   (prod)         /   Inventory_Spike (spike default, unchanged)
# (DBCONN kept as a legacy alias so the existing `DBCONN=` override still works.)
DB_CONN = os.environ.get("IGN_DB_CONN", os.environ.get("DBCONN", "Inventory_Spike"))

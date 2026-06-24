# db_shared — Project Library: the ONE place the Ignition DB-connection (datasource) name lives.
#
# Every app-code module (edi810/edi856/edi_inbound/forecast/asn/hotcall/stockLedger/order/renban/
# order_file/forecast_distribution/auto_purge) binds its `system.db.*` calls to a named gateway
# connection. They used to each hardcode `DATABASE = "Inventory_Spike"`. This module centralizes that
# string so the spike->prod rename (Inventory_Spike -> Inventory) is a SINGLE edit point here, not 13.
#
# Usage in an app module (preserves the existing module-level `DATABASE` symbol, so the ~50 downstream
# `db = database if database is not None else DATABASE` / `def f(..., database=DATABASE)` references are
# untouched):
#
#     from db_shared import CONNECTION as DATABASE
#
# PROD RENAME (David's cut-time decision): change the one line below to "Inventory". On the CPython
# harness side the same value comes from scripts/_ignenv.py (env IGN_DB_CONN); both must agree.
#
# Jython 2.7 (8.1.52 and 8.3 — no scripting delta). Pure constant; no `system` / gateway API used, so it
# loads identically on the gateway and under the headless jython_shim.

CONNECTION = "Inventory_Spike"   # the named Ignition DB-connection; SINGLE prod-rename edit point.

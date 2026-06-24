# report_render/driver — the THIN gateway entry for the M3 render harness (the edi810/856 driver pattern).
#
# run_report(report_key, params, out_dir, ...) : resolve the report's .sql (faithful↔corrected seam) ->
# run it (system.db.runPrepQuery) -> build the cell plan (declarative or the report's custom pure plan) ->
# render the .xlsx (Apache POI) -> write tmp-then-rename to out_dir with the legacy yyyymmddhhmmss00
# filename. NO PrintOut — server-side render REPLACES the failing client Excel/printer path.
#
# `system` is a gateway global at call time (never imported) so this module imports cleanly under
# scripts/e2e/jython_shim.py for the headless e2e (the R8 runner). The .sql files live on disk as the
# source-of-truth NQ substitute (the proven Order/Supplier spike shape); on 8.3 each .sql is promoted to a
# real Named Query and load_sql/runPrepQuery becomes runNamedQuery (IG83-TODO, no driver-logic change).
#
# Render dispatch is DATA-DRIVEN by the def's `render` key:
#   "declarative"               -> code.build_declarative_plan  (Daily Shipping Assy, INVOICE Summary)
#   "build_forecast_detail_plan"-> code.build_forecast_detail_plan
#   "build_lot_location_plan"   -> code.build_lot_location_plan  (special: two .sql -> {"wheels","tires"})

import os as _os

# In the gateway these are sibling Project-Library modules (report_render.code / report_render.report_defs).
# For the headless shim/unit-test we load by file path; this dual-import keeps both paths working.
try:
    import code as _engine
    import report_defs as _defs
except ImportError:                                  # pragma: no cover (gateway package import)
    from report_render import code as _engine
    from report_render import report_defs as _defs


# ---- .sql source-of-truth loader (the NQ substitute on the spike) -------------------------------
# The .sql NQ-substitute files live under the repo's db/namedqueries/reporting/ tree (repo reorg). From
# this module (ignition/project-library/report_render/) that is three dirs up to the repo root, then into
# db/namedqueries/reporting. IG83-TODO: on prod each .sql is a promoted Named Query (runNamedQuery), so
# this on-disk resolution goes away entirely.
_SQL_DIR = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)),
                                           "..", "..", "..", "db", "namedqueries", "reporting"))


def load_sql(sql_key):
    """Read sql/<sql_key>.sql (stripping the leading -- comment header so only the statement is bound).
    IG83-TODO: on prod each .sql is a promoted Named Query -> replace this + runPrepQuery with
    system.db.runNamedQuery(path, paramsDict)."""
    path = _os.path.join(_SQL_DIR, sql_key + ".sql")
    with open(path) as fh:
        lines = fh.read().splitlines()
    # Drop pure-comment lines AND strip trailing inline -- comments, so the bound text is just the SELECT
    # (params bind by ? order; a '?' inside a comment must NOT be treated as a bind marker). No '?' appears
    # inside any string literal in these queries, so a plain index split is safe.
    body = []
    for ln in lines:
        if ln.strip().startswith("--"):
            continue
        if "--" in ln:
            ln = ln[:ln.index("--")]
        body.append(ln)
    return "\n".join(body).strip()


def _args_for(report_def, params):
    """Bind params in the def's declared `params` order to the .sql ? markers (the NQ name-binding port)."""
    p = params or {}
    return [p.get(name) for name in report_def.get("params", [])]


def run_report(report_key, params=None, out_dir=None, database=None, query_variant=None, now=None):
    """Run ONE M3 report end to end. Returns {"path", "rows", "key", "variant"}.

    params: a dict of the report's params (e.g. {"PDate": "20250121"}) PLUS any render-only params the
            custom plans need (PlantName/AssemblerName for Lot Location — gateway config, not a client INI).
    out_dir: the server-side fiReportsOutputDir (gateway config). query_variant: per-call faithful/corrected
            override (else report_defs.QUERY_VARIANT). now: inject a fixed clock for deterministic tests."""
    rdef = _defs.REPORTS[report_key]
    render = rdef.get("render", "declarative")

    # --- LOT LOCATION: special two-.sql dispatch (wheels proc then tires proc into one sheet) ----------
    if render == "build_lot_location_plan":
        qmap = rdef["queries"]
        wheels = _engine.build_dataset(load_sql(qmap["wheels"]), [], database)
        tires = _engine.build_dataset(load_sql(qmap["tires"]), [], database)
        cells, widths = _engine.build_lot_location_plan(rdef, {"wheels": wheels, "tires": tires},
                                                        params=params)
        xlsx = _engine.render_xlsx(cells, widths)
        out_name = (params or {}).get("PlantName", "") + rdef["out_name"]   # legacy <PlantName>LotLocation
        path = _engine.write_output(xlsx, out_name, out_dir, now)
        rows = len(wheels) + len(tires)
        _engine._log("report/%s: %d rows (wheels=%d tires=%d) -> %s"
                     % (report_key, rows, len(wheels), len(tires), path))
        return {"path": path, "rows": rows, "key": report_key, "variant": "faithful"}

    # --- DECLARATIVE + FORECAST DETAIL: one detail .sql (+ optional header .sql) -----------------------
    sql_key = _defs.resolve_query(rdef["query"], query_variant)
    detail = _engine.build_dataset(load_sql(sql_key), _args_for(rdef, params), database)

    header_row = None
    if rdef.get("header_query"):
        # the header band query reuses the report's params order (Daily Shipping Assy: @PDate)
        hdr = _engine.build_dataset(load_sql(rdef["header_query"]), _args_for(rdef, params), database)
        header_row = hdr[0] if len(hdr) > 0 else None

    if render == "declarative":
        cells, widths = _engine.build_declarative_plan(rdef, detail, header_row, params)
    elif render == "build_forecast_detail_plan":
        cells, widths = _engine.build_forecast_detail_plan(rdef, detail, header_row, params)
    else:
        raise ValueError("report_render.driver: unknown render %r for %s" % (render, report_key))

    xlsx = _engine.render_xlsx(cells, widths)
    path = _engine.write_output(xlsx, rdef["out_name"], out_dir, now)
    n = len(detail)
    variant = query_variant or _defs.QUERY_VARIANT
    _engine._log("report/%s[%s]: %d rows -> %s" % (report_key, variant, n, path))
    return {"path": path, "rows": n, "key": report_key, "variant": variant}

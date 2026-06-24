#!/usr/bin/env python3
"""test_edi_inbound_e2e.py — END-TO-END exercise of the REAL edi_inbound 997/824 gateway DRIVER, driven
headlessly through the system.db/system.file shim (jython_shim) — the M1 inbound loop-closer proof.

WHAT THIS PROVES (running the ACTUAL driver code, NOT a proc-EXEC — retro R8)
----------------------------------------------------------------------------
Runs process_inbound / process_one_file / process_997 / process_824 against the spike DB inside the
per-file transaction the driver itself opens:

  (1) 997 SH ACCEPT -> the synthetic 'S' ASN flips to 'A' (site-scoped UPDATE_EINStatus wrap), matched
      by EIN; a 997 IN REJECT -> the synthetic 'S' invoice flips to 'R'. The flip is keyed by EIN+type.
  (2) 997 UNKNOWN-EIN -> @@ROWCOUNT is 0 -> an ALARM row (IN_EDI_ALARM_REJ, '997_UNKNOWN_EIN') is
      written; it is NOT a silent success (the legacy no-@@ROWCOUNT bug fix).
  (3) 824 -> the referenced ASN (matched by manifest on INV_ASN_DETAIL_MST -> parent INV_ASN_MST) flips
      to 'R', AND an alarm row per reject line captures manifest/part/errorText (Q10).
  (4) DUNS GUARD -> a file whose ISA delSL[4] matches no INV_SITES.VC_TMM_DUNS is QUARANTINED (a ledger
      row VC_STATUS='QUARANTINED', file moved to the Quarantine dir), NOT silently dropped (Q11).
  (5) IDEMPOTENT RE-PROCESS -> dropping the SAME 997 twice processes it once (the second poll is a
      ledger NO-OP); the status flip is not double-applied / no duplicate alarms.

HONEST VERIFICATION (no golden inbound sample)
----------------------------------------------
NO captured TEMA 997/824 exists on disk, so BYTE/OFFSET PARITY VS A REAL TEMA FILE IS UNPROVABLE — not
claimed. The fixtures are built to the legacy EDIUpload.pas copy() offsets (the oracle) + the outbound
856 ISA shape (TEMA-as-sender inbound mirror). What IS proven end-to-end: the STATUS-FLOW (status flips
keyed by EIN+type / manifest, site-resolved by DUNS), the @@ROWCOUNT unknown-EIN alarm, the 824 reject
flag + detail, the DUNS quarantine, and idempotency — all through the REAL driver + DB. Pending: confirm
the AK1/NTE/ISA offsets against a real inbound sample (spec §2.3 #1, §3.1, §5.2).

Live data note (einstatus-status-flow-analysis.md §2): Inventory_Live is ALL 'A' status — no 'C'/'S'/'R'
exist to exercise the flip naturally. So we INJECT synthetic 'S' ASN/invoice rows (tagged VC_ADD so
cleanup is exact) and assert the flip off those. The spike is restored as-found.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_edi_inbound_e2e.py
"""
import os
import sys
import shutil
import tempfile
import subprocess

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report, preclean_sentinels, DB_CONN   # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"

SITE_ID = 1
TMM_DUNS = "000000011"          # INV_SITES site 1 VC_TMM_DUNS (the trading-partner DUNS at ISA delSL[4])
BAD_DUNS = "999999999"          # matches no site -> quarantine

# sentinels so cleanup is exact and nothing real is touched.
TEST_LINE = "ZZINB856"          # the synthetic 'S' ASN line name
INV_ADD = "ZZINB810INVSENT0"    # 16-char VC_ADD stamp on the synthetic 'S' invoice
ASN_ADD = "ZZINB856ASNSENT0"    # 16-char VC_ADD stamp on the synthetic 'S' ASN + detail
MANIFEST = "ZZINB001"           # the 824-referenced manifest (8 chars)
PART = "ZZINBPART000"           # 12 chars
# FIX-B multi-ASN: a manifest carried on TWO distinct ASNs (Live proves up to 3 per manifest).
MANIFEST_FANOUT = "ZZINB777"    # 8 chars — shared across 2 synthetic ASNs
ASN_FAN1_ADD = "ZZINB856FAN0SEN0"  # 16-char VC_ADD on the 1st fan-out ASN
ASN2_ADD = "ZZINB856ASN2SEN0"   # 16-char VC_ADD stamp on the 2nd synthetic ASN + its detail

EDI_PY = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..",
    "ignition", "project-library", "edi_inbound", "code.py"))


def sql(db, query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(db, q):
    rows = sql(db, q)
    return rows[0][0] if rows else None


# --- inbound-file fixture builders (legacy offsets + the inbound TEMA-as-sender ISA shape) ----------

def inbound_isa(tmmDuns):
    """The inbound ISA — TEMA as sender. We PLACE the DUNS at split index 4 (delSL[4]) to match how the
    code reads it.

    !!! HONESTY CAVEAT (adversary BLOCKER-1 / KEY FINDING #2): this HARD-CODES the TMM DUNS into the
    delSL[4] slot, so the DUNS-guard cases (4 below: a matching DUNS routes, a non-matching DUNS
    quarantines) test the guard MECHANICS only — they are SELF-CONSISTENT and do NOT prove which X12
    element a REAL inbound TEMA ISA carries the routing DUNS on. On our OWN outbound ISA delSL[4]=ISA04
    is OUR SiteDUNS and the TMM DUNS is at delSL[8]/ISA08; the inbound element is UNPROVABLE without a
    captured golden TEMA file (same pending-golden stance as 856/810 byte-parity). The CODE matches
    VC_TMM_DUNS legacy-faithfully (== AD_GetSiteTMMDUNS); only the element index is the open gap."""
    return ("ISA*00*          *01*" + ("%-10s" % tmmDuns)[:10] +
            "*ZZ*               *01*               *260618*1430*U*00400*000000777*0*P*>")


def ak1(fnId, ein):
    return "AK1*" + ("%-2s" % fnId)[:2] + "*" + ("%09d" % int(ein))


def ak9(code):
    return "AK9*" + (str(code) if code else " ") + "*1*1*0"


def nte(errText, manifest, part):
    line = "NTE*GEN*"
    line += ("%-50s" % errText)[:50]
    line += "*"
    line += ("%-8s" % manifest)[:8]
    line += "*"
    line += ("%-12s" % part)[:12]
    return line


def file_997(tmmDuns, acks):
    """acks: list of (fnId, ein, code-or-None). None code -> no AK9 (a malformed group)."""
    segs = [inbound_isa(tmmDuns), "GS*FA*X*Y*20260618*1430*777*X*004010", "ST*997*0001"]
    for fnId, ein, code in acks:
        segs.append(ak1(fnId, ein))
        if code is not None:
            segs.append(ak9(code))
    segs += ["SE*%d*0001" % (len(segs) - 1), "GE*1*777", "IEA*1*000000777"]
    return "\r\n".join(segs) + "\r\n"


def file_824(tmmDuns, refs):
    segs = [inbound_isa(tmmDuns), "GS*AG*X*Y*20260618*1430*777*X*004010", "ST*824*0001",
            "BGN*11*REF*20260618"]
    segs += [nte(e, mf, p) for (e, mf, p) in refs]
    segs += ["SE*%d*0001" % (len(segs) - 1), "GE*1*777", "IEA*1*000000777"]
    return "\r\n".join(segs) + "\r\n"


def _sentinel_statements():
    """The ordered (child-before-parent), sentinel-scoped DELETEs that remove EXACTLY this suite's
    synthetic rows (every predicate keyed on a ZZINB* VC_ADD / manifest / line / file-name sentinel).
    Shared by cleanup() (end) AND the start-of-suite self-heal preclean (P11)."""
    return [
        "DELETE FROM INV_ASN_DETAIL_MST WHERE VC_ADD IN ('%s','%s','%s') "
        "OR VC_MANIFEST_NUMBER IN ('%s','%s')"
        % (ASN_ADD, ASN_FAN1_ADD, ASN2_ADD, MANIFEST, MANIFEST_FANOUT),
        "DELETE FROM INV_ASN_MST WHERE VC_ADD IN ('%s','%s','%s') OR VC_LINE_NAME = '%s'"
        % (ASN_ADD, ASN_FAN1_ADD, ASN2_ADD, TEST_LINE),
        "DELETE FROM INV_INV_MST WHERE VC_ADD = '%s'" % INV_ADD,
        # the ledger + alarm rows this test created (by the sentinel manifest/part or our file names).
        "DELETE FROM INV_EDI_ALARM_REJ WHERE VC_MANIFEST_NUMBER IN ('%s','%s') "
        "OR VC_SUMMARY LIKE '%%ZZINB%%' OR VC_ALARM_TYPE = '997_UNKNOWN_EIN'"
        % (MANIFEST, MANIFEST_FANOUT),
        "DELETE FROM INV_EDI_INBOUND_LOG WHERE VC_FILE_NAME LIKE 'zzinb_%'",
    ]


def cleanup():
    """Restore the spike as-found: drop the synthetic ASN/invoice/detail + the ledger/alarm rows we made."""
    for stmt in _sentinel_statements():
        sql(INV, stmt)


def main():
    print("=" * 92)
    print(" edi_inbound 997/824 END-TO-END — status flow + @@ROWCOUNT alarm + 824 reject + DUNS guard")
    print(" + idempotency (REAL driver via the shim; NOT a proc-EXEC). No golden sample — see docstring.")
    print("=" * 92)
    rep = Report()
    edi = jython_shim.load_wrapper("edi_inbound_driver", EDI_PY)

    # P11 self-heal: pre-clean leftovers from a KILLED prior run via the shared error-swallowing helper
    # (so a partially-dirty DB still drains even if that run's cleanup() never ran / aborted mid-way).
    preclean_sentinels(lambda s: sql(INV, s), _sentinel_statements(), label="edi_inbound")

    tmpdir = tempfile.mkdtemp(prefix="edi_inbound_e2e_")
    archiveDir = os.path.join(tmpdir, "Archive")
    quarantineDir = os.path.join(tmpdir, "Quarantine")
    asnId = invId = None
    try:
        # --- 0. ground truth: the M1 tables exist + site 1 has the TMM DUNS we feed -----------------
        rep.check("INV_EDI_INBOUND_LOG + INV_EDI_ALARM_REJ exist (spike-edi-inbound-tables.sql applied)",
                  scalar(INV, "SELECT COUNT(*) FROM sys.tables WHERE name IN "
                              "('INV_EDI_INBOUND_LOG','INV_EDI_ALARM_REJ')") == "2", "both present")
        siteDuns = scalar(INV, "SELECT VC_TMM_DUNS FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
        rep.check("site %d VC_TMM_DUNS == %r (the DUNS we put at ISA delSL[4])" % (SITE_ID, TMM_DUNS),
                  (siteDuns or "").strip() == TMM_DUNS, "got %r" % siteDuns)

        # --- 1. inject a synthetic 'S' ASN (with a detail row carrying the 824 manifest) + a 'S' inv --
        # ASN: status 'S' (sent, awaiting ack), a KNOWN EIN we will ack. Live is all 'A' (§2) so we make
        # our own 'S' rows; tag VC_ADD for exact cleanup.
        ASN_EIN = 880001
        INV_EIN = 880002
        UNKNOWN_EIN = 889999      # an EIN that exists on NO row -> the unknown-EIN alarm
        sql(INV,
            "INSERT INTO INV_ASN_MST (IN_ASN_EIN, VC_ASN_STATUS, VC_LINE_NAME, VC_ASSEMBLY_LINE, "
            "VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, IN_QTY, VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (%d, 'S', '%s', '', '0001', '0002', 1, '20260618', '%s', '%s')"
            % (ASN_EIN, TEST_LINE, ASN_ADD, ASN_ADD))
        asnId = int(scalar(INV, "SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_ADD='%s'" % ASN_ADD))
        # a detail row carrying the 824 manifest (the manifest lives on the DETAIL, not the header).
        sql(INV,
            "INSERT INTO INV_ASN_DETAIL_MST (IN_ASN_ID, IN_INV_ID, IN_ASN_EIN, VC_MANIFEST_NUMBER, "
            "VC_ASSY_PART_NUMBER, IN_QTY, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (%d, NULL, %d, '%s', '%s', 1, '%s', '%s')"
            % (asnId, ASN_EIN, MANIFEST, PART, ASN_ADD, ASN_ADD))
        sql(INV,
            "INSERT INTO INV_INV_MST (IN_INV_EIN, VC_INV_STATUS, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (%d, 'S', '%s', '%s')" % (INV_EIN, INV_ADD, INV_ADD))
        invId = int(scalar(INV, "SELECT IN_INV_ID FROM INV_INV_MST WHERE VC_ADD='%s'" % INV_ADD))
        rep.check("synthetic 'S' ASN + detail + 'S' invoice injected",
                  scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId) == "S"
                  and scalar(INV, "SELECT VC_INV_STATUS FROM INV_INV_MST WHERE IN_INV_ID=%d" % invId) == "S",
                  "ASN id=%d / INV id=%d both 'S'" % (asnId, invId))

        # ============================================================================================
        # (1)+(2) 997: SH accept -> ASN 'A'; IN reject -> INV 'R'; unknown EIN -> alarm (no silent OK)
        # ============================================================================================
        print("\n--- (1)(2) 997: SH accept / IN reject / unknown-EIN alarm (one file, 3 AK1 groups) ---")
        f997 = file_997(TMM_DUNS, [("SH", ASN_EIN, "A"), ("IN", INV_EIN, "R"), ("SH", UNKNOWN_EIN, "A")])
        p997 = os.path.join(tmpdir, "zzinb_997.edi")
        with open(p997, "w", newline="") as fh:
            fh.write(f997)
        res = edi.process_one_file(f997, "zzinb_997.edi", database=DB_CONN,
                                   srcPath=p997, archiveDir=archiveDir, quarantineDir=quarantineDir)
        rep.check("997 file outcome == PROCESSED", res["outcome"] == "PROCESSED", res["outcome"])
        rep.check("997 detected as type 997", res["type"] == "997", str(res["type"]))

        asnStatus = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        rep.check("SH accept -> ASN flipped 'S' -> 'A' (site-scoped, keyed by EIN+type)",
                  asnStatus == "A", "ASN status=%r" % asnStatus)
        invStatus = scalar(INV, "SELECT VC_INV_STATUS FROM INV_INV_MST WHERE IN_INV_ID=%d" % invId)
        rep.check("IN reject -> invoice flipped 'S' -> 'R'", invStatus == "R", "INV status=%r" % invStatus)

        applied = res["result"]["applied"]
        unknown = res["result"]["unknownEins"]
        rep.check("2 acks applied (the SH + the IN), 1 unknown-EIN", applied == 2 and len(unknown) == 1,
                  "applied=%d unknown=%d" % (applied, len(unknown)))
        alarmRows = scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                "WHERE VC_ALARM_TYPE='997_UNKNOWN_EIN' AND IN_EIN=%d" % UNKNOWN_EIN)
        rep.check("@@ROWCOUNT==0 unknown EIN -> ONE alarm row (NOT a silent success — legacy bug fix)",
                  alarmRows == "1", "alarm rows=%s" % alarmRows)
        rep.check("the unknown-EIN alarm is unresolved (BIT_RESOLVED=0 -> home-hub surfaces it)",
                  scalar(INV, "SELECT BIT_RESOLVED FROM INV_EDI_ALARM_REJ WHERE IN_EIN=%d" % UNKNOWN_EIN)
                  in ("0", "False", "false"),
                  scalar(INV, "SELECT BIT_RESOLVED FROM INV_EDI_ALARM_REJ WHERE IN_EIN=%d" % UNKNOWN_EIN))
        # the file was archived post-commit (the DB ledger is the authoritative idempotency guard).
        rep.check("processed 997 archived to the Archive dir",
                  os.path.exists(os.path.join(archiveDir, "zzinb_997.edi")), "archived")
        rep.check("a ledger row records the 997 as PROCESSED for site %d" % SITE_ID,
                  scalar(INV, "SELECT VC_STATUS FROM INV_EDI_INBOUND_LOG "
                              "WHERE VC_FILE_NAME='zzinb_997.edi' AND VC_EDI_TYPE='997'") == "PROCESSED",
                  "ledger PROCESSED")

        # ============================================================================================
        # (3) 824: manifest match -> ASN 'R' + an alarm row per reject line with the detail
        # ============================================================================================
        print("\n--- (3) 824: manifest match -> ASN 'R' + reject-detail alarm rows (Q10) ---")
        # the ASN is currently 'A' (just accepted by the 997); the 824 is a later application reject ->
        # it must flip the SAME ASN to 'R' (legacy NEVER did this -> the 6 stuck-'S' on Live, §4 trap).
        refs824 = [("RECEIVING QTY SHORT", MANIFEST, PART),
                   ("LABEL UNREADABLE", MANIFEST, "ZZINBPART999")]   # 2 lines, same manifest
        f824 = file_824(TMM_DUNS, refs824)
        p824 = os.path.join(tmpdir, "zzinb_824.edi")
        with open(p824, "w", newline="") as fh:
            fh.write(f824)
        res824 = edi.process_one_file(f824, "zzinb_824.edi", database=DB_CONN,
                                      srcPath=p824, archiveDir=archiveDir, quarantineDir=quarantineDir)
        rep.check("824 file outcome == PROCESSED", res824["outcome"] == "PROCESSED", res824["outcome"])
        asnStatus = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        rep.check("824 manifest match -> the referenced ASN flipped 'A' -> 'R' (D3 recoverable; "
                  "legacy left it stuck)", asnStatus == "R", "ASN status=%r" % asnStatus)
        rep.check("824 flipped exactly 1 ASN here (this manifest is on 1 ASN; the multi-ASN fan-out "
                  "is exercised in (3b))",
                  res824["result"]["asnsFlagged"] == 1, "asnsFlagged=%d" % res824["result"]["asnsFlagged"])
        rejAlarms = sql(INV, "SELECT VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, VC_ERROR_TEXT "
                             "FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='824_ASN_REJECT' "
                             "AND VC_MANIFEST_NUMBER='%s' ORDER BY IN_ALARM_ID" % MANIFEST)
        rep.check("824 wrote ONE alarm row per reject line (2) with manifest/part/errorText "
                  "(click-to-detail payload)", len(rejAlarms) == 2, "%d alarm rows" % len(rejAlarms))
        if len(rejAlarms) == 2:
            rep.check("  alarm[0] captures (manifest, part, errorText) verbatim",
                      rejAlarms[0][0] == MANIFEST and rejAlarms[0][1] == PART
                      and rejAlarms[0][2] == "RECEIVING QTY SHORT", repr(rejAlarms[0]))
            rep.check("  alarm[1] captures the SECOND reject line's part + text",
                      rejAlarms[1][1] == "ZZINBPART999" and rejAlarms[1][2] == "LABEL UNREADABLE",
                      repr(rejAlarms[1]))

        # ============================================================================================
        # (3b) FIX-B — 824 manifest on TWO ASNs -> BOTH flip 'R' (intended over-flag, SHOULD-FIX-2)
        # ============================================================================================
        print("\n--- (3b) FIX-B: a manifest carried on 2 ASNs -> BOTH flagged 'R'; no unrelated ASN touched ---")
        # Build two NEW synthetic 'S' ASNs whose DETAIL rows both carry MANIFEST_FANOUT. Live proves a
        # manifest maps to up to 3 distinct ASNs; the driver's set-based join flips ALL of them.
        fanIds = []
        # NB: a filtered unique index UX_INV_ASN_MST_LINE_PDATE_NORMAL on (VC_LINE_NAME, VC_PRODUCTION_DATE)
        # forbids two ASNs sharing the setup ASN's (line, prod-date) — so give each fan-out ASN its OWN
        # line name + production date. (They share the MANIFEST on their DETAIL rows; the header line/date
        # is irrelevant to the 824 manifest join.)
        for tag in (ASN_FAN1_ADD, ASN2_ADD):   # two FRESH VC_ADD tags (distinct from the setup ASN)
            ein = 880100 + len(fanIds)
            lineNm = "ZZINBFAN%d" % len(fanIds)        # distinct per ASN (avoid the line/date unique idx)
            prodDt = "2026061%d" % (len(fanIds) + 1)   # distinct production date too (belt+suspenders)
            sql(INV,
                "INSERT INTO INV_ASN_MST (IN_ASN_EIN, VC_ASN_STATUS, VC_LINE_NAME, VC_ASSEMBLY_LINE, "
                "VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, IN_QTY, VC_PRODUCTION_DATE, VC_LAST_UPDATE, "
                "VC_ADD) VALUES (%d, 'S', '%s', '', '0001', '0002', 1, '%s', '%s', '%s')"
                % (ein, lineNm, prodDt, tag, tag))
            fid = int(scalar(INV, "SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_ADD='%s'" % tag))
            sql(INV,
                "INSERT INTO INV_ASN_DETAIL_MST (IN_ASN_ID, IN_INV_ID, IN_ASN_EIN, VC_MANIFEST_NUMBER, "
                "VC_ASSY_PART_NUMBER, IN_QTY, VC_LAST_UPDATE, VC_ADD) "
                "VALUES (%d, NULL, %d, '%s', '%s', 1, '%s', '%s')"
                % (fid, ein, MANIFEST_FANOUT, PART, tag, tag))
            fanIds.append(fid)
        rep.check("two synthetic 'S' ASNs share manifest %s (the fan-out fixture)" % MANIFEST_FANOUT,
                  scalar(INV, "SELECT COUNT(DISTINCT a.IN_ASN_ID) FROM INV_ASN_MST a "
                              "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID "
                              "WHERE d.VC_MANIFEST_NUMBER='%s'" % MANIFEST_FANOUT) == "2", "2 ASNs")
        # snapshot an UNRELATED ASN (the original asnId carries MANIFEST, not MANIFEST_FANOUT).
        unrelatedBefore = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        ffan = file_824(TMM_DUNS, [("MULTI-ASN REJECT", MANIFEST_FANOUT, PART)])
        pfan = os.path.join(tmpdir, "zzinb_824fanout.edi")
        with open(pfan, "w", newline="") as fh:
            fh.write(ffan)
        resFan = edi.process_one_file(ffan, "zzinb_824fanout.edi", database=DB_CONN,
                                      srcPath=pfan, archiveDir=archiveDir, quarantineDir=quarantineDir)
        rep.check("multi-ASN 824 outcome == PROCESSED", resFan["outcome"] == "PROCESSED",
                  resFan["outcome"])
        rep.check("FIX-B: a manifest on 2 ASNs flips BOTH to 'R' (asnsFlagged==2, the intended over-flag)",
                  resFan["result"]["asnsFlagged"] == 2,
                  "asnsFlagged=%d" % resFan["result"]["asnsFlagged"])
        bothR = scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE IN_ASN_ID IN (%d,%d) "
                            "AND VC_ASN_STATUS='R'" % (fanIds[0], fanIds[1]))
        rep.check("  ...both fan-out ASNs are now 'R'", bothR == "2", "%s of 2 are 'R'" % bothR)
        unrelatedAfter = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        rep.check("  ...NO unrelated ASN was touched (the %s-only ASN unchanged)" % MANIFEST,
                  unrelatedBefore == unrelatedAfter,
                  "before=%r after=%r" % (unrelatedBefore, unrelatedAfter))
        fanAlarm = scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                               "WHERE VC_ALARM_TYPE='824_ASN_REJECT' AND VC_MANIFEST_NUMBER='%s'"
                               % MANIFEST_FANOUT)
        rep.check("  ...the reject is captured in the alarm row(s) (one per reject LINE = 1 here)",
                  fanAlarm == "1", "%s alarm rows" % fanAlarm)

        # ============================================================================================
        # (4) DUNS GUARD: a wrong-DUNS file is QUARANTINED, not silently dropped (Q11)
        # ============================================================================================
        # MECHANICS ONLY (adversary BLOCKER-1): a value at delSL[4] matching VC_TMM_DUNS routes; a
        # non-matching one quarantines. This does NOT prove which X12 element a real inbound TEMA ISA
        # uses for the routing DUNS — that is golden-pending (see inbound_isa caveat). Only the
        # split + VC_TMM_DUNS match + quarantine plumbing is exercised here.
        print("\n--- (4) DUNS guard MECHANICS: wrong-DUNS file quarantined (element index golden-pending) ---")
        fbad = file_997(BAD_DUNS, [("SH", ASN_EIN, "A")])
        pbad = os.path.join(tmpdir, "zzinb_baddunsfile.edi")
        with open(pbad, "w", newline="") as fh:
            fh.write(fbad)
        # snapshot the ASN status BEFORE -> the bad-DUNS file must NOT touch it.
        before = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        resbad = edi.process_one_file(fbad, "zzinb_baddunsfile.edi", database=DB_CONN,
                                      srcPath=pbad, archiveDir=archiveDir, quarantineDir=quarantineDir)
        rep.check("wrong-DUNS file outcome == QUARANTINED_NO_DUNS",
                  resbad["outcome"] == "QUARANTINED_NO_DUNS", resbad["outcome"])
        rep.check("wrong-DUNS file moved to the Quarantine dir (NOT left to re-scan, NOT dropped)",
                  os.path.exists(os.path.join(quarantineDir, "zzinb_baddunsfile.edi")), "quarantined")
        after = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        rep.check("wrong-DUNS file did NOT touch any ASN status", before == after,
                  "before=%r after=%r" % (before, after))
        rep.check("a ledger row records the quarantine (re-drop of the same junk is also a no-op)",
                  scalar(INV, "SELECT VC_STATUS FROM INV_EDI_INBOUND_LOG "
                              "WHERE VC_FILE_NAME='zzinb_baddunsfile.edi'") == "QUARANTINED",
                  "ledger QUARANTINED")

        # ============================================================================================
        # (5) IDEMPOTENT RE-PROCESS: the SAME 997 dropped twice processes once
        # ============================================================================================
        print("\n--- (5) idempotency: re-process the SAME 997 -> ledger no-op, no double-apply ---")
        # reset the ASN to 'S' so a (wrong) double-apply would be visible as a re-flip to 'A'.
        sql(INV, "UPDATE INV_ASN_MST SET VC_ASN_STATUS='S' WHERE IN_ASN_ID=%d" % asnId)
        alarmsBefore = scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                   "WHERE VC_ALARM_TYPE='997_UNKNOWN_EIN' AND IN_EIN=%d" % UNKNOWN_EIN)
        # re-drop the identical 997 (same name + content -> same hash -> ledger hit).
        with open(p997, "w", newline="") as fh:
            fh.write(f997)
        resRe = edi.process_one_file(f997, "zzinb_997.edi", database=DB_CONN,
                                     srcPath=p997, archiveDir=archiveDir, quarantineDir=quarantineDir)
        rep.check("re-processing the same 997 -> SKIPPED_ALREADY_PROCESSED (ledger no-op)",
                  resRe["outcome"] == "SKIPPED_ALREADY_PROCESSED", resRe["outcome"])
        asnStatus = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        rep.check("the re-drop did NOT re-apply the flip (ASN stayed 'S', not re-flipped to 'A')",
                  asnStatus == "S", "ASN status=%r" % asnStatus)
        alarmsAfter = scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                  "WHERE VC_ALARM_TYPE='997_UNKNOWN_EIN' AND IN_EIN=%d" % UNKNOWN_EIN)
        rep.check("the re-drop did NOT write a duplicate unknown-EIN alarm",
                  alarmsBefore == alarmsAfter, "before=%s after=%s" % (alarmsBefore, alarmsAfter))
        rep.check("still exactly ONE ledger PROCESSED row for the 997 (not two)",
                  scalar(INV, "SELECT COUNT(*) FROM INV_EDI_INBOUND_LOG "
                              "WHERE VC_FILE_NAME='zzinb_997.edi' AND VC_STATUS='PROCESSED'") == "1",
                  "ledger rows")

        # ============================================================================================
        # (5c) FIX-A — SAME CONTENT, DIFFERENT NAME -> processed ONCE (content is the dedup identity)
        # ============================================================================================
        # The bug: the old (name AND hash) key let a renamed redelivery (VAN/mailer re-stamps the file
        # name) re-process -> an 824 would DOUBLE-flip its ASN to 'R' + write DUPLICATE alarms. The fix
        # dedups on the CONTENT HASH alone, so identical content under a NEW name is a NO-OP.
        print("\n--- (5c) FIX-A: same content under a NEW filename -> processed ONCE (no dup alarm/flip) ---")
        # Use the SAME 824 content we already processed in (3) but a DIFFERENT filename. (3) already
        # PROCESSED this exact body under 'zzinb_824.edi', so the same content under a new name MUST skip.
        # First reset the (3) ASN to 'A' so a (wrong) re-process would be visible as a re-flip to 'R'.
        sql(INV, "UPDATE INV_ASN_MST SET VC_ASN_STATUS='A' WHERE IN_ASN_ID=%d" % asnId)
        dupAlarmsBefore = scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                      "WHERE VC_ALARM_TYPE='824_ASN_REJECT' "
                                      "AND VC_MANIFEST_NUMBER='%s'" % MANIFEST)
        pRename = os.path.join(tmpdir, "zzinb_824_RENAMED.edi")   # identical content, new name
        with open(pRename, "w", newline="") as fh:
            fh.write(f824)                                        # f824 == the (3) body verbatim
        resRename = edi.process_one_file(f824, "zzinb_824_RENAMED.edi", database=DB_CONN,
                                         srcPath=pRename, archiveDir=archiveDir,
                                         quarantineDir=quarantineDir)
        rep.check("same 824 CONTENT under a NEW name -> SKIPPED_ALREADY_PROCESSED (content = identity)",
                  resRename["outcome"] == "SKIPPED_ALREADY_PROCESSED", resRename["outcome"])
        asnAfterRename = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % asnId)
        rep.check("the renamed redelivery did NOT re-flip the ASN (stayed 'A', not re-flipped to 'R')",
                  asnAfterRename == "A", "ASN status=%r" % asnAfterRename)
        dupAlarmsAfter = scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                     "WHERE VC_ALARM_TYPE='824_ASN_REJECT' "
                                     "AND VC_MANIFEST_NUMBER='%s'" % MANIFEST)
        rep.check("the renamed redelivery wrote ZERO duplicate alarms (0 new 824_ASN_REJECT rows)",
                  dupAlarmsBefore == dupAlarmsAfter,
                  "before=%s after=%s" % (dupAlarmsBefore, dupAlarmsAfter))
        rep.check("no PROCESSED ledger row for the renamed file (it was a content-hash no-op)",
                  scalar(INV, "SELECT COUNT(*) FROM INV_EDI_INBOUND_LOG "
                              "WHERE VC_FILE_NAME='zzinb_824_RENAMED.edi'") == "0", "no ledger row")
        # corollary: a DIFFERENT-content file reusing an ALREADY-USED name DOES reprocess (new hash =
        # new event). Drop new 824 content (a different manifest) under the (3) name 'zzinb_824.edi'.
        fDiff = file_824(TMM_DUNS, [("DIFFERENT CONTENT SAME NAME", MANIFEST_FANOUT, PART)])
        rep.check("sanity: the different-content file has a DIFFERENT hash than the (3) 824",
                  edi._hash_text(fDiff) != edi._hash_text(f824), "distinct content")
        pDiff = os.path.join(tmpdir, "zzinb_824.edi")            # REUSE the (3) filename
        with open(pDiff, "w", newline="") as fh:
            fh.write(fDiff)
        resDiff = edi.process_one_file(fDiff, "zzinb_824.edi", database=DB_CONN,
                                       srcPath=pDiff, archiveDir=archiveDir, quarantineDir=quarantineDir)
        rep.check("DIFFERENT content under an ALREADY-USED name -> reprocessed (PROCESSED, new hash)",
                  resDiff["outcome"] == "PROCESSED", resDiff["outcome"])

        # (5b) the poll-loop entry process_inbound() over a fresh dir end-to-end (one new 997).
        print("\n--- (5b) process_inbound() poll over a directory ---")
        polldir = tempfile.mkdtemp(prefix="edi_inbound_poll_")
        try:
            # a NEW (unprocessed) 997 acking the ASN we just reset to 'S'.
            fpoll = file_997(TMM_DUNS, [("SH", ASN_EIN, "A")])
            with open(os.path.join(polldir, "zzinb_poll997.edi"), "w", newline="") as fh:
                fh.write(fpoll)
            pollRes = edi.process_inbound(polldir, database=DB_CONN)
            outcomes = [r["outcome"] for r in pollRes]
            rep.check("process_inbound poll processed the one new 997 (PROCESSED)",
                      outcomes == ["PROCESSED"], str(outcomes))
            rep.check("the polled 997 flipped the ASN back to 'A'",
                      scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d"
                                  % asnId) == "A", "ASN status")
        finally:
            sql(INV, "DELETE FROM INV_EDI_INBOUND_LOG WHERE VC_FILE_NAME='zzinb_poll997.edi'")
            shutil.rmtree(polldir, ignore_errors=True)

    finally:
        cleanup()
        shutil.rmtree(tmpdir, ignore_errors=True)

    # confirm cleanup left nothing behind
    leftAsn = scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_ADD='%s' OR VC_LINE_NAME='%s'"
                          % (ASN_ADD, TEST_LINE))
    leftLog = scalar(INV, "SELECT COUNT(*) FROM INV_EDI_INBOUND_LOG WHERE VC_FILE_NAME LIKE 'zzinb_%'")
    rep.check("spike restored as-found (no synthetic ASN / ledger rows remain)",
              leftAsn == "0" and leftLog == "0", "asn=%s log=%s" % (leftAsn, leftLog))

    print("")
    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())

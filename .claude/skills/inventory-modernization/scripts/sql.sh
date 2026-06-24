#!/usr/bin/env bash
# Helper for reading the UTF-16LE SQL files in this repo.
# Usage:
#   sql.sh cat schema              # dump schema as UTF-8
#   sql.sh cat triggers            # dump triggers.sql as UTF-8
#   sql.sh grep schema "PATTERN"   # grep the schema (case-insensitive)
#   sql.sh proc  "PROC_NAME"       # print one stored procedure body
#   sql.sh list  procs|triggers|tables
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../../.." && pwd)"
# Live schema dump from the production server (2026-06-12), authoritative. RELOCATED to db/ in the
# repo reorg (chore/repo-reorg-ignition); a pointer stub remains at "DB Schema/CreateInventory.sql.MOVED.md".
# The older spaced-name "Create Inventory.superseded-2026-06-01.sql" is the prior snapshot pre-2026-spec-cites point at.
SCHEMA="$ROOT/db/CreateInventory.sql"
TRIGGERS="$ROOT/docs/triggers.sql"
conv() { iconv -f UTF-16LE -t UTF-8 "$1"; }
file_for() { case "$1" in schema) echo "$SCHEMA";; triggers) echo "$TRIGGERS";; *) echo "$1";; esac; }

cmd="${1:-}"; shift || true
case "$cmd" in
  cat)  conv "$(file_for "${1:?which file}")";;
  grep) conv "$(file_for "${1:?which file}")" | grep -iE "${2:?pattern}";;
  list)
    case "${1:?procs|triggers|tables}" in
      procs)    conv "$SCHEMA" | grep -iE "CREATE +PROCEDURE" | sed -E 's/.*CREATE +PROCEDURE +//I; s/\[?dbo\]?\.//I; s/[][]//g; s/[ (].*$//' | sort -u;;
      triggers) conv "$SCHEMA" | grep -iE "CREATE +TRIGGER"   | sed -E 's/.*CREATE +TRIGGER +//I; s/\[?dbo\]?\.//I; s/[][]//g; s/[ (].*$//' | sort -u;;
      tables)   conv "$SCHEMA" | grep -iE "CREATE +TABLE"     | sed -E 's/.*CREATE +TABLE +//I; s/\[?dbo\]?\.//I; s/[][]//g; s/[ (].*$//' | sort -u;;
    esac;;
  proc)
    # print from "CREATE PROCEDURE ...<name>" to the next "CREATE PROCEDURE"/"GO"
    conv "$SCHEMA" | awk -v n="${1:?proc name}" '
      BEGIN{IGNORECASE=1; inblk=0}
      /CREATE +PROCEDURE/ { if (index($0,n)>0){inblk=1} else if(inblk){exit} }
      inblk{print}'
    ;;
  *) echo "usage: sql.sh {cat|grep|list|proc} ..."; exit 1;;
esac

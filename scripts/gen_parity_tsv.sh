#!/usr/bin/env bash
# gen_parity_tsv.sh — regenerate the SC1 parity TSV dumps from SIM_OrderSimulation.
#
# parity_diff.py compares the proc's Section B (grid rows) + Section C (phased
# cells) against the golden Delphi exports. Those dumps live in /tmp and clear
# between sessions, so re-run this before parity_diff.py.
#
# Anchor matches the golden exports: Today=2026-06-15, Line=COROLLA, FillDays=25.
# Requires the mssql-spike container up and SA_PASS exported (see spike-db.sh).
set -euo pipefail

CONTAINER=${CONTAINER:-mssql-spike}
SA_PASS=${SA_PASS:?export SA_PASS first (see scripts/spike-db.sh)}
SQLCMD='/opt/mssql-tools18/bin/sqlcmd -C -S localhost -U sa'
TODAY=${TODAY:-2026-06-15}
LINE=${LINE:-COROLLA}
FILLDAYS=${FILLDAYS:-25}
OUT=${OUT:-/tmp}

dump() {  # $1=PartType(upper) $2=lowercase-suffix $3=Section
  docker exec "$CONTAINER" $SQLCMD -P "$SA_PASS" -d Inventory -s $'\t' -W -h -1 \
    -Q "SET NOCOUNT ON; EXEC dbo.SIM_OrderSimulation @LineName='$LINE', @PartType='$1', @Today='$TODAY', @FillDays=$FILLDAYS, @Section='$3'"
}

for pt in TIRE WHEEL VALVE FILM; do
  lc=$(echo "$pt" | tr '[:upper:]' '[:lower:]')
  # Section B: keep only PART rows (parser ignores SIZE_HEADER/FOOTER)
  dump "$pt" "$lc" B | grep '^PART' > "$OUT/secB_${lc}.tsv"
  dump "$pt" "$lc" C                > "$OUT/secC_${lc}.tsv"
  echo "  $pt -> $OUT/secB_${lc}.tsv ($(wc -l < "$OUT/secB_${lc}.tsv") parts), $OUT/secC_${lc}.tsv ($(wc -l < "$OUT/secC_${lc}.tsv") cells)"
done
echo "DONE. Now: python3 docs/analysis/order/spike/parity_diff.py"

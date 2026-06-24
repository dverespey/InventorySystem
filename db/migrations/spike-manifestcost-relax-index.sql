/* ============================================================================
   spike-manifestcost-relax-index.sql — apply D13.2(b)/D13.3 to the spike DB.
   ----------------------------------------------------------------------------
   The live INV_MANIFEST_COST_MST has a UNIQUE constraint/index
   IX_INV_MANIFEST_COST_MST on VC_ASSY_MANIFEST_NUMBER (global manifest-number
   uniqueness). David ruled (D13.2 option b) this is a legacy on-the-fly quirk to
   CORRECT: the real billing constraint is D6 non-overlapping date windows per assy
   code (enforced app-side by the ManifestCost view's checkWindowOverlap), NOT
   manifest-number uniqueness. The global unique blocks a D6-valid insert (same
   manifest number, different assy code OR a 2nd non-overlapping window), so the
   rebuild relaxes it.

   This DROPs that constraint in the spike so option (b) is real + testable. It is
   the spike's instance of the D13.3 cutover conversion step; the production
   migration script must do the same (and add the D6 per-(site,assy) non-overlap
   constraint as the real backstop — see decisions.md D13.3).

   The restored Inventory.bak re-creates the constraint, so RE-RUN this after any
   spike re-seed (spike-db.sh) to keep ManifestCost behaving per (b).

   Run (sa): export SA_PASS=...; docker exec -i mssql-spike /opt/mssql-tools18/bin/sqlcmd \
               -C -S localhost -U sa -P "$SA_PASS" -d Inventory -i /dev/stdin < this file
   ============================================================================ */
USE Inventory;
GO
IF EXISTS (SELECT 1 FROM sys.indexes
           WHERE object_id = OBJECT_ID('dbo.INV_MANIFEST_COST_MST')
             AND name = 'IX_INV_MANIFEST_COST_MST')
BEGIN
    -- It backs a UNIQUE KEY constraint, so drop it as a constraint (not DROP INDEX).
    ALTER TABLE dbo.INV_MANIFEST_COST_MST DROP CONSTRAINT IX_INV_MANIFEST_COST_MST;
    PRINT 'D13.2(b): dropped IX_INV_MANIFEST_COST_MST (global manifest-number unique).';
END
ELSE
    PRINT 'D13.2(b): IX_INV_MANIFEST_COST_MST already absent (nothing to drop).';
GO
SELECT 'manifest-cost unique indexes remaining' AS check_,
       COUNT(*) AS n
FROM sys.indexes
WHERE object_id = OBJECT_ID('dbo.INV_MANIFEST_COST_MST') AND is_unique = 1;
GO

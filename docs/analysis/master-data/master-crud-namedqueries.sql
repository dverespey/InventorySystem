/* ============================================================================
   master-crud-namedqueries.sql  —  Supplier master CRUD query set (SQL source of truth)
   ----------------------------------------------------------------------------
   Status:  build artifact for the Supplier master CRUD (first master-data rebuild)
   Author:  ignition-developer / 2026-06-16
   Mirrors: the Order spike's named-queries.sql  (one findable query per op/table)
   Design:  docs/analysis/master-data/IGNITION-master-crud-design.md  §A.2, §B.1

   ----------------------------------------------------------------------------
   MECHANISM (per §A.2 MECHANISM CORRECTION, 2026-06-16):
     These are NOT on-disk Ignition Named Query resources. A headless build cannot
     author on-disk NQ resources (the gateway parses a NQ's data.bin as Ignition
     XML serialization; a hand-authored file fails with
     "SAXParseException: Content is not allowed in prolog" — proven on this gateway).
     Instead this file is the CANONICAL SQL, executed at runtime via inline
       system.db.runPrepQuery(sql, args, "Inventory_Spike")
     inside the Perspective views' binding/script transforms (the pattern PROVEN by
     the Order spike: …/views/Order/OrderSpike/view.json).

     Ignition NQ param syntax (:paramName) is shown below for the canonical/Designer
     form. The runtime views use positional "?" placeholders with an ordered args
     list (runPrepQuery), in the same column order documented per query. When this
     project is later opened in the Designer these can be promoted to true NQ
     resources for the single-point-edit benefit (a Designer task, not headless).

   ----------------------------------------------------------------------------
   SITE SEAM (D1):  every read/write of a master table carries a :siteId param,
     even though INV_SUPPLIER_MST has NO site_id column today (verified 2026-06-16:
     COL_LENGTH(...,'site_id') -> NULL). The predicate is the commented `-- IG-SITE:`
     line below; today the param is accepted/validated but NOT applied to a WHERE.
     The Postgres phase (R3) flips every `-- IG-SITE:` line on in one pass, AFTER the
     single-column UNIQUE IX_INV_SUPPLIER_MST is rebuilt composite (site_id, code).
     `siteId` is sourced server-side from session.custom.siteId (defaults to 1 in the
     spike), NEVER from a client param.

   GROUND TRUTH (read off live `Inventory` on mssql-spike, 2026-06-16):
     INV_SUPPLIER_MST: PK IN_SUPPLIER_ID; business key VC_SUPPLIER_CODE varchar(5);
       UNIQUE IX_INV_SUPPLIER_MST on VC_SUPPLIER_CODE (single-column); NO site_id.
     INV_LOGISTICS_MST(IN_LOGISTICS_ID, VC_LOGISTICS_NAME) — genuine surrogate FK.
     INV_PART_TYPE_MST(VC_PART_TYPE) — Create-Order-Sheet combo, code-valued.
     DELETE_SupplierCode trigger (verified live body): (1) nullifies
       INV_PARTS_STOCK_MST.IN_SUPPLIER_ID, (2) DELETEs INV_BREAKDOWN_FC_INF where
       VC_SUPPLIER_CODE matches, (3) DELETEs INV_FORECAST_INF where VC_SUPPLIER_CODE
       matches. refCount (below) mirrors ALL THREE — the R1 data-loss gate.
     Audit recipe below is byte-identical to live INSERT_SupplierInfo (verified).

   8.1 ↔ 8.3:
     # IG83-TODO: replace the yyyymmddHHMMSSff audit string with a real datetime
                  DEFAULT/trigger at the Postgres phase; drop the string form.
     # IG83-TODO: flip every `-- IG-SITE:` predicate on + add composite
                  (site_id, VC_SUPPLIER_CODE) UNIQUE index (R3 ordered migration).
     # IG81-COMPAT: plain parameterized T-SQL; runs identically on 8.1.52 and 8.3.
   ============================================================================ */


/* ----------------------------------------------------------------------------
   Supplier/list   (Query)  — grid rows for the List view
   params (ordered for runPrepQuery args):  searchTerm (String), searchTerm (String), siteId (Int)
     NOTE: searchTerm appears twice positionally (code LIKE, name LIKE) plus the
           '' guard; see the view's args list. Designer NQ form uses :searchTerm once.
   returns: RecordID + display columns; enum codes mapped to labels for the grid.
   ---------------------------------------------------------------------------- */
SELECT  s.IN_SUPPLIER_ID        AS "RecordID",
        s.VC_SUPPLIER_CODE      AS "Supplier Code",
        s.VC_SUPPLIER_NAME      AS "Supplier Name",
        s.VC_CITY               AS "City",
        s.VC_STATE              AS "State",
        l.VC_LOGISTICS_NAME     AS "Logistics",
        CASE s.VC_OUTPUT_FILE WHEN 'T' THEN 'TEXT' WHEN 'E' THEN 'EXCEL' WHEN 'B' THEN 'BOTH' END AS "Output File Type",
        CASE s.VC_INVENTORY_ADD_POINT WHEN 'S' THEN 'SHIPPED' WHEN 'A' THEN 'ARRIVED' END AS "Inventory Add Point"
FROM    INV_SUPPLIER_MST s
        LEFT OUTER JOIN INV_LOGISTICS_MST l ON s.IN_LOGISTICS_ID = l.IN_LOGISTICS_ID
WHERE  (:searchTerm = '' OR s.VC_SUPPLIER_CODE LIKE '%' + :searchTerm + '%'
                        OR s.VC_SUPPLIER_NAME LIKE '%' + :searchTerm + '%')
-- IG-SITE:  AND s.site_id = :siteId
ORDER BY s.VC_SUPPLIER_CODE;
/* Parity: searchTerm='' matches SELECT_SupplierInfo '' ordering/contents (enum
   labels identical). Server-side LIKE search improves on the legacy client exact
   filter (documented divergence, §B.6). */


/* ----------------------------------------------------------------------------
   Supplier/get   (Query)  — one row by surrogate id, for the Detail form
   params:  recordId (Int), siteId (Int)
   returns: all editable columns PLUS IN_LOGISTICS_ID (combo value, D2) AND
            VC_LOGISTICS_NAME (combo display). Enum/part-type stored codes returned
            raw; the view maps code->label client-side.
   ---------------------------------------------------------------------------- */
SELECT  s.IN_SUPPLIER_ID,  s.VC_SUPPLIER_CODE, s.VC_SUPPLIER_NAME,
        s.VC_ADDRESS, s.VC_CITY, s.VC_STATE, s.VC_ZIP, s.VC_COUNTRY,
        s.VC_TEL, s.VC_FAX, s.VC_PERSON, s.VC_EMAIL_ADDRESS,
        s.VC_BREAKDOWN_ORDER_DIRECTORY,
        s.IN_LOGISTICS_ID,                      -- combo value (D2: id, not name)
        l.VC_LOGISTICS_NAME,                    -- combo display
        s.VC_OUTPUT_FILE, s.BIT_ORDER_FILE_TIMESTAMP, s.BIT_SITE_NUMBER_IN_ORDER,
        s.VC_CREATE_ORDER_SHEET, s.VC_INVENTORY_ADD_POINT
FROM    INV_SUPPLIER_MST s
        LEFT OUTER JOIN INV_LOGISTICS_MST l ON s.IN_LOGISTICS_ID = l.IN_LOGISTICS_ID
WHERE   s.IN_SUPPLIER_ID = :recordId
-- IG-SITE:  AND s.site_id = :siteId
;
/* vs legacy: SELECT_SupplierInfo returned the logistics NAME only. get returns
   IN_LOGISTICS_ID too, so the FK combo binds by id (D2). */


/* ----------------------------------------------------------------------------
   Supplier/insert   (Update Query, returns identity)
   params (ordered):  code, name, address, city, state, zip, country, tel, fax,
                      person, email, directory, logisticsId, outputFile,
                      orderFileTimestamp, siteNumberInOrder, createOrderSheet,
                      invAddPoint   ( + siteId, IG-SITE only)
   returns: SCOPE_IDENTITY() AS newId (legacy proc never echoed it — improvement).
   audit:   VC_ADD = byte-identical 16-char yyyymmddHHMMSSff recipe (matches the
            live INSERT_SupplierInfo proc, verified 2026-06-16).
   NOTE: includes VC_COUNTRY (legacy proc omitted it; §B.2 surfaces Country —
         harmless documented divergence). Explicit column list (not positional).
   ---------------------------------------------------------------------------- */
INSERT INTO INV_SUPPLIER_MST
    (VC_SUPPLIER_CODE, VC_SUPPLIER_NAME, VC_ADDRESS, VC_CITY, VC_STATE, VC_ZIP, VC_COUNTRY,
     VC_TEL, VC_FAX, VC_PERSON, VC_EMAIL_ADDRESS, VC_BREAKDOWN_ORDER_DIRECTORY,
     IN_LOGISTICS_ID, VC_OUTPUT_FILE, BIT_ORDER_FILE_TIMESTAMP, BIT_SITE_NUMBER_IN_ORDER,
     VC_CREATE_ORDER_SHEET, VC_INVENTORY_ADD_POINT, VC_ADD
     /* IG-SITE: , site_id */)
VALUES
    (:code, :name, :address, :city, :state, :zip, :country,
     :tel, :fax, :person, :email, :directory,
     :logisticsId, :outputFile, :orderFileTimestamp, :siteNumberInOrder,
     :createOrderSheet, :invAddPoint,
     CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
     /* IG-SITE: , :siteId */);
SELECT CAST(SCOPE_IDENTITY() AS int) AS newId;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   Supplier/update   (Update Query)  — keyed on the surrogate id (D2; rename-safe)
   params (ordered):  code, name, address, city, state, zip, country, tel, fax,
                      person, email, directory, logisticsId, outputFile,
                      orderFileTimestamp, siteNumberInOrder, createOrderSheet,
                      invAddPoint, recordId   ( + siteId, IG-SITE only)
   audit:   VC_LASTUPDATE = same 16-char recipe; VC_ADD untouched.
   ---------------------------------------------------------------------------- */
UPDATE INV_SUPPLIER_MST SET
    VC_SUPPLIER_CODE=:code, VC_SUPPLIER_NAME=:name, VC_ADDRESS=:address, VC_CITY=:city,
    VC_STATE=:state, VC_ZIP=:zip, VC_COUNTRY=:country, VC_TEL=:tel, VC_FAX=:fax,
    VC_PERSON=:person, VC_EMAIL_ADDRESS=:email, VC_BREAKDOWN_ORDER_DIRECTORY=:directory,
    IN_LOGISTICS_ID=:logisticsId, VC_OUTPUT_FILE=:outputFile,
    BIT_ORDER_FILE_TIMESTAMP=:orderFileTimestamp, BIT_SITE_NUMBER_IN_ORDER=:siteNumberInOrder,
    VC_CREATE_ORDER_SHEET=:createOrderSheet, VC_INVENTORY_ADD_POINT=:invAddPoint,
    VC_LASTUPDATE = CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
WHERE IN_SUPPLIER_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   Supplier/checkCodeUnique   (Query)  — uniqueness pre-check (A.4)
   params:  code (String), excludeId (Int, default 0), siteId (Int)
   excludeId = 0 on insert; excludeId = recordId on update (lets a row keep its
   own code = rename support, D2). The live IX_INV_SUPPLIER_MST UNIQUE index is the
   race backstop. Today checks code alone (matches the single-column live index);
   the :siteId predicate is the IG-SITE line, flipped on with the composite index.
   ---------------------------------------------------------------------------- */
SELECT COUNT(*) AS n
FROM   INV_SUPPLIER_MST
WHERE  VC_SUPPLIER_CODE = :code
  AND  IN_SUPPLIER_ID  <> :excludeId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   Supplier/refCount   (Query)  — the D3 RESTRICT delete gate (R1-critical)
   params:  recordId (Int), code (String)
   Counts EVERY table the live DELETE_SupplierCode trigger would touch, so the gate
   blocks any delete that would fire the trigger's forecast cascade:
     - INV_PARTS_STOCK_MST  by IN_SUPPLIER_ID  (the trigger NULLIFIES these)
     - INV_BREAKDOWN_FC_INF by VC_SUPPLIER_CODE (the trigger HARD-DELETEs these)
     - INV_FORECAST_INF     by VC_SUPPLIER_CODE (the trigger HARD-DELETEs these)
   Any non-zero total -> block; never call delete. The supplier `code` is sourced
   from the already-loaded form (it is the deleted row's VC_SUPPLIER_CODE).
   R1 FIX: a parts-only gate was blind to the two forecast hard-deletes; this NQ
   mirrors the full trigger body (verified live, 2026-06-16).
   ---------------------------------------------------------------------------- */
SELECT
    (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST  WHERE IN_SUPPLIER_ID   = :recordId)
  + (SELECT COUNT(*) FROM INV_BREAKDOWN_FC_INF WHERE VC_SUPPLIER_CODE = :code)
  + (SELECT COUNT(*) FROM INV_FORECAST_INF     WHERE VC_SUPPLIER_CODE = :code)
    AS n;


/* ----------------------------------------------------------------------------
   Supplier/delete   (Update Query)  — only reached AFTER refCount = 0 (A.5/B.5)
   param:  recordId (Int)
   The live DELETE_SupplierCode trigger then has nothing to act on (inert by
   construction). NEVER let a delete reach the trigger while references exist.
   ---------------------------------------------------------------------------- */
DELETE FROM INV_SUPPLIER_MST WHERE IN_SUPPLIER_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   lookups/logistics   (Query)  — FK combo source for the Logistics dropdown (D2)
   param:  siteId (Int)
   value = IN_LOGISTICS_ID (genuine surrogate FK), label = VC_LOGISTICS_NAME.
   Blank selection in the view -> save IN_LOGISTICS_ID = NULL (legacy "empty
   logistics saves NULL" behavior, §A.6).
   ---------------------------------------------------------------------------- */
SELECT IN_LOGISTICS_ID AS id, VC_LOGISTICS_NAME AS label
FROM INV_LOGISTICS_MST
/* IG-SITE: WHERE site_id = :siteId */
ORDER BY VC_LOGISTICS_NAME;


/* ----------------------------------------------------------------------------
   lookups/partType   (Query)  — Create-Order-Sheet combo (code-valued, NOT a FK)
   param:  siteId (Int)
   VC_CREATE_ORDER_SHEET stores the part-type CODE, so the dropdown value IS the
   code string (not a surrogate id). Only IN_LOGISTICS_ID is a genuine surrogate FK.
   ---------------------------------------------------------------------------- */
SELECT VC_PART_TYPE AS id, VC_PART_TYPE AS label
FROM INV_PART_TYPE_MST
/* IG-SITE: WHERE site_id = :siteId */
ORDER BY VC_PART_TYPE;

/* R6: NO lookups/addPoint NQ. Inventory-Add-Point is the static S/A enum (D4),
   not a lookup combo — INV_ADD_POINT_INF query dropped as dead. The Output-File
   T/E/B enum is likewise a static dropdown (no NQ). Both declared in the view. */


/* ============================================================================
   SIZE MASTER  —  second master-data rebuild module (leaf master)
   ----------------------------------------------------------------------------
   Status:  build artifact for the Size master CRUD (replicates the PROVEN Supplier
            pattern; strict, leaner subset)
   Author:  ignition-developer / 2026-06-17
   Source:  docs/analysis/master-data/size.md  +  IGNITION-master-crud-design.md §C
   View:    Master/Size/Size  (combined master-detail, route /size)

   SAME MECHANISM as Supplier above: NOT on-disk NQ resources — this is the
   canonical SQL, executed at runtime via inline system.db.runPrepQuery inside the
   view's binding/script transforms (positional '?' placeholders, ordered args).

   GROUND TRUTH (read off live `Inventory` on mssql-spike, 2026-06-17):
     INV_SIZE_MST: PK IN_SIZE_ID identity; business key VC_SIZE_CODE varchar(6);
       UNIQUE IX_INV_SIZE_MST on VC_SIZE_CODE (single-column); NO site_id.
       Data columns: VC_SIZE_NAME varchar(50), IN_USAGE int NULL, IN_DAYS int NULL.
       Audit columns: VC_LAST_UPDATE (WITH underscore — differs from Supplier's
       VC_LASTUPDATE), VC_ADD. 64 rows live.
     LEAF MASTER: no FK combos, no enums (drop all lookups/*). Just 4 user fields.
     DELETE_SizeCode trigger (verified live body, 2026-06-17): does ONE thing —
       UPDATE INV_PARTS_STOCK_MST SET IN_SIZE_ID = NULL FROM ...,DELETED d WHERE
       a.IN_SIZE_ID = d.IN_SIZE_ID. It does NOT touch INV_PARTS_STOCK_MST_HIST,
       and that table DOES carry IN_SIZE_ID (COL_LENGTH = 4) — so a legacy delete
       would leave dangling history FKs. Per D3 (RESTRICT) refCount counts BOTH
       parts AND _HIST and BLOCKS the delete (no nulling, no dangling).
     Audit recipe below is byte-identical to live INSERT_SizeInfo / UPDATE_SizeInfo
       (the same 16-char yyyymmddHHMMSSff recipe Supplier uses).

   8.1 ↔ 8.3:
     # IG83-TODO: replace the yyyymmddHHMMSSff audit string with a real datetime
                  DEFAULT/trigger at the Postgres phase; drop the string form.
     # IG83-TODO: flip every `-- IG-SITE:` predicate on + add composite
                  (site_id, VC_SIZE_CODE) UNIQUE index (R3 ordered migration).
     # IG81-COMPAT: plain parameterized T-SQL; runs identically on 8.1.52 and 8.3.

   DIVERGENCES FROM LEGACY (documented, intentional):
     - Validation added: presence(code, name) + len(code) <= 6. Legacy had NONE.
       (NOTE: <=6, NOT the ==5 Supplier rule — VC_SIZE_CODE is varchar(6).)
     - D8 Bug 1 NOT reproduced: legacy InsertSizeInfo dup-checked SELECT_AssyRatioInfo
       (broadcast codes), so real size dups were never caught. checkCodeUnique below
       checks INV_SIZE_MST.VC_SIZE_CODE (the size's OWN code). excludeId supports
       rename (D2).
     - 0-vs-NULL: IN_USAGE/IN_DAYS written as the entered integer, 0 if blank
       (matches legacy, which forces blanks to 0; never NULL from this form).
   ============================================================================ */


/* ----------------------------------------------------------------------------
   Size/list   (Query)  — grid rows for the List view
   params (ordered):  searchTerm (String), searchTerm (String), siteId (Int)
   returns: RecordID + the 4 display columns (no enums, no joins). Mirrors legacy
            SELECT_SizeInfo '' (5 UI-aliased columns), ORDER BY VC_SIZE_CODE.
   ---------------------------------------------------------------------------- */
SELECT  IN_SIZE_ID     AS "RecordID",
        VC_SIZE_CODE   AS "Size Code",
        VC_SIZE_NAME   AS "Size Name",
        IN_USAGE       AS "Daily Usage",
        IN_DAYS        AS "Safety Days"
FROM    INV_SIZE_MST
WHERE  (:searchTerm = '' OR VC_SIZE_CODE LIKE '%' + :searchTerm + '%'
                        OR VC_SIZE_NAME LIKE '%' + :searchTerm + '%')
-- IG-SITE:  AND site_id = :siteId
ORDER BY VC_SIZE_CODE;
/* Parity: searchTerm='' matches SELECT_SizeInfo '' ordering/contents. Server-side
   LIKE search improves on the legacy client-side exact Filter (documented). */


/* ----------------------------------------------------------------------------
   Size/get   (Query)  — one row by surrogate id, for the Detail form
   params:  recordId (Int), siteId (Int)
   ---------------------------------------------------------------------------- */
SELECT  IN_SIZE_ID, VC_SIZE_CODE, VC_SIZE_NAME, IN_USAGE, IN_DAYS
FROM    INV_SIZE_MST
WHERE   IN_SIZE_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   Size/insert   (Update Query, returns identity)
   params (ordered):  code, name, usage, days   ( + siteId, IG-SITE only)
   returns: SCOPE_IDENTITY() AS newId (legacy proc never echoed it — improvement).
   audit:   VC_ADD = byte-identical 16-char yyyymmddHHMMSSff recipe (matches live
            INSERT_SizeInfo).  usage/days arrive as the entered ints (0 if blank).
   ---------------------------------------------------------------------------- */
INSERT INTO INV_SIZE_MST
    (VC_SIZE_CODE, VC_SIZE_NAME, IN_USAGE, IN_DAYS, VC_ADD
     /* IG-SITE: , site_id */)
VALUES
    (:code, :name, :usage, :days,
     CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
     /* IG-SITE: , :siteId */);
SELECT CAST(SCOPE_IDENTITY() AS int) AS newId;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   Size/update   (Update Query)  — keyed on the surrogate id (D2; rename-safe)
   params (ordered):  code, name, usage, days, recordId  ( + siteId, IG-SITE only)
   audit:   ⚠️ VC_LAST_UPDATE (WITH underscore) = same 16-char recipe; VC_ADD untouched.
   ---------------------------------------------------------------------------- */
UPDATE INV_SIZE_MST SET
    VC_SIZE_CODE=:code, VC_SIZE_NAME=:name, IN_USAGE=:usage, IN_DAYS=:days,
    VC_LAST_UPDATE = CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
WHERE IN_SIZE_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   Size/checkCodeUnique   (Query)  — uniqueness pre-check; D8 Bug-1 FIX
   params:  code (String), excludeId (Int, default 0), siteId (Int)
   Checks INV_SIZE_MST.VC_SIZE_CODE — the size's OWN code — NOT SELECT_AssyRatioInfo
   (the legacy bug). excludeId = 0 on insert; = recordId on update (rename support,
   D2). The live IX_INV_SIZE_MST UNIQUE index is the race backstop.
   ---------------------------------------------------------------------------- */
SELECT COUNT(*) AS n
FROM   INV_SIZE_MST
WHERE  VC_SIZE_CODE = :code
  AND  IN_SIZE_ID  <> :excludeId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   Size/refCount   (Query)  — the D3 RESTRICT delete gate (R1-critical)
   params:  recordId (Int)
   The live DELETE_SizeCode trigger only nullifies INV_PARTS_STOCK_MST.IN_SIZE_ID
   (verified body, 2026-06-17) — it does NOT touch INV_PARTS_STOCK_MST_HIST, which
   ALSO carries IN_SIZE_ID. Per D3 (RESTRICT, no unlink, no dangling), this gate
   counts BOTH current parts AND history rows; any non-zero total -> BLOCK delete.
   (Size's code-keyed cross-module writer UPDATE_SizeUsage is a writer, not a
   reference; it is out of scope for the delete gate.)
   ---------------------------------------------------------------------------- */
SELECT
    (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST      WHERE IN_SIZE_ID = :recordId)
  + (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE IN_SIZE_ID = :recordId)
    AS n;


/* ----------------------------------------------------------------------------
   Size/delete   (Update Query)  — only reached AFTER refCount = 0
   param:  recordId (Int)
   The live DELETE_SizeCode trigger then has nothing to act on (inert by
   construction). NEVER let a delete reach the trigger while references exist.
   ---------------------------------------------------------------------------- */
DELETE FROM INV_SIZE_MST WHERE IN_SIZE_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;

/* Size is a LEAF master: NO lookups/* queries (no FK combos, no enums). */

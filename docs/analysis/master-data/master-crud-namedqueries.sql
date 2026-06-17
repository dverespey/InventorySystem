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
     Instead, each Perspective view INLINES its own copy of this SQL and runs it via
       system.db.runPrepQuery(sql, args, "Inventory_Spike")
     in its binding/script transforms (the pattern PROVEN by the Order spike:
     …/views/Order/OrderSpike/view.json).

     ⚠️ AUTHORITATIVE-AT-RUNTIME = THE VIEW's inline SQL, NOT this file. This file is
     NOT loaded or executed — it is the HUMAN-READABLE AUTHORING SPEC that each view's
     inline SQL is hand-copied from, and the two MUST be kept in sync by hand (the view
     uses positional "?" placeholders + an ordered args list; this spec shows the
     equivalent :paramName form for readability/the future Designer promotion). They CAN
     drift — when editing a query, change BOTH the view and this spec. (Single-point-of-edit
     via real on-disk Named Queries is blocked headless — see the MECHANISM note above; it
     becomes possible only when the project is opened in the Designer and these are promoted
     to true NQ resources, a Designer task. Until then: view = truth, this = synced spec.)

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


/* ============================================================================
   LOGISTICS MASTER  —  third master-data rebuild module (leaf, NAME-keyed)
   ----------------------------------------------------------------------------
   Status:  build artifact for the Logistics master CRUD (replicates the PROVEN Size
            template; leaf master keyed by a NAME, not a 5-char code).
   Author:  ignition-developer / 2026-06-17
   Source:  docs/analysis/master-data/logistics.md  +  IGNITION-master-crud-design.md §C
   View:    Master/Logistics/Logistics  (combined master-detail, route /logistics)

   SAME MECHANISM as Size/Supplier above: NOT on-disk NQ resources — this is the
   canonical SQL, executed at runtime via inline system.db.runPrepQuery inside the
   view's binding/script transforms (positional '?' placeholders, ordered args).

   GROUND TRUTH (read off live `Inventory` on mssql-spike, 2026-06-17):
     INV_LOGISTICS_MST: PK IN_LOGISTICS_ID identity; BUSINESS KEY VC_LOGISTICS_NAME
       varchar(25) — a NAME, NOT a code; UNIQUE IX_INV_LOGISTICS_MST on
       VC_LOGISTICS_NAME (single-column); NO site_id.
       Data columns: VC_ADDRESS(50), VC_CITY(50), VC_STATE(50), VC_ZIP(10),
       VC_COUNTRY(50), VC_TEL(10), VC_FAX(10), VC_PERSON(50), VC_EMAIL_ADDRESS(255),
       VC_BREAKDOWN_ORDER_DIRECTORY(512).
       ⚠️ Audit columns: VC_LASTUPDATE (NO underscore — SAME as Supplier, DIFFERS
       from Size's VC_LAST_UPDATE) + VC_ADD. Both varchar(16). 1 row live.
     LEAF MASTER: no FK combos, no enums (drop all lookups/*).
     DELETE_LogisticsCode trigger (verified live body, 2026-06-17):
       UPDATE INV_SUPPLIER_MST SET IN_LOGISTICS_ID=null FROM INV_SUPPLIER_MST a,
       DELETED d WHERE a.IN_LOGISTICS_ID = d.IN_LOGISTICS_ID.
       It nullifies ONLY supplier FKs and LEAVES INV_PARTS_STOCK_MST.IN_LOGISTICS_ID
       (and _HIST.IN_LOGISTICS_ID — COL_LENGTH=4, present) DANGLING. Per D3 (RESTRICT,
       no unlink, no dangling), refCount below counts ALL THREE (supplier + parts +
       _HIST) and BLOCKS the delete on any non-zero total — this kills the legacy
       dangle (the cleanest D3 win).
     Audit recipe below is byte-identical to the live INSERT/UPDATE_LogisticsInfo
       procs (the same 16-char yyyymmddHHMMSSff recipe Supplier/Size use).

   8.1 ↔ 8.3:
     # IG83-TODO: replace the yyyymmddHHMMSSff audit string with a real datetime
                  DEFAULT/trigger at the Postgres phase; drop the string form.
     # IG83-TODO: flip every `-- IG-SITE:` predicate on + add composite
                  (site_id, VC_LOGISTICS_NAME) UNIQUE index (R3 ordered migration).
     # IG81-COMPAT: plain parameterized T-SQL; runs identically on 8.1.52 and 8.3.

   DIVERGENCES FROM LEGACY (documented, intentional):
     - Validation added: presence(name) + per-site uniqueness on VC_LOGISTICS_NAME.
       Legacy form had NO presence check (a blank name was insertable). NO fixed-
       length rule (there is no code). checkCodeUnique checks the NAME column;
       excludeId supports rename (D2).
     - WIDEN proc-truncated fields: legacy procs narrowed @LogPerson to varchar(25)
       and @LogDirectory to varchar(215). Direct NQ writes the FULL DB widths
       (VC_PERSON 50, VC_BREAKDOWN_ORDER_DIRECTORY 512).
     - VC_COUNTRY surfaced as a field (legacy form hid it) — same posture as Supplier.
     - Delete is D3 RESTRICT (block) instead of the legacy supplier-nullify +
       parts-dangle behavior.
   ============================================================================ */


/* ----------------------------------------------------------------------------
   Logistics/list   (Query)  — grid rows for the List view
   params (ordered):  searchTerm (String), searchTerm (String), siteId (Int)
   returns: RecordID + display columns (no enums, no joins). Mirrors legacy
            SELECT_LogisticsInfo '' ordering (ORDER BY VC_LOGISTICS_NAME).
   ---------------------------------------------------------------------------- */
SELECT  IN_LOGISTICS_ID    AS "RecordID",
        VC_LOGISTICS_NAME  AS "Logistics Name",
        VC_CITY            AS "City",
        VC_STATE           AS "State",
        VC_TEL             AS "Telephone",
        VC_PERSON          AS "Person"
FROM    INV_LOGISTICS_MST
WHERE  (:searchTerm = '' OR VC_LOGISTICS_NAME LIKE '%' + :searchTerm + '%'
                        OR VC_PERSON LIKE '%' + :searchTerm + '%')
-- IG-SITE:  AND site_id = :siteId
ORDER BY VC_LOGISTICS_NAME;
/* Parity: searchTerm='' matches SELECT_LogisticsInfo '' ordering/contents. Server-
   side LIKE search improves on the legacy client-side exact Filter (documented). */


/* ----------------------------------------------------------------------------
   Logistics/get   (Query)  — one row by surrogate id, for the Detail form
   params:  recordId (Int), siteId (Int)
   returns: all editable columns (name + address block + contact + directory +
            country). Surrogate key IN_LOGISTICS_ID resolves the row (D2).
   ---------------------------------------------------------------------------- */
SELECT  IN_LOGISTICS_ID, VC_LOGISTICS_NAME, VC_ADDRESS, VC_CITY, VC_STATE, VC_ZIP,
        VC_COUNTRY, VC_TEL, VC_FAX, VC_PERSON, VC_EMAIL_ADDRESS,
        VC_BREAKDOWN_ORDER_DIRECTORY
FROM    INV_LOGISTICS_MST
WHERE   IN_LOGISTICS_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   Logistics/insert   (Update Query, returns identity)
   params (ordered):  name, address, city, state, zip, country, tel, fax, person,
                      email, directory   ( + siteId, IG-SITE only)
   returns: SCOPE_IDENTITY() AS newId (legacy proc never echoed it — improvement).
   audit:   VC_ADD = byte-identical 16-char yyyymmddHHMMSSff recipe (matches live
            INSERT_LogisticsInfo). WIDENS person->50, directory->512 (full DB width).
   ---------------------------------------------------------------------------- */
INSERT INTO INV_LOGISTICS_MST
    (VC_LOGISTICS_NAME, VC_ADDRESS, VC_CITY, VC_STATE, VC_ZIP, VC_COUNTRY,
     VC_TEL, VC_FAX, VC_PERSON, VC_EMAIL_ADDRESS, VC_BREAKDOWN_ORDER_DIRECTORY, VC_ADD
     /* IG-SITE: , site_id */)
VALUES
    (:name, :address, :city, :state, :zip, :country,
     :tel, :fax, :person, :email, :directory,
     CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
     /* IG-SITE: , :siteId */);
SELECT CAST(SCOPE_IDENTITY() AS int) AS newId;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   Logistics/update   (Update Query)  — keyed on the surrogate id (D2; rename-safe)
   params (ordered):  name, address, city, state, zip, country, tel, fax, person,
                      email, directory, recordId   ( + siteId, IG-SITE only)
   audit:   ⚠️ VC_LASTUPDATE (NO underscore) = same 16-char recipe; VC_ADD untouched.
   ---------------------------------------------------------------------------- */
UPDATE INV_LOGISTICS_MST SET
    VC_LOGISTICS_NAME=:name, VC_ADDRESS=:address, VC_CITY=:city, VC_STATE=:state,
    VC_ZIP=:zip, VC_COUNTRY=:country, VC_TEL=:tel, VC_FAX=:fax, VC_PERSON=:person,
    VC_EMAIL_ADDRESS=:email, VC_BREAKDOWN_ORDER_DIRECTORY=:directory,
    VC_LASTUPDATE = CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
WHERE IN_LOGISTICS_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   Logistics/checkCodeUnique   (Query)  — uniqueness pre-check on the NAME
   params:  name (String), excludeId (Int, default 0), siteId (Int)
   Checks INV_LOGISTICS_MST.VC_LOGISTICS_NAME (the business key — a NAME, not a
   code). excludeId = 0 on insert; = recordId on update (rename support, D2). The
   live IX_INV_LOGISTICS_MST UNIQUE index is the race backstop. NO fixed-length rule.
   ---------------------------------------------------------------------------- */
SELECT COUNT(*) AS n
FROM   INV_LOGISTICS_MST
WHERE  VC_LOGISTICS_NAME = :name
  AND  IN_LOGISTICS_ID  <> :excludeId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   Logistics/refCount   (Query)  — the D3 RESTRICT delete gate (R1-critical)
   params:  recordId (Int)
   The live DELETE_LogisticsCode trigger only nullifies INV_SUPPLIER_MST.IN_LOGISTICS_ID
   (verified body, 2026-06-17) and LEAVES INV_PARTS_STOCK_MST + _HIST DANGLING. Per D3
   (RESTRICT) this gate counts ALL THREE inbound references and BLOCKs on any non-zero
   total — never reaching the trigger while references exist. This kills the legacy
   dangle (parts/_HIST were never cleaned up by the trigger).
   ---------------------------------------------------------------------------- */
SELECT
    (SELECT COUNT(*) FROM INV_SUPPLIER_MST          WHERE IN_LOGISTICS_ID = :recordId)
  + (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST        WHERE IN_LOGISTICS_ID = :recordId)
  + (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST   WHERE IN_LOGISTICS_ID = :recordId)
    AS n;


/* ----------------------------------------------------------------------------
   Logistics/delete   (Update Query)  — only reached AFTER refCount = 0
   param:  recordId (Int)
   The live DELETE_LogisticsCode trigger then has nothing to act on (inert by
   construction). NEVER let a delete reach the trigger while references exist.
   ---------------------------------------------------------------------------- */
DELETE FROM INV_LOGISTICS_MST WHERE IN_LOGISTICS_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;

/* Logistics is a LEAF master: NO lookups/* queries (no FK combos, no enums). */


/* ============================================================================
   RenbanGroup master CRUD  —  INV_RENBAN_GROUP_MST  (fourth master)
   ----------------------------------------------------------------------------
   Author:  ignition-developer / 2026-06-17
   Spec:    docs/analysis/admin/master-remainders.md §A
   Mirrors: the Size canonical template (audit cols VC_LAST_UPDATE + VC_ADD,
            same recipe; single combined view; D3 RESTRICT delete gate).

   GROUND TRUTH (read off live `Inventory` on mssql-spike, verified 2026-06-17):
     - 12 columns. Key IN_RENBAN_ID (int identity). Business key
       VC_RENBAN_GROUP_CODE varchar(5) (e.g. "CMWA"); UNIQUE index
       IX_INV_RENBAN_GROUP_MST is on VC_RENBAN_GROUP_CODE (confirmed).
     - Audit cols: VC_LAST_UPDATE (WITH underscore) + VC_ADD — same as Size,
       NOT Logistics' no-underscore VC_LASTUPDATE.
     - Editable fields: VC_RENBAN_GROUP_CODE (text), VC_RENBAN_GROUP_COUNT
       varchar(3) (3-char zero-padded counter "001".."999"), IN_SHIP_DAYS int,
       and 6 per-weekday IN_SHIP_DAYS_MONDAY..SATURDAY int. (No Sunday column.)
     - 5 rows today: CAP, CMWA, DICAS, HCAP, PACF.

   COUNTER (VC_RENBAN_GROUP_COUNT):  the legacy form left-pads the typed value to
     exactly 3 chars before save (1->"001", 12->"012"). Reproduced in the view's
     Save script (numeric 0..999, then "%03d"). The AUTOMATIC counter increment
     (UPDATE_RenbanGroupCount, the read-then-write race D11#7) is OUT OF SCOPE for
     this CRUD form — it belongs to the Order/RenbanOrder pipeline. This form does
     ONLY the manual zero-padded edit.

   ⚠️ SPEC CORRECTION (R1):  master-remainders.md §A says "Triggers: none" — WRONG.
     The live DB HAS a FOR-DELETE trigger DELETE_RenbanGroupCode. Verified body
     (2026-06-17):
        CREATE TRIGGER [dbo].[DELETE_RenbanGroupCode] ON [dbo].[INV_RENBAN_GROUP_MST]
        FOR DELETE AS
          UPDATE INV_PARTS_STOCK_MST SET IN_RENBAN_ID = null
          FROM INV_PARTS_STOCK_MST p, DELETED d
          WHERE p.IN_RENBAN_ID = d.IN_RENBAN_ID
     It nullifies the CURRENT parts' IN_RENBAN_ID and IGNORES
     INV_PARTS_STOCK_MST_HIST (which ALSO carries IN_RENBAN_ID, COL_LENGTH=4) — the
     same dangling-_HIST pattern as Size/Logistics. Per D3 (RESTRICT) the refCount
     gate counts BOTH parts + _HIST and BLOCKs on any non-zero total, so the trigger
     is never reached while references exist (kills the dangle).
   ============================================================================ */


/* ----------------------------------------------------------------------------
   RenbanGroup/list   (Query)  — grid rows for the List view
   params (ordered):  searchTerm (String), searchTerm (String), siteId (Int)
   returns: RecordID + display columns. ORDER BY VC_RENBAN_GROUP_CODE (parity).
   ---------------------------------------------------------------------------- */
SELECT  IN_RENBAN_ID           AS "RecordID",
        VC_RENBAN_GROUP_CODE   AS "RenbanGroupCode",
        VC_RENBAN_GROUP_COUNT  AS "Count",
        IN_SHIP_DAYS           AS "ShipDays"
FROM    INV_RENBAN_GROUP_MST
WHERE  (:searchTerm = '' OR VC_RENBAN_GROUP_CODE LIKE '%' + :searchTerm + '%')
-- IG-SITE:  AND site_id = :siteId
ORDER BY VC_RENBAN_GROUP_CODE;
/* Parity: searchTerm='' matches the legacy list ordering. Single business-key
   column so the LIKE search is on the code only. */


/* ----------------------------------------------------------------------------
   RenbanGroup/get   (Query)  — one row by surrogate id, for the Detail form
   params:  recordId (Int), siteId (Int)
   returns: code + counter + IN_SHIP_DAYS + the 6 per-weekday ship-days.
   Surrogate key IN_RENBAN_ID resolves the row (D2).
   ---------------------------------------------------------------------------- */
SELECT  IN_RENBAN_ID, VC_RENBAN_GROUP_CODE, VC_RENBAN_GROUP_COUNT, IN_SHIP_DAYS,
        IN_SHIP_DAYS_MONDAY, IN_SHIP_DAYS_TUESDAY, IN_SHIP_DAYS_WEDNESDAY,
        IN_SHIP_DAYS_THURSDAY, IN_SHIP_DAYS_FRIDAY, IN_SHIP_DAYS_SATURDAY
FROM    INV_RENBAN_GROUP_MST
WHERE   IN_RENBAN_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   RenbanGroup/insert   (Update Query, returns identity)
   params (ordered):  code, count(3-char zero-padded), shipDays, mon, tue, wed,
                      thu, fri, sat   ( + siteId, IG-SITE only)
   returns: SCOPE_IDENTITY() AS newId.
   audit:   ⚠️ VC_ADD = byte-identical 16-char yyyymmddHHMMSSff recipe (same as Size).
   counter: VC_RENBAN_GROUP_COUNT is passed already left-padded to 3 chars by the
            view's Save script ("%03d" of the numeric value).
   ---------------------------------------------------------------------------- */
INSERT INTO INV_RENBAN_GROUP_MST
    (VC_RENBAN_GROUP_CODE, VC_RENBAN_GROUP_COUNT, IN_SHIP_DAYS,
     IN_SHIP_DAYS_MONDAY, IN_SHIP_DAYS_TUESDAY, IN_SHIP_DAYS_WEDNESDAY,
     IN_SHIP_DAYS_THURSDAY, IN_SHIP_DAYS_FRIDAY, IN_SHIP_DAYS_SATURDAY, VC_ADD
     /* IG-SITE: , site_id */)
VALUES
    (:code, :count, :shipDays, :mon, :tue, :wed, :thu, :fri, :sat,
     CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
     /* IG-SITE: , :siteId */);
SELECT CAST(SCOPE_IDENTITY() AS int) AS newId;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   RenbanGroup/update   (Update Query)  — keyed on the surrogate id (D2; rename-safe)
   params (ordered):  code, count, shipDays, mon, tue, wed, thu, fri, sat,
                      recordId   ( + siteId, IG-SITE only)
   audit:   ⚠️ VC_LAST_UPDATE (WITH underscore) = same 16-char recipe; VC_ADD untouched.
   ---------------------------------------------------------------------------- */
UPDATE INV_RENBAN_GROUP_MST SET
    VC_RENBAN_GROUP_CODE=:code, VC_RENBAN_GROUP_COUNT=:count, IN_SHIP_DAYS=:shipDays,
    IN_SHIP_DAYS_MONDAY=:mon, IN_SHIP_DAYS_TUESDAY=:tue, IN_SHIP_DAYS_WEDNESDAY=:wed,
    IN_SHIP_DAYS_THURSDAY=:thu, IN_SHIP_DAYS_FRIDAY=:fri, IN_SHIP_DAYS_SATURDAY=:sat,
    VC_LAST_UPDATE = CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
WHERE IN_RENBAN_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;
-- IG83-TODO: replace the yyyymmddHHMMSSff string with a real datetime default at the Postgres phase.


/* ----------------------------------------------------------------------------
   RenbanGroup/checkCodeUnique   (Query)  — uniqueness pre-check on the CODE
   params:  code (String), excludeId (Int, default 0), siteId (Int)
   Checks INV_RENBAN_GROUP_MST.VC_RENBAN_GROUP_CODE (the business key).
   excludeId = 0 on insert; = recordId on update (rename support, D2). The live
   IX_INV_RENBAN_GROUP_MST UNIQUE index is the race backstop. Code is varchar(5)
   -> cap at 5 (no fixed-length==N rule; spec states none).
   ---------------------------------------------------------------------------- */
SELECT COUNT(*) AS n
FROM   INV_RENBAN_GROUP_MST
WHERE  VC_RENBAN_GROUP_CODE = :code
  AND  IN_RENBAN_ID        <> :excludeId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   RenbanGroup/refCount   (Query)  — the D3 RESTRICT delete gate (R1-critical)
   params:  recordId (Int)
   The live DELETE_RenbanGroupCode trigger (verified body, 2026-06-17) only
   nullifies INV_PARTS_STOCK_MST.IN_RENBAN_ID and IGNORES
   INV_PARTS_STOCK_MST_HIST (which ALSO carries IN_RENBAN_ID, COL_LENGTH=4). Per
   D3 (RESTRICT) this gate counts BOTH inbound references and BLOCKs on any
   non-zero total — never reaching the trigger while references exist. This kills
   the legacy _HIST dangle.
   ---------------------------------------------------------------------------- */
SELECT
    (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST       WHERE IN_RENBAN_ID = :recordId)
  + (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST  WHERE IN_RENBAN_ID = :recordId)
    AS n;


/* ----------------------------------------------------------------------------
   RenbanGroup/delete   (Update Query)  — only reached AFTER refCount = 0
   param:  recordId (Int)
   The live DELETE_RenbanGroupCode trigger then has nothing to act on (inert by
   construction). NEVER let a delete reach the trigger while references exist.
   ---------------------------------------------------------------------------- */
DELETE FROM INV_RENBAN_GROUP_MST WHERE IN_RENBAN_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;

/* RenbanGroup is a LEAF master: NO lookups/* queries (no FK combos, no enums). */


/* ============================================================================
   PartsStock master  —  INV_PARTS_STOCK_MST  (the FIFTH / keystone master)
   ----------------------------------------------------------------------------
   Author:  ignition-developer / 2026-06-17
   Design:  IGNITION-master-crud-design.md §A (template) + the keystone deltas below.
   Spec:    docs/analysis/inventory-stock/parts-stock-master.md (read in full).

   Key:  IN_PART_ID  int IDENTITY PK.
   Business key:  VC_PART_NUMBER varchar(12), UNIQUE via IX_INV_PARTS_STOCK_MST.
   Audit:  VC_LAST_UPDATE (UNDERSCORE, like Size) on update; VC_ADD on insert.
           ⚠️ VC_ADD is NOT NULL — the only audit col in the schema that is —
           so it is ALWAYS written on insert (the 16-char yyyymmddHHMMSSff recipe).

   FIVE FK combos (D2: store the surrogate id, display the label):
     Supplier     IN_SUPPLIER_ID    label VC_SUPPLIER_CODE
     Logistics    IN_LOGISTICS_ID   label VC_LOGISTICS_NAME   (blank selection -> NULL)
     Size         IN_SIZE_ID        label VC_SIZE_CODE
     RenbanGroup  IN_RENBAN_ID      label VC_RENBAN_GROUP_CODE
     PartType     IN_PART_TYPE_ID   label VC_PART_TYPE
   Legacy resolved label->id INSIDE the proc; the rebuild passes the id directly (D2).

   LINE dropdown — cross-DB, NOW FIXTURED. Source:
     SELECT DISTINCT LineName FROM VehicleOrder.dbo.LINE ORDER BY LineName
   (spike VehicleOrder.dbo.LINE seeded CAMRY/COROLLA/HIGHLANDER/TACOMA/TUNDRA;
    the ignition_spike login reads it cross-DB via the 3-part name over the
    Inventory_Spike connection — verified 2026-06-17). Stored back as the STRING
    VC_LINE_NAME (NO FK; default 'TUNDRA' if blank).
    Production reads LINE on the ALC_Connection->VehicleOrder catalog; this
    3-part-name read is the faithful spike equivalent.

   ⚠️ IN_QTY IS READ-ONLY — the form does NOT own the on-hand invariant.
     12 qty-triggers (the future stock-ledger) own it. So:
       - display IN_QTY as a READ-ONLY field;
       - on INSERT set IN_QTY = 0 (a new part starts at zero on-hand; the ledger moves it);
       - on UPDATE the SET list OMITS IN_QTY ENTIRELY — never let the CRUD form overwrite
         the live balance. This is an INTENTIONAL IMPROVEMENT over the legacy
         `SET IN_QTY=@QTY`-with-loaded-value round-trip: it removes the stale-overwrite
         race (legacy fed the proc the value loaded into a ReadOnly box, so the written
         value merely equaled the loaded one; the rebuild simply does not touch the column).

   ⚠️ BIT_LOT_SIZE_ORDERS IS STORED INVERTED (order-renban domain memory: 0 = lot-sized TRUE).
     Legacy form stores `not checkbox.checked` and displays `not stored`. Reproduced
     faithfully: checkbox<->stored value is INVERTED on BOTH read and write.
       read:  checkboxSelected = NOT storedBit
       write: storedBit        = NOT checkboxSelected   (0 if checked, 1 if unchecked)

   IN_RENBAN_COUNT is int in the table; legacy passed varchar(3). Written as an int
     (numeric-entry); NO zero-pad here (that pad belongs to the RenbanGroup master's
     own count column, a different field).

   SITE SEAM (D1):  INV_PARTS_STOCK_MST DOES have a site_id column in the spike
     (spike-db.sh added it + seeded 15 site-2 rows for a parts-level isolation test).
     Even so, the `-- IG-SITE:` predicate stays COMMENTED for parallel-run safety and
     to match the template: the legacy Delphi app shares this table, and the composite
     unique index migration (R3) is the single point that turns every predicate on. The
     :siteId param is plumbed view->query but not yet applied to a WHERE.
   ============================================================================ */


/* ----------------------------------------------------------------------------
   PartsStock/list   (Query)  — grid rows for the combined view's grid
   params:  searchTerm (String, default ''), siteId (Int)
   LEFT JOIN all five masters so a NULL/dangling FK still returns the part with a
   blank label (LEFT-JOIN parity with the legacy SELECT_PartsStockInfo). Server-side
   LIKE search on part number + name replaces the legacy in-memory Filter (P7).
   View logs:  SPIKE PartsStock/list: N rows
   ---------------------------------------------------------------------------- */
SELECT  p.IN_PART_ID            AS RecordID,
        p.VC_PART_NUMBER        AS PartNumber,
        p.VC_PARTS_NAME         AS PartsName,
        sup.VC_SUPPLIER_CODE    AS SupplierCode,
        lg.VC_LOGISTICS_NAME    AS LogisticsName,
        sz.VC_SIZE_CODE         AS SizeCode,
        rb.VC_RENBAN_GROUP_CODE AS RenbanCode,
        pt.VC_PART_TYPE         AS PartType,
        p.VC_LINE_NAME          AS LineName,
        p.IN_QTY                AS Qty
FROM    INV_PARTS_STOCK_MST p
        LEFT OUTER JOIN INV_SUPPLIER_MST     sup ON p.IN_SUPPLIER_ID  = sup.IN_SUPPLIER_ID
        LEFT OUTER JOIN INV_LOGISTICS_MST    lg  ON p.IN_LOGISTICS_ID = lg.IN_LOGISTICS_ID
        LEFT OUTER JOIN INV_SIZE_MST         sz  ON p.IN_SIZE_ID      = sz.IN_SIZE_ID
        LEFT OUTER JOIN INV_RENBAN_GROUP_MST rb  ON p.IN_RENBAN_ID    = rb.IN_RENBAN_ID
        LEFT OUTER JOIN INV_PART_TYPE_MST    pt  ON p.IN_PART_TYPE_ID = pt.IN_PART_TYPE_ID
WHERE  (:searchTerm = '' OR p.VC_PART_NUMBER LIKE '%' + :searchTerm + '%'
                        OR p.VC_PARTS_NAME   LIKE '%' + :searchTerm + '%')
-- IG-SITE:  AND p.site_id = :siteId
ORDER BY p.VC_PART_NUMBER
;


/* ----------------------------------------------------------------------------
   PartsStock/get   (Query)  — one row by surrogate id (D2). Returns the FK IDS
   (to bind the combos by id) plus the labels (for display), the inverted lot-size
   bit AS STORED (the view inverts on load), and IN_QTY (read-only display).
   params:  recordId (Int), siteId (Int)
   ---------------------------------------------------------------------------- */
SELECT  p.IN_PART_ID, p.VC_PART_NUMBER, p.VC_PARTS_NAME,
        p.IN_SUPPLIER_ID,  sup.VC_SUPPLIER_CODE,
        p.IN_LOGISTICS_ID, lg.VC_LOGISTICS_NAME,
        p.IN_SIZE_ID,      sz.VC_SIZE_CODE,
        p.IN_RENBAN_ID,    rb.VC_RENBAN_GROUP_CODE,
        p.IN_PART_TYPE_ID, pt.VC_PART_TYPE,
        p.VC_LINE_NAME, p.VC_KANBAN_NUMBER, p.IN_1LOTQTY,
        p.BIT_LOT_SIZE_ORDERS,           -- stored INVERTED; view computes selected = NOT stored
        p.IN_RENBAN_COUNT, p.IN_QTY,     -- IN_QTY read-only
        p.IN_LEADTIME, p.IN_LEADTIME_MONDAY, p.IN_LEADTIME_TUESDAY, p.IN_LEADTIME_WEDNESDAY,
        p.IN_LEADTIME_THURSDAY, p.IN_LEADTIME_FRIDAY, p.IN_LEADTIME_SATURDAY,
        p.IN_SHIP_DAYS, p.IN_SHIP_DAYS_MONDAY, p.IN_SHIP_DAYS_TUESDAY, p.IN_SHIP_DAYS_WEDNESDAY,
        p.IN_SHIP_DAYS_THURSDAY, p.IN_SHIP_DAYS_FRIDAY, p.IN_SHIP_DAYS_SATURDAY,
        p.MO_PART_COST, p.VC_COMMENTS
FROM    INV_PARTS_STOCK_MST p
        LEFT OUTER JOIN INV_SUPPLIER_MST     sup ON p.IN_SUPPLIER_ID  = sup.IN_SUPPLIER_ID
        LEFT OUTER JOIN INV_LOGISTICS_MST    lg  ON p.IN_LOGISTICS_ID = lg.IN_LOGISTICS_ID
        LEFT OUTER JOIN INV_SIZE_MST         sz  ON p.IN_SIZE_ID      = sz.IN_SIZE_ID
        LEFT OUTER JOIN INV_RENBAN_GROUP_MST rb  ON p.IN_RENBAN_ID    = rb.IN_RENBAN_ID
        LEFT OUTER JOIN INV_PART_TYPE_MST    pt  ON p.IN_PART_TYPE_ID = pt.IN_PART_TYPE_ID
WHERE   p.IN_PART_ID = :recordId
-- IG-SITE:  AND p.site_id = :siteId
;


/* ----------------------------------------------------------------------------
   PartsStock/insert   (Update Query, returns identity)
   params (data): code, name, supplierId, logisticsId, sizeId, renbanId, partTypeId,
                  lineName, kanban, lotQty, lotSizeBit(=NOT checkbox), renbanCount,
                  leadtime + 6 weekday, shipdays + 6 weekday, partCost, comments  + siteId
   ⚠️ IN_QTY is set to a LITERAL 0 (new part starts at zero on-hand; the ledger moves it).
   ⚠️ VC_ADD (NOT NULL) written with the 16-char yyyymmddHHMMSSff recipe.
   ⚠️ BIT_LOT_SIZE_ORDERS receives the ALREADY-INVERTED value (NOT checkbox.selected).
   Returns SCOPE_IDENTITY() (legacy never echoed the new id — an improvement).
   Explicit column list (not positional) — kills the P10 schema-order fragility.
   ---------------------------------------------------------------------------- */
INSERT INTO INV_PARTS_STOCK_MST
    (VC_PART_NUMBER, VC_PARTS_NAME, IN_SUPPLIER_ID, IN_LOGISTICS_ID, IN_SIZE_ID,
     IN_RENBAN_ID, IN_PART_TYPE_ID, VC_LINE_NAME, VC_KANBAN_NUMBER, IN_1LOTQTY,
     BIT_LOT_SIZE_ORDERS, IN_RENBAN_COUNT, IN_QTY,
     IN_LEADTIME, IN_LEADTIME_MONDAY, IN_LEADTIME_TUESDAY, IN_LEADTIME_WEDNESDAY,
     IN_LEADTIME_THURSDAY, IN_LEADTIME_FRIDAY, IN_LEADTIME_SATURDAY,
     IN_SHIP_DAYS, IN_SHIP_DAYS_MONDAY, IN_SHIP_DAYS_TUESDAY, IN_SHIP_DAYS_WEDNESDAY,
     IN_SHIP_DAYS_THURSDAY, IN_SHIP_DAYS_FRIDAY, IN_SHIP_DAYS_SATURDAY,
     MO_PART_COST, VC_COMMENTS, VC_ADD
     /* IG-SITE: , site_id */)
VALUES
    (:code, :name, :supplierId, :logisticsId, :sizeId,
     :renbanId, :partTypeId, :lineName, :kanban, :lotQty,
     :lotSizeBit, :renbanCount, 0,                            -- IN_QTY literal 0 (read-only)
     :leadtime, :ltMon, :ltTue, :ltWed, :ltThu, :ltFri, :ltSat,
     :shipdays, :sdMon, :sdTue, :sdWed, :sdThu, :sdFri, :sdSat,
     :partCost, :comments,
     CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
     /* IG-SITE: , :siteId */);
SELECT CAST(SCOPE_IDENTITY() AS int) AS newId;
-- IG83-TODO: at the Postgres phase replace the string audit with a real datetime DEFAULT; drop the string form.


/* ----------------------------------------------------------------------------
   PartsStock/update   (Update Query)  — keys on the surrogate id (D2); rename-safe.
   params:  (same data params as insert) + recordId + siteId
   ⚠️ IN_QTY is DELIBERATELY OMITTED from the SET list — the form never overwrites the
      trigger-maintained on-hand balance (intentional improvement over legacy SET IN_QTY=@QTY).
   ⚠️ VC_LAST_UPDATE (UNDERSCORE) set with the 16-char recipe. VC_ADD untouched.
   ⚠️ BIT_LOT_SIZE_ORDERS receives the ALREADY-INVERTED value.
   ---------------------------------------------------------------------------- */
UPDATE INV_PARTS_STOCK_MST SET
    VC_PART_NUMBER=:code, VC_PARTS_NAME=:name, IN_SUPPLIER_ID=:supplierId,
    IN_LOGISTICS_ID=:logisticsId, IN_SIZE_ID=:sizeId, IN_RENBAN_ID=:renbanId,
    IN_PART_TYPE_ID=:partTypeId, VC_LINE_NAME=:lineName, VC_KANBAN_NUMBER=:kanban,
    IN_1LOTQTY=:lotQty, BIT_LOT_SIZE_ORDERS=:lotSizeBit, IN_RENBAN_COUNT=:renbanCount,
    /* IN_QTY intentionally NOT in the SET list (read-only invariant; see header) */
    IN_LEADTIME=:leadtime, IN_LEADTIME_MONDAY=:ltMon, IN_LEADTIME_TUESDAY=:ltTue,
    IN_LEADTIME_WEDNESDAY=:ltWed, IN_LEADTIME_THURSDAY=:ltThu, IN_LEADTIME_FRIDAY=:ltFri,
    IN_LEADTIME_SATURDAY=:ltSat, IN_SHIP_DAYS=:shipdays, IN_SHIP_DAYS_MONDAY=:sdMon,
    IN_SHIP_DAYS_TUESDAY=:sdTue, IN_SHIP_DAYS_WEDNESDAY=:sdWed, IN_SHIP_DAYS_THURSDAY=:sdThu,
    IN_SHIP_DAYS_FRIDAY=:sdFri, IN_SHIP_DAYS_SATURDAY=:sdSat, MO_PART_COST=:partCost,
    VC_COMMENTS=:comments,
    VC_LAST_UPDATE = CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
WHERE IN_PART_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   PartsStock/checkCodeUnique   (Query)  — per-site uniqueness pre-check on the part
   number. params:  code (String), excludeId (Int, default 0), siteId (Int)
   excludeId=0 for insert; excludeId=recordId for update (lets a row keep its own
   number -> rename support, D2). The live UNIQUE index IX_INV_PARTS_STOCK_MST is the
   race backstop (the view also catches the unique-violation SQLException).
   ---------------------------------------------------------------------------- */
SELECT COUNT(*) AS n
FROM   INV_PARTS_STOCK_MST
WHERE  VC_PART_NUMBER = :code
  AND  IN_PART_ID    <> :excludeId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   PartsStock/refCount   (Query)  — the D3 RESTRICT delete gate (R1-critical).
   params:  recordId (Int), partNumber (String)

   LIVE DELETE_PartNumber trigger (verified body, 2026-06-17): logs to
   Activity.dbo.InsertAct_Log 'INVENTORY','TRIGGER', then BLANKS (SET ... = '') the
   four INV_ASSY_RATIO_MST.VC_*_PART_NUMBER*_CODE columns and the two
   INV_FORECAST_DETAIL_INF.VC_*_PART_NUMBER_CODE columns whose value matches the
   deleted VC_PART_NUMBER. It does NOT delete/adjust the HIST table, the qty ledger,
   IN_QTY, or any transactional child (open orders / rejects / stocktaking /
   part-shipping) — so deleting a part with history/transactional children ORPHANS
   those rows. Per D3 (RESTRICT) the rebuild BLOCKS the delete while ANY of the
   transactional/child references exist (verified each table+column exists, 2026-06-17):
     INV_OPEN_ORDER_INF      by VC_PART_NUMBER   (string key)
     INV_REJECT_INF          by IN_PART_ID       (id key)
     INV_STOCKTAKING_INF     by IN_PART_ID       (id key)
     INV_PART_SHIPPING_INF   by VC_PART_NUMBER   (string key)
     INV_PART_QTY_INF        by VC_PART_NUMBER   (qty-ledger, string key)
     INV_BREAKDOWN_FC_INF    by VC_PART_NUMBER   (string key)
   Block on any non-zero total; NEVER reach the trigger while references exist.

   HIST-SCOPE — RESOLVED (David 2026-06-17, option b): _HIST is DROPPED from the gate.
   INV_PARTS_STOCK_MST has an INSERT_PartsStockMST trigger that snapshots EVERY new part
   into INV_PARTS_STOCK_MST_HIST on insert, so counting _HIST would make NO part ever
   hard-deletable (every part has its own snapshot). A part's OWN append-only audit
   history is NOT an external reference, so it does not block — matching the legacy
   DELETE_PartNumber trigger's own scope. The gate blocks on the 6 TRANSACTIONAL-child
   tables only; a part with no orders/rejects/stocktaking/shipping/qty-ledger/forecast
   rows (e.g. a just-created mistake) is deletable. Any part with real inventory/order
   history still blocks. (NOTE: Size/Logistics/RenbanGroup KEEP _HIST in their gates —
   there it is the dangle-kill for parts' FK in _HIST, a different relationship.)
   ---------------------------------------------------------------------------- */
SELECT
    (SELECT COUNT(*) FROM INV_OPEN_ORDER_INF       WHERE VC_PART_NUMBER = :partNumber)
  + (SELECT COUNT(*) FROM INV_REJECT_INF           WHERE IN_PART_ID     = :recordId)
  + (SELECT COUNT(*) FROM INV_STOCKTAKING_INF      WHERE IN_PART_ID     = :recordId)
  + (SELECT COUNT(*) FROM INV_PART_SHIPPING_INF    WHERE VC_PART_NUMBER = :partNumber)
  + (SELECT COUNT(*) FROM INV_PART_QTY_INF         WHERE VC_PART_NUMBER = :partNumber)
  + (SELECT COUNT(*) FROM INV_BREAKDOWN_FC_INF     WHERE VC_PART_NUMBER = :partNumber)
    AS n
;


/* ----------------------------------------------------------------------------
   PartsStock/delete   (Update Query)  — only reached AFTER refCount = 0.
   param:  recordId (Int)
   The live DELETE_PartNumber trigger then has no matching assy-ratio/forecast-detail
   string codes to blank (inert by construction, since a zero-ref part is not
   referenced anywhere). NEVER let a delete reach the trigger while references exist.
   ---------------------------------------------------------------------------- */
DELETE FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
;


/* ----------------------------------------------------------------------------
   PartsStock FK-combo lookups (D2: value = surrogate id, label = code/name).
   Each is a plain Query; :siteId plumbed (IG-SITE seam commented).
   The LINE lookup is the cross-DB string source (NO FK).
   ---------------------------------------------------------------------------- */
-- lookups/supplier
SELECT IN_SUPPLIER_ID AS id, VC_SUPPLIER_CODE AS label
FROM INV_SUPPLIER_MST  /* IG-SITE: WHERE site_id = :siteId */  ORDER BY VC_SUPPLIER_CODE;

-- lookups/logistics   (blank selection -> NULL on save)
SELECT IN_LOGISTICS_ID AS id, VC_LOGISTICS_NAME AS label
FROM INV_LOGISTICS_MST  /* IG-SITE: WHERE site_id = :siteId */  ORDER BY VC_LOGISTICS_NAME;

-- lookups/size
SELECT IN_SIZE_ID AS id, VC_SIZE_CODE AS label
FROM INV_SIZE_MST  WHERE VC_SIZE_CODE <> ''  /* IG-SITE: AND site_id = :siteId */  ORDER BY VC_SIZE_CODE;

-- lookups/renban
SELECT IN_RENBAN_ID AS id, VC_RENBAN_GROUP_CODE AS label
FROM INV_RENBAN_GROUP_MST  /* IG-SITE: WHERE site_id = :siteId */  ORDER BY VC_RENBAN_GROUP_CODE;

-- lookups/partType
SELECT IN_PART_TYPE_ID AS id, VC_PART_TYPE AS label
FROM INV_PART_TYPE_MST  /* IG-SITE: WHERE site_id = :siteId */  ORDER BY VC_PART_TYPE;

-- lookups/line  (cross-DB, string-valued; value == label; default 'TUNDRA' if blank)
--   Production reads LINE on the ALC_Connection->VehicleOrder catalog; this 3-part-name
--   read over the Inventory_Spike connection is the faithful spike equivalent.
SELECT DISTINCT LineName AS id, LineName AS label
FROM VehicleOrder.dbo.LINE  ORDER BY LineName;

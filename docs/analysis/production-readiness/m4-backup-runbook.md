# M4 hardening — Backup / restore runbook (single-site)

M4 piece-3 hardening item **HD7** (`m4-auth-sites-sourcetruth.md` §6 / D-M4-5). The rebuild has TWO
independent stateful tiers that must BOTH be backed up to recover the system: the **SQL Server database**
(all business data + the procs/triggers that ARE the logic) and the **Ignition gateway** (the project —
views, Named Queries, scripts — plus config and the DB connection definitions). A restore needs both.

Single-site scope (each deployment = one gateway = one site). Concrete, actionable; placeholder values
only (no real credentials, paths, or hosts — those live gateway-side / in ops config, never in this repo).

---

## 1. What to back up (the two tiers)

| Tier | Artifact | Holds | Tool |
|---|---|---|---|
| **A. SQL Server DB** | a `.bak` full/diff/log set | ALL business data + the 182 procs + 25 triggers (the logic) + `INV_SITES` config rows | `BACKUP DATABASE` (SQL Agent or `sqlcmd`) |
| **B. Ignition gateway** | a `.gwbk` gateway backup | the project (views, Named Queries, scripts, project-library), gateway config, **the named DB-connection definitions** (incl. the encrypted credential), tag config, user-source/roles | gateway `Backup` (web UI) or `gwcmd -b` |

The two are independent and BOTH required: a `.gwbk` with a stale/empty DB, or a `.bak` with no gateway,
each restores only half the system. Take them on compatible schedules (below) so a paired restore is
point-in-time-close.

---

## 2. Schedule + retention (single-site)

**A. SQL Server database (`.bak`)**
- **Full backup nightly** (off-hours; the EDI/ASN loop is daytime). `BACKUP DATABASE [Inventory] TO
  DISK = N'<backup-dir>\Inventory_full_YYYYMMDD.bak' WITH INIT, CHECKSUM, COMPRESSION;`
- **Transaction-log backups every 15–30 min during business hours** (DB in FULL recovery model) so the
  recovery-point objective for the revenue-critical EDI loop is ≤30 min: `BACKUP LOG [Inventory] TO DISK
  = N'<backup-dir>\Inventory_log_YYYYMMDD_HHMM.trn' WITH CHECKSUM;` (skip if SIMPLE recovery is accepted
  — then RPO = last nightly full; confirm the recovery-model choice with the DBA/David at cutover).
- **Verify each backup**: `RESTORE VERIFYONLY FROM DISK = N'...';` in the same job; alert on failure.
- **Retention**: keep **14 nightly fulls** on local/fast disk; **the DATAPURGE retention floor is 12
  months** (`INV_SITES.IN_DATA_RETENTION ≥ 12`, the legacy `DataModule.pas:6890` rule) — so monthly fulls
  are kept **≥ 12 months** to cover the same horizon the purge prunes, plus year-end fulls per the
  client's records-retention policy.

**B. Ignition gateway (`.gwbk`)**
- **Gateway backup nightly** (after the project is stable; config changes are infrequent): web UI
  Config → Backup/Restore → Download, or headless `gwcmd -b <path>/gateway_YYYYMMDD.gwbk`.
- **Plus an on-demand `.gwbk` BEFORE every deploy/config change** (a project edit, a Named-Query change,
  a DB-connection edit, an upgrade) — this is the cheap rollback.
- **Retention**: keep **14 nightly** + the **last 5 pre-change** `.gwbk` files. They are small.

**Offsite / 3-2-1**: replicate BOTH the `.bak` set and the `.gwbk` set to a **second location** (offsite
or cloud object storage) at least daily — 3 copies, 2 media, 1 offsite. Encrypt at rest offsite (the
`.bak` carries real client data; the `.gwbk` carries the **encrypted DB credential** — treat both as
sensitive, never to the repo, never to a public bucket).

---

## 3. Restore drill (paired, single-site)

Run as a **scheduled drill** (quarterly minimum) on a NON-prod box — a restore you have never tested is
not a backup.

1. **Stand up SQL Server**, restore the DB from the latest full + the log chain to the target point in
   time:
   ```
   RESTORE DATABASE [Inventory] FROM DISK = N'...Inventory_full_YYYYMMDD.bak'
       WITH MOVE 'Inventory'     TO N'<data>\Inventory.mdf',
            MOVE 'Inventory_log' TO N'<log>\Inventory_log.ldf',
            NORECOVERY, REPLACE;
   -- apply each log in order, last one WITH RECOVERY:
   RESTORE LOG [Inventory] FROM DISK = N'...Inventory_log_....trn' WITH NORECOVERY;  -- ... repeat
   RESTORE LOG [Inventory] FROM DISK = N'...Inventory_log_LAST.trn' WITH RECOVERY;
   ```
   (The spike's `scripts/spike-db.sh` does the same restore for the dev sandbox — reuse it as the
   drill's reference; it reads the logical file names via `RESTORE FILELISTONLY` and MOVEs them.)
2. **Restore the gateway** from the latest `.gwbk`: gateway web UI Config → Backup/Restore → Restore, or
   `gwcmd -r <path>/gateway_YYYYMMDD.gwbk`. This brings back the project + config + **the named DB
   connections**.
3. **Re-point / re-enter the DB credential if needed.** The `.gwbk` carries the connection definition;
   on a fresh host the gateway may need the DB connection re-validated (host/port) and the encrypted
   credential re-supplied from the secret store (see `m4-hardening-secrets.md` §2 — the credential is
   gateway-side, not in the repo, so a fresh restore re-enters it).
4. **Smoke-verify** the restored stack:
   - gateway loads clean — `grep "Unable to deserialize" logs/wrapper.log` returns nothing; no FAULTED
     resources.
   - the DB connection is **Valid** (Config → Databases → Connections shows green).
   - a read works (open the landing hub — KPIs render via the Named Queries) and a guarded write works
     (a ProductionControl session round-trips a throwaway master row, then delete it).
   - `INV_SITES` has the expected site row(s); `IN_EIN_SEQ` is at/after the last issued EDI control
     number (a restored-too-far-back EIN would re-mint a used control number — verify before resuming the
     EDI loop).
5. **Record the drill**: which backups, the point-in-time reached, time-to-recover (RTO), and any gap.

---

## 4. Redundancy posture (Q14 / D-M4-5) — note, don't build

Per Q14, **redundancy is invisible to the application** — the rebuild does not implement HA, does not
fail over in code, and has no app-level awareness of a standby. The posture is documented, not built:

- **Today**: single gateway + single SQL Server instance (matches the legacy single-machine posture).
  Backups (above) are the recovery mechanism; RTO is "restore both tiers on a replacement host."
- **If HA is later required** (a David decision, D-M4-5): it is an **infrastructure** concern handled
  BELOW the app and transparently to it —
  - Ignition **redundant gateway** pair (a standby gateway syncing config from the master); the app code
    is unchanged.
  - SQL Server **Always On availability group** / failover-cluster instance behind the gateway's named
    DB connection; the connection name stays the same, so **no driver, view, or Named Query changes**.
  - This is why every driver references the connection by **name** (`m4-hardening-secrets.md`): swapping
    to an HA listener is a gateway-config edit, never a code change.
- **Decision still open (D-M4-5)**: single instance acceptable for go-live, or HA required? Until decided,
  the recovery guarantee is the tested backup/restore drill above (RPO ≤ 30 min with log backups; RTO =
  restore time on a replacement host).

## 5. Quick checklist (ops)
- [ ] DB FULL recovery model confirmed (or SIMPLE accepted, RPO understood)
- [ ] nightly `.bak` full + 15–30 min log backups, with `VERIFYONLY`, alerting on failure
- [ ] nightly `.gwbk` + an on-demand `.gwbk` before every deploy/config change
- [ ] both sets replicated offsite, encrypted at rest
- [ ] quarterly paired restore drill, RTO/RPO recorded
- [ ] post-restore: gateway loads clean, DB connection Valid, `IN_EIN_SEQ` not rewound below last issued

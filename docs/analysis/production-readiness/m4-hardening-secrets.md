# M4 hardening — Secrets / connection strings → the Ignition gateway

M4 piece-3 hardening item **HD1** (`m4-auth-sites-sourcetruth.md` §6) + legacy hazard **H2** (§1).
This is a **config + confirm** deliverable, not a build: it documents the migration of the legacy
plaintext SQL connection strings out of an INI and into the Ignition gateway's encrypted config, and
records the grep-confirmation that the rebuild's drivers already use the named gateway connection (no
hardcoded credentials anywhere in tracked files).

Confidence: legacy facts read from `DataModule.pas` / `DataModule.dfm`; rebuild facts grepped from the
project-library drivers (commands + output reproduced below).

---

## 1. The legacy posture — secrets in cleartext, on every machine

The legacy Delphi app holds **three SQL Server connection strings, each WITH the password in cleartext**,
in a per-machine INI under `[DATABASE]`, read raw at startup:

- `DataModule.pas:731-733` loads `fiInventoryConnection` / `fiActivityConnection` / `fiALCConnection`
  straight from the INI (`TCIniField` properties) into the three `TADOConnection`s
  (`Inv_Connection`, `Act_Connection`, `ALC_Connection`).
- `DataModule.dfm:10` / `DataModule.dfm:522` carry the design-time `ConnectionString` with
  `Password=...` embedded (the .dfm-baked default before the INI override).

This is **legacy hazard H2** (`m4-auth-sites-sourcetruth.md` §1): the connection string IS the real
security perimeter. Anyone on the LAN with a copy of the INI (or the .dfm default) bypasses the app's
Logon form entirely and reads/writes the DB directly. The app login is cosmetic relative to the DB
perimeter, and **the password sits in plaintext in a file on every install.**

Guardrail already in force in this repo (`CLAUDE.md` "Guardrails"): INI files with connection strings are
**git-ignored**; credentials are never committed or echoed into tracked files. This doc preserves that.

---

## 2. The rebuild posture — a NAMED gateway DB connection, creds in the gateway secret store

The Ignition rebuild does not read a connection string from a file at all. It uses an **Ignition gateway
named Database Connection** (`Inventory_Spike` on the spike → `Inventory` SQL Server DB; the prod
deployment names the prod connection the same and points it at the prod server). The gateway stores the
connection's username/password **encrypted in the gateway config** (`gateway config DB → idb` / the
gateway's `config.idb` + the gateway secret store), NOT in any project file, NOT in the repo.

How a driver references it: every project-library driver names the connection by its **logical name** in
a module-level constant and passes that name to `system.db.*`; the gateway resolves the name to the
stored (encrypted) credentials at call time. Example shape (from the drivers):

```
DATABASE = "Inventory_Spike"            # the Inventory rebuild connection name
ALC_DATABASE = "VehicleOrder"           # the shared cross-DB connection (AD_* reads), also a named conn
...
system.db.runPrepQuery(sql, args, DATABASE)
```

The driver code never sees a JDBC URL, a host, a username, or a password — only the **name**. This is the
HD1 migration:

| Legacy ([DATABASE] INI, plaintext) | Rebuild (gateway named connection, encrypted) |
|---|---|
| `fiInventoryConnection` (`DataModule.pas:731`) | gateway DB connection **`Inventory_Spike`** (prod: same name → prod server) |
| `fiActivityConnection` (`DataModule.pas:732`) | folds into the same gateway connection (activity log writes go through the named conn) |
| `fiALCConnection` (`DataModule.pas:733`) | gateway DB connection **`VehicleOrder`** (the shared ALC/cross-DB reads: `AD_GetSite`, `AD_FRSPULL`, the production calendar) |
| `DataModule.dfm` `Password=...` baked default | **gone** — no password in any file; the gateway holds it encrypted |

**Where the secret actually lives:** Ignition gateway config (Config → Databases → Connections), with the
password field encrypted at rest by the gateway. For a tighter prod posture the password can be sourced
from the gateway **secret store** / an external secrets provider rather than typed into the connection —
either way it is gateway-side and **never in a committed file or a plaintext INI**. (Setting up the named
connection + its credential is a gateway-config step, documented at cutover; it is not code.)

---

## 3. Confirmation — NO hardcoded credentials / NO INI conn strings in the rebuild drivers (grep)

Run from the repo root. Empty output = clean.

```
# (A1) hardcoded JDBC URLs / passwords / Data-Source strings in any driver
grep -rniE "jdbc:|password=|pwd=|Data Source=|Initial Catalog=|user id=|uid=" \
  docs/analysis/*/project-library docs/analysis/*/*/project-library
        -> (no matches)

# (A2) the connection is referenced ONLY by logical name
grep -rhoE "(DB|database)\s*=\s*\"[A-Za-z_]+\"" \
  docs/analysis/*/project-library docs/analysis/*/*/project-library | sort | uniq -c
        -> database="Inventory_Spike"   (the named gateway connection — no URL/creds)

# (A3) any INI read / legacy connection-string field in a driver
grep -rniE "\.ini|\[DATABASE\]|fiInventoryConnection|fiActivityConnection|fiALCConnection|ConnectionString" \
  docs/analysis/*/project-library docs/analysis/*/*/project-library
        -> (no matches)
```

Every driver that touches the DB (`asn`, `order`, `stockLedger`, `order_file`, `hotcall`, `renban`,
`edi810`, `edi_inbound`, `edi856`, `forecast`) names the connection `Inventory_Spike` (and the
cross-DB readers also name `VehicleOrder`) via a module-level constant — **CONFIRMED: zero hardcoded
connection strings, zero embedded passwords, zero INI reads in the rebuild's code.** The shim
(`scripts/e2e/jython_shim.py`) maps that same logical name to the spike container for headless tests; it
also holds no credentials (the SA password comes from `$SA_PASS` in the environment, never a tracked
file).

---

## 4. The standing guardrail (unchanged, restated)

- **No credentials in tracked files.** INI files with connection strings are git-ignored (`CLAUDE.md`);
  the legacy `.bak` files (real client data) are git-ignored; `$SA_PASS` is an env var, never committed.
- **Secrets live gateway-side, encrypted.** The DB connection credential is in the gateway config /
  secret store. The repo (drivers + views + Named Queries) references only the connection's **name**.
- **A schema/credential rotation is a gateway-config edit, not a code change.** Rotating the DB password
  or repointing prod is done in the gateway DB-connection screen; no driver, view, or Named Query changes.

## 5. Residual / handoff
- Creating the named gateway DB connection + entering its (encrypted) credential is a **gateway-config
  step performed at cutover** — out of scope for code; noted in the cutover runbook.
- `m4-auth-sites-sourcetruth.md` §1 H1 (plaintext *user* passwords) is a separate item closed by the
  auth piece (HD2: Internal user source hashes, no plaintext, force first-login reset) — not this doc.

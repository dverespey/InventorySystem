# Ignition Version Strategy — Dev 8.1.52 / Target 8.3

*Scoped 2026-06-12; version verified on the dev box 2026-06-13.* Standing constraint for all Ignition
work in this rebuild. Companion to [`ignition-feasibility.md`](ignition-feasibility.md) and
[`ignition-spike-plan.md`](ignition-spike-plan.md).

## The split

| | Version | Why |
|---|---|---|
| **Local dev / test machine** | **Ignition 8.1.52** (verified `gwcmd -i`) | This is an older **Intel MacBook**; **8.3 will not run on it**. 8.1 is the LTS line and 8.1.52 is its mature tail. All local building and testing happens here. |
| **Production target** | **Ignition 8.3** | The deployed system targets 8.3. The dev box is the temporary limiter and will eventually be replaced. |

> **Note on the constraint as stated:** the working instruction was "dev caps at 8.0." The actual
> installed **and running** gateway is **8.1.52** (`/usr/local/ignition`, gateway up on `:8088`). This
> doc is written to **8.1.52** because that is what's on the box. If a literal-8.0 downgrade is ever
> intended, revisit the `system.util.getVersion()` and Perspective-maturity notes below — they change.

**Consequence:** code is written for **8.3 semantics** but must **load and run on 8.1.52** during
development. The 8.1→8.3 gap is narrow (same Perspective lineage, same Jython 2.7, same DB/named-query
APIs); the few real deltas get a **guard** and an **inline retrofit note**.

## Annotation convention (greppable)

Use these exact markers so the retrofit surface is always discoverable:

- `# IG83-TODO: <what to switch to on 8.3>` — a deliberate 8.1-safe path that **should be upgraded** on
  8.3 (e.g. an Event-Stream-shaped cleanup of a polling timer script).
- `# IG81-COMPAT: <why this shape>` — code written to run on the **8.1.52 dev box**; safe on 8.3 but
  worth revisiting.
- `# IG83-ONLY: <guarded>` — a branch that only executes on 8.3 (runtime-guarded); the dev box takes
  the fallback.

## Version detection

`system.util.getVersion()` **exists on 8.1+**, so it works on the dev box and on 8.3:

```python
v = system.util.getVersion()          # e.g. "8.1.52" / "8.3.x"
major, minor = v.getMajor(), v.getMinor()

def atLeast(maj, min):
    """Guard 8.3-only API paths."""
    return (major, minor) >= (maj, min)
```

```python
if atLeast(8, 3):
    # IG83-ONLY: native 8.3 path (e.g. Event Stream)
    ...
else:
    # IG81-COMPAT: timer/event-script fallback the dev box uses
    ...
```

> Version-agnostic fallback (if ever forced onto a build where `getVersion()` is absent, i.e. literal
> 8.0): `from com.inductiveautomation.ignition.common import BundleUtil; BundleUtil.get().getVersion()`.
> Not needed on 8.1.52, kept here as the escape hatch.

## Known deltas to watch (8.1.52 → 8.3)

| Area | 8.1.52 (dev) | 8.3 (target) | Guard / note |
|---|---|---|---|
| `system.util.getVersion()` | present | present | use it directly |
| Jython | 2.7 | 2.7 | **no delta** — scripting stable (good for EDI X12 + proc calls) |
| `system.db.createSProcCall`, Named Queries | present | present | **no delta** — wrap-the-proc parallel-run path is version-safe |
| Perspective components | mature LTS set | superset | prefer components that exist on 8.1; `# IG83-TODO:` any 8.3-only component you stub |
| Event Streams | **absent** | new in 8.3 (Kafka/HTTP/DB/Gateway sources) | build EDI polling as a **gateway timer/event script** (works on both); `# IG83-TODO:` if a stream is the cleaner eventual form |
| Gateway config / API | file/DB-backed | fully API-driven, JSON config (versionable) | don't build tooling that assumes 8.3's config API |
| Offline Perspective form submit | n/a | new in 8.3 (forms queue offline, submit on reconnect) | don't rely on it; ensure save paths are explicit |

*(Each "8.1.52" cell verified on the actual dev gateway during the spike where it touches the slice.)*

## Impact on the PartsStockMaster spike (Check A)

Check A's UI-velocity measurement runs on **8.1.52 Perspective — the mature LTS line**, so the reading
is **broadly representative** of 8.3, not heavily discounted (this is the key change from the earlier
"8.0 weakest-release" framing — that discount was based on the wrong version).

- A small discount still applies only where an 8.3-only component would have saved time; tag those
  spots, don't let them dominate the estimate.
- The "no scaffold generator" cost is **inherent to Perspective on every version** — that's the real
  signal Check A measures, and 8.1.52 measures it faithfully.
- Checks **B** (`siteScopedQuery()`) and **C** (atomic file I/O via gateway timer scripts) are
  **version-neutral** — they prove out the same on 8.1.52 and carry to 8.3 unchanged.

## Dev gateway facts (verified 2026-06-13)

- Install: `/usr/local/ignition` (default dir). Running as the `apple` user.
- Version: **8.1.52 (64-bit)**, status RUNNING. HTTP `:8088`, HTTPS `:8043`.
- Modules installed include **Perspective, Reporting, Vision, SQL Bridge, Tag Historian, OPC-UA**.
- Reachable: `curl http://localhost:8088/StatusPing` → `{"state":"RUNNING"}`.

## Non-goals

- Don't build 8.3-only features on the dev box and "hope" — if it can't run on 8.1.52, guard it and
  exercise the fallback locally.
- Don't downgrade the *production* design to the dev box's ceiling — target 8.3; just keep it runnable
  on 8.1.52.

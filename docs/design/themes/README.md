# TAI suite Perspective theme — `tai-light` / `tai-dark`

A shared, **gateway-level** design system for *every* Perspective project in the
Failproof/TAI suite (InventorySystem, GALC, MES, Manifest printer, Admin).

In Ignition, **themes are gateway-scoped, not project-scoped.** They live in the
gateway's `…/data/modules/com.inductiveautomation.perspective/themes/` folder and
are shared by all Perspective projects on that gateway. Once these files are
deployed, any project inherits the whole palette just by setting its Perspective
Session `theme` property to `tai-light` or `tai-dark`. This is what makes the
landing-page identity carry through to every form — and across every project.

Decision + concepts: `../landing-mockups/` and memory `project-ui-design-direction`.

## Files / why it's structured this way

| File | Role |
|---|---|
| `tai-light.css`, `tai-dark.css` | **Entry files** — must sit at the themes root; they register the theme and populate the Designer's theme list. Each imports IA's base theme, then our tokens, then the adapter. |
| `tai/tokens.base.css` | Type, radii, motion — identical in light and dark. |
| `tai/tokens.light.css`, `tai/tokens.dark.css` | The brand palette (`--tai-*`). **Edit these to retheme the whole suite.** |
| `tai/map.css` | **Adapter** — maps IA's stock theme variables onto our `--tai-*` tokens so every built-in component recolors for free. **The only file that references IA's internal variable names**; if an Ignition upgrade (8.1 → 8.3) renames them, this is the single place to fix. |
| `tai/fonts.css` | Self-hosted Inter + JetBrains Mono, embedded as base64 data-URIs. **Generated** by `scripts/build_fonts_css.py` from `fonts/` — don't hand-edit. |

We never edit IA's `light/`/`dark/` folders — they're overwritten on gateway
startup. We only *import* from them.

## How another project (e.g. GALC) inherits it

1. Deploy these files to the gateway (see below) — done once per gateway.
2. In that project's Perspective Session Props, set `theme = tai-light`.
3. For custom components, reference the **stable** `--tai-*` tokens in style
   classes / inline styles (they won't shift on upgrade), e.g.
   `backgroundColor: --tai-surface`, `color: --tai-status-danger`,
   `fontFamily: --tai-font-mono`.
4. Optional dark toggle: a button/action running
   `system.perspective.setTheme('tai-dark')` (persist the choice per user).

## Tokens

Light → Dark:

- `--tai-bg` `#F5F7FB` → `#0A111F`  · `--tai-surface` `#FFFFFF` → `#121C2E`
- `--tai-border` `#E3E9F2` → `#243049`  · `--tai-text` `#0E1B2E` → `#E7EEF9`
- `--tai-accent` `#1E5BBF` → `#5B8DE8`
- status (ISA-101, same meaning both modes; lightened for dark):
  `--tai-status-danger` `#D32F2F`/`#EF5350` · `warning` `#F57C00`/`#FB8C00` ·
  `caution` `#FBC02D`/`#FFD54F` · `info` `#1976D2`/`#42A5F5` ·
  `success` `#2E7D32`/`#66BB6A`
- type: `--tai-font-sans` (Inter → Noto Sans → system), `--tai-font-mono` (JetBrains Mono → system)
- radii: `--tai-radius` 12px (cards) · `--tai-radius-control` 8px · `--tai-radius-input` 6px

**Status colors mean "attention" (ISA-101).** Normal UI stays calm; color appears
only when something needs action.

> Fonts: Inter + JetBrains Mono are **self-hosted** — embedded as base64 data-URIs in
> `tai/fonts.css`. (The gateway `/res/perspective/fonts` route serves only the module's
> *bundled* fonts, so data-dir woff2 404 even after a restart; embedding is the portable,
> offline-safe fix.) `local()` is listed first, so an installed copy wins without decoding
> the data-URI. Source woff2 live in `fonts/`; regenerate with
> `python3 scripts/build_fonts_css.py`.

## Deploy to a gateway

```sh
DEST="/usr/local/ignition/data/modules/com.inductiveautomation.perspective/themes"
cp tai-light.css tai-dark.css "$DEST"/
mkdir -p "$DEST/tai" && cp tai/*.css "$DEST/tai/"
```

New entry files appear in the Designer's theme dropdown automatically (the gateway
watches this folder). If they don't show, restart the gateway to force detection —
mind the restart throttle on the dev box. **Editing an imported partial
(`tokens.*` / `map` / `fonts`) does NOT recompile a cached theme — `touch` the entry
file (`tai-light.css` / `tai-dark.css`), or restart, to force a recompile.**

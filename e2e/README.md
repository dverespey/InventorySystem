# Perspective E2E harness (Playwright, headless)

Drives the live dev gateway so development never waits on a manual browser click.
See `docs/automated-ui-testing.md` for the why/approach.

## One-time
    python3 -m pip install playwright && python3 -m playwright install chromium
    cp scripts/e2e/.env.example scripts/e2e/.env   # fill in gateway admin login

The `.env` creds are used ONLY to auto-reset the Perspective 2-hour trial (the dev
gateway is unlicensed). `.env` and `artifacts/` are gitignored.

## Run
    python3 scripts/e2e/reset_trial.py             # reset the 2h Perspective trial
    python3 scripts/e2e/test_order_spike.py        # headless: assert + screenshot
    python3 scripts/e2e/test_order_spike.py --headed   # watch it live (design review)

Screenshots land in `scripts/e2e/artifacts/` (order_spike_load.png, _after_edit.png).
Exit code 1 if any check FAILs. Selectors use component `domId` when present
(harden the view by adding `#spike-order-grid`, `#spike-simulate-btn`, etc.).

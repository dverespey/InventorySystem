"""Reset the Ignition Perspective 2-hour module trial via the gateway web UI.

The dev gateway runs unlicensed → Perspective sessions expire after 2 hours and
must be reset ("Sign in to Reset" on the gateway home). This automates that:
log in to the Ignition IdP, click Reset Trial. Needs gateway admin creds once
(env GATEWAY_USER/GATEWAY_PASS or e2e/.env) — NEVER hardcode them.

Usage:  python3 e2e/reset_trial.py [--headed]
Returns 0 on reset (or if already active), 2 if creds missing, 1 on failure.
"""
import sys
from playwright.sync_api import sync_playwright
import lib


def _expired(page):
    return "Trial Expired" in page.inner_text("body")


def reset_trial(page):
    page.goto(lib.BASE + "/web/home", wait_until="networkidle", timeout=30000)
    body = page.inner_text("body")
    if "Trial" not in body and "trial" not in body:
        return True, "no trial banner (licensed)"
    # Already running (not expired) — Perspective will render; reset not required.
    if not _expired(page):
        if page.query_selector("text=Reset Trial"):   # extend opportunistically
            page.click("text=Reset Trial")
            page.wait_for_load_state("networkidle", timeout=15000)
            return True, "trial active (extended)"
        return True, "trial active (not expired)"

    # Expired + already logged in → just click Reset Trial.
    if page.query_selector("text=Reset Trial"):
        page.click("text=Reset Trial")
        page.wait_for_load_state("networkidle", timeout=15000)
        return True, "reset (was already logged in)"

    user, pw = lib.creds()
    if not user or not pw:
        return False, "NEED_CREDS"

    signin = page.query_selector("text=Sign in to Reset") or page.query_selector("text=Log In")
    if not signin:
        return False, "no sign-in affordance found"
    signin.click()
    # Ignition IdP is a TWO-STEP form: username, Enter -> password field reveals.
    # (Target input[name=password] specifically; the IdP also renders unnamed
    #  decoy/honeypot password inputs — do NOT fill those.)
    page.wait_for_selector("input[name=username]", timeout=20000)
    page.fill("input[name=username]", user)
    page.press("input[name=username]", "Enter")
    page.wait_for_selector("input[name=password]:visible", timeout=20000)
    page.fill("input[name=password]", pw)
    page.press("input[name=password]", "Enter")
    page.wait_for_load_state("networkidle", timeout=20000)

    # back on home — click Reset Trial
    page.goto(lib.BASE + "/web/home", wait_until="networkidle", timeout=30000)
    rt = page.query_selector("text=Reset Trial")
    if rt:
        rt.click()
        page.wait_for_load_state("networkidle", timeout=15000)
        return True, "logged in + reset"
    # Some builds reset on the licensing page; report what we see
    return False, "logged in but no Reset Trial button (check login succeeded)"


def main():
    headed = "--headed" in sys.argv
    with sync_playwright() as p:
        b = p.chromium.launch(headless=not headed)
        pg = b.new_page(viewport={"width": 1400, "height": 900})
        ok, msg = reset_trial(pg)
        pg.screenshot(path=lib.ARTIFACTS + "/trial_reset.png", full_page=True)
        b.close()
    print(("OK: " if ok else "FAIL: ") + msg)
    if msg == "NEED_CREDS":
        print("Set GATEWAY_USER / GATEWAY_PASS in e2e/.env (gitignored) or env.")
        return 2
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

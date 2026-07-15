# Form-UX semantics: `TNewPasswordDlg` — `NewPassword.pas` / `NewPassword.dfm`

**Confirmed LIVE** (`InventorySystem.dpr:56`). Forced-password-reset dialog. **Not reached from
`UserAdmin`** — it is invoked from the shared ancestor `DataModule.TData_Module.ValidateUser`
(`DataModule.pas:6101-6130`), part of the **login** flow, when the operator's password has expired
(`UPDATEDIFF >= IN_PASSWORD_RESET_DAYS`, or that field is NULL). See `auth-users.md` §"Password
expiry / forced reset" for the live-vs-snapshot caveat on whether this branch is actually reachable
against the current schema.

## Dialogs & confirmations
- **Only one dialog, and it's a retry loop, not a confirmation:** if the two password fields don't
  match (or the new password is blank), `SHowMessage('Password does not match')` (`:62` — note the
  literal casing typo `SHowMessage`, still valid Pascal since identifiers are case-insensitive, but
  worth flagging as a code-quality artifact, not a behavior). No armed/two-step confirmation exists
  anywhere on this form — clicking OK with matching, non-blank fields commits immediately.
- **No cancel-confirmation** — clicking Cancel (`ModalResult=2`/`mrCancel`, `.dfm:56`) closes
  immediately with `Cancel:=TRUE`. The caller (`DataModule.ValidateUser:6104-6109`) treats
  `NewPasswordDlg.Cancel=True` as a **failed login** (`Result:=False`, logs `'LOGIN ERR'`) — i.e.
  **declining to set a new password when one is required aborts the login attempt entirely**, it
  does not fall back to the old password.

## Field clear / repopulate
- **`fCancel` defaults to `TRUE` on every show** (`FormShow`, `:47-50 fCancel:=TRUE;
  NewPassword.SetFocus;`) — so if the dialog is closed by any means other than a successful
  `OKBtnClick` (e.g. the window's system-close, if enabled — `BorderStyle=bsDialog` typically hides
  the close box, but no `BorderIcons=[]` override is present in `.dfm` the way `ForecastBreakdown_Form`
  sets it), the caller sees `Cancel=True` by default — **fail-safe**: an unexpected dismissal is
  treated as "no new password provided," not as silently keeping the old one.
- **On mismatch**, only `ConfirmPassword` is cleared and refocused (`:63-64
  ConfirmPAssword.Text:=''; ConfirmPAssword.SetFocus;`) — **`NewPassword` (the first field) is left
  holding whatever the operator typed**, so a retry only needs to re-type the confirmation, not
  both fields. This is the opposite of a "clear both on error" pattern.
- **No pre-fill of either field** — both start blank per `.dfm` design-time defaults (no `Text`
  property set on either `TEdit`); there is no "show the current password" behavior here (contrast
  `UserAdmin`, which does surface the existing plaintext password on user selection).

## Focus & keyboard
- **Initial focus is explicit, not `.dfm`-driven**: `FormShow` calls `NewPassword.SetFocus`
  (`:49`), even though `NewPassword` is already `TabOrder=0` (`.dfm:37`) — belt-and-suspenders, same
  effective result.
- `OKBtn` has `Default = True` (`.dfm:45`) — **Enter from either password field submits** (unlike
  every admin/EDI form covered elsewhere in this batch, which have no Default button at all).
  `CancelBtn` has `Cancel = True` (`.dfm:54`) — **Escape cancels**. This is the **only form in this
  set with both Enter-submits and Escape-cancels wired** via VCL flags rather than left to
  accelerator keys alone.
- On mismatch, focus moves to `ConfirmPassword` (`:64`) as noted above.
- Both edits are `CharCase = ecUpperCase` (`.dfm:34,64`) and `PasswordChar = '*'` (masked,
  `.dfm:36,66`) — passwords are entered upper-case and masked on screen, consistent with the rest
  of the auth family, though the underlying store is still plaintext (see `auth-users.md`).

## Enable/disable state machine
- **None.** OK and Cancel are always enabled; there is no live "passwords match" indicator or
  Enabled-gating of OK — mismatch is caught reactively only when OK is clicked.

## Error surfacing
- The single mismatch/blank case is the **only** error path on this form, surfaced via
  `ShowMessage` (see Dialogs above). There is no try/except in `NewPassword.pas` at all — any
  exception (there is essentially nothing that could throw here beyond the string compare) would
  propagate to the VCL default handler. The actual persistence call
  (`Data_Module.UpdateUserPassword` → `UPDATE_UserPassword`) happens in the **caller**
  (`DataModule.pas:6113`), not on this form, and that proc is **confirmed absent from the schema
  snapshot** (`auth-users.md` §3 — "CALLED BUT ABSENT FROM SCHEMA") — so a live run through this
  path may fail *after* this dialog reports success, with the failure surfacing back in
  `ValidateUser`'s own `MessageDlg('Unable to validate your' + #13 + 'information at this
  time.', mtError, [mbOK], 0)` (`DataModule.pas:6154`), not from this form.

## Cross-refs
- `docs/analysis/admin/auth-users.md` — §"Password expiry / forced reset" (the live-vs-snapshot
  question of whether this dialog is reachable today) and §3 (`UPDATE_UserPassword` missing-proc
  finding).
- `docs/analysis/cross-cutting/form-ux/UserAdmin.md` — the sibling admin screen; note this dialog
  is **not** reached from there.

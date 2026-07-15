# Form-UX semantics: `TConfirmPassword_Form` — `ConfirmPassword.pas` / `ConfirmPassword.dfm`

## Live-vs-dead verdict: **DORMANT — compiles and ships, but its one call site is fully commented
out**

Registered live in `InventorySystem.dpr:22` and compiles cleanly (unlike the two dead forecast
units above). Its only intended call site, `MainMenu.Administration_UserAdmin_MenuItemClick`
(`MainMenu.pas:548-561`), has the **entire gate commented out**:
```
//ConfirmPassword_Form := TConfirmPassword_Form.Create(self);
//If ConfirmPassword_Form.Execute Then
//Begin
  //ConfirmPassword_Form.Free;
  UserAdmin_Form := TUserAdmin_Form.Create(Self);
  UserAdmin_Form.Execute;
  UserAdmin_Form.Free;
//End
//Else
  //ConfirmPassword_Form.Free;
```
So **today, clicking the Administration → User Admin menu item opens `UserAdmin_Form` directly,
with no re-authentication step.** A repo-wide search confirms no other unit references
`ConfirmPassword_Form`/`TConfirmPassword_Form`. Already flagged this way in
`docs/analysis/admin/auth-users.md` §5 ("wired but commented out … Treat as dormant (not dead —
it ships, just isn't reached)"). This file adds the UX-layer detail for that dormant path, in case
the rebuild is asked to reinstate a re-auth gate in front of the Admin/Users screen.

## Dialogs & confirmations
- **Blank password** — `MessageDlg('The Password box is blank.' + #13 + 'Please try again.',
  mtInformation, [mbOK], 0)` (`:44-45`), then refocuses the password field (`:46`). Checked via
  `Trim(Password_Edit.Text) = ''` before attempting any comparison.
- **Wrong password** — `MessageDlg('The password you entered is invalid.' + #13 + 'Please try
  again.', mtInformation, [mbOK], 0)` (`:60-61`), preceded by clearing the field and refocusing
  (`:58-59`) — so the operator sees an already-blanked box the instant the dialog is dismissed,
  ready to retry.
- **Correct password** — `ModalResult := mrOK` (`:52`), no dialog; the form just closes.
  `Data_Module.LogActLog('USER ADMIN', 'Confirmed password to enter User Administration.', 0)`
  fires on success (`:53`); `Data_Module.LogActLog('ADMIN ERR', 'FAILED to match password to enter
  User Administration.', 0)` fires on mismatch (`:57`) — **both logged regardless of the dead call
  site**, i.e. if this form were ever invoked manually or reinstated, the audit trail already
  works.
- **No retry-limit** — the operator can retype indefinitely; there is no lockout/attempt counter
  anywhere in this unit (contrast a modern login throttle).
- **Uncaught-exception fallback in `Execute`** — `showMessage('Unable to generate logon screen.'
  + #13 + 'ERROR:' + #13 + E.Message)` (`:74-77`) — the same copy-pasted "logon screen" text
  found verbatim in `UserAdmin.Execute` (`UserAdmin.pas:67`), confirming both trace to a shared
  template/ancestor pattern in this app's `Execute` idiom, not independently authored text.

## Field clear / repopulate
- **`Execute` blanks the password field on every open**: `Password_Edit.Text:='';` (`:69`), before
  `ShowModal` — so even if this form were reinstated and shown multiple times in a session, no
  previous attempt's text (masked or otherwise) persists.
- **On a wrong-password result**, the field is explicitly re-blanked (`:58`) before the dialog
  even appears — the operator never has to manually clear it to retry.
- **On a blank-password result**, the field is NOT re-blanked (it's already empty by definition of
  that branch) — only refocused.

## Focus & keyboard
- No `ActiveControl` set in `.dfm`; default focus by tab order goes to `bitBtnLogon`
  (`TabOrder=1`) — **not** the password edit (`Password_Edit` is `TabOrder=0` inside `Panel2` but
  `.dfm` doesn't set `ActiveControl` at the form level, and no `FormShow`/`FormActivate` handler
  exists to `SetFocus` the edit either). **This means an operator opening this dialog must click
  or Tab into the password field before typing** — no explicit initial-focus call exists anywhere
  in `ConfirmPassword.pas`, unlike every other password-entry form in this batch
  (`NewPassword.FormShow` and `Logon`, per `auth-users.md`, both explicitly focus their password
  field). Flag as a likely oversight if this dialog is ever reinstated.
- `bitBtnLogon` (`TBitBtn`, caption `&Enter`) has `Default = True` (`.dfm:59`) — Enter submits from
  anywhere on the form. `BitBtnCancel` has `Cancel = True` (`.dfm:83`) — Escape cancels
  (`ModalResult=2`/`mrCancel`).
- Refocus-to-`Password_Edit` happens explicitly on both failure paths (`:46`, `:59`), so after a
  failed attempt the field-focus gap above is moot — it's only the very first open that has no
  explicit focus call.

## Enable/disable state machine
- **None.** Enter/Cancel are always enabled; there is no live-validation Enabled-gating of the
  Enter button based on whether the field is populated.

## Error surfacing
- All via `MessageDlg` (cataloged above) — no log pane on this form; only `LogActLog` calls
  (Activity DB), not surfaced to the operator beyond the dialogs.
- No `try/except` inside `bitBtnLogonClick` itself — the plaintext compare
  (`Data_Module.gobjUser.AppUserPass = Trim(Password_Edit.Text)`, `:50`) cannot realistically
  throw, so this is a non-issue in practice.

## Cross-refs
- `docs/analysis/admin/auth-users.md` §5 — the data-side "dormant, not dead" verdict this file
  elaborates with dialog-by-dialog UX detail.
- `docs/analysis/cross-cutting/form-ux/UserAdmin.md` — the screen this form was meant to gate,
  reached today with no re-auth step.

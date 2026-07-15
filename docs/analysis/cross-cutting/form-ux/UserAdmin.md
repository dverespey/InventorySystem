# Form-UX semantics: `TUserAdmin_Form` — `UserAdmin.pas` / `UserAdmin.dfm`

**Confirmed LIVE** (`InventorySystem.dpr:23`). User CRUD screen: combobox of existing
`(UserID, plaintext Password, IsAdmin)` rows, Insert/Update/Delete/Clear/Close. Reached from
`MainMenu.Administration_UserAdmin_MenuItemClick` (`MainMenu.pas:548-561`) — **directly**, with no
gate (see `ConfirmPassword.md` for the dead re-auth gate that would otherwise sit in front of this
screen). Full data/proc/security spec: `docs/analysis/admin/auth-users.md`.

## Dialogs & confirmations
- **Insert — validation, not armed:**
  - Any of UserID/Password/Confirm blank → `MessageDlg('Please fill in all boxes.',
    mtInformation, [mbOK], 0)` (`:115-116`).
  - Password ≠ Confirm → `MessageDlg('Please make sure that the password' + #13 + 'and confirm
    password fields match.', mtError, [mbOK], 0)` (`:132`).
  - Duplicate (id+password pair already exists) → `MessageDlg('This User ID and Password already
    exist.' + #13 + 'Please try again.', mtError, [mbOK], 0)` (`:122`).
  - **No confirmation before a successful insert** — once validation passes and the dup-check
    clears, the row is written immediately.
- **Update — same two validation dialogs as Insert** (blank-fields `:140-141`; mismatch `:151`),
  **and also no confirmation before a successful update.**
- **Delete — the one armed/two-step confirmation on this form:**
  `MessageDlg('Are you sure you wish to delete ' + UserAdminDetail.ID + '?', mtConfirmation,
  [mbOK, mbCancel], 0) = mrOK` (`:164`). Note the button set is **OK/Cancel, not Yes/No** (unlike
  `SizeMaster`'s delete confirm, which uses `[mbYes, mbNo]`) — inconsistent phrasing across the
  app's delete-confirm dialogs; `mrOK` proceeds, anything else (including dialog dismissal) no-ops.
  **No selection guard needed on Delete itself** — see field-clear note below; a separate blank-
  combobox check runs first.
- **No selection / blank combobox on Delete** — `MessageDlg('Please select a user from the drop
  down list prior to pressing delete.', mtInformation, [mbOK], 0)` (`:160`), checked via
  `Trim(ExistingUserIDs_ComboBox.Text) = ''` **before** reaching the delete-confirm dialog above.
- **Uncaught-exception fallback** — `Execute`'s own `try/except`: `showMessage('Unable to generate
  logon screen.' + #13 + 'ERROR:' + #13 + E.Message)` (`:66-69`) — note the message text says
  "logon screen," a copy-paste artifact from a sibling form (likely `Logon.pas`/
  `ConfirmPassword.pas`), not accurate to this screen. Flag: don't reproduce this misleading string
  in the rebuild.

## Field clear / repopulate
- **Selecting an existing user AUTO-FILLS the password field with the stored plaintext password**
  (`ExistingUserIDs_ComboBoxChange`, `:89-103`): `UserID_Edit.Text := UserAdminDetail.ID;
  Password_Edit.Text := UserAdminDetail.Pass; IsAdmin_CheckBox.Checked := UserAdminDetail.Admin;
  ConfirmPassword_Edit.Text := '';` — the password box is visually masked (`PasswordChar='*'`,
  `.dfm:143`) but its `.Text` holds the real plaintext (loaded into the combobox item object at
  `FormCreate`, `DataModule.pas:6194`, straight from `SELECT VC_PASSWORD FROM INV_USERS`).
  **`ConfirmPassword_Edit` is the one field that IS cleared on selection** — the operator must
  re-type the (now-visible-as-dots) current password into Confirm before Update will accept it
  unchanged. **This is the #135-class case in reverse**: here an incoming value is silently
  populated (not left empty) where a naive rebuild might instead leave the password field blank
  on selection (safer, but a behavior change from legacy) or, worse, round-trip the masked dots
  as a literal value.
- **Selecting index 0 (the blank placeholder row) clears the whole panel:**
  `If ExistingUserIDs_ComboBox.ItemIndex = 0 Then Data_Module.ClearControls(UserAdmin_Panel)`
  (`:93-94`). The blank row is injected by `SetComboBoxesWithUserObj`
  (`DataModule.pas:6191 Items.AddObject('', nil)`) as item 0 ahead of every real user row.
- **`Clear_Button`** (`:105-110`): resets `ComboBox.ItemIndex:=0`, `ClearControls(UserAdmin_Panel)`
  (blanks all Edits, unchecks the checkbox — see `ClearControls`, `DataModule.pas:5976-5998`), then
  `ExistingUserIDs_ComboBox.SetFocus`.
- **After a successful Insert** (`:124-129`): combobox is **fully re-queried** (fresh
  `SetComboBoxesWithUserObj` call, re-adding the blank row + all users including the new one),
  `ItemIndex:=0`, panel cleared, focus to combobox — i.e. **the operator is returned to a blank
  form**, not shown the just-inserted user (contrast `SizeMaster`, which re-selects and echoes back
  the just-inserted row).
- **After a successful Update** (`:148`): only the combobox is re-queried — **the detail panel is
  NOT cleared and NOT re-selected**; `UserID_Edit`/`Password_Edit`/etc. are left showing whatever
  was just typed (which may now be stale relative to the fresh combobox item objects, since those
  are brand-new `TUserAdminDetail` instances). `ItemIndex` is also left wherever it was (now
  pointing at a *different* object after the re-query, since the list was rebuilt) — a real
  stale-selection hazard: the combobox's `ItemIndex` no longer reliably corresponds to the row the
  operator thinks is selected until they interact with it again.
- **After a successful Delete** (`:166-171`): panel cleared, combobox re-queried, `ItemIndex:=0`,
  focus to combobox — same "return to blank" pattern as Insert.

## Focus & keyboard
- No `ActiveControl` in `.dfm`; default focus is `ExistingUserIDs_ComboBox` (`TabOrder=0` inside
  `UserAdmin_Panel`, the first panel in tab order).
- **All text entry is forced uppercase**: `EditKeyPress` (`:78-82`, `If Key in ['a'..'z'] then
  Dec(Key,32)`) is wired to `ExistingUserIDs_ComboBox` (`OnKeyPress`, `.dfm:114` — applies to
  typed filter text even though it's `csDropDownList`, a no-op in practice since that style
  disables free typing) **and** to `UserID_Edit`/`Password_Edit`/`ConfirmPassword_Edit`
  (`.dfm:129,145,161`) — every credential the operator types is coerced to upper-case as they type,
  matching the login-side uppercase-forcing noted in `auth-users.md` §4.
- Explicit `SetFocus` calls: `ExistingUserIDs_ComboBox.SetFocus` after Clear (`:109`), after
  successful Insert (`:128`), after successful Delete (`:170`) — **not** after a successful Update
  (no focus move at all in that branch).
- `Close_Button` has `ModalResult = 2` (`mrCancel`, `.dfm:205`) with **no `Cancel=True`** flag set
  — Escape is not VCL-wired to Close; only clicking it (or its `&Close` accelerator) works. No
  button on this form has `Default=True` either — Enter does not trigger Insert/Update from an
  arbitrary field.
- Accelerator keys present on all five buttons (`&Insert`, `&Update`, `Cl&ear`, `&Close`,
  `&Delete`, `.dfm:177,186,195,204,213`).

## Enable/disable state machine
- **None found.** No button or edit has its `Enabled` toggled anywhere in `UserAdmin.pas` —
  Insert/Update/Delete are always clickable regardless of whether a row is selected; the
  form relies entirely on **reactive validation dialogs** (blank fields, no selection, mismatch,
  duplicate) rather than disabling controls to prevent invalid actions. This matches the
  `SizeMaster` pattern (no proactive Enabled-gating anywhere in that master-data family either).

## Error surfacing
- All surfaces are `MessageDlg`/`showMessage` dialogs (cataloged above) — there is no status
  label or log pane on this form. `Insert_ButtonClick`/`Update_ButtonClick`/`Delete_ButtonClick`
  have no `try/except` of their own; an exception thrown by
  `Data_Module.InsertUser`/`UpdateUserInfo`/`DeleteUserInfo` (all of which internally retry up to 3
  times per `auth-users.md` §"P12 retry-recursion check" before giving up and only `LogActLog`-ing)
  propagates unguarded to the VCL default exception handler if it still fails after retries — a
  generic Delphi error box, not a form-authored message, same posture as `SizeMaster`.

## Cross-refs
- `docs/analysis/admin/auth-users.md` — full proc/schema/security spec (plaintext passwords,
  composite id+password identity, retry-recursion confirmation, target-Ignition role model).
- `docs/analysis/cross-cutting/form-ux/ConfirmPassword.md` — the dead re-auth gate that would
  otherwise sit in front of this screen (`MainMenu.pas:551-560`, fully commented out).
- `docs/analysis/cross-cutting/form-ux/NewPassword.md` — the *other* password dialog in this
  family, reached from `Logon`/`ValidateUser`, not from this screen.

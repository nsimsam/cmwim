# Canadian Muslim Women in Medicine

Task board for the 2026-27 year, built from the annual plan and the Oct 15, 2026 meeting. Served at https://noorsimsam.com/cmwim/ once GitHub Pages is on.

The **Tasks** tab of the [CMWIM task sheet](https://docs.google.com/spreadsheets/d/1M7I2oyw726dc9xS642O1nkgt7-_U7omqBkMSNJGvm5I/edit) in Google Drive is the master list. The team can edit tasks in the Sheet or on the board (after unlocking it with the team passcode). Keep the row 1 headers (Initiative, Task, Owner, Due, Status, Notes, ID) as they are.

- **Initiative** and **Status** must match the dropdown choices on the Lists tab. A task with an unknown initiative shows under "Other".
- **Due** is a date, or blank.
- **ID** is filled in by the script. Leave it alone.

## Editing from the board

The board writes to the Sheet through [`apps-script/Code.gs`](apps-script/Code.gs), deployed from the Sheet as a web app (Extensions → Apps Script, then Deploy → Web app, execute as me, access: anyone). The passcode is the `PASSCODE` script property. The web app URL goes in `API` in `index.html`; while `API` is empty the board is read-only.

The timeline shades the no-event dates (exams, Ramadan, Eid) listed in `AVOID` in `index.html`.

Everything depends on team capacity and member interest, so events can be scaled back, pushed to next year or dropped.

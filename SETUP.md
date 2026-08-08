# Monster truck request page

A one-page site where students request the Monster Energy truck for their event,
or pitch a collab / sponsorship. Submissions land in a Google Sheet.

## 1. Publish the page

Settings > Pages > Source: **Deploy from a branch** > Branch: `main` / `root` > Save.

Live at: `https://kevincecil11.github.io/monster-energy-campus/`

## 2. Connect Google Sheets

1. Create a new Google Sheet. Copy the ID out of the URL:
   `https://docs.google.com/spreadsheets/d/`**`<THIS_PART>`**`/edit`
2. In that Sheet: **Extensions > Apps Script**. Delete the sample code and paste
   everything from `apps-script.gs`.
3. Set `SHEET_ID` at the top. Optionally set `NOTIFY_EMAIL` to get an email per request.
4. **Deploy > New deployment > Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Authorize it (Google will warn about an unverified app: Advanced > Go to project > Allow).
6. Copy the deployment URL, ending in `/exec`.
7. In `index.html`, find `SHEET_ENDPOINT` near the bottom and paste the URL in.
8. Commit. Pages redeploys in about a minute.

Until step 7 is done, the form validates and shows the success screen but logs the
payload to the browser console instead of saving.

## 3. Sheet columns

| Received at | Reference | Request type | Name | Phone | Email | Venue | Address | Event date | Expected crowd | Notes | Status |

`Status` starts as `New`. Change it to `Contacted` / `Confirmed` / `Declined` as
you work the list. Add a dropdown on that column to keep it clean.

## Notes

- Every submission gets a reference code (`MTR-YYMMDD-XXXX`) shown to the user and
  stored in the sheet, so phone follow-ups have something to quote.
- The event date field enforces a 3 week minimum lead time. Change the `21` in
  `index.html` if that window shifts.
- Redeploy the Apps Script (**Deploy > Manage deployments > edit > New version**)
  any time you change `apps-script.gs`.

## Disclaimer

Unofficial student ambassador page. Monster Energy and the claw mark are
trademarks of Monster Beverage Corporation.

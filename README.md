# TDT DATA — Google Apps Script Patient Data Entry

Mobile-first patient data entry and review app using the user's existing Google Sheet.

## Release

**v2.1** — quick-touch input layout: uppercase names, single-row digit keypads for CR and 10-digit mobile numbers, age default 13 with large steppers, height 120 cm and weight 30 kg starting values, selectable sex / YES-NO / Tanner staging tiles, ferritin year selection limited to 2022–2026. Supports nine existing ferritin pairs plus additional ferritin value/date pairs appended after original columns A:AM. IAP 2015 age/sex 50th-percentile reference is included for supported ages.

## Structure

- `TDT_Data_UI.html` — mobile HTML/CSS/JS.
- `TDT_Data_Server.gs` — Apps Script server and guarded Google Sheets access.
- `appsscript.json` — V8 runtime manifest.

## Before deploying

1. **This repository is public.** Source code only: never commit patient data, exports, credentials, `.clasp.json`, or real patient screenshots.
2. In Apps Script **Project Settings → Script properties**, add `TDT_SPREADSHEET_ID` with the real spreadsheet ID (do not commit it to GitHub). In `TDT_Data_Server.gs`, set authorized `editorEmails` / `viewerEmails`, verify `sheetGid` and the exact **39 original column headings** in A:AM. The default empty editor allowlist blocks writes.
3. In the bound Apps Script project, use the same names for the `.gs` and `.html` files and keep `appsscript.json` as the manifest.
4. Run `tdtValidateSetup` and test create/update/CR lookup/ferritin 10+ with fictional records on a **copy** of the target spreadsheet.
5. Use Apps Script **Deploy → Manage deployments → Edit → New version → Deploy** after testing. GitHub commits do **not** deploy the web app automatically.

### Safety

This release has not been tested against live patient data or the actual Google Sheet; the connected account lacked permission. Input defaults for age, height, and weight are convenience placeholders, **not real patient measurements**. Verify before saving. Clinical growth references are not diagnoses.
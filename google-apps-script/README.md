# Worship XLSX Reader (Google Apps Script)

Parses the **SYNC EM Worship** `.xlsx` file directly from Google Drive in memory — no conversion to Google Sheets, no writes to the file.

## Quick setup

1. Open [Google Apps Script](https://script.google.com/) → **New project**.
2. Paste `WorshipSyncReader.gs` into the editor.
3. **(Recommended)** Add SheetJS as a second script file for faster runs:
   - Download [xlsx.full.min.js (v0.20.3)](https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js)
   - In Apps Script: **+** → **Script** → name it `xlsx.full.min.js` → paste the file contents
   - Comment out the `eval(...)` line in `loadSheetJS_()` if you do this
4. Run `testGetCurrentWeekRow()` → approve Drive permissions.
5. Use `getCurrentWeekWorshipRow()` from triggers or other scripts.

## File details

| Setting    | Value                                      |
|-----------|---------------------------------------------|
| Drive ID  | `1BSuNwm8YoW_UJkLoZuWY9Oqgy0hVpG-9`        |
| Sheet tab | `2026`                                      |
| Date col  | A (`Date` header on row 3)                  |
| Format    | `DD/MM/YYYY` (e.g. `05/07/2026`)           |

## How matching works

By default (`matchStrategy: 'sunday'`), the script finds the row whose **column A date equals the Sunday of the current week**. For today (5 Jul 2026, a Sunday), that matches row with `05/07/2026`.

Set `matchStrategy: 'same-week'` in `CONFIG` if you instead want any row whose date falls in the same Sunday–Saturday week as today.

## Return value

```javascript
{
  rowNumber: 32,
  date: Date(2026-07-05),
  headers: ['Date', 'Preparation', 'Call to Worship', 'Adoration', 'Assurance of Pardon'],
  row: {
    Date: '05/07/2026',
    Preparation: '...',
    'Call to Worship': 'Phil 4:4-7',
  },
  rawRow: ['05/07/2026', '...', 'Phil 4:4-7', ...]
}
```

## Important notes

- The Drive file must remain a native `.xlsx` (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`). If you fully convert it to Google Sheets, use `SpreadsheetApp` instead.
- Set your script **time zone** under Project Settings to match your locale (e.g. `Australia/Sydney`) so “current week” is correct.
- The script only **reads** the file; it never modifies Drive or the workbook.

# Employee Form → Excel  (PowerShell edition)

HTML/JS form; a **pure PowerShell** web server appends every submission as a new row in an Excel workbook. No Python, no Node, no build step.

**Excel file:** `D:\employee\data\form.xlsx` (override with `-ExcelPath` or `$env:EXCEL_PATH`)

The `D:\employee\data` folder is **created automatically** if it doesn't exist — you don't need to make it yourself. You only need drive `D:` to exist and be writable.

## Files
```
employee-form/
├── server.ps1           # PowerShell HttpListener server + ImportExcel writer
├── start.cmd            # double-click launcher for Windows
├── static/index.html    # the form UI (HTML + CSS + vanilla JS)
└── data/form.xlsx       # fallback only (used if D: is missing/unwritable, e.g. on Linux/macOS)
```

## Run
```powershell
pwsh -File server.ps1                                  # macOS / Linux / PS7
powershell -ExecutionPolicy Bypass -File server.ps1     # Windows PowerShell 5.1
.\start.cmd                                             # or just double-click
```
Options:
```powershell
pwsh -File server.ps1 -Port 9000                        # force a port
pwsh -File server.ps1 -ExcelPath 'E:\other\form.xlsx'   # different workbook
pwsh -File server.ps1 -NoBrowser                        # don't auto-open browser
```

## Auto dependency install
On startup `server.ps1` checks and installs anything missing, then reports each step:
1. Verifies PowerShell ≥ 5.1 and enables **TLS 1.2** (needed by old Windows PS to reach PSGallery).
2. Installs the **NuGet package provider** if absent.
3. Marks **PSGallery** trusted so the install is non-interactive.
4. Installs **ImportExcel** (`-Scope CurrentUser`, no admin needed) if not already present, then imports it.
5. Confirms `System.Net.HttpListener` is supported.

## Auto port detection
Tries `8080, 8000, 5000, 3000, 8888, 5050, 7070, 9090, 4200, 8081` in order and takes the first free one; if all are busy it asks the OS for any free ephemeral port. A `-Port` you pass is honoured unless it's already taken, in which case it falls back to auto-detect. The chosen port is printed at startup and shown in the page header.

## API
| Method | Route | Purpose |
|---|---|---|
| GET    | `/`                     | The form page |
| GET    | `/api/employees`        | Read all rows back as JSON |
| POST   | `/api/employees`        | Validate and append one row |
| PUT    | `/api/employees/{id}`   | Validate and update one row |
| DELETE | `/api/employees/{id}`   | Permanently delete one row |
| GET    | `/api/download`         | Download `form.xlsx` |
| GET    | `/api/info`             | Active port + real Excel path in use |

## Columns
`ID | Employee Name | Email | Phone | Department | Designation | Joining Date | Salary | Gender | Address | Submitted At`

## Edit or delete a record
Each saved row has **Edit** and **Delete** buttons in the sticky **Actions** column at the left of the table:
- **Edit** loads the row into the form. Choose **Update Record** to save changes or **Cancel Edit** to leave it unchanged.
- **Delete** asks for confirmation and then permanently removes the row from the active workbook.

## How records are stored (important)
The workbook at the predefined path is the **permanent database**. New submissions append rows; edits update the matching row; deletes remove the matching row.

- Records persist across server restarts and reboots until explicitly deleted.
- **You never have to download anything to save data.** `Download a copy (optional)` just hands you a snapshot copy; the master file on disk is already current the instant you hit Save.
- On startup the server reports how many records are already stored, e.g.
  `Existing workbook opened: D:\employee\data\form.xlsx  (37 record(s) already stored - new rows will be appended)`
- A rolling **startup backup** is copied to `data/backups/` (10 most recent kept).
- If the workbook is ever unreadable/corrupt it is **preserved** as `form.xlsx.corrupt-<timestamp>.bak` rather than deleted, and a fresh one is started.
- **Self-healing:** the file and its folder are re-created on demand. If the workbook or the whole `data/` folder is deleted while the server is running, the next request (submit, list, or info) rebuilds it automatically with the styled header row — no restart needed.

### "I don't see the Excel file / data folder"
The workbook is only written where the server can actually write:
1. Look in the folder printed at startup as `Excel file   <path>` — that is the real location, also shown in the page banner and at `/api/info`.
2. If `D:\employee\data` isn't writable (drive missing, no permission, or you're on Linux/macOS where `D:\` is meaningless), the server falls back to `<script folder>\data\form.xlsx` and logs a `!!` warning explaining exactly why.
3. The file is created at startup, so it exists even with zero records (header row only).
4. To force a specific location: `pwsh -File server.ps1 -ExcelPath 'E:\other\form.xlsx'`

## Notes
- **Why a server at all?** Browser JS can't write to a fixed disk path like `D:\employee\data\form.xlsx`. The PowerShell listener does the file write.
- Validation runs in the browser **and** again in PowerShell before anything is written.
- Duplicate emails are rejected with HTTP 409 (case-insensitive).
- Writes are serialised with a named **Mutex**, so concurrent adds, edits, and deletes can't corrupt the workbook.
- Salary is stored as a real number formatted `#,##0.00`; the header row is styled and frozen.
- Binds `http://+:PORT/`; on Windows without admin rights that needs a URL ACL, so it automatically retries on `localhost` only.
- If `D:\employee\data` isn't writable it falls back to `./data/form.xlsx` and the UI shows the real path in use.
- On Linux, ImportExcel may warn `Cannot Autosize ... libgdiplus`. Harmless — column widths are set explicitly. Silence it with `apt-get install -y libgdiplus libc6-dev`.

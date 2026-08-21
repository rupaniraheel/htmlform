<#
=============================================================================
 Employee Form  ->  Excel      (PowerShell HttpListener server)
=============================================================================
 * Installs every dependency it needs before starting (ImportExcel module,
   NuGet provider, PSGallery trust) -- no manual setup.
 * Auto-detects a free TCP port (tries 8080, 8000, 5000, 3000, 8888 ... then
   asks the OS for any free one).
 * Serves static/index.html and a small JSON API.
 * Appends every submitted record as a new row in the Excel workbook.

 Excel file : D:\employee\data\form.xlsx   (override: -ExcelPath or $env:EXCEL_PATH)
              The folder is created automatically if it does not exist.

 Run:
     pwsh -File server.ps1                       # any platform
     powershell -ExecutionPolicy Bypass -File server.ps1     # Windows PS 5.1
     pwsh -File server.ps1 -Port 9000 -ExcelPath 'E:\other\form.xlsx'
=============================================================================
#>

[CmdletBinding()]
param(
    [int]    $Port,                                   # force a port (else auto-detect)
    [string] $ExcelPath = $(if ($env:EXCEL_PATH) { $env:EXCEL_PATH } else { 'D:\employee\data\form.xlsx' }),
    [string] $BindHost  = '0.0.0.0',                  # 0.0.0.0 = listen on all interfaces
    [switch] $NoBrowser                               # don't auto-open the browser
)

$ErrorActionPreference = 'Stop'
$script:Root = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---------------------------------------------------------------------------
# Pretty logging
# ---------------------------------------------------------------------------
function Write-Step { param($m) Write-Host "  ->  $m"           -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "  OK  $m"           -ForegroundColor Green }
function Write-Warn2{ param($m) Write-Host "  !!  $m"           -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "  XX  $m"           -ForegroundColor Red }
function Write-Head {
    param($m)
    Write-Host ''
    Write-Host ("=" * 68) -ForegroundColor DarkGray
    Write-Host "  $m"     -ForegroundColor White
    Write-Host ("=" * 68) -ForegroundColor DarkGray
}

# ===========================================================================
# 1. DEPENDENCIES  --  install anything missing before the app starts
# ===========================================================================
function Initialize-Dependencies {
    Write-Head 'Checking dependencies'

    # --- PowerShell version -------------------------------------------------
    Write-Step "PowerShell $($PSVersionTable.PSVersion) on $([System.Environment]::OSVersion.Platform)"
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw 'PowerShell 5.1 or newer is required.'
    }

    # --- TLS 1.2 (needed by old Windows PowerShell to reach PSGallery) ------
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }

    # --- ImportExcel module -------------------------------------------------
    if (Get-Module -ListAvailable -Name ImportExcel) {
        $v = (Get-Module -ListAvailable -Name ImportExcel |
              Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-Ok "ImportExcel $v already installed"
    }
    else {
        Write-Warn2 'ImportExcel module not found - installing from PSGallery...'

        # NuGet package provider (Windows PowerShell needs this first)
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Write-Step 'Installing NuGet package provider...'
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        }

        # Trust PSGallery so the install is non-interactive
        $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
            Write-Step 'Trusting PSGallery repository...'
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }

        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
        Write-Ok  "ImportExcel $((Get-Module -ListAvailable ImportExcel | Select-Object -First 1).Version) installed"
    }

    Import-Module ImportExcel -DisableNameChecking -WarningAction SilentlyContinue
    Write-Ok 'ImportExcel module loaded'

    # --- HttpListener support ----------------------------------------------
    if (-not [System.Net.HttpListener]::IsSupported) {
        throw 'System.Net.HttpListener is not supported on this system.'
    }
    Write-Ok 'HttpListener supported'
}

# ===========================================================================
# 2. EXCEL FILE  --  resolve a writable path and create the workbook
# ===========================================================================
$script:Headers = @(
    'ID','Employee Name','Email','Phone','Department','Designation',
    'Joining Date','Salary','Gender','Address','Submitted At'
)
$script:Widths = @(6,24,28,18,22,22,14,14,10,34,20)

function Resolve-ExcelPath {
    param([string]$Preferred)

    $isWin = $IsWindows -or $env:OS -eq 'Windows_NT'

    # A Windows-style path (D:\...) is meaningless on Linux/macOS -> use the local
    # data folder instead of creating a literal folder called "D:\employee\data".
    if (-not $isWin -and $Preferred -match '^[A-Za-z]:[\\/]') {
        $fallbackDir = Join-Path $script:Root 'data'
        if (-not (Test-Path -LiteralPath $fallbackDir)) {
            New-Item -ItemType Directory -Path $fallbackDir -Force | Out-Null
        }
        $fb = Join-Path $fallbackDir 'form.xlsx'
        Write-Warn2 "'$Preferred' is a Windows path but this is $([System.Environment]::OSVersion.Platform)."
        Write-Warn2 "Using '$fb' instead. On Windows the configured path is used as-is."
        return $fb
    }

    $folder = Split-Path -Parent $Preferred
    if ([string]::IsNullOrWhiteSpace($folder)) { $folder = $script:Root }

    try {
        if (-not (Test-Path -LiteralPath $folder)) {
            Write-Step "Creating folder: $folder"
            New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
        }
        $probe = Join-Path $folder ".write_test_$PID"
        'ok' | Out-File -LiteralPath $probe -Encoding ascii -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $Preferred
    }
    catch {
        $fallbackDir = Join-Path $script:Root 'data'
        if (-not (Test-Path -LiteralPath $fallbackDir)) {
            New-Item -ItemType Directory -Path $fallbackDir -Force | Out-Null
        }
        $fb = Join-Path $fallbackDir 'form.xlsx'
        Write-Warn2 "Cannot write to '$Preferred'"
        Write-Warn2 "Reason: $($_.Exception.Message)"
        if ($isWin -and $Preferred -match '^([A-Za-z]):') {
            $drive = $Matches[1]
            if (-not (Test-Path -LiteralPath "${drive}:\")) {
                Write-Warn2 "Drive ${drive}: does not exist on this machine."
            }
        }
        Write-Warn2 "Falling back to '$fb'"
        return $fb
    }
}

function Initialize-Workbook {
    param([switch]$Quiet)

    if (Test-Path -LiteralPath $script:XlsxPath) {
        # File already exists -> NEVER recreate it, we only ever append to it.
        # Verify it's readable; if it's corrupt, move it aside instead of losing data.
        try {
            $probe = Open-ExcelPackage -Path $script:XlsxPath
            $sheet = $probe.Workbook.Worksheets['Employees']
            $existing = if ($sheet -and $sheet.Dimension) { $sheet.Dimension.End.Row - 1 } else { 0 }
            Close-ExcelPackage $probe -NoSave
            if (-not $Quiet) {
                Write-Ok "Existing workbook opened: $script:XlsxPath  ($existing record(s) already stored - new rows will be appended)"
            }
            return
        }
        catch {
            $backup = "$script:XlsxPath.corrupt-$(Get-Date -Format 'yyyyMMdd-HHmmss').bak"
            Write-Warn2 "Existing workbook could not be read - preserving it as '$backup' and starting a fresh one."
            Move-Item -LiteralPath $script:XlsxPath -Destination $backup -Force
        }
    }

    # make sure the containing folder still exists (it may have been deleted)
    $dir = Split-Path -Parent $script:XlsxPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Step "Re-created missing folder: $dir"
    }

    Write-Step "Creating workbook: $script:XlsxPath"
    $pkg = Open-ExcelPackage -Path $script:XlsxPath -Create
    $ws  = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Employees' -ClearSheet

    for ($i = 0; $i -lt $script:Headers.Count; $i++) {
        $c = $i + 1
        $ws.Cells[1, $c].Value                       = $script:Headers[$i]
        $ws.Cells[1, $c].Style.Font.Bold             = $true
        $ws.Cells[1, $c].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $ws.Cells[1, $c].Style.Fill.PatternType      = 'Solid'
        $ws.Cells[1, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
        $ws.Cells[1, $c].Style.HorizontalAlignment   = 'Center'
        $ws.Column($c).Width                         = $script:Widths[$i]
    }
    $ws.View.FreezePanes(2, 1)
    Close-ExcelPackage $pkg
    Write-Ok "Workbook created with header row"
}

function Get-Records {
    if (-not (Test-Path -LiteralPath $script:XlsxPath)) { return @() }
    $pkg = Open-ExcelPackage -Path $script:XlsxPath
    try {
        $ws   = $pkg.Workbook.Worksheets['Employees']
        $rows = @()
        if ($ws.Dimension) {
            for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
                $row = @()
                for ($c = 1; $c -le $script:Headers.Count; $c++) {
                    $v = $ws.Cells[$r, $c].Value
                    $row += $(if ($null -eq $v) { '' } else { [string]$v })
                }
                if (($row -join '').Trim()) { $rows += ,$row }
            }
        }
        return ,$rows
    }
    finally { Close-ExcelPackage $pkg -NoSave }
}

function Add-Record {
    param([hashtable]$Data)

    $mutex = New-Object System.Threading.Mutex($false, 'Global\EmployeeFormXlsx')
    [void]$mutex.WaitOne()
    try {
        Initialize-Workbook -Quiet
        $pkg = Open-ExcelPackage -Path $script:XlsxPath
        try {
            $ws = $pkg.Workbook.Worksheets['Employees']
            if (-not $ws) { $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Employees' }

            $lastRow = if ($ws.Dimension) { $ws.Dimension.End.Row } else { 1 }

            # duplicate email guard (column C)
            $email = ([string]$Data.email).Trim().ToLower()
            for ($r = 2; $r -le $lastRow; $r++) {
                $existing = [string]$ws.Cells[$r, 3].Value
                if ($existing.Trim().ToLower() -eq $email) {
                    return @{ ok = $false; duplicate = $true }
                }
            }

            $newRow = $lastRow + 1
            $id     = $lastRow            # header is row 1 -> id == record count + 1

            $ws.Cells[$newRow, 1 ].Value = $id
            $ws.Cells[$newRow, 2 ].Value = ([string]$Data.name).Trim()
            $ws.Cells[$newRow, 3 ].Value = ([string]$Data.email).Trim()
            $ws.Cells[$newRow, 4 ].Value = ([string]$Data.phone).Trim()
            $ws.Cells[$newRow, 5 ].Value = ([string]$Data.department).Trim()
            $ws.Cells[$newRow, 6 ].Value = ([string]$Data.designation).Trim()
            $ws.Cells[$newRow, 7 ].Value = ([string]$Data.joining_date).Trim()

            $sal = ([string]$Data.salary).Trim()
            if ($sal) {
                $ws.Cells[$newRow, 8].Value             = [double]$sal
                $ws.Cells[$newRow, 8].Style.Numberformat.Format = '#,##0.00'
            }

            $ws.Cells[$newRow, 9 ].Value = ([string]$Data.gender).Trim()
            $ws.Cells[$newRow, 10].Value = ([string]$Data.address).Trim()
            $ws.Cells[$newRow, 11].Value = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

            Close-ExcelPackage $pkg -Show:$false
            $pkg = $null
            return @{ ok = $true; id = $id; total = $newRow - 1 }
        }
        finally { if ($pkg) { Close-ExcelPackage $pkg -NoSave } }
    }
    finally { $mutex.ReleaseMutex(); $mutex.Dispose() }
}

# ===========================================================================
# 3. VALIDATION  (server side -- the browser validates too)
# ===========================================================================
function Test-Record {
    param([hashtable]$D)
    $e = @{}

    $name = ([string]$D.name).Trim()
    if ($name.Length -lt 2) { $e['name'] = 'Name must be at least 2 characters.' }

    $email = ([string]$D.email).Trim()
    if ($email -notmatch '^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$') { $e['email'] = 'Enter a valid email address.' }

    $phone = ([string]$D.phone).Trim()
    if ($phone -notmatch '^[0-9+\-\s()]{7,20}$') { $e['phone'] = 'Enter a valid phone number.' }

    if (-not ([string]$D.department).Trim())  { $e['department']  = 'Please select a department.' }
    if (-not ([string]$D.designation).Trim()) { $e['designation'] = 'Designation is required.' }

    $jd = ([string]$D.joining_date).Trim()
    if (-not $jd) {
        $e['joining_date'] = 'Joining date is required.'
    }
    elseif ($jd -notmatch '^\d{4}-\d{2}-\d{2}$') {
        $e['joining_date'] = 'Joining date must be YYYY-MM-DD.'
    }

    $sal = ([string]$D.salary).Trim()
    if ($sal) {
        $num = 0.0
        if (-not [double]::TryParse($sal, [ref]$num)) { $e['salary'] = 'Salary must be a number.' }
        elseif ($num -lt 0)                           { $e['salary'] = 'Salary cannot be negative.' }
    }
    return $e
}

# ===========================================================================
# 4. PORT AUTO-DETECTION
# ===========================================================================
function Test-PortFree {
    param([int]$P)
    try {
        $l = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, $P)
        $l.Start(); $l.Stop()
        return $true
    } catch { return $false }
}

function Get-FreePort {
    param([int]$Requested)

    if ($Requested -gt 0) {
        if (Test-PortFree $Requested) { return $Requested }
        Write-Warn2 "Port $Requested is busy - auto-detecting another one..."
    }

    foreach ($p in 8080, 8000, 5000, 3000, 8888, 5050, 7070, 9090, 4200, 8081) {
        if (Test-PortFree $p) { return $p }
    }

    # let the OS hand us any free ephemeral port
    $l = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, 0)
    $l.Start()
    $p = $l.LocalEndpoint.Port
    $l.Stop()
    return $p
}

# ===========================================================================
# 5. HTTP HELPERS
# ===========================================================================
function Send-Response {
    param($Context, [string]$Body, [string]$ContentType = 'application/json', [int]$Status = 200)
    $res    = $Context.Response
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $res.StatusCode      = $Status
    $res.ContentType     = "$ContentType; charset=utf-8"
    $res.ContentLength64 = $buffer.Length
    $res.Headers.Add('Cache-Control', 'no-store')
    $res.OutputStream.Write($buffer, 0, $buffer.Length)
    $res.OutputStream.Close()
}

function Send-Json {
    param($Context, $Object, [int]$Status = 200)
    Send-Response -Context $Context -Body ($Object | ConvertTo-Json -Depth 6 -Compress) -Status $Status
}

function ConvertTo-Hashtable {
    param($Obj)
    $h = @{}
    if ($null -ne $Obj) {
        foreach ($p in $Obj.PSObject.Properties) { $h[$p.Name] = $p.Value }
    }
    return $h
}

# ===========================================================================
# 6. BOOT
# ===========================================================================
Write-Head 'Employee Form -> Excel   (PowerShell edition)'

Initialize-Dependencies

Write-Head 'Preparing Excel file (permanent record store)'
$script:XlsxPath = Resolve-ExcelPath -Preferred $ExcelPath
Initialize-Workbook

# Keep a rolling safety copy of the master file each time the server starts.
if (Test-Path -LiteralPath $script:XlsxPath) {
    try {
        $bakDir = Join-Path (Split-Path -Parent $script:XlsxPath) 'backups'
        if (-not (Test-Path -LiteralPath $bakDir)) { New-Item -ItemType Directory -Path $bakDir -Force | Out-Null }
        $bak = Join-Path $bakDir ("form-{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Copy-Item -LiteralPath $script:XlsxPath -Destination $bak -Force
        # keep only the 10 most recent backups
        Get-ChildItem -LiteralPath $bakDir -Filter 'form-*.xlsx' |
            Sort-Object LastWriteTime -Descending | Select-Object -Skip 10 |
            Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Ok "Startup backup saved to $bak"
    } catch { Write-Warn2 "Could not create startup backup: $($_.Exception.Message)" }
}

Write-Head 'Starting web server'
$chosenPort = Get-FreePort -Requested $Port
Write-Ok "Auto-detected free port: $chosenPort"

$listener = [System.Net.HttpListener]::new()
$prefixes = @("http://+:$chosenPort/")     # all interfaces (works for the live preview)
foreach ($p in $prefixes) { $listener.Prefixes.Add($p) }

try {
    $listener.Start()
}
catch {
    # Windows without admin rights cannot bind "+" -> fall back to localhost only
    Write-Warn2 "Could not bind '+' (needs admin on Windows). Retrying on localhost..."
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("http://localhost:$chosenPort/")
    $listener.Start()
}

$url = "http://localhost:$chosenPort/"
Write-Host ''
Write-Ok  "Server running at  $url"
Write-Host "      Excel file   $script:XlsxPath" -ForegroundColor Gray
Write-Host "      Press Ctrl+C to stop"          -ForegroundColor DarkGray
Write-Host ''

if (-not $NoBrowser) {
    try {
        if ($IsWindows -or $env:OS -eq 'Windows_NT') { Start-Process $url | Out-Null }
        elseif ($IsMacOS)                            { & open  $url 2>$null }
    } catch { }
}

# ===========================================================================
# 7. REQUEST LOOP
# ===========================================================================
try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $req = $ctx.Request
        $verb = $req.HttpMethod
        $path = $req.Url.AbsolutePath.TrimEnd('/')
        if (-not $path) { $path = '/' }

        try {
            switch -Regex ("$verb $path") {

                # ---------- the form page --------------------------------
                '^GET /$' {
                    $file = Join-Path $script:Root 'static/index.html'
                    if (Test-Path -LiteralPath $file) {
                        Send-Response $ctx (Get-Content -LiteralPath $file -Raw) 'text/html'
                    } else {
                        Send-Response $ctx 'static/index.html not found' 'text/plain' 404
                    }
                    break
                }

                # ---------- which file are we writing to? ----------------
                '^GET /api/info$' {
                    Initialize-Workbook -Quiet     # self-heal if the file was deleted/moved
                    Send-Json $ctx @{ ok = $true; file = $script:XlsxPath; configured = $ExcelPath; port = $chosenPort }
                    break
                }

                # ---------- read all rows back out of the xlsx -----------
                '^GET /api/employees$' {
                    Initialize-Workbook -Quiet     # self-heal if the file was deleted/moved
                    $rows = Get-Records          # already an array-of-arrays
                    if ($null -eq $rows) { $rows = @() }
                    Send-Json $ctx @{ ok = $true; headers = $script:Headers; rows = $rows; file = $script:XlsxPath }
                    break
                }

                # ---------- append a row ---------------------------------
                '^POST /api/employees$' {
                    $reader = [System.IO.StreamReader]::new($req.InputStream, [System.Text.Encoding]::UTF8)
                    $raw    = $reader.ReadToEnd(); $reader.Close()

                    $data = @{}
                    try   { $data = ConvertTo-Hashtable ($raw | ConvertFrom-Json) }
                    catch { Send-Json $ctx @{ ok = $false; message = 'Invalid JSON body.' } 400; break }

                    $errors = Test-Record $data
                    if ($errors.Count -gt 0) {
                        Send-Json $ctx @{ ok = $false; message = 'Please fix the highlighted fields.'; errors = $errors } 400
                        break
                    }

                    $result = Add-Record -Data $data
                    if (-not $result.ok) {
                        Send-Json $ctx @{
                            ok      = $false
                            message = 'This email is already saved in the Excel file.'
                            errors  = @{ email = 'Duplicate email.' }
                        } 409
                        break
                    }

                    Write-Host ("  +  row #{0} saved  ({1})" -f $result.id, $data.name) -ForegroundColor Green
                    Send-Json $ctx @{
                        ok            = $true
                        message       = "Record #$($result.id) saved to Excel."
                        file          = $script:XlsxPath
                        total_records = $result.total
                    }
                    break
                }

                # ---------- download the workbook ------------------------
                '^GET /api/download$' {
                    if (-not (Test-Path -LiteralPath $script:XlsxPath)) { Initialize-Workbook }
                    $bytes = [System.IO.File]::ReadAllBytes($script:XlsxPath)
                    $res   = $ctx.Response
                    $res.StatusCode      = 200
                    $res.ContentType     = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    $res.ContentLength64 = $bytes.Length
                    $res.Headers.Add('Content-Disposition', 'attachment; filename="form.xlsx"')
                    $res.OutputStream.Write($bytes, 0, $bytes.Length)
                    $res.OutputStream.Close()
                    break
                }

                default {
                    Send-Json $ctx @{ ok = $false; message = "No route for $verb $path" } 404
                }
            }
        }
        catch {
            Write-Err $_.Exception.Message
            try { Send-Json $ctx @{ ok = $false; message = "Server error: $($_.Exception.Message)" } 500 } catch { }
        }
    }
}
finally {
    if ($listener) { $listener.Stop(); $listener.Close() }
    Write-Host ''
    Write-Host '  Server stopped.' -ForegroundColor Yellow
}

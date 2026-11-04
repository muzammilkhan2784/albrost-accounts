# Albrost Accounts - small local server.
# It shows the app in the browser and saves entries into the Excel file
# that sits next to the "app" folder. Nothing goes on the internet.

$ErrorActionPreference = 'Stop'
$AppDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root      = Split-Path -Parent $AppDir
$ExcelName = 'Al_Brost_Finance_Supplier_Balances.xlsx'
$ExcelPath = Join-Path $Root $ExcelName
$BackupDir = Join-Path $Root 'Backups'
$Port      = 8765
$Url       = "http://localhost:$Port/"
$IdleMinutes = 30

function Open-Ui {
    $candidates = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )
    $browser = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if ($browser) { Start-Process -FilePath $browser -ArgumentList "--app=$Url", "--window-size=1350,900" }
    else { Start-Process $Url }
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Url)
try { $listener.Start() }
catch {
    # Already running (or port busy): just open the window again.
    Open-Ui
    exit
}
Open-Ui

$static = @{
    '/'              = @('index.html', 'text/html; charset=utf-8')
    '/index.html'    = @('index.html', 'text/html; charset=utf-8')
    '/exceljs.min.js'= @('exceljs.min.js', 'application/javascript')
    '/chart.umd.js'  = @('chart.umd.js', 'application/javascript')
    '/logo.jpg'      = @('logo.jpg', 'image/jpeg')
}

function Send-Bytes($ctx, [int]$code, [byte[]]$bytes, [string]$type) {
    $res = $ctx.Response
    $res.StatusCode = $code
    $res.ContentType = $type
    $res.Headers.Add('Cache-Control', 'no-store')
    $res.ContentLength64 = $bytes.Length
    if ($bytes.Length -gt 0) { $res.OutputStream.Write($bytes, 0, $bytes.Length) }
    $res.OutputStream.Close()
}
function Send-Json($ctx, [int]$code, $obj) {
    $json = $obj | ConvertTo-Json -Compress
    Send-Bytes $ctx $code ([Text.Encoding]::UTF8.GetBytes($json)) 'application/json; charset=utf-8'
}

function Read-ExcelBytes {
    $fs = [IO.File]::Open($ExcelPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $ms = New-Object IO.MemoryStream
        $fs.CopyTo($ms)
        return ,$ms.ToArray()
    } finally { $fs.Close() }
}

function Save-Excel([byte[]]$bytes) {
    if ($bytes.Length -lt 1000) { throw 'empty' }
    # Is the file open in Excel? Then we cannot write.
    if (Test-Path $ExcelPath) {
        try {
            $probe = [IO.File]::Open($ExcelPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            $probe.Close()
        } catch { return 'locked' }
        if (-not (Test-Path $BackupDir)) { New-Item -ItemType Directory -Path $BackupDir | Out-Null }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        Copy-Item $ExcelPath (Join-Path $BackupDir "Backup_$stamp.xlsx") -Force
        # keep the newest 60 backups
        Get-ChildItem $BackupDir -Filter 'Backup_*.xlsx' | Sort-Object Name -Descending |
            Select-Object -Skip 60 | Remove-Item -Force -ErrorAction SilentlyContinue
    }
    $tmp = Join-Path $Root '~saving.tmp'
    [IO.File]::WriteAllBytes($tmp, $bytes)
    try { [IO.File]::Copy($tmp, $ExcelPath, $true) }
    catch { Remove-Item $tmp -Force -ErrorAction SilentlyContinue; return 'locked' }
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    return 'ok'
}

function Handle($ctx) {
    $req  = $ctx.Request
    $path = $req.Url.AbsolutePath
    $hostName = ($req.Headers['Host'] -split ':')[0]
    if ($hostName -ne 'localhost' -and $hostName -ne '127.0.0.1') { Send-Json $ctx 403 @{ error = 'forbidden' }; return }

    if ($req.HttpMethod -eq 'GET' -and $static.ContainsKey($path)) {
        $f = Join-Path $AppDir $static[$path][0]
        Send-Bytes $ctx 200 ([IO.File]::ReadAllBytes($f)) $static[$path][1]
        return
    }
    switch ($path) {
        '/api/ping' { Send-Json $ctx 200 @{ ok = $true }; return }
        '/api/file' {
            if (-not (Test-Path $ExcelPath)) { Send-Json $ctx 404 @{ error = 'missing'; file = $ExcelPath }; return }
            Send-Bytes $ctx 200 (Read-ExcelBytes) 'application/octet-stream'
            return
        }
        '/api/save' {
            if ($req.HttpMethod -ne 'POST') { Send-Json $ctx 405 @{ error = 'method' }; return }
            $ms = New-Object IO.MemoryStream
            $req.InputStream.CopyTo($ms)
            $r = Save-Excel ($ms.ToArray())
            if ($r -eq 'ok') { Send-Json $ctx 200 @{ ok = $true } } else { Send-Json $ctx 423 @{ error = $r } }
            return
        }
        '/api/open-excel' {
            if ($req.HttpMethod -ne 'POST') { Send-Json $ctx 405 @{ error = 'method' }; return }
            Start-Process $ExcelPath
            Send-Json $ctx 200 @{ ok = $true }; return
        }
        '/api/open-backups' {
            if ($req.HttpMethod -ne 'POST') { Send-Json $ctx 405 @{ error = 'method' }; return }
            if (-not (Test-Path $BackupDir)) { New-Item -ItemType Directory -Path $BackupDir | Out-Null }
            Start-Process explorer.exe $BackupDir
            Send-Json $ctx 200 @{ ok = $true }; return
        }
        default { Send-Json $ctx 404 @{ error = 'not found' } }
    }
}

$lastSeen = [DateTime]::Now
while ($listener.IsListening) {
    $async = $listener.BeginGetContext($null, $null)
    while (-not $async.AsyncWaitHandle.WaitOne(5000)) {
        if (([DateTime]::Now - $lastSeen).TotalMinutes -gt $IdleMinutes) { $listener.Stop(); exit }
    }
    $ctx = $listener.EndGetContext($async)
    $lastSeen = [DateTime]::Now
    try { Handle $ctx }
    catch {
        try { Send-Json $ctx 500 @{ error = 'server'; detail = $_.Exception.Message } } catch {}
    }
}

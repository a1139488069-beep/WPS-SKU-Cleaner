param([switch]$NoBrowser)
$ErrorActionPreference = 'Stop'
$skuNodeCommand = Get-Command node.exe -ErrorAction SilentlyContinue
$skuNodePath = if ($skuNodeCommand) { $skuNodeCommand.Source } else {
    Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
}
if (-not (Test-Path -LiteralPath $skuNodePath)) {
    throw 'Node.js was not found. Install Node.js 18 or later, then run this script again.'
}
$skuUrl = 'http://127.0.0.1:39871'
$skuRunning = $false
try {
    $skuHealth = Invoke-RestMethod -Uri "$skuUrl/health" -TimeoutSec 2
    if ($skuHealth.app -ne 'WPS-SKU-Cleaner') { throw 'Port 39871 belongs to another application.' }
    $skuRunning = $true
} catch {
    if ($_.Exception.Message -like '*belongs to another*') { throw }
}
if (-not $skuRunning) {
    $skuStateDir = Join-Path $env:LOCALAPPDATA 'WpsSkuCleaner'
    New-Item -ItemType Directory -Path $skuStateDir -Force | Out-Null
    $skuServerPath = Join-Path $PSScriptRoot 'server.cjs'
    $skuProc = Start-Process -FilePath $skuNodePath -ArgumentList @('"' + $skuServerPath + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $skuStateDir 'server.log') -RedirectStandardError (Join-Path $skuStateDir 'server-error.log')
    $skuProc.Id | Set-Content -LiteralPath (Join-Path $skuStateDir 'server.pid')
    for ($skuAttempt = 0; $skuAttempt -lt 20; $skuAttempt++) {
        Start-Sleep -Milliseconds 200
        try {
            $skuHealth = Invoke-RestMethod -Uri "$skuUrl/health" -TimeoutSec 1
            if ($skuHealth.app -eq 'WPS-SKU-Cleaner') { $skuRunning = $true; break }
        } catch {}
    }
    if (-not $skuRunning) { throw "Local server failed. Check $skuStateDir\server-error.log" }
}
if (-not $NoBrowser) { Start-Process "$skuUrl/install.html" }
Write-Host 'Local server is running. Use the browser page to install/uninstall the WPS add-in.'

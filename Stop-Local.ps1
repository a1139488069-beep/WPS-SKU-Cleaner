$ErrorActionPreference = 'Stop'
$skuPidFile = Join-Path $env:LOCALAPPDATA 'WpsSkuCleaner\server.pid'
if (-not (Test-Path -LiteralPath $skuPidFile)) { Write-Host 'No local server PID recorded.'; exit }
$skuProcessId = [int](Get-Content -LiteralPath $skuPidFile)
$skuProcessInfo = Get-CimInstance Win32_Process -Filter "ProcessId=$skuProcessId"
$skuExpectedPath = Join-Path $PSScriptRoot 'server.cjs'
if ($skuProcessInfo -and $skuProcessInfo.Name -eq 'node.exe' -and $skuProcessInfo.CommandLine.Contains($skuExpectedPath)) {
    Stop-Process -Id $skuProcessId
    Remove-Item -LiteralPath $skuPidFile
    Write-Host 'Local SKU server stopped. Use the install page to uninstall before stopping permanently.'
} elseif ($skuProcessInfo) {
    throw 'Recorded PID belongs to a different process; nothing was stopped.'
} else {
    Remove-Item -LiteralPath $skuPidFile
    Write-Host 'Local SKU server already stopped.'
}

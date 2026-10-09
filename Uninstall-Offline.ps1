param([string]$Destination)
$ErrorActionPreference = 'Stop'
if (-not $Destination) { $Destination = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons' }
$skuManifest = Join-Path $Destination 'publish.xml'
if (-not (Test-Path -LiteralPath $skuManifest)) { Write-Host 'No add-in config found.'; exit }
$skuXml = New-Object System.Xml.XmlDocument
$skuXml.PreserveWhitespace = $true
$skuXml.Load($skuManifest)
if ($skuXml.DocumentElement.Name -ne 'jsplugins') { throw 'Unexpected config; nothing was changed.' }
$skuEntries = @($skuXml.DocumentElement.SelectNodes('*') | Where-Object { $_.GetAttribute('name') -eq 'WPS-SKU-Cleaner' })
if (-not $skuEntries.Count) { Write-Host 'SKU add-in is not registered.'; exit }
Copy-Item -LiteralPath $skuManifest -Destination ($skuManifest + '.sku-uninstall-backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
foreach ($skuEntry in $skuEntries) { [void]$skuXml.DocumentElement.RemoveChild($skuEntry) }
$skuXml.Save($skuManifest)
Write-Host 'SKU add-in registration removed. Save files and restart WPS. Code files and other add-ins are preserved.'

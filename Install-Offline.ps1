param([string]$Destination)
$ErrorActionPreference = 'Stop'
$skuVersion = '1.0.5'
$skuName = 'WPS-SKU-Cleaner'
$skuFolder = $skuName + '_' + $skuVersion
if (-not $Destination) { $Destination = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons' }
$skuDestinationRoot = [IO.Path]::GetFullPath($Destination)
$skuSource = Join-Path $PSScriptRoot ('offline\' + $skuFolder)
foreach ($skuRequired in @('index.html','main.js','ribbon.xml')) {
    if (-not (Test-Path -LiteralPath (Join-Path $skuSource $skuRequired))) { throw "Missing offline file: $skuRequired. Extract the whole project first." }
}
New-Item -ItemType Directory -Path $skuDestinationRoot -Force | Out-Null
$skuTarget = Join-Path $skuDestinationRoot $skuFolder
New-Item -ItemType Directory -Path $skuTarget -Force | Out-Null
foreach ($skuRequired in @('index.html','main.js','ribbon.xml')) {
    Copy-Item -LiteralPath (Join-Path $skuSource $skuRequired) -Destination (Join-Path $skuTarget $skuRequired) -Force
}
$skuManifest = Join-Path $skuDestinationRoot 'publish.xml'
$skuXml = New-Object System.Xml.XmlDocument
$skuXml.PreserveWhitespace = $true
if (Test-Path -LiteralPath $skuManifest) {
    $skuXml.Load($skuManifest)
    if ($skuXml.DocumentElement.Name -ne 'jsplugins') { throw 'Unexpected publish.xml root; config was not changed.' }
    Copy-Item -LiteralPath $skuManifest -Destination ($skuManifest + '.sku-backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
} else { $skuXml.LoadXml('<jsplugins></jsplugins>') }
$skuOldEntries = @($skuXml.DocumentElement.SelectNodes('*') | Where-Object { $_.GetAttribute('name') -eq $skuName })
foreach ($skuOldEntry in $skuOldEntries) { [void]$skuXml.DocumentElement.RemoveChild($skuOldEntry) }
$skuEntry = $skuXml.CreateElement('jsplugin')
foreach ($skuAttribute in @{name=$skuName;type='et';url=$skuFolder;version=$skuVersion;enable='enable_dev';install='null'}.GetEnumerator()) {
    $skuEntry.SetAttribute($skuAttribute.Key,$skuAttribute.Value)
}
[void]$skuXml.DocumentElement.AppendChild($skuEntry)
$skuSettings = New-Object System.Xml.XmlWriterSettings
$skuSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
$skuSettings.Indent = $true
$skuWriter = [System.Xml.XmlWriter]::Create($skuManifest,$skuSettings)
try { $skuXml.Save($skuWriter) } finally { $skuWriter.Dispose() }
Write-Host "Offline SKU add-in $skuVersion installed at $skuTarget"
Write-Host 'Save all WPS files, fully exit WPS, then reopen it. Approve the normal WPS add-in prompt if shown.'
Write-Host 'Look for the SKU tools tab. No Node.js or local HTTP service is required for normal use.'

<#
.SYNOPSIS
Build a dated KEV report while retaining saved Weekly report edits and device history.
.EXAMPLE
.\Refresh-KEVReport.ps1
Open the file-picker interface. The CMD launcher does this when double-clicked.
.EXAMPLE
.\Refresh-KEVReport.ps1 -ExportPath '..\DevicesByKEVwUsers&Incidents Nov 9 2026.xlsx' -ReportDate '2026-11-09' -OutputPath '.\KEV-Report-2026-11-09.xlsx'
.EXAMPLE
.\Refresh-KEVReport.ps1 -ExportPath 'C:\Reports\Defender-export.csv' -ReportDate '2026-11-16' -TrackingWorkbook '.\KEV-Report-2026-11-09.xlsx' -OutputPath '.\KEV-Report-2026-11-16.xlsx'
.NOTES
ExportPath accepts native Defender CSV exports or Excel (.xlsx) exports.
Save and close the previous report before running. Saved report inputs are archived; outputs are never overwritten.
This local pilot uses the Codex bundled runtime on this computer; it does not contact Defender or ServiceNow.
#>
[CmdletBinding()]
param(
    [string]$ExportPath,
    [ValidatePattern('^\d{4}-\d{2}-\d{2}$')][string]$ReportDate,
    [string]$TrackingWorkbook,
    [string]$OutputPath,
    [switch]$Gui,
    [switch]$UpgradeReport
)
$ErrorActionPreference = 'Stop'
$pilotRoot = Split-Path -Parent $PSScriptRoot
if ($Gui -or $PSBoundParameters.Count -eq 0) {
    . (Join-Path $PSScriptRoot 'ReportGui.ps1')
    $guiOptions = @{ PilotRoot = $pilotRoot }
    foreach ($name in @('ExportPath', 'ReportDate', 'TrackingWorkbook', 'OutputPath')) {
        if ($PSBoundParameters.ContainsKey($name)) { $guiOptions[$name] = $PSBoundParameters[$name] }
    }
    Show-KEVReportDialog @guiOptions
    return
}
$requiredNames = if ($UpgradeReport) { @('TrackingWorkbook', 'OutputPath') } else { @('ExportPath', 'ReportDate', 'OutputPath') }
foreach ($name in $requiredNames) {
    if ([string]::IsNullOrWhiteSpace((Get-Variable -Name $name -ValueOnly))) {
        throw "Supply -ExportPath, -ReportDate and -OutputPath, or run without arguments to open the GUI. Missing: -$name."
    }
}
if (!$UpgradeReport) {
    $null = [datetime]::ParseExact($ReportDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $sourcePath = (Resolve-Path -LiteralPath $ExportPath).Path
    if (!(Test-Path -LiteralPath $sourcePath -PathType Leaf) -or [IO.Path]::GetExtension($sourcePath) -notin @('.csv', '.xlsx')) {
        throw 'ExportPath must be an existing .csv or .xlsx Defender export.'
    }
}
$destinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($destinationPath) -ne '.xlsx') { throw 'OutputPath must end in .xlsx.' }
if (Test-Path -LiteralPath $destinationPath) { throw 'Output already exists. Choose a new filename to preserve the previous report.' }
$runtimeRoot = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies'
$nodeExecutable = Join-Path $runtimeRoot 'node\bin\node.exe'
$pythonExecutable = Join-Path $runtimeRoot 'python\python.exe'
$packagesPath = Join-Path $runtimeRoot 'node\node_modules'
if (!(Test-Path -LiteralPath $nodeExecutable) -or !(Test-Path -LiteralPath $pythonExecutable)) {
    throw 'Codex bundled Node.js/Python runtime was not found. Open this project in Codex and check workspace dependencies.'
}
$moduleLink = Join-Path $pilotRoot 'node_modules'
$moduleInfo = Get-Item -LiteralPath $moduleLink -Force -ErrorAction SilentlyContinue
if ($moduleInfo -and $moduleInfo.LinkType -eq 'Junction' -and @($moduleInfo.Target)[0] -ne $packagesPath) {
    # Repoint only the utility's dependency junction, never a real directory.
    [IO.Directory]::Delete($moduleLink)
    $moduleInfo = $null
}
if (!$moduleInfo) {
    New-Item -ItemType Junction -Path $moduleLink -Target $packagesPath | Out-Null
}
$runDirectory = Join-Path $pilotRoot ('validation\run-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $runDirectory | Out-Null
$jsonPath = Join-Path $runDirectory 'prepared.json'
$prepareArgs = @((Join-Path $PSScriptRoot 'prepare_report.py'), '--json', $jsonPath)
$previousPath = $null
if (!$UpgradeReport) { $prepareArgs += @('--export', $sourcePath, '--date', $ReportDate) }
if ($TrackingWorkbook) {
    $previousPath = (Resolve-Path -LiteralPath $TrackingWorkbook).Path
    if (!(Test-Path -LiteralPath $previousPath -PathType Leaf) -or [IO.Path]::GetExtension($previousPath) -ne '.xlsx') { throw 'TrackingWorkbook must be an existing .xlsx report.' }
    if ($UpgradeReport) { $prepareArgs += @('--existing-report', $previousPath) }
    else { $prepareArgs += @('--previous', $previousPath) }
} else {
    # Only the first build may bootstrap from the original annotated workbooks.
    if (Get-ChildItem -LiteralPath $pilotRoot -Filter 'KEV-Report-*.xlsx' -File) {
        throw 'A pilot report already exists. Supply -TrackingWorkbook so its edits are retained.'
    }
    $projectRoot = Split-Path -Parent $pilotRoot
    $nov2 = Join-Path $projectRoot 'AdvancedHuntingResults-AT - Date Passed CISA Known Exploited Vulnerabilities_20261102.xlsx'
    $nov9 = Join-Path $projectRoot 'AdvancedHuntingResults-AT - Date Passed CISA Known Exploited Vulnerabilities_20261109.xlsx'
    # Preserve the local one-time migration when its inputs are available.
    # A fresh checkout can start directly from its first Defender export.
    if ((Test-Path -LiteralPath $nov2 -PathType Leaf) -and (Test-Path -LiteralPath $nov9 -PathType Leaf)) {
        $prepareArgs += @('--legacy', ('2026-11-02|' + $nov2), '--legacy', ('2026-11-09|' + $nov9))
    }
}
& $pythonExecutable -B @prepareArgs
if ($LASTEXITCODE -ne 0) { throw 'Source validation or tracking migration failed. No workbook was generated.' }
if ($previousPath) {
    $archiveDirectory = Join-Path $pilotRoot 'archive'
    New-Item -ItemType Directory -Force -Path $archiveDirectory | Out-Null
    $archiveName = [datetime]::UtcNow.ToString('yyyyMMddTHHmmssZ') + '_' + [guid]::NewGuid().ToString('N') + '_' + [IO.Path]::GetFileName($previousPath)
    $archivePath = Join-Path $archiveDirectory $archiveName
    Copy-Item -LiteralPath $previousPath -Destination $archivePath -ErrorAction Stop
    Write-Output ('Archived saved input: ' + $archivePath)
}
& $nodeExecutable (Join-Path $PSScriptRoot 'build_report.mjs') $jsonPath $destinationPath
if ($LASTEXITCODE -ne 0) { throw 'Workbook generation failed. Inspect the validation output before retrying.' }
& $pythonExecutable -B (Join-Path $PSScriptRoot 'finish_workbook.py') $destinationPath
if ($LASTEXITCODE -ne 0) { throw 'The report was generated but the register could not be hidden; inspect the output before using it.' }
Write-Output ('Created ' + $destinationPath)

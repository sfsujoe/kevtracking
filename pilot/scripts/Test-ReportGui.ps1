$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReportGui.ps1')
$pilotRoot = Split-Path -Parent $PSScriptRoot

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (!$Condition) { throw $Message }
}
function Assert-Throws {
    param([scriptblock]$Action, [string]$Expected)
    try { & $Action | Out-Null } catch {
        if ($_.Exception.Message -like ('*' + $Expected + '*')) { return }
        throw
    }
    throw "Expected validation failure: $Expected"
}

# Use a harmless worker fixture to test the actual hidden-process/timer path.
# Copying the existing workbook tests file arguments without creating report data.
$testRoot = Join-Path $pilotRoot ('validation\gui test & apostrophe''s ' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$previousPath = Join-Path $testRoot 'KEV-Report-fixture.xlsx'
# The form checks paths, not workbook contents; the worker copies this marker.
Set-Content -LiteralPath $previousPath -Value 'Synthetic GUI path-validation fixture; not a workbook.' -Encoding UTF8
$worker = @'
param([string]$ExportPath, [string]$ReportDate, [string]$OutputPath, [string]$TrackingWorkbook)
$ErrorActionPreference = 'Stop'
if (!(Test-Path -LiteralPath $ExportPath)) { throw 'Export argument did not survive quoting.' }
Copy-Item -LiteralPath $TrackingWorkbook -Destination $OutputPath
Write-Output ('Created ' + $OutputPath)
'@
Set-Content -LiteralPath (Join-Path $testRoot 'Refresh-KEVReport.ps1') -Value $worker -Encoding UTF8
$csvPath = Join-Path $testRoot 'MDE export & users.csv'
[pscustomobject]@{ DeviceId = 'csv-test'; DeviceName = 'CSV test'; KEVCount = 1; OverdueKEVs = 1; KEVs = '["CVE-2026-0001"]' } | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
$exportPath = $csvPath
$requestOptions = @{ PilotRoot = $testRoot; ExportPath = $exportPath; TrackingWorkbook = $previousPath; ReportDate = [datetime]'2026-11-16'; OutputDirectory = $testRoot; OutputFileName = 'next report.xlsx' }
$request = Resolve-KEVReportRequest @requestOptions
Assert-True ($request.ReportDate -eq '2026-11-16') 'Report date must be invariant.'
Assert-True ($request.ExportPath -eq $exportPath) 'Export path changed unexpectedly.'
$csvOptions = $requestOptions.Clone(); $csvOptions.ExportPath = $csvPath
$csvRequest = Resolve-KEVReportRequest @csvOptions
Assert-True ($csvRequest.ExportPath -eq $csvPath) 'CSV exports must be accepted.'
$bad = $requestOptions.Clone(); $bad.TrackingWorkbook = $csvPath
Assert-Throws { Resolve-KEVReportRequest @bad } 'previous report must be an .xlsx'
$bad = $requestOptions.Clone(); $bad.ExportPath = ''
Assert-Throws { Resolve-KEVReportRequest @bad } 'existing Defender'
$bad = $requestOptions.Clone(); $bad.TrackingWorkbook = ''
Assert-Throws { Resolve-KEVReportRequest @bad } 'previous report'
$bad = $requestOptions.Clone(); $bad.OutputDirectory = Join-Path $testRoot 'missing'
Assert-Throws { Resolve-KEVReportRequest @bad } 'output folder'
$bad = $requestOptions.Clone(); $bad.OutputFileName = '..\other.xlsx'
Assert-Throws { Resolve-KEVReportRequest @bad } 'report filename'
$bad = $requestOptions.Clone(); $bad.OutputDirectory = $testRoot; $bad.OutputFileName = 'KEV-Report-fixture.xlsx'
Assert-Throws { Resolve-KEVReportRequest @bad } 'already exists'
$arguments = Get-KEVWorkerArguments -PilotRoot $testRoot -Request $request
Assert-True ($arguments.Contains('"' + $exportPath + '"')) 'The export argument must be quoted.'
Assert-True ($arguments.Contains('-NonInteractive')) 'The worker must not prompt for input.'
Write-Output 'File/directory validation and worker quoting passed.'

function Invoke-TestClick {
    param($Button)
    $method = $Button.GetType().GetMethod('OnClick', [Reflection.BindingFlags]'Instance,NonPublic')
    $method.Invoke($Button, @([EventArgs]::Empty)) | Out-Null
}
function Wait-TestWorker {
    param($Ui)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($Ui.State.Busy) {
        [Windows.Forms.Application]::DoEvents()
        if ($watch.Elapsed.TotalSeconds -gt 20) { throw ('Worker timeout: ' + $Ui.Status.Text) }
        Start-Sleep -Milliseconds 50
    }
}
# Exercise the real Browse click handlers without showing native dialogs.
# Capture their state, then simulate Cancel/OK or a native dialog failure.
$script:pickerState = @{ Calls = (New-Object Collections.ArrayList); Result = 'Cancel'; Selection = ''; Fail = $false }
function Show-KEVPickerDialog {
    param($Dialog, $Owner)
    if ($script:pickerState.Fail) { throw 'Picker fixture failure' }
    $isFolder = $Dialog -is [Windows.Forms.FolderBrowserDialog]
    [void]$script:pickerState.Calls.Add([pscustomobject]@{
        Kind = if ($isFolder) { 'Folder' } else { 'File' }
        Title = if ($isFolder) { $Dialog.Description } else { $Dialog.Title }
        InitialPath = if ($isFolder) { $Dialog.SelectedPath } else { $Dialog.FileName }
        Filter = if ($isFolder) { '' } else { $Dialog.Filter }
    })
    if ($script:pickerState.Result -eq 'OK') {
        if ($isFolder) { $Dialog.SelectedPath = $script:pickerState.Selection }
        else { $Dialog.FileName = $script:pickerState.Selection }
    }
    return [Windows.Forms.DialogResult]::$($script:pickerState.Result)
}
$ui = New-KEVReportForm -PilotRoot $testRoot -ExportPath $exportPath -TrackingWorkbook $previousPath -ReportDate '2026-11-16'
try {
    Assert-True (@($ui).Count -eq 1) 'The form factory must return only its UI object.'
    Assert-True ($ui.BrowseButtons.Count -eq 3) 'Expected two file pickers and a folder picker.'
    foreach ($key in @('Export', 'Tracking', 'Folder')) {
        foreach ($initialValue in @('', '   ', (Join-Path $testRoot 'does not exist'))) {
            $ui.Inputs[$key].Text = $initialValue
            $before = $script:pickerState.Calls.Count
            Invoke-TestClick $ui.BrowseButtons[$key]
            Assert-True ($script:pickerState.Calls.Count -eq $before + 1) "Browse for $key must open the picker when its field is blank or invalid."
            Assert-True ($ui.Inputs[$key].Text -eq $initialValue) "Cancel must preserve the $key field."
        }
    }
    foreach ($spec in @(@('Export', $csvPath, 'Choose the Defender export'), @('Tracking', $previousPath, 'Choose the saved report with your latest edits and history'), @('Folder', $testRoot, 'Choose where to save the new KEV report'))) {
        $key = $spec[0]
        $ui.Inputs[$key].Text = $spec[1]
        Invoke-TestClick $ui.BrowseButtons[$key]
        $captured = $script:pickerState.Calls[$script:pickerState.Calls.Count - 1]
        Assert-True ($captured.InitialPath -eq $spec[1]) "Browse for $key must retain its existing selection."
        Assert-True ($captured.Title -eq $spec[2]) "Browse for $key opened the wrong picker."
        if ($key -eq 'Export') { Assert-True ($captured.Filter.Contains('*.csv;*.xlsx')) 'The Defender picker must show CSV and XLSX files.' }
        if ($key -eq 'Tracking') { Assert-True (!$captured.Filter.Contains('*.csv')) 'The tracking picker must request an XLSX report.' }
        $ui.Inputs[$key].Text = ''
        $script:pickerState.Result = 'OK'
        $script:pickerState.Selection = $spec[1]
        Invoke-TestClick $ui.BrowseButtons[$key]
        Assert-True ($ui.Inputs[$key].Text -eq $spec[1]) "Browse for $key did not apply the selection."
        $script:pickerState.Result = 'Cancel'
    }
    $script:pickerState.Fail = $true
    Invoke-TestClick $ui.BrowseButtons.Export
    Assert-True ($ui.Status.Text.Contains('Could not open file picker: Picker fixture failure')) 'Native file-picker errors must be handled in the form.'
    Invoke-TestClick $ui.BrowseButtons.Folder
    Assert-True ($ui.Status.Text.Contains('Could not open folder picker: Picker fixture failure')) 'Native folder-picker errors must be handled in the form.'
    $script:pickerState.Fail = $false
    $ui.Status.Text = 'Ready. Select the export, then choose Create report.'
    Assert-True ((Read-KEVWorkerLog '') -eq '') 'A blank log path must not throw.'
    Write-Output 'Browse handlers passed blank/whitespace/missing paths, Cancel, OK, and picker-error checks.'
    Assert-True ($ui.Inputs.FileName.Text -eq 'KEV-Report-2026-11-16.xlsx') 'Expected a dated output filename.'
    $ui.DatePicker.Value = [datetime]'2026-11-17'
    Assert-True ($ui.Inputs.FileName.Text -eq 'KEV-Report-2026-11-17.xlsx') 'Changing the date must update the default filename.'
    $ui.Inputs.FileName.Text = 'custom report.xlsx'
    $ui.DatePicker.Value = [datetime]'2026-11-18'
    Assert-True ($ui.Inputs.FileName.Text -eq 'custom report.xlsx') 'Changing the date must preserve a custom filename.'
    # Create child handles without showing an interactive window, so the
    # rendered preview includes the form's controls rather than an empty frame.
    $createControls = [Windows.Forms.Control].GetMethod('CreateControl', [Reflection.BindingFlags]'Instance,NonPublic', $null, [type[]]@([bool]), $null)
    $createControls.Invoke($ui.Form, @($true)) | Out-Null
    $ui.Form.PerformLayout()
    foreach ($box in $ui.Inputs.Values) { Assert-True ($box.Width -gt 200) 'A file/folder input has no usable layout width.' }
    $bitmap = New-Object Drawing.Bitmap($ui.Form.Width, $ui.Form.Height)
    try {
        $ui.Form.DrawToBitmap($bitmap, (New-Object Drawing.Rectangle(0, 0, $bitmap.Width, $bitmap.Height)))
        $bitmap.Save((Join-Path $pilotRoot 'validation\gui-preview.png'), [Drawing.Imaging.ImageFormat]::Png)
    } finally { $bitmap.Dispose() }

    Invoke-TestClick $ui.CreateButton
    Assert-True $ui.State.Busy ('Expected a background worker. Status: ' + $ui.Status.Text)
    Assert-True (!$ui.CreateButton.Enabled) 'Duplicate starts must be disabled while busy.'
    Wait-TestWorker $ui
    Assert-True ($ui.Status.Text.StartsWith('Created:')) ('Worker did not succeed: ' + $ui.Status.Text + ' ' + $ui.Log.Text)
    Assert-True $ui.CreateButton.Enabled 'Controls must be reenabled after completion.'
    $copyPath = Join-Path $testRoot 'custom report.xlsx'
    Assert-True ((Get-FileHash -LiteralPath $copyPath).Hash -eq (Get-FileHash -LiteralPath $previousPath).Hash) 'Worker arguments or copy destination changed.'
    Write-Output 'Hidden worker, responsive timer, and success handling passed.'

    Set-Content -LiteralPath (Join-Path $testRoot 'Refresh-KEVReport.ps1') -Value 'Write-Error "Expected fixture failure"; exit 7' -Encoding UTF8
    $ui.Inputs.FileName.Text = 'failed report.xlsx'
    Invoke-TestClick $ui.CreateButton
    Assert-True $ui.State.Busy ('Expected failure worker to start: ' + $ui.Status.Text)
    Wait-TestWorker $ui
    Assert-True ($ui.Status.Text.StartsWith('Report creation failed')) 'Worker failure must be visible.'
    Assert-True ($ui.Log.Text.Contains('Expected fixture failure')) 'Error details must appear in the log.'
    $ui.Inputs.Export.Text = ''
    Invoke-TestClick $ui.CreateButton
    Assert-True (!$ui.State.Busy) 'Invalid input must not start a worker.'
    Assert-True ($ui.Status.Text.Contains('existing Defender')) 'Validation must be visible in the form.'
    Write-Output 'Worker failure, log display, and inline validation passed.'
} finally {
    $ui.Timer.Stop()
    $ui.Timer.Dispose()
    $ui.Form.Dispose()
}
Write-Output 'GUI checks passed.'

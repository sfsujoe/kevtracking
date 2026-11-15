# Windows Forms front end. The report generator runs in a separate hidden process
# so file dialogs, window repainting and progress remain responsive.

function Get-KEVLatestReport {
    param([string]$PilotRoot)
    $files = @(Get-ChildItem -LiteralPath $PilotRoot -Filter 'KEV-Report-*.xlsx' -File)
    $outputs = Join-Path (Split-Path -Parent $PilotRoot) 'outputs'
    if (Test-Path -LiteralPath $outputs -PathType Container) {
        $files += @(Get-ChildItem -LiteralPath $outputs -Filter 'KEV-Report-*.xlsx' -File -Recurse)
    }
    return $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

function Resolve-KEVReportRequest {
    param(
        [string]$PilotRoot,
        [string]$ExportPath,
        [string]$TrackingWorkbook,
        [datetime]$ReportDate,
        [string]$OutputDirectory,
        [string]$OutputFileName
    )
    if ([string]::IsNullOrWhiteSpace($ExportPath) -or !(Test-Path -LiteralPath $ExportPath -PathType Leaf)) {
        throw 'Choose an existing Defender CSV or Excel export.'
    }
    $exportFullPath = (Resolve-Path -LiteralPath $ExportPath).ProviderPath
    if ([IO.Path]::GetExtension($exportFullPath) -notin @('.csv', '.xlsx')) { throw 'The Defender export must be a .csv or .xlsx file.' }
    $previousFullPath = ''
    if (![string]::IsNullOrWhiteSpace($TrackingWorkbook)) {
        if (!(Test-Path -LiteralPath $TrackingWorkbook -PathType Leaf)) { throw 'Choose an existing previous report workbook.' }
        $previousFullPath = (Resolve-Path -LiteralPath $TrackingWorkbook).ProviderPath
        if ([IO.Path]::GetExtension($previousFullPath) -ne '.xlsx') { throw 'The previous report must be an .xlsx file.' }
    } elseif (Get-KEVLatestReport $PilotRoot) {
        throw 'Choose the previous report so its saved Weekly report edits and history are retained.'
    }
    if ([string]::IsNullOrWhiteSpace($OutputDirectory) -or !(Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        throw 'Choose an existing output folder.'
    }
    if ([string]::IsNullOrWhiteSpace($OutputFileName) -or
        $OutputFileName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
        [IO.Path]::GetExtension($OutputFileName) -ne '.xlsx') {
        throw 'Enter a report filename ending in .xlsx, without a directory path.'
    }
    $outputFullPath = Join-Path (Resolve-Path -LiteralPath $OutputDirectory).ProviderPath $OutputFileName
    if (Test-Path -LiteralPath $outputFullPath) {
        throw 'That report already exists. Choose a different filename to preserve it.'
    }
    return [pscustomobject]@{
        ExportPath = $exportFullPath
        TrackingWorkbook = $previousFullPath
        ReportDate = $ReportDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        OutputPath = $outputFullPath
    }
}

function ConvertTo-KEVProcessArgument {
    param([string]$Value)
    # File arguments are literal native arguments, never PowerShell expressions.
    if ($Value.IndexOfAny([char[]]@('"', "`r", "`n")) -ge 0) { throw 'Invalid quote or newline in a file argument.' }
    return '"' + $Value + '"'
}

function Get-KEVWorkerArguments {
    param([string]$PilotRoot, $Request)
    $arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
        (ConvertTo-KEVProcessArgument (Join-Path $PilotRoot 'Refresh-KEVReport.ps1')),
        '-ExportPath', (ConvertTo-KEVProcessArgument $Request.ExportPath),
        '-ReportDate', $Request.ReportDate,
        '-OutputPath', (ConvertTo-KEVProcessArgument $Request.OutputPath))
    if ($Request.TrackingWorkbook) {
        $arguments += @('-TrackingWorkbook', (ConvertTo-KEVProcessArgument $Request.TrackingWorkbook))
    }
    return $arguments -join ' '
}

function Read-KEVWorkerLog {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or !(Test-Path -LiteralPath $Path)) { return '' }
    $stream = $null
    $reader = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8, $true)
        return $reader.ReadToEnd()
    } catch [IO.IOException] {
        return ''
    } finally {
        if ($reader) { $reader.Dispose() } elseif ($stream) { $stream.Dispose() }
    }
}

function Show-KEVPickerDialog {
    param($Dialog, $Owner)
    return $Dialog.ShowDialog($Owner)
}

function New-KEVReportForm {
    param(
        [Parameter(Mandatory = $true)][string]$PilotRoot,
        [string]$ExportPath,
        [string]$ReportDate,
        [string]$TrackingWorkbook,
        [string]$OutputPath
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()
    $PilotRoot = [IO.Path]::GetFullPath($PilotRoot)
    $form = New-Object Windows.Forms.Form
    $form.Text = 'Weekly KEV Report'
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = New-Object Drawing.Size(900, 620)
    $form.MinimumSize = New-Object Drawing.Size(820, 620)
    $form.Font = New-Object Drawing.Font('Segoe UI', 10)
    $form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::Dpi

    $layout = New-Object Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = New-Object Windows.Forms.Padding(18)
    $layout.ColumnCount = 3
    $layout.RowCount = 11
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 175)))
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 100)))
    for ($i = 0; $i -lt 11; $i++) {
        if ($i -eq 8) { [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100))) }
        else { [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize'))) }
    }
    $form.Controls.Add($layout)

    $intro = New-Object Windows.Forms.Label
    $intro.Text = 'Choose the Defender export and your latest saved report. Weekly report edits and device history will carry forward.'
    $intro.AutoSize = $true
    $intro.MaximumSize = New-Object Drawing.Size(825, 0)
    $intro.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 16)
    $layout.Controls.Add($intro, 0, 0)
    $layout.SetColumnSpan($intro, 3)

    $inputs = @{}
    $browsers = @{}
    foreach ($spec in @(
        @(1, 'Export', 'Defender export', 'Browse...'),
        @(2, 'Tracking', 'Previous report', 'Browse...'),
        @(4, 'Folder', 'Output folder', 'Browse...'),
        @(5, 'FileName', 'Report filename', '')
    )) {
        $row = [int]$spec[0]
        $key = $spec[1]
        $label = New-Object Windows.Forms.Label
        $label.Text = $spec[2]
        $label.AutoSize = $true
        $label.Anchor = 'Left'
        $label.Margin = New-Object Windows.Forms.Padding(0, 9, 8, 9)
        $layout.Controls.Add($label, 0, $row)
        $box = New-Object Windows.Forms.TextBox
        $box.Dock = 'Fill'
        $box.Margin = New-Object Windows.Forms.Padding(0, 6, 8, 6)
        $layout.Controls.Add($box, 1, $row)
        $inputs[$key] = $box
        if ($spec[3]) {
            $button = New-Object Windows.Forms.Button
            $button.Text = $spec[3]
            $button.AutoSize = $true
            $button.Dock = 'Fill'
            $button.Margin = New-Object Windows.Forms.Padding(0, 5, 0, 5)
            $layout.Controls.Add($button, 2, $row)
            $browsers[$key] = $button
        } else { $layout.SetColumnSpan($box, 2) }
    }
    $dateLabel = New-Object Windows.Forms.Label
    $dateLabel.Text = 'Report date'
    $dateLabel.AutoSize = $true
    $dateLabel.Anchor = 'Left'
    $layout.Controls.Add($dateLabel, 0, 3)
    $datePicker = New-Object Windows.Forms.DateTimePicker
    $datePicker.Format = 'Custom'
    $datePicker.CustomFormat = 'yyyy-MM-dd'
    $datePicker.Width = 170
    $datePicker.Margin = New-Object Windows.Forms.Padding(0, 6, 8, 6)
    $layout.Controls.Add($datePicker, 1, 3)
    if ($ReportDate) { $datePicker.Value = [datetime]::ParseExact($ReportDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
    $inputs.Export.Text = $ExportPath
    if ($TrackingWorkbook) { $inputs.Tracking.Text = $TrackingWorkbook }
    else {
        $latest = Get-KEVLatestReport $PilotRoot
        if ($latest) { $inputs.Tracking.Text = $latest.FullName }
    }
    $inputs.Folder.Text = $PilotRoot
    $autoName = 'KEV-Report-' + $datePicker.Value.ToString('yyyy-MM-dd') + '.xlsx'
    $inputs.FileName.Text = $autoName
    if ($OutputPath) {
        $fullOutput = [IO.Path]::GetFullPath($OutputPath)
        $inputs.Folder.Text = [IO.Path]::GetDirectoryName($fullOutput)
        $inputs.FileName.Text = [IO.Path]::GetFileName($fullOutput)
    }

    $note = New-Object Windows.Forms.Label
    $note.Text = 'Save and close your edited report first. The utility archives it and creates a new workbook.'
    $note.AutoSize = $true
    $note.ForeColor = [Drawing.Color]::FromArgb(80, 90, 100)
    $note.Margin = New-Object Windows.Forms.Padding(0, 10, 0, 10)
    $layout.Controls.Add($note, 0, 6)
    $layout.SetColumnSpan($note, 3)
    $status = New-Object Windows.Forms.Label
    $status.Text = 'Ready. Select the export, then choose Create report.'
    $status.AutoSize = $true
    $status.MaximumSize = New-Object Drawing.Size(825, 0)
    $status.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 8)
    $layout.Controls.Add($status, 0, 7)
    $layout.SetColumnSpan($status, 3)
    $log = New-Object Windows.Forms.TextBox
    $log.Multiline = $true
    $log.ReadOnly = $true
    $log.ScrollBars = 'Both'
    $log.WordWrap = $false
    $log.Dock = 'Fill'
    $log.BackColor = [Drawing.Color]::White
    $log.Font = New-Object Drawing.Font('Consolas', 9)
    $log.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 10)
    $layout.Controls.Add($log, 0, 8)
    $layout.SetColumnSpan($log, 3)
    $progress = New-Object Windows.Forms.ProgressBar
    $progress.Dock = 'Fill'
    $progress.Height = 8
    $progress.MarqueeAnimationSpeed = 0
    $layout.Controls.Add($progress, 0, 9)
    $layout.SetColumnSpan($progress, 3)
    $buttons = New-Object Windows.Forms.FlowLayoutPanel
    $buttons.AutoSize = $true
    $buttons.Dock = 'Fill'
    $buttons.FlowDirection = 'RightToLeft'
    $buttons.Margin = New-Object Windows.Forms.Padding(0, 10, 0, 0)
    $close = New-Object Windows.Forms.Button
    $close.Text = 'Close'
    $close.AutoSize = $true
    $create = New-Object Windows.Forms.Button
    $create.Text = 'Create report'
    $create.AutoSize = $true
    $create.Padding = New-Object Windows.Forms.Padding(8, 2, 8, 2)
    $buttons.Controls.Add($close)
    $buttons.Controls.Add($create)
    $layout.Controls.Add($buttons, 0, 10)
    $layout.SetColumnSpan($buttons, 3)
    $form.AcceptButton = $create
    $form.CancelButton = $close

    $state = @{ Busy = $false; Process = $null; Stdout = ''; Stderr = ''; OutputPath = ''; AutoName = $autoName; LogDirectory = ''; LogText = '' }
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = 400
    $editable = @($inputs.Export, $inputs.Tracking, $inputs.Folder, $inputs.FileName, $datePicker, $create) + @($browsers.Values)
    $setBusy = {
        param([bool]$Busy)
        $state.Busy = $Busy
        foreach ($control in $editable) { $control.Enabled = !$Busy }
        $close.Enabled = !$Busy
        if ($Busy) { $progress.Style = 'Marquee'; $progress.MarqueeAnimationSpeed = 30 }
        else { $progress.MarqueeAnimationSpeed = 0; $progress.Style = 'Blocks'; $progress.Value = 0 }
    }.GetNewClosure()
    $updateLog = {
        $stdout = Read-KEVWorkerLog $state.Stdout
        $stderr = Read-KEVWorkerLog $state.Stderr
        $text = $stdout
        if ($stderr) { $text += "`r`n" + $stderr }
        if ($text -and $text -ne $state.LogText) {
            $state.LogText = $text
            $log.Text = $text
            $log.SelectionStart = $log.TextLength
            $log.ScrollToCaret()
        }
    }.GetNewClosure()

    $datePicker.Add_ValueChanged({
        if ($inputs.FileName.Text -eq $state.AutoName -or !$inputs.FileName.Text) {
            $inputs.FileName.Text = 'KEV-Report-' + $datePicker.Value.ToString('yyyy-MM-dd') + '.xlsx'
        }
        $state.AutoName = 'KEV-Report-' + $datePicker.Value.ToString('yyyy-MM-dd') + '.xlsx'
    }.GetNewClosure())
    foreach ($key in @('Export', 'Tracking')) {
        $selectedKey = $key
        $browsers[$key].Add_Click({
            $dialog = New-Object Windows.Forms.OpenFileDialog
            try {
                $dialog.Filter = if ($selectedKey -eq 'Export') {
                    'Defender exports (*.csv;*.xlsx)|*.csv;*.xlsx|CSV exports (*.csv)|*.csv|Excel workbooks (*.xlsx)|*.xlsx'
                } else { 'Excel workbooks (*.xlsx)|*.xlsx' }
                $dialog.CheckFileExists = $true
                $dialog.Multiselect = $false
                $dialog.Title = if ($selectedKey -eq 'Export') { 'Choose the Defender export' } else { 'Choose the saved report with your latest edits and history' }
                $dialog.InitialDirectory = $PilotRoot
                $currentPath = $inputs[$selectedKey].Text.Trim()
                if (![string]::IsNullOrWhiteSpace($currentPath) -and (Test-Path -LiteralPath $currentPath -PathType Leaf)) {
                    $dialog.FileName = (Resolve-Path -LiteralPath $currentPath).ProviderPath
                }
                if ((Show-KEVPickerDialog -Dialog $dialog -Owner $form) -eq [Windows.Forms.DialogResult]::OK) { $inputs[$selectedKey].Text = $dialog.FileName }
            } catch {
                $status.Text = 'Could not open file picker: ' + $_.Exception.Message
                $status.ForeColor = [Drawing.Color]::DarkRed
            } finally { $dialog.Dispose() }
        }.GetNewClosure())
    }
    $browsers.Folder.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        try {
            $dialog.Description = 'Choose where to save the new KEV report'
            $dialog.ShowNewFolderButton = $true
            $currentFolder = $inputs.Folder.Text.Trim()
            if (![string]::IsNullOrWhiteSpace($currentFolder) -and (Test-Path -LiteralPath $currentFolder -PathType Container)) { $dialog.SelectedPath = $currentFolder }
            if ((Show-KEVPickerDialog -Dialog $dialog -Owner $form) -eq [Windows.Forms.DialogResult]::OK) { $inputs.Folder.Text = $dialog.SelectedPath }
        } catch {
            $status.Text = 'Could not open folder picker: ' + $_.Exception.Message
            $status.ForeColor = [Drawing.Color]::DarkRed
        } finally { $dialog.Dispose() }
    }.GetNewClosure())

    $create.Add_Click({
        try {
            $request = Resolve-KEVReportRequest -PilotRoot $PilotRoot -ExportPath $inputs.Export.Text.Trim() -TrackingWorkbook $inputs.Tracking.Text.Trim() -ReportDate $datePicker.Value.Date -OutputDirectory $inputs.Folder.Text.Trim() -OutputFileName $inputs.FileName.Text.Trim()
            $state.LogDirectory = Join-Path $PilotRoot ('validation\gui-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $state.LogDirectory | Out-Null
            $state.Stdout = Join-Path $state.LogDirectory 'refresh-output.log'
            $state.Stderr = Join-Path $state.LogDirectory 'refresh-error.log'
            $state.OutputPath = $request.OutputPath
            $log.Text = ''
            $state.LogText = ''
            $status.Text = 'Creating report. This can take a few minutes. Keep this window open.'
            $status.ForeColor = [Drawing.Color]::FromArgb(30, 60, 90)
            & $setBusy $true
            $startOptions = @{
                FilePath = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
                ArgumentList = (Get-KEVWorkerArguments -PilotRoot $PilotRoot -Request $request)
                WorkingDirectory = $PilotRoot
                WindowStyle = 'Hidden'
                RedirectStandardOutput = $state.Stdout
                RedirectStandardError = $state.Stderr
                PassThru = $true
                ErrorAction = 'Stop'
            }
            $state.Process = Start-Process @startOptions
            # Windows PowerShell's Start-Process can otherwise lose the native
            # handle before a later timer reads ExitCode. Keep it open now.
            $null = $state.Process.Handle
            $timer.Start()
        } catch {
            & $setBusy $false
            $status.Text = $_.Exception.Message
            $status.ForeColor = [Drawing.Color]::DarkRed
        }
    }.GetNewClosure())

    $timer.Add_Tick({
        try {
            & $updateLog
            $state.Process.Refresh()
            if (!$state.Process.HasExited) { return }
            $timer.Stop()
            $state.Process.WaitForExit()
            $exitCode = $state.Process.ExitCode
            & $updateLog
            $state.Process.Dispose()
            $state.Process = $null
            & $setBusy $false
            if ($exitCode -eq 0 -and (Test-Path -LiteralPath $state.OutputPath -PathType Leaf)) {
                $status.Text = 'Created: ' + $state.OutputPath
                $status.ForeColor = [Drawing.Color]::DarkGreen
                $log.AppendText("`r`nReport saved to: " + $state.OutputPath)
            } else {
                $status.Text = 'Report creation failed. Details are shown below. Correct the settings and try again.'
                $status.ForeColor = [Drawing.Color]::DarkRed
                $log.AppendText("`r`nExit code: $exitCode`r`nLogs: " + $state.LogDirectory)
            }
        } catch {
            $timer.Stop()
            # Keep the form busy if the generator is still running; closing it
            # must not leave the user unaware of an active report write.
            $running = $state.Process -and !$state.Process.HasExited
            if (!$running) { & $setBusy $false }
            $status.Text = 'Could not read report progress. Logs: ' + $state.LogDirectory
            $log.AppendText("`r`n" + $_.Exception.Message)
            if ($running) { $timer.Start() }
        }
    }.GetNewClosure())
    $close.Add_Click({ $form.Close() }.GetNewClosure())
    $form.Add_FormClosing({
        param($sender, $eventArgs)
        if ($state.Busy) { $eventArgs.Cancel = $true }
    }.GetNewClosure())
    $form.Add_FormClosed({
        $timer.Stop()
        $timer.Dispose()
        if ($state.Process) { $state.Process.Dispose() }
    }.GetNewClosure())
    return [pscustomobject]@{ Form = $form; Inputs = $inputs; DatePicker = $datePicker; CreateButton = $create; CloseButton = $close; BrowseButtons = $browsers; State = $state; Timer = $timer; Status = $status; Log = $log }
}

function Show-KEVReportDialog {
    param([string]$PilotRoot, [string]$ExportPath, [string]$ReportDate, [string]$TrackingWorkbook, [string]$OutputPath)
    $ui = New-KEVReportForm @PSBoundParameters
    try { $ui.Form.ShowDialog() | Out-Null }
    finally { $ui.Timer.Dispose(); $ui.Form.Dispose() }
}

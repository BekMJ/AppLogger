Param(
    [Parameter(Mandatory=$false)] [string] $AppPath,
    [Parameter(Mandatory=$false)] [string] $AppArgs,
    [Parameter(Mandatory=$false)] [string] $AppName,
    [Parameter(Mandatory=$false)] [string] $LogDir = (Join-Path -Path $PSScriptRoot -ChildPath 'Logs'),
    [Parameter(Mandatory=$false)] [switch] $ForceCsvOnly
)

$ErrorActionPreference = 'Stop'

function Ensure-DirectoryExists {
    Param([string] $Directory)
    if (-not (Test-Path -LiteralPath $Directory)) { [void](New-Item -ItemType Directory -Path $Directory) }
}

function Get-LogBaseName {
    Param([string] $Prefix = 'LabUsage')
    $yyyymm = Get-Date -Format 'yyyyMM'
    return "$Prefix-$yyyymm"
}

function Test-ExcelAvailable {
    try {
        $excel = New-Object -ComObject Excel.Application
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        return $true
    } catch {
        return $false
    }
}

function Write-LogToCsv {
    Param(
        [string] $Directory,
        [hashtable] $Entry
    )
    Ensure-DirectoryExists -Directory $Directory
    $base = Get-LogBaseName
    $path = Join-Path -Path $Directory -ChildPath ("$base.csv")

    $header = 'Timestamp,Name,Advisor,Experiment,ComputerName,WindowsUser,App'
    if (-not (Test-Path -LiteralPath $path)) { $header | Out-File -FilePath $path -Encoding UTF8 }

    $csvLine = (
        @(
            $Entry.Timestamp,
            $Entry.Name,
            $Entry.Advisor,
            $Entry.Experiment,
            $Entry.ComputerName,
            $Entry.WindowsUser,
            $Entry.App
        ) | ForEach-Object {
            $value = [string]$_
            if ($value -match '[",\n]') { '"' + ($value -replace '"','""') + '"' } else { $value }
        }
    ) -join ','

    Add-Content -Path $path -Value $csvLine
    return $path
}

function Write-LogToExcelXlsx {
    Param(
        [string] $Directory,
        [hashtable] $Entry
    )
    Ensure-DirectoryExists -Directory $Directory
    $base = Get-LogBaseName
    $xlsxPath = Join-Path -Path $Directory -ChildPath ("$base.xlsx")

    $excel = $null
    $workbook = $null
    $worksheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        if (Test-Path -LiteralPath $xlsxPath) {
            $workbook = $excel.Workbooks.Open($xlsxPath)
        } else {
            $workbook = $excel.Workbooks.Add()
        }

        $worksheet = $workbook.Worksheets.Item(1)
        if (-not $worksheet) { $worksheet = $workbook.Worksheets.Add() }
        $worksheet.Name = 'Log'

        $headerMap = @('Timestamp','Name','Advisor','Experiment','ComputerName','WindowsUser','App')
        $lastRow = $worksheet.UsedRange.Rows.Count
        if ($lastRow -eq 1 -and [string]::IsNullOrWhiteSpace(($worksheet.Cells.Item(1,1)).Text)) { $lastRow = 0 }
        if ($lastRow -lt 1) {
            for ($i=0; $i -lt $headerMap.Count; $i++) { $worksheet.Cells.Item(1, $i+1) = $headerMap[$i] }
            $lastRow = 1
        }

        $targetRow = $lastRow + 1
        $values = @(
            $Entry.Timestamp,
            $Entry.Name,
            $Entry.Advisor,
            $Entry.Experiment,
            $Entry.ComputerName,
            $Entry.WindowsUser,
            $Entry.App
        )
        for ($i=0; $i -lt $values.Count; $i++) { $worksheet.Cells.Item($targetRow, $i+1) = $values[$i] }

        if (Test-Path -LiteralPath $xlsxPath) {
            $workbook.Save()
        } else {
            $workbook.SaveAs($xlsxPath)
        }
        return $xlsxPath
    } finally {
        if ($workbook -ne $null) { $workbook.Close($true) | Out-Null }
        if ($excel -ne $null) { $excel.Quit() }
        if ($worksheet -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
        if ($workbook -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
        if ($excel -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Show-LoggerForm {
    Param([string] $Title = 'Lab Usage Logger')
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(420, 250)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = 'Name:'
    $lblName.Location = New-Object System.Drawing.Point(15, 20)
    $lblName.AutoSize = $true

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(120, 16)
    $txtName.Size = New-Object System.Drawing.Size(260, 20)

    $lblAdvisor = New-Object System.Windows.Forms.Label
    $lblAdvisor.Text = 'Advisor:'
    $lblAdvisor.Location = New-Object System.Drawing.Point(15, 60)
    $lblAdvisor.AutoSize = $true

    $txtAdvisor = New-Object System.Windows.Forms.TextBox
    $txtAdvisor.Location = New-Object System.Drawing.Point(120, 56)
    $txtAdvisor.Size = New-Object System.Drawing.Size(260, 20)

    $lblExperiment = New-Object System.Windows.Forms.Label
    $lblExperiment.Text = 'Experiment:'
    $lblExperiment.Location = New-Object System.Drawing.Point(15, 100)
    $lblExperiment.AutoSize = $true

    $txtExperiment = New-Object System.Windows.Forms.TextBox
    $txtExperiment.Location = New-Object System.Drawing.Point(120, 96)
    $txtExperiment.Size = New-Object System.Drawing.Size(260, 20)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Location = New-Object System.Drawing.Point(200, 150)
    $btnOk.Enabled = $false

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(300, 150)

    $validate = {
        if ([string]::IsNullOrWhiteSpace($txtName.Text) -or [string]::IsNullOrWhiteSpace($txtAdvisor.Text) -or [string]::IsNullOrWhiteSpace($txtExperiment.Text)) {
            $btnOk.Enabled = $false
        } else { $btnOk.Enabled = $true }
    }
    $txtName.Add_TextChanged($validate)
    $txtAdvisor.Add_TextChanged($validate)
    $txtExperiment.Add_TextChanged($validate)

    $result = $null
    $btnOk.Add_Click({
        $result = @{
            Name = $txtName.Text.Trim()
            Advisor = $txtAdvisor.Text.Trim()
            Experiment = $txtExperiment.Text.Trim()
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel

    $form.Controls.AddRange(@($lblName,$txtName,$lblAdvisor,$txtAdvisor,$lblExperiment,$txtExperiment,$btnOk,$btnCancel))

    [void]$form.ShowDialog()
    return $result
}

function Start-TargetApp {
    Param(
        [string] $Path,
        [string] $Arguments
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (Test-Path -LiteralPath $Path) {
        Start-Process -FilePath $Path -ArgumentList $Arguments -WindowStyle Normal | Out-Null
    } else {
        Start-Process -FilePath $Path -ArgumentList $Arguments -WindowStyle Normal | Out-Null
    }
}

# Main
try {
    if (-not [string]::IsNullOrWhiteSpace($AppName)) {
        $title = "Lab Usage Logger — $AppName"
    } elseif (-not [string]::IsNullOrWhiteSpace($AppPath)) {
        $title = "Lab Usage Logger — " + [System.IO.Path]::GetFileNameWithoutExtension($AppPath)
    } else {
        $title = 'Lab Usage Logger'
    }

    $inputData = Show-LoggerForm -Title $title
    if ($null -eq $inputData) { exit 1 }

    $entry = @{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Name = $inputData.Name
        Advisor = $inputData.Advisor
        Experiment = $inputData.Experiment
        ComputerName = $env:COMPUTERNAME
        WindowsUser = $env:USERNAME
        App = if ($AppName) { $AppName } elseif ($AppPath) { [System.IO.Path]::GetFileNameWithoutExtension($AppPath) } else { '' }
    }

    $csvPath = Write-LogToCsv -Directory $LogDir -Entry $entry

    if (-not $ForceCsvOnly) {
        if (Test-ExcelAvailable) {
            try { [void](Write-LogToExcelXlsx -Directory $LogDir -Entry $entry) } catch { }
        }
    }

    Start-TargetApp -Path $AppPath -Arguments $AppArgs
    exit 0
} catch {
    $msg = "AppLogger error: " + $_.Exception.Message
    try { [System.Windows.Forms.MessageBox]::Show($msg, 'AppLogger', 'OK', 'Error') | Out-Null } catch { Write-Error $msg }
    exit 2
}



<#
Weekly Status Report automation -- one-time setup for a new user.

Share this file together with Send-WeeklyStatusReport.ps1 (keep them in the same folder).
Run this script once. It will:
  1. Create a WSR folder (and a Scripts subfolder) for you.
  2. Copy the automation script into place.
  3. Build a blank WeeklyEffortInput.xlsx for you to fill in (My Info / Project Summary /
     Effort Details / Settings tabs).
  4. Register a Windows scheduled task that runs the report every Friday at 7:00 PM.

You never need to open or edit the PowerShell script itself -- everything personal lives
in the Excel file.
#>

param(
    [string]$TargetFolder,
    [string]$TaskName = "WeeklyStatusReport"
)

$ErrorActionPreference = "Stop"

if (-not $TargetFolder) {
    if ($env:OneDrive) {
        $TargetFolder = Join-Path $env:OneDrive "WSR"
    } else {
        $TargetFolder = Join-Path $env:USERPROFILE "WSR"
    }
}

Write-Output "Setting up Weekly Status Report automation in: $TargetFolder"

$scriptsFolder = Join-Path $TargetFolder "Scripts"
New-Item -ItemType Directory -Force -Path $scriptsFolder | Out-Null

# ---- Copy the generic automation script (must be next to this installer) ----
$sourceScript = Join-Path $PSScriptRoot "Send-WeeklyStatusReport.ps1"
if (-not (Test-Path $sourceScript)) {
    throw "Send-WeeklyStatusReport.ps1 was not found next to this installer. Keep both files together in the same folder before running Setup-NewUser.ps1."
}
$destScript = Join-Path $scriptsFolder "Send-WeeklyStatusReport.ps1"
Copy-Item $sourceScript $destScript -Force
Write-Output "Copied automation script to $destScript"

# ---- Build a blank WeeklyEffortInput.xlsx (skip if one already exists, so re-running this is safe) ----
$excelPath = Join-Path $TargetFolder "WeeklyEffortInput.xlsx"
if (Test-Path $excelPath) {
    Write-Output "WeeklyEffortInput.xlsx already exists -- leaving it as-is."
} else {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()

    # -- My Info --
    $wsInfo = $wb.Worksheets.Item(1)
    $wsInfo.Name = "My Info"
    $wsInfo.Cells.Item(1,1) = "Field"
    $wsInfo.Cells.Item(1,2) = "Value"
    $wsInfo.Range("A1:B1").Font.Bold = $true
    $wsInfo.Range("A1:B1").Interior.Color = 0xD9CFC2
    $infoFields = @("Name","Employee ID","Supervisor","Greeting","To Email","CC Email","Subject")
    $r = 2
    foreach ($f in $infoFields) { $wsInfo.Cells.Item($r,1) = $f; $r++ }
    $wsInfo.Range("A1:B$($r-1)").Borders.LineStyle = 1
    $wsInfo.Columns.Item(1).ColumnWidth = 18
    $wsInfo.Columns.Item(2).ColumnWidth = 45

    $noteRow1 = $r + 1
    $wsInfo.Cells.Item($noteRow1,1) = "Note: This tab holds all your personal info used to build the report -- Name (Project Details), Employee ID, Supervisor, Greeting (first line of the email), To/CC email addresses, and the email Subject. Nothing is hardcoded in the script -- everything comes from these values."
    $wsInfo.Cells.Item($noteRow1,1).Font.Italic = $true
    $wsInfo.Range("A$($noteRow1):B$($noteRow1)").Merge() | Out-Null
    $wsInfo.Range("A$($noteRow1)").WrapText = $true
    $wsInfo.Range("A$($noteRow1)").RowHeight = 45

    $noteRow2 = $noteRow1 + 1
    $wsInfo.Cells.Item($noteRow2,1) = "New teammate setup: run Setup-NewUser.ps1 (kept together with Send-WeeklyStatusReport.ps1) once. It creates your own WSR folder, copies the script, builds this file for you, and registers your Friday 7:00 PM scheduled task -- then just fill in this tab with your own details."
    $wsInfo.Cells.Item($noteRow2,1).Font.Italic = $true
    $wsInfo.Range("A$($noteRow2):B$($noteRow2)").Merge() | Out-Null
    $wsInfo.Range("A$($noteRow2)").WrapText = $true
    $wsInfo.Range("A$($noteRow2)").RowHeight = 45

    # -- Project Summary --
    $wsSummary = $wb.Worksheets.Add([Type]::Missing, $wsInfo)
    $wsSummary.Name = "Project Summary"
    $wsSummary.Cells.Item(1,1) = "Project Name"
    $wsSummary.Cells.Item(1,2) = "Summary"
    $wsSummary.Range("A1:B1").Font.Bold = $true
    $wsSummary.Range("A1:B1").Interior.Color = 0xD9CFC2
    $wsSummary.Cells.Item(2,1) = "Example Project"
    $wsSummary.Cells.Item(2,2) = "Replace this row with your own project name and a short summary of what you worked on this week."
    $wsSummary.Range("A1:B2").Borders.LineStyle = 1
    $wsSummary.Range("B2").WrapText = $true
    $wsSummary.Columns.Item(1).ColumnWidth = 22
    $wsSummary.Columns.Item(2).ColumnWidth = 90

    # -- Effort Details --
    $wsEffort = $wb.Worksheets.Add([Type]::Missing, $wsSummary)
    $wsEffort.Name = "Effort Details"
    $headers = @("Day","Project","Task","Description","Tickets","Blockers/Remarks")
    for ($c = 0; $c -lt 6; $c++) { $wsEffort.Cells.Item(1,$c+1) = $headers[$c] }
    $wsEffort.Range("A1:F1").Font.Bold = $true
    $wsEffort.Range("A1:F1").Interior.Color = 0xD9CFC2
    $days = @("Monday","Tuesday","Wednesday","Thursday","Friday")
    $er = 2
    foreach ($d in $days) {
        $wsEffort.Cells.Item($er,1) = $d
        $wsEffort.Cells.Item($er,2) = "Example Project"
        $wsEffort.Cells.Item($er,3) = ""
        $wsEffort.Cells.Item($er,4) = ""
        $wsEffort.Cells.Item($er,5) = ""
        $wsEffort.Cells.Item($er,6) = ""
        $er++
        $wsEffort.Cells.Item($er,1) = $d
        $wsEffort.Cells.Item($er,2) = "Internal"
        $wsEffort.Cells.Item($er,3) = "Non Billable/ Meetings"
        $wsEffort.Cells.Item($er,4) = "Meetings - Daily Internal Standup"
        $wsEffort.Cells.Item($er,5) = ""
        $wsEffort.Cells.Item($er,6) = ""
        $er++
    }
    $wsEffort.Cells.Item($er+1,1) = "Note: Project must exactly match a name from the 'Project Summary' sheet, or be 'Internal'."
    $wsEffort.Cells.Item($er+1,1).Font.Italic = $true
    $wsEffort.Range("A1:F$($er-1)").Borders.LineStyle = 1
    $wsEffort.Range("D2:D$($er-1)").WrapText = $true
    $wsEffort.Columns.Item(1).ColumnWidth = 12
    $wsEffort.Columns.Item(2).ColumnWidth = 16
    $wsEffort.Columns.Item(3).ColumnWidth = 24
    $wsEffort.Columns.Item(4).ColumnWidth = 60
    $wsEffort.Columns.Item(5).ColumnWidth = 16
    $wsEffort.Columns.Item(6).ColumnWidth = 22

    # -- Settings --
    $wsSettings = $wb.Worksheets.Add([Type]::Missing, $wsEffort)
    $wsSettings.Name = "Settings"
    $wsSettings.Cells.Item(1,1) = "Setting"
    $wsSettings.Cells.Item(1,2) = "Value"
    $wsSettings.Range("A1:B1").Font.Bold = $true
    $wsSettings.Range("A1:B1").Interior.Color = 0xD9CFC2
    $wsSettings.Cells.Item(2,1) = "Auto Send (Yes/No)"
    $wsSettings.Cells.Item(2,2) = "No"
    $wsSettings.Cells.Item(3,1) = "Show Project Summary (Yes/No)"
    $wsSettings.Cells.Item(3,2) = "Yes"
    $wsSettings.Cells.Item(4,1) = "Note: 'Auto Send' set to 'Yes' before the Friday run sends the report automatically instead of just saving it as a draft. 'Show Project Summary' set to 'No' hides the narrative Project Summary section from the email entirely, leaving just the Project Details and Effort Details tables."
    $wsSettings.Cells.Item(4,1).Font.Italic = $true
    $wsSettings.Range("A1:B4").Borders.LineStyle = 1
    $wsSettings.Range("A4").WrapText = $true
    $wsSettings.Range("A4").RowHeight = 45
    $wsSettings.Columns.Item(1).ColumnWidth = 24
    $wsSettings.Columns.Item(2).ColumnWidth = 14

    $wsInfo.Activate()
    try { $wb.SaveAs($excelPath, 51) } catch { }
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsInfo) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsSummary) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsEffort) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsSettings) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Output "Created blank input file: $excelPath"
}

# ---- Register the scheduled task (Friday 7:00 PM, catches up if you're logged off at that time) ----
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$destScript`""
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At 7:00PM
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Weekly status report automation" -Force | Out-Null
Write-Output "Registered scheduled task '$TaskName' -- runs every Friday at 7:00 PM."

Write-Output ""
Write-Output "Setup complete. Next steps:"
Write-Output "  1. Open $excelPath"
Write-Output "  2. Fill in the 'My Info' tab with your own Name, Employee ID, Supervisor, Greeting, To/CC emails, and Subject."
Write-Output "  3. Fill in 'Project Summary' and 'Effort Details' with this week's work."
Write-Output "  4. Leave 'Settings' -> Auto Send as 'No' until you're comfortable with the output."
Write-Output "  5. Test it: powershell -NoProfile -ExecutionPolicy Bypass -File `"$destScript`""

Start-Process $excelPath

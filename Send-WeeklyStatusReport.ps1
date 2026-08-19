$ErrorActionPreference = "Stop"

# ---- Compute Monday (Week Beginning) and Friday (Report Date) of the current week ----
# Works no matter which day of the week the script actually runs on.
$today = Get-Date
$dow = [int]$today.DayOfWeek                 # Sunday=0 ... Saturday=6
$daysSinceMonday = ($dow + 6) % 7             # Monday=0, Tuesday=1, ... Sunday=6
$weekBegin  = $today.AddDays(-$daysSinceMonday)
$reportDate = $weekBegin.AddDays(4)           # Friday of that same week
$fmt = "MM/dd/yyyy"

$weekBeginStr  = $weekBegin.ToString($fmt)
$reportDateStr = $reportDate.ToString($fmt)
$dayNames = @("Monday","Tuesday","Wednesday","Thursday","Friday")
$weekDates = 0..4 | ForEach-Object { $weekBegin.AddDays($_).ToString($fmt) }
$dayDateLookup = @{}
for ($i = 0; $i -lt 5; $i++) { $dayDateLookup[$dayNames[$i]] = $weekDates[$i] }

# ================= Read this week's data from the input file =================
# Resolved relative to this script's own location (Scripts\Send-WeeklyStatusReport.ps1),
# so this file is identical for every user -- nothing here needs to be edited per person.
$inputPath = Join-Path (Split-Path $PSScriptRoot -Parent) "WeeklyEffortInput.xlsx"
if (-not (Test-Path $inputPath)) {
    throw "Input file not found: $inputPath"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($inputPath, [Type]::Missing, $true) # read-only

# ---- Sheet 1: Project Summary (defines the canonical list of project names) ----
$wsSummary = $wb.Worksheets.Item("Project Summary")
$projects = @()
$r = 2
while ($true) {
    $name = "$($wsSummary.Cells.Item($r,1).Text)".Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { break }
    $summary = "$($wsSummary.Cells.Item($r,2).Text)".Trim()
    $projects += @{ Name = $name; Summary = $summary }
    $r++
}

# ---- Sheet 2: Effort Details (variable rows: Day, Project, Task, Description, Blockers) ----
$wsEffort = $wb.Worksheets.Item("Effort Details")
$effortEntries = @()
$er = 2
while ($true) {
    $day = "$($wsEffort.Cells.Item($er,1).Text)".Trim()
    if ([string]::IsNullOrWhiteSpace($day)) { break }
    $effortEntries += @{
        Day         = $day
        Project     = "$($wsEffort.Cells.Item($er,2).Text)".Trim()
        Task        = "$($wsEffort.Cells.Item($er,3).Text)".Trim()
        Description = "$($wsEffort.Cells.Item($er,4).Text)".Trim()
        Blockers    = "$($wsEffort.Cells.Item($er,5).Text)".Trim()
    }
    $er++
}

# ---- Sheet 3: Settings (Auto Send toggle) ----
$autoSend = $false
try {
    $wsSettings = $wb.Worksheets.Item("Settings")
    $autoSendValue = "$($wsSettings.Cells.Item(2,2).Text)".Trim()
    $autoSend = ($autoSendValue -ieq "Yes")
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsSettings) | Out-Null
} catch {
    $autoSend = $false   # no Settings sheet found -> safest default is draft-only
}

# ---- Sheet 4: My Info (personal fields -- this is what makes the script identical for every user) ----
$myInfo = @{}
try {
    $wsMyInfo = $wb.Worksheets.Item("My Info")
    $mr = 2
    while ($true) {
        $field = "$($wsMyInfo.Cells.Item($mr,1).Text)".Trim()
        if ([string]::IsNullOrWhiteSpace($field)) { break }
        $myInfo[$field] = "$($wsMyInfo.Cells.Item($mr,2).Text)".Trim()
        $mr++
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsMyInfo) | Out-Null
} catch {
    throw "'My Info' sheet not found in $inputPath -- add it with Name, Employee ID, Supervisor, Greeting, To Email, CC Email, and Subject rows."
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsSummary) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsEffort) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

$userDisplayName = $myInfo["Name"]
$employeeId      = $myInfo["Employee ID"]
$supervisorName  = $myInfo["Supervisor"]
$greetingLine    = $myInfo["Greeting"]
$toEmail         = $myInfo["To Email"]
$ccEmail         = $myInfo["CC Email"]
$emailSubject    = $myInfo["Subject"]

$requiredFields = @("Name","Employee ID","Supervisor","Greeting","To Email","CC Email","Subject")
$missingFields = $requiredFields | Where-Object { [string]::IsNullOrWhiteSpace($myInfo[$_]) }
if ($missingFields.Count -gt 0) {
    throw "'My Info' sheet is missing value(s) for: $($missingFields -join ', ')"
}

if ($projects.Count -eq 0) {
    throw "No project rows found in $inputPath (Project Summary sheet)"
}

# The Project Details field and every Effort Details row are driven off this same list,
# so the project name only ever needs to be typed once, in the Project Summary sheet.
# "Internal" is a standing pseudo-project (daily standup) and doesn't need a Project Summary entry.
$projectFieldValue = ($projects | ForEach-Object { $_.Name }) -join ", "
$knownProjectNames = @("Internal") + ($projects | ForEach-Object { $_.Name })
$unrecognized = $effortEntries | Where-Object { $knownProjectNames -notcontains $_.Project } | ForEach-Object { $_.Project }
if ($unrecognized.Count -gt 0) {
    Write-Output "WARNING: Effort Details has project name(s) not found in Project Summary: $($unrecognized -join ', ')"
}

# ---- Colors ----
$purple     = "#5F358A"
$sectionBlu = "#8EA9DB"
$bandGrey   = "#D9D9D9"
$bandWhite  = "#FFFFFF"
$internalBg = "#F2F2F2"
$border     = "1px solid #BFBFBF"

function HtmlEncode($text) {
    if ([string]::IsNullOrEmpty($text)) { return "&nbsp;" }
    return $text.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace("`n","<br>")
}

$projectRows = ""
foreach ($p in $projects) {
    $projectRows += @"
  <tr>
    <td style="border:$border;padding:8px;">
      <b>$(HtmlEncode $p.Name):</b><br>
      $(HtmlEncode $p.Summary)
    </td>
  </tr>
"@
}

# Group Effort Details rows by weekday (in Monday-Friday order); rows render in the order
# they appear in the sheet for that day, including any "Internal" rows typed there.
$effortRows = ""
foreach ($dayName in $dayNames) {
    $d = $dayDateLookup[$dayName]
    $dayEntries = $effortEntries | Where-Object { $_.Day -eq $dayName }
    foreach ($entry in $dayEntries) {
        $rowBg = if ($entry.Project -eq "Internal") { $internalBg } else { $bandWhite }
        $effortRows += @"
  <tr style="background:$rowBg;">
    <td style="border:$border;padding:6px;text-align:center;">$d</td>
    <td style="border:$border;padding:6px;text-align:center;">$(HtmlEncode $entry.Project)</td>
    <td style="border:$border;padding:6px;">$(HtmlEncode $entry.Task)</td>
    <td style="border:$border;padding:6px;">$(HtmlEncode $entry.Description)</td>
    <td style="border:$border;padding:6px;text-align:center;">$(HtmlEncode $entry.Blockers)</td>
  </tr>
"@
    }
}

$html = @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;margin:0;padding:0;">
<p style="margin:0 0 10px 0;">$greetingLine</p>
<p style="margin:0 0 14px 0;">Please find the weekly status for the week $weekBeginStr to $reportDateStr below. Kindly let me know if you need any further clarification.</p>
<table style="border-collapse:collapse;width:700px;margin:0 0 14px 0;">
  <tr><td colspan="1" style="border:$border;background:$sectionBlu;font-weight:bold;text-align:center;padding:6px;">Project Summary for This Week</td></tr>
$projectRows
</table>
<table style="border-collapse:collapse;width:700px;margin:0 0 14px 0;">
  <tr><td colspan="2" style="border:$border;background:$purple;color:#FFFFFF;font-weight:bold;text-align:center;padding:8px;">Weekly Status Report</td></tr>
  <tr><td colspan="2" style="border:$border;background:$sectionBlu;font-weight:bold;text-align:center;padding:6px;">Project Details</td></tr>
  <tr style="background:$bandGrey;"><td style="border:$border;font-weight:bold;padding:6px;width:180px;">Name</td><td style="border:$border;padding:6px;width:400px;">$userDisplayName</td></tr>
  <tr style="background:$bandWhite;"><td style="border:$border;font-weight:bold;padding:6px;">Project</td><td style="border:$border;padding:6px;">$projectFieldValue</td></tr>
  <tr style="background:$bandGrey;"><td style="border:$border;font-weight:bold;padding:6px;">Week Beginning</td><td style="border:$border;padding:6px;">$weekBeginStr</td></tr>
  <tr style="background:$bandWhite;"><td style="border:$border;font-weight:bold;padding:6px;">Supervisor</td><td style="border:$border;padding:6px;">$supervisorName</td></tr>
  <tr style="background:$bandGrey;"><td style="border:$border;font-weight:bold;padding:6px;">Report Date</td><td style="border:$border;padding:6px;">$reportDateStr</td></tr>
  <tr style="background:$bandWhite;"><td style="border:$border;font-weight:bold;padding:6px;">Emp ID</td><td style="border:$border;padding:6px;">$employeeId</td></tr>
</table>
<table style="border-collapse:collapse;width:700px;margin:0 0 14px 0;">
  <tr><td colspan="5" style="border:$border;background:$sectionBlu;font-weight:bold;text-align:center;padding:6px;">Effort Details</td></tr>
  <tr style="background:$bandGrey;font-weight:bold;">
    <td style="border:$border;text-align:center;padding:6px;width:100px;">Date</td>
    <td style="border:$border;text-align:center;padding:6px;width:90px;">Project</td>
    <td style="border:$border;text-align:center;padding:6px;width:170px;">Task</td>
    <td style="border:$border;text-align:center;padding:6px;width:250px;">Description</td>
    <td style="border:$border;text-align:center;padding:6px;width:90px;">Blockers/Remarks</td>
  </tr>
$effortRows
</table>
<p style="margin:0;">Thanks,</p>
</div>
"@

# ================= Create the Outlook draft (reply to last week's sent email, so every =================
# ================= week's report threads together under one fixed subject) =================
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$sentFolder = $ns.GetDefaultFolder(5) # olFolderSentMail
$sentItems = $sentFolder.Items
$sentItems.Sort("[SentOn]", $true)    # most recent first

$lastSent = $null
for ($i = 1; $i -le $sentItems.Count; $i++) {
    $candidate = $sentItems.Item($i)
    if ($candidate.Subject -eq $emailSubject) {
        $lastSent = $candidate
        break
    }
}

if ($lastSent) {
    $mail = $lastSent.Reply()   # true reply -- inherits conversation/thread headers
} else {
    $mail = $outlook.CreateItem(0) # first-ever run: no prior thread to reply to
}

$mail.To  = $toEmail
$mail.CC  = $ccEmail
$mail.Subject = $emailSubject   # Reply() prepends "RE: " -- reset to the exact fixed subject

$mail.Display()   # triggers Outlook to insert the default signature (and quoted history, if a reply) into HTMLBody
$signatureHtml = $mail.HTMLBody

# Remove the blank placeholder lines Outlook inserts before the signature in an empty compose window
$signatureHtml = $signatureHtml -replace '(<p class=MsoNormal><o:p>&nbsp;</o:p></p>\s*)+(?=<div><p class=MsoNormal><a name="_MailAutoSig")', ''

$bodyTagMatch = [regex]::Match($signatureHtml, '<body[^>]*>')
if ($bodyTagMatch.Success) {
    $insertPos = $bodyTagMatch.Index + $bodyTagMatch.Length
    $newHtml = $signatureHtml.Substring(0, $insertPos) + $html + $signatureHtml.Substring($insertPos)
} else {
    $newHtml = $html + $signatureHtml
}
$mail.HTMLBody = $newHtml

if ($autoSend) {
    $mail.Send()
    $resultMode = "SENT"
} else {
    $mail.Save()      # saved to Drafts, still open for review
    $resultMode = "DRAFT"
}

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

Write-Output "DONE:Mode=$resultMode WeekBeginning=$weekBeginStr ReportDate=$reportDateStr Projects=$projectFieldValue"

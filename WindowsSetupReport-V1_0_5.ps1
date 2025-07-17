# PowerShell System Report Generator with VBScript Log Extraction
Clear-Host

$ScriptName = "setup_info_windows"
$scriptVersion = "V_1_0_1"
$scriptGitHub = "https://github.com/muthdieter"
$scriptDate = "7.2025"

mode 300

Write-Host ""
Write-Host "             ____  __  __"
Write-Host "            |  _ \|  \/  |"
Write-Host "            | | | | |\/| |"
Write-Host "            | |_| | |  | |"
Write-Host "            |____/|_|  |_|"
Write-Host "   "
Write-Host ""
Write-Host "       $scriptGitHub " -ForegroundColor magenta
Write-Host ""
Write-Host "       $ScriptName   " -ForegroundColor Green
write-Host "       $scriptVersion" -ForegroundColor Green
write-host "       $scriptDate   " -ForegroundColor Green
Write-Host ""
Write-Host "  modified original script ( at Github) from Microsoft " -ForegroundColor Magenta
Write-Host ""
Write-Host ""
Pause

Add-Type -AssemblyName System.Windows.Forms
function Pick-Folder($prompt) {
    Write-Host "`n$prompt" -ForegroundColor Cyan
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $browser.Description = $prompt
    $null = $browser.ShowDialog()
    return $browser.SelectedPath
}


$OutputRoot = Pick-Folder "Select destination folder for setup report"
if (-not $OutputRoot -or -not (Test-Path $OutputRoot)) {
    Write-Host "Cancelled or invalid path selected. Exiting." -ForegroundColor Red
    exit 1
}

$ComputerName = $env:COMPUTERNAME
$UserName = $env:USERNAME
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFolder = Join-Path $OutputRoot "setup_report_${UserName}-${ComputerName}_$Timestamp"
New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null

Write-Host "`nCollecting system data into: $ReportFolder" -ForegroundColor Green

# === System Info ===
systeminfo > "$ReportFolder\systeminfo.txt"
Get-ComputerInfo | Out-File "$ReportFolder\computerinfo.txt"
Get-WmiObject Win32_OperatingSystem | Out-File "$ReportFolder\os_details.txt"

# === Windows Update History ===
Get-WmiObject -Class "Win32_QuickFixEngineering" | Out-File "$ReportFolder\windows_updates.txt"

# === Event Logs ===
wevtutil qe System /f:text /c:1000 > "$ReportFolder\eventlog_system.txt"
wevtutil qe Application /f:text /c:1000 > "$ReportFolder\eventlog_app.txt"

# === Temp + Panther ===
$TempFiles = Join-Path $env:TEMP "*"
$PantherFolder = "$env:SystemRoot\Panther"

Get-ChildItem $TempFiles -Recurse -ErrorAction SilentlyContinue | Out-File "$ReportFolder\temp_files.txt"
if (Test-Path $PantherFolder) {
    Get-ChildItem $PantherFolder -Recurse -ErrorAction SilentlyContinue | Out-File "$ReportFolder\panther_files.txt"
}

# === Registry Exports ===
reg export "HKLM\SYSTEM" "$ReportFolder\HKLM_SYSTEM.reg" /y
reg export "HKLM\SOFTWARE" "$ReportFolder\HKLM_SOFTWARE.reg" /y

# === Run VBScript (GetEvents) ===
$vbScript = @"
On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where Logfile = 'System'",,48)
Set objFile = objFSO.CreateTextFile("$ReportFolder\\GetEvents.log", True)

For Each objItem in colItems
    msg = ""
    If Not IsNull(objItem.Message) Then msg = objItem.Message
    objFile.WriteLine "Event ID: " & objItem.EventCode & " | " & objItem.SourceName & " | " & msg
Next

objFile.Close
"@

$vbPath = Join-Path $ReportFolder "GetEvents.vbs"
$vbScript | Set-Content -Path $vbPath -Encoding ASCII
cscript.exe //nologo "$vbPath"

# === Done ===
Write-Host "`n✅ Report completed. Output folder:" -ForegroundColor Green
Write-Host $ReportFolder -ForegroundColor Cyan
Start-Process "explorer.exe" $ReportFolder

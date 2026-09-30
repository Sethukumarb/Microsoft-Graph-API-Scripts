#Requires -Version 5.1
<#
.SYNOPSIS
    Retrieves Windows devices from Microsoft Intune categorized by days since last sync,
    and exports them into a 2-sheet Excel workbook (Summary + All Devices).

.DESCRIPTION
    This script connects to Microsoft Intune via the Microsoft Graph API using
    App Registration credentials (Client ID, Client Secret, Tenant ID).

    It fetches all Windows-managed devices and calculates how many days have passed
    since each device last synced with Intune (using the 'lastSyncDateTime' field).

    Devices are categorized into 10 aging buckets. Every device appears on the single
    'All Devices' sheet with an AgeBucket column; the 'Summary' sheet holds the counts:

        Bucket 01 -   0 to   7 days : Active - recently synced, healthy devices.
        Bucket 02 -   8 to  30 days : Normal - synced within a month, no action needed.
        Bucket 03 -  31 to  60 days : Watch - slightly overdue, monitor closely.
        Bucket 04 -  61 to  90 days : Aging - may have connectivity or policy issues.
        Bucket 05 -  91 to 120 days : Concern - approaching warning threshold.
        Bucket 06 - 121 to 150 days : Warning - action recommended.
        Bucket 07 - 151 to 180 days : Elevated - escalate for investigation.
        Bucket 08 - 181 to 240 days : High Risk - likely offline or abandoned.
        Bucket 09 - 241 to 365 days : Critical - cleanup or recovery required.
        Bucket 10 - 366 days and above : Stale - remove from Intune immediately.

    Output file is saved in the same folder where this script is located.
    A full transcript log is saved alongside the Excel file.

    NO Autopilot device data is fetched.
    NO Azure AD / Entra ID device data is fetched.
    ONLY Intune-managed Windows device objects are queried.

    v2.3 output layout: ONE 'All Devices' sheet with full details (sorted stalest first) and
    ONE 'Summary' sheet (overview, age buckets, join type, MDM, compliance). Requires
    ImportExcel 7.0 or later (Open-ExcelPackage -Create).

    v2.2 adds two report columns: JoinType (Entra Joined / Hybrid / Registered) and MDM
    (friendly management authority derived from managementAgent).

    v2.1 hardening:
        - Strict-mode safe: missing '@odata.nextLink' on the last page no longer throws.
        - Access token is refreshed automatically before expiry and on HTTP 401.
        - Retry with backoff on HTTP 429 / 500 / 502 / 503 / 504 (honors Retry-After).
        - Never-synced devices (0001-01-01 sentinel) correctly shown as "Never Synced".

.OUTPUTS
    - IntuneStaleWindowsDevices_<timestamp>.xlsx  (2 sheets: Summary, All Devices)
    - IntuneStaleDevices_Transcript_<timestamp>.log  (PowerShell transcript)

.NOTES
    Script Name   : Get-IntuneStaleWindowsDevices.ps1
    Author        : Sethu Kumar B
    Version       : 1.2
    Created       : 2026-05-13
    Last Modified : 2026-09-30 (v1.2 - 2-sheet output; v1.0 - added JoinType and MDM columns; v1 - strict-mode fix, token refresh, retry, never-synced fix)

    Required PowerShell Module : ImportExcel
    Install if missing         : Install-Module ImportExcel -Scope CurrentUser

    Required Graph API App Permissions (Application type):
        - DeviceManagementManagedDevices.Read.All

.EXAMPLE
    .\Get-IntuneStaleWindowsDevices.ps1
    Runs the script. Output files are saved to the same folder as the script.
#>

# ============================================================
# CONFIGURATION BLOCK - Update these values before running
# ============================================================

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

# Graph API version. 'beta' is required because the joinType property is not returned by v1.0.
# Set to 'v1.0' to run on v1.0 only - the JoinType column will then show 'Not Returned'.
$GraphVersion = "beta"

# Retry / token settings
$MaxRetries          = 6      # Max retry attempts per Graph call
$BaseDelaySeconds    = 5      # Initial backoff delay (doubles each retry, capped at 120s)
$TokenRefreshMinutes = 5      # Refresh token when fewer than this many minutes remain

# ============================================================

# Make the script independent of strict-mode settings inherited from profile/ISE/parent
Set-StrictMode -Off

# -------------------------------------------------------
# OUTPUT PATH SETUP
# -------------------------------------------------------
$ScriptRoot = $PSScriptRoot
$Timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ExcelFile  = Join-Path $ScriptRoot "IntuneStaleWindowsDevices_$Timestamp.xlsx"
$LogFile    = Join-Path $ScriptRoot "IntuneStaleDevices_Transcript_$Timestamp.log"

# -------------------------------------------------------
# START TRANSCRIPT
# -------------------------------------------------------
Start-Transcript -Path $LogFile -Append
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "  Intune Stale Windows Device Report" -ForegroundColor Cyan
Write-Host "  Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host ""

# -------------------------------------------------------
# HELPER FUNCTIONS
# -------------------------------------------------------

# Requests a new access token and stores it in script-scope variables.
function Get-GraphToken {
    $TokenUrl  = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    $TokenBody = @{
        grant_type    = "client_credentials"
        client_id     = $ClientID
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    $TokenResponse = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $TokenBody -ErrorAction Stop

    $script:AccessToken = $TokenResponse.access_token
    $script:AuthHeaders = @{ Authorization = "Bearer $($script:AccessToken)" }
    $script:TokenExpiry = (Get-Date).AddSeconds([int]$TokenResponse.expires_in)
}

# Refreshes the token if it is close to expiry.
function Update-GraphTokenIfNeeded {
    if ((Get-Date) -ge $script:TokenExpiry.AddMinutes(-$TokenRefreshMinutes)) {
        Write-Host "  [INFO] Access token near expiry - refreshing..." -ForegroundColor DarkYellow
        Get-GraphToken
        Write-Host "  [OK] Token refreshed." -ForegroundColor Green
    }
}

# GET request with token refresh, 401 re-auth and retry/backoff on throttling or transient errors.
function Invoke-GraphGet {
    param([Parameter(Mandatory = $true)][string]$Uri)

    $Attempt = 0
    while ($true) {
        $Attempt++
        Update-GraphTokenIfNeeded

        try {
            return Invoke-RestMethod -Uri $Uri -Headers $script:AuthHeaders -Method Get -ErrorAction Stop
        }
        catch {
            $StatusCode = $null
            $RetryAfter = $null

            if ($_.Exception.Response) {
                try { $StatusCode = [int]$_.Exception.Response.StatusCode } catch { }
                try { $RetryAfter = $_.Exception.Response.Headers["Retry-After"] } catch { }
            }

            if ($Attempt -gt $MaxRetries) {
                throw
            }

            if ($StatusCode -eq 401) {
                Write-Host "  [WARN] HTTP 401 - re-authenticating (attempt $Attempt/$MaxRetries)..." -ForegroundColor DarkYellow
                Get-GraphToken
                continue
            }

            if ($StatusCode -in 429, 500, 502, 503, 504) {
                $Delay = [Math]::Min($BaseDelaySeconds * [Math]::Pow(2, $Attempt - 1), 120)
                $ParsedRetry = 0
                if ($RetryAfter -and [int]::TryParse([string]$RetryAfter, [ref]$ParsedRetry) -and $ParsedRetry -gt 0) {
                    $Delay = $ParsedRetry
                }
                Write-Host ("  [WARN] HTTP {0} - waiting {1}s before retry (attempt {2}/{3})..." -f $StatusCode, $Delay, $Attempt, $MaxRetries) -ForegroundColor DarkYellow
                Start-Sleep -Seconds $Delay
                continue
            }

            # Non-retryable error
            throw
        }
    }
}

# -------------------------------------------------------
# STEP 1 - CHECK IMPORTEXCEL MODULE
# -------------------------------------------------------
Write-Host "[STEP 1/6] Checking for ImportExcel module..." -ForegroundColor Yellow

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "  [!] ImportExcel module not found. Attempting to install..." -ForegroundColor Red
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "  [OK] ImportExcel installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "  [FATAL] Failed to install ImportExcel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Please run: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Red
        Stop-Transcript
        exit 1
    }
}
else {
    Write-Host "  [OK] ImportExcel module found." -ForegroundColor Green
}

Import-Module ImportExcel -ErrorAction Stop

# -------------------------------------------------------
# STEP 2 - AUTHENTICATE TO GRAPH API
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 2/6] Authenticating to Microsoft Graph API..." -ForegroundColor Yellow

try {
    Get-GraphToken
    Write-Host "  [OK] Authentication successful. Token expires at $($script:TokenExpiry.ToString('HH:mm:ss'))." -ForegroundColor Green
}
catch {
    Write-Host "  [FATAL] Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript
    exit 1
}

# -------------------------------------------------------
# STEP 3 - FETCH WINDOWS MANAGED DEVICES FROM INTUNE
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 3/6] Fetching Windows managed devices from Intune..." -ForegroundColor Yellow
Write-Host "  Filter : operatingSystem eq 'Windows' (applied server-side)" -ForegroundColor DarkGray
Write-Host "  Graph  : $GraphVersion" -ForegroundColor DarkGray
Write-Host "  This may take several minutes for large environments..." -ForegroundColor DarkGray

# NOTE: $filter=operatingSystem eq 'Windows' is safe to use WITHOUT $select on managedDevices.
# Do NOT combine $filter + $select on this endpoint - Graph API returns HTTP 400.
$GraphUri       = "https://graph.microsoft.com/$GraphVersion/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$top=999"
$WindowsDevices = [System.Collections.Generic.List[object]]::new()
$PageCount      = 0

try {
    do {
        $PageCount++
        $Response = Invoke-GraphGet -Uri $GraphUri

        # Strict-mode safe property access
        if ($Response.PSObject.Properties['value'] -and $Response.value) {
            foreach ($Device in $Response.value) {
                $WindowsDevices.Add($Device)
            }
        }

        # Last page has no '@odata.nextLink' - treat as end of data
        if ($Response.PSObject.Properties['@odata.nextLink']) {
            $GraphUri = [string]$Response.'@odata.nextLink'
        }
        else {
            $GraphUri = $null
        }

        Write-Host ("  [{0}] Page {1,3} fetched - Running total: {2} Windows devices" -f (Get-Date -Format 'HH:mm:ss'), $PageCount, $WindowsDevices.Count) -ForegroundColor DarkGray

    } while (-not [string]::IsNullOrWhiteSpace($GraphUri))

    Write-Host "  [OK] Fetch complete. Total Windows devices retrieved: $($WindowsDevices.Count)" -ForegroundColor Green
}
catch {
    Write-Host "  [FATAL] Failed to fetch devices: $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript
    exit 1
}

if ($WindowsDevices.Count -eq 0) {
    Write-Host "  [WARN] No Windows devices found. Exiting." -ForegroundColor Red
    Stop-Transcript
    exit 0
}

# -------------------------------------------------------
# STEP 4 - CALCULATE DAYS SINCE LAST SYNC & CATEGORIZE
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 4/6] Calculating days since last sync and categorizing into 10 buckets..." -ForegroundColor Yellow

$Today = Get-Date

# 10 aging buckets: Label, MinDays, MaxDays, Summary row color
$Buckets = @(
    [PSCustomObject]@{ Label="01 - Active (0-7 Days)";        Min=0;   Max=7;   Color="#C6EFCE" }
    [PSCustomObject]@{ Label="02 - Normal (8-30 Days)";       Min=8;   Max=30;  Color="#E2EFDA" }
    [PSCustomObject]@{ Label="03 - Watch (31-60 Days)";       Min=31;  Max=60;  Color="#FFF2CC" }
    [PSCustomObject]@{ Label="04 - Aging (61-90 Days)";       Min=61;  Max=90;  Color="#FFE699" }
    [PSCustomObject]@{ Label="05 - Concern (91-120 Days)";    Min=91;  Max=120; Color="#FCE4D6" }
    [PSCustomObject]@{ Label="06 - Warning (121-150 Days)";   Min=121; Max=150; Color="#F8CBAD" }
    [PSCustomObject]@{ Label="07 - Elevated (151-180 Days)";  Min=151; Max=180; Color="#F4B084" }
    [PSCustomObject]@{ Label="08 - High Risk (181-240 Days)"; Min=181; Max=240; Color="#FF9999" }
    [PSCustomObject]@{ Label="09 - Critical (241-365 Days)";  Min=241; Max=365; Color="#FF7C80" }
    [PSCustomObject]@{ Label="10 - Stale (366+ Days)";        Min=366; Max=[int]::MaxValue; Color="#FF5050" }
)

# Friendly names for joinType (beta) values
$JoinTypeMap = @{
    "azureADJoined"       = "Entra Joined"
    "hybridAzureADJoined" = "Hybrid Entra Joined"
    "azureADRegistered"   = "Entra Registered"
    "unknown"             = "Unknown"
}

# Friendly MDM authority derived from managementAgent
$MdmMap = @{
    "mdm"                              = "Microsoft Intune"
    "intuneClient"                     = "Microsoft Intune (PC agent)"
    "configurationManagerClientMdm"    = "Intune + SCCM (Co-managed)"
    "configurationManagerClientMdmEas" = "Intune + SCCM (Co-managed) + EAS"
    "configurationManagerClient"       = "SCCM only"
    "easMdm"                           = "Intune + EAS"
    "easIntuneClient"                  = "Intune (PC agent) + EAS"
    "eas"                              = "Exchange ActiveSync only"
    "msSense"                          = "Defender for Endpoint (MDE)"
    "microsoft365ManagedMdm"           = "Microsoft 365 Managed MDM"
    "unknown"                          = "Unknown"
}

function Add-Count {
    param([hashtable]$Table, [string]$Key)
    if ($Table.ContainsKey($Key)) { $Table[$Key] = $Table[$Key] + 1 }
    else { $Table[$Key] = 1 }
}

$BucketCounts     = @{}
foreach ($B in $Buckets) { $BucketCounts[$B.Label] = 0 }
$JoinCounts       = @{}
$MdmCounts        = @{}
$ComplianceCounts = @{}
$AllRows          = [System.Collections.Generic.List[object]]::new()

$SkippedNoSync = 0
$NeverSynced   = 0
$Over30        = 0
$Over90        = 0
$Over180       = 0
$Over365       = 0

foreach ($Device in $WindowsDevices) {

    # Parse lastSyncDateTime safely
    $LastSync    = $null
    $RawLastSync = [string]$Device.lastSyncDateTime

    if ([string]::IsNullOrWhiteSpace($RawLastSync) -or $RawLastSync.StartsWith("0001-01-01")) {
        # Never synced (empty or Graph sentinel 0001-01-01) - route to Stale bucket (366+)
        $LastSync = [datetime]::MinValue
        $NeverSynced++
    }
    else {
        try {
            $LastSync = [datetime]::Parse($RawLastSync)
        }
        catch {
            $SkippedNoSync++
            continue
        }

        # Extra guard for any other year-1 style values after timezone conversion
        if ($LastSync.Year -le 1) {
            $LastSync = [datetime]::MinValue
            $NeverSynced++
        }
    }

    $DaysSinceSync = [int]($Today - $LastSync).TotalDays
    if ($DaysSinceSync -lt 0) { $DaysSinceSync = 0 }

    # Find matching bucket
    $BucketLabel = "Unassigned"
    foreach ($B in $Buckets) {
        if ($DaysSinceSync -ge $B.Min -and $DaysSinceSync -le $B.Max) {
            $BucketLabel = $B.Label
            break
        }
    }

    # Resolve JoinType (strict-mode safe) and MDM authority
    $JoinTypeRaw = "Not Returned"
    if ($Device.PSObject.Properties['joinType'] -and -not [string]::IsNullOrWhiteSpace([string]$Device.joinType)) {
        $JoinTypeRaw = [string]$Device.joinType
    }
    $JoinTypeName = $JoinTypeRaw
    if ($JoinTypeMap.ContainsKey($JoinTypeRaw)) { $JoinTypeName = $JoinTypeMap[$JoinTypeRaw] }

    $AgentRaw = [string]$Device.managementAgent
    $MdmName  = $AgentRaw
    if ($MdmMap.ContainsKey($AgentRaw)) { $MdmName = $MdmMap[$AgentRaw] }
    if ([string]::IsNullOrWhiteSpace($MdmName)) { $MdmName = "Unknown" }

    $ComplianceRaw = [string]$Device.complianceState
    if ([string]::IsNullOrWhiteSpace($ComplianceRaw)) { $ComplianceRaw = "unknown" }

    $LastSyncText = "Never Synced"
    if ($LastSync -ne [datetime]::MinValue) { $LastSyncText = $LastSync.ToString("yyyy-MM-dd HH:mm:ss") }

    # Build row object (all details for the single 'All Devices' sheet)
    $Row = [PSCustomObject]@{
        DeviceName        = [string]$Device.deviceName
        SerialNumber      = [string]$Device.serialNumber
        OperatingSystem   = [string]$Device.operatingSystem
        OSVersion         = [string]$Device.osVersion
        LastSyncDateTime  = $LastSyncText
        DaysSinceSync     = $DaysSinceSync
        AgeBucket         = $BucketLabel
        UserPrincipalName = [string]$Device.userPrincipalName
        EnrollmentDate    = [string]$Device.enrolledDateTime
        ComplianceState   = $ComplianceRaw
        JoinType          = $JoinTypeName
        MDM               = $MdmName
        ManagementAgent   = $AgentRaw
        IntuneDeviceId    = [string]$Device.id
    }
    $AllRows.Add($Row)

    # Summary counters
    Add-Count -Table $BucketCounts -Key $BucketLabel
    Add-Count -Table $JoinCounts -Key $JoinTypeName
    Add-Count -Table $MdmCounts -Key $MdmName
    Add-Count -Table $ComplianceCounts -Key $ComplianceRaw
    if ($DaysSinceSync -ge 31)  { $Over30++ }
    if ($DaysSinceSync -ge 91)  { $Over90++ }
    if ($DaysSinceSync -ge 181) { $Over180++ }
    if ($DaysSinceSync -ge 366) { $Over365++ }
}

$TotalDevices = $AllRows.Count

Write-Host "  [OK] Categorization complete:" -ForegroundColor Green
foreach ($B in $Buckets) {
    Write-Host ("       {0,-40} : {1} devices" -f $B.Label, $BucketCounts[$B.Label]) -ForegroundColor DarkGray
}
Write-Host "       Never synced (in Stale bucket)        : $NeverSynced" -ForegroundColor DarkGray
if ($SkippedNoSync -gt 0) {
    Write-Host "       Skipped (sync date parse error)       : $SkippedNoSync" -ForegroundColor DarkGray
}

if ($TotalDevices -eq 0) {
    Write-Host "  [WARN] No devices to report after parsing. Exiting." -ForegroundColor Red
    Stop-Transcript
    exit 0
}

# -------------------------------------------------------
# SUMMARY SHEET HELPERS
# -------------------------------------------------------
Add-Type -AssemblyName System.Drawing

# Writes one titled, styled table on the Summary worksheet. Returns the next free row.
function Write-SummaryBlock {
    param(
        $Ws,
        [int]$Row,
        [string]$Title,
        [string[]]$Headers,
        $Data,
        [string[]]$RowColors
    )

    $Dark     = [System.Drawing.ColorTranslator]::FromHtml("#1F4E78")
    $White    = [System.Drawing.Color]::White
    $Solid    = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $Thin     = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $Center   = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $ColCount = $Headers.Count

    # Section title
    $Ws.Cells[$Row, 1].Value = $Title
    $Ws.Cells[$Row, 1].Style.Font.Bold = $true
    $Ws.Cells[$Row, 1].Style.Font.Size = 12
    $Ws.Cells[$Row, 1].Style.Font.Color.SetColor($Dark)
    $Row++

    # Header row
    $HeaderRow = $Row
    for ($c = 1; $c -le $ColCount; $c++) {
        $Cell = $Ws.Cells[$Row, $c]
        $Cell.Value = $Headers[$c - 1]
        $Cell.Style.Font.Bold = $true
        $Cell.Style.Font.Color.SetColor($White)
        $Cell.Style.Fill.PatternType = $Solid
        $Cell.Style.Fill.BackgroundColor.SetColor($Dark)
        if ($c -gt 1) { $Cell.Style.HorizontalAlignment = $Center }
    }
    $Row++

    # Data rows
    $Index = 0
    foreach ($Item in $Data) {
        for ($c = 1; $c -le $ColCount; $c++) {
            $Cell = $Ws.Cells[$Row, $c]
            $Cell.Value = $Item[$c - 1]
            if ($Headers[$c - 1] -eq "Percent") { $Cell.Style.Numberformat.Format = "0.0%" }
            if ($c -gt 1) { $Cell.Style.HorizontalAlignment = $Center }
            if ($RowColors -and $Index -lt $RowColors.Count -and $RowColors[$Index]) {
                $Cell.Style.Fill.PatternType = $Solid
                $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($RowColors[$Index]))
            }
        }
        if ([string]$Item[0] -eq "TOTAL") {
            $Ws.Cells[$Row, 1, $Row, $ColCount].Style.Font.Bold = $true
        }
        $Row++
        $Index++
    }

    # Borders around the table
    $LastRow = $Row - 1
    $Range = $Ws.Cells[$HeaderRow, 1, $LastRow, $ColCount]
    $Range.Style.Border.Top.Style    = $Thin
    $Range.Style.Border.Bottom.Style = $Thin
    $Range.Style.Border.Left.Style   = $Thin
    $Range.Style.Border.Right.Style  = $Thin
    $Range.Style.Border.Top.Color.SetColor([System.Drawing.Color]::Gray)
    $Range.Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::Gray)
    $Range.Style.Border.Left.Color.SetColor([System.Drawing.Color]::Gray)
    $Range.Style.Border.Right.Color.SetColor([System.Drawing.Color]::Gray)

    return ($Row + 1)
}

# Turns a count hashtable into rows of Category / Devices / Percent (sorted desc) plus a TOTAL row.
function New-CountData {
    param([hashtable]$Counts, [int]$Total)
    $List = [System.Collections.Generic.List[object]]::new()
    foreach ($Entry in ($Counts.GetEnumerator() | Sort-Object Value -Descending)) {
        $Pct = 0.0
        if ($Total -gt 0) { $Pct = [double]$Entry.Value / $Total }
        $List.Add(@([string]$Entry.Key, [int]$Entry.Value, $Pct))
    }
    $List.Add(@("TOTAL", $Total, 1.0))
    return , $List
}

# -------------------------------------------------------
# STEP 5 - BUILD SUMMARY SHEET
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 5/6] Building Summary sheet..." -ForegroundColor Yellow

if (Test-Path $ExcelFile) { Remove-Item $ExcelFile -Force }

try {
    $Pkg = Open-ExcelPackage -Path $ExcelFile -Create
    $Ws  = Add-Worksheet -ExcelPackage $Pkg -WorksheetName "Summary"
    $Ws.View.ShowGridLines = $false
    $Ws.Column(1).Width = 40
    $Ws.Column(2).Width = 24
    $Ws.Column(3).Width = 14

    # Report title
    $Ws.Cells[1, 1].Value = "Intune Windows Devices - Sync Staleness Summary"
    $Ws.Cells[1, 1].Style.Font.Bold = $true
    $Ws.Cells[1, 1].Style.Font.Size = 16
    $Ws.Cells[1, 1].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#1F4E78"))

    $NextRow = 3

    # Overview
    $Overview = [System.Collections.Generic.List[object]]::new()
    $Overview.Add(@("Report generated", $Today.ToString("yyyy-MM-dd HH:mm:ss")))
    $Overview.Add(@("Graph API version", $GraphVersion))
    $Overview.Add(@("Total Windows devices", $TotalDevices))
    $Overview.Add(@("Never synced", $NeverSynced))
    $Overview.Add(@("Not synced 31+ days", $Over30))
    $Overview.Add(@("Not synced 91+ days", $Over90))
    $Overview.Add(@("Not synced 181+ days", $Over180))
    $Overview.Add(@("Not synced 366+ days", $Over365))
    $Overview.Add(@("Skipped (sync date parse error)", $SkippedNoSync))
    $NextRow = Write-SummaryBlock -Ws $Ws -Row $NextRow -Title "Overview" -Headers @("Metric", "Value") -Data $Overview

    # Age buckets (in bucket order, colored)
    $BucketData   = [System.Collections.Generic.List[object]]::new()
    $BucketColors = [System.Collections.Generic.List[string]]::new()
    foreach ($B in $Buckets) {
        $Cnt = [int]$BucketCounts[$B.Label]
        $BucketData.Add(@($B.Label, $Cnt, ([double]$Cnt / $TotalDevices)))
        $BucketColors.Add($B.Color)
    }
    $BucketData.Add(@("TOTAL", $TotalDevices, 1.0))
    $BucketColors.Add("")
    $NextRow = Write-SummaryBlock -Ws $Ws -Row $NextRow -Title "Devices by Days Since Last Sync" -Headers @("Age Bucket", "Devices", "Percent") -Data $BucketData -RowColors $BucketColors.ToArray()

    # Join type
    $JoinData = New-CountData -Counts $JoinCounts -Total $TotalDevices
    $NextRow  = Write-SummaryBlock -Ws $Ws -Row $NextRow -Title "Devices by Join Type" -Headers @("Join Type", "Devices", "Percent") -Data $JoinData

    # MDM
    $MdmData = New-CountData -Counts $MdmCounts -Total $TotalDevices
    $NextRow = Write-SummaryBlock -Ws $Ws -Row $NextRow -Title "Devices by MDM" -Headers @("MDM", "Devices", "Percent") -Data $MdmData

    # Compliance
    $CompData = New-CountData -Counts $ComplianceCounts -Total $TotalDevices
    $NextRow  = Write-SummaryBlock -Ws $Ws -Row $NextRow -Title "Devices by Compliance State" -Headers @("Compliance State", "Devices", "Percent") -Data $CompData

    Close-ExcelPackage -ExcelPackage $Pkg
    Write-Host "  [OK] Summary sheet built." -ForegroundColor Green
}
catch {
    Write-Host "  [FATAL] Failed to build Summary sheet: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Note: this step needs ImportExcel 7.0 or later. Run: Update-Module ImportExcel" -ForegroundColor Red
    Stop-Transcript
    exit 1
}

# -------------------------------------------------------
# STEP 6 - EXPORT ALL DEVICES (SINGLE SHEET)
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 6/6] Writing 'All Devices' sheet ($TotalDevices rows)..." -ForegroundColor Yellow

try {
    # Stalest devices first
    $AllRows | Sort-Object DaysSinceSync -Descending |
        Export-Excel -Path $ExcelFile -WorksheetName "All Devices" `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
            -TableName "TblAllDevices" -TableStyle "Medium2"

    Write-Host "  [OK] Excel file saved successfully." -ForegroundColor Green
}
catch {
    Write-Host "  [FATAL] Failed to write Excel file: $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript
    exit 1
}

# -------------------------------------------------------
# SUMMARY
# -------------------------------------------------------
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "  REPORT COMPLETE" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "  Excel File : $ExcelFile" -ForegroundColor White
Write-Host "  Log File   : $LogFile" -ForegroundColor White
Write-Host "  Sheets     : Summary, All Devices" -ForegroundColor White
Write-Host ""
Write-Host "  DEVICE COUNT SUMMARY:" -ForegroundColor Cyan
Write-Host ("  {0,-40} {1}" -f "Total Windows Devices Fetched", $WindowsDevices.Count) -ForegroundColor White
Write-Host ("  {0,-40} {1}" -f "Devices in report", $TotalDevices) -ForegroundColor White
foreach ($B in $Buckets) {
    Write-Host ("  {0,-40} {1}" -f $B.Label, $BucketCounts[$B.Label]) -ForegroundColor White
}
Write-Host ("  {0,-40} {1}" -f "Never synced (in Stale bucket)", $NeverSynced) -ForegroundColor DarkGray
if ($SkippedNoSync -gt 0) {
    Write-Host ("  {0,-40} {1}" -f "Skipped (sync date parse error)", $SkippedNoSync) -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "  Completed : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host ""

Stop-Transcript
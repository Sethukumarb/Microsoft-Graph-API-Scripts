#Requires -Version 5.1
<#
.SYNOPSIS
    Retrieves Windows devices from Microsoft Intune categorized by days since last sync,
    and exports them into a 10-sheet Excel workbook for full staleness visibility.

.DESCRIPTION
    This script connects to Microsoft Intune via the Microsoft Graph API using
    App Registration credentials (Client ID, Client Secret, Tenant ID).

    It fetches all Windows-managed devices and calculates how many days have passed
    since each device last synced with Intune (using the 'lastSyncDateTime' field).

    Devices are categorized into 10 aging buckets across 10 worksheet tabs:

        Sheet 01 -   0 to   7 days : Active - recently synced, healthy devices.
        Sheet 02 -   8 to  30 days : Normal - synced within a month, no action needed.
        Sheet 03 -  31 to  60 days : Watch - slightly overdue, monitor closely.
        Sheet 04 -  61 to  90 days : Aging - may have connectivity or policy issues.
        Sheet 05 -  91 to 120 days : Concern - approaching warning threshold.
        Sheet 06 - 121 to 150 days : Warning - action recommended.
        Sheet 07 - 151 to 180 days : Elevated - escalate for investigation.
        Sheet 08 - 181 to 240 days : High Risk - likely offline or abandoned.
        Sheet 09 - 241 to 365 days : Critical - cleanup or recovery required.
        Sheet 10 - 366 days and above : Stale - remove from Intune immediately.

    Output file is saved in the same folder where this script is located.
    A full transcript log is saved alongside the Excel file.

    NO Autopilot device data is fetched.
    NO Azure AD / Entra ID device data is fetched.
    ONLY Intune-managed Windows device objects are queried.

    v2.2 adds two report columns: JoinType (Entra Joined / Hybrid / Registered) and MDM
    (friendly management authority derived from managementAgent).

    v2.1 hardening:
        - Strict-mode safe: missing '@odata.nextLink' on the last page no longer throws.
        - Access token is refreshed automatically before expiry and on HTTP 401.
        - Retry with backoff on HTTP 429 / 500 / 502 / 503 / 504 (honors Retry-After).
        - Never-synced devices (0001-01-01 sentinel) correctly shown as "Never Synced".

.OUTPUTS
    - IntuneStaleWindowsDevices_<timestamp>.xlsx  (10-sheet Excel workbook)
    - IntuneStaleDevices_Transcript_<timestamp>.log  (PowerShell transcript)

.NOTES
    Script Name   : Get-IntuneStaleWindowsDevices.ps1
    Author        : Sethu Kumar B
    Version       : 1.0
    Created       : 2026-05-13
    Last Modified : 2026-09-30 (v1.0 - added JoinType and MDM columns; v2.1 - strict-mode fix, token refresh, retry, never-synced fix)

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

# 10 aging buckets - each is a Generic List for performance
# Bucket definition: SheetName, MinDays, MaxDays, TableName, TableStyle, Title
$Buckets = @(
    [PSCustomObject]@{ SheetName="01 - Active (0-7 Days)";        Min=0;   Max=7;   Table="Tbl01"; Style="Medium1"; Title="ACTIVE - Devices Synced Within the Last 7 Days" }
    [PSCustomObject]@{ SheetName="02 - Normal (8-30 Days)";       Min=8;   Max=30;  Table="Tbl02"; Style="Medium1"; Title="NORMAL - Devices Not Synced for 8 to 30 Days" }
    [PSCustomObject]@{ SheetName="03 - Watch (31-60 Days)";       Min=31;  Max=60;  Table="Tbl03"; Style="Medium2"; Title="WATCH - Devices Not Synced for 31 to 60 Days" }
    [PSCustomObject]@{ SheetName="04 - Aging (61-90 Days)";       Min=61;  Max=90;  Table="Tbl04"; Style="Medium2"; Title="AGING - Devices Not Synced for 61 to 90 Days" }
    [PSCustomObject]@{ SheetName="05 - Concern (91-120 Days)";    Min=91;  Max=120; Table="Tbl05"; Style="Medium3"; Title="CONCERN - Devices Not Synced for 91 to 120 Days" }
    [PSCustomObject]@{ SheetName="06 - Warning (121-150 Days)";   Min=121; Max=150; Table="Tbl06"; Style="Medium3"; Title="WARNING - Devices Not Synced for 121 to 150 Days" }
    [PSCustomObject]@{ SheetName="07 - Elevated (151-180 Days)";  Min=151; Max=180; Table="Tbl07"; Style="Medium4"; Title="ELEVATED - Devices Not Synced for 151 to 180 Days" }
    [PSCustomObject]@{ SheetName="08 - High Risk (181-240 Days)"; Min=181; Max=240; Table="Tbl08"; Style="Medium4"; Title="HIGH RISK - Devices Not Synced for 181 to 240 Days" }
    [PSCustomObject]@{ SheetName="09 - Critical (241-365 Days)";  Min=241; Max=365; Table="Tbl09"; Style="Medium6"; Title="CRITICAL - Devices Not Synced for 241 to 365 Days" }
    [PSCustomObject]@{ SheetName="10 - Stale (366+ Days)";        Min=366; Max=[int]::MaxValue; Table="Tbl10"; Style="Medium6"; Title="STALE - Devices Not Synced for 366 Days or More" }
)

# Initialize a data list for each bucket
$BucketData = @{}
foreach ($B in $Buckets) {
    $BucketData[$B.SheetName] = [System.Collections.Generic.List[object]]::new()
}

$SkippedNoSync = 0
$NeverSynced   = 0

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

    # Resolve JoinType (strict-mode safe) and MDM authority
    $JoinTypeRaw = "Not Returned"
    if ($Device.PSObject.Properties['joinType'] -and -not [string]::IsNullOrWhiteSpace([string]$Device.joinType)) {
        $JoinTypeRaw = [string]$Device.joinType
    }
    $JoinTypeName = if ($JoinTypeMap.ContainsKey($JoinTypeRaw)) { $JoinTypeMap[$JoinTypeRaw] } else { $JoinTypeRaw }

    $AgentRaw = [string]$Device.managementAgent
    $MdmName  = if ($MdmMap.ContainsKey($AgentRaw)) { $MdmMap[$AgentRaw] } else { $AgentRaw }

    # Build row object
    $Row = [PSCustomObject]@{
        DeviceName        = [string]$Device.deviceName
        SerialNumber      = [string]$Device.serialNumber
        OperatingSystem   = [string]$Device.operatingSystem
        OSVersion         = [string]$Device.osVersion
        LastSyncDateTime  = if ($LastSync -eq [datetime]::MinValue) { "Never Synced" } else { $LastSync.ToString("yyyy-MM-dd HH:mm:ss") }
        DaysSinceSync     = $DaysSinceSync
        UserPrincipalName = [string]$Device.userPrincipalName
        EnrollmentDate    = [string]$Device.enrolledDateTime
        ComplianceState   = [string]$Device.complianceState
        JoinType          = $JoinTypeName
        MDM               = $MdmName
        ManagementAgent   = $AgentRaw
        IntuneDeviceId    = [string]$Device.id
    }

    # Find matching bucket and add row
    foreach ($B in $Buckets) {
        if ($DaysSinceSync -ge $B.Min -and $DaysSinceSync -le $B.Max) {
            $BucketData[$B.SheetName].Add($Row)
            break
        }
    }
}

Write-Host "  [OK] Categorization complete:" -ForegroundColor Green
foreach ($B in $Buckets) {
    Write-Host ("       {0,-40} : {1} devices" -f $B.SheetName, $BucketData[$B.SheetName].Count) -ForegroundColor DarkGray
}
Write-Host "       Never synced (in Stale bucket)        : $NeverSynced" -ForegroundColor DarkGray
if ($SkippedNoSync -gt 0) {
    Write-Host "       Skipped (sync date parse error)       : $SkippedNoSync" -ForegroundColor DarkGray
}

# -------------------------------------------------------
# STEP 5 - PREPARE EXCEL OUTPUT
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 5/6] Preparing Excel workbook..." -ForegroundColor Yellow

if (Test-Path $ExcelFile) { Remove-Item $ExcelFile -Force }

# -------------------------------------------------------
# STEP 6 - EXPORT TO EXCEL (10 SHEETS)
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 6/6] Writing Excel report to disk..." -ForegroundColor Yellow

$SheetNumber = 0

try {
    foreach ($B in $Buckets) {
        $SheetNumber++
        $Data = $BucketData[$B.SheetName]
        Write-Host ("  [{0,2}/10] Writing: {1} ({2} devices)..." -f $SheetNumber, $B.SheetName, $Data.Count) -ForegroundColor DarkGray

        if ($Data.Count -gt 0) {
            $Data | Export-Excel -Path $ExcelFile -WorksheetName $B.SheetName `
                -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
                -TableName $B.Table -TableStyle $B.Style `
                -Title $B.Title `
                -TitleBold -TitleSize 12
        }
        else {
            [PSCustomObject]@{ Note = "No devices found in this category." } |
                Export-Excel -Path $ExcelFile -WorksheetName $B.SheetName -AutoSize
        }
    }

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
Write-Host ""
Write-Host "  DEVICE COUNT SUMMARY:" -ForegroundColor Cyan
Write-Host ("  {0,-40} {1}" -f "Total Windows Devices Fetched", $WindowsDevices.Count) -ForegroundColor White
foreach ($B in $Buckets) {
    Write-Host ("  {0,-40} {1}" -f $B.SheetName, $BucketData[$B.SheetName].Count) -ForegroundColor White
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
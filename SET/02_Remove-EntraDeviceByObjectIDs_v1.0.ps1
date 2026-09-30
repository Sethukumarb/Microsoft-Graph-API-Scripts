#Requires -Version 5.1
# ===============================================================================
#  Script      :  Remove-EntraDeviceByObjectIDs.ps1
#  Description :  Bulk-delete Entra ID device objects via Microsoft Graph API
#                 using a list of Entra Device OBJECT IDs (not Device ID, not
#                 Intune Managed Device ID). Touches Entra ONLY - no Intune,
#                 no Autopilot records are removed.
#                 Supports Dry Run mode, batch size cap, and CSV result report.
#  Author      :  Sethu Kumar B
#  Version     :  1.0
#  Permissions :  Device.ReadWrite.All (Application, App Registration)
#  API Ref     :  DELETE https://graph.microsoft.com/v1.0/devices/{id}
#                 https://learn.microsoft.com/en-us/graph/api/device-delete
#                 Success = 204 No Content, no response body.
#                 {id} = Entra device object ID. (Docs also allow the
#                 /devices(deviceId='{deviceId}') form - not used here.)
# ===============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ===============================================================================
#  CONFIGURATION - Fill in before running
# ===============================================================================

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

# Input : one Entra device OBJECT ID (GUID) per line
$InputFile    = "$PSScriptRoot\EntraObjectIDs.txt"
$Stamp        = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile      = "$PSScriptRoot\Logs\Remove-EntraDevice_$Stamp.log"
$ReportFile   = "$PSScriptRoot\Logs\Remove-EntraDevice_$Stamp.csv"

# DRY RUN : $true  = simulate only, no deletions (DEFAULT - keep it on for first run)
#           $false = live run, deletions are permanent
$DryRun       = $false

# BATCH CAP : abort if input has more IDs than this
$MaxBatchSize = 2

# Delay between deletes (ms) to stay gentle on Graph throttling
$DelayMs      = 200
# ===============================================================================


# -------------------------------------------------------------------------------
#  LOGGING
# -------------------------------------------------------------------------------
function Write-Log {
    param(
        [string] $Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR","SECTION")]
        [string] $Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp]  $($Level.PadRight(7))  $Message"
    $colour    = switch ($Level) {
        "SUCCESS" { "Green"  }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red"    }
        "SECTION" { "Cyan"   }
        default   { "White"  }
    }
    Write-Host $entry -ForegroundColor $colour
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8
}

function Write-Section {
    param([string]$Title)
    $line = "-" * 79
    Write-Log $line      "SECTION"
    Write-Log "  $Title" "SECTION"
    Write-Log $line      "SECTION"
}


# -------------------------------------------------------------------------------
#  GRAPH API - TOKEN
# -------------------------------------------------------------------------------
function Get-GraphToken {
    $body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ClientID
        client_secret = $ClientSecret
    }
    try {
        $response = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" `
            -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Log "Token acquisition failed. Verify TenantID, ClientID, ClientSecret." "ERROR"
        Write-Log "Detail : $_" "ERROR"
        exit 1
    }
}

# Graph call with simple retry on 429 / 503 / 504
function Invoke-GraphWithRetry {
    param(
        [string]    $Method,
        [string]    $Uri,
        [hashtable] $Headers
    )
    $maxTries = 4
    for ($try = 1; $try -le $maxTries; $try++) {
        try {
            if ($Method -eq "DELETE") {
                # Invoke-WebRequest so the 204 status code can be verified
                return Invoke-WebRequest -Method DELETE -Uri $Uri -Headers $Headers -UseBasicParsing -ErrorAction Stop
            }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -ErrorAction Stop
        }
        catch {
            $code = $null
            if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
            if (($code -eq 429 -or $code -eq 503 -or $code -eq 504) -and $try -lt $maxTries) {
                $wait = 5 * $try
                $ra   = $_.Exception.Response.Headers["Retry-After"]
                if ($ra) { $wait = [int]$ra }
                Write-Log "HTTP $code - retry $try/$($maxTries - 1) in $wait sec" "WARN"
                Start-Sleep -Seconds $wait
            }
            else { throw }
        }
    }
}


# -------------------------------------------------------------------------------
#  INITIALISE - LOG FOLDER
# -------------------------------------------------------------------------------
$logFolder = Split-Path $LogFile
if (-not (Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}


# -------------------------------------------------------------------------------
#  STARTUP BANNER
# -------------------------------------------------------------------------------
Write-Section "Remove-EntraDeviceByObjectIDs  |  v1.0  |  Sethu Kumar B"
Write-Log "Start Time   : $(Get-Date -Format 'dddd, dd MMMM yyyy  HH:mm:ss')" "INFO"
Write-Log "Input File   : $InputFile"  "INFO"
Write-Log "Log File     : $LogFile"    "INFO"
Write-Log "Report File  : $ReportFile" "INFO"
Write-Log "Batch Cap    : $MaxBatchSize device(s) per run" "INFO"

if ($DryRun) {
    Write-Log "*** DRY RUN MODE - ZERO DELETIONS WILL OCCUR ***" "WARN"
}
else {
    Write-Log "*** LIVE MODE - ENTRA DELETIONS ARE PERMANENT ***" "WARN"
}


# -------------------------------------------------------------------------------
#  STEP 1 - VALIDATE INPUT
# -------------------------------------------------------------------------------
Write-Section "Step 1 of 4 - Validating Input File"

if (-not (Test-Path $InputFile)) {
    Write-Log "Input file not found : $InputFile" "ERROR"
    exit 1
}

$raw = @(Get-Content $InputFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

# GUID format check - catches pasted Device IDs with junk, names, or blanks
$guidRegex = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$valid   = @()
$invalid = @()
foreach ($line in $raw) {
    if ($line -match $guidRegex) { $valid += $line.ToLower() } else { $invalid += $line }
}

if ($invalid.Count -gt 0) {
    Write-Log "$($invalid.Count) line(s) are not valid GUIDs and will be skipped:" "WARN"
    foreach ($i in $invalid) { Write-Log "  Invalid : $i" "WARN" }
}

# Remove duplicates
$DeviceIDs = @($valid | Select-Object -Unique)
$dupes     = $valid.Count - $DeviceIDs.Count
if ($dupes -gt 0) { Write-Log "$dupes duplicate ID(s) removed." "WARN" }

if ($DeviceIDs.Count -eq 0) {
    Write-Log "No valid Object IDs to process." "WARN"
    exit 0
}
Write-Log "Loaded $($DeviceIDs.Count) unique valid Object ID(s)." "SUCCESS"


# -------------------------------------------------------------------------------
#  STEP 2 - BATCH GUARD
# -------------------------------------------------------------------------------
Write-Section "Step 2 of 4 - Batch Size Enforcement"

if ($DeviceIDs.Count -gt $MaxBatchSize) {
    Write-Log "ABORTED - $($DeviceIDs.Count) device(s) in input. Max per run: $MaxBatchSize." "ERROR"
    Write-Log "Split the input file or raise `$MaxBatchSize in the CONFIG block." "ERROR"
    exit 1
}
Write-Log "Batch size OK : $($DeviceIDs.Count) / $MaxBatchSize" "SUCCESS"


# -------------------------------------------------------------------------------
#  STEP 3 - AUTHENTICATE
# -------------------------------------------------------------------------------
Write-Section "Step 3 of 4 - Authenticating to Microsoft Graph"

$Token   = Get-GraphToken
$Headers = @{ Authorization = "Bearer $Token" }
Write-Log "Token acquired." "SUCCESS"


# -------------------------------------------------------------------------------
#  STEP 4 - PROCESS
# -------------------------------------------------------------------------------
Write-Section "Step 4 of 4 - Processing Entra Device Objects"

$Success  = 0
$NotFound = 0
$Failed   = 0
$Counter  = 0
$Report   = New-Object System.Collections.ArrayList

foreach ($ObjectId in $DeviceIDs) {
    $Counter++
    $Prefix = "[$Counter/$($DeviceIDs.Count)]"
    $Uri    = "https://graph.microsoft.com/v1.0/devices/$ObjectId"

    $row = [ordered]@{
        ObjectId    = $ObjectId
        DisplayName = ""
        DeviceId    = ""
        OS          = ""
        TrustType   = ""
        Enabled     = ""
        LastSignIn  = ""
        Result      = ""
        Detail      = ""
    }

    try {
        # Look up first so the log shows exactly what is being deleted
        $dev = Invoke-GraphWithRetry -Method GET -Uri $Uri -Headers $Headers

        $row.DisplayName = $dev.displayName
        $row.DeviceId    = $dev.deviceId
        $row.OS          = $dev.operatingSystem
        $row.TrustType   = $dev.trustType
        $row.Enabled     = $dev.accountEnabled
        $row.LastSignIn  = $dev.approximateLastSignInDateTime

        $info = "$($dev.displayName)  |  $($dev.operatingSystem)  |  $($dev.trustType)  |  DeviceId: $($dev.deviceId)  |  ObjectId: $ObjectId"

        if ($DryRun) {
            Write-Log "$Prefix  DRY RUN - WOULD DELETE  |  $info" "WARN"
            $row.Result = "DryRun"
            $Success++
        }
        else {
            $resp = Invoke-GraphWithRetry -Method DELETE -Uri $Uri -Headers $Headers
            if ($resp.StatusCode -eq 204) {
                Write-Log "$Prefix  DELETED (204)  |  $info" "SUCCESS"
                $row.Result = "Deleted"
                $Success++
            }
            else {
                Write-Log "$Prefix  UNEXPECTED STATUS $($resp.StatusCode) (expected 204)  |  $info" "ERROR"
                $row.Result = "Failed"
                $row.Detail = "Unexpected status $($resp.StatusCode)"
                $Failed++
            }
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    catch {
        $code = $null
        if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
        if ($code -eq 404) {
            Write-Log "$Prefix  NOT FOUND  |  ObjectId: $ObjectId (wrong ID type, or already deleted)" "WARN"
            $row.Result = "NotFound"
            $NotFound++
        }
        else {
            if ($code -eq 403) { Write-Log "HTTP 403 - confirm Device.ReadWrite.All (Application) has admin consent." "WARN" }
            Write-Log "$Prefix  FAILED  |  ObjectId: $ObjectId  |  HTTP $code  |  $_" "ERROR"
            $row.Result = "Failed"
            $row.Detail = "HTTP $code - $_"
            $Failed++
        }
    }
    [void]$Report.Add([pscustomobject]$row)
}


# -------------------------------------------------------------------------------
#  SUMMARY
# -------------------------------------------------------------------------------
Write-Section "Execution Summary"

$Report | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8

$modeLabel = if ($DryRun) { "DRY RUN (simulated)" } else { "LIVE (permanent)" }
$failLevel = if ($Failed -gt 0) { "ERROR" } else { "INFO" }

Write-Log "Mode        : $modeLabel" "INFO"
Write-Log "Total Input : $($DeviceIDs.Count)" "INFO"
if ($DryRun) { Write-Log "Would Delete: $Success" "WARN" }
else         { Write-Log "Deleted     : $Success" "SUCCESS" }
Write-Log "Not Found   : $NotFound" "WARN"
Write-Log "Failed      : $Failed" $failLevel
Write-Log "End Time    : $(Get-Date -Format 'dddd, dd MMMM yyyy  HH:mm:ss')" "INFO"
Write-Log "Log         : $LogFile" "INFO"
Write-Log "Report CSV  : $ReportFile" "INFO"

if ($DryRun) {
    Write-Log "Review the log, then set `$DryRun = `$false to execute live." "WARN"
}

Write-Section "Script Complete"


# -------------------------------------------------------------------------------
#  SECURITY CLEANUP
# -------------------------------------------------------------------------------
$ClientSecret = ""
$Token        = ""
$Headers      = $null
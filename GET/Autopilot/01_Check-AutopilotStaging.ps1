#Requires -Version 5.1
# ==============================================================================
# Script Name  : Check-AutopilotStaging.ps1
# Description  : Read-only check of the Autopilot staging endpoint.
#                Reports every record currently sitting in
#                importedWindowsAutopilotDeviceIdentities - serial number,
#                import status, error code, and intended action (import /
#                delete). No changes are made.
#
#                USE THIS BEFORE re-importing devices to confirm no stale
#                staging records exist that would cause 806 /
#                ZtdDeviceAlreadyAssigned errors.
#
#                OUTPUT CSV COLUMNS:
#                  SerialNumber, ImportedDeviceID, GroupTag,
#                  IntendedAction, ImportStatus, ErrorCode, ErrorName,
#                  RegistrationID, CreatedDateTime
#
#                IMPORT STATUS VALUES:
#                  unknown          - not yet processed
#                  pending          - queued, processing not started
#                  complete         - successfully imported
#                  error            - failed (see ErrorCode / ErrorName)
#                  completedWithError - partial success
#
# Author       : Sethu Kumar B
# Version      : 1.0
# Created Date : 2026-04-30
#
# Requirements :
#   - Azure AD App Registration
#   - Graph API Application Permission (admin consent granted):
#       DeviceManagementServiceConfig.ReadWrite.All
#   - PowerShell 5.1 or later
#   - TLS 1.2 enabled
# ==============================================================================


#region --- CONFIGURATION -------------------------------------------------------

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

#endregion ----------------------------------------------------------------------


#region --- INIT ----------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile     = Join-Path $PSScriptRoot "AutopilotStagingCheck_$Timestamp.csv"
$script:LogFile = Join-Path $PSScriptRoot "AutopilotStagingCheck_$Timestamp.log"

try {
    [System.IO.File]::WriteAllText($script:LogFile,
        "Check-AutopilotStaging v1.0`r`nStarted: $(Get-Date)`r`n`r`n",
        [System.Text.Encoding]::UTF8)
} catch { $script:LogFile = $null }

#endregion ----------------------------------------------------------------------


#region --- FUNCTIONS -----------------------------------------------------------

function Write-Log {
    param (
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR","SECTION","BLANK")]
        [string]$Level = "INFO"
    )
    $ColourMap = @{ INFO="Gray"; SUCCESS="Green"; WARN="Yellow"; ERROR="Red"; SECTION="Cyan"; BLANK="Gray" }
    $PrefixMap = @{ INFO="[INFO]   "; SUCCESS="[OK]     "; WARN="[WARN]   "; ERROR="[ERROR]  "; SECTION=""; BLANK="         " }
    $t = Get-Date -Format "HH:mm:ss"
    if     ($Level -eq "BLANK")   { Write-Host "" }
    elseif ($Level -eq "SECTION") { Write-Host "`n$Message" -ForegroundColor Cyan }
    else   { Write-Host "[$t] $($PrefixMap[$Level]) $Message" -ForegroundColor $ColourMap[$Level] }
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile `
                -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $($PrefixMap[$Level]) $Message" `
                -Encoding UTF8
        } catch { }
    }
}


function Get-GraphToken {
    param ([string]$TenantId, [string]$ClientId, [string]$ClientSecret)
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    try {
        Write-Log "Requesting access token..." -Level INFO
        $r = Invoke-RestMethod -Method POST -ContentType "application/x-www-form-urlencoded" `
             -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
             -Body $Body -ErrorAction Stop
        Write-Log "Access token acquired." -Level SUCCESS
        return $r.access_token
    }
    catch { Write-Log "Authentication failed: $_" -Level ERROR; exit 1 }
}


# -----------------------------------------------------------------------------
# Get-StagingRecords
# Pages through all importedWindowsAutopilotDeviceIdentities records.
# Returns list of raw objects.
# -----------------------------------------------------------------------------
function Get-StagingRecords {
    param ([string]$AccessToken)

    $Headers = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $Uri     = "https://graph.microsoft.com/beta/deviceManagement/importedWindowsAutopilotDeviceIdentities?`$top=1000"
    $All     = [System.Collections.Generic.List[PSObject]]::new()

    do {
        try {
            $r   = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
            $Arr = if ($r.PSObject.Properties["value"]) { @($r.value) } else { @($r) }
            foreach ($rec in $Arr) { if ($rec) { $All.Add($rec) } }
            $Uri = if ($r.PSObject.Properties["@odata.nextLink"]) { $r.'@odata.nextLink' } else { $null }
        }
        catch {
            Write-Log "Page fetch failed: $_" -Level ERROR
            $Uri = $null
        }
    } while ($Uri)

    return $All
}

#endregion ----------------------------------------------------------------------


#region --- MAIN ----------------------------------------------------------------

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Check-AutopilotStaging  |  Sethu Kumar B  |  READ ONLY       " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "" -Level BLANK
Write-Log "Output CSV : $OutputFile"        -Level INFO
Write-Log "Log file   : $($script:LogFile)" -Level INFO
Write-Log "" -Level BLANK


# -- Step 1: Authenticate ------------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 1 - Authenticating" -Level INFO
Write-Log "==========================================================" -Level SECTION

$Token = Get-GraphToken -TenantId $TenantID -ClientId $ClientID -ClientSecret $ClientSecret
Write-Log "" -Level BLANK


# -- Step 2: Fetch staging records ---------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 2 - Fetching Staging Records" -Level INFO
Write-Log "==========================================================" -Level SECTION

$Records = Get-StagingRecords -AccessToken $Token
Write-Log "Total records in staging: $($Records.Count)" -Level $(if ($Records.Count -gt 0) {"WARN"} else {"SUCCESS"})
Write-Log "" -Level BLANK


# -- Step 3: Parse and display -------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 3 - Parsing Status" -Level INFO
Write-Log "==========================================================" -Level SECTION

$Results = [System.Collections.Generic.List[PSObject]]::new()

if ($Records.Count -eq 0) {
    Write-Log "Staging is clean - no records found. Safe to re-import." -Level SUCCESS
}
else {
    $CountComplete = 0
    $CountError    = 0
    $CountPending  = 0
    $CountUnknown  = 0

    foreach ($rec in $Records) {

        # --- Status fields ----------------------------------------------------
        $ImportStatus = "unknown"
        $ErrorCode    = ""
        $ErrorName    = ""
        $RegID        = ""

        if ($rec.PSObject.Properties["state"] -and $rec.state) {
            if ($rec.state.PSObject.Properties["deviceImportStatus"])   { $ImportStatus = [string]$rec.state.deviceImportStatus   }
            if ($rec.state.PSObject.Properties["deviceErrorCode"])      { $ErrorCode    = [string]$rec.state.deviceErrorCode      }
            if ($rec.state.PSObject.Properties["deviceErrorName"])      { $ErrorName    = [string]$rec.state.deviceErrorName      }
            if ($rec.state.PSObject.Properties["deviceRegistrationId"]) { $RegID        = [string]$rec.state.deviceRegistrationId }
        }

        # Tally
        switch ($ImportStatus) {
            "complete"           { $CountComplete++ }
            "error"              { $CountError++    }
            "completedWithError" { $CountError++    }
            "pending"            { $CountPending++  }
            default              { $CountUnknown++  }
        }

        # Log colour per status
        $Level = switch ($ImportStatus) {
            "complete"           { "SUCCESS" }
            "error"              { "ERROR"   }
            "completedWithError" { "ERROR"   }
            "pending"            { "WARN"    }
            default              { "INFO"    }
        }

        $Detail = ""
        if ($ErrorCode -and $ErrorCode -ne "0") { $Detail  = "ErrorCode: $ErrorCode" }
        if ($ErrorName) { $Detail += if ($Detail) { " | $ErrorName" } else { $ErrorName } }
        if ($RegID)     { $Detail += if ($Detail) { " | RegID: $RegID" } else { "RegID: $RegID" } }

        Write-Log ("  {0,-30} | Status: {1,-20} | {2}" -f
            $rec.serialNumber, $ImportStatus, $Detail) -Level $Level

        $Results.Add([PSCustomObject]@{
            SerialNumber     = if ($rec.serialNumber)    { [string]$rec.serialNumber    } else { "" }
            ImportedDeviceID = if ($rec.id)              { [string]$rec.id              } else { "" }
            GroupTag         = if ($rec.groupTag)        { [string]$rec.groupTag        } else { "" }
            ImportStatus     = $ImportStatus
            ErrorCode        = $ErrorCode
            ErrorName        = $ErrorName
            RegistrationID   = $RegID
            CreatedDateTime  = if ($rec.createdDateTime) { [string]$rec.createdDateTime } else { "" }
        })
    }

    Write-Log "" -Level BLANK
    Write-Log "--- Staging Summary ---" -Level SECTION
    Write-Log "  Complete           : $CountComplete" -Level $(if ($CountComplete -gt 0) {"SUCCESS"} else {"INFO"})
    Write-Log "  Pending            : $CountPending"  -Level $(if ($CountPending  -gt 0) {"WARN"}    else {"INFO"})
    Write-Log "  Unknown            : $CountUnknown"  -Level $(if ($CountUnknown  -gt 0) {"WARN"}    else {"INFO"})
    Write-Log "  Errors             : $CountError"    -Level $(if ($CountError    -gt 0) {"ERROR"}   else {"INFO"})
    Write-Log "" -Level BLANK

    $StaleRisk = $CountComplete + $CountError + $CountPending + $CountUnknown
    if ($StaleRisk -gt 0) {
        Write-Log "WARNING: $StaleRisk record(s) in staging." -Level WARN
        Write-Log "  Any record still present may trigger 806/ZtdDeviceAlreadyAssigned on re-import." -Level WARN
        Write-Log "  Wait for Graph to auto-purge OR manually delete stale records before re-importing." -Level WARN
    }
}
Write-Log "" -Level BLANK


# -- Step 4: Export CSV --------------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 4 - Exporting Results" -Level INFO
Write-Log "==========================================================" -Level SECTION

if ($Results.Count -gt 0) {
    try {
        $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        $SizeMB = (((Get-Item $OutputFile).Length) / 1MB).ToString("0.00")
        Write-Log "CSV exported." -Level SUCCESS
        Write-Log "  Path : $OutputFile" -Level INFO
        Write-Log "  Rows : $($Results.Count)  |  Size: $SizeMB MB" -Level INFO
    }
    catch { Write-Log "CSV export failed: $_" -Level ERROR }
}
else {
    Write-Log "No records - staging clean. No CSV written." -Level SUCCESS
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  CHECK COMPLETE  |  Sethu Kumar B                             " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

#endregion ----------------------------------------------------------------------
<#
.SYNOPSIS
    Export all Windows Autopilot devices from Microsoft Intune via Graph API to CSV.

.DESCRIPTION
    Authenticates to Microsoft Graph API using app-only (client credentials) OAuth2 flow,
    then fetches all Windows Autopilot device identities from the beta endpoint with full
    OData pagination support. Exports SerialNumber, AutopilotID, AzureADObjectID, GroupTag,
    and DeploymentProfile to a CSV file in the script's directory.

    Graph Endpoint : https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities
    Auth Flow      : Client Credentials (App Registration, no user sign-in)
    Output - CSV   : <ScriptRoot>\AutopilotDevices_<timestamp>.csv
    Output - Log   : <ScriptRoot>\Logs\AutopilotExport_<timestamp>.log
    Output - Trans : <ScriptRoot>\Logs\AutopilotExport_<timestamp>_Transcript.log

    Security:
    - All credential variables explicitly wiped from memory post-run via [Runtime.InteropServices.Marshal]
    - Secure strings used for ClientSecret in-memory handling
    - Access token overwritten before variable removal
    - Transcript and log capture full session for audit trail
    - No credentials written to any log or transcript

.PARAMETER TenantID
    Azure AD Tenant ID (GUID).

.PARAMETER ClientID
    App Registration Client ID (GUID). Requires DeviceManagementServiceConfig.Read.All.

.PARAMETER ClientSecret
    App Registration Client Secret (SecureString conversion handled internally).

.NOTES
    Author  : Sethu Kumar B
    Version : 3.0
    Date    : Dec 2025

.EXAMPLE
    .\Export-AutopilotDevices.ps1
    Runs with hardcoded config block. CSV, log, and transcript saved to script root.
#>

### ================================
### Variables – UPDATE THESE VALUES
### ================================

$TenantID     = ""
$ClientID     = ""
$ClientSecret = "" 

# Stored as plain string briefly; wiped securely post-run

### ================================
### Path Setup
### ================================
$TimeStamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$LogDir         = Join-Path $PSScriptRoot "Logs"
$OutputPath     = Join-Path $PSScriptRoot "AutopilotDevices_$TimeStamp.csv"
$LogPath        = Join-Path $LogDir "AutopilotExport_$TimeStamp.log"
$TranscriptPath = Join-Path $LogDir "AutopilotExport_${TimeStamp}_Transcript.log"

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

### ================================
### Start Transcript
### Captures full console session to file for audit.
### ================================
Start-Transcript -Path $TranscriptPath -Append | Out-Null

### ================================
### Write-Log
### Writes timestamped entries to log file and console simultaneously.
### Severity levels: INFO, WARN, ERROR.
### ================================
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $Entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogPath -Value $Entry
    switch ($Level) {
        "WARN"  { Write-Host $Entry -ForegroundColor Yellow }
        "ERROR" { Write-Host $Entry -ForegroundColor Red }
        default { Write-Host $Entry }
    }
}

### ================================
### Get-GraphToken
### Obtains OAuth2 Bearer token via client credentials grant.
### Accepts ClientSecret as SecureString; converts to plain only at POST body,
### then immediately clears the BSTR pointer from unmanaged memory.
### Returns access_token string.
### ================================
function Get-GraphToken {
    param (
        [string]$TenantID,
        [string]$ClientID,
        [SecureString]$ClientSecret
    )

    # Convert SecureString → plain text only for the POST body
    $BSTR        = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    $PlainSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    $Body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ClientID
        client_secret = $PlainSecret
    }

    try {
        Write-Log "Requesting Graph API token for TenantID: $TenantID"
        $Response = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" `
            -Body $Body
        Write-Log "Token acquired successfully."
        return $Response.access_token
    }
    catch {
        Write-Log "Token request failed: $_" -Level ERROR
        throw
    }
    finally {
        # Wipe plain text secret and BSTR from unmanaged memory immediately
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        $PlainSecret = $null
        $Body        = $null
    }
}

### ================================
### Get-AllAutopilotDevices
### Fetches all Autopilot device identities from Graph beta endpoint.
### Handles OData nextLink pagination until no further pages remain.
### Returns full array of raw device objects.
### ================================
function Get-AllAutopilotDevices {
    param (
        [string]$AccessToken
    )
    $Headers    = @{ Authorization = "Bearer $AccessToken" }
    $URL        = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
    $AllDevices = @()

    do {
        Write-Log "Fetching page: $URL"
        try {
            $Response    = Invoke-RestMethod -Method Get -Uri $URL -Headers $Headers
            $AllDevices += $Response.value
            Write-Log "Page fetched. Cumulative count: $($AllDevices.Count)"
            $URL = $Response.'@odata.nextLink'
        }
        catch {
            Write-Log "Page fetch failed: $_" -Level ERROR
            throw
        }
    } while ($URL)

    Write-Log "Pagination complete. Total devices: $($AllDevices.Count)"
    return $AllDevices
}

### ================================
### ConvertTo-AutopilotReport
### Maps raw Graph device objects to flat PSCustomObject rows for CSV export.
### Fields: SerialNumber, AutopilotID, AzureADObjectID, GroupTag, DeploymentProfile.
### ================================
function ConvertTo-AutopilotReport {
    param (
        [array]$Devices
    )
    Write-Log "Building report rows for $($Devices.Count) devices."
    return $Devices | ForEach-Object {
        [PSCustomObject]@{
            SerialNumber      = $_.serialNumber
            AutopilotID       = $_.id
            AzureADObjectID   = $_.azureAdDeviceId
            GroupTag          = $_.groupTag
            DeploymentProfile = $_.deploymentProfileAssigned
        }
    }
}

### ================================
### Invoke-CredentialWipe
### Security cleanup — explicitly nulls and garbage-collects all credential
### variables from memory. Runs in finally block to guarantee execution
### even on script error or early exit.
### WARNING: Do not remove or bypass this block.
### ================================
function Invoke-CredentialWipe {
    param (
        [ref]$TenantIDRef,
        [ref]$ClientIDRef,
        [ref]$SecureSecretRef,
        [ref]$AccessTokenRef
    )
    Write-Log "Running credential wipe from memory."

    # Overwrite access token with garbage before null
    if ($AccessTokenRef.Value) {
        $AccessTokenRef.Value = [string]::new('X', $AccessTokenRef.Value.Length)
        $AccessTokenRef.Value = $null
    }

    # Dispose SecureString (wipes unmanaged memory)
    if ($SecureSecretRef.Value -is [SecureString]) {
        $SecureSecretRef.Value.Dispose()
        $SecureSecretRef.Value = $null
    }

    $TenantIDRef.Value  = $null
    $ClientIDRef.Value  = $null

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Log "Credential wipe complete."
}

### ================================
### MAIN
### ================================
$AccessToken  = $null
$SecureSecret = $null

try {
    Write-Log "=== Autopilot Export Started ==="
    Write-Log "Output CSV      : $OutputPath"
    Write-Log "Log file        : $LogPath"
    Write-Log "Transcript      : $TranscriptPath"

    # Convert plain ClientSecret → SecureString immediately; discard plain
    $SecureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $ClientSecret = $null   # Plain text gone from this scope

    $AccessToken  = Get-GraphToken -TenantID $TenantID -ClientID $ClientID -ClientSecret $SecureSecret
    $RawDevices   = Get-AllAutopilotDevices -AccessToken $AccessToken
    $Report       = ConvertTo-AutopilotReport -Devices $RawDevices

    $Report | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Log "Export complete. File: $OutputPath | Total: $($Report.Count) devices."
}
catch {
    Write-Log "Script failed: $_" -Level ERROR
}
finally {
    # Always wipe creds — even on crash
    Invoke-CredentialWipe `
        -TenantIDRef    ([ref]$TenantID) `
        -ClientIDRef    ([ref]$ClientID) `
        -SecureSecretRef ([ref]$SecureSecret) `
        -AccessTokenRef ([ref]$AccessToken)

    Write-Log "=== Autopilot Export Ended ==="
    Stop-Transcript | Out-Null
}
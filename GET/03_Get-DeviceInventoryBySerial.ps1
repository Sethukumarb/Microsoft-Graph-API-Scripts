#Requires -Version 5.1
# ==============================================================================
# Script Name  : Get-DeviceInventoryBySerial.ps1
# Description  : Pulls a complete inventory of all objects linked to each
#                device serial number across Intune, Autopilot, and Azure AD.
#
#                FOR EACH SERIAL NUMBER THE SCRIPT REPORTS:
#                  - All Intune managed device records (1 or many duplicates)
#                  - All Autopilot device identity records (1 or many)
#                  - Azure AD device object (looked up from Intune azureADDeviceId)
#
#                READ ONLY - no changes made anywhere.
#
#                INPUT FILE:
#                  input.txt - same folder as this script
#                  One serial number per line. # comments and blank lines ignored.
#
#                OUTPUT FILES (saved to $PSScriptRoot):
#                  DeviceInventory_[timestamp].csv   - full inventory
#                  DeviceInventory_[timestamp].log   - run log
#                  DeviceInventory_Transcript_[timestamp].log - PS transcript
#
#                CSV COLUMNS:
#                  SerialNumber, RecordType, RecordIndex,
#                  DeviceName, Manufacturer, Model,
#                  OSPlatform, OSVersion,
#                  IntuneDeviceID, AzureADDeviceID,
#                  AutopilotDeviceID, GroupTag,
#                  EnrollmentState, ComplianceState, ManagementState,
#                  PrimaryUserUPN, LastSyncDateTime, EnrollmentDate,
#                  ProfileStatus, ProfileName,
#                  AzureADObjectID, AzureADDisplayName,
#                  AzureADIsManaged, AzureADIsCompliant,
#                  Verdict
#
#                VERDICT VALUES (per row):
#                  Intune record         FOUND - 1 record
#                  Intune record         FOUND - DUPLICATE (N of M)
#                  Autopilot record      FOUND - 1 record
#                  Autopilot record      FOUND - DUPLICATE (N of M)
#                  Azure AD object       FOUND in Azure AD
#                  Azure AD object       NOT FOUND in Azure AD
#                  Azure AD object       NO PERMISSION to read Azure AD
#
# Author       : Sethu Kumar B
# Version      : 1.2
# Created Date : 2026-04-17
# Last Modified: 2026-04-17
#
# Requirements :
#   - Azure AD App Registration (READ-ONLY)
#   - Graph API Application Permissions (admin consent granted):
#       DeviceManagementManagedDevices.Read.All   - Intune device records
#       DeviceManagementServiceConfig.Read.All    - Autopilot records
#       Device.Read.All                           - Azure AD device lookup
#                                                   (optional - script handles
#                                                    no-permission gracefully)
#   - PowerShell 5.1 or later
#
# Change Log   :
#   v1.0 - 2026-04-17 - Sethu Kumar B - Initial release.
#   v1.1 - 2026-04-17 - Sethu Kumar B - Windows OS filter on Intune pull.
#   v1.2 - 2026-04-17 - Sethu Kumar B - Azure AD now lists ALL matching
#                        objects per deviceId (not just first). Duplicates
#                        in Entra shown as FOUND - DUPLICATE (N of M) same
#                        as Intune. Azure AD duplicate count in summary.
#                        pull ($filter=operatingSystem eq Windows). Only
#                        Windows devices loaded into lookup table.
# ==============================================================================


#region --- CONFIGURATION -------------------------------------------------------

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

$InputFileName = "inputserialnumber.txt"
$InputPath     = Join-Path $PSScriptRoot $InputFileName

$MaxRetries = 5

#endregion ----------------------------------------------------------------------


#region --- INIT ----------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile     = Join-Path $PSScriptRoot "DeviceInventory_$Timestamp.csv"
$script:LogFile = Join-Path $PSScriptRoot "DeviceInventory_$Timestamp.log"
$TranscriptFile = Join-Path $PSScriptRoot "DeviceInventory_Transcript_$Timestamp.log"

try { Start-Transcript -Path $TranscriptFile -Force | Out-Null } catch { }

try {
    [System.IO.File]::WriteAllText($script:LogFile,
        "Get-DeviceInventoryBySerial v1.0`r`nStarted: $(Get-Date)`r`n`r`n",
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


function Invoke-GraphGetAllPages {
    param (
        [string]$InitialUri,
        [string]$AccessToken,
        [string]$Label,
        [int]$MaxRetries = 5
    )
    $Headers    = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $AllRecords = [System.Collections.Generic.List[PSObject]]::new()
    $Uri        = $InitialUri
    $Page       = 0
    $Total      = 0

    Write-Log "Pulling: $Label" -Level INFO

    do {
        $Page++
        $Attempt = 0
        $Success = $false
        do {
            $Attempt++
            try {
                $r      = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
                $Count  = $r.value.Count
                $Total += $Count
                Write-Log "  Page $Page - $Count records (total: $Total)" -Level INFO
                foreach ($rec in $r.value) { $AllRecords.Add($rec) }
                $Uri     = $r.'@odata.nextLink'
                $Success = $true
            }
            catch {
                $Code = $_.Exception.Response.StatusCode.value__
                if ($Code -eq 429) {
                    $Wait = 60
                    try {
                        $HV = $_.Exception.Response.Headers.GetValues("Retry-After")
                        if ($HV -and $HV.Count -gt 0) { $Wait = [int]$HV[0] }
                    } catch { }
                    $Wait += Get-Random -Minimum 1 -Maximum 10
                    if ($Attempt -lt $MaxRetries) {
                        Write-Log "  429 - waiting ${Wait}s (retry $Attempt/$MaxRetries)..." -Level WARN
                        Start-Sleep -Seconds $Wait
                    } else {
                        Write-Log "  429 persisted. Skipping page." -Level WARN
                        $Uri = $null; $Success = $true
                    }
                } else {
                    Write-Log "  Page $Page failed - HTTP $Code" -Level ERROR
                    $Uri = $null; $Success = $true
                }
            }
        } while (-not $Success -and $Attempt -lt $MaxRetries)
    } while ($Uri)

    Write-Log "$Label - $Total total records." -Level SUCCESS
    return $AllRecords.ToArray()
}


function Format-Date {
    param ([string]$DateString)
    if ([string]::IsNullOrWhiteSpace($DateString) -or $DateString -like "0001*") { return "N/A" }
    try {
        $dt = [datetime]::Parse($DateString, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
        return $dt.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch { return $DateString }
}


# Looks up ALL Azure AD device objects matching the given azureADDeviceId.
# One azureADDeviceId can have multiple Azure AD objects (duplicates in Entra).
# Returns a Generic List of PSObjects - one per Azure AD object found.
# Never throws - errors returned as a single not-found item in the list.
function Get-AzureADDeviceObjects {
    param ([string]$AzureADDeviceId, [string]$AccessToken)

    $Results = [System.Collections.Generic.List[PSObject]]::new()

    if ([string]::IsNullOrWhiteSpace($AzureADDeviceId) -or $AzureADDeviceId -eq "N/A") {
        $Results.Add([PSCustomObject]@{
            Found = $false; Status = "NO AZURE AD DEVICE ID"
            ObjectID = "N/A"; DisplayName = "N/A"; IsManaged = "N/A"; IsCompliant = "N/A"
        })
        return $Results
    }

    $Headers = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $Uri     = "https://graph.microsoft.com/v1.0/devices" +
               "?`$filter=deviceId eq '$AzureADDeviceId'" +
               "&`$select=id,displayName,deviceId,isManaged,isCompliant,operatingSystem"
    try {
        $r = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
        if ($r.value -and $r.value.Count -gt 0) {
            foreach ($d in $r.value) {
                $Results.Add([PSCustomObject]@{
                    Found       = $true
                    Status      = "FOUND in Azure AD"
                    ObjectID    = if ($d.id)                    { [string]$d.id }          else { "N/A" }
                    DisplayName = if ($d.displayName)           { [string]$d.displayName } else { "N/A" }
                    IsManaged   = if ($null -ne $d.isManaged)   { [string]$d.isManaged }   else { "N/A" }
                    IsCompliant = if ($null -ne $d.isCompliant) { [string]$d.isCompliant } else { "N/A" }
                })
            }
        } else {
            $Results.Add([PSCustomObject]@{
                Found = $false; Status = "NOT FOUND in Azure AD"
                ObjectID = "N/A"; DisplayName = "N/A"; IsManaged = "N/A"; IsCompliant = "N/A"
            })
        }
    }
    catch {
        $Code      = $_.Exception.Response.StatusCode.value__
        $StatusMsg = if ($Code -eq 403 -or $Code -eq 401) {
            "NO PERMISSION - Device.Read.All not granted"
        } else { "LOOKUP ERROR - HTTP $Code" }
        $Results.Add([PSCustomObject]@{
            Found = $false; Status = $StatusMsg
            ObjectID = "N/A"; DisplayName = "N/A"; IsManaged = "N/A"; IsCompliant = "N/A"
        })
    }
    return $Results
}

#endregion ----------------------------------------------------------------------


#region --- MAIN ----------------------------------------------------------------

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Device Inventory by Serial  |  Sethu Kumar B  |  READ ONLY   " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "" -Level BLANK
Write-Log "Input file   : $InputPath"         -Level INFO
Write-Log "Output CSV   : $OutputFile"        -Level INFO
Write-Log "Log file     : $($script:LogFile)" -Level INFO
Write-Log "Transcript   : $TranscriptFile"    -Level INFO
Write-Log "" -Level BLANK


# -- Step 1: Read serial numbers ------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 1 - Reading Input File" -Level INFO
Write-Log "==========================================================" -Level SECTION

if (-not (Test-Path $InputPath)) {
    Write-Log "Input file not found: $InputPath" -Level ERROR
    Write-Log "Create input.txt with one serial number per line." -Level ERROR
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}

$Serials = [System.Collections.Generic.List[string]]::new()
foreach ($Line in (Get-Content -Path $InputPath -Encoding UTF8)) {
    $T = $Line.Trim()
    if ([string]::IsNullOrWhiteSpace($T) -or $T.StartsWith("#")) { continue }
    $Serials.Add($T)
}

Write-Log "Serials loaded: $($Serials.Count)" -Level $(if ($Serials.Count -gt 0) {"SUCCESS"} else {"WARN"})
if ($Serials.Count -eq 0) {
    Write-Log "No serials found. Exiting." -Level WARN
    try { Stop-Transcript | Out-Null } catch { }
    exit 0
}
foreach ($s in $Serials) { Write-Log "  -> $s" -Level INFO }
Write-Log "" -Level BLANK


# -- Step 2: Authenticate ------------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 2 - Authenticating" -Level INFO
Write-Log "==========================================================" -Level SECTION
$Token = Get-GraphToken -TenantId $TenantID -ClientId $ClientID -ClientSecret $ClientSecret


# -- Step 3: Bulk pull Intune devices ------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 3 - Pulling Intune Managed Devices (Windows only)" -Level INFO
Write-Log "==========================================================" -Level SECTION

# Server-side filter: Windows only. Reduces payload - skips iOS/Android/macOS.
# $filter + $select is safe on managedDevices (unlike Autopilot endpoint).
$IntuneUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices" +
             "?`$filter=operatingSystem eq 'Windows'" +
             "&`$select=id,serialNumber,deviceName,manufacturer,model,operatingSystem," +
             "osVersion,azureADDeviceId,userPrincipalName,enrolledDateTime," +
             "lastSyncDateTime,complianceState,managementState,deviceEnrollmentType&`$top=1000"

$AllIntune = Invoke-GraphGetAllPages -InitialUri $IntuneUri -AccessToken $Token `
             -Label "Intune Windows managed devices" -MaxRetries $MaxRetries

# Store ALL records per serial as a list
$IntuneBySerial = @{}
foreach ($d in $AllIntune) {
    if (-not [string]::IsNullOrWhiteSpace($d.serialNumber)) {
        $Key = $d.serialNumber.ToLower().Trim()
        if (-not $IntuneBySerial.ContainsKey($Key)) {
            $IntuneBySerial[$Key] = [System.Collections.Generic.List[PSObject]]::new()
        }
        $IntuneBySerial[$Key].Add($d)
    }
}
Write-Log "Intune indexed: $($IntuneBySerial.Count) unique serials from $($AllIntune.Count) total records." -Level SUCCESS
Write-Log "" -Level BLANK


# -- Step 4: Bulk pull Autopilot devices ---------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 4 - Pulling Autopilot Device Identities" -Level INFO
Write-Log "==========================================================" -Level SECTION
# No $select - causes HTTP 500 with $top on this endpoint

$APUri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities" +
         "?`$top=1000"

$AllAP = Invoke-GraphGetAllPages -InitialUri $APUri -AccessToken $Token `
         -Label "Autopilot device identities" -MaxRetries $MaxRetries

$APBySerial = @{}
foreach ($ap in $AllAP) {
    if (-not [string]::IsNullOrWhiteSpace($ap.serialNumber)) {
        $Key = $ap.serialNumber.ToLower().Trim()
        if (-not $APBySerial.ContainsKey($Key)) {
            $APBySerial[$Key] = [System.Collections.Generic.List[PSObject]]::new()
        }
        $APBySerial[$Key].Add($ap)
    }
}
Write-Log "Autopilot indexed: $($APBySerial.Count) unique serials from $($AllAP.Count) total records." -Level SUCCESS
Write-Log "" -Level BLANK


# -- Step 5: Process each serial -----------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 5 - Building Inventory per Serial" -Level INFO
Write-Log "==========================================================" -Level SECTION

$Results = [System.Collections.Generic.List[PSObject]]::new()
$Counter = 0
$Total   = $Serials.Count

foreach ($Serial in $Serials) {
    $Counter++
    $SerialKey = $Serial.ToLower().Trim()

    $IntuneRecords = if ($IntuneBySerial.ContainsKey($SerialKey)) { $IntuneBySerial[$SerialKey] } else { [System.Collections.Generic.List[PSObject]]::new() }
    $APRecords     = if ($APBySerial.ContainsKey($SerialKey))     { $APBySerial[$SerialKey] }     else { [System.Collections.Generic.List[PSObject]]::new() }

    $ICount = $IntuneRecords.Count
    $ACount = $APRecords.Count

    Write-Log "" -Level BLANK
    Write-Log "  [$Counter/$Total] Serial: $Serial" -Level INFO
    Write-Log ("    Intune    : {0} record(s){1}" -f $ICount,
        $(if ($ICount -gt 1) {" *** $ICount DUPLICATES ***"} elseif ($ICount -eq 0) {" - NOT FOUND"} else {""})) `
        -Level $(if ($ICount -gt 1) {"WARN"} elseif ($ICount -eq 0) {"WARN"} else {"SUCCESS"})
    Write-Log ("    Autopilot : {0} record(s){1}" -f $ACount,
        $(if ($ACount -gt 1) {" *** $ACount DUPLICATES ***"} elseif ($ACount -eq 0) {" - NOT FOUND"} else {""})) `
        -Level $(if ($ACount -gt 1) {"WARN"} elseif ($ACount -eq 0) {"WARN"} else {"SUCCESS"})

    # -- Intune rows -----------------------------------------------------------
    if ($ICount -eq 0) {
        $Results.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            RecordType        = "Intune"
            RecordIndex       = "N/A"
            DeviceName        = "N/A"; Manufacturer = "N/A"; Model = "N/A"
            OSPlatform        = "N/A"; OSVersion = "N/A"
            IntuneDeviceID    = "N/A"; AzureADDeviceID = "N/A"
            AutopilotDeviceID = "N/A"; GroupTag = "N/A"
            EnrollmentState   = "N/A"; ComplianceState = "N/A"; ManagementState = "N/A"
            PrimaryUserUPN    = "N/A"; LastSyncDateTime = "N/A"; EnrollmentDate = "N/A"
            ProfileStatus     = "N/A"; ProfileName = "N/A"
            AzureADObjectID   = "N/A"; AzureADDisplayName = "N/A"
            AzureADIsManaged  = "N/A"; AzureADIsCompliant = "N/A"
            Verdict           = "NOT FOUND in Intune"
        })
    }

    $IntuneIdx = 0
    foreach ($d in $IntuneRecords) {
        $IntuneIdx++
        $Verdict = if ($ICount -eq 1) { "FOUND - 1 Intune record" } else { "FOUND - DUPLICATE ($IntuneIdx of $ICount Intune records)" }
        $DeviceName = if ($d.deviceName)      { [string]$d.deviceName }      else { "N/A" }
        $AADId      = if ($d.azureADDeviceId) { [string]$d.azureADDeviceId } else { "N/A" }

        Write-Log "    [Intune $IntuneIdx/$ICount] $DeviceName | ID: $([string]$d.id)" `
                  -Level $(if ($ICount -gt 1) {"WARN"} else {"SUCCESS"})

        $Results.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            RecordType        = "Intune"
            RecordIndex       = "$IntuneIdx of $ICount"
            DeviceName        = $DeviceName
            Manufacturer      = if ($d.manufacturer)         { [string]$d.manufacturer }         else { "N/A" }
            Model             = if ($d.model)                { [string]$d.model }                else { "N/A" }
            OSPlatform        = if ($d.operatingSystem)      { [string]$d.operatingSystem }      else { "N/A" }
            OSVersion         = if ($d.osVersion)            { [string]$d.osVersion }            else { "N/A" }
            IntuneDeviceID    = if ($d.id)                   { [string]$d.id }                   else { "N/A" }
            AzureADDeviceID   = $AADId
            AutopilotDeviceID = "N/A"
            GroupTag          = "N/A"
            EnrollmentState   = "N/A"
            ComplianceState   = if ($d.complianceState)      { [string]$d.complianceState }      else { "N/A" }
            ManagementState   = if ($d.managementState)      { [string]$d.managementState }      else { "N/A" }
            PrimaryUserUPN    = if ($d.userPrincipalName)    { [string]$d.userPrincipalName }    else { "N/A" }
            LastSyncDateTime  = Format-Date $d.lastSyncDateTime
            EnrollmentDate    = Format-Date $d.enrolledDateTime
            ProfileStatus     = "N/A"
            ProfileName       = "N/A"
            AzureADObjectID   = "N/A"
            AzureADDisplayName= "N/A"
            AzureADIsManaged  = "N/A"
            AzureADIsCompliant= "N/A"
            Verdict           = $Verdict
        })
    }

    # -- Autopilot rows --------------------------------------------------------
    if ($ACount -eq 0) {
        $Results.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            RecordType        = "Autopilot"
            RecordIndex       = "N/A"
            DeviceName        = "N/A"; Manufacturer = "N/A"; Model = "N/A"
            OSPlatform        = "N/A"; OSVersion = "N/A"
            IntuneDeviceID    = "N/A"; AzureADDeviceID = "N/A"
            AutopilotDeviceID = "N/A"; GroupTag = "N/A"
            EnrollmentState   = "N/A"; ComplianceState = "N/A"; ManagementState = "N/A"
            PrimaryUserUPN    = "N/A"; LastSyncDateTime = "N/A"; EnrollmentDate = "N/A"
            ProfileStatus     = "N/A"; ProfileName = "N/A"
            AzureADObjectID   = "N/A"; AzureADDisplayName = "N/A"
            AzureADIsManaged  = "N/A"; AzureADIsCompliant = "N/A"
            Verdict           = "NOT FOUND in Autopilot"
        })
    }

    $APIdx = 0
    foreach ($ap in $APRecords) {
        $APIdx++
        $Verdict = if ($ACount -eq 1) { "FOUND - 1 Autopilot record" } else { "FOUND - DUPLICATE ($APIdx of $ACount Autopilot records)" }
        $APModel = if ($ap.model) { [string]$ap.model } else { "N/A" }

        Write-Log "    [Autopilot $APIdx/$ACount] $APModel | GroupTag: $(if ($ap.groupTag) {$ap.groupTag} else {'(empty)'}) | ID: $([string]$ap.id)" `
                  -Level $(if ($ACount -gt 1) {"WARN"} else {"SUCCESS"})

        # Profile name from expanded fields if available
        $ProfileName = "N/A"
        if ($ap.PSObject.Properties["deploymentProfile"] -and $ap.deploymentProfile -and $ap.deploymentProfile.displayName) {
            $ProfileName = [string]$ap.deploymentProfile.displayName
        } elseif ($ap.PSObject.Properties["displayName"] -and -not [string]::IsNullOrWhiteSpace($ap.displayName)) {
            $ProfileName = [string]$ap.displayName
        }

        $Results.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            RecordType        = "Autopilot"
            RecordIndex       = "$APIdx of $ACount"
            DeviceName        = "N/A"
            Manufacturer      = if ($ap.manufacturer) { [string]$ap.manufacturer } else { "N/A" }
            Model             = $APModel
            OSPlatform        = "Windows"
            OSVersion         = "N/A"
            IntuneDeviceID    = if ($ap.managedDeviceId)                       { [string]$ap.managedDeviceId }                       else { "N/A" }
            AzureADDeviceID   = if ($ap.azureActiveDirectoryDeviceId)          { [string]$ap.azureActiveDirectoryDeviceId }          else { "N/A" }
            AutopilotDeviceID = if ($ap.id)                                    { [string]$ap.id }                                    else { "N/A" }
            GroupTag          = if ($ap.groupTag)                              { [string]$ap.groupTag }                              else { "(empty)" }
            EnrollmentState   = if ($ap.enrollmentState)                       { [string]$ap.enrollmentState }                       else { "N/A" }
            ComplianceState   = "N/A"
            ManagementState   = "N/A"
            PrimaryUserUPN    = if ($ap.userPrincipalName)                     { [string]$ap.userPrincipalName }                     else { "N/A" }
            LastSyncDateTime  = Format-Date $ap.lastContactedDateTime
            EnrollmentDate    = "N/A"
            ProfileStatus     = if ($ap.deploymentProfileAssignmentStatus)     { [string]$ap.deploymentProfileAssignmentStatus }     else { "N/A" }
            ProfileName       = $ProfileName
            AzureADObjectID   = "N/A"
            AzureADDisplayName= "N/A"
            AzureADIsManaged  = "N/A"
            AzureADIsCompliant= "N/A"
            Verdict           = $Verdict
        })
    }

    # -- Azure AD rows (one per unique azureADDeviceId from Intune records) ----
    # Collect unique Azure AD Device IDs from all Intune records for this serial
    $AADIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($d in $IntuneRecords) {
        if ($d.azureADDeviceId -and $d.azureADDeviceId -ne "N/A") {
            [void]$AADIds.Add([string]$d.azureADDeviceId)
        }
    }

    if ($AADIds.Count -eq 0) {
        # No Azure AD ID available - add a single not-available row
        $Results.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            RecordType        = "Azure AD"
            RecordIndex       = "N/A"
            DeviceName        = "N/A"; Manufacturer = "N/A"; Model = "N/A"
            OSPlatform        = "N/A"; OSVersion = "N/A"
            IntuneDeviceID    = "N/A"; AzureADDeviceID = "N/A"
            AutopilotDeviceID = "N/A"; GroupTag = "N/A"
            EnrollmentState   = "N/A"; ComplianceState = "N/A"; ManagementState = "N/A"
            PrimaryUserUPN    = "N/A"; LastSyncDateTime = "N/A"; EnrollmentDate = "N/A"
            ProfileStatus     = "N/A"; ProfileName = "N/A"
            AzureADObjectID   = "N/A"; AzureADDisplayName = "N/A"
            AzureADIsManaged  = "N/A"; AzureADIsCompliant = "N/A"
            Verdict           = "Azure AD - no device ID available (device may not be in Intune)"
        })
    }

    $AADIdx = 0
    foreach ($AADId in $AADIds) {
        $AADIdx++
        # Returns ALL Azure AD objects for this device ID - handles duplicates
        $AADObjects = Get-AzureADDeviceObjects -AzureADDeviceId $AADId -AccessToken $Token
        $AADObjCount = $AADObjects.Count
        $FoundCount  = @($AADObjects | Where-Object { $_.Found }).Count

        if ($FoundCount -gt 1) {
            Write-Log ("    [Azure AD] $AADId - *** $FoundCount DUPLICATE AZURE AD OBJECTS ***") -Level WARN
        } elseif ($FoundCount -eq 1) {
            Write-Log "    [Azure AD] $AADId - FOUND in Azure AD" -Level SUCCESS
        } else {
            Write-Log "    [Azure AD] $AADId - $($AADObjects[0].Status)" -Level WARN
        }

        # Find the matching Intune record for device name
        $MatchingIntune = $IntuneRecords | Where-Object { $_.azureADDeviceId -eq $AADId } | Select-Object -First 1
        $DevNameForAAD  = if ($MatchingIntune -and $MatchingIntune.deviceName) { [string]$MatchingIntune.deviceName } else { "N/A" }
        $IntuneIdForAAD = if ($MatchingIntune -and $MatchingIntune.id) { [string]$MatchingIntune.id } else { "N/A" }

        $ObjIdx = 0
        foreach ($AADObj in $AADObjects) {
            $ObjIdx++
            # Verdict shows duplicate count if more than one found
            $Verdict = if (-not $AADObj.Found) {
                "Azure AD - $($AADObj.Status)"
            } elseif ($FoundCount -eq 1) {
                "Azure AD - FOUND - 1 object"
            } else {
                "Azure AD - FOUND - DUPLICATE ($ObjIdx of $FoundCount Azure AD objects)"
            }

            if ($FoundCount -gt 1) {
                Write-Log "      [Azure AD object $ObjIdx/$FoundCount] $($AADObj.DisplayName) | ObjectID: $($AADObj.ObjectID)" -Level WARN
            }

            $Results.Add([PSCustomObject]@{
                SerialNumber      = $Serial
                RecordType        = "Azure AD"
                RecordIndex       = if ($FoundCount -gt 1) { "$ObjIdx of $FoundCount" } else { "$AADIdx" }
                DeviceName        = $DevNameForAAD
                Manufacturer      = "N/A"
                Model             = "N/A"
                OSPlatform        = "N/A"
                OSVersion         = "N/A"
                IntuneDeviceID    = $IntuneIdForAAD
                AzureADDeviceID   = $AADId
                AutopilotDeviceID = "N/A"
                GroupTag          = "N/A"
                EnrollmentState   = "N/A"
                ComplianceState   = "N/A"
                ManagementState   = "N/A"
                PrimaryUserUPN    = "N/A"
                LastSyncDateTime  = "N/A"
                EnrollmentDate    = "N/A"
                ProfileStatus     = "N/A"
                ProfileName       = "N/A"
                AzureADObjectID   = $AADObj.ObjectID
                AzureADDisplayName= $AADObj.DisplayName
                AzureADIsManaged  = $AADObj.IsManaged
                AzureADIsCompliant= $AADObj.IsCompliant
                Verdict           = $Verdict
            })
        }
        Start-Sleep -Milliseconds 200
    }
}


# -- Step 6: Export CSV ---------------------------------------------------------
Write-Log "" -Level BLANK
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 6 - Exporting CSV" -Level INFO
Write-Log "==========================================================" -Level SECTION

try {
    $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    $SizeMB = (((Get-Item $OutputFile).Length) / 1MB).ToString("0.00")
    Write-Log "CSV exported." -Level SUCCESS
    Write-Log "  Path : $OutputFile" -Level INFO
    Write-Log "  Rows : $($Results.Count)  |  Size: $SizeMB MB" -Level INFO
}
catch { Write-Log "CSV export failed: $_" -Level ERROR }


# -- Summary --------------------------------------------------------------------
$IFound    = @($Results | Where-Object { $_.RecordType -eq "Intune"    -and $_.Verdict -notlike "NOT FOUND*" }).Count
$INotFound = @($Results | Where-Object { $_.RecordType -eq "Intune"    -and $_.Verdict -like "NOT FOUND*" }).Count
$IDupes    = @($Results | Where-Object { $_.RecordType -eq "Intune"    -and $_.Verdict -like "*DUPLICATE*" }).Count
$AFound    = @($Results | Where-Object { $_.RecordType -eq "Autopilot" -and $_.Verdict -notlike "NOT FOUND*" }).Count
$ANotFound = @($Results | Where-Object { $_.RecordType -eq "Autopilot" -and $_.Verdict -like "NOT FOUND*" }).Count
$ADupes    = @($Results | Where-Object { $_.RecordType -eq "Autopilot" -and $_.Verdict -like "*DUPLICATE*" }).Count
$AADFound  = @($Results | Where-Object { $_.RecordType -eq "Azure AD" -and $_.Verdict -like "*FOUND*" -and $_.Verdict -notlike "*NOT FOUND*" }).Count
$AADDupes  = @($Results | Where-Object { $_.RecordType -eq "Azure AD" -and $_.Verdict -like "*DUPLICATE*" }).Count
$AADMissing= @($Results | Where-Object { $_.RecordType -eq "Azure AD" -and $_.Verdict -like "*NOT FOUND*" }).Count
$AADNoPerm = @($Results | Where-Object { $_.RecordType -eq "Azure AD" -and $_.Verdict -like "*NO PERMISSION*" }).Count

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  COMPLETE - READ ONLY - NO CHANGES MADE                       " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "  Serials processed       : $Total"        -Level INFO
Write-Log "" -Level BLANK
Write-Log "  -- Intune -----------------------------------------" -Level SECTION
Write-Log "  Records found           : $IFound"       -Level $(if ($IFound    -gt 0) {"SUCCESS"} else {"WARN"})
Write-Log "  Not found               : $INotFound"    -Level $(if ($INotFound -gt 0) {"WARN"}    else {"INFO"})
Write-Log "  Duplicates              : $IDupes"       -Level $(if ($IDupes    -gt 0) {"WARN"}    else {"INFO"})
Write-Log "" -Level BLANK
Write-Log "  -- Autopilot --------------------------------------" -Level SECTION
Write-Log "  Records found           : $AFound"       -Level $(if ($AFound    -gt 0) {"SUCCESS"} else {"WARN"})
Write-Log "  Not found               : $ANotFound"    -Level $(if ($ANotFound -gt 0) {"WARN"}    else {"INFO"})
Write-Log "  Duplicates              : $ADupes"       -Level $(if ($ADupes    -gt 0) {"WARN"}    else {"INFO"})
Write-Log "" -Level BLANK
Write-Log "  -- Azure AD ----------------------------------------" -Level SECTION
Write-Log "  Found in Azure AD       : $AADFound"     -Level $(if ($AADFound  -gt 0) {"SUCCESS"} else {"INFO"})
Write-Log "  Duplicates in Azure AD  : $AADDupes"     -Level $(if ($AADDupes   -gt 0) {"WARN"}   else {"INFO"})
Write-Log "  Not found in Azure AD   : $AADMissing"   -Level $(if ($AADMissing -gt 0) {"WARN"}   else {"INFO"})
Write-Log "  No read permission      : $AADNoPerm"    -Level $(if ($AADNoPerm  -gt 0) {"WARN"}   else {"INFO"})
Write-Log "" -Level BLANK
Write-Log "  Output CSV   : $OutputFile"        -Level INFO
Write-Log "  Log file     : $($script:LogFile)" -Level INFO
Write-Log "  Transcript   : $TranscriptFile"    -Level INFO
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

try { Stop-Transcript | Out-Null } catch { }

#endregion ----------------------------------------------------------------------
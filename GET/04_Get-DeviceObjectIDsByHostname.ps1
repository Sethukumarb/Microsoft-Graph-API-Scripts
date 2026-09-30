#Requires -Version 5.1
# ==============================================================================
# Script Name  : Get-DeviceObjectIDsByHostname.ps1
# Description  : Given a list of device hostnames from a text file, searches
#                Azure AD for ALL matching device objects per hostname and
#                returns both the Azure AD Object ID and Intune Device ID.
#
#                If a hostname has multiple entries (duplicates), ALL of them
#                are listed — each as a separate row in the CSV and a separate
#                line in the BulkImport TXT.
#
#                WHY THIS SCRIPT EXISTS:
#                  When adding devices to a static Entra group using Bulk Import,
#                  you need the Azure AD Object ID of each device. This script
#                  resolves hostnames to Object IDs so you can upload the TXT
#                  directly into the Entra portal Bulk Import flow.
#
#                HOW IT WORKS:
#                  1. Read hostnames from DeviceList.txt (one per line)
#                  2. For each hostname:
#                     a. Search Azure AD  - GET /v1.0/devices?$filter=displayName eq
#                     b. Search Intune    - GET /beta/deviceManagement/managedDevices
#                     c. Cross-reference  - match AAD objects to Intune records
#                        using deviceId / azureADDeviceId
#                     d. Flag duplicates  - if more than one AAD object found
#                  3. Export full CSV report (all details)
#                  4. Export BulkImport TXT (Azure AD Object IDs only)
#                  5. Print console summary + save log
#
#                DUPLICATE HANDLING:
#                  If a hostname returns 2+ Azure AD objects, ALL are included
#                  in both the CSV and BulkImport TXT. The IsDuplicate column
#                  in the CSV flags these rows so you can review before importing.
#
#                CSV COLUMNS:
#                  Hostname, AADObjectID, AADDeviceID, IntuneDeviceID,
#                  SerialNumber, PrimaryUser, EnrolledDate, LastSync,
#                  DaysSinceSync, OS, OSVersion, ComplianceState,
#                  ManagementState, OwnerType, IsDuplicate, AADObjectCount,
#                  FoundInAAD, FoundInIntune, Notes
#
#                INPUT FILE:
#                  DeviceList.txt - same folder as this script
#                  One hostname per line. Blank lines ignored.
#
#                  Example:
#                    W-5k2ChJSOHb
#                    W-VegpHYqeeM
#                    W-LT-00142
#
#                OUTPUT FILES (saved to $PSScriptRoot):
#                  Get-DeviceObjectIDs_[timestamp].csv
#                  Get-DeviceObjectIDs_[timestamp]_BulkImport.txt
#                  Get-DeviceObjectIDs_[timestamp].log
#
#                HOW TO USE THE BULK IMPORT TXT:
#                  1. Open Entra portal - Groups - your group - Members
#                  2. Click "Bulk Add Members"
#                  3. Upload the _BulkImport.txt file
#                  4. Review and confirm
#
# Author       : Sethu Kumar B
# Version      : 1.0
# Created Date : 2026-04-14
#
# Requirements :
#   - Azure AD App Registration
#   - Graph API Application Permissions (admin consent granted):
#       Device.Read.All                          - search Azure AD device objects
#       DeviceManagementManagedDevices.Read.All  - search Intune records
#   - PowerShell 5.1 or later
#
# Change Log   :
#   v1.0 - 2026-04-14 - Endpoint Engineering Team - Initial release.
# ==============================================================================


#region --- CONFIGURATION -----------------------------------------------------

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

# -- INPUT FILE ----------------------------------------------------------------
# Plain text file with one hostname per line.
# Saved in the same folder as this script.
# ------------------------------------------------------------------------------
$InputFileName = "DeviceList.txt"
$InputFilePath = Join-Path $PSScriptRoot $InputFileName

# -- OUTPUT FOLDER -------------------------------------------------------------
# Leave blank to save all outputs to the same folder as this script.
# ------------------------------------------------------------------------------
$OutputFolder = ""

#endregion --------------------------------------------------------------------


#region --- FUNCTIONS ---------------------------------------------------------

# ------------------------------------------------------------------------------
# FUNCTION : Write-Log
# Purpose  : Writes timestamped colour-coded messages to console and log file.
# ------------------------------------------------------------------------------
function Write-Log {
    param (
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR","SECTION","BLANK")]
        [string]$Level = "INFO"
    )
    $ColourMap = @{
        INFO    = "Gray";  SUCCESS = "Green"; WARN  = "Yellow"
        ERROR   = "Red";   SECTION = "Cyan";  BLANK = "Gray"
    }
    $PrefixMap = @{
        INFO    = "[INFO]   "; SUCCESS = "[OK]     "; WARN    = "[WARN]   "
        ERROR   = "[ERROR]  "; SECTION = "[=======]"; BLANK   = "         "
    }
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


# ------------------------------------------------------------------------------
# FUNCTION : Get-GraphAccessToken
# Purpose  : Authenticates via client credentials and returns a bearer token.
# ------------------------------------------------------------------------------
function Get-GraphAccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    try {
        Write-Log "Requesting access token..." -Level INFO
        $r = Invoke-RestMethod -Method POST `
             -ContentType "application/x-www-form-urlencoded" `
             -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
             -Body $Body -ErrorAction Stop
        Write-Log "Access token acquired." -Level SUCCESS
        return $r.access_token
    }
    catch {
        Write-Log "Authentication failed: $_" -Level ERROR
        exit 1
    }
}


# ------------------------------------------------------------------------------
# FUNCTION : Get-AADObjectsByHostname
# Purpose  : Searches Azure AD for ALL device objects matching a hostname.
#            Returns all matches including duplicates.
# ------------------------------------------------------------------------------
function Get-AADObjectsByHostname {
    param ([string]$Hostname, [string]$AccessToken)

    $Headers  = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $SafeName = $Hostname -replace "'", "''"
    $Uri      = "https://graph.microsoft.com/v1.0/devices" +
                "?`$filter=displayName eq '$SafeName'" +
                "&`$select=id,deviceId,displayName,registrationDateTime," +
                "approximateLastSignInDateTime,operatingSystem,operatingSystemVersion," +
                "isCompliant,managementType,physicalIds,trustType,accountEnabled"
    try {
        $r = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
        return @($r.value)
    }
    catch {
        Write-Log "  Azure AD search failed for '$Hostname': $_" -Level WARN
        return @()
    }
}


# ------------------------------------------------------------------------------
# FUNCTION : Get-IntuneRecordsByHostname
# Purpose  : Searches Intune for ALL managed device records matching a hostname.
#            Returns all matches including duplicates.
# ------------------------------------------------------------------------------
function Get-IntuneRecordsByHostname {
    param ([string]$Hostname, [string]$AccessToken)

    $Headers  = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $SafeName = $Hostname -replace "'", "''"
    $Uri      = "https://graph.microsoft.com/beta/deviceManagement/managedDevices" +
                "?`$filter=deviceName eq '$SafeName'"
    try {
        $r = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
        return @($r.value)
    }
    catch {
        Write-Log "  Intune search failed for '$Hostname': $_" -Level WARN
        return @()
    }
}


# ------------------------------------------------------------------------------
# FUNCTION : Get-SerialFromPhysicalIds
# Purpose  : Extracts serial number from AAD physicalIds array.
# ------------------------------------------------------------------------------
function Get-SerialFromPhysicalIds {
    param ([array]$PhysicalIds)
    if (-not $PhysicalIds -or $PhysicalIds.Count -eq 0) { return "" }
    foreach ($e in $PhysicalIds) {
        if ($e -match '^\[SerialNumber\]:(.+)$') { return $Matches[1].Trim() }
    }
    foreach ($e in $PhysicalIds) {
        if ($e -match '^\[OrderID\]:(.+)$') { return $Matches[1].Trim() }
    }
    return ""
}


# ------------------------------------------------------------------------------
# FUNCTION : Get-DaysSince
# Purpose  : Returns days since an ISO datetime string. 9999 = never/missing.
# ------------------------------------------------------------------------------
function Get-DaysSince {
    param ([string]$DateString)
    if ([string]::IsNullOrWhiteSpace($DateString) -or $DateString -like "0001*") { return 9999 }
    try {
        $dt = [datetime]::Parse($DateString, $null,
              [System.Globalization.DateTimeStyles]::RoundtripKind)
        return [math]::Round(((Get-Date) - $dt).TotalDays, 1)
    }
    catch { return 9999 }
}


# ------------------------------------------------------------------------------
# FUNCTION : Format-Date
# Purpose  : Formats ISO datetime to readable yyyy-MM-dd HH:mm.
# ------------------------------------------------------------------------------
function Format-Date {
    param ([string]$DateString)
    if ([string]::IsNullOrWhiteSpace($DateString) -or $DateString -like "0001*") { return "Never" }
    try {
        return ([datetime]::Parse($DateString, $null,
            [System.Globalization.DateTimeStyles]::RoundtripKind)).ToString("yyyy-MM-dd HH:mm")
    }
    catch { return $DateString }
}


# ------------------------------------------------------------------------------
# FUNCTION : Resolve-DeviceRow
# Purpose  : For each AAD object found for a hostname, builds one CSV row.
#            Cross-references the AAD object to any matching Intune record
#            using AADDeviceId. Flags duplicates. Returns a PSCustomObject.
# ------------------------------------------------------------------------------
function Resolve-DeviceRow {
    param (
        [string]$Hostname,
        [PSObject]$AADObject,
        [array]$IntuneRecords,
        [int]$AADObjectCount
    )

    $IsDuplicate = ($AADObjectCount -gt 1)

    # Cross-reference to Intune using deviceId
    $LinkedIntune   = $null
    $IntuneDeviceID = "Not found in Intune"
    $PrimaryUser    = "N/A"
    $EnrolledDate   = "N/A"
    $LastSync       = "N/A"
    $DaysSinceSync  = "N/A"
    $ComplianceState = "N/A"
    $ManagementState = "N/A"
    $OwnerType      = "N/A"
    $SerialNumber   = Get-SerialFromPhysicalIds -PhysicalIds $AADObject.physicalIds

    if (-not [string]::IsNullOrWhiteSpace($AADObject.deviceId)) {
        $LinkedIntune = $IntuneRecords | Where-Object {
            $_.azureADDeviceId -and
            $_.azureADDeviceId.ToLower() -eq $AADObject.deviceId.ToLower()
        } | Select-Object -First 1
    }

    if ($LinkedIntune) {
        $IntuneDeviceID  = $LinkedIntune.id
        $PrimaryUser     = if ($LinkedIntune.userPrincipalName) { $LinkedIntune.userPrincipalName } else { "No Primary User" }
        $EnrolledDate    = Format-Date $LinkedIntune.enrolledDateTime
        $LastSync        = Format-Date $LinkedIntune.lastSyncDateTime
        $DaysSinceSync   = Get-DaysSince $LinkedIntune.lastSyncDateTime
        $ComplianceState = $LinkedIntune.complianceState
        $ManagementState = $LinkedIntune.managementState
        $OwnerType       = $LinkedIntune.managedDeviceOwnerType
        if ([string]::IsNullOrWhiteSpace($SerialNumber)) {
            $SerialNumber = $LinkedIntune.serialNumber
        }
    }

    # Build notes
    $Notes = ""
    if ($IsDuplicate) {
        $Notes = "DUPLICATE - $AADObjectCount AAD objects share this hostname"
    }
    elseif ($IntuneDeviceID -eq "Not found in Intune") {
        $Notes = "AAD object exists but no matching Intune record found"
    }
    else {
        $Notes = "OK - single entry, linked to Intune"
    }

    # DaysSinceSync display
    $DaysSinceSyncDisplay = if ($DaysSinceSync -eq 9999) { "Never" } else { $DaysSinceSync }

    return [PSCustomObject]@{
        Hostname         = $Hostname
        AADObjectID      = $AADObject.id
        AADDeviceID      = $AADObject.deviceId
        IntuneDeviceID   = $IntuneDeviceID
        SerialNumber     = $SerialNumber
        PrimaryUser      = $PrimaryUser
        EnrolledDate     = $EnrolledDate
        LastSync         = $LastSync
        DaysSinceSync    = $DaysSinceSyncDisplay
        OS               = $AADObject.operatingSystem
        OSVersion        = $AADObject.operatingSystemVersion
        ComplianceState  = $ComplianceState
        ManagementState  = $ManagementState
        OwnerType        = $OwnerType
        AccountEnabled   = $AADObject.accountEnabled
        TrustType        = $AADObject.trustType
        RegisteredDate   = Format-Date $AADObject.registrationDateTime
        LastActivity     = Format-Date $AADObject.approximateLastSignInDateTime
        IsDuplicate      = if ($IsDuplicate) { "YES" } else { "NO" }
        AADObjectCount   = $AADObjectCount
        FoundInAAD       = "YES"
        FoundInIntune    = if ($LinkedIntune) { "YES" } else { "NO" }
        Notes            = $Notes
    }
}

#endregion --------------------------------------------------------------------


#region --- MAIN --------------------------------------------------------------

# -- Validate config ----------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($TenantID) -or
    [string]::IsNullOrWhiteSpace($ClientID) -or
    [string]::IsNullOrWhiteSpace($ClientSecret)) {
    Write-Host "[ERROR]   TenantID, ClientID, and ClientSecret must all be set in the CONFIGURATION region." -ForegroundColor Red
    exit 1
}

# -- Initialise output paths --------------------------------------------------
$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$ResolvedOutput = if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $PSScriptRoot } else { $OutputFolder }
if (-not (Test-Path $ResolvedOutput)) {
    New-Item -ItemType Directory -Path $ResolvedOutput -Force | Out-Null
}

$OutputCSV      = Join-Path $ResolvedOutput ("Get-DeviceObjectIDs_" + $Timestamp + ".csv")
$BulkImportTXT  = Join-Path $ResolvedOutput ("Get-DeviceObjectIDs_" + $Timestamp + "_BulkImport.txt")
$script:LogFile = Join-Path $ResolvedOutput ("Get-DeviceObjectIDs_" + $Timestamp + ".log")

try {
    [System.IO.File]::WriteAllText(
        $script:LogFile,
        "Get-DeviceObjectIDs v1.0`r`nStarted : $(Get-Date)`r`nInput   : $InputFilePath`r`n`r`n",
        [System.Text.Encoding]::UTF8
    )
} catch { $script:LogFile = $null }

# -- Banner -------------------------------------------------------------------
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Get-DeviceObjectIDs  |  READ ONLY                            " -ForegroundColor Cyan
Write-Host "  Endpoint Engineering Team  |  Western Digital                " -ForegroundColor Cyan
Write-Host "  Resolves hostnames to AAD Object IDs + Intune Device IDs.    " -ForegroundColor Cyan
Write-Host "  Duplicates are listed. NO changes made.                      " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "" -Level BLANK
Write-Log "Input TXT       : $InputFilePath"   -Level INFO
Write-Log "Output CSV      : $OutputCSV"        -Level INFO
Write-Log "BulkImport TXT  : $BulkImportTXT"   -Level INFO
Write-Log "Log file        : $($script:LogFile)" -Level INFO

# -- STEP 1: Read input file --------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 1 - Reading Hostname List" -Level INFO
Write-Log "==========================================================" -Level SECTION

if (-not (Test-Path $InputFilePath)) {
    Write-Log "Input file not found: $InputFilePath" -Level ERROR
    Write-Log "Create DeviceList.txt in the same folder as this script." -Level ERROR
    Write-Log "One hostname per line." -Level ERROR
    exit 1
}

try {
    $RawLines = Get-Content -Path $InputFilePath -Encoding UTF8
    Write-Log "File loaded - $($RawLines.Count) raw lines." -Level SUCCESS
}
catch {
    Write-Log "Failed to read input file: $_" -Level ERROR
    exit 1
}

$Hostnames = $RawLines |
    ForEach-Object { $_.Trim() } |
    Where-Object   { -not [string]::IsNullOrWhiteSpace($_) } |
    Sort-Object    -Unique

Write-Log "Unique hostnames to process: $($Hostnames.Count)" -Level SUCCESS
foreach ($h in $Hostnames) { Write-Log "  -> $h" -Level INFO }

# -- STEP 2: Authenticate -----------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 2 - Authenticating" -Level INFO
Write-Log "==========================================================" -Level SECTION

$Token = Get-GraphAccessToken -TenantId $TenantID -ClientId $ClientID -ClientSecret $ClientSecret

# -- STEP 3: Resolve each hostname --------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 3 - Resolving Hostnames to Object IDs" -Level INFO
Write-Log "==========================================================" -Level SECTION

$AllRows        = [System.Collections.Generic.List[PSObject]]::new()
$BulkImportIDs  = [System.Collections.Generic.List[string]]::new()
$Counter        = 0
$Total          = $Hostnames.Count
$NotFoundCount  = 0
$DuplicateCount = 0
$CleanCount     = 0

foreach ($Hostname in $Hostnames) {

    $Counter++
    Write-Log "" -Level BLANK
    Write-Log "  [$Counter/$Total] $Hostname" -Level WARN

    # Search Azure AD
    Write-Log "    Searching Azure AD..." -Level INFO
    $AADObjects = Get-AADObjectsByHostname -Hostname $Hostname -AccessToken $Token
    Write-Log "    Azure AD objects found : $($AADObjects.Count)" `
        -Level $(if ($AADObjects.Count -eq 0) { "ERROR" } elseif ($AADObjects.Count -gt 1) { "WARN" } else { "SUCCESS" })

    # Search Intune
    Write-Log "    Searching Intune..." -Level INFO
    $IntuneRecords = Get-IntuneRecordsByHostname -Hostname $Hostname -AccessToken $Token
    Write-Log "    Intune records found   : $($IntuneRecords.Count)" `
        -Level $(if ($IntuneRecords.Count -eq 0) { "WARN" } elseif ($IntuneRecords.Count -gt 1) { "WARN" } else { "SUCCESS" })

    # Not found in Azure AD at all
    if ($AADObjects.Count -eq 0) {
        Write-Log "    -> NOT FOUND in Azure AD" -Level ERROR
        $NotFoundCount++
        $AllRows.Add([PSCustomObject]@{
            Hostname        = $Hostname
            AADObjectID     = "NOT FOUND"
            AADDeviceID     = ""
            IntuneDeviceID  = if ($IntuneRecords.Count -gt 0) { $IntuneRecords[0].id } else { "NOT FOUND" }
            SerialNumber    = ""
            PrimaryUser     = ""
            EnrolledDate    = ""
            LastSync        = ""
            DaysSinceSync   = ""
            OS              = ""
            OSVersion       = ""
            ComplianceState = ""
            ManagementState = ""
            OwnerType       = ""
            AccountEnabled  = ""
            TrustType       = ""
            RegisteredDate  = ""
            LastActivity    = ""
            IsDuplicate     = "N/A"
            AADObjectCount  = 0
            FoundInAAD      = "NO"
            FoundInIntune   = if ($IntuneRecords.Count -gt 0) { "YES" } else { "NO" }
            Notes           = "Hostname not found in Azure AD - cannot add to group"
        })
        continue
    }

    # Flag duplicates
    if ($AADObjects.Count -gt 1) {
        $DuplicateCount++
        Write-Log ("    -> DUPLICATE - {0} AAD objects share this hostname" -f $AADObjects.Count) -Level WARN
    }
    else {
        $CleanCount++
    }

    # Build one row per AAD object
    foreach ($AADObj in $AADObjects) {
        $Row = Resolve-DeviceRow `
               -Hostname      $Hostname `
               -AADObject     $AADObj `
               -IntuneRecords $IntuneRecords `
               -AADObjectCount $AADObjects.Count

        $AllRows.Add($Row)
        $BulkImportIDs.Add($AADObj.id)

        $DupTag = if ($Row.IsDuplicate -eq "YES") { " [DUPLICATE]" } else { "" }
        Write-Log ("    AADObjectID  : {0}{1}" -f $AADObj.id, $DupTag) -Level $(if ($Row.IsDuplicate -eq "YES") { "WARN" } else { "SUCCESS" })
        Write-Log ("    IntuneID     : {0}" -f $Row.IntuneDeviceID) -Level INFO
        Write-Log ("    Serial       : {0}" -f $Row.SerialNumber)   -Level INFO
        Write-Log ("    User         : {0}" -f $Row.PrimaryUser)    -Level INFO
        Write-Log ("    Last Sync    : {0} ({1} days ago)" -f $Row.LastSync, $Row.DaysSinceSync) -Level INFO
    }

    Start-Sleep -Milliseconds 200
}

# -- STEP 4: Export CSV -------------------------------------------------------
Write-Log "" -Level BLANK
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 4 - Exporting CSV" -Level INFO
Write-Log "==========================================================" -Level SECTION

try {
    $AllRows | Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding UTF8
    $SizeMB   = (((Get-Item $OutputCSV).Length) / 1MB).ToString("0.00")
    $RowCount  = ($AllRows | Measure-Object).Count
    Write-Log "CSV exported successfully." -Level SUCCESS
    Write-Log "  Path : $OutputCSV"        -Level INFO
    Write-Log "  Rows : $RowCount   |   Size: $SizeMB MB" -Level INFO
}
catch { Write-Log "CSV export failed: $_" -Level ERROR }

# -- STEP 5: Export BulkImport TXT --------------------------------------------
Write-Log "" -Level BLANK
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 5 - Exporting BulkImport TXT" -Level INFO
Write-Log "==========================================================" -Level SECTION

try {
    if ($BulkImportIDs.Count -gt 0) {
        [System.IO.File]::WriteAllLines(
            $BulkImportTXT,
            $BulkImportIDs.ToArray(),
            [System.Text.Encoding]::UTF8
        )
        Write-Log "BulkImport TXT exported successfully." -Level SUCCESS
        Write-Log "  Path       : $BulkImportTXT" -Level INFO
        Write-Log "  Object IDs : $($BulkImportIDs.Count)" -Level INFO
        Write-Log "  NOTE: Duplicate hostnames will have multiple IDs in this file." -Level WARN
        Write-Log "        Review the CSV and remove any IDs you do not want to add." -Level WARN
    }
    else {
        Write-Log "No Azure AD Object IDs found - BulkImport TXT not created." -Level WARN
    }
}
catch { Write-Log "BulkImport TXT export failed: $_" -Level ERROR }

# -- Summary ------------------------------------------------------------------
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  COMPLETE - READ ONLY - NO CHANGES MADE                       " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "  Hostnames processed     : $Total"         -Level INFO
Write-Log "  Clean  (single entry)   : $CleanCount"    -Level $(if ($CleanCount    -gt 0) { "SUCCESS" } else { "INFO" })
Write-Log "  Duplicates (2+ entries) : $DuplicateCount" -Level $(if ($DuplicateCount -gt 0) { "WARN"    } else { "INFO" })
Write-Log "  Not found in Azure AD   : $NotFoundCount"  -Level $(if ($NotFoundCount  -gt 0) { "ERROR"   } else { "INFO" })
Write-Log "  Total Object IDs in TXT : $($BulkImportIDs.Count)" -Level INFO
Write-Log "" -Level BLANK
Write-Log "  -- How to use the BulkImport TXT --------------------------------" -Level SECTION
Write-Log "  1. Review CSV - check duplicate rows before importing"          -Level INFO
Write-Log "  2. Remove any Object IDs from TXT you do not want added"        -Level INFO
Write-Log "  3. Entra portal - Groups - your group - Members - Bulk Add"     -Level INFO
Write-Log "  4. Upload the BulkImport TXT file"                              -Level INFO
Write-Log "" -Level BLANK
Write-Log "  -- Output files --------------------------------------------------" -Level SECTION
Write-Log "  CSV            : $OutputCSV"        -Level INFO
Write-Log "  BulkImport TXT : $BulkImportTXT"   -Level INFO
Write-Log "  Log file       : $($script:LogFile)" -Level INFO
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

#endregion --------------------------------------------------------------------
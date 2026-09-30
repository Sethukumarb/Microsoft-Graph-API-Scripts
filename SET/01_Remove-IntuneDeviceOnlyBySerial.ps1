#Requires -Version 5.1
# ==============================================================================
# Script Name  : Remove-IntuneDeviceOnlyBySerial.ps1
# Description  : For each input serial number:
#                  - Finds ALL Intune managed device records
#                  - Finds ALL Autopilot device records  (report only, no delete)
#                  - Finds ALL Entra ID (Azure AD) device objects (report only, no delete)
#                  - DELETES Intune managed device records ONLY
#
#                WHAT THIS SCRIPT DELETES:
#                  Intune Managed Device   - DELETE /deviceManagement/managedDevices/{id}
#
#                WHAT THIS SCRIPT SKIPS (report only):
#                  Autopilot Device        - NOT touched. Listed in Sheet 1 only.
#                  Entra ID Device Object  - NOT touched. Listed in Sheet 1 + Sheet 3.
#                                           Sheet 3 is formatted for handoff to Entra team.
#
#                LOOKUP STRATEGY:
#                  Serial -> bulk pull all Intune devices   -> hashtable lookup
#                  Serial -> bulk pull all Autopilot devices -> hashtable lookup
#                  Serial -> Entra ID lookup per Intune record (azureADDeviceId)
#                  No $filter used on managedDevices or windowsAutopilotDeviceIdentities
#                  (both endpoints return HTTP 500 with $filter/$select + $top)
#
#                DRY RUN MODE ($DryRun = $true - DEFAULT):
#                  Shows exactly what would be deleted. No DELETE calls made.
#                  Review output Excel then set $DryRun = $false to execute.
#
#                INPUT FILE:
#                  remove_inputserialnumbers.txt - same folder as this script
#                  One serial number per line. # comments and blank lines ignored.
#
#                OUTPUT FILE (saved to $PSScriptRoot):
#                  RemoveIntuneDevices_[timestamp].xlsx   - 3-sheet Excel workbook
#                  RemoveIntuneDevices_[timestamp].log    - detailed run log
#                  RemoveIntuneDevices_Transcript_[timestamp].log - PS transcript
#
#                EXCEL WORKBOOK STRUCTURE:
#                  Sheet 1 - Full Discovery
#                    All records found for each input serial.
#                    RecordType: Intune | Autopilot | EntraID
#                    Columns: SerialNumber, Hostname, RecordType, IntuneDeviceID,
#                             AzureADDeviceID, AzureADObjectID, AutopilotDeviceID,
#                             PlannedAction, Manufacturer, Model, OS, UPN, Timestamp
#
#                  Sheet 2 - Intune Deletions
#                    Only Intune records. Shows delete result.
#                    Columns: SerialNumber, Hostname, IntuneDeviceID,
#                             DeleteStatus, DeleteNote, Manufacturer, Model, Timestamp
#
#                  Sheet 3 - Entra ID Handoff
#                    Entra ID objects found but NOT deleted.
#                    Formatted for sharing with Entra/Azure AD team.
#                    Columns: SerialNumber, Hostname, AzureADDeviceID, AzureADObjectID,
#                             Status, Note, Timestamp
#
#                DELETE STATUS VALUES:
#                  DELETED   - Successfully deleted
#                  DRY RUN   - Would be deleted (DryRun = $true)
#                  FAILED    - Delete call failed (see DeleteNote)
#                  NOT FOUND - No Intune record for this serial
#
# Requires     : ImportExcel module  (Install-Module ImportExcel -Scope CurrentUser)
#
# Author       : Sethu Kumar B
# Version      : 1.0
# Created Date : 2026-05-14
#
# Permissions required (admin consent granted):
#   DeviceManagementManagedDevices.ReadWrite.All  - read + delete Intune devices
#   DeviceManagementServiceConfig.ReadWrite.All   - read Autopilot devices
#   Device.Read.All                               - read Entra ID device objects
# ==============================================================================


#region --- CONFIGURATION -------------------------------------------------------

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

# -- DRY RUN -------------------------------------------------------------------

# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# WARNING: $DryRun = $false WILL PERMANENTLY DELETE INTUNE DEVICE RECORDS.
#          THIS CANNOT BE UNDONE.
#          Always run with $DryRun = $true first and review the Excel output
#          before setting $DryRun = $false.
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

$DryRun = $true

# -- INPUT FILE ----------------------------------------------------------------
$InputFileName = "remove_inputserialnumbers.txt"
$InputPath     = Join-Path $PSScriptRoot $InputFileName

# -- THROTTLE ------------------------------------------------------------------
$MaxRetries    = 5

# -- BATCH CAP ----------------------------------------------------------------
# Abort if input exceeds this count. Safety gate before live deletions.
# 0 = no limit.
$MaxBatchSize  = 1

#endregion ----------------------------------------------------------------------


#region --- INIT ----------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$DryLabel       = if ($DryRun) { "[DRY RUN] " } else { "" }
$OutputFile     = Join-Path $PSScriptRoot "${DryLabel}RemoveIntuneDevices_$Timestamp.xlsx"
$script:LogFile = Join-Path $PSScriptRoot "${DryLabel}RemoveIntuneDevices_$Timestamp.log"
$TranscriptFile = Join-Path $PSScriptRoot "${DryLabel}RemoveIntuneDevices_Transcript_$Timestamp.log"

try { Start-Transcript -Path $TranscriptFile -Force | Out-Null } catch { }

try {
    [System.IO.File]::WriteAllText($script:LogFile,
        "Remove-IntuneDeviceBySerial v1.0`r`nStarted: $(Get-Date)`r`nDryRun: $DryRun`r`n`r`n",
        [System.Text.Encoding]::UTF8)
} catch { $script:LogFile = $null }

# Verify ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[ERROR] ImportExcel module not found." -ForegroundColor Red
    Write-Host "        Run: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}
Import-Module ImportExcel -ErrorAction Stop

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
    catch {
        $ErrCode = $null
        if ($null -ne $_.Exception.PSObject.Properties['Response'] -and $null -ne $_.Exception.Response) {
            $ErrCode = $_.Exception.Response.StatusCode.value__
        }
        $ErrMsg = if ($ErrCode) { "HTTP $ErrCode - $($_.Exception.Message)" } else { $_.Exception.Message }
        Write-Log "Authentication failed: $ErrMsg" -Level ERROR
        exit 1
    }
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
                $Uri     = if ($null -ne $r.PSObject.Properties['@odata.nextLink']) { $r.'@odata.nextLink' } else { $null }
                $Success = $true
            }
            catch {
                $Code = $null
                if ($null -ne $_.Exception.PSObject.Properties['Response'] -and $null -ne $_.Exception.Response) {
                    $Code = $_.Exception.Response.StatusCode.value__
                }
                if ($Code -eq 429) {
                    $Wait = 60
                    try {
                        $HV = $_.Exception.Response.Headers.GetValues("Retry-After")
                        if ($HV -and $HV.Count -gt 0) { $Wait = [int]$HV[0] }
                    } catch { }
                    $Wait += Get-Random -Minimum 1 -Maximum 10
                    if ($Attempt -lt $MaxRetries) {
                        Write-Log "  429 throttled - waiting ${Wait}s (retry $Attempt/$MaxRetries)..." -Level WARN
                        Start-Sleep -Seconds $Wait
                    }
                    else {
                        Write-Log "  429 persisted after $MaxRetries attempts. Skipping page." -Level WARN
                        $Uri     = $null
                        $Success = $true
                    }
                }
                else {
                    $ErrDetail = if ($Code) { "HTTP $Code" } else { "No HTTP response (network/timeout/auth error)" }
                    Write-Log "  Page $Page failed - $ErrDetail : $($_.Exception.Message)" -Level ERROR
                    $Uri     = $null
                    $Success = $true
                }
            }
        } while (-not $Success -and $Attempt -lt $MaxRetries)

    } while ($Uri)

    Write-Log "$Label - $Total total records." -Level SUCCESS
    return $AllRecords.ToArray()
}


# Lookup Entra ID device object by azureADDeviceId (GUID from Intune record).
# Returns object with both id (ObjectID) and deviceId (AzureADDeviceID), or $null.
function Get-EntraDevice {
    param ([string]$AzureADDeviceId, [string]$AccessToken)
    if ([string]::IsNullOrWhiteSpace($AzureADDeviceId) -or $AzureADDeviceId -eq "N/A") { return $null }
    $Headers = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $Uri     = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$AzureADDeviceId'" +
               "&`$select=id,displayName,deviceId,operatingSystem,isCompliant,isManaged"
    try {
        $r = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
        if ($r.value -and $r.value.Count -gt 0) { return $r.value[0] }
        return $null
    }
    catch { return $null }
}


function Invoke-GraphDelete {
    param ([string]$Uri, [string]$AccessToken)
    $Headers = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    try {
        Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        $Code = $null
        if ($null -ne $_.Exception.PSObject.Properties['Response'] -and $null -ne $_.Exception.Response) {
            $Code = $_.Exception.Response.StatusCode.value__
        }
        $ErrDetail = if ($Code) { "HTTP $Code" } else { "No HTTP response (network/timeout/auth error)" }
        throw "$ErrDetail : $($_.Exception.Message)"
    }
}


# Build styled Excel workbook with 3 sheets using ImportExcel
function Export-ResultsToExcel {
    param (
        [string]$Path,
        [System.Collections.Generic.List[PSObject]]$Sheet1Data,
        [System.Collections.Generic.List[PSObject]]$Sheet2Data,
        [System.Collections.Generic.List[PSObject]]$Sheet3Data,
        [bool]$IsDryRun
    )

    # Remove existing file if present
    if (Test-Path $Path) { Remove-Item $Path -Force }

    $ModeLabel = if ($IsDryRun) { "DRY RUN" } else { "LIVE" }

    # ---- SHEET 1 - Full Discovery ----
    $S1 = if ($Sheet1Data.Count -gt 0) { $Sheet1Data } else {
        @([PSCustomObject]@{ Note = "No records found" })
    }
    $ExcelPkg = $S1 | Export-Excel -Path $Path -WorksheetName "Sheet1 - Full Discovery" `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru `
        -TableName "DiscoveryTable" -TableStyle Medium2

    $WS1 = $ExcelPkg.Workbook.Worksheets["Sheet1 - Full Discovery"]

    # Color-code Action column (col 8 = PlannedAction)
    if ($Sheet1Data.Count -gt 0) {
        for ($row = 2; $row -le ($Sheet1Data.Count + 1); $row++) {
            $actionCell = $WS1.Cells[$row, 8]
            $actionVal  = [string]$actionCell.Value
            switch ($actionVal) {
                "WILL DELETE" {
                    $actionCell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $actionCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $actionCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(192, 0, 0))
                }
                "SKIPPED"     {
                    $actionCell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $actionCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $actionCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
                }
                "NOT FOUND"   {
                    $actionCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $actionCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 156))
                }
            }
        }
    }

    # Tab color Sheet 1
    $WS1.TabColor = [System.Drawing.Color]::FromArgb(31, 78, 121)

    # ---- SHEET 2 - Intune Deletions ----
    $S2 = if ($Sheet2Data.Count -gt 0) { $Sheet2Data } else {
        @([PSCustomObject]@{ Note = "No Intune records found or deleted" })
    }
    $ExcelPkg = $S2 | Export-Excel -ExcelPackage $ExcelPkg -WorksheetName "Sheet2 - Intune Deletions" `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru `
        -TableName "IntuneDeleteTable" -TableStyle Medium3

    $WS2 = $ExcelPkg.Workbook.Worksheets["Sheet2 - Intune Deletions"]

    # Color-code DeleteStatus column (col 4)
    if ($Sheet2Data.Count -gt 0) {
        for ($row = 2; $row -le ($Sheet2Data.Count + 1); $row++) {
            $statusCell = $WS2.Cells[$row, 4]
            $statusVal  = [string]$statusCell.Value
            switch ($statusVal) {
                "DELETED"   {
                    $statusCell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $statusCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $statusCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))
                }
                "DRY RUN"   {
                    $statusCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $statusCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 156))
                }
                "FAILED"    {
                    $statusCell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $statusCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $statusCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(192, 0, 0))
                }
                "NOT FOUND" {
                    $statusCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $statusCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(242, 242, 242))
                }
            }
        }
    }

    $WS2.TabColor = [System.Drawing.Color]::FromArgb(192, 0, 0)

    # ---- SHEET 3 - Entra ID Handoff ----
    $S3 = if ($Sheet3Data.Count -gt 0) { $Sheet3Data } else {
        @([PSCustomObject]@{ Note = "No Entra ID objects found for input serials" })
    }
    $ExcelPkg = $S3 | Export-Excel -ExcelPackage $ExcelPkg -WorksheetName "Sheet3 - Entra ID Handoff" `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru `
        -TableName "EntraHandoffTable" -TableStyle Medium6

    $WS3 = $ExcelPkg.Workbook.Worksheets["Sheet3 - Entra ID Handoff"]
    $WS3.TabColor = [System.Drawing.Color]::FromArgb(0, 70, 127)

    # Add header note row above data on Sheet 3 (insert row 1, push data down)
    $WS3.InsertRow(1, 1)
    $WS3.Cells[1, 1].Value = "Entra ID Handoff - $(Get-Date -Format 'yyyy-MM-dd') | Mode: $ModeLabel | These devices were NOT deleted by the script. Please delete manually from Entra portal."
    $WS3.Cells[1, 1].Style.Font.Bold = $true
    $WS3.Cells[1, 1].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 70, 127))
    $merged = $WS3.Cells[1, 1, 1, 7]
    $merged.Merge = $true

    # Save and close
    Close-ExcelPackage $ExcelPkg -Show:$false

    Write-Log "Excel workbook saved: $Path" -Level SUCCESS
}

#endregion ----------------------------------------------------------------------


#region --- MAIN ----------------------------------------------------------------

# Banner
Write-Host ""
if ($DryRun) {
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  Remove-IntuneDeviceBySerial v1.0  |  DRY RUN                " -ForegroundColor Cyan
    Write-Host "  Sethu Kumar B                                                 " -ForegroundColor Cyan
    Write-Host "  Deletes: Intune only. Autopilot + Entra ID: report only.     " -ForegroundColor Cyan
    Write-Host "  Review Excel output then set DryRun = false to execute.      " -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
}
else {
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host "  Remove-IntuneDeviceBySerial v1.0  |  LIVE - DELETES NOW     " -ForegroundColor Red
    Write-Host "  Sethu Kumar B                                                 " -ForegroundColor Red
    Write-Host "  !! INTUNE DEVICE RECORDS WILL BE PERMANENTLY DELETED !!      " -ForegroundColor Red
    Write-Host "  Autopilot and Entra ID records will NOT be touched.           " -ForegroundColor Red
    Write-Host "================================================================" -ForegroundColor Red
}
Write-Log "" -Level BLANK
Write-Log "Mode         : $(if ($DryRun) {'DRY RUN - no deletes'} else {'LIVE - Intune deletes will execute'})" `
          -Level $(if ($DryRun) {"INFO"} else {"WARN"})
Write-Log "Input file   : $InputPath"         -Level INFO
Write-Log "Output Excel : $OutputFile"        -Level INFO
Write-Log "Log file     : $($script:LogFile)" -Level INFO
Write-Log "Transcript   : $TranscriptFile"    -Level INFO
Write-Log "" -Level BLANK


# -- Step 1: Read serial numbers -----------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 1 - Reading Input File" -Level INFO
Write-Log "==========================================================" -Level SECTION

if (-not (Test-Path $InputPath)) {
    Write-Log "Input file not found: $InputPath" -Level ERROR
    Write-Log "Create $InputFileName with one serial number per line." -Level ERROR
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
if ($MaxBatchSize -gt 0 -and $Serials.Count -gt $MaxBatchSize) {
    Write-Log "BATCH CAP EXCEEDED: $($Serials.Count) serials in input, limit is $MaxBatchSize." -Level ERROR
    Write-Log "Reduce input file or increase MaxBatchSize in config. Exiting." -Level ERROR
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}
foreach ($s in $Serials) { Write-Log "  -> $s" -Level INFO }
Write-Log "" -Level BLANK


# -- Step 2: Authenticate ------------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 2 - Authenticating" -Level INFO
Write-Log "==========================================================" -Level SECTION

$Token = Get-GraphToken -TenantId $TenantID -ClientId $ClientID -ClientSecret $ClientSecret


# -- Step 3: Bulk pull Intune Windows devices into hashtable -------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 3 - Building Intune Device Lookup Table (Windows only)" -Level INFO
Write-Log "==========================================================" -Level SECTION

$IntuneUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices" +
             "?`$filter=operatingSystem eq 'Windows'" +
             "&`$select=id,serialNumber,deviceName,manufacturer,model," +
             "azureADDeviceId,userPrincipalName,operatingSystem&`$top=1000"

$AllIntune = Invoke-GraphGetAllPages -InitialUri $IntuneUri -AccessToken $Token `
             -Label "Intune Windows managed devices" -MaxRetries $MaxRetries

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
$TotalIntuneRecords = ($IntuneBySerial.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
Write-Log "Intune lookup table: $($IntuneBySerial.Count) unique serials, $TotalIntuneRecords total records." -Level SUCCESS
Write-Log "" -Level BLANK


# -- Step 4: Bulk pull Autopilot devices into hashtable -----------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 4 - Building Autopilot Device Lookup Table (report only)" -Level INFO
Write-Log "==========================================================" -Level SECTION

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
$TotalAPRecords = ($APBySerial.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
Write-Log "Autopilot lookup table: $($APBySerial.Count) unique serials, $TotalAPRecords total records." -Level SUCCESS
Write-Log "" -Level BLANK


# -- Step 5: Process each serial -----------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 5 - Processing Serials" -Level INFO
Write-Log "==========================================================" -Level SECTION

if ($DryRun) {
    Write-Log "DRY RUN active - no DELETE calls will be made." -Level WARN
} else {
    Write-Log "LIVE mode - Intune DELETE calls will execute." -Level WARN
}
Write-Log "" -Level BLANK

# Output lists for 3 sheets
$Sheet1 = [System.Collections.Generic.List[PSObject]]::new()   # Full Discovery
$Sheet2 = [System.Collections.Generic.List[PSObject]]::new()   # Intune Deletions
$Sheet3 = [System.Collections.Generic.List[PSObject]]::new()   # Entra ID Handoff

$Counter          = 0
$TotalSerials     = $Serials.Count
$DeletedIntune    = 0
$FailedCount      = 0
$NotFoundIntune   = 0
$EntraFoundTotal  = 0

foreach ($Serial in $Serials) {
    $Counter++
    $SerialKey = $Serial.ToLower().Trim()
    $RunTS     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($IntuneBySerial.ContainsKey($SerialKey)) { $IntuneDevices = $IntuneBySerial[$SerialKey] } else { $IntuneDevices = @() }
    if ($APBySerial.ContainsKey($SerialKey))     { $APDevices     = $APBySerial[$SerialKey] }     else { $APDevices     = @() }
    $IntuneDevices = @($IntuneDevices)
    $APDevices     = @($APDevices)

    Write-Log "  [$Counter/$TotalSerials] Serial: $Serial  |  Intune: $($IntuneDevices.Count)  Autopilot: $($APDevices.Count)" `
              -Level INFO

    # Shared device info - pull from first record available
    $Manufacturer = "N/A"; $Model = "N/A"
    if ($IntuneDevices.Count -gt 0) {
        $Manufacturer = if ($IntuneDevices[0].manufacturer) { [string]$IntuneDevices[0].manufacturer } else { "N/A" }
        $Model        = if ($IntuneDevices[0].model)        { [string]$IntuneDevices[0].model }        else { "N/A" }
    } elseif ($APDevices.Count -gt 0) {
        $Manufacturer = if ($APDevices[0].manufacturer) { [string]$APDevices[0].manufacturer } else { "N/A" }
        $Model        = if ($APDevices[0].model)        { [string]$APDevices[0].model }        else { "N/A" }
    }

    # ---- Intune records ----
    if ($IntuneDevices.Count -eq 0) {
        $NotFoundIntune++
        Write-Log "    [Intune] NOT FOUND" -Level WARN

        # Sheet 1 - not found row
        $Sheet1.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            Hostname          = "N/A"
            RecordType        = "Intune"
            IntuneDeviceID    = "NOT FOUND"
            AzureADDeviceID   = "N/A"
            AzureADObjectID   = "N/A"
            AutopilotDeviceID = "N/A"
            PlannedAction     = "NOT FOUND"
            Manufacturer      = $Manufacturer
            Model             = $Model
            OS                = "N/A"
            UPN               = "N/A"
            Timestamp         = $RunTS
        })

        # Sheet 2 - not found row
        $Sheet2.Add([PSCustomObject]@{
            SerialNumber   = $Serial
            Hostname       = "N/A"
            IntuneDeviceID = "NOT FOUND"
            DeleteStatus   = "NOT FOUND"
            DeleteNote     = "No Intune managed device record for this serial"
            Manufacturer   = $Manufacturer
            Model          = $Model
            Timestamp      = $RunTS
        })
    }

    $IntuneIdx = 0
    foreach ($IntuneDevice in $IntuneDevices) {
        $IntuneIdx++
        $IntuneId   = [string]$IntuneDevice.id
        $Hostname   = if ($IntuneDevice.deviceName)      { [string]$IntuneDevice.deviceName }      else { "N/A" }
        $AzureADId  = if ($IntuneDevice.azureADDeviceId) { [string]$IntuneDevice.azureADDeviceId } else { "N/A" }
        $UPN        = if ($IntuneDevice.userPrincipalName){ [string]$IntuneDevice.userPrincipalName } else { "N/A" }
        $OS         = if ($IntuneDevice.operatingSystem)  { [string]$IntuneDevice.operatingSystem }  else { "N/A" }
        $DupeNote   = if ($IntuneDevices.Count -gt 1) { " (Duplicate $IntuneIdx of $($IntuneDevices.Count))" } else { "" }

        $DeleteStatus = ""; $DeleteNote = ""

        if ($DryRun) {
            Write-Log "    [Intune$DupeNote] FOUND $Hostname ($IntuneId) - DRY RUN" -Level WARN
            $DeleteStatus = "DRY RUN"
            $DeleteNote   = "Would be deleted - set DryRun = false to execute"
            $DeletedIntune++
        }
        else {
            try {
                $DelUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$IntuneId"
                Invoke-GraphDelete -Uri $DelUri -AccessToken $Token | Out-Null
                Write-Log "    [Intune$DupeNote] DELETED $Hostname ($IntuneId)" -Level SUCCESS
                $DeleteStatus = "DELETED"
                $DeleteNote   = "Successfully deleted"
                $DeletedIntune++
            }
            catch {
                Write-Log "    [Intune$DupeNote] FAILED $Hostname ($IntuneId) - $_" -Level ERROR
                $DeleteStatus = "FAILED"
                $DeleteNote   = [string]$_
                $FailedCount++
            }
            Start-Sleep -Seconds 1
        }

        # Sheet 1 - Intune row
        $Sheet1.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            Hostname          = $Hostname
            RecordType        = "Intune$DupeNote"
            IntuneDeviceID    = $IntuneId
            AzureADDeviceID   = $AzureADId
            AzureADObjectID   = "N/A"
            AutopilotDeviceID = "N/A"
            PlannedAction     = if ($DryRun) { "WILL DELETE" } else { $DeleteStatus }
            Manufacturer      = $Manufacturer
            Model             = $Model
            OS                = $OS
            UPN               = $UPN
            Timestamp         = $RunTS
        })

        # Sheet 2 - Intune deletion row
        $Sheet2.Add([PSCustomObject]@{
            SerialNumber   = $Serial
            Hostname       = $Hostname
            IntuneDeviceID = $IntuneId
            DeleteStatus   = $DeleteStatus
            DeleteNote     = $DeleteNote
            Manufacturer   = $Manufacturer
            Model          = $Model
            Timestamp      = $RunTS
        })

        # ---- Entra ID lookup per Intune record ----
        if ($AzureADId -ne "N/A") {
            Write-Log "    [EntraID] Looking up $AzureADId ..." -Level INFO
            $EntraObj = Get-EntraDevice -AzureADDeviceId $AzureADId -AccessToken $Token
            if ($EntraObj) {
                $EntraFoundTotal++
                $EntraObjId = [string]$EntraObj.id
                Write-Log "    [EntraID] FOUND - ObjectID: $EntraObjId - SKIPPED (no delete)" -Level INFO

                # Sheet 1 - EntraID row
                $Sheet1.Add([PSCustomObject]@{
                    SerialNumber      = $Serial
                    Hostname          = $Hostname
                    RecordType        = "EntraID"
                    IntuneDeviceID    = $IntuneId
                    AzureADDeviceID   = $AzureADId
                    AzureADObjectID   = $EntraObjId
                    AutopilotDeviceID = "N/A"
                    PlannedAction     = "SKIPPED"
                    Manufacturer      = $Manufacturer
                    Model             = $Model
                    OS                = if ($EntraObj.operatingSystem) { [string]$EntraObj.operatingSystem } else { $OS }
                    UPN               = $UPN
                    Timestamp         = $RunTS
                })

                # Sheet 3 - Entra ID handoff row
                $Sheet3.Add([PSCustomObject]@{
                    SerialNumber    = $Serial
                    Hostname        = $Hostname
                    AzureADDeviceID = $AzureADId
                    AzureADObjectID = $EntraObjId
                    Status          = "SKIPPED - Not deleted by script"
                    Note            = "Delete manually from Entra portal. Object ID is the key identifier for deletion."
                    Timestamp       = $RunTS
                })
            }
            else {
                Write-Log "    [EntraID] NOT FOUND for $AzureADId (no permission or already deleted)" -Level WARN

                # Sheet 1 - EntraID not found row
                $Sheet1.Add([PSCustomObject]@{
                    SerialNumber      = $Serial
                    Hostname          = $Hostname
                    RecordType        = "EntraID"
                    IntuneDeviceID    = $IntuneId
                    AzureADDeviceID   = $AzureADId
                    AzureADObjectID   = "NOT FOUND"
                    AutopilotDeviceID = "N/A"
                    PlannedAction     = "NOT FOUND"
                    Manufacturer      = $Manufacturer
                    Model             = $Model
                    OS                = $OS
                    UPN               = $UPN
                    Timestamp         = $RunTS
                })
            }
        }
        else {
            Write-Log "    [EntraID] No azureADDeviceId on Intune record - skipping lookup" -Level WARN
        }
    }

    if (-not $DryRun -and $IntuneDevices.Count -gt 0) { Start-Sleep -Seconds 2 }

    # ---- Autopilot records (report only - no delete) ----
    if ($APDevices.Count -eq 0) {
        Write-Log "    [Autopilot] NOT FOUND in Autopilot registry" -Level WARN
        $Sheet1.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            Hostname          = "N/A"
            RecordType        = "Autopilot"
            IntuneDeviceID    = "N/A"
            AzureADDeviceID   = "N/A"
            AzureADObjectID   = "N/A"
            AutopilotDeviceID = "NOT FOUND"
            PlannedAction     = "NOT FOUND"
            Manufacturer      = $Manufacturer
            Model             = $Model
            OS                = "N/A"
            UPN               = "N/A"
            Timestamp         = $RunTS
        })
    }

    $APIdx = 0
    foreach ($APDevice in $APDevices) {
        $APIdx++
        $APId      = [string]$APDevice.id
        $APModel   = if ($APDevice.model) { [string]$APDevice.model } else { $Model }
        $APMfr     = if ($APDevice.manufacturer) { [string]$APDevice.manufacturer } else { $Manufacturer }
        $DupeNote  = if ($APDevices.Count -gt 1) { " (Duplicate $APIdx of $($APDevices.Count))" } else { "" }

        Write-Log "    [Autopilot$DupeNote] FOUND $APId - SKIPPED (not deleted by design)" -Level INFO

        $Sheet1.Add([PSCustomObject]@{
            SerialNumber      = $Serial
            Hostname          = "N/A"
            RecordType        = "Autopilot$DupeNote"
            IntuneDeviceID    = "N/A"
            AzureADDeviceID   = "N/A"
            AzureADObjectID   = "N/A"
            AutopilotDeviceID = $APId
            PlannedAction     = "SKIPPED"
            Manufacturer      = $APMfr
            Model             = $APModel
            OS                = "N/A"
            UPN               = "N/A"
            Timestamp         = $RunTS
        })
    }

    Write-Log "" -Level BLANK
}


# -- Step 6: Export Excel -------------------------------------------------------
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 6 - Exporting Excel Workbook" -Level INFO
Write-Log "==========================================================" -Level SECTION

try {
    Export-ResultsToExcel -Path $OutputFile `
        -Sheet1Data $Sheet1 -Sheet2Data $Sheet2 -Sheet3Data $Sheet3 `
        -IsDryRun $DryRun
    $SizeMB = (((Get-Item $OutputFile).Length) / 1MB).ToString("0.00")
    Write-Log "Excel exported successfully." -Level SUCCESS
    Write-Log "  Path         : $OutputFile"        -Level INFO
    Write-Log "  Size         : $SizeMB MB"         -Level INFO
    Write-Log "  Sheet 1 rows : $($Sheet1.Count)"   -Level INFO
    Write-Log "  Sheet 2 rows : $($Sheet2.Count)"   -Level INFO
    Write-Log "  Sheet 3 rows : $($Sheet3.Count)"   -Level INFO
}
catch { Write-Log "Excel export failed: $_" -Level ERROR }


# -- Summary -------------------------------------------------------------------
Write-Host ""
Write-Host "================================================================" -ForegroundColor $(if ($DryRun) {"Cyan"} else {"Green"})
Write-Host "  $(if ($DryRun) {'DRY RUN COMPLETE - NO CHANGES MADE'} else {'COMPLETE'})" -ForegroundColor $(if ($DryRun) {"Cyan"} else {"Green"})
Write-Host "================================================================" -ForegroundColor $(if ($DryRun) {"Cyan"} else {"Green"})

Write-Log "  Mode                   : $(if ($DryRun) {'DRY RUN'} else {'LIVE'})" `
          -Level $(if ($DryRun) {"INFO"} else {"WARN"})
Write-Log "  Total serials          : $TotalSerials"    -Level INFO
Write-Log "  Intune deleted         : $DeletedIntune"   -Level $(if ($DeletedIntune   -gt 0) {"SUCCESS"} else {"INFO"})
Write-Log "  Intune not found       : $NotFoundIntune"  -Level $(if ($NotFoundIntune  -gt 0) {"WARN"}    else {"INFO"})
Write-Log "  Failures               : $FailedCount"     -Level $(if ($FailedCount     -gt 0) {"ERROR"}   else {"INFO"})
Write-Log "  Entra ID found (Sheet3): $EntraFoundTotal" -Level $(if ($EntraFoundTotal -gt 0) {"WARN"}    else {"INFO"})
Write-Log "  Autopilot              : NOT TOUCHED (report only in Sheet 1)" -Level INFO
Write-Log "" -Level BLANK
if ($DryRun) {
    Write-Log "  Review Excel Sheet 2 then set DryRun = false to execute deletions." -Level WARN
    Write-Log "  Share Sheet 3 with Entra ID team for manual cleanup." -Level INFO
}
Write-Log "  Output Excel : $OutputFile"        -Level INFO
Write-Log "  Log file     : $($script:LogFile)" -Level INFO
Write-Log "  Transcript   : $TranscriptFile"    -Level INFO
Write-Host "================================================================" -ForegroundColor $(if ($DryRun) {"Cyan"} else {"Green"})
Write-Host ""

# Clear credentials
$ClientSecret = ""
$Token        = ""

try { Stop-Transcript | Out-Null } catch { }

#endregion ----------------------------------------------------------------------
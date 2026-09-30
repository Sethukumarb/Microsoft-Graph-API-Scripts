#Requires -Version 5.1
# ==============================================================================
# Script Name  : Get-UserDetailsByUPN.ps1
# Description  : Fetches full user profile details from Azure AD / Entra ID
#                via Microsoft Graph API using User Principal Name (UPN) as
#                the sole lookup key.
#
#                LOOKUP STRATEGY:
#                  Single-pass exact UPN match only.
#                    $filter=userPrincipalName eq 'user@domain.com'
#                    Post-validated with -ieq (case-insensitive exact match)
#                  ConsistencyLevel: eventual + $count=true used for reliable
#                  Graph filter behaviour across all account types.
#
#                INPUT FILE:
#                  UPNs.txt - same folder as this script
#                  One UPN per line. Blank lines and # comments ignored.
#
#                  Accepted input format:
#                    john.doe@contoso.com
#                    jane.smith@contoso.com
#
#                OUTPUT FILES (saved to $PSScriptRoot):
#                  UserDetailsByUPN_[timestamp].csv  - full results
#                  UserDetailsByUPN_[timestamp].log  - run log
#
#                CSV COLUMNS:
#                  InputUPN, LookupStatus,
#                  DisplayName, UserPrincipalName, Mail,
#                  Department, JobTitle, OfficeLocation,
#                  AccountEnabled,
#                  CompanyName, UsageLocation,
#                  Country, State, City, StreetAddress,
#                  UserId
#
#                LOOKUP STATUS VALUES:
#                  FOUND     - UPN matched, details resolved
#                  NOT FOUND - no user found for this UPN
#                  ERROR     - API error during lookup
#
#                READ ONLY: Only GET requests. No changes made anywhere.
#
# Author       : Sethu Kumar B
# Version      : 1.0
# Created Date : 2026-04-28
# Last Modified: 2026-04-28
#
# Requirements :
#   - Azure AD App Registration (READ-ONLY)
#   - Graph API Application Permission (admin consent granted):
#       User.Read.All  - search and read Azure AD user profiles
#   - PowerShell 5.1 or later
#   - TLS 1.2 enabled
#
# Change Log   :
#   v1.0 - 2026-04-28 - Sethu Kumar B - Initial release. Single-pass exact UPN
#                        lookup with full profile field retrieval including
#                        location, company, and address attributes.
# ==============================================================================


#region --- CONFIGURATION -------------------------------------------------------

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

# Input file - must be in the same folder as this script
$InputFileName = "UPNs.txt"
$InputPath     = Join-Path $PSScriptRoot $InputFileName

#endregion ----------------------------------------------------------------------


#region --- INIT ----------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile     = Join-Path $PSScriptRoot "UserDetailsByUPN_$Timestamp.csv"
$script:LogFile = Join-Path $PSScriptRoot "UserDetailsByUPN_$Timestamp.log"

#endregion ----------------------------------------------------------------------


#region --- FUNCTIONS -----------------------------------------------------------

# -----------------------------------------------------------------------------
# Write-Log
# Writes a timestamped, colour-coded message to console and log file.
# -----------------------------------------------------------------------------
function Write-Log {
    param (
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR","SECTION","BLANK")]
        [string]$Level = "INFO"
    )
    $ColourMap = @{
        INFO    = "Gray"
        SUCCESS = "Green"
        WARN    = "Yellow"
        ERROR   = "Red"
        SECTION = "Cyan"
        BLANK   = "Gray"
    }
    $PrefixMap = @{
        INFO    = "[INFO]   "
        SUCCESS = "[OK]     "
        WARN    = "[WARN]   "
        ERROR   = "[ERROR]  "
        SECTION = "[=======]"
        BLANK   = "         "
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


# -----------------------------------------------------------------------------
# Get-GraphToken
# Authenticates with Entra ID using client credentials and returns an
# access token for Microsoft Graph API.
# -----------------------------------------------------------------------------
function Get-GraphToken {
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
        Write-Log "Requesting access token from Entra ID..." -Level INFO
        $Response = Invoke-RestMethod `
            -Method      POST `
            -Uri         "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -ContentType "application/x-www-form-urlencoded" `
            -Body        $Body `
            -ErrorAction Stop
        Write-Log "Access token acquired successfully." -Level SUCCESS
        return $Response.access_token
    }
    catch {
        Write-Log "Authentication failed: $_" -Level ERROR
        exit 1
    }
}


# -----------------------------------------------------------------------------
# Get-UserByUPN
# Single-pass exact UPN lookup via Microsoft Graph.
#
# Uses:
#   $filter=userPrincipalName eq 'upn'
#   ConsistencyLevel: eventual + $count=true for reliable filter support
#   PowerShell -ieq post-validation as secondary exact-match gate
#
# Returns the matched user object, or $null if not found.
# Handles 429 throttling with one automatic retry.
# -----------------------------------------------------------------------------
function Get-UserByUPN {
    param (
        [string]$UPN,
        [string]$AccessToken
    )

    $Select  = "id,displayName,userPrincipalName,mail,department,jobTitle,officeLocation,accountEnabled,companyName,usageLocation,country,state,city,streetAddress"
    $Encoded = [Uri]::EscapeDataString("userPrincipalName eq '$UPN'")
    $Uri     = "https://graph.microsoft.com/v1.0/users?`$filter=$Encoded&`$select=$Select&`$count=true"

    $Headers = @{
        Authorization      = "Bearer $AccessToken"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual"
    }

    try {
        Write-Log "    Querying Graph API..." -Level INFO
        $Response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop

        if (-not $Response.value -or $Response.value.Count -eq 0) {
            return $null
        }

        # Post-validate: -ieq guarantees character-exact, case-insensitive match
        $Match = $Response.value | Where-Object { $_.userPrincipalName -ieq $UPN } | Select-Object -First 1

        return $Match
    }
    catch {
        $StatusCode = $_.Exception.Response.StatusCode.value__

        if ($StatusCode -eq 404) { return $null }

        if ($StatusCode -eq 429) {
            $WaitSeconds = 30
            try { $WaitSeconds = [int]$_.Exception.Response.Headers["Retry-After"] } catch { }
            Write-Log "    429 throttled - waiting ${WaitSeconds}s before retry..." -Level WARN
            Start-Sleep -Seconds $WaitSeconds
            try {
                $Response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
                if (-not $Response.value -or $Response.value.Count -eq 0) { return $null }
                return $Response.value | Where-Object { $_.userPrincipalName -ieq $UPN } | Select-Object -First 1
            }
            catch {
                Write-Log "    Retry after throttle also failed: $_" -Level WARN
                return $null
            }
        }

        Write-Log "    Graph API error (HTTP $StatusCode): $_" -Level WARN
        throw
    }
}


# -----------------------------------------------------------------------------
# Build-ResultRow
# Builds a single structured CSV output row from a user object.
# All fields default to "N/A" if the user object is null or property missing.
# -----------------------------------------------------------------------------
function Build-ResultRow {
    param (
        [string]$InputUPN,
        [string]$LookupStatus,
        [PSObject]$User
    )

    return [PSCustomObject]@{
        InputUPN          = $InputUPN
        LookupStatus      = $LookupStatus
        DisplayName       = if ($User -and $User.displayName)              { [string]$User.displayName }              else { "N/A" }
        UserPrincipalName = if ($User -and $User.userPrincipalName)        { [string]$User.userPrincipalName }        else { "N/A" }
        Mail              = if ($User -and $User.mail)                     { [string]$User.mail }                     else { "N/A" }
        Department        = if ($User -and $User.department)               { [string]$User.department }               else { "N/A" }
        JobTitle          = if ($User -and $User.jobTitle)                 { [string]$User.jobTitle }                 else { "N/A" }
        OfficeLocation    = if ($User -and $User.officeLocation)           { [string]$User.officeLocation }           else { "N/A" }
        AccountEnabled    = if ($User -and $null -ne $User.accountEnabled) { [string]$User.accountEnabled }           else { "N/A" }
        CompanyName       = if ($User -and $User.companyName)              { [string]$User.companyName }              else { "N/A" }
        UsageLocation     = if ($User -and $User.usageLocation)            { [string]$User.usageLocation }            else { "N/A" }
        Country           = if ($User -and $User.country)                  { [string]$User.country }                  else { "N/A" }
        State             = if ($User -and $User.state)                    { [string]$User.state }                    else { "N/A" }
        City              = if ($User -and $User.city)                     { [string]$User.city }                     else { "N/A" }
        StreetAddress     = if ($User -and $User.streetAddress)            { [string]$User.streetAddress }            else { "N/A" }
        UserId            = if ($User -and $User.id)                       { [string]$User.id }                       else { "N/A" }
    }
}

#endregion ----------------------------------------------------------------------


#region --- MAIN ----------------------------------------------------------------

# Initialise log file
try {
    [System.IO.File]::WriteAllText(
        $script:LogFile,
        "Get-UserDetailsByUPN v1.0`r`nStarted : $(Get-Date)`r`nInput   : $InputPath`r`n`r`n",
        [System.Text.Encoding]::UTF8
    )
} catch {
    $script:LogFile = $null
}

# Banner
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Get User Details by UPN  |  Sethu Kumar B  |  v1.0           " -ForegroundColor Cyan
Write-Host "  READ ONLY - No changes made to any system                     " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "" -Level BLANK
Write-Log "Script root  : $PSScriptRoot"       -Level INFO
Write-Log "Input file   : $InputPath"          -Level INFO
Write-Log "Output CSV   : $OutputFile"         -Level INFO
Write-Log "Log file     : $($script:LogFile)"  -Level INFO
Write-Log "" -Level BLANK


# ==============================================================================
# STEP 1 - Read and validate input file
# ==============================================================================
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 1 - Reading Input File" -Level INFO
Write-Log "==========================================================" -Level SECTION

if (-not (Test-Path $InputPath)) {
    Write-Log "Input file not found: $InputPath" -Level ERROR
    Write-Log "Create UPNs.txt with one UPN per line in the same folder as this script." -Level ERROR
    exit 1
}

$RawLines = Get-Content -Path $InputPath -Encoding UTF8 -ErrorAction Stop
$UPNList  = [System.Collections.Generic.List[string]]::new()

foreach ($Line in $RawLines) {
    $Trimmed = $Line.Trim()
    if ([string]::IsNullOrWhiteSpace($Trimmed)) { continue }
    if ($Trimmed.StartsWith("#"))               { continue }
    # Basic UPN format check - must contain @
    if ($Trimmed -notmatch '@') {
        Write-Log "  SKIPPED (not a valid UPN format - missing @): $Trimmed" -Level WARN
        continue
    }
    $UPNList.Add($Trimmed)
}

Write-Log "File loaded - $($RawLines.Count) raw line(s), $($UPNList.Count) valid UPN(s)." -Level SUCCESS
Write-Log "" -Level BLANK

if ($UPNList.Count -eq 0) {
    Write-Log "No valid UPNs found in input file. Nothing to process. Exiting." -Level WARN
    exit 0
}

Write-Log "UPNs to resolve:" -Level INFO
foreach ($u in $UPNList) { Write-Log "  -> $u" -Level INFO }
Write-Log "" -Level BLANK


# ==============================================================================
# STEP 2 - Authenticate with Entra ID
# ==============================================================================
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 2 - Authenticating with Entra ID" -Level INFO
Write-Log "==========================================================" -Level SECTION

$Token = Get-GraphToken -TenantId $TenantID -ClientId $ClientID -ClientSecret $ClientSecret
Write-Log "" -Level BLANK


# ==============================================================================
# STEP 3 - Fetch details for each UPN
# ==============================================================================
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 3 - Fetching User Details" -Level INFO
Write-Log "==========================================================" -Level SECTION
Write-Log "" -Level BLANK

$Results       = [System.Collections.Generic.List[PSObject]]::new()
$Counter       = 0
$Total         = $UPNList.Count
$FoundCount    = 0
$NotFoundCount = 0
$ErrorCount    = 0

foreach ($UPN in $UPNList) {
    $Counter++
    Write-Log "  [$Counter / $Total] Looking up: '$UPN'" -Level INFO

    try {
        $User = Get-UserByUPN -UPN $UPN -AccessToken $Token

        if ($null -eq $User) {
            Write-Log "    RESULT: NOT FOUND" -Level WARN
            $Results.Add((Build-ResultRow -InputUPN $UPN -LookupStatus "NOT FOUND" -User $null))
            $NotFoundCount++
        }
        else {
            $Email = if ($User.mail) { $User.mail } else { $User.userPrincipalName }
            Write-Log "    RESULT: FOUND  ->  $($User.displayName)  |  $Email  |  $($User.department)" -Level SUCCESS
            $Results.Add((Build-ResultRow -InputUPN $UPN -LookupStatus "FOUND" -User $User))
            $FoundCount++
        }
    }
    catch {
        Write-Log "    ERROR fetching '$UPN': $_" -Level ERROR
        $Results.Add((Build-ResultRow -InputUPN $UPN -LookupStatus "ERROR" -User $null))
        $ErrorCount++
    }

    Write-Log "" -Level BLANK
    Start-Sleep -Milliseconds 100
}


# ==============================================================================
# STEP 4 - Export results to CSV
# ==============================================================================
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 4 - Exporting CSV" -Level INFO
Write-Log "==========================================================" -Level SECTION

if ($Results.Count -gt 0) {
    try {
        $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        $SizeMB = ((Get-Item $OutputFile).Length / 1MB).ToString("0.00")
        Write-Log "CSV exported successfully." -Level SUCCESS
        Write-Log "  Path : $OutputFile"                             -Level INFO
        Write-Log "  Rows : $($Results.Count)  |  Size: $SizeMB MB" -Level INFO
    }
    catch {
        Write-Log "CSV export failed: $_" -Level ERROR
        exit 1
    }
} else {
    Write-Log "No results to export." -Level WARN
}


# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  COMPLETE - READ ONLY - NO CHANGES MADE                       " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "  Total UPNs           : $Total"         -Level INFO
Write-Log "  Found                : $FoundCount"    -Level $(if ($FoundCount    -gt 0) { "SUCCESS" } else { "INFO" })
Write-Log "  Not found            : $NotFoundCount" -Level $(if ($NotFoundCount -gt 0) { "WARN"    } else { "INFO" })
Write-Log "  Errors               : $ErrorCount"    -Level $(if ($ErrorCount    -gt 0) { "ERROR"   } else { "INFO" })
Write-Log "" -Level BLANK
Write-Log "  Output CSV  : $OutputFile"         -Level INFO
Write-Log "  Log file    : $($script:LogFile)"  -Level INFO
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

#endregion ----------------------------------------------------------------------
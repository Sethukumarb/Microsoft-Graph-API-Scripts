#Requires -Version 5.1
# ==============================================================================
# Script Name  : Get-UserEmailFromUsername.ps1
# Description  : Resolves email addresses (UPN and mail) for one or more users
#                by looking up their username (sAMAccountName / display name /
#                full UPN) in Azure AD / Entra ID via Microsoft Graph API.
#
#                LOOKUP STRATEGY:
#                  For each username in the input file the script tries:
#                    Pass 1 - Exact UPN match
#                             $filter=userPrincipalName eq 'username@domain'
#                             Post-validated with -ieq (case-insensitive exact)
#                    Pass 2 - Exact displayName match
#                             $filter=displayName eq 'username'
#                             Post-validated with -ieq (case-insensitive exact)
#
#                  Stops at the first pass that returns a validated result.
#                  Both passes use ConsistencyLevel: eventual + $count=true
#                  so Graph filters work reliably for all account types.
#
#                  Case behaviour:
#                    "sethu kumar b" matches "Sethu Kumar B"  -> FOUND
#                    "Sethu Kumar"   does NOT match "Sethu Kumar B" -> NOT FOUND
#                    Characters must be identical; only case is flexible.
#
#                INPUT FILE:
#                  Usernames.txt - same folder as this script
#                  One username per line. Blank lines and # comments ignored.
#
#                  Accepted input formats (any mix):
#                    john.doe              <- sAMAccountName / UPN prefix
#                    john.doe@contoso.com  <- full UPN
#                    John Doe              <- display name (exact, any case)
#
#                OUTPUT FILES (saved to $PSScriptRoot):
#                  UserEmailDetails_[timestamp].csv  - full results
#                  UserEmailDetails_[timestamp].log  - run log
#
#                CSV COLUMNS:
#                  InputUsername, MatchCount, LookupStatus,
#                  DisplayName, UserPrincipalName, Mail,
#                  Department, JobTitle, OfficeLocation,
#                  AccountEnabled, UserId,
#                  CompanyName, UsageLocation,
#                  Country, State, City, StreetAddress
#
#                LOOKUP STATUS VALUES:
#                  FOUND          - exactly one match, email resolved
#                  MULTIPLE MATCH - more than one user matched the input
#                  NOT FOUND      - no user found for this input
#                  ERROR          - API error during lookup
#
#                READ ONLY: Only GET requests. No changes made anywhere.
#
# Author       : Sethu Kumar B
# Version      : 2.1
# Created Date : 2026-04-15
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
#   v1.0 - 2026-04-15 - Sethu Kumar B - Initial release. Four-pass lookup
#                        strategy (UPN exact, displayName exact, displayName
#                        startsWith, UPN startsWith). All matches returned
#                        when multiple users found for same input.
#
#   v2.0 - 2026-04-15 - Sethu Kumar B - Reduced to two-pass exact-only lookup.
#                        Added ConsistencyLevel: eventual + $count=true headers
#                        so displayName eq filter works reliably in Entra ID.
#                        Added PowerShell -ieq post-validation for guaranteed
#                        case-insensitive exact character matching. Removed
#                        startsWith passes to prevent partial matches.
#
#   v2.1 - 2026-04-28 - Sethu Kumar B - Added new output fields: CompanyName,
#                        UsageLocation, Country, State, City, StreetAddress.
#                        Updated $Select, Build-ResultRow, and CSV column docs.
# ==============================================================================


#region --- CONFIGURATION -------------------------------------------------------

$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

# Input file - must be in the same folder as this script
$InputFileName = "Usernames.txt"
$InputPath     = Join-Path $PSScriptRoot $InputFileName

#endregion ----------------------------------------------------------------------


#region --- INIT ----------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile     = Join-Path $PSScriptRoot "UserEmailDetails_$Timestamp.csv"
$script:LogFile = Join-Path $PSScriptRoot "UserEmailDetails_$Timestamp.log"

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
# Invoke-GraphSearch
# Sends a single GET request to Microsoft Graph and returns the value array.
#
# Key headers sent:
#   ConsistencyLevel: eventual  - required for $filter on displayName and other
#                                 non-default-indexed properties in Entra ID.
#                                 Without this, Graph silently returns empty
#                                 results even when the user exists.
#   $count=true (in URI)        - required alongside ConsistencyLevel: eventual
#                                 for advanced query support.
#
# Returns an empty array on 404 or any unrecoverable error.
# Handles 429 throttling with one automatic retry.
# -----------------------------------------------------------------------------
function Invoke-GraphSearch {
    param (
        [string]$Uri,
        [string]$AccessToken
    )

    $Headers = @{
        Authorization      = "Bearer $AccessToken"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual"   # Required for advanced $filter queries
    }

    try {
        $Response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
        if ($Response.value) { return @($Response.value) }
        return @()
    }
    catch {
        $StatusCode = $_.Exception.Response.StatusCode.value__

        if ($StatusCode -eq 404) {
            return @()
        }

        if ($StatusCode -eq 429) {
            # Throttled - honour Retry-After header, then retry once
            $WaitSeconds = 30
            try { $WaitSeconds = [int]$_.Exception.Response.Headers["Retry-After"] } catch { }
            Write-Log "  429 throttled - waiting ${WaitSeconds}s before retry..." -Level WARN
            Start-Sleep -Seconds $WaitSeconds
            try {
                $Response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
                if ($Response.value) { return @($Response.value) }
                return @()
            }
            catch {
                Write-Log "  Retry after throttle also failed: $_" -Level WARN
                return @()
            }
        }

        Write-Log "  Graph API error (HTTP $StatusCode): $_" -Level WARN
        return @()
    }
}


# -----------------------------------------------------------------------------
# Resolve-UserByUsername
# Two-pass exact lookup strategy for a given username string.
#
# Pass 1: Exact UPN match        - fastest, for full UPN or UPN-prefix input
# Pass 2: Exact displayName match - for display name input (e.g. "John Doe")
#
# Both passes:
#   - Use Graph $filter with eq operator (case-insensitive on Graph side)
#   - Use $count=true + ConsistencyLevel: eventual for reliable filter support
#   - Apply PowerShell -ieq post-validation as a secondary exact-match gate
#     to ensure no unexpected partial or fuzzy results slip through
#
# Case behaviour:
#   Input "sethu kumar b"  matches Display Name "Sethu Kumar B"  -> FOUND
#   Input "Sethu Kumar"    does NOT match "Sethu Kumar B"        -> NOT FOUND
#
# Returns array of matching user objects. Empty array if nothing found.
# -----------------------------------------------------------------------------
function Resolve-UserByUsername {
    param (
        [string]$Username,
        [string]$AccessToken
    )

    # Fields to retrieve for each matched user
    $Select = "id,displayName,userPrincipalName,mail,department,jobTitle,officeLocation,accountEnabled,companyName,usageLocation,country,state,city,streetAddress"

    # ------------------------------------------------------------------
    # Pass 1 - Exact UPN match
    # Handles: john.doe@contoso.com or john.doe (Graph eq on UPN)
    # ------------------------------------------------------------------
    Write-Log "    Pass 1: Exact UPN match (case-insensitive)..." -Level INFO

    $Encoded  = [Uri]::EscapeDataString("userPrincipalName eq '$Username'")
    $Uri      = "https://graph.microsoft.com/v1.0/users?`$filter=$Encoded&`$select=$Select&`$count=true"
    $GraphHits = Invoke-GraphSearch -Uri $Uri -AccessToken $AccessToken

    # Post-validate: PowerShell -ieq ensures character-exact, case-insensitive match
    $Validated = @($GraphHits | Where-Object { $_.userPrincipalName -ieq $Username })

    if ($Validated.Count -gt 0) {
        Write-Log "    Pass 1 matched: $($Validated.Count) result(s) after exact validation." -Level SUCCESS
        return $Validated
    }

    Write-Log "    Pass 1: No exact UPN match." -Level INFO

    # ------------------------------------------------------------------
    # Pass 2 - Exact displayName match
    # Handles: "Sethu Kumar B" or "John Doe" (Graph eq on displayName)
    # Requires ConsistencyLevel: eventual to work reliably in Entra ID.
    # ------------------------------------------------------------------
    Write-Log "    Pass 2: Exact displayName match (case-insensitive)..." -Level INFO

    $Encoded  = [Uri]::EscapeDataString("displayName eq '$Username'")
    $Uri      = "https://graph.microsoft.com/v1.0/users?`$filter=$Encoded&`$select=$Select&`$count=true"
    $GraphHits = Invoke-GraphSearch -Uri $Uri -AccessToken $AccessToken

    # Post-validate: PowerShell -ieq ensures character-exact, case-insensitive match
    $Validated = @($GraphHits | Where-Object { $_.displayName -ieq $Username })

    if ($Validated.Count -gt 0) {
        Write-Log "    Pass 2 matched: $($Validated.Count) result(s) after exact validation." -Level SUCCESS
        return $Validated
    }

    Write-Log "    Pass 2: No exact displayName match." -Level INFO
    Write-Log "    No match found on any pass." -Level WARN
    return @()
}


# -----------------------------------------------------------------------------
# Build-ResultRow
# Builds a single structured CSV output row from a user object.
# All fields default to "N/A" if the user object is null or property missing.
# -----------------------------------------------------------------------------
function Build-ResultRow {
    param (
        [string]$InputUsername,
        [int]$MatchCount,
        [string]$LookupStatus,
        [PSObject]$User
    )

    return [PSCustomObject]@{
        InputUsername     = $InputUsername
        MatchCount        = $MatchCount
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
        "Get-UserEmailFromUsername v2.1`r`nStarted : $(Get-Date)`r`nInput   : $InputPath`r`n`r`n",
        [System.Text.Encoding]::UTF8
    )
} catch {
    $script:LogFile = $null
}

# Banner
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Get User Email from Username  |  Sethu Kumar B  |  v2.1      " -ForegroundColor Cyan
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
    Write-Log "Create Usernames.txt with one username per line in the same folder as this script." -Level ERROR
    exit 1
}

$RawLines  = Get-Content -Path $InputPath -Encoding UTF8 -ErrorAction Stop
$Usernames = [System.Collections.Generic.List[string]]::new()

foreach ($Line in $RawLines) {
    $Trimmed = $Line.Trim()
    if ([string]::IsNullOrWhiteSpace($Trimmed)) { continue }  # skip blank lines
    if ($Trimmed.StartsWith("#"))               { continue }  # skip comments
    $Usernames.Add($Trimmed)
}

Write-Log "File loaded - $($RawLines.Count) raw line(s), $($Usernames.Count) valid username(s)." -Level SUCCESS
Write-Log "" -Level BLANK

if ($Usernames.Count -eq 0) {
    Write-Log "No usernames found in input file. Nothing to process. Exiting." -Level WARN
    exit 0
}

Write-Log "Usernames to resolve:" -Level INFO
foreach ($u in $Usernames) { Write-Log "  -> $u" -Level INFO }
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
# STEP 3 - Resolve each username via Microsoft Graph
# ==============================================================================
Write-Log "==========================================================" -Level SECTION
Write-Log "STEP 3 - Resolving Usernames" -Level INFO
Write-Log "==========================================================" -Level SECTION
Write-Log "Strategy : Exact match only (UPN then DisplayName)" -Level INFO
Write-Log "Case     : Case-insensitive, character-exact" -Level INFO
Write-Log "" -Level BLANK

$Results       = [System.Collections.Generic.List[PSObject]]::new()
$Counter       = 0
$Total         = $Usernames.Count
$FoundCount    = 0
$MultipleCount = 0
$NotFoundCount = 0
$ErrorCount    = 0

foreach ($Username in $Usernames) {
    $Counter++
    Write-Log "  [$Counter / $Total] Resolving: '$Username'" -Level INFO

    try {
        $Matches = Resolve-UserByUsername -Username $Username -AccessToken $Token

        if ($Matches.Count -eq 0) {
            # --------------------------------------------------------
            # NOT FOUND - no user matched on either pass
            # --------------------------------------------------------
            Write-Log "    RESULT: NOT FOUND" -Level WARN
            $Results.Add(
                (Build-ResultRow -InputUsername $Username -MatchCount 0 `
                 -LookupStatus "NOT FOUND" -User $null)
            )
            $NotFoundCount++

        } elseif ($Matches.Count -eq 1) {
            # --------------------------------------------------------
            # FOUND - exactly one user matched (ideal case)
            # --------------------------------------------------------
            $User  = $Matches[0]
            $Email = if ($User.mail) { $User.mail } else { $User.userPrincipalName }
            Write-Log "    RESULT: FOUND  ->  $($User.displayName)  |  $Email" -Level SUCCESS
            $Results.Add(
                (Build-ResultRow -InputUsername $Username -MatchCount 1 `
                 -LookupStatus "FOUND" -User $User)
            )
            $FoundCount++

        } else {
            # --------------------------------------------------------
            # MULTIPLE MATCH - more than one user shares this exact
            # display name (unlikely but possible in large tenants).
            # All matches are returned so you can identify the right one.
            # --------------------------------------------------------
            Write-Log "    RESULT: MULTIPLE MATCH - $($Matches.Count) users share this exact name." -Level WARN
            foreach ($User in $Matches) {
                $Email = if ($User.mail) { $User.mail } else { $User.userPrincipalName }
                Write-Log "      -> $($User.displayName)  |  $Email  |  Dept: $($User.department)" -Level WARN
                $Results.Add(
                    (Build-ResultRow -InputUsername $Username -MatchCount $Matches.Count `
                     -LookupStatus "MULTIPLE MATCH" -User $User)
                )
            }
            $MultipleCount++
        }
    }
    catch {
        # --------------------------------------------------------
        # ERROR - unexpected exception during resolution
        # --------------------------------------------------------
        Write-Log "    ERROR resolving '$Username': $_" -Level ERROR
        $Results.Add(
            (Build-ResultRow -InputUsername $Username -MatchCount 0 `
             -LookupStatus "ERROR" -User $null)
        )
        $ErrorCount++
    }

    Write-Log "" -Level BLANK

    # Small delay to be polite to the Graph API
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
Write-Log "  Total usernames      : $Total"         -Level INFO
Write-Log "  Found (exact match)  : $FoundCount"    -Level $(if ($FoundCount    -gt 0) { "SUCCESS" } else { "INFO" })
Write-Log "  Multiple matches     : $MultipleCount" -Level $(if ($MultipleCount -gt 0) { "WARN"    } else { "INFO" })
Write-Log "  Not found            : $NotFoundCount" -Level $(if ($NotFoundCount -gt 0) { "WARN"    } else { "INFO" })
Write-Log "  Errors               : $ErrorCount"    -Level $(if ($ErrorCount    -gt 0) { "ERROR"   } else { "INFO" })
Write-Log "" -Level BLANK
Write-Log "  Output CSV  : $OutputFile"         -Level INFO
Write-Log "  Log file    : $($script:LogFile)"  -Level INFO
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

#endregion ----------------------------------------------------------------------
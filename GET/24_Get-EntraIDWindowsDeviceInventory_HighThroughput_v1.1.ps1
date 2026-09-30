#Requires -Version 5.1
<#
.SYNOPSIS
    Fetches all Windows device objects from Entra ID (Azure AD), enriches each record
    with serial number (from Intune or Autopilot) and Intune Device ID where available,
    and exports the full inventory into an 11-sheet Excel workbook categorized by
    days since last sign-in.

.DESCRIPTION
    This script connects to Microsoft Graph API using App Registration credentials
    (Client ID, Client Secret, Tenant ID) and performs the following:

    STEP 1 — Fetch all Entra ID device objects where operatingSystem eq 'Windows'.
             Only pure Windows devices are fetched. Windows Server, iOS, Android,
             macOS are excluded.

    STEP 2 — Fetch all Windows Intune managed device objects.
             Used to cross-reference serial number and Intune Device ID.
             Join key: Entra ID deviceId matches Intune azureADDeviceId.

    STEP 3 — Fetch all Windows Autopilot device objects.
             Used as a fallback serial number source if Intune has no match.
             Join key: Entra ID deviceId matches Autopilot azureADDeviceId.

    STEP 4 — For each Entra ID device:
             - Match to Intune object  → pull IntuneDeviceId, SerialNumber ONLY
               NOTE: Intune PrimaryUser is NOT mapped — Entra ID objects may be old
               duplicates that do not correspond to the current active Intune user.
               User context from Intune would be misleading on stale Entra objects.
             - If no Intune match      → try Autopilot object for SerialNumber + AutopilotObjectId
             - Owner name + UPN pulled from /devices/{id}/registeredOwners using Graph
               $batch (20 per call). Toggle with $IncludeOwner in the config block.
             - Calculate DaysSinceSignIn from approximateLastSignInDateTime

    STEP 5 — Categorize all devices into 11 aging buckets based on DaysSinceSignIn
             and export to a single Excel workbook with 11 worksheet tabs.

    Sheet layout:
        Sheet 01 -   0 to   7 days : Active    — recently signed in, healthy.
        Sheet 02 -   8 to  30 days : Normal    — signed in within a month.
        Sheet 03 -  31 to  60 days : Watch     — slightly overdue, monitor.
        Sheet 04 -  61 to  90 days : Aging     — possible connectivity issue.
        Sheet 05 -  91 to 120 days : Concern   — approaching warning threshold.
        Sheet 06 - 121 to 150 days : Warning   — action recommended.
        Sheet 07 - 151 to 180 days : Elevated  — escalate for investigation.
        Sheet 08 - 181 to 240 days : High Risk — likely offline or abandoned.
        Sheet 09 - 241 to 365 days : Critical  — cleanup or recovery required.
        Sheet 10 - 366 to 700 days : Stale     — remove from Entra ID immediately.
        Sheet 11 - 701+ days       : Defunct   — long-abandoned, priority cleanup.

    Output columns per device:
        DeviceName, EntraDeviceId, EntraObjectId, OperatingSystem, OSVersion,
        AccountEnabled, TrustType, JoinType, DeviceOwnerName, DeviceOwnerUPN,
        MDMProvider, SecuritySettingsManagement, IsManaged, IsCompliant,
        RegisteredDateTime, ApproximateLastSignInDateTime, DaysSinceSignIn,
        SerialNumber, IntuneDeviceId, AutopilotObjectId, IntuneRecord, ManagementAgent

    Output file saved to the same folder as this script.
    Full transcript log saved alongside the Excel file.

.OUTPUTS
    - EntraIDWindowsDeviceInventory_<timestamp>.xlsx   (11-sheet Excel workbook)
    - EntraIDWindowsDevices_Transcript_<timestamp>.log (PowerShell transcript)

.NOTES
    Script Name   : Get-EntraIDWindowsDeviceInventory.ps1
    Author        : Sethu Kumar B
    Version       : 1.1
    Created       : 2026-05-13
    Last Modified : 2026-05-13 (v1.1 — added portal-style columns: owner name/UPN via $batch, compliance, registered date, security settings mgmt; token refresh)

    Required PowerShell Module : ImportExcel
    Install if missing         : Install-Module ImportExcel -Scope CurrentUser

    Required Graph API App Permissions (Application type):
        - Device.Read.All                          (Entra ID devices + registered owners)
        - DeviceManagementManagedDevices.Read.All  (Intune devices)
        - DeviceManagementServiceConfig.Read.All   (Autopilot devices)

.EXAMPLE
    .\Get-EntraIDWindowsDeviceInventory.ps1
    Runs the script. Output files are saved to the same folder as the script.
#>

# ============================================================
# CONFIGURATION BLOCK — Update these values before running
# ============================================================
$TenantID     = ""
$ClientID     = ""
$ClientSecret = ""

# $true  = fetch registered owner name + UPN for every device (uses $batch, 20 devices per call).
#           Adds significant runtime on large tenants. $false = skip, owner columns left blank.
$IncludeOwner = $true

# Parallel threads for the owner lookup. 6 is a reasonable start. If you see many throttle
# pauses in the log, lower it (3-4). If none, you can try 8-10.
$OwnerThreads = 6
# ============================================================

# -------------------------------------------------------
# OUTPUT PATH SETUP
# -------------------------------------------------------
$ScriptRoot = $PSScriptRoot
$Timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ExcelFile  = Join-Path $ScriptRoot "EntraIDWindowsDeviceInventory_$Timestamp.xlsx"
$LogFile    = Join-Path $ScriptRoot "EntraIDWindowsDevices_Transcript_$Timestamp.log"

# -------------------------------------------------------
# START TRANSCRIPT
# -------------------------------------------------------
Start-Transcript -Path $LogFile -Append
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "  Entra ID Windows Device Inventory Report" -ForegroundColor Cyan
Write-Host "  Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host ""

# -------------------------------------------------------
# STEP 1 — CHECK IMPORTEXCEL MODULE
# -------------------------------------------------------
Write-Host "[STEP 1/7] Checking for ImportExcel module..." -ForegroundColor Yellow

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "  [!] ImportExcel module not found. Attempting to install..." -ForegroundColor Red
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "  [OK] ImportExcel installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "  [FATAL] Failed to install ImportExcel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Run manually: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Red
        Stop-Transcript
        exit 1
    }
}
else {
    Write-Host "  [OK] ImportExcel module found." -ForegroundColor Green
}

Import-Module ImportExcel -ErrorAction Stop

# -------------------------------------------------------
# STEP 2 — AUTHENTICATE TO GRAPH API
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 2/7] Authenticating to Microsoft Graph API..." -ForegroundColor Yellow

$TokenUrl  = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$TokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $TokenResponse = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $TokenBody -ErrorAction Stop
    $AccessToken   = $TokenResponse.access_token
    $AuthHeaders   = @{ Authorization = "Bearer $AccessToken" }
    Write-Host "  [OK] Authentication successful." -ForegroundColor Green
}
catch {
    Write-Host "  [FATAL] Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript
    exit 1
}

# -------------------------------------------------------
# HELPER — Generic paginated Graph API fetch function
# Accepts a URI, returns a Generic List of all result objects
# -------------------------------------------------------
function Invoke-GraphGetAll {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Label
    )

    $Results   = [System.Collections.Generic.List[object]]::new()
    $PageCount = 0

    do {
        $PageCount++
        try {
            $Response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -ErrorAction Stop
        }
        catch {
            Write-Host "  [ERROR] Failed fetching $Label page ${PageCount}: $($_.Exception.Message)" -ForegroundColor Red
            break
        }

        if ($Response.value) {
            foreach ($Item in $Response.value) { $Results.Add($Item) }
        }

        # FIX: last page has no @odata.nextLink. Under Set-StrictMode, reading a
        # missing property throws PropertyNotFoundStrict. Check it exists first.
        $Uri = $null
        if ($Response.PSObject.Properties['@odata.nextLink']) {
            $Uri = $Response.'@odata.nextLink'
        }

        Write-Host ("  Page {0,3} fetched — Running total: {1} {2}" -f $PageCount, $Results.Count, $Label) -ForegroundColor DarkGray

    } while ($Uri)

    return $Results
}

# -------------------------------------------------------
# HELPER — Strict-mode-safe property read.
# Returns $null instead of throwing when a property is missing.
# -------------------------------------------------------
function Get-SafeProp {
    param($Object, [string]$Name)
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $null
}

# -------------------------------------------------------
# HELPER — Refresh the access token in place when it gets old.
# $Headers is a hashtable (reference type), so updating it here
# updates it for every caller. Tokens last ~60-90 min.
# -------------------------------------------------------
$TokenIssuedAt = Get-Date
function Update-GraphToken {
    param([hashtable]$Headers)
    if (((Get-Date) - $script:TokenIssuedAt).TotalMinutes -lt 40) { return }
    $R = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $TokenBody -ErrorAction Stop
    $Headers['Authorization'] = "Bearer $($R.access_token)"
    $script:TokenIssuedAt = Get-Date
    Write-Host "  [INFO] Access token refreshed." -ForegroundColor DarkGray
}

# -------------------------------------------------------
# HELPER — Registered owners for all devices via Graph $batch.
# 20 requests per batch call, several batches in parallel (runspace pool).
# Retries 429/503/504 using Retry-After; all threads pause together on throttle.
# Returns hashtable: Entra object id -> @{ Name = "..."; UPN = "..." }
# Multiple owners are joined with "; ".
# -------------------------------------------------------
function Get-DeviceOwnerMap {
    param([object[]]$Devices, [hashtable]$Headers, [int]$Threads = 6)

    $Arr   = @($Devices)
    $Total = $Arr.Count

    # Shared, thread-safe state
    $Shared  = [hashtable]::Synchronized(@{ Token = $Headers['Authorization']; PauseUntil = [datetime]::MinValue })
    $Results = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new()
    $Queue   = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

    # Work items = comma-separated string of up to 20 device object ids
    for ($i = 0; $i -lt $Total; $i += 20) {
        $Last = [Math]::Min($i + 19, $Total - 1)
        $Ids  = @($Arr[$i..$Last] | ForEach-Object { [string]$_.id })
        $Queue.Enqueue(($Ids -join ','))
    }

    # Worker: runs in its own runspace, pulls chunks until the queue is empty
    $Worker = {
        param($Queue, $Results, $Shared)

        $BatchUri = 'https://graph.microsoft.com/v1.0/$batch'
        function SafeProp($Obj, $Name) {
            if ($null -ne $Obj -and $Obj.PSObject.Properties[$Name]) { return $Obj.$Name }
            return $null
        }

        $Chunk = $null
        while ($Queue.TryDequeue([ref]$Chunk)) {
            $Pending = @($Chunk -split ',')
            $Attempt = 0

            while ($Pending.Count -gt 0 -and $Attempt -lt 8) {
                $Attempt++

                # Honor a shared pause set by any thread that got throttled
                while ((Get-Date) -lt $Shared.PauseUntil) { Start-Sleep -Milliseconds 500 }

                $Hdrs     = @{ Authorization = $Shared.Token }
                $Requests = @($Pending | ForEach-Object {
                    @{ id = $_; method = 'GET'; url = "/devices/$_/registeredOwners/microsoft.graph.user?`$select=displayName,userPrincipalName" }
                })
                $Body = @{ requests = $Requests } | ConvertTo-Json -Depth 5 -Compress

                try {
                    $Resp = Invoke-RestMethod -Uri $BatchUri -Headers $Hdrs -Method Post -Body $Body -ContentType 'application/json' -ErrorAction Stop
                }
                catch {
                    $Wait = 5 * $Attempt
                    try {
                        $RA = $_.Exception.Response.Headers['Retry-After']
                        if ($RA) { $Wait = [Math]::Max($Wait, [int]$RA) }
                    } catch { }
                    $Shared.PauseUntil = (Get-Date).AddSeconds([Math]::Min($Wait, 120))
                    continue
                }

                $Retry = @()
                foreach ($R in @($Resp.responses)) {
                    $Id     = [string]$R.id
                    $Status = [int]$R.status

                    if ($Status -eq 200) {
                        $Names = @()
                        $Upns  = @()
                        foreach ($U in @((SafeProp (SafeProp $R 'body') 'value'))) {
                            $N = [string](SafeProp $U 'displayName')
                            $P = [string](SafeProp $U 'userPrincipalName')
                            if ($N) { $Names += $N }
                            if ($P) { $Upns  += $P }
                        }
                        $Results[$Id] = @{ Name = ($Names -join '; '); UPN = ($Upns -join '; '); Failed = $false }
                    }
                    elseif ($Status -in 429, 503, 504) {
                        $Retry += $Id
                        $RA = SafeProp (SafeProp $R 'headers') 'Retry-After'
                        $W  = 5 * $Attempt
                        if ($RA) { $W = [Math]::Max($W, [int]$RA) }
                        $Until = (Get-Date).AddSeconds([Math]::Min($W, 120))
                        if ($Until -gt $Shared.PauseUntil) { $Shared.PauseUntil = $Until }
                    }
                    else {
                        # 404 = no owner / object gone (fine). Anything else = failure.
                        $Results[$Id] = @{ Name = ''; UPN = ''; Failed = ($Status -ne 404) }
                    }
                }
                $Pending = @($Retry)
            }

            # Gave up on these after all retries
            foreach ($Id in $Pending) { $Results[$Id] = @{ Name = ''; UPN = ''; Failed = $true } }
        }
    }

    # Start the pool
    $Pool = [runspacefactory]::CreateRunspacePool(1, $Threads)
    $Pool.Open()
    $Jobs = @()
    for ($t = 0; $t -lt $Threads; $t++) {
        $PS = [powershell]::Create()
        $PS.RunspacePool = $Pool
        [void]$PS.AddScript($Worker.ToString()).AddArgument($Queue).AddArgument($Results).AddArgument($Shared)
        $Jobs += [PSCustomObject]@{ PS = $PS; Handle = $PS.BeginInvoke() }
    }

    # Monitor: refresh token for all workers, print progress
    $LastPrint = Get-Date
    while (@($Jobs | Where-Object { -not $_.Handle.IsCompleted }).Count -gt 0) {
        Start-Sleep -Seconds 2
        Update-GraphToken -Headers $Headers
        $Shared.Token = $Headers['Authorization']

        if (((Get-Date) - $LastPrint).TotalSeconds -ge 30) {
            $LastPrint = Get-Date
            $Done = [Math]::Min($Results.Count, $Total)
            Write-Host ("  Owners resolved: {0}/{1} ({2}%)" -f $Done, $Total, [int](100 * $Done / $Total)) -ForegroundColor DarkGray
        }
    }

    foreach ($J in $Jobs) {
        try { $J.PS.EndInvoke($J.Handle) | Out-Null } catch { }
        if ($J.PS.HadErrors) {
            foreach ($E in $J.PS.Streams.Error) { Write-Host "  [WORKER ERROR] $($E.ToString())" -ForegroundColor Red }
        }
        $J.PS.Dispose()
    }
    $Pool.Close()
    $Pool.Dispose()

    # Collect
    $Map    = @{}
    $Failed = 0
    foreach ($K in $Results.Keys) {
        $V = $Results[$K]
        $Map[$K] = @{ Name = $V.Name; UPN = $V.UPN }
        if ($V.Failed) { $Failed++ }
    }
    Write-Host ("  Owners resolved: {0}/{1}" -f $Map.Count, $Total) -ForegroundColor DarkGray
    if ($Failed -gt 0) {
        Write-Host "  [WARN] Owner lookup failed for $Failed devices (left blank)." -ForegroundColor Red
    }
    return $Map
}

# -------------------------------------------------------
# STEP 3 — FETCH ENTRA ID WINDOWS DEVICE OBJECTS
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 3/7] Fetching Entra ID Windows device objects..." -ForegroundColor Yellow
Write-Host "  Filter : operatingSystem eq 'Windows' (server-side)" -ForegroundColor DarkGray

# NOTE: $filter and $select CANNOT be combined on the /devices endpoint reliably.
# Fetch with $filter only, no $select — all fields returned, owner lookup done separately.
$EntraUri     = "https://graph.microsoft.com/v1.0/devices?`$filter=operatingSystem eq 'Windows'&`$top=999"
$EntraDevices = Invoke-GraphGetAll -Uri $EntraUri -Headers $AuthHeaders -Label "Entra ID Windows devices"

Write-Host "  [OK] Total Entra ID Windows devices fetched: $($EntraDevices.Count)" -ForegroundColor Green

if ($EntraDevices.Count -eq 0) {
    Write-Host "  [WARN] No Entra ID Windows devices found. Exiting." -ForegroundColor Red
    Stop-Transcript
    exit 0
}

# -------------------------------------------------------
# STEP 4 — FETCH INTUNE WINDOWS MANAGED DEVICES
# Used for: IntuneDeviceId, SerialNumber (primary source), IntunePrimaryUser
# Join key : Intune azureADDeviceId = Entra ID deviceId
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 4/7] Fetching Intune Windows managed devices (for cross-reference)..." -ForegroundColor Yellow
Write-Host "  Filter : operatingSystem eq 'Windows' (server-side)" -ForegroundColor DarkGray

$IntuneUri     = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$top=999"
$IntuneDevices = Invoke-GraphGetAll -Uri $IntuneUri -Headers $AuthHeaders -Label "Intune Windows devices"

Write-Host "  [OK] Total Intune Windows devices fetched: $($IntuneDevices.Count)" -ForegroundColor Green

# Build Intune lookup hashtable keyed by azureADDeviceId (= Entra deviceId)
# Stores: IntuneDeviceId, SerialNumber, PrimaryUser (userPrincipalName)
$IntuneMap = @{}
foreach ($D in $IntuneDevices) {
    $Key = [string]$D.azureADDeviceId
    if (-not [string]::IsNullOrWhiteSpace($Key) -and -not $IntuneMap.ContainsKey($Key)) {
        $IntuneMap[$Key] = [PSCustomObject]@{
            IntuneDeviceId = [string]$D.id
            SerialNumber   = [string]$D.serialNumber
            ManagementAgent = [string](Get-SafeProp $D 'managementAgent')
        }
    }
}
Write-Host "  [OK] Intune lookup map built: $($IntuneMap.Count) unique entries." -ForegroundColor Green

# -------------------------------------------------------
# STEP 5 — FETCH AUTOPILOT DEVICES
# Used for: AutopilotObjectId, SerialNumber (fallback if no Intune match)
# Join key : Autopilot azureActiveDirectoryDeviceId = Entra ID deviceId
# NOTE: $select and $expand cannot be combined with $top on this endpoint (HTTP 500).
#       Fetch without $select or $top restrictions.
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 5/7] Fetching Windows Autopilot device objects (for cross-reference)..." -ForegroundColor Yellow

$AutopilotUri     = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities"
$AutopilotDevices = Invoke-GraphGetAll -Uri $AutopilotUri -Headers $AuthHeaders -Label "Autopilot devices"

Write-Host "  [OK] Total Autopilot devices fetched: $($AutopilotDevices.Count)" -ForegroundColor Green

# Build Autopilot lookup hashtable keyed by azureActiveDirectoryDeviceId
$AutopilotMap = @{}
foreach ($A in $AutopilotDevices) {
    $Key = [string]$A.azureActiveDirectoryDeviceId   # FIX: Autopilot property name (not azureADDeviceId)
    if (-not [string]::IsNullOrWhiteSpace($Key) -and -not $AutopilotMap.ContainsKey($Key)) {
        $AutopilotMap[$Key] = [PSCustomObject]@{
            AutopilotObjectId = [string]$A.id
            SerialNumber      = [string]$A.serialNumber
        }
    }
}
Write-Host "  [OK] Autopilot lookup map built: $($AutopilotMap.Count) unique entries." -ForegroundColor Green

# -------------------------------------------------------
# STEP 6 — ENRICH ENTRA ID DEVICES + CATEGORIZE
# For each Entra ID device:
#   - Lookup registered owner (single Graph call per device — batched awareness)
#   - Match Intune record  → IntuneDeviceId, SerialNumber, PrimaryUser
#   - Match Autopilot record → AutopilotObjectId, SerialNumber (fallback)
#   - Calculate DaysSinceSignIn
#   - Route to correct aging bucket
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 6/7] Enriching devices and categorizing into 11 buckets..." -ForegroundColor Yellow
Write-Host "  Owner lookup (if enabled) can take a long time on large tenants..." -ForegroundColor DarkGray

$Today = Get-Date

# Registered owners (name + UPN) fetched up front in batches
$OwnerMap = @{}
if ($IncludeOwner) {
    Write-Host "  Fetching registered owners via `$batch, $OwnerThreads parallel threads (20 devices per call)..." -ForegroundColor DarkGray
    $OwnerMap = Get-DeviceOwnerMap -Devices $EntraDevices -Headers $AuthHeaders -Threads $OwnerThreads
    Write-Host "  [OK] Owner lookup complete: $($OwnerMap.Count) devices resolved." -ForegroundColor Green
}
else {
    Write-Host "  [SKIP] Owner lookup disabled (`$IncludeOwner = `$false). Owner columns left blank." -ForegroundColor DarkGray
}

# 11 aging buckets
$Buckets = @(
    [PSCustomObject]@{ SheetName="01 - Active (0-7 Days)";          Min=0;   Max=7;            Table="Tbl01"; Style="Medium1"; Title="ACTIVE — Entra ID Windows Devices with Sign-In Within Last 7 Days" }
    [PSCustomObject]@{ SheetName="02 - Normal (8-30 Days)";         Min=8;   Max=30;           Table="Tbl02"; Style="Medium1"; Title="NORMAL — Entra ID Windows Devices Not Signed In for 8 to 30 Days" }
    [PSCustomObject]@{ SheetName="03 - Watch (31-60 Days)";         Min=31;  Max=60;           Table="Tbl03"; Style="Medium2"; Title="WATCH — Entra ID Windows Devices Not Signed In for 31 to 60 Days" }
    [PSCustomObject]@{ SheetName="04 - Aging (61-90 Days)";         Min=61;  Max=90;           Table="Tbl04"; Style="Medium2"; Title="AGING — Entra ID Windows Devices Not Signed In for 61 to 90 Days" }
    [PSCustomObject]@{ SheetName="05 - Concern (91-120 Days)";      Min=91;  Max=120;          Table="Tbl05"; Style="Medium3"; Title="CONCERN — Entra ID Windows Devices Not Signed In for 91 to 120 Days" }
    [PSCustomObject]@{ SheetName="06 - Warning (121-150 Days)";     Min=121; Max=150;          Table="Tbl06"; Style="Medium3"; Title="WARNING — Entra ID Windows Devices Not Signed In for 121 to 150 Days" }
    [PSCustomObject]@{ SheetName="07 - Elevated (151-180 Days)";    Min=151; Max=180;          Table="Tbl07"; Style="Medium4"; Title="ELEVATED — Entra ID Windows Devices Not Signed In for 151 to 180 Days" }
    [PSCustomObject]@{ SheetName="08 - High Risk (181-240 Days)";   Min=181; Max=240;          Table="Tbl08"; Style="Medium4"; Title="HIGH RISK — Entra ID Windows Devices Not Signed In for 181 to 240 Days" }
    [PSCustomObject]@{ SheetName="09 - Critical (241-365 Days)";    Min=241; Max=365;          Table="Tbl09"; Style="Medium6"; Title="CRITICAL — Entra ID Windows Devices Not Signed In for 241 to 365 Days" }
    [PSCustomObject]@{ SheetName="10 - Stale (366-700 Days)";       Min=366; Max=700;          Table="Tbl10"; Style="Medium6"; Title="STALE — Entra ID Windows Devices Not Signed In for 366 to 700 Days" }
    [PSCustomObject]@{ SheetName="11 - Defunct (701+ Days)";        Min=701; Max=[int]::MaxValue; Table="Tbl11"; Style="Medium6"; Title="DEFUNCT — Entra ID Windows Devices Not Signed In for Over 700 Days" }
)

$BucketData = @{}
foreach ($B in $Buckets) {
    $BucketData[$B.SheetName] = [System.Collections.Generic.List[object]]::new()
}

$ProcessedCount  = 0
$SkippedCount    = 0
$TotalDevices    = $EntraDevices.Count

foreach ($Device in $EntraDevices) {

    $ProcessedCount++

    # Progress indicator every 50 devices
    if ($ProcessedCount % 50 -eq 0 -or $ProcessedCount -eq $TotalDevices) {
        Write-Host ("  Processed {0}/{1} devices..." -f $ProcessedCount, $TotalDevices) -ForegroundColor DarkGray
    }

    # --- Parse approximateLastSignInDateTime ---
    $LastSignIn    = $null
    $DaysSinceSignIn = 0

    if (-not [string]::IsNullOrWhiteSpace([string]$Device.approximateLastSignInDateTime)) {
        try {
            $LastSignIn      = [datetime]::Parse([string]$Device.approximateLastSignInDateTime)
            $DaysSinceSignIn = [int]($Today - $LastSignIn).TotalDays
        }
        catch {
            # Cannot parse — treat as never signed in
            $LastSignIn      = $null
            $DaysSinceSignIn = [int]::MaxValue
        }
    }
    else {
        # Never signed in — route to Defunct bucket
        $DaysSinceSignIn = [int]::MaxValue
    }

    # --- Registered owner (from pre-fetched batch map) ---
    $DeviceOwnerName = ""
    $DeviceOwnerUPN  = ""
    $ObjKey = [string]$Device.id
    if ($OwnerMap.ContainsKey($ObjKey)) {
        $DeviceOwnerName = $OwnerMap[$ObjKey].Name
        $DeviceOwnerUPN  = $OwnerMap[$ObjKey].UPN
    }

    # --- Cross-reference Intune (primary source) ---
    $EntraDeviceId  = [string]$Device.deviceId   # This is the azureADDeviceId used in Intune/Autopilot
    $IntuneDeviceId = ""
    $SerialNumber   = ""
    $AutopilotObjId = ""
    $ManagementAgent = ""
    $IntuneRecord   = "No"

    if ($IntuneMap.ContainsKey($EntraDeviceId)) {
        $IntuneRef      = $IntuneMap[$EntraDeviceId]
        $IntuneDeviceId = $IntuneRef.IntuneDeviceId
        $SerialNumber   = $IntuneRef.SerialNumber
        $ManagementAgent = $IntuneRef.ManagementAgent
        $IntuneRecord   = "Yes"
    }

    # --- Cross-reference Autopilot (serial fallback + ObjectId) ---
    if ($AutopilotMap.ContainsKey($EntraDeviceId)) {
        $APRef          = $AutopilotMap[$EntraDeviceId]
        $AutopilotObjId = $APRef.AutopilotObjectId
        # Only use Autopilot serial if Intune didn't provide one
        if ([string]::IsNullOrWhiteSpace($SerialNumber)) {
            $SerialNumber = $APRef.SerialNumber
        }
    }

    # --- Join type (derived from trustType) and MDM info (from Entra device object) ---
    $TrustTypeVal  = [string](Get-SafeProp $Device 'trustType')
    $JoinTypeLabel = switch ($TrustTypeVal) {
        "AzureAd"   { "Entra Joined" }
        "ServerAd"  { "Hybrid Entra Joined" }
        "Workplace" { "Entra Registered" }
        default     { "Unknown" }
    }
    $MdmAppIdVal   = [string](Get-SafeProp $Device 'mdmAppId')
    $MdmProvider   = if ([string]::IsNullOrWhiteSpace($MdmAppIdVal)) { "None" }
                     elseif ($MdmAppIdVal -eq "0000000a-0000-0000-c000-000000000000") { "Intune" }
                     else { "Other ($MdmAppIdVal)" }
    $IsManagedText = [string](Get-SafeProp $Device 'isManaged')
    $IsCompliantText = [string](Get-SafeProp $Device 'isCompliant')

    # Security settings management (Entra managementType; MicrosoftSense = Defender for Endpoint)
    $MgmtTypeVal = [string](Get-SafeProp $Device 'managementType')
    $SecSettingsMgmt = if ($MgmtTypeVal -eq "MicrosoftSense") { "Microsoft Defender for Endpoint" } else { $MgmtTypeVal }

    # Registered date
    $RegRaw = [string](Get-SafeProp $Device 'registrationDateTime')
    $RegisteredText = ""
    if (-not [string]::IsNullOrWhiteSpace($RegRaw)) {
        try   { $RegisteredText = ([datetime]::Parse($RegRaw)).ToString("yyyy-MM-dd HH:mm:ss") }
        catch { $RegisteredText = $RegRaw }
    }

    # --- Build output row ---
    $Row = [PSCustomObject]@{
        DeviceName                    = [string]$Device.displayName
        EntraDeviceId                 = $EntraDeviceId
        EntraObjectId                 = [string]$Device.id
        OperatingSystem               = [string]$Device.operatingSystem
        OSVersion                     = [string]$Device.operatingSystemVersion
        AccountEnabled                = [string]$Device.accountEnabled
        TrustType                     = [string]$Device.trustType
        JoinType                      = $JoinTypeLabel
        DeviceOwnerName               = $DeviceOwnerName
        DeviceOwnerUPN                = $DeviceOwnerUPN
        MDMProvider                   = $MdmProvider
        SecuritySettingsManagement    = $SecSettingsMgmt
        IsManaged                     = $IsManagedText
        IsCompliant                   = $IsCompliantText
        RegisteredDateTime            = $RegisteredText
        ApproximateLastSignInDateTime = if ($null -ne $LastSignIn) { $LastSignIn.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never Signed In" }
        DaysSinceSignIn               = if ($DaysSinceSignIn -eq [int]::MaxValue) { 99999 } else { $DaysSinceSignIn }
        SerialNumber                  = $SerialNumber
        IntuneDeviceId                = $IntuneDeviceId
        AutopilotObjectId             = $AutopilotObjId
        IntuneRecord                  = $IntuneRecord
        ManagementAgent               = $ManagementAgent
    }

    # --- Route to correct bucket ---
    $Routed = $false
    foreach ($B in $Buckets) {
        if ($Row.DaysSinceSignIn -ge $B.Min -and $Row.DaysSinceSignIn -le $B.Max) {
            $BucketData[$B.SheetName].Add($Row)
            $Routed = $true
            break
        }
    }

    if (-not $Routed) {
        # Safety net — route anything unmatched to Defunct
        $BucketData["11 - Defunct (701+ Days)"].Add($Row)
        $SkippedCount++
    }
}

Write-Host "  [OK] Enrichment and categorization complete." -ForegroundColor Green
Write-Host "  Total processed : $ProcessedCount" -ForegroundColor DarkGray

Write-Host ""
Write-Host "  BUCKET SUMMARY:" -ForegroundColor Cyan
foreach ($B in $Buckets) {
    Write-Host ("       {0,-45} : {1} devices" -f $B.SheetName, $BucketData[$B.SheetName].Count) -ForegroundColor DarkGray
}

# -------------------------------------------------------
# STEP 7 — EXPORT TO EXCEL (11 SHEETS)
# -------------------------------------------------------
Write-Host ""
Write-Host "[STEP 7/7] Writing Excel report to disk..." -ForegroundColor Yellow

if (Test-Path $ExcelFile) { Remove-Item $ExcelFile -Force }

$SheetNumber = 0

try {
    foreach ($B in $Buckets) {
        $SheetNumber++
        $Data = $BucketData[$B.SheetName]
        Write-Host ("  [{0,2}/11] Writing: {1} ({2} devices)..." -f $SheetNumber, $B.SheetName, $Data.Count) -ForegroundColor DarkGray

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
# FINAL SUMMARY
# -------------------------------------------------------
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "  REPORT COMPLETE" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "  Excel File : $ExcelFile" -ForegroundColor White
Write-Host "  Log File   : $LogFile" -ForegroundColor White
Write-Host ""
Write-Host "  DEVICE COUNT SUMMARY:" -ForegroundColor Cyan
Write-Host ("  {0,-45} {1}" -f "Total Entra ID Windows Devices",   $EntraDevices.Count)    -ForegroundColor White
Write-Host ("  {0,-45} {1}" -f "Total Intune Windows Devices",     $IntuneDevices.Count)   -ForegroundColor White
Write-Host ("  {0,-45} {1}" -f "Total Autopilot Devices",          $AutopilotDevices.Count) -ForegroundColor White
Write-Host ("  {0,-45} {1}" -f "Intune Matches Found",             $IntuneMap.Count)        -ForegroundColor White
Write-Host ("  {0,-45} {1}" -f "Autopilot Matches Found",          $AutopilotMap.Count)     -ForegroundColor White
Write-Host ""
foreach ($B in $Buckets) {
    Write-Host ("  {0,-45} {1}" -f $B.SheetName, $BucketData[$B.SheetName].Count) -ForegroundColor White
}
Write-Host ""
Write-Host "  Completed : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host ""

Stop-Transcript
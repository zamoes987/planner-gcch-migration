<#
.SYNOPSIS
  One script to migrate Microsoft Planner (basic plans) and Planner premium plans (Project Plan 3 /
  formerly Project for the web) from a Microsoft 365 Commercial tenant to a Microsoft 365 GCC High tenant.

.DESCRIPTION
  Run it step by step with -Step. Every step logs to logs\ and is safe to rerun.

    Init            Create config\ templates next to this script (settings.json, users.csv, urlmap.csv).
    Setup           Create the app registration + certificate + Graph permissions for -Side Source|Destination
                    and write the IDs into config\settings.json (sign in as an admin of that tenant).
    Discover        List every Microsoft 365 group in the source that has plans -> config\groups.csv.
    Preflight       Go/no-go checks against your real tenants (-Scope All|Basic|Premium).
    Export          Export basic plans from the commercial tenant to export\basic (read-only on source).
    ExportTaskChat  Add task chat (the 2026 replacement for comments) to the export. Signs in as a user.
    DryRun          Resolve every user/group mapping against GCC High; writes nothing.
    Import          Import basic plans into GCC High (resumable). -OnlyIds <planId> for a pilot.
    Verify          Reconcile GCC High against the export -> logs\verify-basic.csv.
    ExportPremium   Export premium plans from the commercial Dataverse environment.
    ImportPremium   Import premium plans into GCC High via the Project schedule APIs (resumable).

  Facts this relies on, with sources, are in README.md. Project Online itself retired on 2026-09-30
  and cannot be migrated through any API.

.EXAMPLE
  .\Migrate-PlannerToGccHigh.ps1 -Step Init
  .\Migrate-PlannerToGccHigh.ps1 -Step Setup -Side Source
  .\Migrate-PlannerToGccHigh.ps1 -Step Setup -Side Destination
  .\Migrate-PlannerToGccHigh.ps1 -Step Discover
  .\Migrate-PlannerToGccHigh.ps1 -Step Preflight
  .\Migrate-PlannerToGccHigh.ps1 -Step Export
  .\Migrate-PlannerToGccHigh.ps1 -Step ExportTaskChat
  .\Migrate-PlannerToGccHigh.ps1 -Step DryRun
  .\Migrate-PlannerToGccHigh.ps1 -Step Import -OnlyIds <onePlanId>
  .\Migrate-PlannerToGccHigh.ps1 -Step Import
  .\Migrate-PlannerToGccHigh.ps1 -Step Verify
  .\Migrate-PlannerToGccHigh.ps1 -Step ExportPremium
  .\Migrate-PlannerToGccHigh.ps1 -Step ImportPremium
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateSet('Init', 'Setup', 'Discover', 'Preflight', 'Export', 'ExportTaskChat', 'DryRun', 'Import', 'Verify', 'ExportPremium', 'ImportPremium')]
    [string]$Step,
    [ValidateSet('Source', 'Destination')][string]$Side,
    [ValidateSet('All', 'Basic', 'Premium')][string]$Scope = 'All',
    [string]$SettingsPath = (Join-Path $PSScriptRoot 'config\settings.json'),
    [string[]]$OnlyIds,
    [switch]$TryNamedResources,
    [switch]$SkipComments,
    [int]$SampleTasks = 200,
    [int]$CertMonths = 3
)
# No StrictMode: Graph/Dataverse omit optional properties (e.g. @odata.nextLink) and strict mode would throw on them.
$ErrorActionPreference = 'Stop'

#region ==================== settings, logging, state ====================

function Read-MigrationSettings {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { throw "Settings file not found: $Path. Run: .\Migrate-PlannerToGccHigh.ps1 -Step Init" }
    $s = Get-Content $Path -Raw | ConvertFrom-Json
    $root = Split-Path (Resolve-Path $Path) -Parent | Split-Path -Parent
    foreach ($name in 'ExportFolder', 'StateFolder', 'LogFolder') {
        if (-not [System.IO.Path]::IsPathRooted($s.$name)) { $s.$name = Join-Path $root $s.$name }
        New-Item -ItemType Directory -Force -Path $s.$name | Out-Null
    }
    $script:LogFile = Join-Path $s.LogFolder ("migration-{0:yyyyMMdd}.log" -f (Get-Date))
    return $s
}

function Write-MigLog {
    param([ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO', [Parameter(Mandatory)][string]$Message)
    $line = "{0:u} [{1}] {2}" -f (Get-Date), $Level, $Message
    $color = @{ INFO = 'Gray'; WARN = 'Yellow'; ERROR = 'Red'; OK = 'Green' }[$Level]
    Write-Host $line -ForegroundColor $color
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line }
}

function Add-CsvRow {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Row)
    if (Test-Path $Path) { $Row | Export-Csv -Path $Path -Append -NoTypeInformation }
    else { $Row | Export-Csv -Path $Path -NoTypeInformation }
}

function Get-IdMap {
    # source-id -> destination-id, persisted after every create so reruns resume instead of duplicating.
    # Ordinal (case-sensitive) keys: Planner IDs are case-sensitive per the Graph docs.
    param([Parameter(Mandatory)][string]$Path)
    $map = [System.Collections.Hashtable]::new([System.StringComparer]::Ordinal)
    if (Test-Path $Path) {
        (Get-Content $Path -Raw | ConvertFrom-Json -AsHashtable).GetEnumerator() | ForEach-Object { $map[$_.Key] = $_.Value }
    }
    return $map
}

function Save-IdMap {
    param([Parameter(Mandatory)][hashtable]$Map, [Parameter(Mandatory)][string]$Path)
    $tmp = "$Path.tmp"
    $Map | ConvertTo-Json -Depth 3 | Set-Content -Path $tmp -Encoding utf8
    Move-Item -Path $tmp -Destination $Path -Force   # write-then-rename: a crash never leaves a half-written state file
}

function Get-UserMap([string]$ConfigDir) {
    $m = @{}
    $p = Join-Path $ConfigDir 'users.csv'
    if (Test-Path $p) { Import-Csv $p | Where-Object sourceUpn | ForEach-Object { $m[$_.sourceUpn.Trim().ToLowerInvariant()] = $_.destUpn.Trim().ToLowerInvariant() } }
    $m
}
function Get-DomainMap($Settings) {
    $m = @{}
    if ($Settings.DomainMap) { $Settings.DomainMap.PSObject.Properties | ForEach-Object { $m[$_.Name.ToLowerInvariant()] = ([string]$_.Value).ToLowerInvariant() } }
    $m
}

#endregion

#region ==================== pure transforms (unit tested) ====================

function ConvertTo-PlannerRefKey {
    # Planner reference keys are URLs with . : % @ # encoded (Graph: plannerExternalReferences). '%' first so %20 survives.
    param([Parameter(Mandatory)][string]$Url)
    $Url.Replace('%', '%25').Replace('.', '%2E').Replace(':', '%3A').Replace('@', '%40').Replace('#', '%23')
}

function ConvertFrom-PlannerRefKey {
    param([Parameter(Mandatory)][string]$Key)
    $k = $Key -replace '%2[Ee]', '.' -replace '%3[Aa]', ':' -replace '%40', '@' -replace '%23', '#'
    $k -replace '%25', '%'
}

function Get-MappedUrl {
    # Rewrites a URL using the longest matching source prefix from urlmap.csv; unchanged if nothing matches.
    param([Parameter(Mandatory)][string]$Url, [object[]]$UrlMap = @())
    $hit = $UrlMap | Where-Object { $Url.StartsWith($_.sourcePrefix, [StringComparison]::OrdinalIgnoreCase) } |
        Sort-Object { $_.sourcePrefix.Length } -Descending | Select-Object -First 1
    if ($hit) { return $hit.destPrefix + $Url.Substring($hit.sourcePrefix.Length) }
    return $Url
}

function Resolve-DestUpn {
    # users.csv row wins; otherwise DomainMap (contoso.com -> contoso.us); otherwise $null (unmapped).
    param([string]$SourceUpn, [hashtable]$UserMap = @{}, [hashtable]$DomainMap = @{})
    if ([string]::IsNullOrWhiteSpace($SourceUpn)) { return $null }
    $key = $SourceUpn.ToLowerInvariant()
    if ($UserMap.ContainsKey($key)) { return $UserMap[$key] }
    $parts = $key.Split('@')
    if ($parts.Count -eq 2 -and $DomainMap.ContainsKey($parts[1])) { return "$($parts[0])@$($DomainMap[$parts[1]])" }
    return $null
}

function Get-SortedByOrderHint {
    # Planner order hints sort by ordinal character value; shorter first when one is a prefix (Graph: order hints).
    param([object[]]$Items, [string]$Property = 'orderHint')
    if (-not $Items) { return @() }
    $list = [System.Collections.Generic.List[object]]::new()
    $Items | ForEach-Object { $list.Add($_) }
    $list.Sort([System.Comparison[object]] { param($x, $y) [string]::CompareOrdinal([string]$x.$Property, [string]$y.$Property) })
    return , $list.ToArray()
}

function Get-AppendOrderHint {
    # "<previous> <next>!" with next empty = place after previous; first item in an empty list = " !".
    param([string]$PreviousHint)
    if ([string]::IsNullOrEmpty($PreviousHint)) { return ' !' }
    return "$PreviousHint !"
}

function ConvertTo-PlainText {
    param([string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return '' }
    $t = $Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', '' -replace '(?i)<br\s*/?>', "`n" -replace '(?i)</(p|div|li)>', "`n" -replace '<[^>]+>', ''
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    ($t -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join "`n"
}

function ConvertTo-CommentPost {
    # Task chat messages (Graph beta plannerTaskChatMessage) -> the same shape as group posts, so one formatter handles both.
    param([object[]]$Messages, [hashtable]$NameById = @{})
    foreach ($m in $Messages) {
        if (-not $m -or $m.deletedTime) { continue }
        $uid = if ($m.createdBy -and $m.createdBy.user) { [string]$m.createdBy.user.id } else { '' }
        $name = if ($uid -and $NameById[$uid]) { $NameById[$uid] } else { 'unknown' }
        [pscustomobject]@{ createdDateTime = $m.createdDateTime; from = [pscustomobject]@{ emailAddress = [pscustomobject]@{ name = $name } }; body = [pscustomobject]@{ content = $m.content } }
    }
}

function Format-CommentsText {
    # Comments and task chat cannot be written into GCC High as conversations; they are carried as text.
    param([object[]]$Posts)
    $Posts = @($Posts | Where-Object { $_ })
    if (-not $Posts) { return '' }
    $lines = foreach ($p in ($Posts | Sort-Object { [datetime]$_.createdDateTime })) {
        $who = if ($p.from -and $p.from.emailAddress) { $p.from.emailAddress.name } else { 'unknown' }
        "[{0:yyyy-MM-dd HH:mm}Z] {1}: {2}" -f ([datetime]$p.createdDateTime).ToUniversalTime(), $who, (ConvertTo-PlainText $p.body.content)
    }
    "`n`n--- Comments migrated from commercial tenant ---`n" + ($lines -join "`n")
}

function Get-RetryDelaySeconds {
    # Honor Retry-After (Graph and Dataverse send it on 429); otherwise exponential backoff capped at 120s.
    param($Headers, [int]$Attempt)
    $ra = $null
    if ($Headers -is [System.Collections.IDictionary]) {
        foreach ($k in @($Headers.Keys)) { if ($k -ieq 'Retry-After') { $ra = @($Headers[$k])[0] } }
    }
    elseif ($Headers) {
        foreach ($kv in $Headers) { if ($kv.Key -ieq 'Retry-After') { $ra = @($kv.Value)[0] } }
    }
    $sec = 0
    if ($ra -and [int]::TryParse([string]$ra, [ref]$sec) -and $sec -gt 0) { return $sec }
    return [Math]::Min(120, [Math]::Pow(2, $Attempt))
}

function Split-IntoChunks {
    param([object[]]$Items, [int]$Size)
    $chunks = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Items.Count; $i += $Size) { $chunks.Add(@($Items[$i..([Math]::Min($i + $Size, $Items.Count) - 1)])) }
    return , $chunks.ToArray()
}

function Get-PlannerAssignmentsBody {
    # Maps source assignee IDs -> destination IDs. Returns @{ Body = <hashtable>; Unmapped = <string[]> }.
    param($SourceAssignments, [hashtable]$SourceUserUpn, [hashtable]$DestUserIdByUpn, [hashtable]$UserMap = @{}, [hashtable]$DomainMap = @{})
    $body = @{}; $unmapped = @()
    if ($SourceAssignments) {
        foreach ($p in $SourceAssignments.PSObject.Properties) {
            $srcUpn = $SourceUserUpn[$p.Name]
            $destUpn = Resolve-DestUpn -SourceUpn $srcUpn -UserMap $UserMap -DomainMap $DomainMap
            $destId = if ($destUpn) { $DestUserIdByUpn[$destUpn.ToLowerInvariant()] } else { $null }
            if ($destId) { $body[$destId] = @{ '@odata.type' = '#microsoft.graph.plannerAssignment'; orderHint = ' !' } }
            else { $unmapped += $(if ($srcUpn) { $srcUpn } else { "id:$($p.Name)" }) }
        }
    }
    @{ Body = $body; Unmapped = $unmapped }
}

function Get-TaskDetailsBody {
    # PATCH body for plannerTaskDetails: description (+comments), previewType, checklist (new keys), references (URL-mapped keys).
    param($SourceDetails, [object[]]$UrlMap = @(), [string]$CommentsText = '')
    $body = @{}
    $desc = [string]$SourceDetails.description
    if ($CommentsText) { $desc = $desc + $CommentsText }
    if ($desc) { $body.description = $desc }
    if ($SourceDetails.previewType) { $body.previewType = $SourceDetails.previewType }
    $checklist = @{}
    if ($SourceDetails.checklist) {
        $items = @($SourceDetails.checklist.PSObject.Properties | ForEach-Object { $_.Value })
        foreach ($c in (Get-SortedByOrderHint -Items $items)) {
            $checklist[[guid]::NewGuid().ToString()] = @{ '@odata.type' = '#microsoft.graph.plannerChecklistItem'; title = $c.title; isChecked = [bool]$c.isChecked }
        }
    }
    if ($checklist.Count) { $body.checklist = $checklist }
    $refs = @{}
    if ($SourceDetails.references) {
        foreach ($r in $SourceDetails.references.PSObject.Properties) {
            $url = Get-MappedUrl -Url (ConvertFrom-PlannerRefKey $r.Name) -UrlMap $UrlMap
            $ref = @{ '@odata.type' = '#microsoft.graph.plannerExternalReference' }
            if ($r.Value.alias) { $ref.alias = $r.Value.alias }
            if ($r.Value.type) { $ref.type = $r.Value.type }
            $refs[(ConvertTo-PlannerRefKey $url)] = $ref
        }
    }
    if ($refs.Count) { $body.references = $refs }
    return $body
}

function Get-AuthorityHost {
    # Microsoft Graph national cloud deployments: global login.microsoftonline.com, US Government login.microsoftonline.us.
    param([Parameter(Mandatory)][string]$AzEnvironment)
    switch ($AzEnvironment) {
        'AzureCloud' { 'https://login.microsoftonline.com' }
        'AzureUSGovernment' { 'https://login.microsoftonline.us' }
        default { throw "Unsupported AzEnvironment '$AzEnvironment' (use AzureCloud or AzureUSGovernment)" }
    }
}

function ConvertTo-FormBody([hashtable]$Fields) {
    ($Fields.GetEnumerator() | Sort-Object Key | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join '&'
}

#endregion

#region ==================== Microsoft Graph ====================

function Connect-GraphApp {
    # App-only, certificate auth. 'Global' = commercial, 'USGov' = GCC High (graph.microsoft.us).
    param([Parameter(Mandatory)]$Side)
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }   # one cloud at a time; tokens are not interchangeable
    Connect-MgGraph -Environment $Side.GraphEnvironment -TenantId $Side.TenantId -ClientId $Side.ClientId `
        -CertificateThumbprint $Side.CertificateThumbprint -NoWelcome -ContextScope Process
    $ctx = Get-MgContext
    Write-MigLog OK "Graph connected: env=$($ctx.Environment) tenant=$($ctx.TenantId) auth=$($ctx.AuthType)"
}

function Invoke-Graph {
    param(
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Uri,
        $Body,
        [hashtable]$Headers = @{},
        [int]$MaxAttempts = 8
    )
    for ($attempt = 1; ; $attempt++) {
        $sc = $null; $rh = $null
        $p = @{ Method = $Method; Uri = $Uri; Headers = $Headers; SkipHttpErrorCheck = $true
                StatusCodeVariable = 'sc'; ResponseHeadersVariable = 'rh'; OutputType = 'PSObject' }
        if ($null -ne $Body) { $p.Body = ($Body | ConvertTo-Json -Depth 20 -Compress); $p.ContentType = 'application/json' }
        $resp = Invoke-MgGraphRequest @p
        $status = if ($null -ne $sc) { [int]$sc }
        else {
            # Fail safe if the status variable wasn't populated: never treat a Graph error payload as success.
            $hasError = ($resp -is [System.Collections.IDictionary] -and $resp.Contains('error')) -or
                        ($resp -isnot [System.Collections.IDictionary] -and $resp -and $resp.PSObject.Properties.Name -contains 'error')
            if (-not $hasError) { 200 }
            else { switch -Regex ([string]$resp.error.code) { '^(TooManyRequests|activityLimitReached)$' { 429 } '^(serviceNotAvailable|ServiceUnavailable)$' { 503 } default { 400 } } }
        }
        if ($status -lt 400) { return $resp }
        if (($status -in @(429, 503, 504)) -and ($attempt -lt $MaxAttempts)) {
            $wait = Get-RetryDelaySeconds -Headers $rh -Attempt $attempt
            Write-MigLog WARN "Graph $status on $Method $Uri - waiting ${wait}s (attempt $attempt/$MaxAttempts)"
            Start-Sleep -Seconds $wait
            continue
        }
        $msg = try { $resp.error.message } catch { "$resp" }
        $ex = [System.Exception]::new("Graph $Method $Uri failed: HTTP $status - $msg")
        $ex.Data['StatusCode'] = $status
        throw $ex
    }
}

function Get-GraphAll {
    param([Parameter(Mandatory)][string]$Uri)
    $out = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-Graph -Uri $next
        if ($page.value) { $page.value | ForEach-Object { $out.Add($_) } }
        $next = $page.'@odata.nextLink'
    }
    return , $out.ToArray()
}

#endregion

#region ==================== Dataverse Web API ====================

function Invoke-OAuthForm([string]$Uri, [hashtable]$Fields) {
    $r = Invoke-WebRequest -Method POST -Uri $Uri -Body (ConvertTo-FormBody $Fields) -ContentType 'application/x-www-form-urlencoded' -SkipHttpErrorCheck
    [pscustomobject]@{ Status = [int]$r.StatusCode; Json = $(if ($r.Content) { $r.Content | ConvertFrom-Json } else { $null }) }
}

function Set-DvTokenFromResponse($Json) {
    $script:DvToken = $Json.access_token
    $script:DvTokenExpires = (Get-Date).AddSeconds([int]$Json.expires_in)
    if ($Json.refresh_token) { $script:DvRefresh = $Json.refresh_token }
}

function Request-DeviceCodeToken {
    # OAuth 2.0 device authorization grant (Microsoft identity platform), against the cloud's own authority.
    $d = $script:DvDevice
    $base = "$($d.Authority)/$($d.TenantId)/oauth2/v2.0"
    $r = Invoke-OAuthForm "$base/devicecode" @{ client_id = $d.ClientId; scope = $d.Scope }
    if ($r.Status -ge 400) { throw "Device code request failed: $($r.Json.error) $($r.Json.error_description)" }
    Write-Host $r.Json.message -ForegroundColor Cyan
    $interval = [Math]::Max(1, [int]$r.Json.interval)
    $deadline = (Get-Date).AddSeconds([int]$r.Json.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        $t = Invoke-OAuthForm "$base/token" @{ grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; client_id = $d.ClientId; device_code = $r.Json.device_code }
        if ($t.Status -lt 400) { Set-DvTokenFromResponse $t.Json; return }
        if ($t.Json.error -eq 'authorization_pending') { continue }
        if ($t.Json.error -eq 'slow_down') { $interval += 5; continue }
        throw "Device code sign-in failed: $($t.Json.error) $($t.Json.error_description)"
    }
    throw 'Device code expired before sign-in completed'
}

function Connect-Dataverse {
    # AuthMode Az: delegated sign-in via Az.Accounts (Microsoft Learn Dataverse PowerShell quickstart pattern).
    # AuthMode DeviceCode: your own public-client app registration with the Dynamics CRM user_impersonation permission.
    param([Parameter(Mandatory)][string]$EnvironmentUrl, [Parameter(Mandatory)][string]$AzEnvironment, [Parameter(Mandatory)][string]$TenantId,
          [ValidateSet('Az', 'DeviceCode')][string]$AuthMode = 'Az', [string]$ClientId)
    $script:DvUrl = $EnvironmentUrl.TrimEnd('/') + '/'
    $script:DvToken = $null; $script:DvRefresh = $null; $script:DvAuth = $AuthMode
    if ($AuthMode -eq 'Az') { Connect-AzAccount -Environment $AzEnvironment -Tenant $TenantId -Scope Process | Out-Null }
    else {
        if (-not $ClientId) { throw 'DataverseAuth=DeviceCode needs DataverseClientId (public-client app with the Dynamics CRM user_impersonation delegated permission)' }
        # Learn: public clients use "<environment-url>/user_impersonation"; offline_access returns a refresh token.
        $script:DvDevice = @{ Authority = (Get-AuthorityHost $AzEnvironment); TenantId = $TenantId; ClientId = $ClientId
                              Scope = "$($EnvironmentUrl.TrimEnd('/'))/user_impersonation offline_access" }
        Request-DeviceCodeToken
    }
    $who = Invoke-Dv -Path 'WhoAmI'
    Write-MigLog OK "Dataverse connected ($AuthMode): $script:DvUrl userId=$($who.UserId)"
    return $who
}

function Connect-DataverseSide($SideCfg) {
    $mode = if ($SideCfg.DataverseAuth) { $SideCfg.DataverseAuth } else { 'Az' }
    Connect-Dataverse -EnvironmentUrl $SideCfg.DataverseUrl -AzEnvironment $SideCfg.AzEnvironment -TenantId $SideCfg.TenantId -AuthMode $mode -ClientId $SideCfg.DataverseClientId
}

function Get-DvToken {
    if ($script:DvToken -and $script:DvTokenExpires -gt (Get-Date).AddMinutes(5)) { return $script:DvToken }
    if ($script:DvAuth -eq 'DeviceCode') {
        $d = $script:DvDevice
        if ($script:DvRefresh) {
            $t = Invoke-OAuthForm "$($d.Authority)/$($d.TenantId)/oauth2/v2.0/token" @{ grant_type = 'refresh_token'; client_id = $d.ClientId; refresh_token = $script:DvRefresh; scope = $d.Scope }
            if ($t.Status -lt 400) { Set-DvTokenFromResponse $t.Json; return $script:DvToken }
        }
        Request-DeviceCodeToken
        return $script:DvToken
    }
    $t = Get-AzAccessToken -ResourceUrl $script:DvUrl -AsSecureString
    $script:DvToken = ConvertFrom-SecureString -SecureString $t.Token -AsPlainText
    $script:DvTokenExpires = $t.ExpiresOn.LocalDateTime
    return $script:DvToken
}

function Invoke-Dv {
    param(
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Path,     # relative to <env>/api/data/v9.2/ or an absolute nextLink
        $Body,
        [hashtable]$ExtraHeaders = @{},
        [int]$MaxAttempts = 8
    )
    $uri = if ($Path -match '^https://') { $Path } else { $script:DvUrl + 'api/data/v9.2/' + $Path }
    for ($attempt = 1; ; $attempt++) {
        $h = @{ Authorization = "Bearer $(Get-DvToken)"; Accept = 'application/json'; 'OData-MaxVersion' = '4.0'; 'OData-Version' = '4.0' }
        $ExtraHeaders.GetEnumerator() | ForEach-Object { $h[$_.Key] = $_.Value }
        $p = @{ Method = $Method; Uri = $uri; Headers = $h; SkipHttpErrorCheck = $true }
        if ($null -ne $Body) { $p.Body = [System.Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 20 -Compress)); $p.ContentType = 'application/json; charset=utf-8' }
        $r = Invoke-WebRequest @p
        $status = [int]$r.StatusCode
        if ($status -lt 400) {
            if ($r.Content -and $r.Content.Length -gt 0) { return ($r.Content | ConvertFrom-Json) }
            return $null
        }
        if (($status -in @(429, 503)) -and ($attempt -lt $MaxAttempts)) {
            $wait = Get-RetryDelaySeconds -Headers $r.Headers -Attempt $attempt
            Write-MigLog WARN "Dataverse $status on $Method $Path - waiting ${wait}s"
            Start-Sleep -Seconds $wait
            continue
        }
        $msg = try { ($r.Content | ConvertFrom-Json).error.message } catch { [string]$r.Content }
        $ex = [System.Exception]::new("Dataverse $Method $Path failed: HTTP $status - $msg")
        $ex.Data['StatusCode'] = $status
        throw $ex
    }
}

function Get-DvAll {
    param([Parameter(Mandatory)][string]$Path)
    $hdr = @{ Prefer = 'odata.include-annotations="*",odata.maxpagesize=500' }
    $out = [System.Collections.Generic.List[object]]::new()
    $next = $Path
    while ($next) {
        $page = Invoke-Dv -Path $next -ExtraHeaders $hdr
        if ($page.value) { $page.value | ForEach-Object { $out.Add($_) } }
        $next = $page.'@odata.nextLink'
    }
    return , $out.ToArray()
}

function Get-DvTableInfo {
    # Runtime metadata discovery: table/column/relationship names are read from the environment, not assumed.
    param([Parameter(Mandatory)][string]$LogicalName)
    try { $e = Invoke-Dv -Path "EntityDefinitions(LogicalName='$LogicalName')?`$select=EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute" }
    catch { if ($_.Exception.Data['StatusCode'] -eq 404) { return $null }; throw }
    $attrs = (Invoke-Dv -Path "EntityDefinitions(LogicalName='$LogicalName')/Attributes?`$select=LogicalName,AttributeType,IsValidForCreate").value
    $rels = (Invoke-Dv -Path "EntityDefinitions(LogicalName='$LogicalName')/ManyToOneRelationships?`$select=ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName").value
    [pscustomobject]@{
        LogicalName = $LogicalName; EntitySetName = $e.EntitySetName; PrimaryId = $e.PrimaryIdAttribute; PrimaryName = $e.PrimaryNameAttribute
        CreatableColumns = @($attrs | Where-Object { $_.IsValidForCreate } | ForEach-Object { $_.LogicalName })
        Lookups = @($rels)
    }
}

function Get-DvLookupTargets {
    param([Parameter(Mandatory)]$Row)
    $result = @{}
    foreach ($p in $Row.PSObject.Properties) {
        if ($p.Name -match '^_(.+)_value$' -and $p.Value) {
            $attr = $Matches[1]
            $result[$attr] = @{ Id = [string]$p.Value; Entity = $Row."$($p.Name)@Microsoft.Dynamics.CRM.lookuplogicalname" }
        }
    }
    return $result
}

$script:DvSystemColumns = @(
    'createdon', 'createdby', 'createdonbehalfby', 'modifiedon', 'modifiedby', 'modifiedonbehalfby', 'ownerid', 'owninguser',
    'owningteam', 'owningbusinessunit', 'overriddencreatedon', 'importsequencenumber', 'timezoneruleversionnumber',
    'utcconversiontimezonecode', 'versionnumber', 'statecode', 'statuscode', 'transactioncurrencyid', 'exchangerate')

function New-DvEntityPayload {
    <# Builds a Project schedule API entity from an exported source row.
       Columns = documented allowlist + verified extras, intersected with the destination's creatable columns
       (-AllCreatable: every creatable non-system column). Lookups rebound with <navprop>@odata.bind only when the
       target is in IdMap (created, or planned in this run); otherwise dropped and reported. Source GUID reused as the key. #>
    param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)]$TableInfo, [hashtable]$SetNameByEntity = @{},
          [hashtable]$IdMap = @{}, [string[]]$Columns = @(), [switch]$AllCreatable)
    $payload = [ordered]@{ '@odata.type' = "Microsoft.Dynamics.CRM.$($TableInfo.LogicalName)" }
    $srcId = [string]$Row.($TableInfo.PrimaryId)
    $payload[$TableInfo.PrimaryId] = $(if ($IdMap.ContainsKey($srcId)) { $IdMap[$srcId] } else { $srcId })
    $dropped = @()
    $lookupAttrs = @{}; $TableInfo.Lookups | ForEach-Object { $lookupAttrs[$_.ReferencingAttribute] = $_ }
    $wanted = if ($AllCreatable) { $TableInfo.CreatableColumns | Where-Object { $_ -notin $script:DvSystemColumns -and $_ -ne $TableInfo.PrimaryId } } else { $Columns }
    $targets = Get-DvLookupTargets -Row $Row
    foreach ($col in $wanted) {
        if ($col -notin $TableInfo.CreatableColumns) { $dropped += "$col(not creatable in destination)"; continue }
        if ($lookupAttrs.ContainsKey($col)) {
            if (-not $targets.ContainsKey($col)) { continue }
            $t = $targets[$col]
            $set = $SetNameByEntity[$t.Entity]
            $destId = if ($IdMap.ContainsKey($t.Id)) { $IdMap[$t.Id] } else { $null }
            if (-not $set -or -not $destId) { $dropped += "$col(->$($t.Entity) not migrated)"; continue }
            $payload["$($lookupAttrs[$col].ReferencingEntityNavigationPropertyName)@odata.bind"] = "/$set($destId)"
        }
        elseif ($Row.PSObject.Properties.Name -contains $col -and $null -ne $Row.$col) { $payload[$col] = $Row.$col }
    }
    [pscustomobject]@{ Payload = $payload; Dropped = $dropped }
}

#endregion

#region ==================== steps ====================

function Invoke-StepInit {
    $cfgDir = Split-Path $SettingsPath -Parent
    New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null
    if (-not (Test-Path $SettingsPath)) {
        [ordered]@{
            Source      = [ordered]@{ TenantId = ''; ClientId = ''; CertificateThumbprint = ''; GraphEnvironment = 'Global'; AzEnvironment = 'AzureCloud'
                                      DataverseUrl = 'https://yourorg.crm.dynamics.com/'; DataverseAuth = 'Az'; DataverseClientId = '' }
            Destination = [ordered]@{ TenantId = ''; ClientId = ''; CertificateThumbprint = ''; GraphEnvironment = 'USGov'; AzEnvironment = 'AzureUSGovernment'
                                      DataverseUrl = 'https://yourorg.crm.microsoftdynamics.us/'; DataverseAuth = 'Az'; DataverseClientId = '' }
            DomainMap = [ordered]@{ 'contoso.com' = 'contoso.us' }
            PreflightTestGroupId = ''
            CommentsMode = 'AppendToDescription'
            ExportFolder = 'export'; StateFolder = 'state'; LogFolder = 'logs'
            Premium = [ordered]@{ ExtraColumns = [ordered]@{ msdyn_project = @(); msdyn_projecttask = @('msdyn_description', 'msdyn_progress', 'msdyn_priority') }
                                  OperationSetChunkSize = 190; WaitForOperationSetMinutes = 15 }
        } | ConvertTo-Json -Depth 6 | Set-Content -Path $SettingsPath -Encoding utf8
        Write-Host "Created $SettingsPath" -ForegroundColor Green
    } else { Write-Host "Exists, left unchanged: $SettingsPath" }
    foreach ($f in @(@{ n = 'users.csv'; h = 'sourceUpn,destUpn' }, @{ n = 'urlmap.csv'; h = 'sourcePrefix,destPrefix' })) {
        $p = Join-Path $cfgDir $f.n
        if (-not (Test-Path $p)) { Set-Content -Path $p -Value $f.h -Encoding utf8; Write-Host "Created $p" -ForegroundColor Green }
    }
    Write-Host @"

Next:
 1. In settings.json set Source.TenantId, Destination.TenantId, both DataverseUrl values (premium only),
    DomainMap, and PreflightTestGroupId (a TEST Microsoft 365 group in GCC High).
 2. users.csv: only for people whose UPN doesn't follow DomainMap. urlmap.csv: old SharePoint URL prefix -> new.
 3. .\Migrate-PlannerToGccHigh.ps1 -Step Setup -Side Source      (admin of the commercial tenant)
    .\Migrate-PlannerToGccHigh.ps1 -Step Setup -Side Destination (admin of the GCC High tenant)
"@
}

function Invoke-StepSetup {
    # Creates the app registration, a self-signed certificate (Microsoft Learn recipe), the Graph application
    # permissions (role IDs looked up at runtime from the Microsoft Graph service principal) and grants them
    # via appRoleAssignedTo (Learn: grant API permissions programmatically). Writes IDs into settings.json.
    if (-not $Side) { throw '-Side Source|Destination is required for -Step Setup' }
    $raw = Get-Content $SettingsPath -Raw | ConvertFrom-Json
    $cfg = $raw.$Side
    $perms = if ($Side -eq 'Source') { 'Tasks.Read.All', 'Group.Read.All', 'User.Read.All' } else { 'Tasks.ReadWrite.All', 'Group.Read.All', 'User.Read.All' }
    $appName = "Planner GCC High Migration ($Side)"
    $null = Read-MigrationSettings -Path $SettingsPath   # sets up logging
    $tenantArg = @{}; if ($cfg.TenantId -match '^[0-9a-fA-F-]{36}$' -and $cfg.TenantId -ne '00000000-0000-0000-0000-000000000000') { $tenantArg.TenantId = $cfg.TenantId }
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }
    Connect-MgGraph -Environment $cfg.GraphEnvironment @tenantArg -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All' -NoWelcome -ContextScope Process
    $ctx = Get-MgContext
    Write-MigLog INFO "Signed in to $($ctx.Environment) tenant $($ctx.TenantId) as $($ctx.Account)"

    $graphSp = @((Invoke-Graph -Uri "/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles").value)[0]
    if (-not $graphSp) { throw 'Microsoft Graph service principal not found in this tenant' }
    $roles = foreach ($p in $perms) {
        $r = $graphSp.appRoles | Where-Object { $_.value -eq $p -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $r) { throw "Application permission $p is not offered by Microsoft Graph in this cloud" }
        $r
    }

    $cert = New-SelfSignedCertificate -Subject "CN=$appName" -CertStoreLocation 'Cert:\CurrentUser\My' -KeyExportPolicy NonExportable `
        -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddMonths($CertMonths)
    Write-MigLog OK "Certificate created in Cert:\CurrentUser\My, thumbprint $($cert.Thumbprint), expires $($cert.NotAfter)"

    $app = Invoke-Graph -Method POST -Uri '/v1.0/applications' -Body @{
        displayName = $appName; signInAudience = 'AzureADMyOrg'
        requiredResourceAccess = @(@{ resourceAppId = '00000003-0000-0000-c000-000000000000'; resourceAccess = @($roles | ForEach-Object { @{ id = $_.id; type = 'Role' } }) })
        keyCredentials = @(@{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [Convert]::ToBase64String($cert.RawData); displayName = "CN=$appName" })
    }
    if (-not $app.appId) { throw 'Application create returned no appId' }
    $sp = Invoke-Graph -Method POST -Uri '/v1.0/servicePrincipals' -Body @{ appId = $app.appId }
    Write-MigLog OK "App '$appName' appId=$($app.appId), service principal $($sp.id)"

    foreach ($r in $roles) {
        for ($i = 1; ; $i++) {
            try {
                Invoke-Graph -Method POST -Uri "/v1.0/servicePrincipals/$($graphSp.id)/appRoleAssignedTo" -Body @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $r.id } | Out-Null
                Write-MigLog OK "Granted $($r.value)"; break
            }
            catch {
                if ($i -ge 6) { throw }
                Write-MigLog WARN "Grant $($r.value) not accepted yet ($($_.Exception.Message)); retrying (directory replication)"
                Start-Sleep -Seconds 10
            }
        }
    }
    $raw.$Side.TenantId = $ctx.TenantId; $raw.$Side.ClientId = $app.appId; $raw.$Side.CertificateThumbprint = $cert.Thumbprint
    $raw | ConvertTo-Json -Depth 8 | Set-Content -Path $SettingsPath -Encoding utf8
    Write-MigLog OK "settings.json updated for $Side. If Preflight returns 401/403 right away, wait a few minutes for the grant to apply and rerun."
    Write-MigLog WARN "The certificate's private key is non-exportable: run every step from this Windows account on this machine."
}

function Invoke-StepDiscover {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    Connect-GraphApp -Side $s.Source
    $groups = Get-GraphAll -Uri "/v1.0/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,mail"
    Write-MigLog INFO "Scanning $($groups.Count) Microsoft 365 groups for plans..."
    $rows = foreach ($g in $groups) {
        try { $plans = Get-GraphAll -Uri "/v1.0/groups/$($g.id)/planner/plans?`$select=id,title" }
        catch { Write-MigLog WARN "Group $($g.displayName) ($($g.id)): $($_.Exception.Message)"; continue }
        if ($plans.Count) { [pscustomobject]@{ sourceGroupId = $g.id; destGroupId = ''; sourceDisplayName = $g.displayName; sourceMail = $g.mail; planCount = $plans.Count } }
    }
    $target = Join-Path $cfgDir 'groups.csv'
    if (Test-Path $target) { $target = Join-Path $cfgDir 'groups.discovered.csv' }
    @($rows) | Export-Csv -Path $target -NoTypeInformation
    Write-MigLog OK "$(@($rows).Count) groups have plans -> $target. Fill in destGroupId (the GCC High group ID) for each row you want migrated."
}

function Invoke-StepExport {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    Connect-GraphApp -Side $s.Source
    $outRoot = Join-Path $s.ExportFolder 'basic'
    New-Item -ItemType Directory -Force -Path $outRoot | Out-Null
    $groupsCsv = Join-Path $cfgDir 'groups.csv'
    if (-not (Test-Path $groupsCsv)) { throw "Missing $groupsCsv. Run -Step Discover first." }
    $usersPath = Join-Path $outRoot 'users.json'
    $users = if (Test-Path $usersPath) { Get-Content $usersPath -Raw | ConvertFrom-Json -AsHashtable } else { @{} }
    function Resolve-SourceUser([string]$Id) {
        if ([string]::IsNullOrEmpty($Id) -or $users.ContainsKey($Id)) { return }
        try {
            $u = Invoke-Graph -Uri "/v1.0/users/$Id`?`$select=id,userPrincipalName,displayName,mail"
            $users[$Id] = @{ upn = $u.userPrincipalName; displayName = $u.displayName; mail = $u.mail }
        }
        catch { $users[$Id] = @{ upn = $null; displayName = $null; mail = $null }; Write-MigLog WARN "Source user $Id not resolvable (deleted?)" }
    }
    $inventoryPath = Join-Path $outRoot 'inventory.csv'
    foreach ($g in (Import-Csv $groupsCsv)) {
        $plans = Get-GraphAll -Uri "/v1.0/groups/$($g.sourceGroupId)/planner/plans"
        Write-MigLog INFO "Group $($g.sourceDisplayName) ($($g.sourceGroupId)): $($plans.Count) plan(s)"
        foreach ($plan in $plans) {
            $planDetails = Invoke-Graph -Uri "/v1.0/planner/plans/$($plan.id)/details"
            $buckets = Get-GraphAll -Uri "/v1.0/planner/plans/$($plan.id)/buckets"
            $tasks = Get-GraphAll -Uri "/v1.0/planner/plans/$($plan.id)/tasks"
            $commentPosts = 0; $threads = 0
            $taskRecords = foreach ($t in $tasks) {
                $td = Invoke-Graph -Uri "/v1.0/planner/tasks/$($t.id)/details"
                $posts = @()
                if (-not $SkipComments -and $t.conversationThreadId) {
                    $threads++
                    try { $posts = Get-GraphAll -Uri "/v1.0/groups/$($g.sourceGroupId)/threads/$($t.conversationThreadId)/posts" }
                    catch { Write-MigLog WARN "Comments for task '$($t.title)' not readable: $($_.Exception.Message)" }
                    $commentPosts += $posts.Count
                }
                if ($t.assignments) { $t.assignments.PSObject.Properties.Name | ForEach-Object { Resolve-SourceUser $_ } }
                [pscustomobject]@{ task = $t; details = $td; comments = @($posts) }
            }
            $record = [pscustomobject]@{ exportedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); sourceGroupId = $g.sourceGroupId
                                         plan = $plan; planDetails = $planDetails; buckets = @($buckets); tasks = @($taskRecords) }
            $dir = Join-Path $outRoot $g.sourceGroupId
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
            $record | ConvertTo-Json -Depth 40 | Set-Content -Path (Join-Path $dir "$($plan.id).json") -Encoding utf8
            $users | ConvertTo-Json -Depth 5 | Set-Content -Path $usersPath -Encoding utf8
            Add-CsvRow -Path $inventoryPath -Row ([pscustomobject]@{
                sourceGroupId = $g.sourceGroupId; groupName = $g.sourceDisplayName; planId = $plan.id; planTitle = $plan.title
                buckets = $buckets.Count; tasks = $tasks.Count
                checklistItems = ($tasks | Measure-Object -Property checklistItemCount -Sum).Sum
                references = ($tasks | Measure-Object -Property referenceCount -Sum).Sum
                tasksWithCommentThreads = $threads; commentPosts = $commentPosts; exportedAtUtc = $record.exportedAtUtc })
            Write-MigLog OK "Exported '$($plan.title)': $($buckets.Count) buckets, $($tasks.Count) tasks, $commentPosts comment posts"
        }
    }
    Write-MigLog OK "Basic export complete -> $outRoot. Next: -Step ExportTaskChat (comments made since the 2026 task-chat change)."
}

function Invoke-StepExportTaskChat {
    # Task chat replaced comments for basic plans in early 2026 (MC1193421). It is readable only through Graph BETA
    # 'List messages', Global cloud only, delegated only (Tasks.Read). So: sign in as a user who is a member of the
    # source groups, read chat, store it in the export. It is later appended to the description as text, because
    # 'Create message' is not available in US Gov L4.
    $s = Read-MigrationSettings -Path $SettingsPath
    $exportRoot = Join-Path $s.ExportFolder 'basic'
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }
    Connect-MgGraph -Environment $s.Source.GraphEnvironment -TenantId $s.Source.TenantId -Scopes 'Tasks.Read', 'User.ReadBasic.All' -NoWelcome -ContextScope Process
    Write-MigLog OK "Signed in (delegated) as $((Get-MgContext).Account) for task chat"
    $users = Get-Content (Join-Path $exportRoot 'users.json') -Raw | ConvertFrom-Json -AsHashtable
    $nameById = @{}; foreach ($k in $users.Keys) { if ($users[$k].displayName) { $nameById[$k] = $users[$k].displayName } }
    $problems = Join-Path $s.LogFolder 'taskchat-export-problems.csv'
    $files = Get-ChildItem -Path $exportRoot -Recurse -Filter '*.json' | Where-Object Name -ne 'users.json'
    $total = 0
    foreach ($f in $files) {
        $rec = Get-Content $f.FullName -Raw | ConvertFrom-Json
        if ($OnlyIds -and $rec.plan.id -notin $OnlyIds) { continue }
        foreach ($tr in $rec.tasks) {
            try { $msgs = Get-GraphAll -Uri "/beta/planner/tasks/$($tr.task.id)/messages" }
            catch {
                Add-CsvRow -Path $problems -Row ([pscustomobject]@{ plan = $rec.plan.title; task = $tr.task.title; issue = $_.Exception.Message })
                continue
            }
            foreach ($m in $msgs) {
                $uid = if ($m.createdBy -and $m.createdBy.user) { [string]$m.createdBy.user.id }
                if ($uid -and -not $nameById.ContainsKey($uid)) {
                    try { $nameById[$uid] = (Invoke-Graph -Uri "/v1.0/users/$uid`?`$select=displayName").displayName } catch { $nameById[$uid] = 'unknown' }
                }
            }
            $posts = @(ConvertTo-CommentPost -Messages $msgs -NameById $nameById)
            $tr | Add-Member -NotePropertyName chatPosts -NotePropertyValue $posts -Force
            $total += $posts.Count
        }
        $rec | Add-Member -NotePropertyName chatExportedAtUtc -NotePropertyValue (Get-Date).ToUniversalTime().ToString('o') -Force
        $rec | ConvertTo-Json -Depth 40 | Set-Content -Path $f.FullName -Encoding utf8
        Write-MigLog OK "Task chat added for '$($rec.plan.title)'"
    }
    Write-MigLog OK "Task chat export done: $total messages. Tasks the signed-in user couldn't read are in $problems."
}

function Invoke-StepImport([switch]$DryRun) {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    $exportRoot = Join-Path $s.ExportFolder 'basic'
    $userMap = Get-UserMap $cfgDir; $domainMap = Get-DomainMap $s
    $groupMap = @{}
    Import-Csv (Join-Path $cfgDir 'groups.csv') | Where-Object destGroupId | ForEach-Object { $groupMap[$_.sourceGroupId.Trim()] = $_.destGroupId.Trim() }
    $urlMapPath = Join-Path $cfgDir 'urlmap.csv'
    $urlMap = if (Test-Path $urlMapPath) { @(Import-Csv $urlMapPath | Where-Object sourcePrefix) } else { @() }
    $srcUsers = Get-Content (Join-Path $exportRoot 'users.json') -Raw | ConvertFrom-Json -AsHashtable
    $srcUpnById = @{}; foreach ($k in $srcUsers.Keys) { $srcUpnById[$k] = $srcUsers[$k].upn }

    Connect-GraphApp -Side $s.Destination
    $statePath = Join-Path $s.StateFolder 'basic-idmap.json'
    $idMap = Get-IdMap -Path $statePath
    $problemsCsv = Join-Path $s.LogFolder 'basic-import-problems.csv'
    function Write-Problem([string]$Plan, [string]$Item, [string]$Issue) {
        Add-CsvRow -Path $problemsCsv -Row ([pscustomobject]@{ timeUtc = (Get-Date).ToUniversalTime().ToString('o'); plan = $Plan; item = $Item; issue = $Issue })
    }
    $destIdByUpn = @{}
    $needed = $srcUpnById.Values | ForEach-Object { Resolve-DestUpn -SourceUpn $_ -UserMap $userMap -DomainMap $domainMap } | Where-Object { $_ } | Sort-Object -Unique
    foreach ($upn in $needed) {
        try { $destIdByUpn[$upn] = (Invoke-Graph -Uri "/v1.0/users/$upn`?`$select=id").id }
        catch { Write-MigLog WARN "Destination user $upn not found - their assignments will be skipped"; Write-Problem '' $upn 'destination user not found' }
    }
    Write-MigLog INFO "Resolved $($destIdByUpn.Count)/$(@($needed).Count) destination users"
    $memberCache = @{}
    function Get-DestMembers([string]$GroupId) {
        if (-not $memberCache.ContainsKey($GroupId)) {
            $set = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($m in (Get-GraphAll -Uri "/v1.0/groups/$GroupId/members?`$select=id")) { [void]$set.Add($m.id) }
            $memberCache[$GroupId] = $set
        }
        $memberCache[$GroupId]
    }

    foreach ($f in (Get-ChildItem -Path $exportRoot -Recurse -Filter '*.json' | Where-Object Name -ne 'users.json')) {
        $rec = Get-Content $f.FullName -Raw | ConvertFrom-Json
        $srcPlan = $rec.plan
        if ($OnlyIds -and $srcPlan.id -notin $OnlyIds) { continue }
        $destGroup = $groupMap[$rec.sourceGroupId]
        if (-not $destGroup) { Write-MigLog WARN "Plan '$($srcPlan.title)': no destGroupId in groups.csv - skipped"; Write-Problem $srcPlan.title $rec.sourceGroupId 'no destination group mapping'; continue }
        try { $members = Get-DestMembers $destGroup }
        catch { Write-MigLog ERROR "Destination group $destGroup not readable: $($_.Exception.Message)"; Write-Problem $srcPlan.title $destGroup 'destination group not readable'; continue }

        $pk = "plan:$($srcPlan.id)"
        if (-not $idMap[$pk]) {
            if ($DryRun) { Write-MigLog INFO "[DryRun] would create plan '$($srcPlan.title)' in group $destGroup" }
            else {
                $p = Invoke-Graph -Method POST -Uri '/v1.0/planner/plans' -Body @{ container = @{ containerId = $destGroup; type = 'group' }; title = $srcPlan.title }
                if (-not $p.id) { throw "Plan create for '$($srcPlan.title)' returned no id" }
                $idMap[$pk] = $p.id; Save-IdMap $idMap $statePath
                Write-MigLog OK "Created plan '$($srcPlan.title)' -> $($p.id)"
            }
        }
        $destPlanId = $idMap[$pk]

        $cats = @{}
        if ($rec.planDetails.categoryDescriptions) { foreach ($c in $rec.planDetails.categoryDescriptions.PSObject.Properties) { if ($c.Name -match '^category\d+$' -and $c.Value) { $cats[$c.Name] = $c.Value } } }
        if (-not $DryRun -and $cats.Count -and -not $idMap["plandetails:$($srcPlan.id)"]) {
            $d = Invoke-Graph -Uri "/v1.0/planner/plans/$destPlanId/details"
            Invoke-Graph -Method PATCH -Uri "/v1.0/planner/plans/$destPlanId/details" -Headers @{ 'If-Match' = $d.'@odata.etag' } -Body @{ categoryDescriptions = $cats } | Out-Null
            $idMap["plandetails:$($srcPlan.id)"] = 'done'; Save-IdMap $idMap $statePath
        }

        $prevHint = $null
        foreach ($b in (Get-SortedByOrderHint -Items $rec.buckets)) {
            $bk = "bucket:$($b.id)"
            if ($idMap[$bk]) { $prevHint = $idMap["buckethint:$($b.id)"]; continue }
            if ($DryRun) { continue }
            $nb = Invoke-Graph -Method POST -Uri '/v1.0/planner/buckets' -Body @{ name = $b.name; planId = $destPlanId; orderHint = (Get-AppendOrderHint $prevHint) }
            if (-not $nb.id) { throw "Bucket create for '$($b.name)' returned no id" }
            $idMap[$bk] = $nb.id; $idMap["buckethint:$($b.id)"] = $nb.orderHint; $prevHint = $nb.orderHint
            Save-IdMap $idMap $statePath
        }

        $created = 0; $failed = 0
        foreach ($tr in $rec.tasks) {
            $t = $tr.task
            $a = Get-PlannerAssignmentsBody -SourceAssignments $t.assignments -SourceUserUpn $srcUpnById -DestUserIdByUpn $destIdByUpn -UserMap $userMap -DomainMap $domainMap
            foreach ($u in $a.Unmapped) { Write-Problem $srcPlan.title $t.title "assignee not mapped: $u" }
            foreach ($destUserId in @($a.Body.Keys)) { if (-not $members.Contains($destUserId)) { Write-Problem $srcPlan.title $t.title "assignee $destUserId is not a member of destination group $destGroup" } }
            if ($DryRun) { continue }

            $tk = "task:$($t.id)"
            if (-not $idMap[$tk]) {
                $body = @{ planId = $destPlanId; title = $t.title }
                if ($t.bucketId -and $idMap["bucket:$($t.bucketId)"]) { $body.bucketId = $idMap["bucket:$($t.bucketId)"] }
                foreach ($field in 'startDateTime', 'dueDateTime') { if ($t.$field) { $body[$field] = $t.$field } }
                if ($null -ne $t.priority) { $body.priority = [int]$t.priority }
                if ($t.percentComplete) { $body.percentComplete = [int]$t.percentComplete }
                $applied = @{}
                if ($t.appliedCategories) { foreach ($c in $t.appliedCategories.PSObject.Properties) { if ($c.Value -eq $true) { $applied[$c.Name] = $true } } }
                if ($applied.Count) { $body.appliedCategories = $applied }
                if ($a.Body.Count) { $body.assignments = $a.Body }
                try {
                    $nt = Invoke-Graph -Method POST -Uri '/v1.0/planner/tasks' -Body $body
                    if (-not $nt.id) { throw 'task create returned no id' }
                    $idMap[$tk] = $nt.id; Save-IdMap $idMap $statePath; $created++
                }
                catch { $failed++; Write-MigLog ERROR "Task '$($t.title)': $($_.Exception.Message)"; Write-Problem $srcPlan.title $t.title "task create failed: $($_.Exception.Message)"; continue }
            }
            $destTaskId = $idMap[$tk]

            $dk = "taskdetails:$($t.id)"
            if ($idMap[$dk]) { continue }
            $comments = if ($s.CommentsMode -eq 'AppendToDescription') { Format-CommentsText -Posts (@($tr.comments) + @($tr.chatPosts)) } else { '' }
            $db = Get-TaskDetailsBody -SourceDetails $tr.details -UrlMap $urlMap -CommentsText $comments
            if ($db.Count) {
                try {
                    $d = Invoke-Graph -Uri "/v1.0/planner/tasks/$destTaskId/details"
                    Invoke-Graph -Method PATCH -Uri "/v1.0/planner/tasks/$destTaskId/details" -Headers @{ 'If-Match' = $d.'@odata.etag' } -Body $db | Out-Null
                }
                catch {
                    $code = $_.Exception.Data['StatusCode']
                    if (-not $comments -or $code -ne 400) {
                        # Transient/conflict (404/409/412/5xx): leave pending so the next run retries with a fresh ETag.
                        Write-Problem $srcPlan.title $t.title "details failed (HTTP $code), will retry on next run: $($_.Exception.Message)"
                        continue
                    }
                    # HTTP 400 with comments appended (most likely description length): retry without comments; keep them in a sidecar file.
                    $sidecarDir = Join-Path $s.ExportFolder 'comments-not-imported'
                    New-Item -ItemType Directory -Force -Path $sidecarDir | Out-Null
                    Set-Content -Path (Join-Path $sidecarDir "$($t.id).txt") -Value ("Plan: $($srcPlan.title)`nTask: $($t.title)" + $comments) -Encoding utf8
                    Write-Problem $srcPlan.title $t.title "details with comments rejected (HTTP 400); imported without comments; comments saved to comments-not-imported\$($t.id).txt"
                    try {
                        $db = Get-TaskDetailsBody -SourceDetails $tr.details -UrlMap $urlMap -CommentsText ''
                        $d = Invoke-Graph -Uri "/v1.0/planner/tasks/$destTaskId/details"
                        if ($db.Count) { Invoke-Graph -Method PATCH -Uri "/v1.0/planner/tasks/$destTaskId/details" -Headers @{ 'If-Match' = $d.'@odata.etag' } -Body $db | Out-Null }
                    }
                    catch { Write-Problem $srcPlan.title $t.title "details failed: $($_.Exception.Message)"; continue }
                }
            }
            $idMap[$dk] = 'done'; Save-IdMap $idMap $statePath
        }
        $mode = if ($DryRun) { '[DryRun] ' } else { '' }
        Write-MigLog OK "$mode'$($srcPlan.title)': $($rec.tasks.Count) source tasks, $created created this run, $failed failed"
    }
    Write-MigLog OK "Import pass finished. Review $problemsCsv, then run -Step Verify"
}

function Invoke-StepVerify {
    $s = Read-MigrationSettings -Path $SettingsPath
    $exportRoot = Join-Path $s.ExportFolder 'basic'
    $idMap = Get-IdMap -Path (Join-Path $s.StateFolder 'basic-idmap.json')
    Connect-GraphApp -Side $s.Destination
    $out = Join-Path $s.LogFolder 'verify-basic.csv'
    if (Test-Path $out) { Remove-Item $out }
    $rows = [System.Collections.Generic.List[object]]::new()
    function Add-Check($Plan, $Item, $Check, $Expected, $Actual) {
        $rows.Add([pscustomobject]@{ plan = $Plan; item = $Item; check = $Check; expected = $Expected; actual = $Actual; result = $(if ("$Expected" -eq "$Actual") { 'PASS' } else { 'FAIL' }) })
    }
    foreach ($f in (Get-ChildItem -Path $exportRoot -Recurse -Filter '*.json' | Where-Object Name -ne 'users.json')) {
        $rec = Get-Content $f.FullName -Raw | ConvertFrom-Json
        $destPlanId = $idMap["plan:$($rec.plan.id)"]
        if (-not $destPlanId) { Add-Check $rec.plan.title '(plan)' 'migrated' 'yes' 'no'; continue }
        $destBuckets = Get-GraphAll -Uri "/v1.0/planner/plans/$destPlanId/buckets"
        $destTasks = Get-GraphAll -Uri "/v1.0/planner/plans/$destPlanId/tasks"
        Add-Check $rec.plan.title '(plan)' 'bucket count' $rec.buckets.Count $destBuckets.Count
        Add-Check $rec.plan.title '(plan)' 'task count' $rec.tasks.Count $destTasks.Count
        $byId = [System.Collections.Hashtable]::new([System.StringComparer]::Ordinal)
        $destTasks | ForEach-Object { $byId[$_.id] = $_ }
        foreach ($tr in $rec.tasks) {
            $t = $tr.task
            $d = $byId[$idMap["task:$($t.id)"]]
            if (-not $d) { Add-Check $rec.plan.title $t.title 'task exists' 'yes' 'no'; continue }
            Add-Check $rec.plan.title $t.title 'title' $t.title $d.title
            Add-Check $rec.plan.title $t.title 'percentComplete' ([int]$t.percentComplete) ([int]$d.percentComplete)
            Add-Check $rec.plan.title $t.title 'checklistItemCount' ([int]$t.checklistItemCount) ([int]$d.checklistItemCount)
            Add-Check $rec.plan.title $t.title 'referenceCount' ([int]$t.referenceCount) ([int]$d.referenceCount)
            $srcA = if ($t.assignments) { @($t.assignments.PSObject.Properties).Count } else { 0 }
            $dstA = if ($d.assignments) { @($d.assignments.PSObject.Properties).Count } else { 0 }
            Add-Check $rec.plan.title $t.title 'assignee count' $srcA $dstA
        }
    }
    $rows | Export-Csv -Path $out -NoTypeInformation
    $fail = @($rows | Where-Object result -eq 'FAIL')
    Write-MigLog $(if ($fail.Count) { 'WARN' } else { 'OK' }) "Verification: $($rows.Count) checks, $($fail.Count) FAIL -> $out (assignee FAILs are expected where users were unmapped)"
}

function Invoke-StepPreflight {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    $results = [System.Collections.Generic.List[object]]::new()
    function Add-Result([string]$Area, [string]$Check, [string]$Status, [string]$Detail) {
        $results.Add([pscustomobject]@{ area = $Area; check = $Check; status = $Status; detail = $Detail })
        Write-MigLog (@{ PASS = 'OK'; WARN = 'WARN'; FAIL = 'ERROR'; INFO = 'INFO' }[$Status]) "[$Area] $Check : $Status $Detail"
    }
    function Test-Step([string]$Area, [string]$Check, [scriptblock]$Do) {
        try { $d = & $Do; Add-Result $Area $Check 'PASS' "$d"; return $true }
        catch { Add-Result $Area $Check 'FAIL' $_.Exception.Message; return $false }
    }
    Add-Result 'Local' 'PowerShell version' $(if ($PSVersionTable.PSVersion -ge [version]'7.4') { 'PASS' } else { 'FAIL' }) $PSVersionTable.PSVersion.ToString()
    foreach ($m in 'Microsoft.Graph.Authentication', 'Az.Accounts') {
        $mod = Get-Module -ListAvailable $m | Sort-Object Version -Descending | Select-Object -First 1
        if ($mod) { Add-Result 'Local' "module $m" 'PASS' $mod.Version.ToString() } else { Add-Result 'Local' "module $m" 'FAIL' "Install-Module $m -Scope CurrentUser" }
    }

    if ($Scope -in 'All', 'Basic') {
        $groupsCsv = Join-Path $cfgDir 'groups.csv'
        $groupRows = if (Test-Path $groupsCsv) { @(Import-Csv $groupsCsv) } else { @() }
        if (-not $groupRows) { Add-Result 'Basic' 'groups.csv' 'WARN' 'missing/empty - run -Step Discover first' }
        if (Test-Step 'Basic-Source' 'connect app-only (Global)' { Connect-GraphApp -Side $s.Source; (Get-MgContext).TenantId }) {
            $sampled = 0; $withThread = 0; $postsReadable = $null
            foreach ($g in $groupRows) {
                if ($sampled -ge $SampleTasks) { break }
                try {
                    foreach ($plan in (Get-GraphAll -Uri "/v1.0/groups/$($g.sourceGroupId)/planner/plans")) {
                        foreach ($t in (Get-GraphAll -Uri "/v1.0/planner/plans/$($plan.id)/tasks")) {
                            if ($sampled -ge $SampleTasks) { break }
                            $sampled++
                            if ($t.conversationThreadId) {
                                $withThread++
                                if ($null -eq $postsReadable) {
                                    try { $null = Get-GraphAll -Uri "/v1.0/groups/$($g.sourceGroupId)/threads/$($t.conversationThreadId)/posts"; $postsReadable = $true }
                                    catch { $postsReadable = $false; Add-Result 'Basic-Source' 'read comment posts' 'FAIL' $_.Exception.Message }
                                }
                            }
                        }
                    }
                }
                catch { Add-Result 'Basic-Source' "read plans in $($g.sourceDisplayName)" 'FAIL' $_.Exception.Message }
            }
            Add-Result 'Basic-Source' 'sampled tasks with a comment thread' 'INFO' "$withThread of $sampled"
            if ($postsReadable) { Add-Result 'Basic-Source' 'read comment posts' 'PASS' 'legacy comments readable' }
            Add-Result 'Basic-Source' 'task chat' 'INFO' 'Comments made after the 2026 task-chat change are only readable with -Step ExportTaskChat (delegated sign-in).'
        }
        if (Test-Step 'Basic-Dest' 'connect app-only (USGov)' { Connect-GraphApp -Side $s.Destination; (Get-MgContext).TenantId }) {
            $usersCsv = Join-Path $cfgDir 'users.csv'
            if (Test-Path $usersCsv) {
                $missing = @(foreach ($u in (Import-Csv $usersCsv | Where-Object destUpn)) { try { $null = Invoke-Graph -Uri "/v1.0/users/$($u.destUpn)`?`$select=id" } catch { $u } })
                if ($missing.Count) { $missing | Export-Csv (Join-Path $s.LogFolder 'preflight-missing-users.csv') -NoTypeInformation }
                Add-Result 'Basic-Dest' 'users.csv destUpn resolve' $(if ($missing.Count) { 'WARN' } else { 'PASS' }) "$($missing.Count) not found (logs\preflight-missing-users.csv)"
            }
            foreach ($g in ($groupRows | Where-Object destGroupId)) {
                Test-Step 'Basic-Dest' "group $($g.destGroupId)" { (Invoke-Graph -Uri "/v1.0/groups/$($g.destGroupId)?`$select=displayName").displayName } | Out-Null
            }
            $tg = $s.PreflightTestGroupId
            if (Test-Step 'Basic-Dest' 'create plan (Tasks.ReadWrite.All app-only)' {
                    $script:pfPlan = Invoke-Graph -Method POST -Uri '/v1.0/planner/plans' -Body @{ container = @{ containerId = $tg; type = 'group' }; title = "PREFLIGHT $(Get-Date -Format s)" }
                    $script:pfPlan.id }) {
                Test-Step 'Basic-Dest' 'create bucket' { $script:pfBucket = Invoke-Graph -Method POST -Uri '/v1.0/planner/buckets' -Body @{ name = 'Preflight'; planId = $script:pfPlan.id; orderHint = ' !' }; $script:pfBucket.id } | Out-Null
                Test-Step 'Basic-Dest' 'create task' { $script:pfTask = Invoke-Graph -Method POST -Uri '/v1.0/planner/tasks' -Body @{ planId = $script:pfPlan.id; bucketId = $script:pfBucket.id; title = 'Preflight task'; percentComplete = 50 }; $script:pfTask.id } | Out-Null
                Test-Step 'Basic-Dest' 'update task details (description, checklist, link)' {
                    $d = Invoke-Graph -Uri "/v1.0/planner/tasks/$($script:pfTask.id)/details"
                    $body = Get-TaskDetailsBody -SourceDetails ([pscustomobject]@{ description = 'Preflight'; previewType = 'automatic'
                            checklist = [pscustomobject]@{ x = [pscustomobject]@{ title = 'item'; isChecked = $true; orderHint = ' !' } }
                            references = [pscustomobject]@{ (ConvertTo-PlannerRefKey 'https://learn.microsoft.com/graph') = [pscustomobject]@{ alias = 'Learn'; type = 'Other' } } })
                    Invoke-Graph -Method PATCH -Uri "/v1.0/planner/tasks/$($script:pfTask.id)/details" -Headers @{ 'If-Match' = $d.'@odata.etag' } -Body $body | Out-Null
                    $back = Invoke-Graph -Uri "/v1.0/planner/tasks/$($script:pfTask.id)"
                    if ($back.checklistItemCount -ne 1 -or $back.referenceCount -ne 1) { throw "read-back mismatch: checklist=$($back.checklistItemCount) refs=$($back.referenceCount)" }
                    'checklist=1 refs=1'
                } | Out-Null
                Test-Step 'Basic-Dest' 'delete canary plan' {
                    $p = Invoke-Graph -Uri "/v1.0/planner/plans/$($script:pfPlan.id)"
                    Invoke-Graph -Method DELETE -Uri "/v1.0/planner/plans/$($script:pfPlan.id)" -Headers @{ 'If-Match' = $p.'@odata.etag' } | Out-Null
                    'deleted'
                } | Out-Null
            }
        }
    }

    if ($Scope -in 'All', 'Premium') {
        if (Test-Step 'Premium-Source' "sign in + WhoAmI ($($s.Source.AzEnvironment))" { (Connect-DataverseSide $s.Source).UserId }) {
            Test-Step 'Premium-Source' 'count premium plans' {
                $ti = Get-DvTableInfo -LogicalName 'msdyn_project'
                if (-not $ti) { throw 'msdyn_project table not found - wrong environment URL?' }
                $r = Invoke-Dv -Path "$($ti.EntitySetName)?`$select=$($ti.PrimaryId)&`$count=true" -ExtraHeaders @{ Prefer = 'odata.maxpagesize=1' }
                "$($r.'@odata.count') projects visible to this account"
            } | Out-Null
        }
        if (Test-Step 'Premium-Dest' "sign in + WhoAmI ($($s.Destination.AzEnvironment))" { (Connect-DataverseSide $s.Destination).UserId }) {
            $tiP = Get-DvTableInfo -LogicalName 'msdyn_project'
            if (-not $tiP) { Add-Result 'Premium-Dest' 'Project tables present' 'FAIL' 'msdyn_project not found: premium plans not provisioned in this environment' }
            else {
                Add-Result 'Premium-Dest' 'Project tables present' 'PASS' $tiP.EntitySetName
                if (Test-Step 'Premium-Dest' 'msdyn_CreateProjectV1 (needs Project license on this user)' {
                        $r = Invoke-Dv -Method POST -Path 'msdyn_CreateProjectV1' -Body @{ Project = @{ '@odata.type' = 'Microsoft.Dynamics.CRM.msdyn_project'; msdyn_subject = "PREFLIGHT $(Get-Date -Format s) - delete me" } }
                        $script:pfProj = $r.ProjectId; $script:pfProj }) {
                    $projBind = "/$($tiP.EntitySetName)($script:pfProj)"
                    Test-Step 'Premium-Dest' 'msdyn_CreateTeamMemberV1 (generic)' {
                        $r = Invoke-Dv -Method POST -Path 'msdyn_CreateTeamMemberV1' -Body @{ TeamMember = @{ '@odata.type' = 'Microsoft.Dynamics.CRM.msdyn_projectteam'; msdyn_projectteamid = [guid]::NewGuid().ToString(); 'msdyn_project@odata.bind' = $projBind; msdyn_name = 'Preflight member' } }
                        $script:pfTm = $r.TeamMemberId; $script:pfTm
                    } | Out-Null
                    Test-Step 'Premium-Dest' 'OperationSet: bucket + task + assignment, execute, read back' {
                        $os = (Invoke-Dv -Method POST -Path 'msdyn_CreateOperationSetV1' -Body @{ ProjectId = $script:pfProj; Description = 'preflight' }).OperationSetId
                        $bId = [guid]::NewGuid().ToString(); $tId = [guid]::NewGuid().ToString()
                        Invoke-Dv -Method POST -Path 'msdyn_PssCreateV1' -Body @{ OperationSetId = $os; Entity = @{ '@odata.type' = 'Microsoft.Dynamics.CRM.msdyn_projectbucket'; msdyn_projectbucketid = $bId; msdyn_name = 'Preflight bucket'; 'msdyn_project@odata.bind' = $projBind } } | Out-Null
                        Invoke-Dv -Method POST -Path 'msdyn_PssCreateV1' -Body @{ OperationSetId = $os; Entity = @{ '@odata.type' = 'Microsoft.Dynamics.CRM.msdyn_projecttask'; msdyn_projecttaskid = $tId; msdyn_subject = 'Preflight task'; 'msdyn_project@odata.bind' = $projBind; 'msdyn_projectbucket@odata.bind' = "/msdyn_projectbuckets($bId)"; msdyn_outlinelevel = 1; msdyn_scheduledstart = (Get-Date).ToUniversalTime().ToString('o'); msdyn_scheduledend = (Get-Date).AddDays(3).ToUniversalTime().ToString('o') } } | Out-Null
                        if ($script:pfTm) {
                            Invoke-Dv -Method POST -Path 'msdyn_PssCreateV1' -Body @{ OperationSetId = $os; Entity = @{ '@odata.type' = 'Microsoft.Dynamics.CRM.msdyn_resourceassignment'; msdyn_resourceassignmentid = [guid]::NewGuid().ToString(); msdyn_name = 'Preflight assignment'; 'msdyn_taskid@odata.bind' = "/msdyn_projecttasks($tId)"; 'msdyn_projectteamid@odata.bind' = "/msdyn_projectteams($script:pfTm)"; 'msdyn_projectid@odata.bind' = $projBind } } | Out-Null
                        }
                        Invoke-Dv -Method POST -Path 'msdyn_ExecuteOperationSetV1' -Body @{ OperationSetId = $os } | Out-Null
                        $deadline = (Get-Date).AddMinutes(5)
                        while ((Get-Date) -lt $deadline) {
                            try { Invoke-Dv -Path "msdyn_projecttasks($tId)?`$select=msdyn_projecttaskid" | Out-Null; return "task visible (operation set $os)" } catch { Start-Sleep -Seconds 10 }
                        }
                        throw "task not visible 5 min after execute (operation set $os) - check Settings > Schedule Integration > PSS Error Logs"
                    } | Out-Null
                    $first = if (Test-Path (Join-Path $cfgDir 'users.csv')) { Import-Csv (Join-Path $cfgDir 'users.csv') | Where-Object destUpn | Select-Object -First 1 }
                    if ($first) {
                        try {
                            $su = @((Invoke-Dv -Path "systemusers?`$select=systemuserid&`$filter=domainname eq '$($first.destUpn)'").value)[0]
                            $br = if ($su) { @((Invoke-Dv -Path "bookableresources?`$select=bookableresourceid&`$filter=_userid_value eq $($su.systemuserid)").value)[0] }
                            Add-Result 'Premium-Dest' "bookable resource for $($first.destUpn)" $(if ($br) { 'PASS' } else { 'WARN' }) $(if ($br) { 'found - -TryNamedResources can assign real users' } else { 'not found - imports use generic team members for this user' })
                        }
                        catch { Add-Result 'Premium-Dest' 'bookable resource lookup' 'WARN' $_.Exception.Message }
                    }
                    Add-Result 'Premium-Dest' 'cleanup' 'INFO' "Delete the 'PREFLIGHT ... - delete me' premium plan in Planner (GCC High) by hand."
                }
            }
        }
    }
    $csv = Join-Path $s.LogFolder 'preflight.csv'
    $results | Export-Csv -Path $csv -NoTypeInformation
    $results | Format-Table area, check, status, detail -AutoSize -Wrap | Out-Host
    $fails = @($results | Where-Object status -eq 'FAIL')
    if ($fails.Count) { Write-MigLog ERROR "$($fails.Count) FAIL(s). Fix them before migrating. Premium FAILs only block the premium path. -> $csv" }
    else { Write-MigLog OK "Preflight passed -> $csv" }
}

function Invoke-StepExportPremium {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    $outRoot = Join-Path $s.ExportFolder 'premium'
    New-Item -ItemType Directory -Force -Path $outRoot | Out-Null
    Connect-DataverseSide $s.Source | Out-Null
    $childTables = 'msdyn_projectbucket', 'msdyn_projecttask', 'msdyn_projectteam', 'msdyn_resourceassignment',
                   'msdyn_projecttaskdependency', 'msdyn_projectchecklist', 'msdyn_projectlabel', 'msdyn_projecttasktolabel'
    $info = @{}
    foreach ($t in @('msdyn_project') + $childTables) {
        $ti = Get-DvTableInfo -LogicalName $t
        if ($ti) { $info[$t] = $ti } else { Write-MigLog WARN "Table $t not present in source environment - skipped" }
    }
    if (-not $info['msdyn_project']) { throw "msdyn_project table not found in $($s.Source.DataverseUrl). Wrong environment URL?" }
    $info | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $outRoot '_source-metadata.json') -Encoding utf8
    $projects = Get-DvAll -Path $info['msdyn_project'].EntitySetName
    $filterPath = Join-Path $cfgDir 'premium-projects.csv'
    if ($OnlyIds) { $wanted = @($OnlyIds | ForEach-Object { $_.ToLowerInvariant() }) }
    elseif (Test-Path $filterPath) { $wanted = @(Import-Csv $filterPath | ForEach-Object { $_.sourceProjectId.Trim().ToLowerInvariant() }) }
    else { $wanted = $null }
    if ($wanted) { $projects = @($projects | Where-Object { ([string]$_.($info['msdyn_project'].PrimaryId)).ToLowerInvariant() -in $wanted }) }
    $taskPk = $info['msdyn_projecttask'].PrimaryId
    $teamPk = if ($info['msdyn_projectteam']) { $info['msdyn_projectteam'].PrimaryId }
    Write-MigLog INFO "Exporting $($projects.Count) premium plan(s)"
    function Get-ChildRows([string]$Table, [string]$ProjectId, [string[]]$TaskIds) {
        $ti = $info[$Table]; if (-not $ti) { return , @() }
        $projLookup = $ti.Lookups | Where-Object ReferencedEntity -eq 'msdyn_project' | Select-Object -First 1
        if ($projLookup) { return Get-DvAll -Path "$($ti.EntitySetName)?`$filter=_$($projLookup.ReferencingAttribute)_value eq $ProjectId" }
        $taskLookup = $ti.Lookups | Where-Object ReferencedEntity -eq 'msdyn_projecttask' | Select-Object -First 1
        if (-not $taskLookup -or -not $TaskIds) { Write-MigLog WARN "$Table has no project/task lookup (or no tasks) - skipped"; return , @() }
        $rows = @()
        foreach ($chunk in (Split-IntoChunks -Items $TaskIds -Size 50)) {
            $vals = ($chunk | ForEach-Object { "'$_'" }) -join ','
            $rows += Get-DvAll -Path "$($ti.EntitySetName)?`$filter=Microsoft.Dynamics.CRM.In(PropertyName='$($taskLookup.ReferencingAttribute)',PropertyValues=[$vals])"
        }
        return , $rows
    }
    $inventory = Join-Path $outRoot 'inventory.csv'; $labelsCsv = Join-Path $outRoot 'labels-manual.csv'
    foreach ($p in $projects) {
        $pid_ = $p.($info['msdyn_project'].PrimaryId)
        $tables = [ordered]@{}
        $tables['msdyn_projecttask'] = Get-ChildRows 'msdyn_projecttask' $pid_ @()
        $taskIds = @($tables['msdyn_projecttask'] | ForEach-Object { $_.$taskPk })
        foreach ($t in $childTables | Where-Object { $_ -ne 'msdyn_projecttask' }) { $tables[$t] = Get-ChildRows $t $pid_ $taskIds }
        $teamUpn = @{}
        $resLookup = if ($info['msdyn_projectteam']) { $info['msdyn_projectteam'].Lookups | Where-Object ReferencedEntity -eq 'bookableresource' | Select-Object -First 1 }
        foreach ($tm in $tables['msdyn_projectteam']) {
            if (-not $resLookup) { break }
            $resId = $tm."_$($resLookup.ReferencingAttribute)_value"
            if (-not $resId) { continue }
            try {
                $br = Invoke-Dv -Path "bookableresources($resId)?`$select=name,_userid_value"
                if ($br._userid_value) { $teamUpn[$tm.$teamPk] = (Invoke-Dv -Path "systemusers($($br._userid_value))?`$select=domainname").domainname }
            }
            catch { Write-MigLog WARN "Team member '$($tm.msdyn_name)': user lookup failed ($($_.Exception.Message)) - will import as generic member" }
        }
        [pscustomobject]@{ exportedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); sourceUrl = $s.Source.DataverseUrl; project = $p; tables = $tables; teamUpn = $teamUpn } |
            ConvertTo-Json -Depth 40 | Set-Content -Path (Join-Path $outRoot "$pid_.json") -Encoding utf8
        # Labels can't be created through the schedule API (Learn), so they are recorded for manual re-application.
        $lab = $info['msdyn_projectlabel']; $t2l = $info['msdyn_projecttasktolabel']
        if ($lab -and $t2l -and $tables['msdyn_projecttasktolabel'].Count) {
            $labelName = @{}; foreach ($l in $tables['msdyn_projectlabel']) { $labelName[$l.($lab.PrimaryId)] = $l.($lab.PrimaryName) }
            $taskName = @{}; foreach ($tk in $tables['msdyn_projecttask']) { $taskName[$tk.$taskPk] = $tk.msdyn_subject }
            $toLabel = $t2l.Lookups | Where-Object ReferencedEntity -eq 'msdyn_projectlabel' | Select-Object -First 1
            $toTask = $t2l.Lookups | Where-Object ReferencedEntity -eq 'msdyn_projecttask' | Select-Object -First 1
            foreach ($row in $tables['msdyn_projecttasktolabel']) {
                Add-CsvRow -Path $labelsCsv -Row ([pscustomobject]@{ project = $p.msdyn_subject
                    task = $taskName[$row."_$($toTask.ReferencingAttribute)_value"]; label = $labelName[$row."_$($toLabel.ReferencingAttribute)_value"] })
            }
        }
        $counts = [ordered]@{ projectId = $pid_; project = $p.msdyn_subject }
        foreach ($k in $tables.Keys) { $counts[$k] = @($tables[$k]).Count }
        Add-CsvRow -Path $inventory -Row ([pscustomobject]$counts)
        Write-MigLog OK "Exported premium plan '$($p.msdyn_subject)': $(@($tables['msdyn_projecttask']).Count) tasks"
    }
    Write-MigLog OK "Premium export complete -> $outRoot"
}

function Invoke-StepImportPremium {
    # MUST run as a destination USER with a Project license (Learn: app/system/integration users can't call the schedule APIs).
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    $exportRoot = Join-Path $s.ExportFolder 'premium'
    $userMap = Get-UserMap $cfgDir; $domainMap = Get-DomainMap $s
    Connect-DataverseSide $s.Destination | Out-Null
    $tableNames = 'msdyn_project', 'msdyn_projectbucket', 'msdyn_projecttask', 'msdyn_projectteam', 'msdyn_resourceassignment', 'msdyn_projecttaskdependency', 'msdyn_projectchecklist'
    $info = @{}; $setByEntity = @{}
    foreach ($t in $tableNames) {
        $ti = Get-DvTableInfo -LogicalName $t
        if ($ti) { $info[$t] = $ti; $setByEntity[$t] = $ti.EntitySetName } else { Write-MigLog WARN "Table $t not present in destination - its rows will be skipped" }
    }
    if (-not $info['msdyn_project']) { throw "msdyn_project not found in $($s.Destination.DataverseUrl). Premium plans not provisioned there (create one premium plan in Planner first) or wrong URL." }
    # Columns documented in Microsoft Learn's Project schedule API samples; extras only if creatable in destination metadata.
    $cols = @{
        msdyn_project = @(); msdyn_projectbucket = @('msdyn_name', 'msdyn_project')
        msdyn_projecttask = @('msdyn_subject', 'msdyn_project', 'msdyn_projectbucket', 'msdyn_parenttask', 'msdyn_outlinelevel', 'msdyn_scheduledstart', 'msdyn_scheduledend', 'msdyn_start', 'msdyn_effort')
        msdyn_resourceassignment = @('msdyn_name', 'msdyn_taskid', 'msdyn_projectteamid', 'msdyn_projectid')
        msdyn_projecttaskdependency = @('msdyn_project', 'msdyn_predecessortask', 'msdyn_successortask', 'msdyn_linktype')
    }
    if ($s.Premium -and $s.Premium.ExtraColumns) {
        foreach ($p in $s.Premium.ExtraColumns.PSObject.Properties) {
            foreach ($c in @($p.Value)) {
                if ($info[$p.Name] -and $c -in $info[$p.Name].CreatableColumns) { $cols[$p.Name] += $c; Write-MigLog INFO "Extra column $($p.Name).$c is creatable in destination - included" }
                else { Write-MigLog WARN "Extra column $($p.Name).$c is not creatable in destination - skipped" }
            }
        }
    }
    $chunkSize = [Math]::Min(200, [int]$(if ($s.Premium.OperationSetChunkSize) { $s.Premium.OperationSetChunkSize } else { 190 }))
    $waitMin = [int]$(if ($s.Premium.WaitForOperationSetMinutes) { $s.Premium.WaitForOperationSetMinutes } else { 15 })
    $statePath = Join-Path $s.StateFolder 'premium-idmap.json'
    $idMap = Get-IdMap -Path $statePath
    $problemsCsv = Join-Path $s.LogFolder 'premium-import-problems.csv'
    function Write-Problem([string]$Project, [string]$Item, [string]$Issue) {
        Add-CsvRow -Path $problemsCsv -Row ([pscustomobject]@{ timeUtc = (Get-Date).ToUniversalTime().ToString('o'); project = $Project; item = $Item; issue = $Issue })
    }
    function Test-DvRowExists([string]$Table, [string]$Id) {
        try { Invoke-Dv -Path "$($info[$Table].EntitySetName)($Id)?`$select=$($info[$Table].PrimaryId)" | Out-Null; return $true }
        catch { if ($_.Exception.Data['StatusCode'] -eq 404) { return $false }; throw }
    }
    function Get-NavProp([string]$Table, [string]$Attribute) {
        $l = $info[$Table].Lookups | Where-Object ReferencingAttribute -eq $Attribute | Select-Object -First 1
        if ($l) { $l.ReferencingEntityNavigationPropertyName } else { $Attribute }
    }
    function Invoke-OperationSets([string]$DestProjectId, [string]$Label, [object[]]$Items, [hashtable]$Planned) {
        foreach ($chunk in (Split-IntoChunks -Items $Items -Size $chunkSize)) {
            $osId = (Invoke-Dv -Method POST -Path 'msdyn_CreateOperationSetV1' -Body @{ ProjectId = $DestProjectId; Description = "GCC High migration $(Get-Date -Format s)" }).OperationSetId
            $queued = [System.Collections.Generic.List[object]]::new()
            foreach ($it in $chunk) {
                $ti = $info[$it.Table]
                $srcId = [string]$it.Row.($ti.PrimaryId)
                $pl = New-DvEntityPayload -Row $it.Row -TableInfo $ti -SetNameByEntity $setByEntity -IdMap $Planned -Columns $cols[$it.Table] -AllCreatable:([bool]$it.AllCreatable)
                # Mirror the documented sample: top-level tasks get msdyn_outlinelevel; child tasks get msdyn_parenttask instead.
                if ($it.Table -eq 'msdyn_projecttask' -and $pl.Payload.Contains("$(Get-NavProp 'msdyn_projecttask' 'msdyn_parenttask')@odata.bind")) { $pl.Payload.Remove('msdyn_outlinelevel') }
                foreach ($d in $pl.Dropped) { Write-Problem $Label "$($it.Table) $srcId" "not carried: $d" }
                try {
                    Invoke-Dv -Method POST -Path 'msdyn_PssCreateV1' -Body @{ Entity = $pl.Payload; OperationSetId = $osId } | Out-Null
                    $queued.Add([pscustomobject]@{ Table = $it.Table; Src = $srcId; Dest = [string]$pl.Payload[$ti.PrimaryId] })
                }
                catch { Write-Problem $Label "$($it.Table) $srcId" "PssCreateV1 rejected: $($_.Exception.Message)" }
            }
            if (-not $queued.Count) { try { Invoke-Dv -Method POST -Path 'msdyn_AbandonOperationSetV1' -Body @{ OperationSetId = $osId } | Out-Null } catch { }; continue }
            Invoke-Dv -Method POST -Path 'msdyn_ExecuteOperationSetV1' -Body @{ OperationSetId = $osId } | Out-Null
            # Confirm by reading back (Learn: OperationSet failure status/logs aren't available through the API).
            $deadline = (Get-Date).AddMinutes($waitMin)
            while ($queued.Count -and (Get-Date) -lt $deadline) {
                foreach ($q in @($queued)) { if (Test-DvRowExists $q.Table $q.Dest) { $idMap[$q.Src] = $q.Dest; [void]$queued.Remove($q) } }
                Save-IdMap $idMap $statePath
                if ($queued.Count) { Start-Sleep -Seconds 10 }
            }
            foreach ($q in $queued) { Write-Problem $Label "$($q.Table) $($q.Src)" "not present $waitMin min after ExecuteOperationSetV1 (set $osId). Check destination Settings > Schedule Integration > Operation Sets / PSS Error Logs" }
        }
    }

    foreach ($f in (Get-ChildItem -Path $exportRoot -Filter '*.json' | Where-Object { $_.Name -notlike '_*' })) {
        $rec = Get-Content $f.FullName -Raw | ConvertFrom-Json
        $projTi = $info['msdyn_project']
        $srcProjId = [string]$rec.project.($projTi.PrimaryId)
        if ($OnlyIds -and $srcProjId -notin $OnlyIds) { continue }
        $label = [string]$rec.project.msdyn_subject
        if (-not $idMap[$srcProjId]) {
            if (Test-DvRowExists 'msdyn_project' $srcProjId) { $idMap[$srcProjId] = $srcProjId }
            else {
                $pp = [ordered]@{ '@odata.type' = 'Microsoft.Dynamics.CRM.msdyn_project'; $projTi.PrimaryId = $srcProjId; msdyn_subject = $label }
                foreach ($c in $cols['msdyn_project']) { if ($null -ne $rec.project.$c) { $pp[$c] = $rec.project.$c } }
                try { $r = Invoke-Dv -Method POST -Path 'msdyn_CreateProjectV1' -Body @{ Project = $pp } }
                catch {
                    Write-MigLog WARN "CreateProjectV1 with source ID failed ($($_.Exception.Message)); retrying with a new ID"
                    $pp.Remove($projTi.PrimaryId)
                    $r = Invoke-Dv -Method POST -Path 'msdyn_CreateProjectV1' -Body @{ Project = $pp }
                }
                if (-not $r.ProjectId) { throw "msdyn_CreateProjectV1 for '$label' returned no ProjectId" }
                $idMap[$srcProjId] = [string]$r.ProjectId
                Write-MigLog OK "Created premium plan '$label' -> $($r.ProjectId)"
            }
            Save-IdMap $idMap $statePath
        }
        $destProjId = $idMap[$srcProjId]

        $teamTi = $info['msdyn_projectteam']
        if ($teamTi) {
            foreach ($tm in @($rec.tables.msdyn_projectteam)) {
                if (-not $tm) { continue }
                $src = [string]$tm.($teamTi.PrimaryId)
                if ($idMap[$src]) { continue }
                $projNav = Get-NavProp 'msdyn_projectteam' 'msdyn_project'
                $payload = [ordered]@{ '@odata.type' = 'Microsoft.Dynamics.CRM.msdyn_projectteam'; $teamTi.PrimaryId = $src; "$projNav@odata.bind" = "/$($projTi.EntitySetName)($destProjId)"; msdyn_name = $tm.msdyn_name }
                $named = $false
                $srcUpn = if ($rec.teamUpn) { $rec.teamUpn.$src }
                if ($TryNamedResources -and $srcUpn) {
                    $destUpn = Resolve-DestUpn -SourceUpn $srcUpn -UserMap $userMap -DomainMap $domainMap
                    $br = $null
                    try {
                        $su = @((Invoke-Dv -Path "systemusers?`$select=systemuserid&`$filter=domainname eq '$destUpn'").value)[0]
                        if ($su) { $br = @((Invoke-Dv -Path "bookableresources?`$select=bookableresourceid&`$filter=_userid_value eq $($su.systemuserid)").value)[0] }
                    } catch { Write-Problem $label $tm.msdyn_name "named resource lookup failed: $($_.Exception.Message)" }
                    $resLookup = $teamTi.Lookups | Where-Object ReferencedEntity -eq 'bookableresource' | Select-Object -First 1
                    if ($br -and $resLookup) { $payload["$($resLookup.ReferencingEntityNavigationPropertyName)@odata.bind"] = "/bookableresources($($br.bookableresourceid))"; $named = $true }
                    else { Write-Problem $label $tm.msdyn_name "no destination bookable resource for '$destUpn' - created as generic member" }
                }
                try {
                    $r = Invoke-Dv -Method POST -Path 'msdyn_CreateTeamMemberV1' -Body @{ TeamMember = $payload }
                    if (-not $r.TeamMemberId) { throw 'msdyn_CreateTeamMemberV1 returned no TeamMemberId' }
                    $idMap[$src] = [string]$r.TeamMemberId
                }
                catch {
                    if (-not $named) { Write-Problem $label $tm.msdyn_name "team member create failed: $($_.Exception.Message)"; continue }
                    Write-Problem $label $tm.msdyn_name "named member failed ($($_.Exception.Message)) - retried as generic"
                    @($payload.Keys | Where-Object { $_ -like '*@odata.bind' -and $_ -notlike "$projNav@*" }) | ForEach-Object { $payload.Remove($_) }
                    try { $r = Invoke-Dv -Method POST -Path 'msdyn_CreateTeamMemberV1' -Body @{ TeamMember = $payload }; $idMap[$src] = [string]$r.TeamMemberId }
                    catch { Write-Problem $label $tm.msdyn_name "team member create failed: $($_.Exception.Message)"; continue }
                }
                Save-IdMap $idMap $statePath
            }
        }

        # Planned map: everything already created + everything this run will attempt (source ID reused as destination ID).
        $planned = [System.Collections.Hashtable]::new($idMap, [System.StringComparer]::Ordinal)
        $planned[$srcProjId] = $destProjId
        foreach ($tn in 'msdyn_projectbucket', 'msdyn_projecttask', 'msdyn_resourceassignment', 'msdyn_projecttaskdependency', 'msdyn_projectchecklist') {
            if (-not $info[$tn]) { continue }
            foreach ($row in @($rec.tables.$tn)) { if ($row) { $id = [string]$row.($info[$tn].PrimaryId); if (-not $planned.ContainsKey($id)) { $planned[$id] = $id } } }
        }
        function Get-Pending([string]$Table) {
            if (-not $info[$Table]) { return }
            @($rec.tables.$Table) | Where-Object { $_ -and -not $idMap[[string]$_.($info[$Table].PrimaryId)] }
        }
        $taskPkD = $info['msdyn_projecttask'].PrimaryId
        $parentOf = @{}; foreach ($t in @($rec.tables.msdyn_projecttask)) { if ($t) { $parentOf[[string]$t.$taskPkD] = [string]$t._msdyn_parenttask_value } }
        $depth = @{}
        function Get-Depth([string]$Id) {
            if ($depth.ContainsKey($Id)) { return $depth[$Id] }
            $d = 0; $cur = $Id; $guard = 0
            while ($parentOf[$cur] -and $guard -lt 50) { $d++; $cur = $parentOf[$cur]; $guard++ }
            $depth[$Id] = $d; return $d
        }
        $tasks = @(@(Get-Pending 'msdyn_projecttask') | Sort-Object { Get-Depth ([string]$_.$taskPkD) }, { $_.msdyn_displaysequence })
        $items = @()
        $items += @(Get-Pending 'msdyn_projectbucket' | ForEach-Object { @{ Table = 'msdyn_projectbucket'; Row = $_ } })
        $items += @($tasks | ForEach-Object { @{ Table = 'msdyn_projecttask'; Row = $_ } })
        if ($items.Count) { Invoke-OperationSets $destProjId $label $items $planned }
        $items = @()
        $items += @(Get-Pending 'msdyn_resourceassignment' | ForEach-Object { @{ Table = 'msdyn_resourceassignment'; Row = $_ } })
        $items += @(Get-Pending 'msdyn_projecttaskdependency' | ForEach-Object { @{ Table = 'msdyn_projecttaskdependency'; Row = $_ } })
        if ($items.Count) { Invoke-OperationSets $destProjId $label $items $planned }
        $items = @(Get-Pending 'msdyn_projectchecklist' | ForEach-Object { @{ Table = 'msdyn_projectchecklist'; Row = $_; AllCreatable = $true } })
        if ($items.Count) { Invoke-OperationSets $destProjId $label $items $planned }

        $total = 0; $done = 0
        foreach ($tn in 'msdyn_projectbucket', 'msdyn_projecttask', 'msdyn_projectteam', 'msdyn_resourceassignment', 'msdyn_projecttaskdependency', 'msdyn_projectchecklist') {
            if (-not $info[$tn]) { continue }
            foreach ($row in @($rec.tables.$tn)) { if ($row) { $total++; if ($idMap[[string]$row.($info[$tn].PrimaryId)]) { $done++ } } }
        }
        Write-MigLog $(if ($done -eq $total) { 'OK' } else { 'WARN' }) "Premium plan '$label': $done/$total rows present in destination"
    }
    Write-MigLog OK "Premium import pass finished. Review $problemsCsv. Then share each plan with its group in Planner and re-apply labels from export\premium\labels-manual.csv."
}

#endregion

# Dot-sourced (tests): load functions only.
if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $Step) { Get-Help $PSCommandPath -Detailed; return }

switch ($Step) {
    'Init' { Invoke-StepInit }
    'Setup' { Invoke-StepSetup }
    'Discover' { Invoke-StepDiscover }
    'Preflight' { Invoke-StepPreflight }
    'Export' { Invoke-StepExport }
    'ExportTaskChat' { Invoke-StepExportTaskChat }
    'DryRun' { Invoke-StepImport -DryRun }
    'Import' { Invoke-StepImport }
    'Verify' { Invoke-StepVerify }
    'ExportPremium' { Invoke-StepExportPremium }
    'ImportPremium' { Invoke-StepImportPremium }
}

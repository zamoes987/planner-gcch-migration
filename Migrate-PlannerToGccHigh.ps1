<#
.SYNOPSIS
  Migrates Microsoft Planner (basic plans) and Planner premium plans (Project Plan 3 / formerly Project for the web)
  from a Microsoft 365 Commercial tenant to a Microsoft 365 GCC High tenant. Two commands:

    .\Migrate-PlannerToGccHigh.ps1 -Step Export     # run once, signed in to the COMMERCIAL tenant
    .\Migrate-PlannerToGccHigh.ps1 -Step Import     # run once, signed in to the GCC HIGH tenant

.DESCRIPTION
  Export (commercial):
    - first run only: registers the migration app + certificate (browser sign-in as a commercial admin)
    - finds every Microsoft 365 group that has plans and exports plans, buckets, tasks, checklists, links,
      labels, assignments, comments, group owners/members
    - exports task chat (browser sign-in as a user who is a member of the groups)
    - exports premium plans from every Dataverse environment that has them (browser sign-in)
  Import (GCC High):
    - first run only: registers the migration app + certificate (browser sign-in as a GCC High admin)
    - maps users automatically (same alias on the GCC High domains), maps groups by alias/name and CREATES
      missing groups with the mapped owners and members
    - preflight test (throwaway plan), import, verification report, premium import
  Everything is logged to logs\ and is safe to rerun: it resumes and never duplicates.

  Optional overrides (only if the automatic mapping gets something wrong; see logs\user-mapping.csv):
    config\users.csv   sourceUpn,destUpn
    config\groups.csv  sourceGroupId,destGroupId
    config\urlmap.csv  sourcePrefix,destPrefix

  Advanced single steps: Setup -Side Source|Destination, Discover, Preflight, ExportBasic, ExportTaskChat,
  ExportPremium, MapTargets, DryRun, ImportBasic, Verify, ImportPremium, Init.
  Facts this relies on, with Microsoft sources, are in README.md.

.EXAMPLE
  .\Migrate-PlannerToGccHigh.ps1 -Step Export
  .\Migrate-PlannerToGccHigh.ps1 -Step Import
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateSet('Export', 'Import', 'Setup', 'Discover', 'Preflight', 'ExportBasic', 'ExportTaskChat', 'ExportPremium',
                 'MapTargets', 'DryRun', 'ImportBasic', 'Verify', 'ImportPremium', 'Init')]
    [string]$Step,
    [ValidateSet('Source', 'Destination')][string]$Side,
    [ValidateSet('All', 'Basic', 'Premium')][string]$Scope = 'All',
    [string]$SettingsPath = (Join-Path $PSScriptRoot 'config\settings.json'),
    [string[]]$OnlyIds,
    [switch]$TryNamedResources,
    [switch]$SkipComments,
    [switch]$NoModuleInstall,
    [int]$SampleTasks = 200,
    [int]$CertMonths = 3
)
# No StrictMode: Graph/Dataverse omit optional properties (e.g. @odata.nextLink) and strict mode would throw on them.
$ErrorActionPreference = 'Stop'

#region ==================== settings, logging, state ====================

function Get-DefaultSettings {
    [ordered]@{
        Source      = [ordered]@{ TenantId = ''; ClientId = ''; CertificateThumbprint = ''; GraphEnvironment = 'Global'; AzEnvironment = 'AzureCloud'
                                  DataverseUrl = ''; DataverseAuth = 'Az'; DataverseClientId = '' }
        Destination = [ordered]@{ TenantId = ''; ClientId = ''; CertificateThumbprint = ''; GraphEnvironment = 'USGov'; AzEnvironment = 'AzureUSGovernment'
                                  DataverseUrl = ''; DataverseAuth = 'Az'; DataverseClientId = ''; FallbackOwnerUpn = '' }
        DomainMap = [ordered]@{}
        CreateMissingGroups = $true
        AddMissingAssigneesToGroups = $true
        PreflightTestGroupId = ''
        CommentsMode = 'AppendToDescription'
        NewGroupRetrySeconds = 30
        ExportFolder = 'export'; StateFolder = 'state'; LogFolder = 'logs'
        Premium = [ordered]@{ ExtraColumns = [ordered]@{ msdyn_project = @(); msdyn_projecttask = @('msdyn_description', 'msdyn_progress', 'msdyn_priority') }
                              OperationSetChunkSize = 190; WaitForOperationSetMinutes = 15 }
    }
}

function Invoke-StepInit {
    $cfgDir = Split-Path $SettingsPath -Parent
    New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null
    if (-not (Test-Path $SettingsPath)) {
        Get-DefaultSettings | ConvertTo-Json -Depth 6 | Set-Content -Path $SettingsPath -Encoding utf8
        Write-Host "Created $SettingsPath (defaults; nothing needs editing for a normal run)" -ForegroundColor Green
    }
}

function Read-MigrationSettings {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { Invoke-StepInit }
    $s = Get-Content $Path -Raw | ConvertFrom-Json
    # Fill in any setting missing from an older settings file.
    $def = Get-DefaultSettings | ConvertTo-Json -Depth 6 | ConvertFrom-Json
    foreach ($p in $def.PSObject.Properties) {
        if ($s.PSObject.Properties.Name -notcontains $p.Name) { $s | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value }
        elseif ($p.Name -in 'Source', 'Destination') {
            foreach ($q in $p.Value.PSObject.Properties) { if ($s.($p.Name).PSObject.Properties.Name -notcontains $q.Name) { $s.($p.Name) | Add-Member -NotePropertyName $q.Name -NotePropertyValue $q.Value } }
        }
    }
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
    if (Test-Path $Path) { (Get-Content $Path -Raw | ConvertFrom-Json -AsHashtable).GetEnumerator() | ForEach-Object { $map[$_.Key] = $_.Value } }
    return $map
}

function Save-IdMap {
    param([Parameter(Mandatory)][hashtable]$Map, [Parameter(Mandatory)][string]$Path)
    $tmp = "$Path.tmp"
    $Map | ConvertTo-Json -Depth 6 | Set-Content -Path $tmp -Encoding utf8
    Move-Item -Path $tmp -Destination $Path -Force   # write-then-rename: a crash never leaves a half-written file
}

function Get-UserMap([string]$ConfigDir) {
    $m = @{}
    $p = Join-Path $ConfigDir 'users.csv'
    if (Test-Path $p) { Import-Csv $p | Where-Object { $_.sourceUpn -and $_.destUpn } | ForEach-Object { $m[$_.sourceUpn.Trim().ToLowerInvariant()] = $_.destUpn.Trim().ToLowerInvariant() } }
    $m
}

function Get-PlanFiles([string]$ExportRoot) {
    Get-ChildItem -Path $ExportRoot -Recurse -Filter '*.json' | Where-Object { $_.Name -ne 'users.json' -and $_.Name -notlike '_*' }
}

function Initialize-Prerequisites([string[]]$Need) {
    # Installs missing modules for the current user (no admin rights needed) unless -NoModuleInstall.
    $probe = @{ Graph = @('Connect-MgGraph', 'Microsoft.Graph.Authentication'); Az = @('Connect-AzAccount', 'Az.Accounts') }
    foreach ($n in $Need) {
        $cmd, $mod = $probe[$n]
        if (Get-Command $cmd -ErrorAction SilentlyContinue) { continue }
        if ($NoModuleInstall) { throw "PowerShell module $mod is missing: Install-Module $mod -Scope CurrentUser" }
        Write-MigLog INFO "Installing PowerShell module $mod for the current user..."
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
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
    param([Parameter(Mandatory)][string]$Url, [object[]]$UrlMap = @())
    $hit = $UrlMap | Where-Object { $_.sourcePrefix -and $Url.StartsWith($_.sourcePrefix, [StringComparison]::OrdinalIgnoreCase) } |
        Sort-Object { $_.sourcePrefix.Length } -Descending | Select-Object -First 1
    if ($hit) { return $hit.destPrefix + $Url.Substring($hit.sourcePrefix.Length) }
    return $Url
}

function Get-AutoDomainMap {
    # Every verified source domain -> the destination default domain (Graph organization.verifiedDomains).
    param([object[]]$SourceDomains, [object[]]$DestDomains)
    $m = @{}
    $def = ($DestDomains | Where-Object { $_.isDefault } | Select-Object -First 1).name
    if (-not $def) { return $m }
    foreach ($d in $SourceDomains) { $n = ([string]$d.name).ToLowerInvariant(); if ($n -and -not $m.ContainsKey($n)) { $m[$n] = $def.ToLowerInvariant() } }
    $m
}

function Get-SharePointUrlMap {
    # BEST GUESS for link rewriting: SharePoint/OneDrive hosts derived from each tenant's initial domain
    # (<name>.onmicrosoft.com -> https://<name>.sharepoint.com/, <name>.onmicrosoft.us -> https://<name>.sharepoint.us/).
    # Logged so it can be checked; config\urlmap.csv overrides it.
    param([object[]]$SourceDomains, [object[]]$DestDomains)
    $si = [string](($SourceDomains | Where-Object { $_.isInitial } | Select-Object -First 1).name)
    $di = [string](($DestDomains | Where-Object { $_.isInitial } | Select-Object -First 1).name)
    if ($si -notmatch '^([a-z0-9-]+)\.onmicrosoft\.(com|us)$') { return @() }
    $sx = $Matches[1]; $st = $Matches[2]
    if ($di -notmatch '^([a-z0-9-]+)\.onmicrosoft\.(com|us)$') { return @() }
    $dx = $Matches[1]; $dt = $Matches[2]
    @([pscustomobject]@{ sourcePrefix = "https://$sx.sharepoint.$st/"; destPrefix = "https://$dx.sharepoint.$dt/" },
      [pscustomobject]@{ sourcePrefix = "https://$sx-my.sharepoint.$st/"; destPrefix = "https://$dx-my.sharepoint.$dt/" })
}

function Get-UserMatchCandidates {
    # Ordered destination UPN guesses for one source user: users.csv, DomainMap, same alias on each destination domain,
    # mail alias on each destination domain. Each returns @{ Upn; Method }.
    param($SourceUser, [string[]]$DestDomains = @(), [hashtable]$UserMap = @{}, [hashtable]$DomainMap = @{})
    $out = [System.Collections.Generic.List[object]]::new(); $seen = @{}
    function Add-C([string]$u, [string]$m) { if ($u -and -not $seen.ContainsKey($u)) { $seen[$u] = 1; $out.Add(@{ Upn = $u; Method = $m }) } }
    $upn = ([string]$SourceUser.upn).ToLowerInvariant()
    if ($upn) {
        if ($UserMap.ContainsKey($upn)) { return , @(@{ Upn = $UserMap[$upn]; Method = 'users.csv' }) }
        $local, $dom = $upn.Split('@', 2)
        if ($dom -and $DomainMap.ContainsKey($dom)) { Add-C "$local@$($DomainMap[$dom])" 'domain map' }
        foreach ($d in $DestDomains) { Add-C "$local@$($d.ToLowerInvariant())" "same alias @$d" }
    }
    $mail = ([string]$SourceUser.mail).ToLowerInvariant()
    if ($mail -match '^([^@]+)@') { $ml = $Matches[1]; foreach ($d in $DestDomains) { Add-C "$ml@$($d.ToLowerInvariant())" "mail alias @$d" } }
    , $out.ToArray()
}

function Get-SortedByOrderHint {
    param([object[]]$Items, [string]$Property = 'orderHint')
    if (-not $Items) { return @() }
    $list = [System.Collections.Generic.List[object]]::new()
    $Items | ForEach-Object { $list.Add($_) }
    $list.Sort([System.Comparison[object]] { param($x, $y) [string]::CompareOrdinal([string]$x.$Property, [string]$y.$Property) })
    return , $list.ToArray()
}

function Get-AppendOrderHint {
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
    # Task chat messages (Graph beta plannerTaskChatMessage) -> group-post shape, so one formatter handles both.
    param([object[]]$Messages, [hashtable]$NameById = @{})
    foreach ($m in $Messages) {
        if (-not $m -or $m.deletedTime) { continue }
        $uid = if ($m.createdBy -and $m.createdBy.user) { [string]$m.createdBy.user.id } else { '' }
        $name = if ($uid -and $NameById[$uid]) { $NameById[$uid] } else { 'unknown' }
        [pscustomobject]@{ createdDateTime = $m.createdDateTime; from = [pscustomobject]@{ emailAddress = [pscustomobject]@{ name = $name } }; body = [pscustomobject]@{ content = $m.content } }
    }
}

function Format-CommentsText {
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
    param($Headers, [int]$Attempt)
    $ra = $null
    if ($Headers -is [System.Collections.IDictionary]) { foreach ($k in @($Headers.Keys)) { if ($k -ieq 'Retry-After') { $ra = @($Headers[$k])[0] } } }
    elseif ($Headers) { foreach ($kv in $Headers) { if ($kv.Key -ieq 'Retry-After') { $ra = @($kv.Value)[0] } } }
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
    # Source assignee IDs -> destination IDs (from the user mapping). Returns @{ Body; Unmapped }.
    param($SourceAssignments, [hashtable]$DestIdBySourceId = @{}, [hashtable]$SourceUpnById = @{}, [string[]]$OnlyDestIds)
    $body = @{}; $unmapped = @()
    if ($SourceAssignments) {
        foreach ($p in $SourceAssignments.PSObject.Properties) {
            $destId = $DestIdBySourceId[$p.Name]
            if ($destId -and (-not $PSBoundParameters.ContainsKey('OnlyDestIds') -or $destId -in $OnlyDestIds)) {
                $body[$destId] = @{ '@odata.type' = '#microsoft.graph.plannerAssignment'; orderHint = ' !' }
            }
            elseif (-not $destId) { $unmapped += $(if ($SourceUpnById[$p.Name]) { $SourceUpnById[$p.Name] } else { "id:$($p.Name)" }) }
        }
    }
    @{ Body = $body; Unmapped = $unmapped }
}

function Get-TaskDetailsBody {
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
    param([Parameter(Mandatory)][string]$AzEnvironment)
    switch ($AzEnvironment) {
        'AzureCloud' { 'https://login.microsoftonline.com' }
        'AzureUSGovernment' { 'https://login.microsoftonline.us' }
        default { throw "Unsupported AzEnvironment '$AzEnvironment' (use AzureCloud or AzureUSGovernment)" }
    }
}

function Get-GraphBaseUrl([string]$GraphEnvironment) {
    # Microsoft Graph national cloud deployments.
    switch ($GraphEnvironment) { 'Global' { 'https://graph.microsoft.com' } 'USGov' { 'https://graph.microsoft.us' } default { throw "Unsupported GraphEnvironment '$GraphEnvironment'" } }
}

function Get-GlobalDiscoUrl([string]$AzEnvironment) {
    # Dataverse Global Discovery Service per cloud (Learn: Discover user organizations).
    switch ($AzEnvironment) { 'AzureCloud' { 'https://globaldisco.crm.dynamics.com' } 'AzureUSGovernment' { 'https://globaldisco.crm.microsoftdynamics.us' } default { throw "Unsupported AzEnvironment '$AzEnvironment'" } }
}

function ConvertTo-FormBody([hashtable]$Fields) {
    ($Fields.GetEnumerator() | Sort-Object Key | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join '&'
}

function ConvertTo-ODataString([string]$Value) { $Value.Replace("'", "''") }

#endregion

#region ==================== Microsoft Graph ====================

function Connect-GraphApp {
    # App-only, certificate auth. 'Global' = commercial, 'USGov' = GCC High (graph.microsoft.us).
    param([Parameter(Mandatory)]$Side)
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }   # one cloud at a time; tokens are not interchangeable
    try {
        Connect-MgGraph -Environment $Side.GraphEnvironment -TenantId $Side.TenantId -ClientId $Side.ClientId `
            -CertificateThumbprint $Side.CertificateThumbprint -NoWelcome -ContextScope Process
    }
    catch { throw "App sign-in failed ($($_.Exception.Message)). If this is a new machine or the certificate expired, run: .\Migrate-PlannerToGccHigh.ps1 -Step Setup -Side <Source|Destination>" }
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

function Get-OrgDomains {
    # First Graph call of Export/Import: translate a 401/403 right after Setup into plain advice.
    $tid = (Get-MgContext).TenantId
    try { @((Invoke-Graph -Uri "/v1.0/organization/$tid`?`$select=verifiedDomains").verifiedDomains) }
    catch {
        if ($_.Exception.Data['StatusCode'] -in 401, 403) { throw "The migration app's permissions aren't active yet (HTTP $($_.Exception.Data['StatusCode'])). If Setup just ran, wait 5 minutes and rerun. Otherwise rerun -Step Setup for this side." }
        throw
    }
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

function Connect-AzIfNeeded($SideCfg) {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    $ok = $ctx -and $ctx.Environment -and $ctx.Environment.Name -eq $SideCfg.AzEnvironment -and (-not $SideCfg.TenantId -or [string]$ctx.Tenant.Id -eq $SideCfg.TenantId)
    if (-not $ok) {
        $a = @{ Environment = $SideCfg.AzEnvironment; Scope = 'Process' }
        if ($SideCfg.TenantId) { $a.Tenant = $SideCfg.TenantId }
        Connect-AzAccount @a | Out-Null
    }
}

function Connect-DataverseSide($SideCfg, [string]$EnvironmentUrl) {
    # AuthMode Az: delegated sign-in via Az.Accounts (Learn Dataverse PowerShell quickstart pattern).
    # AuthMode DeviceCode: your own public-client app with the Dynamics CRM user_impersonation permission.
    $url = if ($EnvironmentUrl) { $EnvironmentUrl } else { $SideCfg.DataverseUrl }
    if (-not $url) { throw 'No Dataverse environment URL' }
    $mode = if ($SideCfg.DataverseAuth) { $SideCfg.DataverseAuth } else { 'Az' }
    $script:DvUrl = $url.TrimEnd('/') + '/'
    $script:DvToken = $null; $script:DvRefresh = $null; $script:DvAuth = $mode
    if ($mode -eq 'Az') { Connect-AzIfNeeded $SideCfg }
    else {
        if (-not $SideCfg.DataverseClientId) { throw 'DataverseAuth=DeviceCode needs DataverseClientId (public-client app with the Dynamics CRM user_impersonation delegated permission)' }
        # Learn: public clients use "<environment-url>/user_impersonation"; offline_access returns a refresh token.
        $script:DvDevice = @{ Authority = (Get-AuthorityHost $SideCfg.AzEnvironment); TenantId = $SideCfg.TenantId; ClientId = $SideCfg.DataverseClientId
                              Scope = "$($url.TrimEnd('/'))/user_impersonation offline_access" }
        Request-DeviceCodeToken
    }
    $who = Invoke-Dv -Path 'WhoAmI'
    Write-MigLog OK "Dataverse connected ($mode): $script:DvUrl userId=$($who.UserId)"
    return $who
}

function Get-DataverseEnvironmentUrls($SideCfg) {
    # DataverseUrl if set; otherwise every enabled environment the signed-in user can see (Global Discovery Service).
    if ($SideCfg.DataverseUrl) { return , @($SideCfg.DataverseUrl) }
    if ($SideCfg.DataverseAuth -eq 'DeviceCode') { throw 'Set DataverseUrl in config\settings.json when DataverseAuth = DeviceCode' }
    Connect-AzIfNeeded $SideCfg
    $gds = Get-GlobalDiscoUrl $SideCfg.AzEnvironment
    $t = Get-AzAccessToken -ResourceUrl "$gds/" -AsSecureString
    $tok = ConvertFrom-SecureString -SecureString $t.Token -AsPlainText
    $r = Invoke-WebRequest -Uri "$gds/api/discovery/v2.0/Instances?`$select=Url,FriendlyName,UniqueName,State" -Headers @{ Authorization = "Bearer $tok"; Accept = 'application/json' } -SkipHttpErrorCheck
    if ([int]$r.StatusCode -ge 400) { throw "Global Discovery Service returned HTTP $($r.StatusCode)" }
    $envs = @(($r.Content | ConvertFrom-Json).value | Where-Object { [int]$_.State -eq 0 })
    foreach ($e in $envs) { Write-MigLog INFO "Dataverse environment found: $($e.FriendlyName) $($e.Url)" }
    , @($envs | ForEach-Object { $_.Url })
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
        [Parameter(Mandatory)][string]$Path,
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
            $result[$Matches[1]] = @{ Id = [string]$p.Value; Entity = $Row."$($p.Name)@Microsoft.Dynamics.CRM.lookuplogicalname" }
        }
    }
    return $result
}

$script:DvSystemColumns = @(
    'createdon', 'createdby', 'createdonbehalfby', 'modifiedon', 'modifiedby', 'modifiedonbehalfby', 'ownerid', 'owninguser',
    'owningteam', 'owningbusinessunit', 'overriddencreatedon', 'importsequencenumber', 'timezoneruleversionnumber',
    'utcconversiontimezonecode', 'versionnumber', 'statecode', 'statuscode', 'transactioncurrencyid', 'exchangerate')

function New-DvEntityPayload {
    <# Builds a Project schedule API entity from an exported source row. Columns = documented allowlist + verified extras,
       intersected with the destination's creatable columns (-AllCreatable: every creatable non-system column). Lookups are
       rebound only when the target is in IdMap (created, or planned in this run); otherwise dropped and reported. #>
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

#region ==================== Setup (app registration) ====================

function Get-RequiredPermissions([string]$SideName) {
    if ($SideName -eq 'Source') { 'Tasks.Read.All', 'Group.Read.All', 'User.Read.All', 'Organization.Read.All' }
    else { 'Tasks.ReadWrite.All', 'Group.Read.All', 'Group.Create', 'GroupMember.ReadWrite.All', 'User.Read.All', 'Organization.Read.All' }
}

function Invoke-StepSetup([string]$SetupSide) {
    # Idempotent. Creates (or updates) the app registration, a non-exportable self-signed certificate, and the Graph
    # application permissions (IDs read at runtime from the Microsoft Graph service principal), granted via appRoleAssignedTo.
    if (-not $SetupSide) { $SetupSide = $Side }
    if (-not $SetupSide) { throw '-Side Source|Destination is required for -Step Setup' }
    Initialize-Prerequisites @('Graph')
    $null = Read-MigrationSettings -Path $SettingsPath
    $raw = Get-Content $SettingsPath -Raw | ConvertFrom-Json
    $cfg = $raw.$SetupSide
    $perms = Get-RequiredPermissions $SetupSide
    $appName = "Planner GCC High Migration ($SetupSide)"
    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $tenantArg = @{}; if ($cfg.TenantId -match '^[0-9a-fA-F-]{36}$') { $tenantArg.TenantId = $cfg.TenantId }
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }
    Write-MigLog INFO "Sign in as an ADMIN of the $SetupSide tenant in the browser window..."
    Connect-MgGraph -Environment $cfg.GraphEnvironment @tenantArg -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All' -NoWelcome -ContextScope Process
    $ctx = Get-MgContext
    Write-MigLog INFO "Signed in to $($ctx.Environment) tenant $($ctx.TenantId) as $($ctx.Account)"

    $graphSp = @((Invoke-Graph -Uri "/v1.0/servicePrincipals?`$filter=appId eq '$graphAppId'&`$select=id,appRoles").value)[0]
    if (-not $graphSp) { throw 'Microsoft Graph service principal not found in this tenant' }
    $roles = foreach ($p in $perms) {
        $r = $graphSp.appRoles | Where-Object { $_.value -eq $p -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $r) { throw "Application permission $p is not offered by Microsoft Graph in this cloud" }
        $r
    }
    $rra = @(@{ resourceAppId = $graphAppId; resourceAccess = @($roles | ForEach-Object { @{ id = $_.id; type = 'Role' } }) })

    $certOk = $cfg.CertificateThumbprint -and (Test-Path "Cert:\CurrentUser\My\$($cfg.CertificateThumbprint)")
    if ($certOk) { $cert = Get-Item "Cert:\CurrentUser\My\$($cfg.CertificateThumbprint)" }
    else {
        $cert = New-SelfSignedCertificate -Subject "CN=$appName" -CertStoreLocation 'Cert:\CurrentUser\My' -KeyExportPolicy NonExportable `
            -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddMonths($CertMonths)
        Write-MigLog OK "Certificate created in Cert:\CurrentUser\My, thumbprint $($cert.Thumbprint), expires $($cert.NotAfter)"
    }
    $keyCred = @{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [Convert]::ToBase64String($cert.RawData); displayName = "CN=$appName" }

    $app = $null
    if ($cfg.ClientId) { $app = @((Invoke-Graph -Uri "/v1.0/applications?`$filter=appId eq '$($cfg.ClientId)'&`$select=id,appId").value)[0] }
    if (-not $app) {
        $app = Invoke-Graph -Method POST -Uri '/v1.0/applications' -Body @{ displayName = $appName; signInAudience = 'AzureADMyOrg'; requiredResourceAccess = $rra; keyCredentials = @($keyCred) }
        if (-not $app.appId) { throw 'Application create returned no appId' }
        Write-MigLog OK "App '$appName' created, appId=$($app.appId)"
    }
    else {
        $patch = @{ requiredResourceAccess = $rra }
        if (-not $certOk) { $patch.keyCredentials = @($keyCred) }   # replaces the old certificate (it isn't on this machine)
        Invoke-Graph -Method PATCH -Uri "/v1.0/applications/$($app.id)" -Body $patch | Out-Null
        Write-MigLog OK "App appId=$($app.appId) already exists - permissions$(if (-not $certOk) { ' and certificate' }) updated"
    }
    $sp = @((Invoke-Graph -Uri "/v1.0/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$select=id").value)[0]
    if (-not $sp) { $sp = Invoke-Graph -Method POST -Uri '/v1.0/servicePrincipals' -Body @{ appId = $app.appId } }
    $have = @((Get-GraphAll -Uri "/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments") | ForEach-Object { [string]$_.appRoleId })
    foreach ($r in ($roles | Where-Object { [string]$_.id -notin $have })) {
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
    $raw.$SetupSide.TenantId = $ctx.TenantId; $raw.$SetupSide.ClientId = $app.appId; $raw.$SetupSide.CertificateThumbprint = $cert.Thumbprint
    if ($SetupSide -eq 'Destination' -and -not $raw.Destination.FallbackOwnerUpn) {
        $raw.Destination | Add-Member -NotePropertyName FallbackOwnerUpn -NotePropertyValue ([string]$ctx.Account) -Force
    }
    $raw | ConvertTo-Json -Depth 8 | Set-Content -Path $SettingsPath -Encoding utf8
    Write-MigLog OK "Setup done for $SetupSide. New permissions can take a few minutes to apply."
    Write-MigLog WARN "The certificate's private key is non-exportable: run every step from this Windows account on this machine."
}

function Confirm-AppReady([string]$SideName) {
    $raw = Get-Content $SettingsPath -Raw | ConvertFrom-Json
    $cfg = $raw.$SideName
    $certOk = [bool]$cfg.SkipCertCheck -or ($cfg.CertificateThumbprint -and (Test-Path "Cert:\CurrentUser\My\$($cfg.CertificateThumbprint)"))
    if (-not $cfg.ClientId -or -not $certOk) {
        Write-MigLog INFO "First run on this machine for the $SideName tenant: registering the migration app."
        Invoke-StepSetup -SetupSide $SideName
        Write-MigLog INFO 'Waiting 60s for the new permissions to apply...'
        Start-Sleep -Seconds $(if ($cfg.SkipCertCheck) { 0 } else { 60 })
    }
}

#endregion

#region ==================== Export (commercial) ====================

function Invoke-StepDiscover {
    $s = Read-MigrationSettings -Path $SettingsPath
    Connect-GraphApp -Side $s.Source
    $rows = Find-GroupsWithPlans
    $target = Join-Path (Split-Path $SettingsPath -Parent) 'groups.discovered.csv'
    @($rows) | Export-Csv -Path $target -NoTypeInformation
    Write-MigLog OK "$(@($rows).Count) groups have plans -> $target"
}

function Find-GroupsWithPlans {
    $groups = Get-GraphAll -Uri "/v1.0/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,mail"
    Write-MigLog INFO "Scanning $($groups.Count) Microsoft 365 groups for plans..."
    foreach ($g in $groups) {
        try { $plans = Get-GraphAll -Uri "/v1.0/groups/$($g.id)/planner/plans?`$select=id,title" }
        catch { Write-MigLog WARN "Group $($g.displayName) ($($g.id)): $($_.Exception.Message)"; continue }
        if ($plans.Count) { [pscustomobject]@{ sourceGroupId = $g.id; destGroupId = ''; sourceDisplayName = $g.displayName; sourceMail = $g.mail; planCount = $plans.Count } }
    }
}

function Invoke-StepExportBasic {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    Connect-GraphApp -Side $s.Source
    $outRoot = Join-Path $s.ExportFolder 'basic'
    New-Item -ItemType Directory -Force -Path $outRoot | Out-Null
    $problems = Join-Path $s.LogFolder 'export-problems.csv'

    @{ verifiedDomains = @(Get-OrgDomains) } | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $outRoot '_source-org.json') -Encoding utf8

    $groupsCsv = Join-Path $cfgDir 'groups.csv'
    if (Test-Path $groupsCsv) { $groupRows = @(Import-Csv $groupsCsv | Where-Object sourceGroupId); Write-MigLog INFO "Using $($groupRows.Count) groups from config\groups.csv" }
    else {
        $groupRows = @(Find-GroupsWithPlans)
        $groupRows | Export-Csv -Path $groupsCsv -NoTypeInformation
        Write-MigLog OK "$($groupRows.Count) groups with plans found -> config\groups.csv (all of them will be migrated)"
    }

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
    $planCount = 0; $taskCount = 0; $skipped = 0
    foreach ($g in $groupRows) {
        $gid = $g.sourceGroupId
        try {
            $info = Invoke-Graph -Uri "/v1.0/groups/$gid`?`$select=id,displayName,mailNickname,description,visibility"
            # (Get-GraphAll) in parentheses so the pipeline gets the items, not the array as one object.
            $owners = @((Get-GraphAll -Uri "/v1.0/groups/$gid/owners?`$select=id,userPrincipalName") | Where-Object { $_.userPrincipalName })
            $members = @((Get-GraphAll -Uri "/v1.0/groups/$gid/members?`$select=id,userPrincipalName") | Where-Object { $_.userPrincipalName })
            ($owners + $members) | ForEach-Object { Resolve-SourceUser $_.id }
            $dir = Join-Path $outRoot $gid
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
            [pscustomobject]@{ id = $gid; displayName = $info.displayName; mailNickname = $info.mailNickname; description = $info.description; visibility = $info.visibility
                               ownerIds = @($owners | ForEach-Object { $_.id }); memberIds = @($members | ForEach-Object { $_.id }) } |
                ConvertTo-Json -Depth 5 | Set-Content (Join-Path $dir '_group.json') -Encoding utf8
            $plans = Get-GraphAll -Uri "/v1.0/groups/$gid/planner/plans"
        }
        catch { $skipped++; Write-MigLog ERROR "Group $($g.sourceDisplayName) ($gid) skipped: $($_.Exception.Message)"; Add-CsvRow $problems ([pscustomobject]@{ group = $gid; plan = ''; issue = $_.Exception.Message }); continue }
        Write-MigLog INFO "Group $($info.displayName) ($gid): $($plans.Count) plan(s)"
        foreach ($plan in $plans) {
            try {
                $planDetails = Invoke-Graph -Uri "/v1.0/planner/plans/$($plan.id)/details"
                $buckets = Get-GraphAll -Uri "/v1.0/planner/plans/$($plan.id)/buckets"
                $tasks = Get-GraphAll -Uri "/v1.0/planner/plans/$($plan.id)/tasks"
                $commentPosts = 0; $threads = 0
                $taskRecords = foreach ($t in $tasks) {
                    $td = Invoke-Graph -Uri "/v1.0/planner/tasks/$($t.id)/details"
                    $posts = @()
                    if (-not $SkipComments -and $t.conversationThreadId) {
                        $threads++
                        try { $posts = Get-GraphAll -Uri "/v1.0/groups/$gid/threads/$($t.conversationThreadId)/posts" }
                        catch { Write-MigLog WARN "Comments for task '$($t.title)' not readable: $($_.Exception.Message)" }
                        $commentPosts += $posts.Count
                    }
                    if ($t.assignments) { $t.assignments.PSObject.Properties.Name | ForEach-Object { Resolve-SourceUser $_ } }
                    [pscustomobject]@{ task = $t; details = $td; comments = @($posts) }
                }
                $record = [pscustomobject]@{ exportedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); sourceGroupId = $gid
                                             plan = $plan; planDetails = $planDetails; buckets = @($buckets); tasks = @($taskRecords) }
                $record | ConvertTo-Json -Depth 40 | Set-Content -Path (Join-Path $dir "$($plan.id).json") -Encoding utf8
                $users | ConvertTo-Json -Depth 5 | Set-Content -Path $usersPath -Encoding utf8
                Add-CsvRow -Path $inventoryPath -Row ([pscustomobject]@{
                    sourceGroupId = $gid; groupName = $info.displayName; planId = $plan.id; planTitle = $plan.title
                    buckets = $buckets.Count; tasks = $tasks.Count
                    checklistItems = ($tasks | Measure-Object -Property checklistItemCount -Sum).Sum
                    references = ($tasks | Measure-Object -Property referenceCount -Sum).Sum
                    tasksWithCommentThreads = $threads; commentPosts = $commentPosts; exportedAtUtc = $record.exportedAtUtc })
                Write-MigLog OK "Exported '$($plan.title)': $($buckets.Count) buckets, $($tasks.Count) tasks, $commentPosts comment posts"
                $planCount++; $taskCount += $tasks.Count
            }
            catch { $skipped++; Write-MigLog ERROR "Plan '$($plan.title)' skipped: $($_.Exception.Message)"; Add-CsvRow $problems ([pscustomobject]@{ group = $gid; plan = $plan.title; issue = $_.Exception.Message }) }
        }
    }
    $users | ConvertTo-Json -Depth 5 | Set-Content -Path $usersPath -Encoding utf8
    Write-MigLog OK "Basic export complete -> $outRoot"
    return [pscustomobject]@{ Plans = $planCount; Tasks = $taskCount; Skipped = $skipped }
}

function Invoke-StepExportTaskChat {
    # Task chat replaced comments for basic plans in early 2026 (MC1193421). Readable only through Graph BETA
    # 'List messages': Global cloud, delegated (Tasks.Read). Appended to the description as text on import,
    # because 'Create message' is not available in US Gov L4.
    $s = Read-MigrationSettings -Path $SettingsPath
    $exportRoot = Join-Path $s.ExportFolder 'basic'
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }
    Write-MigLog INFO 'Task chat: sign in (browser) as a user who is a MEMBER of the groups being migrated...'
    Connect-MgGraph -Environment $s.Source.GraphEnvironment -TenantId $s.Source.TenantId -Scopes 'Tasks.Read', 'User.ReadBasic.All' -NoWelcome -ContextScope Process
    Write-MigLog OK "Signed in (delegated) as $((Get-MgContext).Account) for task chat"
    $users = Get-Content (Join-Path $exportRoot 'users.json') -Raw | ConvertFrom-Json -AsHashtable
    $nameById = @{}; foreach ($k in $users.Keys) { if ($users[$k].displayName) { $nameById[$k] = $users[$k].displayName } }
    $problems = Join-Path $s.LogFolder 'taskchat-export-problems.csv'
    $total = 0
    foreach ($f in (Get-PlanFiles $exportRoot)) {
        $rec = Get-Content $f.FullName -Raw | ConvertFrom-Json
        if ($OnlyIds -and $rec.plan.id -notin $OnlyIds) { continue }
        foreach ($tr in $rec.tasks) {
            try { $msgs = Get-GraphAll -Uri "/beta/planner/tasks/$($tr.task.id)/messages" }
            catch { Add-CsvRow -Path $problems -Row ([pscustomobject]@{ plan = $rec.plan.title; task = $tr.task.title; issue = $_.Exception.Message }); continue }
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
    }
    Write-MigLog OK "Task chat export done: $total messages. Tasks this user couldn't read: $problems"
    return $total
}

function Invoke-StepExportPremium {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    Initialize-Prerequisites @('Az')
    $outRoot = Join-Path $s.ExportFolder 'premium'
    New-Item -ItemType Directory -Force -Path $outRoot | Out-Null
    $childTables = 'msdyn_projectbucket', 'msdyn_projecttask', 'msdyn_projectteam', 'msdyn_resourceassignment',
                   'msdyn_projecttaskdependency', 'msdyn_projectchecklist', 'msdyn_projectlabel', 'msdyn_projecttasktolabel'
    $filterPath = Join-Path $cfgDir 'premium-projects.csv'
    $wanted = if ($OnlyIds) { @($OnlyIds | ForEach-Object { $_.ToLowerInvariant() }) }
              elseif (Test-Path $filterPath) { @(Import-Csv $filterPath | ForEach-Object { $_.sourceProjectId.Trim().ToLowerInvariant() }) } else { $null }
    $inventory = Join-Path $outRoot 'inventory.csv'; $labelsCsv = Join-Path $outRoot 'labels-manual.csv'
    $exported = 0
    foreach ($envUrl in (Get-DataverseEnvironmentUrls $s.Source)) {
        Connect-DataverseSide $s.Source $envUrl | Out-Null
        $info = @{}
        foreach ($t in @('msdyn_project') + $childTables) { $ti = Get-DvTableInfo -LogicalName $t; if ($ti) { $info[$t] = $ti } }
        if (-not $info['msdyn_project']) { Write-MigLog INFO "No premium plans (Project tables) in $envUrl"; continue }
        $projects = Get-DvAll -Path $info['msdyn_project'].EntitySetName
        if ($wanted) { $projects = @($projects | Where-Object { ([string]$_.($info['msdyn_project'].PrimaryId)).ToLowerInvariant() -in $wanted }) }
        if (-not $projects.Count) { Write-MigLog INFO "No premium plans in $envUrl"; continue }
        $info | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $outRoot "_source-metadata-$(([uri]$envUrl).Host).json") -Encoding utf8
        $taskPk = $info['msdyn_projecttask'].PrimaryId
        $teamPk = if ($info['msdyn_projectteam']) { $info['msdyn_projectteam'].PrimaryId }
        Write-MigLog INFO "Exporting $($projects.Count) premium plan(s) from $envUrl"
        function Get-ChildRows([string]$Table, [string]$ProjectId, [string[]]$TaskIds) {
            $ti = $info[$Table]; if (-not $ti) { return , @() }
            $projLookup = $ti.Lookups | Where-Object ReferencedEntity -eq 'msdyn_project' | Select-Object -First 1
            if ($projLookup) { return Get-DvAll -Path "$($ti.EntitySetName)?`$filter=_$($projLookup.ReferencingAttribute)_value eq $ProjectId" }
            $taskLookup = $ti.Lookups | Where-Object ReferencedEntity -eq 'msdyn_projecttask' | Select-Object -First 1
            if (-not $taskLookup -or -not $TaskIds) { return , @() }
            $rows = @()
            foreach ($chunk in (Split-IntoChunks -Items $TaskIds -Size 50)) {
                $vals = ($chunk | ForEach-Object { "'$_'" }) -join ','
                $rows += Get-DvAll -Path "$($ti.EntitySetName)?`$filter=Microsoft.Dynamics.CRM.In(PropertyName='$($taskLookup.ReferencingAttribute)',PropertyValues=[$vals])"
            }
            return , $rows
        }
        foreach ($p in $projects) {
            try {
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
                    catch { Write-MigLog WARN "Team member '$($tm.msdyn_name)': user lookup failed - will import as generic member" }
                }
                [pscustomobject]@{ exportedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); sourceUrl = $envUrl; project = $p; tables = $tables; teamUpn = $teamUpn } |
                    ConvertTo-Json -Depth 40 | Set-Content -Path (Join-Path $outRoot "$pid_.json") -Encoding utf8
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
                $counts = [ordered]@{ projectId = $pid_; project = $p.msdyn_subject; environment = $envUrl }
                foreach ($k in $tables.Keys) { $counts[$k] = @($tables[$k]).Count }
                Add-CsvRow -Path $inventory -Row ([pscustomobject]$counts)
                $exported++
                Write-MigLog OK "Exported premium plan '$($p.msdyn_subject)': $(@($tables['msdyn_projecttask']).Count) tasks"
            }
            catch { Write-MigLog ERROR "Premium plan '$($p.msdyn_subject)' skipped: $($_.Exception.Message)" }
        }
    }
    Write-MigLog OK "Premium export complete: $exported plan(s) -> $outRoot"
    return $exported
}

function Invoke-StepExport {
    # One command for the commercial side. Ends with a summary that makes any skipped part impossible to miss.
    Initialize-Prerequisites @('Graph')
    $s = Read-MigrationSettings -Path $SettingsPath
    Confirm-AppReady 'Source'
    $basic = Invoke-StepExportBasic
    $chat = 'ok'; $chatCount = 0; $prem = 'ok'; $premCount = 0
    try { $chatCount = Invoke-StepExportTaskChat }
    catch { $chat = "SKIPPED: $($_.Exception.Message). Fix: sign in as a member of the groups, then run -Step ExportTaskChat (older comments are exported either way)." }
    try { $premCount = Invoke-StepExportPremium }
    catch { $prem = "SKIPPED: $($_.Exception.Message). Fix: set Source.DataverseUrl in config\settings.json (Power Platform admin center > Environments), then run -Step ExportPremium." }
    @{ exportedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); basicPlans = $basic.Plans; basicTasks = $basic.Tasks; basicSkipped = $basic.Skipped
       taskChat = $chat; taskChatMessages = $chatCount; premium = $prem; premiumPlans = $premCount } |
        ConvertTo-Json | Set-Content (Join-Path $s.StateFolder 'export-status.json') -Encoding utf8
    $lines = @(
        '===================== EXPORT SUMMARY =====================',
        "Basic plans   : $($basic.Plans) plans, $($basic.Tasks) tasks$(if ($basic.Skipped) { "  ($($basic.Skipped) SKIPPED - see logs\export-problems.csv)" })",
        "Task chat     : $(if ($chat -eq 'ok') { "$chatCount messages" } else { $chat })",
        "Premium plans : $(if ($prem -eq 'ok') { "$premCount plans" } else { $prem })",
        '==========================================================')
    $lines | Set-Content (Join-Path $s.LogFolder 'export-summary.txt') -Encoding utf8
    $allOk = ($chat -eq 'ok') -and ($prem -eq 'ok') -and -not $basic.Skipped
    foreach ($l in $lines) { Write-MigLog $(if ($allOk) { 'OK' } else { 'WARN' }) $l }
    Write-MigLog OK "Next, on this same machine and Windows account: .\Migrate-PlannerToGccHigh.ps1 -Step Import"
}

#endregion

#region ==================== Import (GCC High) ====================

function Invoke-StepMapTargets([switch]$DryRun) {
    # Maps every source user and group to GCC High. Creates missing groups (with mapped owners/members) unless -DryRun
    # or CreateMissingGroups is false. Writes state\targets.json, logs\user-mapping.csv, logs\group-mapping.csv.
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    $exportRoot = Join-Path $s.ExportFolder 'basic'
    Connect-GraphApp -Side $s.Destination
    $graphBase = Get-GraphBaseUrl $s.Destination.GraphEnvironment
    $srcDomains = @((Get-Content (Join-Path $exportRoot '_source-org.json') -Raw | ConvertFrom-Json).verifiedDomains)
    $dstDomains = @(Get-OrgDomains)
    $dstDomainNames = @(@($dstDomains | Where-Object isDefault) + @($dstDomains | Where-Object { -not $_.isDefault }) | ForEach-Object { $_.name })

    $domainMap = @{}
    if ($s.DomainMap -and @($s.DomainMap.PSObject.Properties).Count) { $s.DomainMap.PSObject.Properties | ForEach-Object { $domainMap[$_.Name.ToLowerInvariant()] = ([string]$_.Value).ToLowerInvariant() } }
    else { $domainMap = Get-AutoDomainMap -SourceDomains $srcDomains -DestDomains $dstDomains }
    $urlMapPath = Join-Path $cfgDir 'urlmap.csv'
    $urlMap = if ((Test-Path $urlMapPath) -and @(Import-Csv $urlMapPath | Where-Object sourcePrefix).Count) { @(Import-Csv $urlMapPath | Where-Object sourcePrefix) }
              else { @(Get-SharePointUrlMap -SourceDomains $srcDomains -DestDomains $dstDomains) }
    foreach ($u in $urlMap) { Write-MigLog INFO "Link rewrite: $($u.sourcePrefix) -> $($u.destPrefix)" }

    # ---- users ----
    $userMap = Get-UserMap $cfgDir
    $srcUsers = Get-Content (Join-Path $exportRoot 'users.json') -Raw | ConvertFrom-Json -AsHashtable
    $cache = @{}
    function Find-DestUser([string]$Upn) {
        if (-not $cache.ContainsKey($Upn)) {
            try { $cache[$Upn] = Invoke-Graph -Uri "/v1.0/users/$Upn`?`$select=id,userPrincipalName" } catch { $cache[$Upn] = $null }
        }
        $cache[$Upn]
    }
    $users = @{}
    $userReport = Join-Path $s.LogFolder 'user-mapping.csv'; if (Test-Path $userReport) { Remove-Item $userReport }
    foreach ($sid in $srcUsers.Keys) {
        $su = $srcUsers[$sid]
        $hit = $null; $method = 'NOT FOUND'
        foreach ($c in (Get-UserMatchCandidates -SourceUser $su -DestDomains $dstDomainNames -UserMap $userMap -DomainMap $domainMap)) {
            $u = Find-DestUser $c.Upn
            if ($u) { $hit = $u; $method = $c.Method; break }
        }
        if (-not $hit -and $su.displayName) {
            try {
                $byName = @((Invoke-Graph -Uri "/v1.0/users?`$filter=displayName eq '$(ConvertTo-ODataString $su.displayName)'&`$select=id,userPrincipalName").value)
                if ($byName.Count -eq 1) { $hit = $byName[0]; $method = 'display name (unique)' }
            } catch { }
        }
        if ($hit) { $users[$sid] = @{ destId = $hit.id; destUpn = $hit.userPrincipalName; method = $method } }
        Add-CsvRow $userReport ([pscustomobject]@{ sourceUpn = $su.upn; sourceName = $su.displayName; destUpn = $(if ($hit) { $hit.userPrincipalName } else { '' }); method = $method })
    }
    Write-MigLog $(if ($users.Count -eq $srcUsers.Count) { 'OK' } else { 'WARN' }) "Users mapped: $($users.Count)/$($srcUsers.Count) (details: $userReport; add misses to config\users.csv and rerun)"

    # ---- groups ----
    $groupsCsv = Join-Path $cfgDir 'groups.csv'
    $override = @{}
    if (Test-Path $groupsCsv) { Import-Csv $groupsCsv | Where-Object { $_.destGroupId } | ForEach-Object { $override[$_.sourceGroupId.Trim()] = $_.destGroupId.Trim() } }
    $statePath = Join-Path $s.StateFolder 'targets.json'
    $prev = if (Test-Path $statePath) { Get-Content $statePath -Raw | ConvertFrom-Json -AsHashtable } else { @{ groups = @{} } }
    $groups = @{}
    $groupReport = Join-Path $s.LogFolder 'group-mapping.csv'; if (Test-Path $groupReport) { Remove-Item $groupReport }
    $fallbackOwner = if ($s.Destination.FallbackOwnerUpn) { Find-DestUser $s.Destination.FallbackOwnerUpn }
    foreach ($gf in (Get-ChildItem -Path $exportRoot -Recurse -Filter '_group.json')) {
        $g = Get-Content $gf.FullName -Raw | ConvertFrom-Json
        $destId = $null; $method = ''; $created = $false
        if ($override[$g.id]) { $destId = $override[$g.id]; $method = 'groups.csv' }
        elseif ($prev.groups -and $prev.groups[$g.id]) { $destId = $prev.groups[$g.id].destGroupId; $method = $prev.groups[$g.id].method; $created = [bool]$prev.groups[$g.id].created }
        if (-not $destId -and $g.mailNickname) {
            $m = @((Invoke-Graph -Uri "/v1.0/groups?`$filter=mailNickname eq '$(ConvertTo-ODataString $g.mailNickname)'&`$select=id,displayName").value)
            if ($m.Count -eq 1) { $destId = $m[0].id; $method = 'same alias' }
        }
        if (-not $destId -and $g.displayName) {
            $m = @((Invoke-Graph -Uri "/v1.0/groups?`$filter=displayName eq '$(ConvertTo-ODataString $g.displayName)'&`$select=id,groupTypes").value | Where-Object { $_.groupTypes -contains 'Unified' })
            if ($m.Count -eq 1) { $destId = $m[0].id; $method = 'same name' }
        }
        if (-not $destId -and $s.CreateMissingGroups -and -not $DryRun) {
            $ownerIds = @($g.ownerIds | ForEach-Object { $users[$_].destId } | Where-Object { $_ } | Select-Object -Unique)
            if (-not $ownerIds -and $fallbackOwner) { $ownerIds = @($fallbackOwner.id) }
            $memberIds = @(@($g.memberIds + $g.ownerIds) | ForEach-Object { $users[$_].destId } | Where-Object { $_ } | Select-Object -Unique)
            # Microsoft Learn: at most 20 owners+members in the create call; owners required or the group can't be modified later.
            $bindOwners = @($ownerIds | Select-Object -First 20)
            $bindMembers = @($memberIds | Select-Object -First ([Math]::Max(0, 20 - $bindOwners.Count)))
            $body = @{ displayName = $g.displayName; mailNickname = $g.mailNickname; groupTypes = @('Unified'); mailEnabled = $true; securityEnabled = $false }
            if ($g.description) { $body.description = $g.description }
            if ($g.visibility -in 'Private', 'Public') { $body.visibility = $g.visibility }
            if ($bindOwners) { $body['owners@odata.bind'] = @($bindOwners | ForEach-Object { "$graphBase/v1.0/users/$_" }) }
            if ($bindMembers) { $body['members@odata.bind'] = @($bindMembers | ForEach-Object { "$graphBase/v1.0/users/$_" }) }
            if (-not $bindOwners) { Write-MigLog WARN "Group '$($g.displayName)': no owner could be mapped; it will be created without an owner (it can't be edited later). Set Destination.FallbackOwnerUpn to avoid this." }
            $ng = $null
            try { $ng = Invoke-Graph -Method POST -Uri '/v1.0/groups' -Body $body }
            catch {
                Write-MigLog WARN "Group create '$($g.displayName)' failed ($($_.Exception.Message)); retrying with alias '$($g.mailNickname)-gcch'"
                $body.mailNickname = "$($g.mailNickname)-gcch"
                try { $ng = Invoke-Graph -Method POST -Uri '/v1.0/groups' -Body $body }
                catch { Write-MigLog ERROR "Group '$($g.displayName)' could not be created: $($_.Exception.Message). Its plans will be skipped; map it in config\groups.csv and rerun." }
            }
            if ($ng.id) {
                $destId = $ng.id; $method = 'created'; $created = $true
                Write-MigLog OK "Created GCC High group '$($g.displayName)' ($destId)"
                foreach ($mid in ($memberIds | Where-Object { $_ -notin $bindMembers })) { Add-DestGroupMember -GroupId $destId -UserId $mid -GraphBase $graphBase | Out-Null }
            }
        }
        if ($destId) { $groups[$g.id] = @{ destGroupId = $destId; method = $method; created = $created; displayName = $g.displayName } }
        elseif ($DryRun) { $method = 'would be CREATED' }
        else { $method = 'NOT MAPPED (CreateMissingGroups is off)' }
        Add-CsvRow $groupReport ([pscustomobject]@{ sourceGroupId = $g.id; sourceName = $g.displayName; destGroupId = $destId; method = $method })
    }
    $data = @{ users = $users; groups = $groups; urlMap = @($urlMap | ForEach-Object { @{ sourcePrefix = $_.sourcePrefix; destPrefix = $_.destPrefix } }); domainMap = $domainMap }
    if (-not $DryRun) { $data | Save-TargetsFile -Path $statePath }
    Write-MigLog OK "Groups: $(@($groups.Values | Where-Object { $_.method -ne 'created' }).Count) matched, $(@($groups.Values | Where-Object created).Count) created (details: $groupReport)"
    return $data
}

function Save-TargetsFile {
    param([Parameter(Mandatory, ValueFromPipeline)]$Data, [Parameter(Mandatory)][string]$Path)
    process { $tmp = "$Path.tmp"; $Data | ConvertTo-Json -Depth 8 | Set-Content -Path $tmp -Encoding utf8; Move-Item $tmp $Path -Force }
}

function Add-DestGroupMember([string]$GroupId, [string]$UserId, [string]$GraphBase) {
    # Learn: a recently created group can return 400 until it replicates - retry after a brief delay.
    for ($i = 1; $i -le 6; $i++) {
        try { Invoke-Graph -Method POST -Uri "/v1.0/groups/$GroupId/members/`$ref" -Body @{ '@odata.id' = "$GraphBase/v1.0/directoryObjects/$UserId" } | Out-Null; return $true }
        catch {
            if ($_.Exception.Message -match 'already exist') { return $true }
            if ($i -eq 6) { Write-MigLog WARN "Could not add $UserId to group $GroupId : $($_.Exception.Message)"; return $false }
            Start-Sleep -Seconds 5
        }
    }
}

function New-PlanWithRetry([string]$GroupId, [string]$Title, [bool]$NewGroup, [int]$WaitSeconds) {
    # A group created minutes ago may not accept Planner yet (directory replication): retry 403/404 for new groups only.
    for ($i = 1; ; $i++) {
        try { return Invoke-Graph -Method POST -Uri '/v1.0/planner/plans' -Body @{ container = @{ containerId = $GroupId; type = 'group' }; title = $Title } }
        catch {
            $code = $_.Exception.Data['StatusCode']
            if (-not $NewGroup -or $code -notin 403, 404 -or $i -ge 10) { throw }
            Write-MigLog WARN "New group not ready for Planner yet (HTTP $code) - waiting ${WaitSeconds}s"
            Start-Sleep -Seconds $WaitSeconds
        }
    }
}

function Invoke-StepImportBasic([switch]$DryRun) {
    $s = Read-MigrationSettings -Path $SettingsPath
    $exportRoot = Join-Path $s.ExportFolder 'basic'
    $targetsPath = Join-Path $s.StateFolder 'targets.json'
    $targets = if (Test-Path $targetsPath) { Get-Content $targetsPath -Raw | ConvertFrom-Json -AsHashtable } else { Invoke-StepMapTargets -DryRun:$DryRun }
    $destIdBySrc = @{}; foreach ($k in $targets.users.Keys) { $destIdBySrc[$k] = $targets.users[$k].destId }
    $urlMap = @($targets.urlMap | ForEach-Object { [pscustomobject]$_ })
    $srcUsers = Get-Content (Join-Path $exportRoot 'users.json') -Raw | ConvertFrom-Json -AsHashtable
    $srcUpnById = @{}; foreach ($k in $srcUsers.Keys) { $srcUpnById[$k] = $srcUsers[$k].upn }
    $graphBase = Get-GraphBaseUrl $s.Destination.GraphEnvironment

    Connect-GraphApp -Side $s.Destination
    $statePath = Join-Path $s.StateFolder 'basic-idmap.json'
    $idMap = Get-IdMap -Path $statePath
    $problemsCsv = Join-Path $s.LogFolder 'basic-import-problems.csv'
    function Write-Problem([string]$Plan, [string]$Item, [string]$Issue) {
        Add-CsvRow -Path $problemsCsv -Row ([pscustomobject]@{ timeUtc = (Get-Date).ToUniversalTime().ToString('o'); plan = $Plan; item = $Item; issue = $Issue })
    }
    $memberCache = @{}
    function Get-DestMembers([string]$GroupId) {
        if (-not $memberCache.ContainsKey($GroupId)) {
            $set = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($m in (Get-GraphAll -Uri "/v1.0/groups/$GroupId/members?`$select=id")) { [void]$set.Add($m.id) }
            $memberCache[$GroupId] = $set
        }
        return , $memberCache[$GroupId]   # comma: return the HashSet itself (PowerShell would otherwise unroll it)
    }

    foreach ($f in (Get-PlanFiles $exportRoot)) {
        $rec = Get-Content $f.FullName -Raw | ConvertFrom-Json
        $srcPlan = $rec.plan
        if ($OnlyIds -and $srcPlan.id -notin $OnlyIds) { continue }
        $tg = $targets.groups[$rec.sourceGroupId]
        if (-not $tg) {
            if ($DryRun) { Write-MigLog INFO "[DryRun] plan '$($srcPlan.title)': its group would be created in GCC High, then the plan imported"; continue }
            Write-MigLog WARN "Plan '$($srcPlan.title)': its group isn't mapped (see logs\group-mapping.csv) - skipped"; Write-Problem $srcPlan.title $rec.sourceGroupId 'group not mapped'; continue
        }
        $destGroup = $tg.destGroupId
        try {
            $members = Get-DestMembers $destGroup
            $pk = "plan:$($srcPlan.id)"
            if (-not $idMap[$pk]) {
                $existing = @()
                try { $existing = @((Get-GraphAll -Uri "/v1.0/groups/$destGroup/planner/plans?`$select=id,title") | Where-Object { $_.title -eq $srcPlan.title }) }
                catch { if (-not $tg.created) { throw } }   # a just-created group may not answer Planner yet
                if ($existing) {
                    Write-MigLog WARN "Plan '$($srcPlan.title)' already exists in the GCC High group and isn't in this machine's state - skipped to avoid a duplicate. Rename or delete it there, then rerun."
                    Write-Problem $srcPlan.title $destGroup 'plan with the same title already exists in destination group (skipped)'; continue
                }
                if ($DryRun) { Write-MigLog INFO "[DryRun] would create plan '$($srcPlan.title)' in group $destGroup" }
                else {
                    $p = New-PlanWithRetry -GroupId $destGroup -Title $srcPlan.title -NewGroup ([bool]$tg.created) -WaitSeconds ([int]$s.NewGroupRetrySeconds)
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
                $a = Get-PlannerAssignmentsBody -SourceAssignments $t.assignments -DestIdBySourceId $destIdBySrc -SourceUpnById $srcUpnById
                foreach ($u in $a.Unmapped) { Write-Problem $srcPlan.title $t.title "assignee not mapped: $u" }
                foreach ($destUserId in @($a.Body.Keys)) {
                    if (-not $members.Contains($destUserId)) {
                        if ($DryRun -or -not $s.AddMissingAssigneesToGroups) { Write-Problem $srcPlan.title $t.title "assignee $destUserId is not a member of destination group $destGroup" }
                        elseif (Add-DestGroupMember -GroupId $destGroup -UserId $destUserId -GraphBase $graphBase) { [void]$members.Add($destUserId); Write-MigLog INFO "Added assignee $destUserId to group $destGroup" }
                    }
                }
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
                    # Attempts: all mapped assignees -> only current group members -> no assignees. A task is never lost over an assignment.
                    $attempts = @($a.Body)
                    if ($a.Body.Count) {
                        $attempts += (Get-PlannerAssignmentsBody -SourceAssignments $t.assignments -DestIdBySourceId $destIdBySrc -SourceUpnById $srcUpnById -OnlyDestIds @($members)).Body
                        $attempts += @{}
                    }
                    $nt = $null; $lastErr = $null
                    foreach ($asg in $attempts) {
                        $try = $body.Clone(); if ($asg.Count) { $try.assignments = $asg }
                        try { $nt = Invoke-Graph -Method POST -Uri '/v1.0/planner/tasks' -Body $try; if ($asg.Count -lt $a.Body.Count) { Write-Problem $srcPlan.title $t.title "created with $($asg.Count) of $($a.Body.Count) assignees ($lastErr)" }; break }
                        catch { $lastErr = $_.Exception.Message; if ($_.Exception.Data['StatusCode'] -notin 400, 403) { break } }
                    }
                    if (-not $nt -or -not $nt.id) { $failed++; Write-MigLog ERROR "Task '$($t.title)': $lastErr"; Write-Problem $srcPlan.title $t.title "task create failed: $lastErr"; continue }
                    $idMap[$tk] = $nt.id; Save-IdMap $idMap $statePath; $created++
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
                            Write-Problem $srcPlan.title $t.title "details failed (HTTP $code), will retry on next run: $($_.Exception.Message)"
                            continue
                        }
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
        catch { Write-MigLog ERROR "Plan '$($srcPlan.title)' failed: $($_.Exception.Message) - continuing with the next plan (rerun to retry)"; Write-Problem $srcPlan.title '' "plan failed: $($_.Exception.Message)" }
    }
    Write-MigLog OK "Basic import pass finished. Problems (if any): $problemsCsv"
}

function Invoke-StepVerify {
    $s = Read-MigrationSettings -Path $SettingsPath
    $exportRoot = Join-Path $s.ExportFolder 'basic'
    $idMap = Get-IdMap -Path (Join-Path $s.StateFolder 'basic-idmap.json')
    $targetsPath = Join-Path $s.StateFolder 'targets.json'
    $mapped = @{}
    if (Test-Path $targetsPath) { $t = Get-Content $targetsPath -Raw | ConvertFrom-Json -AsHashtable; foreach ($k in $t.users.Keys) { $mapped[$k] = 1 } }
    Connect-GraphApp -Side $s.Destination
    $out = Join-Path $s.LogFolder 'verify-basic.csv'
    if (Test-Path $out) { Remove-Item $out }
    $rows = [System.Collections.Generic.List[object]]::new()
    function Add-Check($Plan, $Item, $Check, $Expected, $Actual) {
        $rows.Add([pscustomobject]@{ plan = $Plan; item = $Item; check = $Check; expected = $Expected; actual = $Actual; result = $(if ("$Expected" -eq "$Actual") { 'PASS' } else { 'FAIL' }) })
    }
    foreach ($f in (Get-PlanFiles $exportRoot)) {
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
            # Expected assignees = source assignees that have a GCC High user.
            $srcA = if ($t.assignments) { @(@($t.assignments.PSObject.Properties) | Where-Object { $_ -and ($mapped.Count -eq 0 -or $mapped.ContainsKey($_.Name)) }).Count } else { 0 }
            $dstA = if ($d.assignments) { @($d.assignments.PSObject.Properties).Count } else { 0 }
            Add-Check $rec.plan.title $t.title 'assignee count (mapped users)' $srcA $dstA
        }
    }
    $rows | Export-Csv -Path $out -NoTypeInformation
    $fail = @($rows | Where-Object result -eq 'FAIL')
    Write-MigLog $(if ($fail.Count) { 'WARN' } else { 'OK' }) "Verification: $($rows.Count) checks, $($fail.Count) FAIL -> $out"
    return $fail.Count
}

function Get-DestinationDataverseUrl($s) {
    # DataverseUrl if set (no extra sign-in); else the single discovered environment that has the Project tables.
    if ($s.Destination.DataverseUrl) { return $s.Destination.DataverseUrl }
    $cands = @()
    foreach ($u in (Get-DataverseEnvironmentUrls $s.Destination)) {
        Connect-DataverseSide $s.Destination $u | Out-Null
        if (Get-DvTableInfo -LogicalName 'msdyn_project') { $cands += $u }
    }
    if ($cands.Count -eq 1) { return $cands[0] }
    if (-not $cands.Count) { throw 'No GCC High Dataverse environment has premium plans (Project tables). Create one premium plan in Planner once, or set Destination.DataverseUrl.' }
    throw "Several GCC High environments have premium plans: $($cands -join ', '). Set Destination.DataverseUrl in config\settings.json."
}

function Invoke-StepPreflight([switch]$DestinationOnly, [string]$ForScope, [string]$TestGroupId, [bool]$TestGroupIsNew) {
    $s = Read-MigrationSettings -Path $SettingsPath
    $cfgDir = Split-Path $SettingsPath -Parent
    $scopeNow = if ($ForScope) { $ForScope } else { $Scope }
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

    if ($scopeNow -in 'All', 'Basic') {
        if (-not $DestinationOnly -and (Test-Step 'Basic-Source' 'connect app-only (commercial)' { Connect-GraphApp -Side $s.Source; (Get-MgContext).TenantId })) {
            $groupsCsv = Join-Path $cfgDir 'groups.csv'
            $groupRows = if (Test-Path $groupsCsv) { @(Import-Csv $groupsCsv) } else { @() }
            $sampled = 0; $withThread = 0
            foreach ($g in $groupRows) {
                if ($sampled -ge $SampleTasks) { break }
                try {
                    foreach ($plan in (Get-GraphAll -Uri "/v1.0/groups/$($g.sourceGroupId)/planner/plans")) {
                        foreach ($t in (Get-GraphAll -Uri "/v1.0/planner/plans/$($plan.id)/tasks")) { if ($sampled -ge $SampleTasks) { break }; $sampled++; if ($t.conversationThreadId) { $withThread++ } }
                    }
                }
                catch { Add-Result 'Basic-Source' "read plans in $($g.sourceDisplayName)" 'FAIL' $_.Exception.Message }
            }
            Add-Result 'Basic-Source' 'sampled tasks with a comment thread' 'INFO' "$withThread of $sampled"
        }
        if (Test-Step 'Basic-Dest' 'connect app-only (GCC High)' { Connect-GraphApp -Side $s.Destination; (Get-MgContext).TenantId }) {
            $tg = if ($TestGroupId) { $TestGroupId } else { $s.PreflightTestGroupId }; $tgNew = $TestGroupIsNew
            if (-not $tg) {
                $tp = Join-Path $s.StateFolder 'targets.json'
                if (Test-Path $tp) {
                    # Prefer a group that already existed; a just-created one may still be replicating.
                    $all = @((Get-Content $tp -Raw | ConvertFrom-Json -AsHashtable).groups.Values | Sort-Object { [bool]$_.created })
                    if ($all) { $tg = $all[0].destGroupId; $tgNew = [bool]$all[0].created }
                }
            }
            if (-not $tg) { Add-Result 'Basic-Dest' 'test group' 'FAIL' 'No destination group known yet (run -Step MapTargets, or set PreflightTestGroupId)' }
            elseif (Test-Step 'Basic-Dest' 'create plan (Tasks.ReadWrite.All app-only)' {
                    $script:pfPlan = New-PlanWithRetry -GroupId $tg -Title "PREFLIGHT $(Get-Date -Format s)" -NewGroup $tgNew -WaitSeconds ([int]$s.NewGroupRetrySeconds)
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
                Test-Step 'Basic-Dest' 'delete preflight plan' {
                    $p = Invoke-Graph -Uri "/v1.0/planner/plans/$($script:pfPlan.id)"
                    Invoke-Graph -Method DELETE -Uri "/v1.0/planner/plans/$($script:pfPlan.id)" -Headers @{ 'If-Match' = $p.'@odata.etag' } | Out-Null
                    'deleted'
                } | Out-Null
            }
        }
    }

    if ($scopeNow -in 'All', 'Premium') {
        try { Initialize-Prerequisites @('Az') } catch { Add-Result 'Premium' 'Az.Accounts module' 'FAIL' $_.Exception.Message }
        if (-not $DestinationOnly) {
            Test-Step 'Premium-Source' "sign in + find premium plans ($($s.Source.AzEnvironment))" {
                $n = 0
                foreach ($u in (Get-DataverseEnvironmentUrls $s.Source)) {
                    Connect-DataverseSide $s.Source $u | Out-Null
                    $ti = Get-DvTableInfo -LogicalName 'msdyn_project'
                    if ($ti) { $n += [int](Invoke-Dv -Path "$($ti.EntitySetName)?`$select=$($ti.PrimaryId)&`$count=true" -ExtraHeaders @{ Prefer = 'odata.maxpagesize=1' }).'@odata.count' }
                }
                "$n projects visible to this account"
            } | Out-Null
        }
        if (Test-Step 'Premium-Dest' "sign in + WhoAmI ($($s.Destination.AzEnvironment))" { $script:pfDvUrl = Get-DestinationDataverseUrl $s; (Connect-DataverseSide $s.Destination $script:pfDvUrl).UserId }) {
            $tiP = Get-DvTableInfo -LogicalName 'msdyn_project'
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
                Add-Result 'Premium-Dest' 'cleanup' 'INFO' "Delete the 'PREFLIGHT ... - delete me' premium plan in Planner (GCC High) by hand."
            }
        }
    }
    $csv = Join-Path $s.LogFolder 'preflight.csv'
    $results | Export-Csv -Path $csv -NoTypeInformation
    $results | Format-Table area, check, status, detail -AutoSize -Wrap | Out-Host
    $fails = @($results | Where-Object status -eq 'FAIL')
    if ($fails.Count) { Write-MigLog WARN "$($fails.Count) preflight FAIL(s) -> $csv" } else { Write-MigLog OK "Preflight passed -> $csv" }
    return , $results.ToArray()
}

function Invoke-StepImportPremium {
    # MUST run as a GCC High USER with a Project license (Learn: app/system/integration users can't call the schedule APIs).
    $s = Read-MigrationSettings -Path $SettingsPath
    $exportRoot = Join-Path $s.ExportFolder 'premium'
    if (-not (Get-ChildItem -Path $exportRoot -Filter '*.json' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '_*' })) { Write-MigLog INFO 'No premium plans were exported - nothing to import.'; return }
    Initialize-Prerequisites @('Az')
    $targetsPath = Join-Path $s.StateFolder 'targets.json'
    $destUpnBySrcUpn = @{}
    if (Test-Path $targetsPath) {
        $tg = Get-Content $targetsPath -Raw | ConvertFrom-Json -AsHashtable
        $srcUsers = Get-Content (Join-Path $s.ExportFolder 'basic\users.json') -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -AsHashtable
        if ($srcUsers) { foreach ($k in $tg.users.Keys) { if ($srcUsers[$k].upn) { $destUpnBySrcUpn[$srcUsers[$k].upn.ToLowerInvariant()] = $tg.users[$k].destUpn } } }
    }
    $dvUrl = Get-DestinationDataverseUrl $s
    Connect-DataverseSide $s.Destination $dvUrl | Out-Null
    $tableNames = 'msdyn_project', 'msdyn_projectbucket', 'msdyn_projecttask', 'msdyn_projectteam', 'msdyn_resourceassignment', 'msdyn_projecttaskdependency', 'msdyn_projectchecklist'
    $info = @{}; $setByEntity = @{}
    foreach ($t in $tableNames) {
        $ti = Get-DvTableInfo -LogicalName $t
        if ($ti) { $info[$t] = $ti; $setByEntity[$t] = $ti.EntitySetName } else { Write-MigLog WARN "Table $t not present in destination - its rows will be skipped" }
    }
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
                if ($it.Delete) {
                    try { Invoke-Dv -Method POST -Path 'msdyn_PssDeleteV1' -Body @{ RecordId = $it.Id; EntityLogicalName = $it.Table; OperationSetId = $osId } | Out-Null; $queued.Add([pscustomobject]@{ Table = $it.Table; Src = ''; Dest = $it.Id; Delete = $true }) }
                    catch { Write-Problem $Label "$($it.Table) $($it.Id)" "PssDeleteV1 rejected: $($_.Exception.Message)" }
                    continue
                }
                $srcId = [string]$it.Row.($ti.PrimaryId)
                $pl = New-DvEntityPayload -Row $it.Row -TableInfo $ti -SetNameByEntity $setByEntity -IdMap $Planned -Columns $cols[$it.Table] -AllCreatable:([bool]$it.AllCreatable)
                if ($it.Table -eq 'msdyn_projecttask' -and $pl.Payload.Contains("$(Get-NavProp 'msdyn_projecttask' 'msdyn_parenttask')@odata.bind")) { $pl.Payload.Remove('msdyn_outlinelevel') }
                foreach ($d in $pl.Dropped) { Write-Problem $Label "$($it.Table) $srcId" "not carried: $d" }
                try {
                    Invoke-Dv -Method POST -Path 'msdyn_PssCreateV1' -Body @{ Entity = $pl.Payload; OperationSetId = $osId } | Out-Null
                    $queued.Add([pscustomobject]@{ Table = $it.Table; Src = $srcId; Dest = [string]$pl.Payload[$ti.PrimaryId]; Delete = $false })
                }
                catch { Write-Problem $Label "$($it.Table) $srcId" "PssCreateV1 rejected: $($_.Exception.Message)" }
            }
            if (-not $queued.Count) { try { Invoke-Dv -Method POST -Path 'msdyn_AbandonOperationSetV1' -Body @{ OperationSetId = $osId } | Out-Null } catch { }; continue }
            Invoke-Dv -Method POST -Path 'msdyn_ExecuteOperationSetV1' -Body @{ OperationSetId = $osId } | Out-Null
            $deadline = (Get-Date).AddMinutes($waitMin)
            while ($queued.Count -and (Get-Date) -lt $deadline) {
                foreach ($q in @($queued)) {
                    $exists = Test-DvRowExists $q.Table $q.Dest
                    if ($q.Delete -and -not $exists) { [void]$queued.Remove($q) }
                    elseif (-not $q.Delete -and $exists) { $idMap[$q.Src] = $q.Dest; [void]$queued.Remove($q) }
                }
                Save-IdMap $idMap $statePath
                if ($queued.Count) { Start-Sleep -Seconds 10 }
            }
            foreach ($q in $queued) { Write-Problem $Label "$($q.Table) $($q.Src)$($q.Dest)" "not $(if ($q.Delete) { 'deleted' } else { 'present' }) $waitMin min after ExecuteOperationSetV1 (set $osId). Check destination Settings > Schedule Integration > Operation Sets / PSS Error Logs" }
        }
    }

    foreach ($f in (Get-ChildItem -Path $exportRoot -Filter '*.json' | Where-Object { $_.Name -notlike '_*' })) {
        $rec = Get-Content $f.FullName -Raw | ConvertFrom-Json
        $projTi = $info['msdyn_project']
        $srcProjId = [string]$rec.project.($projTi.PrimaryId)
        if ($OnlyIds -and $srcProjId -notin $OnlyIds) { continue }
        $label = [string]$rec.project.msdyn_subject
        try {
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
                        $destUpn = $destUpnBySrcUpn[$srcUpn.ToLowerInvariant()]
                        $br = $null
                        if ($destUpn) {
                            try {
                                $su = @((Invoke-Dv -Path "systemusers?`$select=systemuserid&`$filter=domainname eq '$destUpn'").value)[0]
                                if ($su) { $br = @((Invoke-Dv -Path "bookableresources?`$select=bookableresourceid&`$filter=_userid_value eq $($su.systemuserid)").value)[0] }
                            } catch { Write-Problem $label $tm.msdyn_name "named resource lookup failed: $($_.Exception.Message)" }
                        }
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

            # Remove the empty default bucket that msdyn_CreateProjectV1 creates (Learn), if it holds no tasks.
            $bTi = $info['msdyn_projectbucket']; $tTi = $info['msdyn_projecttask']
            $bProj = $bTi.Lookups | Where-Object ReferencedEntity -eq 'msdyn_project' | Select-Object -First 1
            $tBucket = $tTi.Lookups | Where-Object ReferencedEntity -eq 'msdyn_projectbucket' | Select-Object -First 1
            if ($bProj -and $tBucket) {
                $mine = @($idMap.Values)
                $extra = @((Get-DvAll -Path "$($bTi.EntitySetName)?`$select=$($bTi.PrimaryId)&`$filter=_$($bProj.ReferencingAttribute)_value eq $destProjId") | ForEach-Object { [string]$_.($bTi.PrimaryId) } | Where-Object { $_ -notin $mine })
                $del = @(foreach ($bid in $extra) {
                    $used = Get-DvAll -Path "$($tTi.EntitySetName)?`$select=$($tTi.PrimaryId)&`$filter=_$($tBucket.ReferencingAttribute)_value eq $bid"
                    if (-not $used.Count) { @{ Table = 'msdyn_projectbucket'; Id = $bid; Delete = $true } }
                })
                if ($del.Count) { Invoke-OperationSets $destProjId $label $del $planned }
            }

            $total = 0; $done = 0
            foreach ($tn in 'msdyn_projectbucket', 'msdyn_projecttask', 'msdyn_projectteam', 'msdyn_resourceassignment', 'msdyn_projecttaskdependency', 'msdyn_projectchecklist') {
                if (-not $info[$tn]) { continue }
                foreach ($row in @($rec.tables.$tn)) { if ($row) { $total++; if ($idMap[[string]$row.($info[$tn].PrimaryId)]) { $done++ } } }
            }
            Add-CsvRow (Join-Path $s.LogFolder 'verify-premium.csv') ([pscustomobject]@{ project = $label; rowsExpected = $total; rowsPresent = $done; result = $(if ($done -eq $total) { 'PASS' } else { 'FAIL' }) })
            Write-MigLog $(if ($done -eq $total) { 'OK' } else { 'WARN' }) "Premium plan '$label': $done/$total rows present in destination"
        }
        catch { Write-MigLog ERROR "Premium plan '$label' failed: $($_.Exception.Message) - continuing (rerun to retry)"; Write-Problem $label '' "plan failed: $($_.Exception.Message)" }
    }
    Write-MigLog OK "Premium import pass finished. Then: share each premium plan with its group in Planner, re-apply labels from export\premium\labels-manual.csv."
}

function Invoke-StepImport {
    # One command for the GCC High side.
    Initialize-Prerequisites @('Graph')
    $s = Read-MigrationSettings -Path $SettingsPath
    if (-not (Test-Path (Join-Path $s.ExportFolder 'basic\users.json'))) { throw "No export found in $($s.ExportFolder). Run -Step Export first (same machine, same Windows account)." }
    Confirm-AppReady 'Destination'
    $hasPremium = [bool](Get-ChildItem -Path (Join-Path $s.ExportFolder 'premium') -Filter '*.json' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '_*' })
    # 1) Preview the mapping without writing. 2) Test GCC High writes in an EXISTING matched group (throwaway plan).
    # 3) Only then create missing groups and import. If no group matched, one group must be created first to test in.
    $preview = Invoke-StepMapTargets -DryRun
    $testGroup = if ($s.PreflightTestGroupId) { $s.PreflightTestGroupId } else { @($preview.groups.Values | ForEach-Object { $_.destGroupId })[0] }
    $testIsNew = $false
    if (-not $testGroup) {
        Write-MigLog WARN 'No existing GCC High group matched, so groups are created BEFORE the write test (a test needs a group).'
        $real = Invoke-StepMapTargets
        $first = @($real.groups.Values)[0]
        if ($first) { $testGroup = $first.destGroupId; $testIsNew = [bool]$first.created }
    }
    $pf = Invoke-StepPreflight -DestinationOnly -ForScope $(if ($hasPremium) { 'All' } else { 'Basic' }) -TestGroupId $testGroup -TestGroupIsNew $testIsNew
    $basicFail = @($pf | Where-Object { $_.area -eq 'Basic-Dest' -and $_.status -eq 'FAIL' -and $_.check -notlike 'delete*' })
    if ($basicFail) { throw "Preflight failed in GCC High ($($basicFail[0].check): $($basicFail[0].detail)). No plans were imported$(if (-not $testIsNew) { ' and no groups were created' }). If Setup just ran, wait 5 minutes and rerun -Step Import." }
    $null = Invoke-StepMapTargets
    Invoke-StepImportBasic
    $vf = Invoke-StepVerify
    if ($hasPremium) {
        $premOk = -not @($pf | Where-Object { $_.area -like 'Premium*' -and $_.status -eq 'FAIL' })
        if ($premOk) { Invoke-StepImportPremium }
        else { Write-MigLog WARN 'Premium plans NOT imported: the GCC High premium preflight failed (see logs\preflight.csv). Basic plans are done. export\premium\ keeps the data.' }
    }
    else {
        $es = Join-Path $s.StateFolder 'export-status.json'
        $st = if (Test-Path $es) { Get-Content $es -Raw | ConvertFrom-Json }
        if ($st -and $st.premium -ne 'ok') { Write-MigLog WARN "PREMIUM PLANS WERE NOT IMPORTED because the export skipped them: $($st.premium)" }
        else { Write-MigLog INFO 'No premium plans were found in the commercial tenant, so none were imported.' }
    }
    Write-MigLog OK "IMPORT FINISHED. Review: logs\user-mapping.csv, logs\group-mapping.csv, logs\verify-basic.csv ($vf FAIL), logs\*problems.csv"
}

#endregion

# Dot-sourced (tests): load functions only.
if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $Step) { Get-Help $PSCommandPath; return }

switch ($Step) {
    'Export' { Invoke-StepExport }
    'Import' { Invoke-StepImport }
    'Init' { Invoke-StepInit }
    'Setup' { Invoke-StepSetup -SetupSide $Side }
    'Discover' { Invoke-StepDiscover }
    'Preflight' { Invoke-StepPreflight | Out-Null }
    'ExportBasic' { $null = Invoke-StepExportBasic }
    'ExportTaskChat' { $null = Invoke-StepExportTaskChat }
    'ExportPremium' { $null = Invoke-StepExportPremium }
    'MapTargets' { $null = Invoke-StepMapTargets }
    'DryRun' {
        # Same first-run app setup as Import, so DryRun works straight after Export on a fresh machine.
        Initialize-Prerequisites @('Graph')
        $null = Read-MigrationSettings -Path $SettingsPath
        Confirm-AppReady 'Destination'
        Invoke-StepImportBasic -DryRun
    }
    'ImportBasic' { Invoke-StepImportBasic }
    'Verify' { Invoke-StepVerify | Out-Null }
    'ImportPremium' { Invoke-StepImportPremium }
}

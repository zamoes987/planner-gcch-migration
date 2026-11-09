#Requires -Version 7.4
<#
  Offline end-to-end test of the TWO-COMMAND flow of Migrate-PlannerToGccHigh.ps1, with NO mapping files:
    -Step Export   (discover groups, export plans + group owners/members + task chat)
    -Step Import   (auto user/group mapping, create missing group, preflight, import, verify)
    -Step Import   again (must change nothing)
    -Step Setup -Side Destination
  against in-memory fakes of the Graph endpoints in both clouds. Proves wiring only; Microsoft's service behavior
  is what the Import's built-in preflight checks in your tenants.
  Run:  pwsh -File .\tests\Run-OfflineBasicE2E.ps1
#>
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$script = Join-Path $repo 'Migrate-PlannerToGccHigh.ps1'
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("gcch-e2e-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path (Join-Path $work 'config') -Force | Out-Null
@{
    Source = @{ TenantId = 'src'; ClientId = 'c'; CertificateThumbprint = 't'; GraphEnvironment = 'Global'; SkipCertCheck = $true }
    Destination = @{ TenantId = 'dst'; ClientId = 'c'; CertificateThumbprint = 't'; GraphEnvironment = 'USGov'; SkipCertCheck = $true }
    NewGroupRetrySeconds = 1
} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $work 'config\settings.json')
$settings = Join-Path $work 'config\settings.json'

# ---------- fake data ----------
function global:New-Obj($h) { $h | ConvertTo-Json -Depth 20 | ConvertFrom-Json }
$global:Src = @{
    org = @{ verifiedDomains = @(@{ name = 'contoso.com'; isDefault = $true; isInitial = $false }, @{ name = 'contoso.onmicrosoft.com'; isDefault = $false; isInitial = $true }) }
    groups = @{
        'g-src-1' = @{ id = 'g-src-1'; displayName = 'Eng'; mailNickname = 'eng'; description = 'Engineering'; visibility = 'Private'; owners = @('u1'); members = @('u1', 'u3') }
        'g-src-2' = @{ id = 'g-src-2'; displayName = 'Ops'; mailNickname = 'ops'; description = 'Operations'; visibility = 'Public'; owners = @('u3'); members = @('u3', 'u1') }
        'g-empty' = @{ id = 'g-empty'; displayName = 'No plans'; mailNickname = 'np'; owners = @(); members = @() } }
    plans = @{ 'g-src-1' = @(@{ id = 'P1'; title = 'Launch plan' }); 'g-src-2' = @(@{ id = 'P2'; title = 'Ops plan' }); 'g-empty' = @() }
    details = @{ P1 = @{ categoryDescriptions = @{ category1 = 'Red label'; category2 = $null } }; P2 = @{ categoryDescriptions = @{} } }
    buckets = @{ P1 = @(@{ id = 'B1'; name = 'Later'; orderHint = 'b' }, @{ id = 'B2'; name = 'First'; orderHint = 'a' }); P2 = @(@{ id = 'B3'; name = 'Ops bucket'; orderHint = 'a' }) }
    tasks = @{
        P1 = @(
            @{ id = 'T1'; title = 'Write spec'; bucketId = 'B1'; percentComplete = 50; priority = 3; conversationThreadId = 'th1'
               appliedCategories = @{ category1 = $true }; checklistItemCount = 2; referenceCount = 1
               assignments = @{ u1 = @{ orderHint = 'x' }; u2 = @{ orderHint = 'y' }; u3 = @{ orderHint = 'z' } } },
            @{ id = 'T2'; title = 'Ship it'; bucketId = 'B2'; percentComplete = 100; priority = 5; checklistItemCount = 0; referenceCount = 0; assignments = @{} })
        P2 = @(@{ id = 'T3'; title = 'Patch servers'; bucketId = 'B3'; percentComplete = 0; priority = 1; checklistItemCount = 0; referenceCount = 0; assignments = @{ u3 = @{ orderHint = 'a' } } }) }
    taskDetails = @{
        T1 = @{ description = 'Spec body'; previewType = 'automatic'
                checklist = @{ c1 = @{ title = 'one'; isChecked = $true; orderHint = 'a' }; c2 = @{ title = 'two'; isChecked = $false; orderHint = 'b' } }
                references = @{ 'https%3A//contoso%2Esharepoint%2Ecom/sites/eng/spec%2Edocx' = @{ alias = 'Spec'; type = 'Word' } } }
        T2 = @{ description = ''; previewType = 'automatic'; checklist = @{}; references = @{} }
        T3 = @{ description = 'Ops work'; previewType = 'automatic'; checklist = @{}; references = @{} } }
    posts = @(
        @{ createdDateTime = '2026-10-02T09:00:00Z'; from = @{ emailAddress = @{ name = 'Ann' } }; body = @{ content = '<p>Second &amp; last</p>' } },
        @{ createdDateTime = '2026-10-01T09:00:00Z'; from = @{ emailAddress = @{ name = 'Bob' } }; body = @{ content = '<div>First</div>' } })
    chat = @{
        T1 = @(@{ id = 'm1'; content = '<div>Chat hello</div>'; createdDateTime = '2026-10-03T09:00:00Z'; deletedTime = $null; createdBy = @{ user = @{ id = 'u1' } } },
               @{ id = 'm2'; content = 'deleted one'; createdDateTime = '2026-10-04T09:00:00Z'; deletedTime = '2026-10-05T00:00:00Z'; createdBy = @{ user = @{ id = 'u9' } } })
        T2 = @(); T3 = @() }
    users = @{ u1 = @{ id = 'u1'; userPrincipalName = 'a@contoso.com'; displayName = 'A'; mail = 'a@contoso.com' }
               u3 = @{ id = 'u3'; userPrincipalName = 'c@contoso.com'; displayName = 'C'; mail = 'c@contoso.com' } }   # u2, u9 deleted (404)
}
$global:Dst = @{
    org = @{ verifiedDomains = @(@{ name = 'contoso.us'; isDefault = $true; isInitial = $false }, @{ name = 'contosogov.onmicrosoft.us'; isDefault = $false; isInitial = $true }) }
    users = @{ 'a@contoso.us' = 'D-u1'; 'c@contoso.us' = 'D-u3'; 'admin@fake' = 'D-admin' }
    groupNick = @{ 'eng' = 'g-dst-1' }
    members = @{ 'g-dst-1' = [System.Collections.Generic.List[string]]@('D-u1') }
    newGroups = @{}; plans = @{}; planCats = @{}; buckets = @{}; tasks = @{}; details = @{}; n = 0
    throttled = $false; fail412Once = $true; memberFailOnce = $true; newPlanFailOnce = $true
    apps = @(); grants = @(); grantFailOnce = $true }
$global:FakeMode = $null; $global:LastScopes = $null; $global:Events = [System.Collections.Generic.List[string]]::new()

function global:Connect-MgGraph { param($Environment, $TenantId, $ClientId, $CertificateThumbprint, [string[]]$Scopes, [switch]$NoWelcome, $ContextScope) $global:FakeMode = $Environment; $global:LastScopes = $Scopes }
function global:Disconnect-MgGraph { $global:FakeMode = $null }
function global:Get-MgContext { [pscustomobject]@{ Environment = $global:FakeMode; TenantId = 'fake-tenant'; AuthType = $(if ($global:LastScopes) { 'Delegated' } else { 'AppOnly' }); Account = 'admin@fake' } }
function global:New-SelfSignedCertificate { param($Subject, $CertStoreLocation, $KeyExportPolicy, $KeySpec, $KeyLength, $KeyAlgorithm, $HashAlgorithm, $NotAfter)
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new($Subject, $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $req.CreateSelfSigned([DateTimeOffset]::Now, [DateTimeOffset]$NotAfter) }

$global:FakeHandler = {
    param($Method, $Uri, $Body, $Headers)
    $path = ($Uri -split '\?')[0]
    $q = [uri]::UnescapeDataString($(if ($Uri -match '\?(.*)$') { $Matches[1] } else { '' }))
    $b = if ($Body) { $Body | ConvertFrom-Json -AsHashtable } else { $null }
    $status = 200; $hdr = @{}; $out = $null
    if ($Method -in 'PATCH', 'DELETE' -and $path -like '/v1.0/planner/*' -and -not $Headers['If-Match']) { throw "fake: $Method without If-Match" }
    if ($global:FakeMode -eq 'Global') {
        switch -Regex ($path) {
            '^/v1.0/organization/fake-tenant$' { $out = New-Obj $global:Src.org; break }
            '^/v1.0/groups$' { $out = New-Obj @{ value = @($global:Src.groups.Values | ForEach-Object { @{ id = $_.id; displayName = $_.displayName; mail = "$($_.mailNickname)@contoso.com" } }) }; break }
            '^/v1.0/groups/([\w-]+)$' { $g = $global:Src.groups[$Matches[1]]; $out = New-Obj @{ id = $g.id; displayName = $g.displayName; mailNickname = $g.mailNickname; description = $g.description; visibility = $g.visibility }; break }
            '^/v1.0/groups/([\w-]+)/(owners|members)$' { $g = $global:Src.groups[$Matches[1]]; $ids = $g[$Matches[2]]
                $out = New-Obj @{ value = @($ids | ForEach-Object { @{ id = $_; userPrincipalName = $global:Src.users[$_].userPrincipalName } }) }; break }
            '^/v1.0/groups/([\w-]+)/planner/plans$' { $out = New-Obj @{ value = @($global:Src.plans[$Matches[1]]) }; break }
            '^/v1.0/planner/plans/(P\d)/details$' { $out = New-Obj $global:Src.details[$Matches[1]]; break }
            '^/v1.0/planner/plans/(P\d)/buckets$' { $out = New-Obj @{ value = @($global:Src.buckets[$Matches[1]]) }; break }
            '^/v1.0/planner/plans/(P\d)/tasks$' { $out = New-Obj @{ value = @($global:Src.tasks[$Matches[1]]) }; break }
            '^/v1.0/planner/tasks/(T\d)/details$' { $out = New-Obj $global:Src.taskDetails[$Matches[1]]; break }
            '^/beta/planner/tasks/(T\d)/messages$' {
                if (-not $global:LastScopes -or 'Tasks.Read' -notin $global:LastScopes) { $status = 403; $out = New-Obj @{ error = @{ message = 'task chat needs delegated Tasks.Read' } } }
                else { $out = New-Obj @{ value = @($global:Src.chat[$Matches[1]]) } }; break }
            '^/v1.0/groups/g-src-1/threads/th1/posts$' { $out = New-Obj @{ value = $global:Src.posts }; break }
            '^/v1.0/users/(.+)$' { if ($global:Src.users[$Matches[1]]) { $out = New-Obj $global:Src.users[$Matches[1]] } else { $status = 404; $out = New-Obj @{ error = @{ message = 'gone' } } }; break }
            default { throw "fake source: unhandled $Method $path" }
        }
    }
    elseif ($global:FakeMode -eq 'USGov') {
        $allGroups = @{}; $global:Dst.groupNick.Values | ForEach-Object { $allGroups[$_] = 1 }; $global:Dst.newGroups.Keys | ForEach-Object { $allGroups[$_] = 1 }
        switch -Regex ("$Method $path") {
            '^GET /v1.0/organization/fake-tenant$' { $out = New-Obj $global:Dst.org; break }
            '^GET /v1.0/users$' { $out = New-Obj @{ value = @() }; break }   # displayName fallback: no matches
            '^GET /v1.0/users/(.+)$' { $id = $global:Dst.users[$Matches[1]]; if ($id) { $out = New-Obj @{ id = $id; userPrincipalName = $Matches[1] } } else { $status = 404; $out = New-Obj @{ error = @{ message = 'no user' } } }; break }
            '^GET /v1.0/groups$' {
                if ($q -match "mailNickname eq '([^']+)'") { $id = $global:Dst.groupNick[$Matches[1]]; $out = New-Obj @{ value = @($(if ($id) { @{ id = $id; displayName = $Matches[1] } })) } }
                else { $out = New-Obj @{ value = @() } }; break }
            '^POST /v1.0/groups$' {
                $global:Events.Add('group-create')
                $id = 'g-new-' + ($global:Dst.newGroups.Count + 1); $global:Dst.newGroups[$id] = $b
                $global:Dst.members[$id] = [System.Collections.Generic.List[string]]@($b['members@odata.bind'] | ForEach-Object { ($_ -split '/')[-1] })
                $status = 201; $out = New-Obj @{ id = $id; displayName = $b.displayName }; break }
            '^POST /v1.0/groups/([\w-]+)/members/\$ref$' {
                if ($global:Dst.memberFailOnce) { $global:Dst.memberFailOnce = $false; $status = 400; $out = New-Obj @{ error = @{ message = 'The source resource object or one of the objects being referenced don''t exist.' } }; break }
                $global:Dst.members[$Matches[1]].Add(($b['@odata.id'] -split '/')[-1]); $status = 204; break }
            '^GET /v1.0/groups/([\w-]+)/members$' { $out = New-Obj @{ value = @($global:Dst.members[$Matches[1]] | ForEach-Object { @{ id = $_ } }) }; break }
            '^GET /v1.0/groups/([\w-]+)/planner/plans$' { $gid = $Matches[1]
                if ($global:Dst.newGroups.ContainsKey($gid) -and $global:Dst.newPlanFailOnce) { $status = 404; $out = New-Obj @{ error = @{ message = 'group not provisioned for Planner yet' } }; break }
                $out = New-Obj @{ value = @($global:Dst.plans.GetEnumerator() | Where-Object { $_.Value.container.containerId -eq $gid } | ForEach-Object { @{ id = $_.Key; title = $_.Value.title } }) }; break }
            '^GET /v1.0/groups/([\w-]+)$' { $out = New-Obj @{ displayName = "Group $($Matches[1])" }; break }
            '^POST /v1.0/planner/plans$' {
                $gid = $b.container.containerId
                if (-not $allGroups.ContainsKey($gid) -or $b.container.type -ne 'group') { throw "fake: bad container $gid" }
                $global:Events.Add($(if ($b.title -like 'PREFLIGHT*') { "preflight-plan:$gid" } else { 'plan-create' }))
                if ($global:Dst.newGroups.ContainsKey($gid) -and $global:Dst.newPlanFailOnce) { $global:Dst.newPlanFailOnce = $false; $status = 404; $out = New-Obj @{ error = @{ message = 'group not provisioned for Planner yet (forced)' } }; break }
                $id = 'DP' + (++$global:Dst.n); $global:Dst.plans[$id] = $b; $out = New-Obj @{ id = $id; title = $b.title }; break }
            '^GET /v1.0/planner/plans/(DP\d+)$' { $out = New-Obj @{ id = $Matches[1]; '@odata.etag' = 'W/"p"' }; break }
            '^DELETE /v1.0/planner/plans/(DP\d+)$' { $global:Dst.plans.Remove($Matches[1]); $status = 204; break }
            '^GET /v1.0/planner/plans/(DP\d+)/details$' { $out = New-Obj @{ '@odata.etag' = 'W/"pd"'; categoryDescriptions = $global:Dst.planCats[$Matches[1]] }; break }
            '^PATCH /v1.0/planner/plans/(DP\d+)/details$' { $global:Dst.planCats[$Matches[1]] = $b.categoryDescriptions; $status = 204; break }
            '^POST /v1.0/planner/buckets$' {
                $id = 'DB' + (++$global:Dst.n); $hint = 'H{0:D3}' -f $global:Dst.n
                $global:Dst.buckets[$id] = @{ id = $id; name = $b.name; planId = $b.planId; orderHint = $hint; sentHint = $b.orderHint }
                $out = New-Obj $global:Dst.buckets[$id]; break }
            '^GET /v1.0/planner/plans/(DP\d+)/buckets$' { $p = $Matches[1]; $out = New-Obj @{ value = @($global:Dst.buckets.Values | Where-Object { $_.planId -eq $p }) }; break }
            '^POST /v1.0/planner/tasks$' {
                if (-not $global:Dst.throttled) { $global:Dst.throttled = $true; $status = 429; $hdr = @{ 'Retry-After' = @('1') }; $out = New-Obj @{ error = @{ message = 'throttled' } }; break }
                $id = 'DT' + (++$global:Dst.n)
                $b.id = $id; $b.checklistItemCount = 0; $b.referenceCount = 0; if (-not $b.assignments) { $b.assignments = @{} }
                $global:Dst.tasks[$id] = $b; $global:Dst.details[$id] = @{ description = ''; checklist = @{}; references = @{} }
                $out = New-Obj $b; break }
            '^GET /v1.0/planner/plans/(DP\d+)/tasks$' { $p = $Matches[1]; $out = New-Obj @{ value = @($global:Dst.tasks.Values | Where-Object { $_.planId -eq $p }) }; break }
            '^GET /v1.0/planner/tasks/(DT\d+)$' { $out = New-Obj $global:Dst.tasks[$Matches[1]]; break }
            '^GET /v1.0/planner/tasks/(DT\d+)/details$' { $d = $global:Dst.details[$Matches[1]].Clone(); $d['@odata.etag'] = 'W/"td"'; $out = New-Obj $d; break }
            '^PATCH /v1.0/planner/tasks/(DT\d+)/details$' {
                $id = $Matches[1]; $d = $global:Dst.details[$id]
                if ($global:Dst.fail412Once -and $global:Dst.tasks[$id].title -eq 'Write spec') { $global:Dst.fail412Once = $false; $status = 412; $out = New-Obj @{ error = @{ message = 'etag mismatch (forced)' } }; break }
                if ($b.description) { $d.description = $b.description }
                if ($b.checklist) { $b.checklist.GetEnumerator() | ForEach-Object { $d.checklist[$_.Key] = $_.Value } }
                if ($b.references) { $b.references.GetEnumerator() | ForEach-Object { $d.references[$_.Key] = $_.Value } }
                $global:Dst.tasks[$id].checklistItemCount = $d.checklist.Count; $global:Dst.tasks[$id].referenceCount = $d.references.Count
                $status = 204; break }
            # --- Setup (app registration) ---
            '^GET /v1.0/servicePrincipals$' {
                if ($q -match "appId eq '00000003-0000-0000-c000-000000000000'") {
                    $out = New-Obj @{ value = @(@{ id = 'graph-sp'; appRoles = @('Tasks.ReadWrite.All', 'Tasks.Read.All', 'Group.Read.All', 'Group.Create', 'GroupMember.ReadWrite.All', 'User.Read.All', 'Organization.Read.All' | ForEach-Object { @{ id = "r-$_"; value = $_; allowedMemberTypes = @('Application') } }) }) } }
                else { $out = New-Obj @{ value = @() } }; break }
            '^GET /v1.0/applications$' { $out = New-Obj @{ value = @() }; break }
            '^POST /v1.0/applications$' { $global:Dst.apps += , $b; $out = New-Obj @{ id = 'app-obj'; appId = 'app-123' }; break }
            '^POST /v1.0/servicePrincipals$' { $out = New-Obj @{ id = 'client-sp'; appId = $b.appId }; break }
            '^GET /v1.0/servicePrincipals/client-sp/appRoleAssignments$' { $out = New-Obj @{ value = @() }; break }
            '^POST /v1.0/servicePrincipals/graph-sp/appRoleAssignedTo$' {
                if ($global:Dst.grantFailOnce) { $global:Dst.grantFailOnce = $false; $status = 400; $out = New-Obj @{ error = @{ message = 'principal not found yet (forced replication delay)' } }; break }
                $global:Dst.grants += , $b; $status = 201; $out = New-Obj @{ id = 'grant' }; break }
            default { throw "fake dest: unhandled $Method $path ? $q" }
        }
    }
    else { throw 'fake: not connected' }
    return @{ Status = $status; Headers = $hdr; Out = $out }
}
# Compiled cmdlet so -StatusCodeVariable / -ResponseHeadersVariable land in the CALLER's scope, like the real binary cmdlet.
$asm = Add-Type -PassThru -TypeDefinition @'
using System.Collections;
using System.Management.Automation;
namespace FakeGraph {
    public static class Handler { public static ScriptBlock Script; }
    [Cmdlet("Invoke", "MgGraphRequest")]
    public class InvokeFake : PSCmdlet {
        [Parameter] public string Method = "GET";
        [Parameter] public string Uri;
        [Parameter] public object Body;
        [Parameter] public IDictionary Headers;
        [Parameter] public string ContentType;
        [Parameter] public SwitchParameter SkipHttpErrorCheck;
        [Parameter] public string StatusCodeVariable;
        [Parameter] public string ResponseHeadersVariable;
        [Parameter] public string OutputType;
        protected override void ProcessRecord() {
            var r = (Hashtable)Handler.Script.Invoke(Method, Uri, Body, Headers)[0].BaseObject;
            if (!string.IsNullOrEmpty(StatusCodeVariable)) SessionState.PSVariable.Set(StatusCodeVariable, r["Status"]);
            if (!string.IsNullOrEmpty(ResponseHeadersVariable)) SessionState.PSVariable.Set(ResponseHeadersVariable, r["Headers"]);
            if (r["Out"] != null) WriteObject(r["Out"]);
        }
    }
}
'@
Import-Module -Assembly ($asm | Select-Object -First 1).Assembly
[FakeGraph.Handler]::Script = $global:FakeHandler

# ---------- run ----------
$fail = 0
function Check([string]$Name, [bool]$Cond) { if ($Cond) { Write-Host "PASS  $Name" -ForegroundColor Green } else { Write-Host "FAIL  $Name" -ForegroundColor Red; $script:fail++ } }
function Run([string]$Step, [hashtable]$Extra = @{}) { & $script -Step $Step -SettingsPath $settings -NoModuleInstall @Extra }

# ===== EXPORT (one command) =====
Run Export
$gcsv = @(Import-Csv (Join-Path $work 'config\groups.csv'))
Check 'export: discovered the 2 groups that have plans (no groups.csv needed)' ($gcsv.Count -eq 2 -and 'g-empty' -notin $gcsv.sourceGroupId)
$p1 = Get-Content (Join-Path $work 'export\basic\g-src-1\P1.json') -Raw | ConvertFrom-Json
$g2 = Get-Content (Join-Path $work 'export\basic\g-src-2\_group.json') -Raw | ConvertFrom-Json
Check 'export: plan, comments, task chat (deleted message dropped)' ($p1.tasks.Count -eq 2 -and @($p1.tasks[0].comments).Count -eq 2 -and @($p1.tasks[0].chatPosts).Count -eq 1)
Check 'export: group alias, visibility, owners, members captured' ($g2.mailNickname -eq 'ops' -and $g2.visibility -eq 'Public' -and ($g2.ownerIds -join ',') -eq 'u3' -and @($g2.memberIds).Count -eq 2)
$sum = Get-Content (Join-Path $work 'logs\export-summary.txt') -Raw
Check 'export summary: basic counts, task chat count, premium SKIPPED with reason and fix' (
    $sum -match 'Basic plans\s+: 2 plans, 3 tasks' -and $sum -match 'Task chat\s+: 1 messages' -and $sum -match 'Premium plans : SKIPPED: .*Fix: set Source\.DataverseUrl')
$global:LastScopes = $null

# ===== DRY RUN (preview, must write nothing) =====
Run DryRun
$gmDry = @(Import-Csv (Join-Path $work 'logs\group-mapping.csv'))
Check 'dry run: no groups, plans or memberships created; Ops reported as "would be CREATED"' (
    $global:Dst.newGroups.Count -eq 0 -and $global:Dst.plans.Count -eq 0 -and @($global:Dst.members['g-dst-1']).Count -eq 1 -and
    ($gmDry | Where-Object sourceGroupId -eq 'g-src-2').method -eq 'would be CREATED' -and -not (Test-Path (Join-Path $work 'state\targets.json')))

# ===== IMPORT (one command) =====
Run Import
Check 'GCC High write test ran in the EXISTING group before any group was created' (
    $global:Events.IndexOf('preflight-plan:g-dst-1') -ge 0 -and $global:Events.IndexOf('preflight-plan:g-dst-1') -lt $global:Events.IndexOf('group-create'))
Check 'import says loudly that premium plans were not imported, and why' ((Get-Content (Get-ChildItem (Join-Path $work 'logs') -Filter 'migration-*.log').FullName -Raw) -match 'PREMIUM PLANS WERE NOT IMPORTED because the export skipped them')
$um = @(Import-Csv (Join-Path $work 'logs\user-mapping.csv'))
Check 'users auto-mapped by domain (contoso.com -> contoso.us); deleted user reported' (
    ($um | Where-Object sourceUpn -eq 'a@contoso.com').destUpn -eq 'a@contoso.us' -and ($um | Where-Object sourceUpn -eq 'c@contoso.com').method -eq 'domain map' -and @($um | Where-Object method -eq 'NOT FOUND').Count -eq 1)
$gm = @(Import-Csv (Join-Path $work 'logs\group-mapping.csv'))
Check 'groups: Eng matched by alias, Ops created' (($gm | Where-Object sourceGroupId -eq 'g-src-1').method -eq 'same alias' -and ($gm | Where-Object sourceGroupId -eq 'g-src-2').method -eq 'created')
$ng = $global:Dst.newGroups['g-new-1']
Check 'created group: same name/alias/visibility, Unified, owner + members bound with graph.microsoft.us URLs' (
    $ng.displayName -eq 'Ops' -and $ng.mailNickname -eq 'ops' -and $ng.visibility -eq 'Public' -and ($ng.groupTypes -join ',') -eq 'Unified' -and
    ($ng['owners@odata.bind'] -join ',') -eq 'https://graph.microsoft.us/v1.0/users/D-u3' -and @($ng['members@odata.bind']).Count -eq 2)
Check 'missing assignee added to the existing group after a replication retry' ('D-u3' -in $global:Dst.members['g-dst-1'] -and -not $global:Dst.memberFailOnce)
$destPlans = @($global:Dst.plans.Values)
Check 'both plans imported; preflight plan cleaned up; new-group plan created after a not-ready retry' ($destPlans.Count -eq 2 -and -not $global:Dst.newPlanFailOnce -and @($destPlans | Where-Object { $_.title -like 'PREFLIGHT*' }).Count -eq 0)
$t1 = $global:Dst.tasks.Values | Where-Object title -eq 'Write spec'
$t3 = $global:Dst.tasks.Values | Where-Object title -eq 'Patch servers'
$t1d = $global:Dst.details[$t1.id]
Check 'T1 assigned to both mapped users (D-u1, D-u3); T3 to D-u3' ((@($t1.assignments.Keys | Sort-Object) -join ',') -eq 'D-u1,D-u3' -and (@($t3.assignments.Keys) -join ',') -eq 'D-u3')
Check 'retried after a 429 and still created the task' ($global:Dst.throttled)
$bk = @($global:Dst.buckets.Values | Where-Object { $global:Dst.plans[$_.planId].title -eq 'Launch plan' } | Sort-Object { $_.orderHint })
Check 'buckets created in source order with append hints' ($bk[0].name -eq 'First' -and $bk[1].name -eq 'Later' -and $bk[0].sentHint -eq ' !' -and $bk[1].sentHint -eq "$($bk[0].orderHint) !")
Check 'label name copied' (($global:Dst.planCats.Values | Where-Object { $_.category1 }).category1 -eq 'Red label')
Check 'transient 412: comments not moved to sidecar, details left pending' (-not (Test-Path (Join-Path $work 'export\comments-not-imported')) -and -not $t1d.description)

# ===== IMPORT again (resume + idempotent) =====
$writesBefore = $global:Dst.n; $groupsBefore = $global:Dst.newGroups.Count
Run Import
$t1d = $global:Dst.details[$t1.id]
$liveTasks = @($global:Dst.tasks.Values | Where-Object { $global:Dst.plans.ContainsKey($_.planId) })   # ignore tasks of deleted preflight plans
Check 'rerun: no new plans/tasks/groups (preflight plan is created and deleted each run)' ($global:Dst.newGroups.Count -eq $groupsBefore -and $liveTasks.Count -eq 3 -and @($global:Dst.plans.Values).Count -eq 2)
Check 'link rewritten automatically to the GCC High SharePoint host' (@($t1d.references.Keys)[0] -eq 'https%3A//contosogov%2Esharepoint%2Eus/sites/eng/spec%2Edocx')
Check 'rerun finished the pending details: comments + task chat, chronological, deleted chat excluded' (
    $t1d.description -match 'Spec body' -and $t1d.description.IndexOf('Bob: First') -lt $t1d.description.IndexOf('Ann: Second & last') -and
    $t1d.description.IndexOf('Ann: Second & last') -lt $t1d.description.IndexOf('A: Chat hello') -and $t1d.description -notmatch 'deleted one' -and $t1d.checklist.Count -eq 2)
$v = @(Import-Csv (Join-Path $work 'logs\verify-basic.csv'))
Check 'verify: zero FAIL (expected assignees = mapped users)' (@($v | Where-Object result -eq 'FAIL').Count -eq 0 -and $v.Count -gt 10)

# ===== Setup (idempotent app registration) =====
Run Setup @{ Side = 'Destination' }
$cfg = (Get-Content $settings -Raw | ConvertFrom-Json).Destination
$app = $global:Dst.apps[0]
Check 'setup: admin scopes requested' ('Application.ReadWrite.All' -in $global:LastScopes -and 'AppRoleAssignment.ReadWrite.All' -in $global:LastScopes)
Check 'setup: destination app requests the 6 documented permissions as Role' (
    (@($app.requiredResourceAccess[0].resourceAccess | ForEach-Object { "$($_.id):$($_.type)" }) -join ',') -eq 'r-Tasks.ReadWrite.All:Role,r-Group.Read.All:Role,r-Group.Create:Role,r-GroupMember.ReadWrite.All:Role,r-User.Read.All:Role,r-Organization.Read.All:Role')
Check 'setup: certificate public key uploaded' ($app.keyCredentials[0].type -eq 'AsymmetricX509Cert' -and [Convert]::FromBase64String($app.keyCredentials[0].key).Length -gt 500)
Check 'setup: 6 grants after one forced replication retry' ($global:Dst.grants.Count -eq 6)
Check 'setup: settings.json updated (ClientId, TenantId, thumbprint, fallback owner)' ($cfg.ClientId -eq 'app-123' -and $cfg.TenantId -eq 'fake-tenant' -and $cfg.CertificateThumbprint -match '^[0-9A-F]{40}$' -and $cfg.FallbackOwnerUpn -eq 'admin@fake')

Remove-Item $work -Recurse -Force
if ($fail) { Write-Host "$fail check(s) FAILED" -ForegroundColor Red; exit 1 } else { Write-Host 'ALL OFFLINE BASIC E2E CHECKS PASSED' -ForegroundColor Green }

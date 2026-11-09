#Requires -Version 7.4
<#
  Offline end-to-end test of the BASIC path of Migrate-PlannerToGccHigh.ps1:
  Discover -> Export -> ExportTaskChat -> DryRun -> Import -> Import (rerun) -> Verify -> Preflight -> Setup,
  against an in-memory fake of the Graph endpoints the script calls. No tenant needed.
  Proves the script's wiring only; Microsoft's service behavior is what -Step Preflight checks in your tenants.
  Run:  pwsh -File .\tests\Run-OfflineBasicE2E.ps1
#>
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$script = Join-Path $repo 'Migrate-PlannerToGccHigh.ps1'
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("gcch-e2e-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path (Join-Path $work 'config') -Force | Out-Null
@{
    Source = @{ TenantId = 'src'; ClientId = 'c'; CertificateThumbprint = 't'; GraphEnvironment = 'Global' }
    Destination = @{ TenantId = 'dst'; ClientId = 'c'; CertificateThumbprint = 't'; GraphEnvironment = 'USGov' }
    DomainMap = @{}; CommentsMode = 'AppendToDescription'; ExportFolder = 'export'; StateFolder = 'state'; LogFolder = 'logs'; PreflightTestGroupId = 'g-test'
} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $work 'config\settings.json')
"sourceGroupId,destGroupId,sourceDisplayName`ng-src-1,g-dst-1,Eng" | Set-Content (Join-Path $work 'config\groups.csv')
"sourceUpn,destUpn`na@contoso.com,a@contoso.us" | Set-Content (Join-Path $work 'config\users.csv')
"sourcePrefix,destPrefix`nhttps://contoso.sharepoint.com/,https://contoso.sharepoint.us/" | Set-Content (Join-Path $work 'config\urlmap.csv')
$settings = Join-Path $work 'config\settings.json'

# ---------- fake Graph ----------
function global:New-Obj($h) { $h | ConvertTo-Json -Depth 20 | ConvertFrom-Json }
$global:Src = @{
    plans   = @(@{ id = 'P1'; title = 'Launch plan' })
    details = @{ categoryDescriptions = @{ category1 = 'Red label'; category2 = $null } }
    buckets = @(@{ id = 'B1'; name = 'Later'; orderHint = 'b' }, @{ id = 'B2'; name = 'First'; orderHint = 'a' })
    tasks   = @(
        @{ id = 'T1'; title = 'Write spec'; bucketId = 'B1'; percentComplete = 50; priority = 3; conversationThreadId = 'th1'
           appliedCategories = @{ category1 = $true }; checklistItemCount = 2; referenceCount = 1
           assignments = @{ u1 = @{ orderHint = 'x' }; u2 = @{ orderHint = 'y' } } },
        @{ id = 'T2'; title = 'Ship it'; bucketId = 'B2'; percentComplete = 100; priority = 5; checklistItemCount = 0; referenceCount = 0; assignments = @{} })
    taskDetails = @{
        T1 = @{ description = 'Spec body'; previewType = 'automatic'
                checklist = @{ c1 = @{ title = 'one'; isChecked = $true; orderHint = 'a' }; c2 = @{ title = 'two'; isChecked = $false; orderHint = 'b' } }
                references = @{ 'https%3A//contoso%2Esharepoint%2Ecom/sites/eng/spec%2Edocx' = @{ alias = 'Spec'; type = 'Word' } } }
        T2 = @{ description = ''; previewType = 'automatic'; checklist = @{}; references = @{} } }
    posts = @(
        @{ createdDateTime = '2026-10-02T09:00:00Z'; from = @{ emailAddress = @{ name = 'Ann' } }; body = @{ content = '<p>Second &amp; last</p>' } },
        @{ createdDateTime = '2026-10-01T09:00:00Z'; from = @{ emailAddress = @{ name = 'Bob' } }; body = @{ content = '<div>First</div>' } })
    chat = @{
        T1 = @(@{ id = 'm1'; content = '<div>Chat hello</div>'; createdDateTime = '2026-10-03T09:00:00Z'; deletedTime = $null; createdBy = @{ user = @{ id = 'u1' } } },
               @{ id = 'm2'; content = 'deleted one'; createdDateTime = '2026-10-04T09:00:00Z'; deletedTime = '2026-10-05T00:00:00Z'; createdBy = @{ user = @{ id = 'u9' } } })
        T2 = @() }
    users = @{ u1 = @{ id = 'u1'; userPrincipalName = 'a@contoso.com'; displayName = 'A' } }   # u2 / u9 = deleted users (404)
}
$global:Dst = @{ users = @{ 'a@contoso.us' = 'D-u1' }; members = @('D-u1'); plans = @{}; planCats = @{}; buckets = @{}; tasks = @{}; details = @{}; n = 0
                 throttled = $false; fail412Once = $true; failDelete = $false; apps = @(); sps = @(); grants = @(); grantFailOnce = $true }
$global:FakeMode = $null; $global:LastScopes = $null

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
    $b = if ($Body) { $Body | ConvertFrom-Json -AsHashtable } else { $null }
    $status = 200; $hdr = @{}; $out = $null
    if ($Method -in 'PATCH', 'DELETE' -and -not $Headers['If-Match']) { throw "fake: $Method without If-Match" }
    if ($global:FakeMode -eq 'Global') {
        switch -Regex ($path) {
            '^/v1.0/groups$' { $out = New-Obj @{ value = @(@{ id = 'g-src-1'; displayName = 'Eng'; mail = 'eng@contoso.com' }, @{ id = 'g-empty'; displayName = 'No plans'; mail = 'np@contoso.com' }) } }
            '^/v1.0/groups/g-empty/planner/plans$' { $out = New-Obj @{ value = @() } }
            '^/v1.0/groups/g-src-1/planner/plans$' { $out = New-Obj @{ value = $global:Src.plans } }
            '^/v1.0/planner/plans/P1/details$' { $out = New-Obj $global:Src.details }
            '^/v1.0/planner/plans/P1/buckets$' { $out = New-Obj @{ value = $global:Src.buckets } }
            '^/v1.0/planner/plans/P1/tasks$' { $out = New-Obj @{ value = $global:Src.tasks } }
            '^/v1.0/planner/tasks/(T\d)/details$' { $out = New-Obj $global:Src.taskDetails[$Matches[1]] }
            '^/beta/planner/tasks/(T\d)/messages$' {
                if (-not $global:LastScopes -or 'Tasks.Read' -notin $global:LastScopes) { $status = 403; $out = New-Obj @{ error = @{ message = 'task chat needs delegated Tasks.Read' } } }
                else { $out = New-Obj @{ value = @($global:Src.chat[$Matches[1]]) } } }
            '^/v1.0/groups/g-src-1/threads/th1/posts$' { $out = New-Obj @{ value = $global:Src.posts } }
            '^/v1.0/users/(.+)$' { if ($global:Src.users[$Matches[1]]) { $out = New-Obj $global:Src.users[$Matches[1]] } else { $status = 404; $out = New-Obj @{ error = @{ message = 'gone' } } } }
            default { throw "fake source: unhandled $Method $path" }
        }
    }
    elseif ($global:FakeMode -eq 'USGov') {
        switch -Regex ("$Method $path") {
            '^GET /v1.0/users/(.+)$' { $id = $global:Dst.users[$Matches[1]]; if ($id) { $out = New-Obj @{ id = $id } } else { $status = 404; $out = New-Obj @{ error = @{ message = 'no user' } } } }
            '^GET /v1.0/groups/g-dst-1/members$' { $out = New-Obj @{ value = @($global:Dst.members | ForEach-Object { @{ id = $_ } }) } }
            '^GET /v1.0/groups/(g-[\w-]+)$' { $out = New-Obj @{ displayName = "Group $($Matches[1])" } }
            '^GET /v1.0/planner/plans/(DP\d+)$' { $out = New-Obj @{ id = $Matches[1]; '@odata.etag' = 'W/"p"' } }
            '^DELETE /v1.0/planner/plans/(DP\d+)$' {
                if ($global:Dst.failDelete) { $status = 403; $out = New-Obj @{ error = @{ message = 'delete denied (forced)' } }; break }
                $global:Dst.plans.Remove($Matches[1]); $status = 204 }
            '^GET /v1.0/planner/tasks/(DT\d+)$' { $out = New-Obj $global:Dst.tasks[$Matches[1]] }
            '^POST /v1.0/planner/plans$' {
                if ($b.container.containerId -notin 'g-dst-1', 'g-test' -or $b.container.type -ne 'group') { throw 'fake: bad container' }
                $id = 'DP' + (++$global:Dst.n); $global:Dst.plans[$id] = $b; $out = New-Obj @{ id = $id; title = $b.title } }
            '^GET /v1.0/planner/plans/(DP\d+)/details$' { $out = New-Obj @{ '@odata.etag' = 'W/"pd"'; categoryDescriptions = $global:Dst.planCats[$Matches[1]] } }
            '^PATCH /v1.0/planner/plans/(DP\d+)/details$' { $global:Dst.planCats[$Matches[1]] = $b.categoryDescriptions; $status = 204 }
            '^POST /v1.0/planner/buckets$' {
                $id = 'DB' + (++$global:Dst.n); $hint = 'H{0:D3}' -f $global:Dst.n
                $global:Dst.buckets[$id] = @{ id = $id; name = $b.name; planId = $b.planId; orderHint = $hint; sentHint = $b.orderHint }
                $out = New-Obj $global:Dst.buckets[$id] }
            '^GET /v1.0/planner/plans/(DP\d+)/buckets$' { $p = $Matches[1]; $out = New-Obj @{ value = @($global:Dst.buckets.Values | Where-Object { $_.planId -eq $p }) } }
            '^POST /v1.0/planner/tasks$' {
                if (-not $global:Dst.throttled) { $global:Dst.throttled = $true; $status = 429; $hdr = @{ 'Retry-After' = @('1') }; $out = New-Obj @{ error = @{ message = 'throttled' } }; break }
                $id = 'DT' + (++$global:Dst.n)
                $b.id = $id; $b.checklistItemCount = 0; $b.referenceCount = 0; if (-not $b.assignments) { $b.assignments = @{} }
                $global:Dst.tasks[$id] = $b; $global:Dst.details[$id] = @{ description = ''; checklist = @{}; references = @{} }
                $out = New-Obj $b }
            '^GET /v1.0/planner/plans/(DP\d+)/tasks$' { $p = $Matches[1]; $out = New-Obj @{ value = @($global:Dst.tasks.Values | Where-Object { $_.planId -eq $p }) } }
            '^GET /v1.0/planner/tasks/(DT\d+)/details$' { $d = $global:Dst.details[$Matches[1]].Clone(); $d['@odata.etag'] = 'W/"td"'; $out = New-Obj $d }
            '^PATCH /v1.0/planner/tasks/(DT\d+)/details$' {
                $id = $Matches[1]; $d = $global:Dst.details[$id]
                if ($global:Dst.fail412Once -and $global:Dst.tasks[$id].title -eq 'Write spec') {
                    $global:Dst.fail412Once = $false; $status = 412; $out = New-Obj @{ error = @{ message = 'etag mismatch (forced)' } }; break }
                if ($b.description) { $d.description = $b.description }
                if ($b.checklist) { $b.checklist.GetEnumerator() | ForEach-Object { $d.checklist[$_.Key] = $_.Value } }
                if ($b.references) { $b.references.GetEnumerator() | ForEach-Object { $d.references[$_.Key] = $_.Value } }
                $global:Dst.tasks[$id].checklistItemCount = $d.checklist.Count; $global:Dst.tasks[$id].referenceCount = $d.references.Count
                $status = 204 }
            # --- Setup step (app registration) ---
            '^GET /v1.0/servicePrincipals$' {
                $out = New-Obj @{ value = @(@{ id = 'graph-sp'; appRoles = @(
                    @{ id = 'r-trw'; value = 'Tasks.ReadWrite.All'; allowedMemberTypes = @('Application') },
                    @{ id = 'r-tr'; value = 'Tasks.Read.All'; allowedMemberTypes = @('Application') },
                    @{ id = 'r-g'; value = 'Group.Read.All'; allowedMemberTypes = @('Application') },
                    @{ id = 'r-u'; value = 'User.Read.All'; allowedMemberTypes = @('Application') }) }) } }
            '^POST /v1.0/applications$' { $global:Dst.apps += , $b; $out = New-Obj @{ id = 'app-obj'; appId = 'app-123' } }
            '^POST /v1.0/servicePrincipals$' { $global:Dst.sps += , $b; $out = New-Obj @{ id = 'client-sp'; appId = $b.appId } }
            '^POST /v1.0/servicePrincipals/graph-sp/appRoleAssignedTo$' {
                if ($global:Dst.grantFailOnce) { $global:Dst.grantFailOnce = $false; $status = 400; $out = New-Obj @{ error = @{ message = 'principal not found yet (forced replication delay)' } }; break }
                $global:Dst.grants += , $b; $status = 201; $out = New-Obj @{ id = 'grant' } }
            default { throw "fake dest: unhandled $Method $path" }
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
function Run([string]$Step, [hashtable]$Extra = @{}) { & $script -Step $Step -SettingsPath $settings @Extra }

Run Discover
$disc = @(Import-Csv (Join-Path $work 'config\groups.discovered.csv'))
Check 'discover: only the group that has plans is listed (groups.csv kept)' ($disc.Count -eq 1 -and $disc[0].sourceGroupId -eq 'g-src-1' -and $disc[0].planCount -eq '1')

Run Export
$exported = Get-Content (Join-Path $work 'export\basic\g-src-1\P1.json') -Raw | ConvertFrom-Json
Check 'export: 2 tasks, 2 buckets, 2 comment posts on T1' ($exported.tasks.Count -eq 2 -and $exported.buckets.Count -eq 2 -and @($exported.tasks[0].comments).Count -eq 2)

Run ExportTaskChat
$exported = Get-Content (Join-Path $work 'export\basic\g-src-1\P1.json') -Raw | ConvertFrom-Json
Check 'task chat: delegated Tasks.Read sign-in, 1 live message kept, deleted one dropped' ('Tasks.Read' -in $global:LastScopes -and @($exported.tasks[0].chatPosts).Count -eq 1 -and $exported.tasks[0].chatPosts[0].from.emailAddress.name -eq 'A')
$global:LastScopes = $null

Run DryRun
Check 'dry run writes nothing' ($global:Dst.plans.Count -eq 0 -and $global:Dst.tasks.Count -eq 0)

Run Import
$writesBefore = $global:Dst.n
$t1first = $global:Dst.tasks.Values | Where-Object title -eq 'Write spec'
Check 'transient 412 on details: comments NOT moved to sidecar, details left pending' (
    -not (Test-Path (Join-Path $work 'export\comments-not-imported')) -and -not $global:Dst.details[$t1first.id].description -and
    @(Import-Csv (Join-Path $work 'logs\basic-import-problems.csv') | Where-Object issue -like '*HTTP 412*will retry*').Count -eq 1)
Run Import   # rerun must not duplicate and must finish the pending details
Check 'rerun performed zero creates' ($global:Dst.n -eq $writesBefore)
Check 'one plan, two buckets, two tasks after rerun (idempotent)' ($global:Dst.plans.Count -eq 1 -and $global:Dst.buckets.Count -eq 2 -and $global:Dst.tasks.Count -eq 2)
Check 'retried after a 429 and still created the task' ($global:Dst.throttled)
$bk = @($global:Dst.buckets.Values | Sort-Object { $_.orderHint })
Check 'buckets created in source order (First, Later) with append hints' ($bk[0].name -eq 'First' -and $bk[1].name -eq 'Later' -and $bk[0].sentHint -eq ' !' -and $bk[1].sentHint -eq "$($bk[0].orderHint) !")
$t1 = $global:Dst.tasks.Values | Where-Object title -eq 'Write spec'
$t1d = $global:Dst.details[$t1.id]
Check 'T1 assigned only to mapped user D-u1' ((@($t1.assignments.Keys) -join ',') -eq 'D-u1')
Check 'T1 keeps percentComplete/priority/label' ($t1.percentComplete -eq 50 -and $t1.priority -eq 3 -and $t1.appliedCategories.category1 -eq $true)
Check 'T1 checklist 2, link rewritten to sharepoint.us' ($t1d.checklist.Count -eq 2 -and (@($t1d.references.Keys)[0] -eq 'https%3A//contoso%2Esharepoint%2Eus/sites/eng/spec%2Edocx'))
Check 'T1 description: comments + task chat, chronological, deleted chat excluded' (
    $t1d.description -match 'Spec body' -and $t1d.description.IndexOf('Bob: First') -lt $t1d.description.IndexOf('Ann: Second & last') -and
    $t1d.description.IndexOf('Ann: Second & last') -lt $t1d.description.IndexOf('A: Chat hello') -and $t1d.description -notmatch 'deleted one')
Check 'label name copied' ($global:Dst.planCats['DP1'].category1 -eq 'Red label')
$problems = Import-Csv (Join-Path $work 'logs\basic-import-problems.csv')
Check 'deleted source user reported as unmapped' (@($problems | Where-Object issue -like 'assignee not mapped: id:u2*').Count -ge 1)

Run Verify
$v = Import-Csv (Join-Path $work 'logs\verify-basic.csv')
$unexpected = @($v | Where-Object { $_.result -eq 'FAIL' -and -not ($_.item -eq 'Write spec' -and $_.check -eq 'assignee count') })
Check 'verify: only the expected assignee-count FAIL (u2 unmapped)' ($unexpected.Count -eq 0 -and @($v | Where-Object result -eq 'FAIL').Count -eq 1)

$global:Dst.failDelete = $true
$plansBefore = $global:Dst.plans.Count
Run Preflight @{ Scope = 'Basic' }
$pf = @(Import-Csv (Join-Path $work 'logs\preflight.csv'))
function Get-Status([string]$Check) { ($pf | Where-Object check -like "$Check*" | Select-Object -First 1).status }
Check 'preflight: canary plan/bucket/task/details all PASS' ((Get-Status 'create plan') -eq 'PASS' -and (Get-Status 'create bucket') -eq 'PASS' -and (Get-Status 'create task') -eq 'PASS' -and (Get-Status 'update task details') -eq 'PASS')
Check 'preflight: comment-thread sampling + post read' ((($pf | Where-Object check -eq 'sampled tasks with a comment thread').detail -eq '1 of 2') -and (Get-Status 'read comment posts') -eq 'PASS')
Check 'preflight: forced delete failure recorded as FAIL, script completed' ((Get-Status 'delete canary plan') -eq 'FAIL' -and $global:Dst.plans.Count -eq $plansBefore + 1)
Check 'preflight: users.csv and dest group resolve' ((Get-Status 'users.csv destUpn resolve') -eq 'PASS' -and (Get-Status 'group g-dst-1') -eq 'PASS')

Run Setup @{ Side = 'Destination' }
$cfg = (Get-Content $settings -Raw | ConvertFrom-Json).Destination
$app = $global:Dst.apps[0]
Check 'setup: admin scopes requested' ('Application.ReadWrite.All' -in $global:LastScopes -and 'AppRoleAssignment.ReadWrite.All' -in $global:LastScopes)
Check 'setup: app requests exactly Tasks.ReadWrite.All, Group.Read.All, User.Read.All as Role' (
    (@($app.requiredResourceAccess[0].resourceAccess | ForEach-Object { "$($_.id):$($_.type)" }) -join ',') -eq 'r-trw:Role,r-g:Role,r-u:Role' -and $app.requiredResourceAccess[0].resourceAppId -eq '00000003-0000-0000-c000-000000000000')
Check 'setup: certificate public key uploaded' ($app.keyCredentials[0].type -eq 'AsymmetricX509Cert' -and $app.keyCredentials[0].usage -eq 'Verify' -and [Convert]::FromBase64String($app.keyCredentials[0].key).Length -gt 500)
Check 'setup: 3 grants after one forced replication retry' ($global:Dst.grants.Count -eq 3 -and ($global:Dst.grants | Where-Object { $_.principalId -ne 'client-sp' -or $_.resourceId -ne 'graph-sp' }).Count -eq 0)
Check 'setup: settings.json updated (ClientId, TenantId, thumbprint)' ($cfg.ClientId -eq 'app-123' -and $cfg.TenantId -eq 'fake-tenant' -and $cfg.CertificateThumbprint -match '^[0-9A-F]{40}$')

Remove-Item $work -Recurse -Force
if ($fail) { Write-Host "$fail check(s) FAILED" -ForegroundColor Red; exit 1 } else { Write-Host 'ALL OFFLINE BASIC E2E CHECKS PASSED' -ForegroundColor Green }

#Requires -Version 7.4
<#
  Offline end-to-end test of the PREMIUM path: 4-Export-Premium -> 5-Import-Premium (run twice) against an
  in-memory fake of the Dataverse Web API + Project schedule API actions.
  The fake metadata below is deliberately partly invented (e.g. label/checklist column names) to prove the
  scripts read names from metadata instead of assuming them. The fake also enforces that every
  @odata.bind target already exists or was queued earlier in the same OperationSet, and <=200 ops per set.
  It proves wiring only - Microsoft's real behavior is checked by 0-Preflight.ps1 -Scope Premium.
#>
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("gcch-pe2e-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path (Join-Path $work 'config') -Force | Out-Null
@{
    Source = @{ TenantId = 'src'; AzEnvironment = 'AzureCloud'; DataverseUrl = 'https://src.crm.dynamics.com/' }
    Destination = @{ TenantId = 'dst'; AzEnvironment = 'AzureUSGovernment'; DataverseUrl = 'https://dst.crm.microsoftdynamics.us/' }
    DomainMap = @{}; ExportFolder = 'export'; StateFolder = 'state'; LogFolder = 'logs'
    Premium = @{ ExtraColumns = @{ msdyn_projecttask = @('msdyn_description', 'msdyn_progress') }; OperationSetChunkSize = 3; WaitForOperationSetMinutes = 1 }
} | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $work 'config\settings.json')
$settings = Join-Path $work 'config\settings.json'
"sourceUpn,destUpn`na@contoso.com,a@contoso.us" | Set-Content (Join-Path $work 'config\users.csv')

# ---------- fake metadata: table -> set, pk, name, columns (name=creatable?), lookups (attr=entity) ----------
$global:Meta = @{
    msdyn_project               = @{ set = 'msdyn_projects'; pk = 'msdyn_projectid'; name = 'msdyn_subject'; cols = @{ msdyn_projectid = 1; msdyn_subject = 1 }; lk = @{} }
    msdyn_projectbucket         = @{ set = 'msdyn_projectbuckets'; pk = 'msdyn_projectbucketid'; name = 'msdyn_name'; cols = @{ msdyn_projectbucketid = 1; msdyn_name = 1; msdyn_project = 1 }; lk = @{ msdyn_project = 'msdyn_project' } }
    msdyn_projecttask           = @{ set = 'msdyn_projecttasks'; pk = 'msdyn_projecttaskid'; name = 'msdyn_subject'
                                     cols = @{ msdyn_projecttaskid = 1; msdyn_subject = 1; msdyn_project = 1; msdyn_projectbucket = 1; msdyn_parenttask = 1; msdyn_outlinelevel = 1; msdyn_scheduledstart = 1; msdyn_scheduledend = 1; msdyn_start = 1; msdyn_effort = 1; msdyn_description = 1; msdyn_progress = 0; createdon = 0 }
                                     lk = @{ msdyn_project = 'msdyn_project'; msdyn_projectbucket = 'msdyn_projectbucket'; msdyn_parenttask = 'msdyn_projecttask' } }
    msdyn_projectteam           = @{ set = 'msdyn_projectteams'; pk = 'msdyn_projectteamid'; name = 'msdyn_name'; cols = @{ msdyn_projectteamid = 1; msdyn_name = 1; msdyn_project = 1; msdyn_resourceid = 1 }; lk = @{ msdyn_project = 'msdyn_project'; msdyn_resourceid = 'bookableresource' } }
    msdyn_resourceassignment    = @{ set = 'msdyn_resourceassignments'; pk = 'msdyn_resourceassignmentid'; name = 'msdyn_name'; cols = @{ msdyn_resourceassignmentid = 1; msdyn_name = 1; msdyn_taskid = 1; msdyn_projectteamid = 1; msdyn_projectid = 1 }; lk = @{ msdyn_taskid = 'msdyn_projecttask'; msdyn_projectteamid = 'msdyn_projectteam'; msdyn_projectid = 'msdyn_project' } }
    msdyn_projecttaskdependency = @{ set = 'msdyn_projecttaskdependencies'; pk = 'msdyn_projecttaskdependencyid'; name = 'msdyn_name'; cols = @{ msdyn_projecttaskdependencyid = 1; msdyn_project = 1; msdyn_predecessortask = 1; msdyn_successortask = 1; msdyn_linktype = 1 }; lk = @{ msdyn_project = 'msdyn_project'; msdyn_predecessortask = 'msdyn_projecttask'; msdyn_successortask = 'msdyn_projecttask' } }
    msdyn_projectchecklist      = @{ set = 'msdyn_projectchecklists'; pk = 'msdyn_projectchecklistid'; name = 'msdyn_name'; cols = @{ msdyn_projectchecklistid = 1; msdyn_name = 1; msdyn_projecttask = 1; msdyn_completed = 1; createdon = 0; ownerid = 1 }; lk = @{ msdyn_projecttask = 'msdyn_projecttask' } }
    msdyn_projectlabel          = @{ set = 'msdyn_projectlabels'; pk = 'msdyn_projectlabelid'; name = 'msdyn_projectlabeltext'; cols = @{ msdyn_projectlabelid = 1; msdyn_projectlabeltext = 1; msdyn_project = 1 }; lk = @{ msdyn_project = 'msdyn_project' } }
    msdyn_projecttasktolabel    = @{ set = 'msdyn_projecttasktolabels'; pk = 'msdyn_projecttasktolabelid'; name = 'msdyn_name'; cols = @{ msdyn_projecttasktolabelid = 1; msdyn_project = 1; msdyn_projecttask = 1; msdyn_projectlabel = 1 }; lk = @{ msdyn_project = 'msdyn_project'; msdyn_projecttask = 'msdyn_projecttask'; msdyn_projectlabel = 'msdyn_projectlabel' } }
}
function global:New-Row([string]$Table, [hashtable]$Props, [hashtable]$Lookups = @{}) {
    $r = [ordered]@{}; $Props.GetEnumerator() | ForEach-Object { $r[$_.Key] = $_.Value }
    foreach ($k in $Lookups.Keys) { $r["_$($k)_value"] = $Lookups[$k]; $r["_$($k)_value@Microsoft.Dynamics.CRM.lookuplogicalname"] = $global:Meta[$Table].lk[$k] }
    $r
}
$global:SrcRows = @{
    msdyn_projects                = @(New-Row msdyn_project @{ msdyn_projectid = 'p1'; msdyn_subject = 'Premium launch' })
    msdyn_projectbuckets          = @(New-Row msdyn_projectbucket @{ msdyn_projectbucketid = 'b1'; msdyn_name = 'Phase 1' } @{ msdyn_project = 'p1' })
    msdyn_projecttasks            = @(  # deliberately child-first to prove parent-first ordering
        New-Row msdyn_projecttask @{ msdyn_projecttaskid = 't3'; msdyn_subject = 'Build'; msdyn_outlinelevel = 3; msdyn_effort = 8; msdyn_description = 'deep'; msdyn_progress = 0.2; createdon = '2026-01-01' } @{ msdyn_project = 'p1'; msdyn_projectbucket = 'b1'; msdyn_parenttask = 't2' }
        New-Row msdyn_projecttask @{ msdyn_projecttaskid = 't2'; msdyn_subject = 'Plan'; msdyn_outlinelevel = 2 } @{ msdyn_project = 'p1'; msdyn_projectbucket = 'b1'; msdyn_parenttask = 't1' }
        New-Row msdyn_projecttask @{ msdyn_projecttaskid = 't1'; msdyn_subject = 'Design'; msdyn_outlinelevel = 1; msdyn_description = 'top' } @{ msdyn_project = 'p1'; msdyn_projectbucket = 'b1' })
    msdyn_projectteams            = @(New-Row msdyn_projectteam @{ msdyn_projectteamid = 'tm1'; msdyn_name = 'Jane Doe' } @{ msdyn_project = 'p1' })
    msdyn_resourceassignments     = @(New-Row msdyn_resourceassignment @{ msdyn_resourceassignmentid = 'ra1'; msdyn_name = 'Jane on Plan' } @{ msdyn_taskid = 't2'; msdyn_projectteamid = 'tm1'; msdyn_projectid = 'p1' })
    msdyn_projecttaskdependencies = @(New-Row msdyn_projecttaskdependency @{ msdyn_projecttaskdependencyid = 'd1'; msdyn_linktype = 192350000 } @{ msdyn_project = 'p1'; msdyn_predecessortask = 't1'; msdyn_successortask = 't3' })
    msdyn_projectchecklists       = @(New-Row msdyn_projectchecklist @{ msdyn_projectchecklistid = 'cl1'; msdyn_name = 'Review'; msdyn_completed = $true; createdon = '2026-01-01'; ownerid = 'someone' } @{ msdyn_projecttask = 't1' })
    msdyn_projectlabels           = @(New-Row msdyn_projectlabel @{ msdyn_projectlabelid = 'l1'; msdyn_projectlabeltext = 'Urgent' } @{ msdyn_project = 'p1' })
    msdyn_projecttasktolabels     = @(New-Row msdyn_projecttasktolabel @{ msdyn_projecttasktolabelid = 'tl1' } @{ msdyn_project = 'p1'; msdyn_projecttask = 't1'; msdyn_projectlabel = 'l1' })
}
$global:Dst = @{ rows = @{}; payloads = @{}; opsets = @{}; pss = 0; createProject = 0 }
$global:DvMode = $null

function global:Connect-AzAccount { param($Environment, $Tenant, $Scope) $global:DvMode = $Environment }
function global:Get-AzAccessToken { param($ResourceUrl, [switch]$AsSecureString)
    [pscustomobject]@{ Token = (ConvertTo-SecureString 'fake' -AsPlainText -Force); ExpiresOn = [DateTimeOffset]::Now.AddHours(1) } }
$global:Login = [System.Collections.Generic.List[object]]::new(); $global:LastAuth = @{}; $global:PendingOnce = $true
function global:Invoke-WebRequest {
    param($Method = 'GET', $Uri, $Headers, $Body, $ContentType, [switch]$SkipHttpErrorCheck)
    # --- Entra device-code endpoints (OAuth 2.0 device authorization grant) ---
    if ($Uri -match '^https://login\.microsoftonline\.(us|com)/([^/]+)/oauth2/v2\.0/(devicecode|token)$') {
        $form = @{}; ([string]$Body -split '&') | ForEach-Object { $kv = $_ -split '=', 2; $form[$kv[0]] = [uri]::UnescapeDataString($kv[1]) }
        $global:Login.Add([pscustomobject]@{ cloud = $Matches[1]; tenant = $Matches[2]; kind = $Matches[3]; form = $form })
        $o = if ($Matches[3] -eq 'devicecode') { @{ device_code = 'dc1'; user_code = 'ABCD'; message = 'To sign in, enter code ABCD (fake)'; interval = 1; expires_in = 60 } }
             elseif ($global:PendingOnce) { $global:PendingOnce = $false; @{ error = 'authorization_pending' } }
             else { @{ access_token = 'tok-dc'; expires_in = 3600; refresh_token = 'rt' } }
        $st = if ($o.error) { 400 } else { 200 }
        return [pscustomobject]@{ StatusCode = $st; Content = ($o | ConvertTo-Json); Headers = @{} }
    }
    $side = if ($Uri.StartsWith('https://src.crm.dynamics.com/api/data/v9.2/')) { 'src' } elseif ($Uri.StartsWith('https://dst.crm.microsoftdynamics.us/api/data/v9.2/')) { 'dst' } else { throw "fake dv: unexpected host $Uri" }
    $global:LastAuth[$side] = $Headers.Authorization
    $rel = $Uri.Substring($Uri.IndexOf('/api/data/v9.2/') + '/api/data/v9.2/'.Length)
    $b = if ($Body) { [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json -AsHashtable } else { $null }
    $status = 200; $out = $null
    switch -Regex ("$Method $rel") {
        '^GET WhoAmI$' { $out = @{ UserId = "user-$side" }; break }
        "^GET EntityDefinitions\(LogicalName='(\w+)'\)/Attributes" { $m = $global:Meta[$Matches[1]]; $out = @{ value = @($m.cols.GetEnumerator() | ForEach-Object { @{ LogicalName = $_.Key; AttributeType = 'String'; IsValidForCreate = [bool]$_.Value } }) }; break }
        "^GET EntityDefinitions\(LogicalName='(\w+)'\)/ManyToOneRelationships" { $m = $global:Meta[$Matches[1]]; $out = @{ value = @($m.lk.GetEnumerator() | ForEach-Object { @{ ReferencingAttribute = $_.Key; ReferencedEntity = $_.Value; ReferencingEntityNavigationPropertyName = $_.Key } }) }; break }
        "^GET EntityDefinitions\(LogicalName='(\w+)'\)\?" { $m = $global:Meta[$Matches[1]]; if ($m) { $out = @{ EntitySetName = $m.set; PrimaryIdAttribute = $m.pk; PrimaryNameAttribute = $m.name } } else { $status = 404; $out = @{ error = @{ message = 'no table' } } }; break }
        '^GET (\w+)\?\$filter=_(\w+)_value eq ([\w-]+)$' { $set = $Matches[1]; $a = $Matches[2]; $v = $Matches[3]; $out = @{ value = @($global:SrcRows[$set] | Where-Object { $_["_$($a)_value"] -eq $v }) }; break }
        "^GET (\w+)\?\`$filter=Microsoft\.Dynamics\.CRM\.In\(PropertyName='(\w+)',PropertyValues=\[(.*)\]\)$" {
            $set = $Matches[1]; $a = $Matches[2]; $vals = $Matches[3] -split ',' | ForEach-Object { $_.Trim("'") }
            $out = @{ value = @($global:SrcRows[$set] | Where-Object { $_["_$($a)_value"] -in $vals }) }; break }
        '^GET (\w+)\?\$select=\w+&\$count=true$' { $out = @{ '@odata.count' = @($global:SrcRows[$Matches[1]]).Count; value = @() }; break }
        "^GET systemusers\?\`$select=systemuserid&\`$filter=domainname eq 'a@contoso\.us'$" { $out = @{ value = @(@{ systemuserid = 'su1' }) }; break }
        '^GET bookableresources\?\$select=bookableresourceid&\$filter=_userid_value eq su1$' { $out = @{ value = @(@{ bookableresourceid = 'br1' }) }; break }
        '^GET (\w+)$' { $out = @{ value = @($global:SrcRows[$Matches[1]]) }; break }
        '^GET (\w+)\(([\w-]+)\)\?\$select=' { if ($global:Dst.rows.ContainsKey($Matches[2])) { $out = @{ id = $Matches[2] } } else { $status = 404; $out = @{ error = @{ message = 'not found' } } }; break }
        '^POST msdyn_CreateProjectV1$' { $global:Dst.createProject++; $id = if ($b.Project.msdyn_projectid) { $b.Project.msdyn_projectid } else { [guid]::NewGuid().ToString() }; $global:Dst.rows[$id] = 'msdyn_project'; $global:Dst.payloads[$id] = $b.Project; $out = @{ ProjectId = $id }; break }
        '^POST msdyn_CreateTeamMemberV1$' {
            $tm = $b.TeamMember
            if (-not ($tm['msdyn_project@odata.bind'] -match '\(([\w-]+)\)$' -and $global:Dst.rows.ContainsKey($Matches[1]))) { $status = 400; $out = @{ error = @{ message = 'team member: project missing' } }; break }
            $global:Dst.rows[$tm.msdyn_projectteamid] = 'msdyn_projectteam'; $global:Dst.payloads[$tm.msdyn_projectteamid] = $tm; $out = @{ TeamMemberId = $tm.msdyn_projectteamid }; break }
        '^POST msdyn_CreateOperationSetV1$' { $os = [guid]::NewGuid().ToString(); $global:Dst.opsets[$os] = [System.Collections.Generic.List[object]]::new(); $out = @{ OperationSetId = $os }; break }
        '^POST msdyn_PssCreateV1$' {
            $global:Dst.pss++
            $q = $global:Dst.opsets[$b.OperationSetId]; $e = $b.Entity
            if ($q.Count -ge 200) { $status = 400; $out = @{ error = @{ message = 'more than 200 operations' } }; break }
            $table = $e['@odata.type'] -replace '^Microsoft\.Dynamics\.CRM\.', ''
            $pk = $global:Meta[$table].pk
            foreach ($k in @($e.Keys | Where-Object { $_ -like '*@odata.bind' })) {
                $target = ($e[$k] -replace '^/\w+\(', '') -replace '\)$', ''
                if (-not ($global:Dst.rows.ContainsKey($target) -or @($q | Where-Object { $_.id -eq $target }).Count)) { $status = 400; $out = @{ error = @{ message = "$table.$k -> $target does not exist yet" } }; break }
            }
            if ($status -ne 200) { break }
            $q.Add(@{ id = $e[$pk]; table = $table; payload = $e }); $out = @{ OperationSetResponse = '{}' }; break }
        '^POST msdyn_ExecuteOperationSetV1$' { foreach ($x in $global:Dst.opsets[$b.OperationSetId]) { $global:Dst.rows[$x.id] = $x.table; $global:Dst.payloads[$x.id] = $x.payload }; $status = 200; $out = @{}; break }
        '^POST msdyn_AbandonOperationSetV1$' { $out = @{}; break }
        default { throw "fake dv: unhandled $Method $rel" }
    }
    [pscustomobject]@{ StatusCode = $status; Content = $(if ($null -ne $out) { $out | ConvertTo-Json -Depth 20 } else { '' }); Headers = @{} }
}

$fail = 0
function Check([string]$Name, [bool]$Cond) { if ($Cond) { Write-Host "PASS  $Name" -ForegroundColor Green } else { Write-Host "FAIL  $Name" -ForegroundColor Red; $script:fail++ } }

$script = Join-Path $repo 'Migrate-PlannerToGccHigh.ps1'
function Run([string]$Step, [hashtable]$Extra = @{}) { & $script -Step $Step -SettingsPath $settings @Extra }

Run ExportPremium
$ex =Get-Content (Join-Path $work 'export\premium\p1.json') -Raw | ConvertFrom-Json
Check 'export: 3 tasks, checklist found via task-lookup In() query' (@($ex.tables.msdyn_projecttask).Count -eq 3 -and @($ex.tables.msdyn_projectchecklist).Count -eq 1)
$labels = @(Import-Csv (Join-Path $work 'export\premium\labels-manual.csv'))
Check 'labels-manual.csv uses the discovered label name column' ($labels.Count -eq 1 -and $labels[0].label -eq 'Urgent' -and $labels[0].task -eq 'Design')

Run ImportPremium
$r = $global:Dst.rows
Check 'all 9 rows created (project, bucket, 3 tasks, team, assignment, dependency, checklist)' (@('p1', 'b1', 't1', 't2', 't3', 'tm1', 'ra1', 'd1', 'cl1' | Where-Object { $r.ContainsKey($_) }).Count -eq 9)
$p = $global:Dst.payloads
Check 'source GUIDs reused as destination IDs' ($r['p1'] -eq 'msdyn_project')
Check 'child task bound to parent, without outline level' ($p['t2']['msdyn_parenttask@odata.bind'] -eq '/msdyn_projecttasks(t1)' -and -not $p['t2'].Contains('msdyn_outlinelevel'))
Check 'top task keeps outline level' ($p['t1']['msdyn_outlinelevel'] -eq 1)
Check 'verified extra column included, non-creatable one skipped' ($p['t1']['msdyn_description'] -eq 'top' -and -not $p['t3'].Contains('msdyn_progress') -and -not $p['t3'].Contains('createdon'))
Check 'checklist via metadata: task bound, system/non-creatable columns excluded' ($p['cl1']['msdyn_projecttask@odata.bind'] -eq '/msdyn_projecttasks(t1)' -and $p['cl1']['msdyn_completed'] -eq $true -and -not $p['cl1'].Contains('ownerid') -and -not $p['cl1'].Contains('createdon'))
Check 'assignment bound to task, team member, project' ($p['ra1']['msdyn_taskid@odata.bind'] -eq '/msdyn_projecttasks(t2)' -and $p['ra1']['msdyn_projectteamid@odata.bind'] -eq '/msdyn_projectteams(tm1)')
$probPath = Join-Path $work 'logs\premium-import-problems.csv'
$probs = if (Test-Path $probPath) { @(Import-Csv $probPath) } else { @() }
Check 'no PssCreate rejections or missing rows' (@($probs | Where-Object { $_.issue -match 'rejected|not present' }).Count -eq 0)

$pssBefore = $global:Dst.pss; $cpBefore = $global:Dst.createProject
Run ImportPremium
Check 'rerun: zero PssCreate and zero CreateProject calls (idempotent)' ($global:Dst.pss -eq $pssBefore -and $global:Dst.createProject -eq $cpBefore)

Run Preflight @{ Scope = 'Premium' }
$pf = @(Import-Csv (Join-Path $work 'logs\preflight.csv'))
function Get-Status([string]$Check) { ($pf | Where-Object check -like "$Check*" | Select-Object -First 1).status }
Check 'preflight: source sign-in + project count' ((Get-Status 'sign in + WhoAmI (AzureCloud)') -eq 'PASS' -and ($pf | Where-Object check -eq 'count premium plans').detail -like '1 projects*')
Check 'preflight: destination canary (project, team member, operation set read-back) PASS' (
    (Get-Status 'msdyn_CreateProjectV1') -eq 'PASS' -and (Get-Status 'msdyn_CreateTeamMemberV1') -eq 'PASS' -and (Get-Status 'OperationSet') -eq 'PASS')
Check 'preflight: named bookable resource lookup' ((Get-Status 'bookable resource for a@contoso.us') -eq 'PASS')

# Device-code fallback for the GCC High Dataverse sign-in (no Az.Accounts dependency).
$j = Get-Content $settings -Raw | ConvertFrom-Json
$j.Destination | Add-Member -NotePropertyName DataverseAuth -NotePropertyValue 'DeviceCode' -Force
$j.Destination | Add-Member -NotePropertyName DataverseClientId -NotePropertyValue 'pc-1' -Force
$j | ConvertTo-Json -Depth 8 | Set-Content $settings
$global:LastAuth = @{}
Run Preflight @{ Scope = 'Premium' }
$pf = @(Import-Csv (Join-Path $work 'logs\preflight.csv'))
$dc = @($global:Login | Where-Object kind -eq 'devicecode')
$tok = @($global:Login | Where-Object kind -eq 'token')
Check 'device code: US Gov authority, destination tenant, user_impersonation + offline_access scope' (
    $dc.Count -eq 1 -and $dc[0].cloud -eq 'us' -and $dc[0].tenant -eq 'dst' -and $dc[0].form.client_id -eq 'pc-1' -and
    $dc[0].form.scope -eq 'https://dst.crm.microsoftdynamics.us/user_impersonation offline_access')
Check 'device code: polled through authorization_pending, then used the issued token' ($tok.Count -eq 2 -and $global:LastAuth['dst'] -eq 'Bearer tok-dc')
Check 'device code: destination canary still PASS' ((Get-Status 'sign in + WhoAmI (AzureUSGovernment)') -eq 'PASS' -and (Get-Status 'msdyn_CreateProjectV1') -eq 'PASS' -and (Get-Status 'OperationSet') -eq 'PASS')

Remove-Item $work -Recurse -Force
if ($fail) { Write-Host "$fail check(s) FAILED" -ForegroundColor Red; exit 1 } else { Write-Host 'ALL OFFLINE PREMIUM E2E CHECKS PASSED' -ForegroundColor Green }

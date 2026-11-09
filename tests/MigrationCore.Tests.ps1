# Pester tests for the network-free functions in Migrate-PlannerToGccHigh.ps1 (dot-sourced: functions only, no step runs).
# Compatible with Pester 3.4 (in-box on Windows). Run:  Invoke-Pester .\tests\MigrationCore.Tests.ps1
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here '..\Migrate-PlannerToGccHigh.ps1')

Describe 'Task chat -> comment text (Graph beta plannerTaskChatMessage)' {
    $msgs = @(
        [pscustomobject]@{ content = '<div>Chat <b>hello</b></div>'; createdDateTime = '2026-10-03T09:00:00Z'; deletedTime = $null; createdBy = [pscustomobject]@{ user = [pscustomobject]@{ id = 'u1' } } },
        [pscustomobject]@{ content = 'gone'; createdDateTime = '2026-10-04T09:00:00Z'; deletedTime = '2026-10-05T00:00:00Z'; createdBy = [pscustomobject]@{ user = [pscustomobject]@{ id = 'u1' } } },
        [pscustomobject]@{ content = 'who?'; createdDateTime = '2026-10-06T09:00:00Z'; deletedTime = $null; createdBy = [pscustomobject]@{ user = [pscustomobject]@{ id = 'u9' } } })
    $posts = @(ConvertTo-CommentPost -Messages $msgs -NameById @{ u1 = 'Ann' })
    It 'skips deleted messages' { $posts.Count | Should Be 2 }
    It 'names known authors and marks unknown ones' { ($posts | ForEach-Object { $_.from.emailAddress.name }) -join ',' | Should Be 'Ann,unknown' }
    It 'merges with legacy posts chronologically and ignores nulls' {
        $legacy = [pscustomobject]@{ createdDateTime = '2026-10-01T09:00:00Z'; from = [pscustomobject]@{ emailAddress = [pscustomobject]@{ name = 'Bob' } }; body = [pscustomobject]@{ content = 'old comment' } }
        $t = Format-CommentsText -Posts (@($legacy) + $posts + @($null))
        $t.IndexOf('Bob: old comment') | Should BeLessThan $t.IndexOf('Ann: Chat hello')
    }
}

Describe 'OAuth helpers' {
    It 'uses the US Government authority for AzureUSGovernment' { Get-AuthorityHost 'AzureUSGovernment' | Should Be 'https://login.microsoftonline.us' }
    It 'uses the global authority for AzureCloud' { Get-AuthorityHost 'AzureCloud' | Should Be 'https://login.microsoftonline.com' }
    It 'rejects unknown clouds' {
        # Explicit try/catch: Pester 3.4's bare 'Should Throw' misreports even { throw 'x' } on PowerShell 7.6.
        $msg = $null; try { Get-AuthorityHost 'Mars' | Out-Null } catch { $msg = $_.Exception.Message }
        $msg | Should Match 'Unsupported AzEnvironment'
    }
    It 'form-encodes values' { ConvertTo-FormBody @{ scope = 'https://x.us/user_impersonation offline_access'; a = '1' } | Should Be 'a=1&scope=https%3A%2F%2Fx.us%2Fuser_impersonation%20offline_access' }
}

Describe 'Planner reference keys (Graph docs: plannerExternalReferences)' {
    It 'encodes exactly like the Microsoft Learn example' {
        ConvertTo-PlannerRefKey 'https://contoso.sharepoint.com/teams/agile/documents/AnnualReport.pptx' |
            Should Be 'https%3A//contoso%2Esharepoint%2Ecom/teams/agile/documents/AnnualReport%2Epptx'
    }
    It 'decodes the Learn update-task-details example key' {
        ConvertFrom-PlannerRefKey 'http%3A//developer%2Emicrosoft%2Ecom' | Should Be 'http://developer.microsoft.com'
    }
    It 'round-trips URLs that already contain %20, @ and #' {
        $u = 'https://x.sharepoint.com/sites/a/My%20Doc.docx?web=1#p@2'
        ConvertFrom-PlannerRefKey (ConvertTo-PlannerRefKey $u) | Should Be $u
    }
}

Describe 'Get-MappedUrl' {
    $map = @(
        [pscustomobject]@{ sourcePrefix = 'https://contoso.sharepoint.com/'; destPrefix = 'https://contoso.sharepoint.us/' },
        [pscustomobject]@{ sourcePrefix = 'https://contoso.sharepoint.com/sites/eng/'; destPrefix = 'https://contoso.sharepoint.us/sites/engineering/' })
    It 'uses the longest matching prefix' {
        Get-MappedUrl 'https://contoso.sharepoint.com/sites/eng/doc.docx' $map | Should Be 'https://contoso.sharepoint.us/sites/engineering/doc.docx'
    }
    It 'is case-insensitive on the prefix' {
        Get-MappedUrl 'HTTPS://CONTOSO.SHAREPOINT.COM/x' $map | Should Be 'https://contoso.sharepoint.us/x'
    }
    It 'leaves unmapped URLs unchanged' {
        Get-MappedUrl 'https://example.org/a' $map | Should Be 'https://example.org/a'
    }
}

Describe 'Resolve-DestUpn' {
    $um = @{ 'jane@contoso.com' = 'jdoe@contoso.us' }
    $dm = @{ 'contoso.com' = 'contoso.us' }
    It 'prefers users.csv' { Resolve-DestUpn 'Jane@Contoso.com' $um $dm | Should Be 'jdoe@contoso.us' }
    It 'falls back to the domain map' { Resolve-DestUpn 'bob@contoso.com' $um $dm | Should Be 'bob@contoso.us' }
    It 'returns null when nothing maps' { Resolve-DestUpn 'x@fabrikam.com' $um $dm | Should BeNullOrEmpty }
    It 'returns null for empty input' { Resolve-DestUpn '' $um $dm | Should BeNullOrEmpty }
}

Describe 'Order hints (Graph docs: Using order hints in Planner)' {
    It 'sorts ordinally, space before digits before letters, shorter prefix first' {
        $items = @('adhg', 'ab', '5637', ' 5637!', 'a') | ForEach-Object { [pscustomobject]@{ orderHint = $_ } }
        ((Get-SortedByOrderHint -Items $items) | ForEach-Object { $_.orderHint }) -join '|' | Should Be ' 5637!|5637|a|ab|adhg'
    }
    It 'builds append hints per the documented "<prev> <next>!" format' {
        Get-AppendOrderHint $null | Should Be ' !'
        Get-AppendOrderHint 'adhg' | Should Be 'adhg !'
    }
}

Describe 'Comments as text' {
    It 'strips HTML and decodes entities' {
        ConvertTo-PlainText '<html><body><div>Hello &amp; <b>bye</b></div><p>line2</p></body></html>' | Should Be "Hello & bye`nline2"
    }
    It 'orders posts by time and labels the author' {
        $posts = @(
            [pscustomobject]@{ createdDateTime = '2026-01-02T10:00:00Z'; from = [pscustomobject]@{ emailAddress = [pscustomobject]@{ name = 'B' } }; body = [pscustomobject]@{ content = 'second' } },
            [pscustomobject]@{ createdDateTime = '2026-01-01T10:00:00Z'; from = [pscustomobject]@{ emailAddress = [pscustomobject]@{ name = 'A' } }; body = [pscustomobject]@{ content = 'first' } })
        $t = Format-CommentsText $posts
        $t.IndexOf('A: first') | Should BeLessThan $t.IndexOf('B: second')
        $t | Should Match 'Comments migrated'
    }
    It 'returns empty for no posts' { Format-CommentsText @() | Should Be '' }
}

Describe 'Get-RetryDelaySeconds' {
    It 'honors Retry-After from a dictionary' { Get-RetryDelaySeconds @{ 'Retry-After' = @('7') } 1 | Should Be 7 }
    It 'honors Retry-After from KeyValuePair headers' {
        $h = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string, string[]]]]::new()
        $h.Add([System.Collections.Generic.KeyValuePair[string, string[]]]::new('Retry-After', [string[]]@('9')))
        Get-RetryDelaySeconds $h 1 | Should Be 9
    }
    It 'falls back to exponential backoff capped at 120' {
        Get-RetryDelaySeconds $null 3 | Should Be 8
        Get-RetryDelaySeconds $null 10 | Should Be 120
    }
}

Describe 'Split-IntoChunks' {
    It 'splits 5 items into 2+2+1' {
        $c = Split-IntoChunks -Items @(1, 2, 3, 4, 5) -Size 2
        $c.Count | Should Be 3
        ($c[2] -join ',') | Should Be '5'
    }
}

Describe 'Get-PlannerAssignmentsBody' {
    $src = [pscustomobject]@{ 'u1' = [pscustomobject]@{ orderHint = 'x' }; 'u2' = [pscustomobject]@{ orderHint = 'y' } }
    $r = Get-PlannerAssignmentsBody -SourceAssignments $src -SourceUserUpn @{ u1 = 'a@contoso.com'; u2 = 'b@fabrikam.com' } `
        -DestUserIdByUpn @{ 'a@contoso.us' = 'd1' } -DomainMap @{ 'contoso.com' = 'contoso.us' }
    It 'maps resolvable assignees to destination IDs' { @($r.Body.Keys) -join ',' | Should Be 'd1' }
    It 'uses the plannerAssignment odata type' { $r.Body['d1']['@odata.type'] | Should Be '#microsoft.graph.plannerAssignment' }
    It 'reports unmapped assignees' { $r.Unmapped -join ',' | Should Be 'b@fabrikam.com' }
}

Describe 'Get-TaskDetailsBody' {
    $details = [pscustomobject]@{
        description = 'Do the thing'; previewType = 'checklist'
        checklist   = [pscustomobject]@{ k2 = [pscustomobject]@{ title = 'second'; isChecked = $false; orderHint = 'b' }; k1 = [pscustomobject]@{ title = 'first'; isChecked = $true; orderHint = 'a' } }
        references  = [pscustomobject]@{ 'https%3A//contoso%2Esharepoint%2Ecom/sites/a/f%2Edocx' = [pscustomobject]@{ alias = 'Spec'; type = 'Word' } }
    }
    $map = @([pscustomobject]@{ sourcePrefix = 'https://contoso.sharepoint.com/'; destPrefix = 'https://contoso.sharepoint.us/' })
    $b = Get-TaskDetailsBody -SourceDetails $details -UrlMap $map -CommentsText "`n--c--"
    It 'appends comments to the description' { $b.description | Should Be "Do the thing`n--c--" }
    It 'creates new checklist keys and keeps items' { $b.checklist.Count | Should Be 2 }
    It 'rewrites and re-encodes reference URLs' {
        @($b.references.Keys)[0] | Should Be 'https%3A//contoso%2Esharepoint%2Eus/sites/a/f%2Edocx'
        $b.references.Values[0].alias | Should Be 'Spec'
    }
    It 'copies previewType' { $b.previewType | Should Be 'checklist' }
}

Describe 'Invoke-Graph fail-safe when the status variable is not populated' {
    # A script-function stand-in cannot set the caller's -StatusCodeVariable (unlike the real binary cmdlet),
    # so this exercises the fallback that inspects the Graph error payload.
    It 'throws on a Graph error payload instead of returning it as success' {
        function global:Invoke-MgGraphRequest { param($Method, $Uri, $Body, $Headers, $ContentType, [switch]$SkipHttpErrorCheck, $StatusCodeVariable, $ResponseHeadersVariable, $OutputType)
            [pscustomobject]@{ error = [pscustomobject]@{ code = 'Forbidden'; message = 'nope' } } }
        { Invoke-Graph -Uri '/v1.0/x' } | Should Throw 'nope'
        Remove-Item function:\Invoke-MgGraphRequest
    }
    It 'retries a TooManyRequests payload, then returns the success' {
        $global:calls = 0
        function global:Invoke-MgGraphRequest { param($Method, $Uri, $Body, $Headers, $ContentType, [switch]$SkipHttpErrorCheck, $StatusCodeVariable, $ResponseHeadersVariable, $OutputType)
            $global:calls++
            if ($global:calls -eq 1) { [pscustomobject]@{ error = [pscustomobject]@{ code = 'TooManyRequests'; message = 'slow down' } } }
            else { [pscustomobject]@{ id = 'ok' } } }
        Mock Start-Sleep { }
        (Invoke-Graph -Uri '/v1.0/x').id | Should Be 'ok'
        $global:calls | Should Be 2
        Remove-Item function:\Invoke-MgGraphRequest
    }
}

Describe 'New-DvEntityPayload' {
    $ti = [pscustomobject]@{
        LogicalName = 'msdyn_projecttask'; PrimaryId = 'msdyn_projecttaskid'; PrimaryName = 'msdyn_subject'
        CreatableColumns = @('msdyn_projecttaskid', 'msdyn_subject', 'msdyn_project', 'msdyn_parenttask', 'msdyn_effort', 'createdon')
        Lookups = @(
            [pscustomobject]@{ ReferencingAttribute = 'msdyn_project'; ReferencedEntity = 'msdyn_project'; ReferencingEntityNavigationPropertyName = 'msdyn_project' },
            [pscustomobject]@{ ReferencingAttribute = 'msdyn_parenttask'; ReferencedEntity = 'msdyn_projecttask'; ReferencingEntityNavigationPropertyName = 'msdyn_parenttask' })
    }
    $row = [pscustomobject]@{
        msdyn_projecttaskid = 't1'; msdyn_subject = 'Task 1'; msdyn_effort = 4.0; msdyn_progress = 0.5; createdon = '2026-01-01'
        '_msdyn_project_value' = 'p1'; '_msdyn_project_value@Microsoft.Dynamics.CRM.lookuplogicalname' = 'msdyn_project'
        '_msdyn_parenttask_value' = 't0'; '_msdyn_parenttask_value@Microsoft.Dynamics.CRM.lookuplogicalname' = 'msdyn_projecttask'
    }
    $sets = @{ msdyn_project = 'msdyn_projects'; msdyn_projecttask = 'msdyn_projecttasks' }
    $cols = @('msdyn_subject', 'msdyn_project', 'msdyn_parenttask', 'msdyn_effort', 'msdyn_progress')

    It 'reuses the source primary key and sets the odata type' {
        $r = New-DvEntityPayload -Row $row -TableInfo $ti -SetNameByEntity $sets -IdMap @{ p1 = 'P1'; t0 = 't0' } -Columns $cols
        $r.Payload['msdyn_projecttaskid'] | Should Be 't1'
        $r.Payload['@odata.type'] | Should Be 'Microsoft.Dynamics.CRM.msdyn_projecttask'
    }
    It 'rebinds lookups to mapped destination IDs' {
        $r = New-DvEntityPayload -Row $row -TableInfo $ti -SetNameByEntity $sets -IdMap @{ p1 = 'P1'; t0 = 't0' } -Columns $cols
        $r.Payload['msdyn_project@odata.bind'] | Should Be '/msdyn_projects(P1)'
        $r.Payload['msdyn_parenttask@odata.bind'] | Should Be '/msdyn_projecttasks(t0)'
    }
    It 'drops lookups whose target was not migrated, and says so' {
        $r = New-DvEntityPayload -Row $row -TableInfo $ti -SetNameByEntity $sets -IdMap @{ p1 = 'P1' } -Columns $cols
        $r.Payload.Contains('msdyn_parenttask@odata.bind') | Should Be $false
        ($r.Dropped -join ';') | Should Match 'msdyn_parenttask'
    }
    It 'skips columns that are not creatable in the destination' {
        $r = New-DvEntityPayload -Row $row -TableInfo $ti -SetNameByEntity $sets -IdMap @{ p1 = 'P1'; t0 = 't0' } -Columns $cols
        $r.Payload.Contains('msdyn_progress') | Should Be $false
        ($r.Dropped -join ';') | Should Match 'msdyn_progress\(not creatable'
    }
    It 'AllCreatable excludes system columns' {
        $r = New-DvEntityPayload -Row $row -TableInfo $ti -SetNameByEntity $sets -IdMap @{ p1 = 'P1'; t0 = 't0' } -AllCreatable
        $r.Payload.Contains('createdon') | Should Be $false
        $r.Payload['msdyn_effort'] | Should Be 4.0
    }
}

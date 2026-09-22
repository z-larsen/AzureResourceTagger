BeforeAll {
    $modulePath = Join-Path (Split-Path $PSScriptRoot) 'ResourceTagger.Core.psm1'
    $module = Import-Module $modulePath -Force -PassThru
    $context = [pscustomobject]@{
        Environment  = [pscustomobject]@{ Name = 'AzureCloud' }
        Tenant       = [pscustomobject]@{ Id = '00000000-0000-0000-0000-000000000001' }
        Subscription = [pscustomobject]@{ Id = '00000000-0000-0000-0000-000000000002' }
        Account      = [pscustomobject]@{ Id = 'tester@example.invalid' }
    }
    $scope = New-TagScope -Context $context
    $resourceId = "/subscriptions/$($scope.SubscriptionId)/resourceGroups/example/providers/Microsoft.Storage/storageAccounts/example"
    $target = [pscustomobject]@{ Id = $resourceId; Name = 'example'; Kind = 'storageAccounts' }

    # Stub the Azure boundary so tests cannot autoload or call real Azure cmdlets.
    & $module {
        function script:Get-AzTag {
            param($ResourceId, $DefaultProfile, $ErrorAction)
            throw 'Unmocked Azure read in test'
        }
        function script:Update-AzTag {
            param($ResourceId, $Tag, $Operation, $DefaultProfile, $ErrorAction)
            throw 'Unmocked Azure write in test'
        }
    }
}

Describe 'Scope identity and target boundaries' {
    It 'matches the same identity independently of dropdown indices' {
        Test-TagScopeMatch -Expected $scope -Actual (New-TagScope -Context $context) | Should -BeTrue
    }

    It 'rejects a change to <Field>' -TestCases @(
        @{ Field = 'Environment'; Value = 'AzureUSGovernment' }
        @{ Field = 'TenantId'; Value = 'different-tenant' }
        @{ Field = 'SubscriptionId'; Value = 'different-subscription' }
        @{ Field = 'AccountId'; Value = 'another@example.invalid' }
        @{ Field = 'ResourceGroup'; Value = 'another-group' }
    ) {
        param($Field, $Value)
        $changed = $scope.PSObject.Copy()
        $changed.$Field = $Value
        Test-TagScopeMatch -Expected $scope -Actual $changed | Should -BeFalse
        { Assert-TagScope -Expected $scope -Actual $changed } | Should -Throw '*Re-scan*'
    }

    It 'does not retain mutable context identity properties' {
        $changedContext = [pscustomobject]@{
            Environment = [pscustomobject]@{ Name = 'AzureCloud' }
            Tenant = [pscustomobject]@{ Id = $scope.TenantId }
            Subscription = [pscustomobject]@{ Id = $scope.SubscriptionId }
            Account = [pscustomobject]@{ Id = $scope.AccountId }
        }
        $snapshot = New-TagScope -Context $changedContext
        $changedContext.Subscription.Id = 'changed'
        $snapshot.SubscriptionId | Should -Be $scope.SubscriptionId
    }

    It 'rejects missing context identity instead of treating it as a valid scope' {
        { New-TagScope -Context ([pscustomobject]@{}) } | Should -Throw
        Test-TagScopeMatch -Expected $null -Actual $scope | Should -BeFalse
    }

    It 'checks resource group boundaries without prefix collisions' {
        $groupScope = New-TagScope -Context $context -ResourceGroup 'example'
        { Assert-TagTargetScope -ResourceId $resourceId -Scope $groupScope } | Should -Not -Throw
        { Assert-TagTargetScope -ResourceId ($resourceId -replace '/example/', '/example-other/') -Scope $groupScope } |
            Should -Throw '*outside*'
    }

    It 'rejects targets from a different subscription' {
        { Assert-TagTargetScope -ResourceId ($resourceId -replace $scope.SubscriptionId, 'another-subscription') -Scope $scope } |
            Should -Throw '*outside*'
    }
}

Describe 'Local inventory filters' {
    BeforeAll {
        $resources = @(
            [pscustomobject]@{ name = 'tagged'; id = 'one'; type = 'Microsoft.Test/widgets'; resourceGroup = 'rg'; tags = @{ Owner = 'Alice' } },
            [pscustomobject]@{ name = 'untagged'; id = 'two'; type = 'Microsoft.Test/widgets'; resourceGroup = 'rg'; tags = @{} }
        )
        $groups = @(
            [pscustomobject]@{ name = 'one'; location = 'test'; tags = @{ Owner = 'Alice' } },
            [pscustomobject]@{ name = 'two'; location = 'test'; tags = @{} }
        )
    }

    It 'normalizes object and dictionary tags with case-insensitive keys' {
        (ConvertTo-TagHashtable ([pscustomobject]@{ Owner = 'Alice' }))['owner'] | Should -BeExactly 'Alice'
        (ConvertTo-TagHashtable @{ Owner = 'Alice' })['OWNER'] | Should -BeExactly 'Alice'
        (ConvertTo-TagHashtable $null).Count | Should -Be 0
    }

    It 'rejects unsupported tag data rather than silently treating it as untagged' {
        { ConvertTo-TagHashtable 'not a tag dictionary' } | Should -Throw
    }

    It 'handles zero, one, and several required tags' {
        @(Get-RequiredTagName -Text '').Count | Should -Be 0
        @(Get-RequiredTagName -Text ' Owner ').Count | Should -Be 1
        @(Get-RequiredTagName -Text 'Owner, Environment, owner, ,').Count | Should -Be 2
        @(Get-ResourceGroupRow -Resources $groups -RequiredTags @(Get-RequiredTagName 'Owner') -OnlyMissing).Count |
            Should -Be 1
    }

    It 'restores all rows after filtering without changing the backing inventory' {
        @(Get-ResourceRow -Resources $resources -Filter Untagged).Count | Should -Be 1
        @(Get-ResourceRow -Resources $resources -Filter All).Count | Should -Be 2
        $resources.Count | Should -Be 2
    }

    It 'filters missing tag names case-insensitively and responds to a new name' {
        @(Get-ResourceRow -Resources $resources -Filter MissingTag -TagName 'owner').Count | Should -Be 1
        @(Get-ResourceRow -Resources $resources -Filter MissingTag -TagName 'Environment').Count | Should -Be 2
    }

    It 'shows all resources while a missing-tag filter has no tag name yet' {
        @(Get-ResourceRow -Resources $resources -Filter MissingTag -TagName '').Count | Should -Be 2
    }

    It 'recomputes group gaps against the current required-tag input' {
        @(Get-ResourceGroupRow -Resources $groups -RequiredTags @('Owner') -OnlyMissing).Count | Should -Be 1
        @(Get-ResourceGroupRow -Resources $groups -RequiredTags @('Environment') -OnlyMissing).Count | Should -Be 2
        @(Get-ResourceGroupRow -Resources $groups -RequiredTags @('Owner')).Count | Should -Be 2
    }
}

Describe 'Shared tag change planning' {
    It 'skips existing tags when overwrite is disabled and sends only new keys' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice'; CostCenter = 'A' } `
            -Tags @{ Owner = 'Bob'; Environment = 'Test' }
        $plan.Payload.Count | Should -Be 1
        $plan.Payload['Environment'] | Should -BeExactly 'Test'
        ($plan.Changes | Where-Object TagName -eq Owner).Action | Should -Be 'Skip'
        ($plan.Changes | Where-Object TagName -eq Owner).PreviousValue | Should -BeExactly 'Alice'
    }

    It 'skips unchanged values even when overwrite is enabled' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = 'Alice' } -Overwrite
        $plan.Status | Should -Be 'Skipped'
        $plan.Payload.Count | Should -Be 0
    }

    It 'plans an overwrite that changes only value casing' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Environment = 'Prod' } -Tags @{ Environment = 'prod' } -Overwrite
        $plan.Payload['Environment'] | Should -BeExactly 'prod'
        $plan.Changes[0].Action | Should -Be 'Change'
    }

    It 'preserves empty values and whitespace as actual values' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{} -Tags @{ Empty = ''; Spaced = ' prod ' }
        $plan.Payload['Empty'] | Should -BeExactly ''
        $plan.Payload['Spaced'] | Should -BeExactly ' prod '
    }

    It 'uses case-sensitive matching for removal values' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Environment = 'Prod' } `
            -Tags @{ Environment = 'prod' } -Operation Delete -MatchValue
        $plan.Status | Should -Be 'Skipped'
        $plan.Payload.Count | Should -Be 0
    }

    It 'distinguishes an exact empty-value match from any value' {
        $exact = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = '' } -Operation Delete -MatchValue
        $any = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = '' } -Operation Delete
        $empty = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = '' } -Tags @{ Owner = '' } -Operation Delete -MatchValue
        $exact.Status | Should -Be 'Skipped'
        $any.Payload['Owner'] | Should -BeExactly 'Alice'
        $empty.Status | Should -Be 'Planned'
        $empty.Payload['Owner'] | Should -BeExactly ''
    }

    It 'never plans removal of an absent key' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{} -Tags @{ Owner = '' } -Operation Delete
        $plan.Status | Should -Be 'Skipped'
    }

    It 'evaluates untagged and first-missing-tag scopes using current tags' {
        $untagged = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Other = 'value' } `
            -Tags @{ Owner = 'Alice' } -Selection Untagged
        $missing = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } `
            -Tags @{ Owner = 'Bob'; Environment = 'Test' } -Selection MissingTag -SelectionTagName Owner
        $untagged.Status | Should -Be 'Skipped'
        $missing.Status | Should -Be 'Skipped'
    }

    It 'exposes preview actions and old/new values without exposing context objects' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } `
            -Tags @{ Owner = 'Bob'; NewTag = '' } -Overwrite
        $rows = @(Get-TagPlanRow -Plan @($plan))
        $rows.Count | Should -Be 2
        ($rows | Where-Object TagName -eq Owner).Before | Should -BeExactly 'Alice'
        ($rows | Where-Object TagName -eq Owner).After | Should -BeExactly 'Bob'
        ($rows | Where-Object TagName -eq NewTag).Before | Should -Be '(absent)'
        ($rows | Where-Object TagName -eq NewTag).After | Should -Be '(empty)'
        $rows[0].PSObject.Properties.Name | Should -Not -Contain 'DefaultProfile'
    }
}

Describe 'Spreadsheet-safe CSV rows' {
    It 'neutralizes <Case> as visible text without modifying the input' -TestCases @(
        @{ Case = 'equals'; Value = '=1+1=2' }
        @{ Case = 'plus'; Value = '+1+1' }
        @{ Case = 'minus'; Value = '-1+1' }
        @{ Case = 'at sign'; Value = '@SUM(A1:A2)' }
        @{ Case = 'space-prefixed formula'; Value = '  =1+1' }
        @{ Case = 'apostrophe-prefixed formula'; Value = "'=1+1" }
        @{ Case = 'tab-prefixed formula'; Value = "`t=1+1" }
        @{ Case = 'carriage-return-prefixed formula'; Value = "`r=1+1" }
        @{ Case = 'line-feed-prefixed formula'; Value = "`n=1+1" }
        @{ Case = 'null-prefixed formula'; Value = ([string][char]0 + '=1+1') }
        @{ Case = 'BOM-prefixed formula'; Value = ([string][char]0xFEFF + '=1+1') }
        @{ Case = 'full-width equals'; Value = ([string][char]0xFF1D + '1+1') }
        @{ Case = 'full-width plus'; Value = ([string][char]0xFF0B + '1+1') }
        @{ Case = 'full-width minus'; Value = ([string][char]0xFF0D + '1+1') }
        @{ Case = 'full-width at sign'; Value = ([string][char]0xFF20 + 'SUM(A1:A2)') }
        @{ Case = 'leading tab'; Value = "`tordinary" }
        @{ Case = 'leading carriage return'; Value = "`rordinary" }
        @{ Case = 'leading line feed'; Value = "`nordinary" }
    ) {
        param($Case, $Value)
        $original = [pscustomobject]@{ Tags = $Value }
        $safe = $original | ConvertTo-SpreadsheetSafeRow
        $safe.Tags | Should -BeExactly ('[text] ' + $Value)
        $original.Tags | Should -BeExactly $Value
        ($safe | ConvertTo-SpreadsheetSafeRow).Tags | Should -BeExactly $safe.Tags
    }

    It 'protects every text column and preserves column order and numeric values' {
        $original = [pscustomobject][ordered]@{
            Scope = '=1'; Name = '+1'; Location = '-1'; ResourceGroup = '@name'
            Type = '=type'; TagCount = 2; Tags = '=1+1=2'; Empty = ''; Optional = $null
        }
        $safe = $original | ConvertTo-SpreadsheetSafeRow
        foreach ($name in @('Scope', 'Name', 'Location', 'ResourceGroup', 'Type', 'Tags')) {
            $safe.$name | Should -BeExactly ('[text] ' + $original.$name)
        }
        ($safe.PSObject.Properties.Name -join ',') | Should -BeExactly ($original.PSObject.Properties.Name -join ',')
        $safe.TagCount | Should -BeOfType ([int])
        $safe.TagCount | Should -Be 2
        $safe.Empty | Should -BeExactly ''
        $safe.Optional | Should -BeNullOrEmpty
    }

    It 'preserves ordinary values and lets CSV serialization escape quotes and delimiters' {
        $text = "Owner=first,second;`"quoted`"`nsecond line"
        $rows = @(
            [pscustomobject]@{ Name = 'normal'; Tags = $text },
            [pscustomobject]@{ Name = 'formula'; Tags = '=1+1";,=1+1' }
        )
        $roundTrip = @($rows | ConvertTo-SpreadsheetSafeRow | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)
        $roundTrip.Count | Should -Be 2
        @($roundTrip[0].PSObject.Properties).Count | Should -Be 2
        $roundTrip[0].Name | Should -BeExactly 'normal'
        $roundTrip[0].Tags | Should -BeExactly $text
        $roundTrip[1].Tags | Should -BeExactly '[text] =1+1";,=1+1'
    }
}

Describe 'Azure read/write boundary and conflict handling' {
    BeforeEach {
        Mock Get-AzTag -ModuleName ResourceTagger.Core {
            [pscustomobject]@{ Properties = [pscustomobject]@{ TagsProperty = @{ Owner = 'Alice' } } }
        }
        Mock Update-AzTag -ModuleName ResourceTagger.Core {}
    }

    It 'only reads during planning and always uses the captured profile' {
        $plan = @(Get-TagOperationPlan -Targets @($target) -Scope $scope -DefaultProfile $context -Tags @{ Owner = 'Bob' })
        $plan[0].Status | Should -Be 'Skipped'
        Should -Invoke Get-AzTag -ModuleName ResourceTagger.Core -Times 1 -Exactly -ParameterFilter {
            $ResourceId -eq $target.Id -and $DefaultProfile -eq $context
        }
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0
    }

    It 'surfaces read failures as errors instead of planning writes against an empty tag set' {
        Mock Get-AzTag -ModuleName ResourceTagger.Core { throw 'Access denied' }
        $plan = @(Get-TagOperationPlan -Targets @($target) -Scope $scope -DefaultProfile $context -Tags @{ Owner = 'Bob' })
        $plan[0].Status | Should -Be 'Error'
        $plan[0].Detail | Should -Match 'Access denied'
        $plan[0].Payload.Count | Should -Be 0
    }

    It 'rejects malformed ARM response shape <Shape> without a write plan' -TestCases @(
        @{ Shape = 'Null' },
        @{ Shape = 'MissingProperties' },
        @{ Shape = 'NullProperties' },
        @{ Shape = 'MissingTags' }
    ) {
        param($Shape)
        Mock Get-AzTag -ModuleName ResourceTagger.Core {
            switch ($Shape) {
                'Null' { $null }
                'MissingProperties' { [pscustomobject]@{ Unexpected = 'value' } }
                'NullProperties' { [pscustomobject]@{ Properties = $null } }
                'MissingTags' { [pscustomobject]@{ Properties = [pscustomobject]@{ Unexpected = 'value' } } }
            }
        }
        $plan = @(Get-TagOperationPlan -Targets @($target) -Scope $scope -DefaultProfile $context -Tags @{ Owner = 'Bob' })
        $plan[0].Status | Should -Be 'Error'
        $plan[0].Payload.Count | Should -Be 0
        $plan[0].Detail | Should -Match 'invalid tag response'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0 -Exactly
    }

    It 'revalidates captured identity immediately before each planning read' {
        $mutableContext = $context.PSObject.Copy()
        $mutableContext.Account = $context.Account.PSObject.Copy()
        $changeAccount = { $mutableContext.Account.Id = 'changed@example.invalid' }.GetNewClosure()
        $plan = @(Get-TagOperationPlan -Targets @($target) -Scope $scope -DefaultProfile $mutableContext `
            -Tags @{ Owner = 'Bob' } -OnProgress $changeAccount)
        $plan[0].Status | Should -Be 'Error'
        Should -Invoke Get-AzTag -ModuleName ResourceTagger.Core -Times 0 -Exactly
    }

    It 'refuses mismatched context before making an Azure request' {
        $wrong = $scope.PSObject.Copy()
        $wrong.TenantId = 'different'
        { Get-TagOperationPlan -Targets @($target) -Scope $wrong -DefaultProfile $context -Tags @{ Owner = 'Bob' } } | Should -Throw
        Should -Invoke Get-AzTag -ModuleName ResourceTagger.Core -Times 0
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0
    }

    It 'refuses an out-of-scope target without reading it' {
        $outside = [pscustomobject]@{ Id = '/subscriptions/other/resourceGroups/rg'; Name = 'rg'; Kind = 'ResourceGroup' }
        $plan = @(Get-TagOperationPlan -Targets @($outside) -Scope $scope -DefaultProfile $context -Tags @{ Owner = 'Bob' })
        $plan[0].Status | Should -Be 'Error'
        Should -Invoke Get-AzTag -ModuleName ResourceTagger.Core -Times 0
    }

    It 'sends a delta-only merge, passes context, and verifies the written value' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice'; CostCenter = 'old' } `
            -Tags @{ Owner = 'Bob' } -Overwrite
        Mock Get-AzTag -ModuleName ResourceTagger.Core {
            [pscustomobject]@{ Properties = [pscustomobject]@{ TagsProperty = @{ Owner = 'Alice'; CostCenter = 'new' } } }
        }
        Mock Update-AzTag -ModuleName ResourceTagger.Core {
            Mock Get-AzTag -ModuleName ResourceTagger.Core {
                [pscustomobject]@{ Properties = [pscustomobject]@{ TagsProperty = @{ Owner = 'Bob'; CostCenter = 'new' } } }
            }
        }
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Success'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 1 -Exactly -ParameterFilter {
            $ResourceId -eq $target.Id -and $Tag.Count -eq 1 -and $Tag['Owner'] -ceq 'Bob' -and
            $Operation -eq 'Merge' -and $DefaultProfile -eq $context
        }
        Should -Invoke Get-AzTag -ModuleName ResourceTagger.Core -Times 2 -Exactly
    }

    It 'skips a target whose affected tag changed after preview' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'BeforePreview' } `
            -Tags @{ Owner = 'Bob' } -Overwrite
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Conflict'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0
    }

    It 'does not overwrite a tag that appeared after it was previewed as absent' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{} -Tags @{ Owner = 'Bob' }
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Conflict'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0
    }

    It 'does not apply an untagged-only plan after another tag is added' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{} -Tags @{ Environment = 'Test' } -Selection Untagged
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Conflict'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0
    }

    It 'does not claim success when the server did not apply the requested value' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = 'Bob' } -Overwrite
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Unverified'
        $result[0].Detail | Should -Match 'verif'
    }

    It 'does not claim removal success if the key still exists' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = '' } -Operation Delete
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Unverified'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 1 -ParameterFilter {
            $Operation -eq 'Delete' -and $Tag['Owner'] -ceq 'Alice'
        }
    }

    It 'verifies a successful deletion without sending unrelated keys' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice'; Other = 'keep' } `
            -Tags @{ Owner = 'Alice' } -Operation Delete -MatchValue
        Mock Update-AzTag -ModuleName ResourceTagger.Core {
            Mock Get-AzTag -ModuleName ResourceTagger.Core {
                [pscustomobject]@{ Properties = [pscustomobject]@{ TagsProperty = @{ Other = 'keep' } } }
            }
        }
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Success'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 1 -Exactly -ParameterFilter {
            $Tag.Count -eq 1 -and $Tag.Owner -ceq 'Alice' -and $Operation -eq 'Delete'
        }
    }

    It 'does not delete a tag whose value changed after preview' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'alice' } `
            -Tags @{ Owner = 'alice' } -Operation Delete -MatchValue
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Conflict'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0 -Exactly
    }

    It 'reports an unverified write when the verification read fails' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } `
            -Tags @{ Owner = 'Bob' } -Overwrite
        Mock Update-AzTag -ModuleName ResourceTagger.Core {
            Mock Get-AzTag -ModuleName ResourceTagger.Core { throw 'Verification read unavailable' }
        }
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Unverified'
        $result[0].Detail | Should -Match 'uncertain'
    }

    It 'passes the same delta and context to a custom writer and verifies its result' {
        $calls = [System.Collections.Generic.List[object]]::new()
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice'; Other = 'keep' } `
            -Tags @{ Owner = 'Bob' } -Overwrite
        $writer = {
            param($id, $delta, $operation, $azureContext)
            $calls.Add([pscustomobject]@{ Id = $id; Tags = $delta; Operation = $operation; Context = $azureContext })
            Mock Get-AzTag -ModuleName ResourceTagger.Core {
                [pscustomobject]@{ Properties = [pscustomobject]@{ TagsProperty = @{ Owner = 'Bob'; Other = 'keep' } } }
            }
        }
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context `
            -UpdateAction $writer -Confirm:$false)
        $result[0].Status | Should -Be 'Success'
        $calls.Count | Should -Be 1
        $calls[0].Id | Should -Be $target.Id
        $calls[0].Tags.Count | Should -Be 1
        $calls[0].Tags.Owner | Should -BeExactly 'Bob'
        $calls[0].Context | Should -Be $context
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0 -Exactly
    }

    It 'reports uncertain outcomes when a write throws or times out' {
        Mock Update-AzTag -ModuleName ResourceTagger.Core { throw 'Timed out' }
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = 'Bob' } -Overwrite
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Unverified'
        $result[0].Detail | Should -Match 'Timed out'
    }

    It 'supports WhatIf without writing' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = 'Bob' } -Overwrite
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -WhatIf)
        $result[0].Status | Should -Be 'Skipped'
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0
    }

    It 'never executes skipped or errored preview entries' {
        $plan = New-TagChangePlan -Target $target -Scope $scope -CurrentTags @{ Owner = 'Alice' } -Tags @{ Owner = 'Bob' }
        $result = @(Invoke-TagOperationPlan -Plan @($plan) -Scope $scope -DefaultProfile $context -Confirm:$false)
        $result[0].Status | Should -Be 'Skipped'
        Should -Invoke Get-AzTag -ModuleName ResourceTagger.Core -Times 0
        Should -Invoke Update-AzTag -ModuleName ResourceTagger.Core -Times 0
    }
}

Describe 'Azure Resource Tagger UI' {
    BeforeAll {
        $root = Split-Path $PSScriptRoot
        $tokens = $null
        $parseErrors = $null
        $appAst = [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $root 'Start-ResourceTagger.ps1'), [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors.Count) { throw ($parseErrors.Message -join '; ') }
        $firstUiStatement = $appAst.EndBlock.Statements | Where-Object {
            $_ -is [System.Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -eq '$xamlPath'
        }
        # Run the real initialization and event bindings, without dependency prompts or ShowDialog.
        $initialization = ($appAst.EndBlock.Statements | Where-Object {
                $_.Extent.StartOffset -ge $firstUiStatement.Extent.StartOffset -and
                $_.Extent.Text -ne '$window.ShowDialog() | Out-Null'
            } | ForEach-Object { $_.Extent.Text }) -join "`n"

        $uiModule = New-Module -Name 'ResourceTagger.UiHarness' -ArgumentList $root, $initialization -ScriptBlock {
            param($AppRoot, $Initialization)
            Set-StrictMode -Version Latest
            $ErrorActionPreference = 'Stop'
            $script:ScriptRoot = $AppRoot
            $script:Version = 'test'
            Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
            Import-Module (Join-Path $AppRoot 'ResourceTagger.Core.psm1') -Force
            & (Get-Module ResourceTagger.Core) {
                function script:Get-AzTag {
                    param($ResourceId, $DefaultProfile, $ErrorAction)
                    throw 'Unmocked Azure read in UI test'
                }
                function script:Update-AzTag {
                    param($ResourceId, $Tag, $Operation, $DefaultProfile, $ErrorAction)
                    throw 'Unmocked Azure write in UI test'
                }
            }
            function Get-AzContext { param($ErrorAction); throw 'Unmocked context lookup in UI test' }
            function Get-AzTenant { param($DefaultProfile, $ErrorAction); throw 'Unmocked tenant lookup in UI test' }
            function Get-AzSubscription { param($TenantId, $DefaultProfile, $ErrorAction); throw 'Unmocked subscription lookup in UI test' }
            function Get-AzResourceGroup { param($DefaultProfile, $ErrorAction); throw 'Unmocked group lookup in UI test' }
            function Set-AzContext { param($Subscription, $Tenant, $DefaultProfile, $Scope, $ErrorAction); throw 'Unmocked context change in UI test' }
            function Connect-AzAccount { param($Environment, $TenantId, $Scope, $ErrorAction); throw 'Unmocked login in UI test' }
            function Search-AzGraph { param($Query, $Subscription, $First, $SkipToken, $DefaultProfile, $ErrorAction); throw 'Unmocked graph lookup in UI test' }
            . ([scriptblock]::Create($Initialization))
        }
        Import-Module $uiModule -DisableNameChecking
    }

    BeforeEach {
        Mock Flush-UI -ModuleName ResourceTagger.UiHarness {}
        Mock Show-TaggerError -ModuleName ResourceTagger.UiHarness {}
        Mock Show-TagPlanPreview -ModuleName ResourceTagger.UiHarness { $false }
        Mock Read-ResourceTagInput -ModuleName ResourceTagger.UiHarness {
            [pscustomobject]@{ TagName = 'Owner'; TagValue = ' Alice ' }
        }
        Mock Select-TagResourceGroups -ModuleName ResourceTagger.UiHarness { param($Groups); $Groups[1] }
        Mock Get-TagOperationPlan -ModuleName ResourceTagger.UiHarness {
            param($Targets, $Scope, $Tags, $Operation, $Overwrite, $MatchValue, $Selection, $SelectionTagName)
            $current = if ($Operation -eq 'Delete') { @{ Owner = 'old' } } else { @{} }
            foreach ($targetRecord in $Targets) {
                New-TagChangePlan -Target $targetRecord -Scope $Scope -CurrentTags $current -Tags $Tags `
                    -Operation $Operation -Overwrite:$Overwrite -MatchValue:$MatchValue `
                    -Selection $Selection -SelectionTagName $SelectionTagName
            }
        }
        Mock Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness {
            param($Plan)
            foreach ($entry in $Plan) {
                $copy = $entry.PSObject.Copy()
                if ($copy.Status -eq 'Planned') { $copy.Status = 'Success' }
                $copy
            }
        }
        Mock Set-AzContext -ModuleName ResourceTagger.UiHarness {
            param($Subscription, $Tenant, $DefaultProfile)
            [pscustomobject]@{
                Environment = $DefaultProfile.Environment
                Tenant = [pscustomobject]@{ Id = $Tenant }
                Subscription = [pscustomobject]@{ Id = $Subscription }
                Account = $DefaultProfile.Account
            }
        }
        Mock Get-AzResourceGroup -ModuleName ResourceTagger.UiHarness {
            [pscustomobject]@{ ResourceGroupName = 'example' }
        }
        & $uiModule {
            $script:Busy = $true
            Clear-TagScan
            $script:Connected = $true
            $script:ActiveContext = [pscustomobject]@{
                Environment = [pscustomobject]@{ Name = 'AzureCloud' }
                Tenant = [pscustomobject]@{ Id = '00000000-0000-0000-0000-000000000001' }
                Subscription = [pscustomobject]@{ Id = '00000000-0000-0000-0000-000000000002' }
                Account = [pscustomobject]@{ Id = 'tester@example.invalid' }
            }
            $script:Subscriptions = @(
                [pscustomobject]@{ Id = $script:ActiveContext.Subscription.Id; Name = 'Test subscription' },
                [pscustomobject]@{ Id = '00000000-0000-0000-0000-000000000003'; Name = 'Second subscription' }
            )
            $ui.SubscriptionSelector.Items.Clear()
            foreach ($sub in $script:Subscriptions) { [void]$ui.SubscriptionSelector.Items.Add($sub.Name) }
            $ui.SubscriptionSelector.SelectedIndex = 0
            $ui.ScopeLevel.SelectedIndex = 0
            $ui.RGSelector.Items.Clear()
            [void]$ui.RGSelector.Items.Add('(All Resource Groups)')
            [void]$ui.RGSelector.Items.Add('example')
            $ui.RGSelector.SelectedIndex = 0
            $ui.RGFilterTag.SelectedIndex = 0
            $ui.RequiredTagsInput.Text = ''
            $ui.ResFilterTag.SelectedIndex = 0
            $ui.ResFilterTagName.Text = ''
            $ui.ApplyScope.SelectedIndex = 0
            $ui.RemoveScope.SelectedIndex = 0
            $ui.RemoveTagValueFilter.Text = ''
            $ui.RemoveMatchValueCheck.IsChecked = $false
            foreach ($name in @('DryRunCheck','ResDryRunCheck','RemoveDryRunCheck')) { $ui[$name].IsChecked = $true }
            foreach ($name in @('OverwriteCheck','ResOverwriteCheck')) { $ui[$name].IsChecked = $false }
            foreach ($name in @('ApplyResultsGrid','RemoveResultsGrid')) { $ui[$name].ItemsSource = $null }
            $script:TagQueue.Clear()
            $script:TagQueue.Add([pscustomobject]@{ TagName = 'CostCentre'; TagValue = 'C100' })
            $prefix = "/subscriptions/$($script:ActiveContext.Subscription.Id)/resourceGroups"
            $script:AllRGs = @(
                [pscustomobject]@{ id = "$prefix/example"; name = 'example'; location = 'test'; tags = @{ Owner = 'old' } },
                [pscustomobject]@{ id = "$prefix/second"; name = 'second'; location = 'test'; tags = @{} }
            )
            $script:AllResources = @(
                [pscustomobject]@{
                    id = "$prefix/example/providers/Microsoft.Test/widgets/one"; name = 'one'
                    type = 'Microsoft.Test/widgets'; resourceGroup = 'example'; location = 'test'; tags = @{ Owner = 'old' }
                },
                [pscustomobject]@{
                    id = "$prefix/example/providers/Microsoft.Test/widgets/two"; name = 'two'
                    type = 'Microsoft.Test/widgets'; resourceGroup = 'example'; location = 'test'; tags = @{}
                }
            )
            $script:ScanScope = New-TagScope -Context $script:ActiveContext
            $script:ScanContext = $script:ActiveContext
            $script:ScannedAt = Get-Date
            Update-ResourceGroupView
            Update-ResourceView
            Update-RemoveTagList
            Set-TaggerBusy $false
        }
    }

    AfterAll {
        if ($uiModule) {
            & $uiModule { $script:Busy = $false; $window.Close() }
            Remove-Module $uiModule -Force
        }
    }

    Describe 'WPF controls, filters, and scope invalidation' {
        It 'resolves every named control and defaults every operation to dry run' {
            $state = & $uiModule {
                [pscustomobject]@{
                    Bound = $ui.Count -eq $controlNames.Count
                    Dry = $ui.DryRunCheck.IsChecked -and $ui.ResDryRunCheck.IsChecked -and $ui.RemoveDryRunCheck.IsChecked
                }
            }
            $state.Bound | Should -BeTrue
            $state.Dry | Should -BeTrue
        }

        It 'rebuilds resource filters and restores all rows without another Azure query' {
            $counts = & $uiModule {
                $ui.ResFilterTag.SelectedIndex = 1
                @($ui.ResourceGrid.ItemsSource).Count
                $ui.ResFilterTag.SelectedIndex = 0
                @($ui.ResourceGrid.ItemsSource).Count
                $ui.ResFilterTag.SelectedIndex = 2
                $ui.ResFilterTagName.Text = 'owner'
                @($ui.ResourceGrid.ItemsSource).Count
                $ui.ResFilterTagName.Text = 'CostCentre'
                @($ui.ResourceGrid.ItemsSource).Count
            }
            $counts | Should -Be @(1, 2, 1, 2)
        }

        It 'recomputes required-tag gaps as text changes and restores filtered groups' {
            $counts = & $uiModule {
                $ui.RequiredTagsInput.Text = 'Owner'
                $ui.RGFilterTag.SelectedIndex = 1
                @($ui.RGGrid.ItemsSource).Count
                $ui.RequiredTagsInput.Text = 'Owner, CostCentre'
                @($ui.RGGrid.ItemsSource).Count
                $ui.RGFilterTag.SelectedIndex = 0
                @($ui.RGGrid.ItemsSource).Count
            }
            $counts | Should -Be @(1, 2, 2)
        }

        It 'invalidates scanned data when <Control> changes' -TestCases @(
            @{ Control = 'ScopeLevel' },
            @{ Control = 'RGSelector' },
            @{ Control = 'SubscriptionSelector' }
        ) {
            param($Control)
            $state = & $uiModule {
                param($controlName)
                $ui[$controlName].SelectedIndex = 1
                [pscustomobject]@{
                    Scope = $script:ScanScope
                    Rows = $script:AllRGs.Count + $script:AllResources.Count
                    CanWrite = $ui.ApplyTagsButton.IsEnabled -or $ui.RemoveTagsButton.IsEnabled -or $ui.ResTagSelectedButton.IsEnabled
                    CanExport = $ui.ExportButton.IsEnabled
                }
            } $Control
            $state.Scope | Should -BeNullOrEmpty
            $state.Rows | Should -Be 0
            $state.CanWrite | Should -BeFalse
            $state.CanExport | Should -BeFalse
        }

        It 'clears scan state even when reconnecting fails before authentication' {
            & $uiModule { Connect-ToAzure -AzureEnvironment AzureCloud }
            Should -Invoke Show-TaggerError -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly
            $state = & $uiModule { [pscustomobject]@{ Scope = $script:ScanScope; Connected = $script:Connected; Busy = $script:Busy } }
            $state.Scope | Should -BeNullOrEmpty
            $state.Connected | Should -BeFalse
            $state.Busy | Should -BeFalse
        }

        It 'disables conflicting controls and blocks reentrant tagging events while busy' {
            $enabled = & $uiModule {
                Set-TaggerBusy $true
                $eventId = [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent
                foreach ($name in @('ApplyTagsButton','RemoveTagsButton','ResTagSelectedButton','CommercialButton','ScanButton')) {
                    $ui[$name].RaiseEvent([System.Windows.RoutedEventArgs]::new($eventId))
                }
                @('ApplyTagsButton','RemoveTagsButton','CommercialButton','GovButton','ScanButton','ExportButton',
                    'SubscriptionSelector','ScopeLevel','RGSelector','DryRunCheck','AddTagButton') |
                    Where-Object { $ui[$_].IsEnabled }
            }
            $enabled | Should -BeNullOrEmpty
            Should -Invoke Get-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
        }

        It 'enables exact-value input only when explicitly selected' {
            $states = & $uiModule {
                $ui.RemoveTagValueFilter.IsEnabled
                $ui.RemoveMatchValueCheck.IsChecked = $true
                $ui.RemoveMatchValueCheck.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                $ui.RemoveTagValueFilter.IsEnabled
            }
            $states | Should -Be @($false, $true)
        }
    }

    Describe 'Shared workflow integration' {
        It 'routes apply scope <Index> through the same planner' -TestCases @(
            @{ Index = 0; Count = 2; Selection = 'All' },
            @{ Index = 1; Count = 2; Selection = 'MissingTag' },
            @{ Index = 2; Count = 2; Selection = 'All' },
            @{ Index = 3; Count = 2; Selection = 'Untagged' },
            @{ Index = 4; Count = 1; Selection = 'All' }
        ) {
            param($Index, $Count, $Selection)
            & $uiModule {
                param($index)
                $ui.ApplyScope.SelectedIndex = $index
                $ui.ApplyTagsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            } $Index
            $expectedCount = $Count
            $expectedSelection = $Selection
            Should -Invoke Get-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly -ParameterFilter {
                $Targets.Count -eq $expectedCount -and $Selection -eq $expectedSelection -and
                $SelectionTagName -eq 'CostCentre' -and $Tags.CostCentre -eq 'C100' -and
                $DefaultProfile.Subscription.Id -eq '00000000-0000-0000-0000-000000000002'
            }
            Should -Invoke Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
        }

        It 'routes individual resources through the planner without trimming tag values' {
            & $uiModule {
                $ui.ResourceGrid.SelectedItem = $ui.ResourceGrid.Items[0]
                $ui.ResTagSelectedButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            }
            Should -Invoke Get-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly -ParameterFilter {
                $Targets.Count -eq 1 -and $Targets[0].Name -eq 'one' -and $Tags.Owner -ceq ' Alice ' -and $Operation -eq 'Merge'
            }
            Should -Invoke Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
        }

        It 'routes removal scope <Index> through fresh planning including previously absent keys' -TestCases @(
            @{ Index = 0; Count = 2 },
            @{ Index = 1; Count = 2 },
            @{ Index = 2; Count = 4 }
        ) {
            param($Index, $Count)
            & $uiModule {
                param($index)
                $ui.RemoveScope.SelectedIndex = $index
                $ui.RemoveMatchValueCheck.IsChecked = $true
                $ui.RemoveTagValueFilter.Text = ''
                $ui.RemoveTagsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            } $Index
            $expectedCount = $Count
            Should -Invoke Get-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly -ParameterFilter {
                $Operation -eq 'Delete' -and $Targets.Count -eq $expectedCount -and $MatchValue -and $Tags.Owner -ceq ''
            }
            Should -Invoke Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
        }

        It 'never executes a dry run even if the preview returns approval' {
            Mock Show-TagPlanPreview -ModuleName ResourceTagger.UiHarness { $true }
            & $uiModule { Invoke-TagWorkflow -Targets @(Get-ScanTarget ResourceGroups) -Tags @{ Owner = 'Alice' } -DryRun $true }
            Should -Invoke Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
            (& $uiModule { $script:ScanScope }) | Should -Not -BeNullOrEmpty
        }

        It 'does not execute a live operation when preview is cancelled' {
            & $uiModule { Invoke-TagWorkflow -Targets @(Get-ScanTarget ResourceGroups) -Tags @{ Owner = 'Alice' } -DryRun $false }
            Should -Invoke Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
            (& $uiModule { $script:ScanScope }) | Should -Not -BeNullOrEmpty
        }

        It 'invalidates inventory after confirmed execution and retains results' {
            Mock Show-TagPlanPreview -ModuleName ResourceTagger.UiHarness { $true }
            & $uiModule { Invoke-TagWorkflow -Targets @(Get-ScanTarget ResourceGroups) -Tags @{ Owner = 'Alice' } -DryRun $false }
            Should -Invoke Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly -ParameterFilter {
                $DefaultProfile.Account.Id -eq 'tester@example.invalid' -and $null -ne $UpdateAction
            }
            $state = & $uiModule {
                [pscustomobject]@{
                    Scope = $script:ScanScope; Results = @($ui.ApplyResultsGrid.ItemsSource).Count
                    CanWrite = $ui.ApplyTagsButton.IsEnabled; Busy = $script:Busy
                }
            }
            $state.Scope | Should -BeNullOrEmpty
            $state.Results | Should -Be 2
            $state.CanWrite | Should -BeFalse
            $state.Busy | Should -BeFalse
        }

        It 'invalidates inventory and surfaces an execution failure without claiming success' {
            Mock Show-TagPlanPreview -ModuleName ResourceTagger.UiHarness { $true }
            Mock Invoke-TagOperationPlan -ModuleName ResourceTagger.UiHarness { throw 'Simulated uncertain failure' }
            & $uiModule { Invoke-TagWorkflow -Targets @(Get-ScanTarget ResourceGroups) -Tags @{ Owner = 'Alice' } -DryRun $false }
            Should -Invoke Show-TaggerError -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly
            (& $uiModule { $script:ScanScope }) | Should -BeNullOrEmpty
            (& $uiModule { $ui.ApplyStatusText.Text }) | Should -Match 'Simulated uncertain failure'
        }

        It 'blocks a changed context before planning, even if the button was still enabled' {
            & $uiModule {
                $script:ActiveContext.Account.Id = 'changed@example.invalid'
                Invoke-TagWorkflow -Targets @(Get-ScanTarget ResourceGroups) -Tags @{ Owner = 'Alice' }
            }
            Should -Invoke Get-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
            Should -Invoke Show-TaggerError -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly
            (& $uiModule { $script:ScanScope }) | Should -BeNullOrEmpty
        }

        It 'deduplicates targets before planning' {
            & $uiModule {
                $target = @(Get-ScanTarget ResourceGroups)[0]
                Invoke-TagWorkflow -Targets @($target, $target) -Tags @{ Owner = 'Alice' }
            }
            Should -Invoke Get-TagOperationPlan -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly -ParameterFilter { $Targets.Count -eq 1 }
        }
    }

    Describe 'Preview presentation and Azure boundaries' {
        It 'neutralizes untrusted cells through the actual CSV export event without changing inventory' {
            $exportPath = Join-Path $TestDrive 'tag-report.csv'
            $saveDialog = [pscustomobject]@{ Filter = ''; FileName = ''; Destination = $exportPath }
            $saveDialog | Add-Member -MemberType ScriptMethod -Name ShowDialog -Value {
                $this.FileName = $this.Destination
                return $true
            }
            Mock New-Object -ModuleName ResourceTagger.UiHarness { $saveDialog } -ParameterFilter {
                $TypeName -eq 'Microsoft.Win32.SaveFileDialog'
            }
            & $uiModule {
                $script:AllRGs[0].tags = @{ '=1+1' = '2' }
                $script:AllResources[0].tags = @{ '=1+1' = '2' }
                $script:AllResources[0].name = '+1+1'
                $ui.ExportButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            }
            Should -Invoke Show-TaggerError -ModuleName ResourceTagger.UiHarness -Times 0 -Exactly
            $rows = @(Import-Csv -LiteralPath $exportPath)
            $rows.Count | Should -Be 4
            @($rows | Where-Object Tags -eq '[text] =1+1=2').Count | Should -Be 2
            @($rows | Where-Object Name -eq '[text] +1+1').Count | Should -Be 1
            (& $uiModule { $script:AllRGs[0].tags['=1+1'] }) | Should -BeExactly '2'
            (& $uiModule { $script:AllResources[0].name }) | Should -BeExactly '+1+1'
            (& $uiModule { $ui.StatusText.Text }) | Should -Match 'Formula-like cells are prefixed'
            (& $uiModule { $script:Busy }) | Should -BeFalse
        }

        It 'shows exact preview counts and hides the dry-run apply button' {
            $state = & $uiModule {
                $target = @(Get-ScanTarget ResourceGroups)[0]
                $plan = @(New-TagChangePlan -Target $target -Scope $script:ScanScope -CurrentTags @{ Owner = 'old' } `
                        -Tags @{ Owner = 'Alice'; CostCentre = 'C100' })
                $owner = $window
                $window = $null
                try {
                    $dialog = New-TagPreviewWindow -Plan $plan -Scope $script:ScanScope -DryRun $true
                    [pscustomobject]@{
                        Enabled = $dialog.FindName('ApplyButton').IsEnabled
                        Visibility = $dialog.FindName('ApplyButton').Visibility.ToString()
                        Summary = $dialog.FindName('SummaryText').Text
                        Scope = $dialog.FindName('ScopeText').Text
                        Rows = @($dialog.FindName('PlanGrid').ItemsSource)
                    }
                    $dialog.Close()
                } finally { $window = $owner }
            }
            $state.Enabled | Should -BeFalse
            $state.Visibility | Should -Be 'Collapsed'
            $state.Summary | Should -Match '1 tag change\(s\) planned, 1 skipped'
            $state.Scope | Should -Match 'tester@example.invalid'
            $state.Rows.Count | Should -Be 2
        }

        It 'reads every Resource Graph page using the explicit profile' {
            Mock Search-AzGraph -ModuleName ResourceTagger.UiHarness {
                param($SkipToken)
                if ($SkipToken) { [pscustomobject]@{ Data = @([pscustomobject]@{ id = 'two' }); SkipToken = $null } }
                else { [pscustomobject]@{ Data = @([pscustomobject]@{ id = 'one' }); SkipToken = 'next' } }
            }
            $rows = & $uiModule { @(Search-AzGraphSafe -Query 'test query' -Subscriptions @('test') -DefaultProfile $script:ActiveContext) }
            @($rows).Count | Should -Be 2
            Should -Invoke Search-AzGraph -ModuleName ResourceTagger.UiHarness -Times 2 -Exactly -ParameterFilter {
                $DefaultProfile.Account.Id -eq 'tester@example.invalid'
            }
        }

        It 'captures scan identity only after both queries succeed' {
            $scanFixtureGroups = & $uiModule { $script:AllRGs }
            $scanFixtureResources = & $uiModule { $script:AllResources }
            Mock Search-AzGraphSafe -ModuleName ResourceTagger.UiHarness {
                param($Query)
                if ($Query -like 'resourcecontainers*') { $scanFixtureGroups }
                else { $scanFixtureResources }
            }
            & $uiModule {
                $ui.ScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            }
            Should -Invoke Search-AzGraphSafe -ModuleName ResourceTagger.UiHarness -Times 2 -Exactly -ParameterFilter {
                $DefaultProfile.Account.Id -eq 'tester@example.invalid'
            }
            (& $uiModule { $ui.StatusText.Text }) | Should -BeLike 'Scan complete*'
            (& $uiModule { $script:ScanScope.SubscriptionId }) | Should -Be '00000000-0000-0000-0000-000000000002'
            (& $uiModule { @($ui.ResourceGrid.ItemsSource).Count }) | Should -Be 2
            (& $uiModule { $ui.ApplyTagsButton.IsEnabled }) | Should -BeTrue
        }

        It 'does not retain a partial or previous inventory after a failed scan' {
            Mock Search-AzGraphSafe -ModuleName ResourceTagger.UiHarness { throw 'Simulated scan failure' }
            & $uiModule { $ui.ScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
            (& $uiModule { $script:ScanScope }) | Should -BeNullOrEmpty
            (& $uiModule { $script:AllRGs.Count + $script:AllResources.Count }) | Should -Be 0
            (& $uiModule { $ui.ExportButton.IsEnabled }) | Should -BeFalse
            Should -Invoke Show-TaggerError -ModuleName ResourceTagger.UiHarness -Times 1 -Exactly
        }

        It 'passes the context into a parameterized worker instead of constructing executable tag text' {
            $worker = $appAst.Find({
                    param($node)
                    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-AzTagUpdateSafe'
                }, $true)
            $worker.Extent.Text | Should -Match '\.AddArgument\(\$DefaultProfile\)'
            $write = $worker.Find({
                    param($node)
                    $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Update-AzTag'
                }, $true)
            $write.Extent.Text | Should -Match '-DefaultProfile \$azureProfile -ErrorAction Stop'
            $write.Extent.Text | Should -Match '-Tag \$tag'
        }
    }

    Describe 'Isolated background write worker' {
        BeforeEach {
            Mock New-TagWriteRunspace -ModuleName ResourceTagger.UiHarness {
                $session = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                $stub = [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Update-AzTag', @'
param($ResourceId, $Tag, $Operation, $DefaultProfile, $ErrorAction)
if ($DefaultProfile.Fail) { throw 'Simulated worker failure' }
if ($DefaultProfile.Delay) { Start-Sleep -Milliseconds $DefaultProfile.Delay }
[void]$DefaultProfile.Calls.Add([pscustomobject]@{
    ResourceId = $ResourceId; Tag = $Tag; Operation = $Operation
    AccountId = $DefaultProfile.Account.Id; ErrorAction = $ErrorAction
})
'@)
                $session.Commands.Add($stub)
                [runspacefactory]::CreateRunspace($session)
            }
            $workerContext = [pscustomobject]@{
                Account = [pscustomobject]@{ Id = 'worker@example.invalid' }
                Fail = $false
                Delay = 0
                Calls = [System.Collections.Generic.List[object]]::new()
            }
        }

        It 'passes literal values and the captured profile for <Operation> writes' -TestCases @(
            @{ Operation = 'Merge' }, @{ Operation = 'Delete' }
        ) {
            param($Operation)
            & $uiModule {
                param($azureContext, $op)
                Invoke-AzTagUpdateSafe -ResourceId '/subscriptions/test/resourceGroups/example' `
                    -Tag @{ Owner = "a'; throw 'not executable" } -Operation $op -DefaultProfile $azureContext
            } $workerContext $Operation
            $workerContext.Calls.Count | Should -Be 1
            $workerContext.Calls[0].Operation | Should -Be $Operation
            $workerContext.Calls[0].Tag.Owner | Should -BeExactly "a'; throw 'not executable"
            $workerContext.Calls[0].AccountId | Should -Be 'worker@example.invalid'
            $workerContext.Calls[0].ErrorAction | Should -Be 'Stop'
        }

        It 'propagates a worker error instead of reporting completion' {
            $workerContext.Fail = $true
            {
                & $uiModule {
                    param($azureContext)
                    Invoke-AzTagUpdateSafe -ResourceId '/subscriptions/test/resourceGroups/example' `
                        -Tag @{ Owner = 'Alice' } -Operation Merge -DefaultProfile $azureContext
                } $workerContext
            } | Should -Throw '*Simulated worker failure*'
            $workerContext.Calls.Count | Should -Be 0
        }

        It 'stops and disposes a timed-out worker with an explicit error' {
            $workerContext.Delay = 5000
            {
                & $uiModule {
                    param($azureContext)
                    Invoke-AzTagUpdateSafe -ResourceId '/subscriptions/test/resourceGroups/example' `
                        -Tag @{ Owner = 'Alice' } -Operation Merge -DefaultProfile $azureContext -TimeoutSeconds 1
                } $workerContext
            } | Should -Throw '*timed out*'
            $workerContext.Calls.Count | Should -Be 0
        }
    }
}

<#
.SYNOPSIS
    Azure Resource Tagger - Scan existing tags and bulk-apply new tags to Azure resources.

.DESCRIPTION
    A WPF-based tool that connects to your Azure tenant, scans resource groups and
    resources for existing tags, identifies tagging gaps, and lets you bulk-apply
    tags at scale.  Designed for governance and compliance workflows such as
    preparing a subscription for Azure Policy tag enforcement.

.NOTES
    Version : 1.0.0
    Author  : Zach Larsen
    Requires: Az.Accounts, Az.Resources, Az.ResourceGraph
#>

#Requires -Version 5.1

param(
    [switch]$Debug
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptRoot = $PSScriptRoot
$script:Version    = '1.0.0'

# ─────────────────────────────────────────────────────────────────
# WPF bootstrap
# ─────────────────────────────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ─────────────────────────────────────────────────────────────────
# Verify required Az modules
# ─────────────────────────────────────────────────────────────────
$requiredModules = @('Az.Accounts','Az.Resources','Az.ResourceGraph')
foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        [System.Windows.MessageBox]::Show(
            "Required module '$mod' is not installed.`n`nRun:  Install-Module $mod -Scope CurrentUser",
            'Missing Dependency', 'OK', 'Error') | Out-Null
        return
    }
}

# ─────────────────────────────────────────────────────────────────
# Load XAML
# ─────────────────────────────────────────────────────────────────
$xamlPath = Join-Path $script:ScriptRoot 'gui\MainWindow.xaml'
if (-not (Test-Path $xamlPath)) {
    [System.Windows.MessageBox]::Show("Cannot find $xamlPath", 'Error', 'OK', 'Error') | Out-Null
    return
}

[xml]$xaml = Get-Content $xamlPath -Raw
$reader   = New-Object System.Xml.XmlNodeReader $xaml
$window   = [Windows.Markup.XamlReader]::Load($reader)

# ─────────────────────────────────────────────────────────────────
# Resolve named controls
# ─────────────────────────────────────────────────────────────────
$controlNames = @(
    'VersionLabel','TenantLabel',
    'CommercialButton','GovButton','ScanButton','ExportButton',
    'ScopeLevel','SubscriptionSelector','RGSelector',
    'RGCountText','ResourceCountText','TagCoverageText','UntaggedRGText','UniqueTagsText',
    'TagSummaryGrid',
    'RGFilterTag','RequiredTagsInput','RequiredTagsPlaceholder','RGGrid',
    'ResFilterTag','ResFilterTagName','ResourceGrid',
    'ApplyTagName','ApplyTagValue','AddTagButton',
    'TagQueueGrid','ClearTagsButton','RemoveTagButton',
    'ApplyScope','OverwriteCheck','DryRunCheck',
    'ApplyTagsButton','ApplyStatusText','ApplyResultsGrid',
    'ProgressBar','StatusText','MainTabs'
)

$ui = @{}
foreach ($name in $controlNames) {
    $ctrl = $window.FindName($name)
    if ($ctrl) { $ui[$name] = $ctrl }
}

$ui.VersionLabel.Text = "v$($script:Version)"

# ─────────────────────────────────────────────────────────────────
# State
# ─────────────────────────────────────────────────────────────────
$script:Connected       = $false
$script:Environment     = ''
$script:Subscriptions   = @()
$script:AllRGs          = @()
$script:AllResources    = @()
$script:TagQueue        = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
$ui.TagQueueGrid.ItemsSource = $script:TagQueue

# Lock icon characters (surrogates for PS 5.1 compat)
$script:LockOpen   = [char]::ConvertFromUtf32(0x1F513)
$script:LockClosed = [char]::ConvertFromUtf32(0x1F512)

# Placeholder text behavior for Required Tags input
$ui.RequiredTagsInput.Add_GotFocus({
    $ui.RequiredTagsPlaceholder.Visibility = 'Collapsed'
})
$ui.RequiredTagsInput.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($ui.RequiredTagsInput.Text)) {
        $ui.RequiredTagsPlaceholder.Visibility = 'Visible'
    }
})

# ─────────────────────────────────────────────────────────────────
# Helper: Update status bar
# ─────────────────────────────────────────────────────────────────
function Update-Status {
    param([string]$Message, [int]$Progress = -1)
    $ui.StatusText.Text = $Message
    if ($Progress -ge 0) { $ui.ProgressBar.Value = $Progress }
    [System.Windows.Forms.Application]::DoEvents()
}

Add-Type -AssemblyName System.Windows.Forms   # for DoEvents

# ─────────────────────────────────────────────────────────────────
# Helper: Convert tag object (hashtable, OrderedDictionary,
#         PSCustomObject, etc.) to a plain hashtable
# ─────────────────────────────────────────────────────────────────
function ConvertTo-TagHashtable {
    param($Tags)
    $map = @{}
    if ($null -eq $Tags) { return $map }
    if ($Tags -is [System.Collections.IDictionary]) {
        foreach ($k in $Tags.Keys) { $map[$k] = $Tags[$k] }
    } elseif ($Tags -is [PSCustomObject]) {
        $Tags.PSObject.Properties | ForEach-Object { $map[$_.Name] = $_.Value }
    } else {
        try { $Tags.PSObject.Properties | ForEach-Object { $map[$_.Name] = $_.Value } } catch {}
    }
    return $map
}

# ─────────────────────────────────────────────────────────────────
# Helper: Safe Resource Graph query with paging
# ─────────────────────────────────────────────────────────────────
function Search-AzGraphSafe {
    param(
        [string]$Query,
        [string[]]$Subscriptions
    )
    $all   = [System.Collections.Generic.List[object]]::new()
    $skip  = $null
    $first = 1000

    do {
        $params = @{
            Query        = $Query
            Subscription = $Subscriptions
            First        = $first
        }
        if ($skip) { $params['SkipToken'] = $skip }

        $result = Search-AzGraph @params
        if ($result.Data) { $all.AddRange($result.Data) }
        elseif ($result) {
            foreach ($r in $result) { $all.Add($r) }
        }
        $skip = $result.SkipToken
    } while ($skip)

    return $all
}

# ─────────────────────────────────────────────────────────────────
# Shared connect logic
# ─────────────────────────────────────────────────────────────────
function Connect-ToAzure {
    param([string]$AzureEnvironment)

    $ui.CommercialButton.IsEnabled = $false
    $ui.GovButton.IsEnabled        = $false
    $ui.ScanButton.IsEnabled       = $false

    $envLabel = if ($AzureEnvironment -eq 'AzureUSGovernment') { 'Gov' } else { 'Commercial' }
    $btn      = if ($AzureEnvironment -eq 'AzureUSGovernment') { $ui.GovButton } else { $ui.CommercialButton }
    $btn.Content = "$($script:LockOpen) $envLabel Tenant"

    try {
        Update-Status "Connecting to Azure $envLabel..." 10

        # Disable Az 12+ interactive subscription picker
        $env:AZURE_LOGIN_EXPERIENCE_V2 = 'Off'

        $ctx = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $ctx -or $ctx.Environment.Name -ne $AzureEnvironment) {
            $window.WindowState = 'Minimized'
            try {
                Connect-AzAccount -Environment $AzureEnvironment -ErrorAction Stop | Out-Null
            } finally {
                $window.WindowState = 'Normal'
                $window.Activate()
            }
            $ctx = Get-AzContext
        }

        $tenantId = $ctx.Tenant.Id
        $script:Environment = $AzureEnvironment

        Update-Status 'Listing subscriptions...' 30

        $script:Subscriptions = @(Get-AzSubscription -TenantId $tenantId -ErrorAction Stop |
            Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)

        $ui.SubscriptionSelector.Items.Clear()
        foreach ($sub in $script:Subscriptions) {
            $item = "$($sub.Name)  ($($sub.Id))"
            $ui.SubscriptionSelector.Items.Add($item) | Out-Null
        }
        if ($ui.SubscriptionSelector.Items.Count -gt 0) {
            $ui.SubscriptionSelector.SelectedIndex = 0
        }

        $ui.TenantLabel.Text = "Tenant: $tenantId  |  $($ctx.Account.Id)  |  $AzureEnvironment"
        $ui.SubscriptionSelector.IsEnabled = $true
        $ui.ScanButton.IsEnabled  = $true
        $script:Connected = $true

        $btn.Content = "$($script:LockClosed) $envLabel Tenant"
        $subCount = @($script:Subscriptions).Count
        Update-Status "Connected to $envLabel - $subCount subscriptions found" 100
    }
    catch {
        Update-Status "Connection failed: $($_.Exception.Message)" 0
        $btn.Content = "$envLabel Tenant"
        [System.Windows.MessageBox]::Show(
            "Failed to connect:`n$($_.Exception.Message)",
            'Connection Error', 'OK', 'Error') | Out-Null
    }

    $ui.CommercialButton.IsEnabled = $true
    $ui.GovButton.IsEnabled        = $true
}

# ─────────────────────────────────────────────────────────────────
# Commercial Tenant button
# ─────────────────────────────────────────────────────────────────
$ui.CommercialButton.Add_Click({
    Connect-ToAzure -AzureEnvironment 'AzureCloud'
})

# ─────────────────────────────────────────────────────────────────
# Gov Tenant button
# ─────────────────────────────────────────────────────────────────
$ui.GovButton.Add_Click({
    Connect-ToAzure -AzureEnvironment 'AzureUSGovernment'
})

# ─────────────────────────────────────────────────────────────────
# Subscription selection → populate RGs
# ─────────────────────────────────────────────────────────────────
$ui.SubscriptionSelector.Add_SelectionChanged({
    $idx = $ui.SubscriptionSelector.SelectedIndex
    if ($idx -lt 0) { return }
    $sub = $script:Subscriptions[$idx]
    Set-AzContext -SubscriptionId $sub.Id -ErrorAction SilentlyContinue | Out-Null

    $ui.RGSelector.Items.Clear()
    $ui.RGSelector.Items.Add('(All Resource Groups)') | Out-Null
    try {
        $rgs = Get-AzResourceGroup | Sort-Object ResourceGroupName
        foreach ($rg in $rgs) {
            $ui.RGSelector.Items.Add($rg.ResourceGroupName) | Out-Null
        }
    } catch {}
    $ui.RGSelector.SelectedIndex = 0
    $ui.RGSelector.IsEnabled = ($ui.ScopeLevel.SelectedIndex -eq 1)
})

# ─────────────────────────────────────────────────────────────────
# Scope level change → enable/disable RG selector
# ─────────────────────────────────────────────────────────────────
$ui.ScopeLevel.Add_SelectionChanged({
    $ui.RGSelector.IsEnabled = ($ui.ScopeLevel.SelectedIndex -eq 1)
})

# ─────────────────────────────────────────────────────────────────
# SCAN
# ─────────────────────────────────────────────────────────────────
$ui.ScanButton.Add_Click({
    $subIdx = $ui.SubscriptionSelector.SelectedIndex
    if ($subIdx -lt 0) { return }
    $sub = $script:Subscriptions[$subIdx]
    $subId = $sub.Id

    try {
        Update-Status 'Scanning resource groups...' 10
        $ui.ScanButton.IsEnabled = $false

        # --- Resource Groups ---
        $rgFilter = $null
        if ($ui.ScopeLevel.SelectedIndex -eq 1 -and $ui.RGSelector.SelectedIndex -gt 0) {
            $rgFilter = $ui.RGSelector.SelectedItem.ToString()
        }

        $rgQuery = "resourcecontainers | where type == 'microsoft.resources/subscriptions/resourcegroups' | project name, id, location, tags, subscriptionId"
        $allRGs  = Search-AzGraphSafe -Query $rgQuery -Subscriptions @($subId)

        if ($rgFilter) {
            $allRGs = $allRGs | Where-Object { $_.name -eq $rgFilter }
        }

        Update-Status 'Scanning resources...' 40

        # --- Resources ---
        if ($rgFilter) {
            $safeRGFilter = $rgFilter -replace "['\\\"]", ''
            $resQuery = "resources | where resourceGroup =~ '$safeRGFilter' | project name, type, resourceGroup, location, tags, subscriptionId, id"
        } else {
            $resQuery = "resources | project name, type, resourceGroup, location, tags, subscriptionId, id"
        }
        $allResources = Search-AzGraphSafe -Query $resQuery -Subscriptions @($subId)

        $script:AllRGs       = $allRGs
        $script:AllResources = $allResources

        Update-Status 'Building tag summary...' 70

        # --- Required tags ---
        $requiredTags = @()
        $rtText = $ui.RequiredTagsInput.Text.Trim()
        if ($rtText) {
            $requiredTags = $rtText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }

        # --- Tag key inventory ---
        $tagKeyCount = @{}
        $rgSorted    = [System.Collections.Generic.List[PSObject]]::new()
        $untaggedRGs = 0

        foreach ($rg in $allRGs) {
            $tagMap = ConvertTo-TagHashtable $rg.tags

            $missingKeys = @()
            foreach ($rt in $requiredTags) {
                if (-not $tagMap.ContainsKey($rt)) { $missingKeys += $rt }
            }

            foreach ($k in $tagMap.Keys) {
                if (-not $tagKeyCount.ContainsKey($k)) { $tagKeyCount[$k] = 0 }
                $tagKeyCount[$k]++
            }

            if ($tagMap.Count -eq 0) { $untaggedRGs++ }

            $rgSorted.Add([PSCustomObject]@{
                Name          = $rg.name
                Location      = $rg.location
                TagCount      = $tagMap.Count
                MissingTags   = if ($missingKeys.Count -gt 0) { $missingKeys -join ', ' } else { '-' }
                Tags          = ($tagMap.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
            })
        }

        # --- Resource rows ---
        $resSorted = [System.Collections.Generic.List[PSObject]]::new()
        $taggedRes = 0
        foreach ($res in $allResources) {
            $tagMap = ConvertTo-TagHashtable $res.tags
            if ($tagMap.Count -gt 0) { $taggedRes++ }

            $shortType = ($res.type -split '/')[-1]
            $resSorted.Add([PSCustomObject]@{
                Name          = $res.name
                Type          = $shortType
                ResourceGroup = $res.resourceGroup
                TagCount      = $tagMap.Count
                Tags          = ($tagMap.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
            })
        }

        # --- Tag summary grid ---
        $tagSummary = $tagKeyCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            [PSCustomObject]@{
                TagKey     = $_.Key
                RGsCovered = $_.Value
                TotalRGs   = $allRGs.Count
                Coverage   = if ($allRGs.Count) { '{0:P0}' -f ($_.Value / $allRGs.Count) } else { '0%' }
            }
        }

        # --- Coverage pct ---
        $coveragePct = if ($allResources.Count) { [math]::Round(($taggedRes / $allResources.Count) * 100, 1) } else { 0 }

        # --- Unique tag keys ---
        $uniqueTags = $tagKeyCount.Keys.Count

        # --- Bind grids ---
        $ui.TagSummaryGrid.ItemsSource = @($tagSummary)
        $ui.RGGrid.ItemsSource         = @($rgSorted)
        $ui.ResourceGrid.ItemsSource   = @($resSorted)

        # --- Summary cards ---
        $ui.RGCountText.Text       = $allRGs.Count.ToString()
        $ui.ResourceCountText.Text = $allResources.Count.ToString()
        $ui.TagCoverageText.Text   = "$coveragePct%"
        $ui.UntaggedRGText.Text    = $untaggedRGs.ToString()
        $ui.UniqueTagsText.Text    = $uniqueTags.ToString()

        $ui.ExportButton.IsEnabled  = $true
        $ui.ApplyTagsButton.IsEnabled = $true
        $ui.ScanButton.IsEnabled    = $true

        Update-Status "Scan complete - $($allRGs.Count) RGs, $($allResources.Count) resources" 100
    }
    catch {
        $ui.ScanButton.IsEnabled = $true
        Update-Status "Scan error: $($_.Exception.Message)" 0
        [System.Windows.MessageBox]::Show(
            "Scan failed:`n$($_.Exception.Message)",
            'Scan Error', 'OK', 'Error') | Out-Null
    }
})

# ─────────────────────────────────────────────────────────────────
# RG FILTER change
# ─────────────────────────────────────────────────────────────────
$ui.RGFilterTag.Add_SelectionChanged({
    if (-not $script:AllRGs -or $script:AllRGs.Count -eq 0) { return }

    $requiredTags = @()
    $rtText = $ui.RequiredTagsInput.Text.Trim()
    if ($rtText) { $requiredTags = $rtText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } }

    $source = $ui.RGGrid.ItemsSource
    if ($ui.RGFilterTag.SelectedIndex -eq 1 -and $requiredTags.Count -gt 0) {
        $ui.RGGrid.ItemsSource = @($source | Where-Object { $_.MissingTags -ne '-' })
    } else {
        # Re-scan to reload full list
        $ui.ScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    }
})

# ─────────────────────────────────────────────────────────────────
# Resource filter: enable tag name textbox when 'Missing Specific Tag'
# ─────────────────────────────────────────────────────────────────
$ui.ResFilterTag.Add_SelectionChanged({
    $ui.ResFilterTagName.IsEnabled = ($ui.ResFilterTag.SelectedIndex -eq 2)

    if (-not $script:AllResources -or $script:AllResources.Count -eq 0) { return }

    switch ($ui.ResFilterTag.SelectedIndex) {
        1 { # Untagged
            $ui.ResourceGrid.ItemsSource = @($ui.ResourceGrid.ItemsSource | Where-Object { $_.TagCount -eq 0 })
        }
        default {
            # Reload from scan
        }
    }
})

# ─────────────────────────────────────────────────────────────────
# ADD TAG to queue
# ─────────────────────────────────────────────────────────────────
$ui.AddTagButton.Add_Click({
    $tagName  = $ui.ApplyTagName.Text.Trim()
    $tagValue = $ui.ApplyTagValue.Text.Trim()

    if (-not $tagName) {
        [System.Windows.MessageBox]::Show('Tag name is required.', 'Validation', 'OK', 'Warning') | Out-Null
        return
    }

    # Check for duplicate
    $exists = $script:TagQueue | Where-Object { $_.TagName -eq $tagName }
    if ($exists) {
        $result = [System.Windows.MessageBox]::Show(
            "Tag '$tagName' is already in the queue. Replace the value?",
            'Duplicate Tag', 'YesNo', 'Question')
        if ($result -eq 'Yes') {
            $script:TagQueue.Remove($exists)
        } else { return }
    }

    $script:TagQueue.Add([PSCustomObject]@{ TagName = $tagName; TagValue = $tagValue })

    $ui.ApplyTagName.Text  = ''
    $ui.ApplyTagValue.Text = ''
    $ui.ApplyTagName.Focus()
})

# ─────────────────────────────────────────────────────────────────
# CLEAR / REMOVE tag queue items
# ─────────────────────────────────────────────────────────────────
$ui.ClearTagsButton.Add_Click({
    $script:TagQueue.Clear()
})

$ui.RemoveTagButton.Add_Click({
    $sel = $ui.TagQueueGrid.SelectedItem
    if ($sel) { $script:TagQueue.Remove($sel) }
})

# ─────────────────────────────────────────────────────────────────
# APPLY TAGS
# ─────────────────────────────────────────────────────────────────
$ui.ApplyTagsButton.Add_Click({
    if ($script:TagQueue.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Add at least one tag to the queue first.', 'No Tags', 'OK', 'Warning') | Out-Null
        return
    }

    $isDryRun  = $ui.DryRunCheck.IsChecked
    $overwrite = $ui.OverwriteCheck.IsChecked
    $scopeIdx  = $ui.ApplyScope.SelectedIndex
    $subIdx    = $ui.SubscriptionSelector.SelectedIndex
    if ($subIdx -lt 0) { return }
    $sub = $script:Subscriptions[$subIdx]

    # Build tag hashtable from queue
    $tagsToApply = @{}
    foreach ($t in $script:TagQueue) {
        $tagsToApply[$t.TagName] = $t.TagValue
    }

    $modeLabel = if ($isDryRun) { 'DRY RUN' } else { 'LIVE' }

    if (-not $isDryRun) {
        $confirm = [System.Windows.MessageBox]::Show(
            "You are about to apply $($tagsToApply.Count) tag(s) to resources in '$($sub.Name)'.`n`nThis is a LIVE operation. Continue?",
            'Confirm Tag Application', 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') { return }
    }

    try {
        $ui.ApplyTagsButton.IsEnabled = $false
        $results = [System.Collections.Generic.List[PSObject]]::new()

        # Determine targets
        $targets = @()
        switch ($scopeIdx) {
            0 { # All RGs in scope
                $targets = $script:AllRGs | ForEach-Object {
                    [PSCustomObject]@{ Id = $_.id; Name = $_.name; Kind = 'ResourceGroup' }
                }
            }
            1 { # RGs missing the first queued tag
                $firstTag = $script:TagQueue[0].TagName
                foreach ($rg in $script:AllRGs) {
                    $tagMap = ConvertTo-TagHashtable $rg.tags
                    if (-not $tagMap.ContainsKey($firstTag)) {
                        $targets += [PSCustomObject]@{ Id = $rg.id; Name = $rg.name; Kind = 'ResourceGroup' }
                    }
                }
            }
            2 { # All resources in scope
                $targets = $script:AllResources | ForEach-Object {
                    [PSCustomObject]@{ Id = $_.id; Name = $_.name; Kind = ($_.type -split '/')[-1] }
                }
            }
            3 { # Untagged resources
                foreach ($res in $script:AllResources) {
                    $tagMap = ConvertTo-TagHashtable $res.tags
                    if ($tagMap.Count -eq 0) {
                        $targets += [PSCustomObject]@{ Id = $res.id; Name = $res.name; Kind = ($res.type -split '/')[-1] }
                    }
                }
            }
        }

        $total = $targets.Count
        $done  = 0

        foreach ($target in $targets) {
            $done++
            $pct = [math]::Round(($done / [math]::Max($total,1)) * 100)
            Update-Status "[$modeLabel] Tagging $done / $total - $($target.Name)" $pct

            $status = 'Success'
            $detail = ''

            try {
                if ($isDryRun) {
                    $status = 'DryRun'
                    $detail = ($tagsToApply.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                } else {
                    # Get current tags
                    $resource  = Get-AzTag -ResourceId $target.Id -ErrorAction Stop
                    $existing  = @{}
                    if ($resource.Properties -and $resource.Properties.TagsProperty) {
                        foreach ($kv in $resource.Properties.TagsProperty.GetEnumerator()) {
                            $existing[$kv.Key] = $kv.Value
                        }
                    }

                    $merged = @{}
                    foreach ($kv in $existing.GetEnumerator()) { $merged[$kv.Key] = $kv.Value }

                    $applied  = @()
                    $skipped  = @()
                    foreach ($kv in $tagsToApply.GetEnumerator()) {
                        if ($merged.ContainsKey($kv.Key) -and -not $overwrite) {
                            $skipped += $kv.Key
                        } else {
                            $merged[$kv.Key] = $kv.Value
                            $applied += $kv.Key
                        }
                    }

                    if ($applied.Count -gt 0) {
                        Update-AzTag -ResourceId $target.Id -Tag $merged -Operation Merge -ErrorAction Stop | Out-Null
                        $detail = "Applied: $($applied -join ', ')"
                        if ($skipped.Count -gt 0) { $detail += " | Skipped (exists): $($skipped -join ', ')" }
                    } else {
                        $status = 'Skipped'
                        $detail = "All tags already exist"
                    }
                }
            }
            catch {
                $status = 'Error'
                $detail = $_.Exception.Message
            }

            $results.Add([PSCustomObject]@{
                Resource = $target.Name
                Kind     = $target.Kind
                Status   = $status
                Detail   = $detail
            })
        }

        $ui.ApplyResultsGrid.ItemsSource = @($results)

        $successCount = ($results | Where-Object { $_.Status -in 'Success','DryRun' }).Count
        $errorCount   = ($results | Where-Object { $_.Status -eq 'Error' }).Count
        $ui.ApplyStatusText.Text = "$modeLabel complete - $successCount succeeded, $errorCount failed out of $total"

        $ui.ApplyTagsButton.IsEnabled = $true
        Update-Status "$modeLabel tagging complete - $total targets processed" 100
    }
    catch {
        $ui.ApplyTagsButton.IsEnabled = $true
        Update-Status "Apply error: $($_.Exception.Message)" 0
        [System.Windows.MessageBox]::Show(
            "Tag application failed:`n$($_.Exception.Message)",
            'Apply Error', 'OK', 'Error') | Out-Null
    }
})

# ─────────────────────────────────────────────────────────────────
# EXPORT TO CSV
# ─────────────────────────────────────────────────────────────────
$ui.ExportButton.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter   = 'CSV Files (*.csv)|*.csv'
    $dlg.FileName = "AzureTagReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog()) {
        try {
            $export = [System.Collections.Generic.List[PSObject]]::new()
            foreach ($rg in $script:AllRGs) {
                $tagMap = ConvertTo-TagHashtable $rg.tags
                $export.Add([PSCustomObject]@{
                    Scope         = 'ResourceGroup'
                    Name          = $rg.name
                    Location      = $rg.location
                    ResourceGroup = $rg.name
                    Type          = 'microsoft.resources/subscriptions/resourcegroups'
                    TagCount      = $tagMap.Count
                    Tags          = ($tagMap.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                })
            }
            foreach ($res in $script:AllResources) {
                $tagMap = ConvertTo-TagHashtable $res.tags
                $export.Add([PSCustomObject]@{
                    Scope         = 'Resource'
                    Name          = $res.name
                    Location      = $res.location
                    ResourceGroup = $res.resourceGroup
                    Type          = $res.type
                    TagCount      = $tagMap.Count
                    Tags          = ($tagMap.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                })
            }
            $export | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            Update-Status "Exported $($export.Count) rows to $($dlg.FileName)" 100
        }
        catch {
            [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
        }
    }
})

# ─────────────────────────────────────────────────────────────────
# Show window
# ─────────────────────────────────────────────────────────────────
$window.ShowDialog() | Out-Null

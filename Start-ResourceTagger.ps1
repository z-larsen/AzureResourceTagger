<#
.SYNOPSIS
    Azure Resource Tagger - Scan existing tags and bulk-apply new tags to Azure resources.

.DESCRIPTION
    A WPF-based tool that connects to your Azure tenant, scans resource groups and
    resources for existing tags, identifies tagging gaps, and lets you bulk-apply
    tags at scale.  Designed for governance and compliance workflows such as
    preparing a subscription for Azure Policy tag enforcement.

.NOTES
    Version : 1.2.0
    Author  : Zac Larsen
    Requires: Az.Accounts, Az.Resources, Az.ResourceGraph
#>

#Requires -Version 5.1

param(
    [switch]$Debug
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptRoot = $PSScriptRoot
$script:Version    = '1.2.0'

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
    'ResTagSelectedButton','ResOverwriteCheck','ResTagStatusText',
    'ApplyTagName','ApplyTagValue','AddTagButton',
    'TagQueueGrid','ClearTagsButton','RemoveTagButton',
    'ApplyScope','OverwriteCheck','DryRunCheck',
    'ApplyTagsButton','ApplyStatusText','ApplyResultsGrid',
    'RemoveTagSelector','RemoveTagValueFilter','RemoveTagValuePlaceholder',
    'RefreshTagListButton','RemoveScope','RemoveDryRunCheck',
    'RemoveTagsButton','RemoveStatusText','RemoveResultsGrid',
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
$script:LastScanSubIdx  = -1
$script:LastScanScope   = -1
$script:LastScanRG      = ''
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
# Helper: Flush WPF dispatcher (process pending UI messages)
# ─────────────────────────────────────────────────────────────────
$script:FlushingUI = $false
function Flush-UI {
    if ($script:FlushingUI) { return }   # prevent reentrancy
    $script:FlushingUI = $true
    try {
        $frame = [System.Windows.Threading.DispatcherFrame]::new()
        $window.Dispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [Action]{ $frame.Continue = $false }
        )
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
    } catch {}
    $script:FlushingUI = $false
}

# ─────────────────────────────────────────────────────────────────
# Helper: Update status bar
# ─────────────────────────────────────────────────────────────────
function Update-Status {
    param([string]$Message, [int]$Progress = -1)
    $ui.StatusText.Text = $Message
    if ($Progress -ge 0) { $ui.ProgressBar.Value = $Progress }
    Flush-UI
}

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

# Helper: Safely get the tags property from a Resource Graph object
function Get-SafeTags {
    param($Obj)
    if ($null -eq $Obj) { return $null }
    if ($Obj.PSObject.Properties.Match('tags').Count -gt 0) { return $Obj.tags }
    return $null
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

        Flush-UI
        $result = Search-AzGraph @params
        Flush-UI

        # Newer Az.ResourceGraph returns PSResourceGraphResponse with .Data
        # Older versions return the array of rows directly
        $hasData = $result.PSObject.Properties.Match('Data').Count -gt 0
        if ($hasData -and $null -ne $result.Data) {
            foreach ($r in $result.Data) {
                if ($null -ne $r) { $all.Add($r) }
            }
        } else {
            # $result itself is the iterable set of rows
            foreach ($r in $result) {
                if ($null -ne $r -and $r.PSObject.Properties.Match('id').Count -gt 0) {
                    $all.Add($r)
                }
            }
        }

        $hasSkip = $result.PSObject.Properties.Match('SkipToken').Count -gt 0
        $skip = if ($hasSkip) { $result.SkipToken } else { $null }
    } while ($skip)

    return $all
}

# ─────────────────────────────────────────────────────────────────
# Tenant picker dialog
# ─────────────────────────────────────────────────────────────────
function Show-TenantPicker {
    param([object[]]$Tenants)

    $pickerXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Select Tenant" Width="520" Height="420"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#F0F0F0" FontFamily="Segoe UI">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Select the tenant to use:" FontSize="14" FontWeight="SemiBold"
                   Foreground="#333" Margin="0,0,0,12"/>
        <ListBox Grid.Row="1" Name="TenantList" FontSize="13" Margin="0,0,0,12"
                 BorderBrush="#CCC" BorderThickness="1"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="OkBtn" Content="Select" Width="90" Height="32" FontSize="13" FontWeight="SemiBold"
                    Background="#0078D4" Foreground="White" BorderThickness="0" Margin="0,0,8,0" IsEnabled="False"/>
            <Button Name="CancelBtn" Content="Cancel" Width="90" Height="32" FontSize="13"
                    Background="White" Foreground="#333" BorderBrush="#CCC" BorderThickness="1"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $rdr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($pickerXaml))
    $dlg = [System.Windows.Markup.XamlReader]::Load($rdr)

    $list      = $dlg.FindName('TenantList')
    $okBtn     = $dlg.FindName('OkBtn')
    $cancelBtn = $dlg.FindName('CancelBtn')

    foreach ($t in $Tenants) {
        $display = if ($t.Name -and $t.Name -ne $t.TenantId) {
            "$($t.Name)  ($($t.TenantId))"
        } else {
            "$($t.TenantId)"
        }
        $item = [System.Windows.Controls.ListBoxItem]::new()
        $item.Content = $display
        $item.Tag = $t.TenantId
        $list.Items.Add($item) | Out-Null
    }

    $list.Add_SelectionChanged({ $okBtn.IsEnabled = ($list.SelectedItem -ne $null) })
    $list.Add_MouseDoubleClick({ if ($list.SelectedItem) { $dlg.DialogResult = $true; $dlg.Close() } })
    $okBtn.Add_Click({ $dlg.DialogResult = $true; $dlg.Close() })
    $cancelBtn.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }

    $picked = $dlg.ShowDialog()
    if ($picked -and $list.SelectedItem) {
        return $list.SelectedItem.Tag
    }
    return $null
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

        # List accessible tenants and show picker
        Update-Status 'Loading accessible tenants...' 20
        $tenants = @(Get-AzTenant -ErrorAction SilentlyContinue)

        if ($tenants.Count -eq 0) {
            throw 'No accessible tenants found.'
        }

        $selectedTenantId = Show-TenantPicker -Tenants $tenants
        if (-not $selectedTenantId) {
            Update-Status 'Tenant selection cancelled.' 0
            $btn.Content = "$envLabel Tenant"
            $ui.CommercialButton.IsEnabled = $true
            $ui.GovButton.IsEnabled        = $true
            return
        }

        # Switch tenant if needed
        if ($selectedTenantId -ne $ctx.Tenant.Id) {
            Update-Status "Switching to tenant $selectedTenantId..." 25
            $window.WindowState = 'Minimized'
            try {
                Connect-AzAccount -Environment $AzureEnvironment -TenantId $selectedTenantId -ErrorAction Stop | Out-Null
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
    Flush-UI

    $ui.RGSelector.Items.Clear()
    $ui.RGSelector.Items.Add('(All Resource Groups)') | Out-Null
    try {
        Flush-UI
        $rgs = Get-AzResourceGroup | Sort-Object ResourceGroupName
        Flush-UI
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
        $ui.ScanButton.IsEnabled = $false
        $ui.ApplyTagsButton.IsEnabled = $false
        $ui.ExportButton.IsEnabled = $false

        # Disable scope controls during scan to prevent Flush-UI from
        # triggering cascading SelectionChanged handlers
        $ui.SubscriptionSelector.IsEnabled = $false
        $ui.ScopeLevel.IsEnabled = $false
        $ui.RGSelector.IsEnabled = $false

        # Clear previous results so WPF isn't rendering stale data during scan
        $ui.TagSummaryGrid.ItemsSource = $null
        $ui.RGGrid.ItemsSource         = $null
        $ui.ResourceGrid.ItemsSource   = $null
        $ui.RGCountText.Text       = '-'
        $ui.ResourceCountText.Text = '-'
        $ui.TagCoverageText.Text   = '-'
        $ui.UntaggedRGText.Text    = '-'
        $ui.UniqueTagsText.Text    = '-'

        Update-Status 'Scanning resource groups...' 10
        Flush-UI

        # --- Resource Groups ---
        $rgFilter = $null
        if ($ui.ScopeLevel.SelectedIndex -eq 1 -and $ui.RGSelector.SelectedIndex -gt 0) {
            $rgFilter = $ui.RGSelector.SelectedItem.ToString()
        }

        $rgQuery = "resourcecontainers | where type == 'microsoft.resources/subscriptions/resourcegroups' | project name, id, location, tags, subscriptionId"
        if ($rgFilter) {
            $safeRGName = $rgFilter -replace "[`'`"]", ''
            $rgQuery = "resourcecontainers | where type == 'microsoft.resources/subscriptions/resourcegroups' | where name =~ '$safeRGName' | project name, id, location, tags, subscriptionId"
        }
        $allRGs  = Search-AzGraphSafe -Query $rgQuery -Subscriptions @($subId)
        Flush-UI

        Update-Status 'Scanning resources...' 40
        Flush-UI

        # --- Resources ---
        if ($rgFilter) {
            $safeRGFilter = $rgFilter -replace "[`'`"]", ''
            $resQuery = "resources | where resourceGroup =~ `'$safeRGFilter`' | project name, type, resourceGroup, location, tags, subscriptionId, id"
        } else {
            $resQuery = "resources | project name, type, resourceGroup, location, tags, subscriptionId, id"
        }
        $allResources = Search-AzGraphSafe -Query $resQuery -Subscriptions @($subId)
        Flush-UI

        $script:AllRGs       = $allRGs
        $script:AllResources = $allResources

        Update-Status 'Building tag summary...' 70
        Flush-UI

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
            if ($rg.PSObject.Properties.Match('name').Count -eq 0) { continue }
            $tagMap = ConvertTo-TagHashtable (Get-SafeTags $rg)

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
            if ($res.PSObject.Properties.Match('name').Count -eq 0) { continue }
            $tagMap = ConvertTo-TagHashtable (Get-SafeTags $res)
            if ($tagMap.Count -gt 0) { $taggedRes++ }

            $shortType = ($res.type -split '/')[-1]
            $resSorted.Add([PSCustomObject]@{
                Name          = $res.name
                Type          = $shortType
                ResourceGroup = $res.resourceGroup
                TagCount      = $tagMap.Count
                Tags          = ($tagMap.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                ResourceId    = $res.id
            })
        }

        # --- Tag summary grid ---
        $tagSummary = $tagKeyCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            [PSCustomObject]@{
                TagKey     = $_.Key
                RGsCovered = $_.Value
                TotalRGs   = @($allRGs).Count
                Coverage   = if (@($allRGs).Count) { '{0:P0}' -f ($_.Value / @($allRGs).Count) } else { '0%' }
            }
        }

        # --- Coverage pct ---
        $coveragePct = if (@($allResources).Count) { [math]::Round(($taggedRes / @($allResources).Count) * 100, 1) } else { 0 }

        # --- Unique tag keys ---
        $uniqueTags = $tagKeyCount.Keys.Count

        # --- Bind grids ---
        Update-Status 'Updating display...' 85
        Flush-UI
        $ui.TagSummaryGrid.ItemsSource = @($tagSummary)
        Flush-UI
        $ui.RGGrid.ItemsSource         = @($rgSorted)
        Flush-UI
        $ui.ResourceGrid.ItemsSource   = @($resSorted)
        Flush-UI

        # --- Summary cards ---
        $ui.RGCountText.Text       = @($allRGs).Count.ToString()
        $ui.ResourceCountText.Text = @($allResources).Count.ToString()
        $ui.TagCoverageText.Text   = "$coveragePct%"
        $ui.UntaggedRGText.Text    = $untaggedRGs.ToString()
        $ui.UniqueTagsText.Text    = $uniqueTags.ToString()

        $ui.ExportButton.IsEnabled  = $true
        $ui.ApplyTagsButton.IsEnabled = $true
        $ui.ScanButton.IsEnabled    = $true
        $ui.SubscriptionSelector.IsEnabled = $true
        $ui.ScopeLevel.IsEnabled    = $true
        $ui.RGSelector.IsEnabled    = ($ui.ScopeLevel.SelectedIndex -eq 1)
        Flush-UI

        # Auto-populate Remove Tags dropdown
        Update-Status 'Discovering tag keys...' 90
        Flush-UI
        $ui.RemoveTagSelector.Items.Clear()
        $removeTagKeys = @{}
        foreach ($rg in $script:AllRGs) {
            $tagMap = ConvertTo-TagHashtable (Get-SafeTags $rg)
            foreach ($k in $tagMap.Keys) { $removeTagKeys[$k] = $true }
        }
        foreach ($res in $script:AllResources) {
            $tagMap = ConvertTo-TagHashtable (Get-SafeTags $res)
            foreach ($k in $tagMap.Keys) { $removeTagKeys[$k] = $true }
        }
        # Also pull from ARM tags API (catches resources ARG doesn't index)
        try {
            Flush-UI
            $armTags = Get-AzTag -ErrorAction SilentlyContinue
            Flush-UI
            foreach ($t in $armTags) {
                if ($t.TagName) { $removeTagKeys[$t.TagName] = $true }
            }
        } catch {}
        foreach ($k in ($removeTagKeys.Keys | Sort-Object)) {
            $ui.RemoveTagSelector.Items.Add($k) | Out-Null
        }
        if ($ui.RemoveTagSelector.Items.Count -gt 0) {
            $ui.RemoveTagSelector.SelectedIndex = 0
        }
        $ui.RemoveTagsButton.IsEnabled = ($ui.RemoveTagSelector.Items.Count -gt 0)
        Flush-UI

        Update-Status "Scan complete - $(@($allRGs).Count) RGs, $(@($allResources).Count) resources" 100
        Flush-UI

        # Record scan scope for stale-data detection
        $script:LastScanSubIdx = $ui.SubscriptionSelector.SelectedIndex
        $script:LastScanScope  = $ui.ScopeLevel.SelectedIndex
        $script:LastScanRG     = if ($ui.ScopeLevel.SelectedIndex -eq 1 -and $ui.RGSelector.SelectedIndex -gt 0) { $ui.RGSelector.SelectedItem.ToString() } else { '' }
    }
    catch {
        $ui.ScanButton.IsEnabled = $true
        $ui.SubscriptionSelector.IsEnabled = $true
        $ui.ScopeLevel.IsEnabled = $true
        $ui.RGSelector.IsEnabled = ($ui.ScopeLevel.SelectedIndex -eq 1)
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
    if (-not $script:AllRGs -or @($script:AllRGs).Count -eq 0) { return }

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

    if (-not $script:AllResources -or @($script:AllResources).Count -eq 0) { return }

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
# RESOURCES TAB - Enable/disable Tag Selected button on selection
# ─────────────────────────────────────────────────────────────────
$ui.ResourceGrid.Add_SelectionChanged({
    $ui.ResTagSelectedButton.IsEnabled = ($ui.ResourceGrid.SelectedItems.Count -gt 0)
    $ui.ResTagStatusText.Text = "$($ui.ResourceGrid.SelectedItems.Count) selected"
})

# ─────────────────────────────────────────────────────────────────
# RESOURCES TAB - Apply tag to selected resources inline
# ─────────────────────────────────────────────────────────────────
$ui.ResTagSelectedButton.Add_Click({
    $selected = @($ui.ResourceGrid.SelectedItems)
    if ($selected.Count -eq 0) { return }

    # Pop a small dialog asking for tag name and value
    $tagDlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Apply Tag" Width="400" Height="220"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#F0F0F0" FontFamily="Segoe UI">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="80"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Row="0" Grid.Column="0" Text="Tag Name:" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8"/>
        <TextBox Grid.Row="0" Grid.Column="1" Name="TagNameBox" FontSize="13" Padding="4" Margin="0,0,0,8"/>
        <TextBlock Grid.Row="1" Grid.Column="0" Text="Tag Value:" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8"/>
        <TextBox Grid.Row="1" Grid.Column="1" Name="TagValueBox" FontSize="13" Padding="4" Margin="0,0,0,8"/>
        <TextBlock Grid.Row="2" Grid.ColumnSpan="2" Name="InfoLabel" FontSize="12" Foreground="#666"
                   Margin="0,0,0,4"/>
        <StackPanel Grid.Row="4" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="ApplyBtn" Content="Apply" Width="90" Height="32" FontSize="13" FontWeight="SemiBold"
                    Background="#107C10" Foreground="White" BorderThickness="0" Margin="0,0,8,0"/>
            <Button Name="CancelBtn" Content="Cancel" Width="90" Height="32" FontSize="13"
                    Background="White" Foreground="#333" BorderBrush="#CCC" BorderThickness="1"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $rdr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($tagDlgXaml))
    $tagDlg = [System.Windows.Markup.XamlReader]::Load($rdr)

    $tagNameBox  = $tagDlg.FindName('TagNameBox')
    $tagValueBox = $tagDlg.FindName('TagValueBox')
    $infoLabel   = $tagDlg.FindName('InfoLabel')
    $applyBtn    = $tagDlg.FindName('ApplyBtn')
    $cancelDlgBtn = $tagDlg.FindName('CancelBtn')

    $infoLabel.Text = "Applying to $($selected.Count) resource(s)"

    $applyBtn.Add_Click({ $tagDlg.DialogResult = $true; $tagDlg.Close() }.GetNewClosure())
    $cancelDlgBtn.Add_Click({ $tagDlg.DialogResult = $false; $tagDlg.Close() }.GetNewClosure())

    $result = $tagDlg.ShowDialog()
    if (-not $result) { return }

    $tagName  = $tagNameBox.Text.Trim()
    $tagValue = $tagValueBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($tagName)) {
        [System.Windows.MessageBox]::Show('Tag name cannot be empty.', 'Validation', 'OK', 'Warning') | Out-Null
        return
    }

    $overwrite = $ui.ResOverwriteCheck.IsChecked
    $successCount = 0
    $skipCount    = 0
    $errorCount   = 0
    $total        = $selected.Count

    foreach ($resObj in $selected) {
        try {
            Update-Status "Tagging $($resObj.Name)..." ([math]::Round(($successCount + $skipCount + $errorCount) / [math]::Max($total,1) * 100))
            Flush-UI

            $resId = $resObj.ResourceId
            Flush-UI
            $resource = Get-AzTag -ResourceId $resId -ErrorAction Stop
            Flush-UI
            $existing = @{}
            if ($resource.Properties -and $resource.Properties.TagsProperty) {
                foreach ($kv in $resource.Properties.TagsProperty.GetEnumerator()) {
                    $existing[$kv.Key] = $kv.Value
                }
            }

            if ($existing.ContainsKey($tagName) -and -not $overwrite) {
                $skipCount++
            } else {
                $tagHash = @{ $tagName = $tagValue }
                Update-AzTag -ResourceId $resId -Tag $tagHash -Operation Merge -ErrorAction Stop | Out-Null
                Flush-UI
                $successCount++
            }
            Flush-UI
        } catch {
            $errorCount++
            $lastErr = $_.Exception.Message
        }
    }

    $statusMsg = "Done: $successCount applied, $skipCount skipped, $errorCount errors"
    if ($errorCount -gt 0 -and $lastErr) { $statusMsg += " - $lastErr" }
    $ui.ResTagStatusText.Text = $statusMsg
    Update-Status "Tag applied to $successCount resources" 100
})

# ─────────────────────────────────────────────────────────────────
# APPLY TAGS
# ─────────────────────────────────────────────────────────────────
$ui.ApplyTagsButton.Add_Click({
    $isDryRun  = $ui.DryRunCheck.IsChecked
    $overwrite = $ui.OverwriteCheck.IsChecked
    $scopeIdx  = $ui.ApplyScope.SelectedIndex
    $subIdx    = $ui.SubscriptionSelector.SelectedIndex
    if ($subIdx -lt 0) { return }
    $sub = $script:Subscriptions[$subIdx]

    # Warn if scope has changed since last scan
    $currentRG = if ($ui.ScopeLevel.SelectedIndex -eq 1 -and $ui.RGSelector.SelectedIndex -gt 0) { $ui.RGSelector.SelectedItem.ToString() } else { '' }
    $scopeChanged = ($subIdx -ne $script:LastScanSubIdx) -or
                    ($ui.ScopeLevel.SelectedIndex -ne $script:LastScanScope) -or
                    ($currentRG -ne $script:LastScanRG)
    if ($scopeChanged -and $script:LastScanSubIdx -ge 0) {
        $warn = [System.Windows.MessageBox]::Show(
            "The scope has changed since the last scan. The tags will be applied based on the previous scan data.`n`nRe-scan first to pick up the new scope, or click Yes to continue with the existing scan data.",
            'Scope Changed', 'YesNo', 'Warning')
        if ($warn -ne 'Yes') { return }
    }

    # ── Selected Resource Groups picker mode ────────────────
    if ($scopeIdx -eq 4) {
        if ($script:TagQueue.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Add at least one tag to the queue first.', 'No Tags', 'OK', 'Warning') | Out-Null
            return
        }
        if (-not $script:AllRGs -or @($script:AllRGs).Count -eq 0) {
            [System.Windows.MessageBox]::Show('Run a scan first so resource groups are available.', 'No Scan Data', 'OK', 'Warning') | Out-Null
            return
        }

        # Build picker dialog
        $pickerXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Select Resource Groups" Width="560" Height="520"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#F0F0F0" FontFamily="Segoe UI">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Select the resource groups to tag:" FontSize="14" FontWeight="SemiBold"
                   Foreground="#333" Margin="0,0,0,8"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <Button Name="SelectAllBtn" Content="Select All" Width="90" Height="28" FontSize="12"
                    Background="White" Foreground="#0078D4" BorderBrush="#0078D4" BorderThickness="1" Margin="0,0,8,0"/>
            <Button Name="SelectNoneBtn" Content="Select None" Width="90" Height="28" FontSize="12"
                    Background="White" Foreground="#0078D4" BorderBrush="#0078D4" BorderThickness="1"/>
        </StackPanel>
        <ListBox Grid.Row="2" Name="RGList" FontSize="13" Margin="0,0,0,12"
                 BorderBrush="#CCC" BorderThickness="1" SelectionMode="Extended"/>
        <TextBlock Grid.Row="3" Name="CountLabel" Text="0 selected" FontSize="12" Foreground="#666" Margin="0,0,0,8"/>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="OkBtn" Content="Apply" Width="90" Height="32" FontSize="13" FontWeight="SemiBold"
                    Background="#107C10" Foreground="White" BorderThickness="0" Margin="0,0,8,0" IsEnabled="False"/>
            <Button Name="CancelBtn" Content="Cancel" Width="90" Height="32" FontSize="13"
                    Background="White" Foreground="#333" BorderBrush="#CCC" BorderThickness="1"/>
        </StackPanel>
    </Grid>
</Window>
"@

        $rdr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($pickerXaml))
        $dlg = [System.Windows.Markup.XamlReader]::Load($rdr)

        $rgList       = $dlg.FindName('RGList')
        $okBtn        = $dlg.FindName('OkBtn')
        $cancelBtn    = $dlg.FindName('CancelBtn')
        $selectAllBtn = $dlg.FindName('SelectAllBtn')
        $selectNoneBtn = $dlg.FindName('SelectNoneBtn')
        $countLabel   = $dlg.FindName('CountLabel')

        foreach ($rg in ($script:AllRGs | Sort-Object { $_.name })) {
            $item = [System.Windows.Controls.ListBoxItem]::new()
            $item.Content = $rg.name
            $item.Tag = $rg
            $rgList.Items.Add($item) | Out-Null
        }

        $rgList.Add_SelectionChanged({
            $count = $rgList.SelectedItems.Count
            $countLabel.Text = "$count selected"
            $okBtn.IsEnabled = ($count -gt 0)
        }.GetNewClosure())

        $selectAllBtn.Add_Click({ $rgList.SelectAll() }.GetNewClosure())
        $selectNoneBtn.Add_Click({ $rgList.UnselectAll() }.GetNewClosure())
        $okBtn.Add_Click({ $dlg.DialogResult = $true; $dlg.Close() }.GetNewClosure())
        $cancelBtn.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() }.GetNewClosure())

        $picked = $dlg.ShowDialog()
        if (-not $picked -or $rgList.SelectedItems.Count -eq 0) { return }

        $selectedRGs = @($rgList.SelectedItems | ForEach-Object { $_.Tag })

        $tagsToApply = @{}
        foreach ($t in $script:TagQueue) { $tagsToApply[$t.TagName] = $t.TagValue }

        $modeLabel = if ($isDryRun) { 'DRY RUN' } else { 'LIVE' }
        $rgCount   = @($selectedRGs).Count

        if ($isDryRun) {
            $msg = "Preview applying $($tagsToApply.Count) tag(s) to $rgCount resource group(s).`n`nProceed with dry run?"
            $confirm = [System.Windows.MessageBox]::Show($msg, 'Confirm Dry Run', 'YesNo', 'Question')
        } else {
            $msg = "You are about to apply $($tagsToApply.Count) tag(s) to $rgCount resource group(s).`n`nThis is a LIVE operation. Continue?"
            $confirm = [System.Windows.MessageBox]::Show($msg, 'Confirm Tag Application', 'YesNo', 'Warning')
        }
        if ($confirm -ne 'Yes') { return }

        try {
            $ui.ApplyTagsButton.IsEnabled = $false
            $results = [System.Collections.Generic.List[PSObject]]::new()
            $done = 0

            foreach ($rgObj in $selectedRGs) {
                $done++
                $pct = [math]::Round(($done / [math]::Max($rgCount,1)) * 100)
                Update-Status "[$modeLabel] Tagging RG $done / $rgCount - $($rgObj.name)" $pct

                $status = 'Success'
                $detail = ''

                try {
                    if ($isDryRun) {
                        $status = 'DryRun'
                        $detail = ($tagsToApply.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                    } else {
                        $resource = Get-AzTag -ResourceId $rgObj.id -ErrorAction Stop
                        $existing = @{}
                        if ($resource.Properties -and $resource.Properties.TagsProperty) {
                            foreach ($kv in $resource.Properties.TagsProperty.GetEnumerator()) {
                                $existing[$kv.Key] = $kv.Value
                            }
                        }

                        $merged = @{}
                        foreach ($kv in $existing.GetEnumerator()) { $merged[$kv.Key] = $kv.Value }

                        $applied = @()
                        $skipped = @()
                        foreach ($kv in $tagsToApply.GetEnumerator()) {
                            if ($merged.ContainsKey($kv.Key) -and -not $overwrite) {
                                $skipped += $kv.Key
                            } else {
                                $merged[$kv.Key] = $kv.Value
                                $applied += $kv.Key
                            }
                        }

                        if (@($applied).Count -gt 0) {
                            Update-AzTag -ResourceId $rgObj.id -Tag $merged -Operation Merge -ErrorAction Stop | Out-Null
                            $detail = "Applied: $($applied -join ', ')"
                            if (@($skipped).Count -gt 0) { $detail += " | Skipped (exists): $($skipped -join ', ')" }
                        } else {
                            $status = 'Skipped'
                            $detail = 'All tags already exist'
                        }
                    }
                } catch {
                    $status = 'Error'
                    $detail = $_.Exception.Message
                }

                $results.Add([PSCustomObject]@{
                    Resource = $rgObj.name; Kind = 'ResourceGroup'
                    Status = $status; Detail = $detail
                })
                $ui.ApplyResultsGrid.ItemsSource = @($results)
                Flush-UI
            }

            $ui.ApplyResultsGrid.ItemsSource = @($results)
            Flush-UI

            $successCount = @($results | Where-Object { $_.Status -in 'Success','DryRun' }).Count
            $errorCount   = @($results | Where-Object { $_.Status -eq 'Error' }).Count
            $ui.ApplyStatusText.Text = "$modeLabel complete - $successCount succeeded, $errorCount failed out of $rgCount"

            $ui.ApplyTagsButton.IsEnabled = $true
            Update-Status "$modeLabel tagging complete - $rgCount resource groups processed" 100
            Flush-UI
        }
        catch {
            $ui.ApplyTagsButton.IsEnabled = $true
            Update-Status "Apply error: $($_.Exception.Message)" 0
            [System.Windows.MessageBox]::Show(
                "Tag application failed:`n$($_.Exception.Message)",
                'Apply Error', 'OK', 'Error') | Out-Null
        }
        return
    }

    # ── Normal queue-based mode ─────────────────────────────
    if ($script:TagQueue.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Add at least one tag to the queue first.', 'No Tags', 'OK', 'Warning') | Out-Null
        return
    }

    # Build tag hashtable from queue
    $tagsToApply = @{}
    foreach ($t in $script:TagQueue) {
        $tagsToApply[$t.TagName] = $t.TagValue
    }

    $modeLabel = if ($isDryRun) { 'DRY RUN' } else { 'LIVE' }

    if ($isDryRun) {
        $msg = "Preview applying $($tagsToApply.Count) tag(s) to resources in $($sub.Name).`n`nProceed with dry run?"
        $confirm = [System.Windows.MessageBox]::Show(
            $msg, 'Confirm Dry Run', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    } else {
        $msg = "You are about to apply $($tagsToApply.Count) tag(s) to resources in $($sub.Name)." + "`n`nThis is a LIVE operation. Continue?"
        $confirm = [System.Windows.MessageBox]::Show(
            $msg, 'Confirm Tag Application', 'YesNo', 'Warning')
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
                    $tagMap = ConvertTo-TagHashtable (Get-SafeTags $rg)
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
                    $tagMap = ConvertTo-TagHashtable (Get-SafeTags $res)
                    if ($tagMap.Count -eq 0) {
                        $targets += [PSCustomObject]@{ Id = $res.id; Name = $res.name; Kind = ($res.type -split '/')[-1] }
                    }
                }
            }
        }

        $total = @($targets).Count
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

                    if (@($applied).Count -gt 0) {
                        Update-AzTag -ResourceId $target.Id -Tag $merged -Operation Merge -ErrorAction Stop | Out-Null
                        $detail = "Applied: $($applied -join ', ')"
                        if (@($skipped).Count -gt 0) { $detail += " | Skipped (exists): $($skipped -join ', ')" }
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
            $ui.ApplyResultsGrid.ItemsSource = @($results)
            Flush-UI
        }

        $ui.ApplyResultsGrid.ItemsSource = @($results)
        Flush-UI

        $successCount = @($results | Where-Object { $_.Status -in 'Success','DryRun' }).Count
        $errorCount   = @($results | Where-Object { $_.Status -eq 'Error' }).Count
        $ui.ApplyStatusText.Text = "$modeLabel complete - $successCount succeeded, $errorCount failed out of $total"

        $ui.ApplyTagsButton.IsEnabled = $true
        Update-Status "$modeLabel tagging complete - $total targets processed" 100
        Flush-UI
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
# REMOVE TAGS - Placeholder text behavior
# ─────────────────────────────────────────────────────────────────
$ui.RemoveTagValueFilter.Add_GotFocus({
    $ui.RemoveTagValuePlaceholder.Visibility = 'Collapsed'
})
$ui.RemoveTagValueFilter.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($ui.RemoveTagValueFilter.Text)) {
        $ui.RemoveTagValuePlaceholder.Visibility = 'Visible'
    }
})

# ─────────────────────────────────────────────────────────────────
# REMOVE TAGS - Refresh tag list from scan data
# ─────────────────────────────────────────────────────────────────
$ui.RefreshTagListButton.Add_Click({
    $ui.RemoveTagSelector.Items.Clear()
    $tagKeys = @{}

    foreach ($rg in $script:AllRGs) {
        $tagMap = ConvertTo-TagHashtable (Get-SafeTags $rg)
        foreach ($k in $tagMap.Keys) { $tagKeys[$k] = $true }
    }
    foreach ($res in $script:AllResources) {
        $tagMap = ConvertTo-TagHashtable (Get-SafeTags $res)
        foreach ($k in $tagMap.Keys) { $tagKeys[$k] = $true }
    }
    # Also pull from ARM tags API (catches resources ARG doesn't index)
    try {
        $armTags = Get-AzTag -ErrorAction SilentlyContinue
        foreach ($t in $armTags) {
            if ($t.TagName) { $tagKeys[$t.TagName] = $true }
        }
    } catch {}

    foreach ($k in ($tagKeys.Keys | Sort-Object)) {
        $ui.RemoveTagSelector.Items.Add($k) | Out-Null
    }
    if ($ui.RemoveTagSelector.Items.Count -gt 0) {
        $ui.RemoveTagSelector.SelectedIndex = 0
    }
    $ui.RemoveTagsButton.IsEnabled = ($ui.RemoveTagSelector.Items.Count -gt 0)
    Update-Status "Tag list refreshed - $(@($tagKeys.Keys).Count) unique keys found" 100
})

# ─────────────────────────────────────────────────────────────────
# REMOVE TAGS - Execute removal
# ─────────────────────────────────────────────────────────────────
$ui.RemoveTagsButton.Add_Click({
    $tagToRemove = $ui.RemoveTagSelector.Text.Trim()
    if (-not $tagToRemove) {
        [System.Windows.MessageBox]::Show('Select or enter a tag name to remove.', 'No Tag Selected', 'OK', 'Warning') | Out-Null
        return
    }

    $valueFilter = $ui.RemoveTagValueFilter.Text.Trim()
    $isDryRun    = $ui.RemoveDryRunCheck.IsChecked
    $scopeIdx    = $ui.RemoveScope.SelectedIndex
    $modeLabel   = if ($isDryRun) { 'DRY RUN' } else { 'LIVE' }

    $scopeDesc = @('all RGs', 'all resources', 'all RGs and resources')[$scopeIdx]
    if ($isDryRun) {
        $removeMsg = "Preview removing tag '$tagToRemove' from $scopeDesc."
        if ($valueFilter) { $removeMsg += " (only where value = '$valueFilter')" }
        $removeMsg += "`n`nProceed with dry run?"
        $confirm = [System.Windows.MessageBox]::Show($removeMsg, 'Confirm Dry Run', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    } else {
        $removeMsg = "You are about to REMOVE tag '$tagToRemove' from $scopeDesc."
        if ($valueFilter) { $removeMsg += " (only where value = '$valueFilter')" }
        $removeMsg += "`n`nThis is a LIVE operation. Continue?"
        $confirm = [System.Windows.MessageBox]::Show($removeMsg, 'Confirm Tag Removal', 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') { return }
    }

    try {
        $ui.RemoveTagsButton.IsEnabled = $false
        $results = [System.Collections.Generic.List[PSObject]]::new()

        # Build target list based on scope
        $targets = [System.Collections.Generic.List[PSObject]]::new()

        if ($scopeIdx -eq 0 -or $scopeIdx -eq 2) {
            foreach ($rg in $script:AllRGs) {
                $tagMap = ConvertTo-TagHashtable (Get-SafeTags $rg)
                if ($tagMap.ContainsKey($tagToRemove)) {
                    if (-not $valueFilter -or $tagMap[$tagToRemove] -eq $valueFilter) {
                        $targets.Add([PSCustomObject]@{
                            Id = $rg.id; Name = $rg.name; Kind = 'ResourceGroup'
                            CurrentValue = $tagMap[$tagToRemove]
                        })
                    }
                }
            }
        }
        if ($scopeIdx -eq 1 -or $scopeIdx -eq 2) {
            foreach ($res in $script:AllResources) {
                $tagMap = ConvertTo-TagHashtable (Get-SafeTags $res)
                if ($tagMap.ContainsKey($tagToRemove)) {
                    if (-not $valueFilter -or $tagMap[$tagToRemove] -eq $valueFilter) {
                        $targets.Add([PSCustomObject]@{
                            Id = $res.id; Name = $res.name; Kind = ($res.type -split '/')[-1]
                            CurrentValue = $tagMap[$tagToRemove]
                        })
                    }
                }
            }
        }

        $total = @($targets).Count
        $done  = 0

        foreach ($target in $targets) {
            $done++
            $pct = [math]::Round(($done / [math]::Max($total,1)) * 100)
            Update-Status "[$modeLabel] Removing '$tagToRemove' - $done / $total - $($target.Name)" $pct

            $status = 'Success'
            $detail = ''

            try {
                if ($isDryRun) {
                    $status = 'DryRun'
                    $detail = "Would remove $tagToRemove=$($target.CurrentValue)"
                } else {
                    $tagToDelete = @{ $tagToRemove = $target.CurrentValue }
                    Update-AzTag -ResourceId $target.Id -Tag $tagToDelete -Operation Delete -ErrorAction Stop | Out-Null
                    $detail = "Removed $tagToRemove=$($target.CurrentValue)"
                }
            }
            catch {
                $status = 'Error'
                $detail = $_.Exception.Message
            }

            $results.Add([PSCustomObject]@{
                Resource      = $target.Name
                Kind          = $target.Kind
                Status        = $status
                PreviousValue = $target.CurrentValue
                Detail        = $detail
            })
            $ui.RemoveResultsGrid.ItemsSource = @($results)
            Flush-UI
        }

        $ui.RemoveResultsGrid.ItemsSource = @($results)
        Flush-UI

        $successCount = @($results | Where-Object { $_.Status -in 'Success','DryRun' }).Count
        $errorCount   = @($results | Where-Object { $_.Status -eq 'Error' }).Count
        $ui.RemoveStatusText.Text = "$modeLabel complete - $successCount succeeded, $errorCount failed out of $total"

        $ui.RemoveTagsButton.IsEnabled = $true
        Update-Status "$modeLabel removal complete - $total targets processed" 100
        Flush-UI
    }
    catch {
        $ui.RemoveTagsButton.IsEnabled = $true
        Update-Status "Remove error: $($_.Exception.Message)" 0
        [System.Windows.MessageBox]::Show(
            "Tag removal failed:`n$($_.Exception.Message)",
            'Remove Error', 'OK', 'Error') | Out-Null
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
                $tagMap = ConvertTo-TagHashtable (Get-SafeTags $rg)
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
                $tagMap = ConvertTo-TagHashtable (Get-SafeTags $res)
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
            Update-Status "Exported $(@($export).Count) rows to $($dlg.FileName)" 100
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

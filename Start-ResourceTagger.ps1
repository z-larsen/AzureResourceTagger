<#
.SYNOPSIS
    Azure Resource Tagger - Scan existing tags and bulk-apply new tags to Azure resources.

.DESCRIPTION
    A WPF-based tool that connects to your Azure tenant, scans resource groups and
    resources for existing tags, identifies tagging gaps, and lets you bulk-apply
    tags at scale.  Designed for governance and compliance workflows such as
    preparing a subscription for Azure Policy tag enforcement.

.NOTES
    Version : 1.3.0
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
$script:Version    = '1.3.0'
Import-Module (Join-Path $script:ScriptRoot 'ResourceTagger.Core.psm1') -Force

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

[xml]$xaml = Get-Content $xamlPath -Raw -Encoding UTF8
$reader   = New-Object System.Xml.XmlNodeReader $xaml
$window   = [Windows.Markup.XamlReader]::Load($reader)
$reader.Dispose()

# ─────────────────────────────────────────────────────────────────
# Resolve named controls
# ─────────────────────────────────────────────────────────────────
$controlNames = @(
    'VersionLabel','TenantLabel',
    'CommercialButton','GovButton','ScanButton','ExportButton',
    'ScopeLevel','SubscriptionSelector','RGSelector','ScanScopeText',
    'RGCountText','ResourceCountText','TagCoverageText','UntaggedRGText','UniqueTagsText',
    'TagSummaryGrid',
    'RGFilterTag','RequiredTagsInput','RequiredTagsPlaceholder','RGGrid',
    'ResFilterTag','ResFilterTagName','ResourceGrid',
    'ResTagSelectedButton','ResOverwriteCheck','ResDryRunCheck','ResTagStatusText',
    'ApplyTagName','ApplyTagValue','AddTagButton',
    'TagQueueGrid','ClearTagsButton','RemoveTagButton',
    'ApplyScope','OverwriteCheck','DryRunCheck',
    'ApplyTagsButton','ApplyStatusText','ApplyResultsGrid',
    'RemoveTagSelector','RemoveTagValueFilter','RemoveTagValuePlaceholder','RemoveMatchValueCheck',
    'RefreshTagListButton','RemoveScope','RemoveDryRunCheck',
    'RemoveTagsButton','RemoveStatusText','RemoveResultsGrid',
    'ProgressBar','StatusText','MainTabs'
)

$ui = @{}
foreach ($name in $controlNames) {
    $ctrl = $window.FindName($name)
    if (-not $ctrl) { throw "Required UI control '$name' was not found." }
    $ui[$name] = $ctrl
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
$script:ActiveContext   = $null
$script:ScanScope       = $null
$script:ScanContext     = $null
$script:ScannedAt       = $null
$script:Busy            = $false
$script:Scanning        = $false
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
# Helper: Flush WPF dispatcher (process pending UI updates)
# ─────────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
function Flush-UI {
    [System.Windows.Forms.Application]::DoEvents()
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

function Show-TaggerError {
    param([string]$Message, [string]$Title = 'Error')
    [System.Windows.MessageBox]::Show($window, $Message, $Title, 'OK', 'Error') | Out-Null
}

# ─────────────────────────────────────────────────────────────────
# Scope and operation state
# ─────────────────────────────────────────────────────────────────
function Update-ControlState {
    $idle = -not $script:Busy
    $connected = $script:Connected -and $null -ne $script:ActiveContext
    $hasScan = $null -ne $script:ScanScope
    foreach ($name in @(
        'CommercialButton','GovButton','ScopeLevel','RGFilterTag','RequiredTagsInput',
        'ResFilterTag','ResourceGrid','ApplyTagName','ApplyTagValue','AddTagButton','TagQueueGrid',
        'ClearTagsButton','RemoveTagButton','ApplyScope','OverwriteCheck','DryRunCheck',
        'ResOverwriteCheck','ResDryRunCheck','RemoveTagSelector','RemoveMatchValueCheck',
        'RemoveScope','RemoveDryRunCheck'
    )) {
        $ui[$name].IsEnabled = $idle
    }
    $ui.SubscriptionSelector.IsEnabled = $idle -and $connected
    $ui.ScanButton.IsEnabled = $idle -and $connected -and $ui.SubscriptionSelector.SelectedIndex -ge 0
    $ui.RGSelector.IsEnabled = $idle -and $connected -and $ui.ScopeLevel.SelectedIndex -eq 1
    $ui.ResFilterTagName.IsEnabled = $idle -and $ui.ResFilterTag.SelectedIndex -eq 2
    $ui.RemoveTagValueFilter.IsEnabled = $idle -and $ui.RemoveMatchValueCheck.IsChecked
    $ui.RefreshTagListButton.IsEnabled = $idle -and $hasScan
    $ui.ExportButton.IsEnabled = $idle -and $hasScan
    $ui.ApplyTagsButton.IsEnabled = $idle -and $hasScan
    $ui.RemoveTagsButton.IsEnabled = $idle -and $hasScan
    $ui.ResTagSelectedButton.IsEnabled = $idle -and $hasScan -and $ui.ResourceGrid.SelectedItems.Count -gt 0
}

function Set-TaggerBusy {
    param([bool]$Busy)
    $script:Busy = $Busy
    Update-ControlState
}

function Clear-TagScan {
    param([string]$Reason = 'Scan the selected scope before making changes.')
    $script:ScanScope = $null
    $script:ScanContext = $null
    $script:ScannedAt = $null
    $script:AllRGs = @()
    $script:AllResources = @()
    foreach ($name in @('TagSummaryGrid','RGGrid','ResourceGrid')) { $ui[$name].ItemsSource = $null }
    foreach ($name in @('RGCountText','ResourceCountText','TagCoverageText','UntaggedRGText','UniqueTagsText')) {
        $ui[$name].Text = '-'
    }
    $ui.RemoveTagSelector.Items.Clear()
    $ui.RemoveTagSelector.Text = ''
    $ui.ScanScopeText.Text = $Reason
    Update-ControlState
}

function Get-CurrentTagScope {
    $index = $ui.SubscriptionSelector.SelectedIndex
    if (-not $script:Connected -or $null -eq $script:ActiveContext -or $index -lt 0) {
        throw 'Connect and select a subscription before scanning.'
    }
    if ($script:ActiveContext.Subscription.Id -ne $script:Subscriptions[$index].Id) {
        throw 'The Azure context does not match the selected subscription. Reconnect and re-scan.'
    }
    $group = if ($ui.ScopeLevel.SelectedIndex -eq 1 -and $ui.RGSelector.SelectedIndex -gt 0) {
        $ui.RGSelector.SelectedItem.ToString()
    } else { '' }
    return New-TagScope -Context $script:ActiveContext -ResourceGroup $group
}

function Update-ResourceGroupView {
    $required = @(Get-RequiredTagName -Text $ui.RequiredTagsInput.Text)
    $ui.RGGrid.ItemsSource = @(Get-ResourceGroupRow -Resources $script:AllRGs -RequiredTags $required `
        -OnlyMissing:($ui.RGFilterTag.SelectedIndex -eq 1))
}

function Update-ResourceView {
    $filter = @('All','Untagged','MissingTag')[$ui.ResFilterTag.SelectedIndex]
    $ui.ResourceGrid.ItemsSource = @(Get-ResourceRow -Resources $script:AllResources -Filter $filter `
        -TagName $ui.ResFilterTagName.Text.Trim())
    if ($filter -eq 'MissingTag' -and [string]::IsNullOrWhiteSpace($ui.ResFilterTagName.Text)) {
        $ui.ResTagStatusText.Text = 'Enter a tag name to filter.'
    }
}

function Update-RemoveTagList {
    $tagKeys = @{}
    foreach ($resource in @($script:AllRGs) + @($script:AllResources)) {
        $tags = ConvertTo-TagHashtable (Get-SafeTags $resource)
        foreach ($key in $tags.Keys) { $tagKeys[$key] = $true }
    }
    $ui.RemoveTagSelector.Items.Clear()
    foreach ($key in ($tagKeys.Keys | Sort-Object)) { [void]$ui.RemoveTagSelector.Items.Add($key) }
    if ($ui.RemoveTagSelector.Items.Count -gt 0) { $ui.RemoveTagSelector.SelectedIndex = 0 }
    Update-ControlState
}

# ─────────────────────────────────────────────────────────────────
# Helper: Safe Resource Graph query with paging
# ─────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────
# Run Update-AzTag in a background runspace so the UI stays
# responsive during bulk live tag operations.
# ─────────────────────────────────────────────────────────────────
function New-TagWriteRunspace {
    return [runspacefactory]::CreateRunspace()
}

function Invoke-AzTagUpdateSafe {
    param(
        [string]$ResourceId,
        [hashtable]$Tag,
        [ValidateSet('Merge','Delete')][string]$Operation = 'Delete',
        [Parameter(Mandatory)]$DefaultProfile,
        [int]$TimeoutSeconds = 60
    )
    $rs = New-TagWriteRunspace
    $ps = [powershell]::Create()
    try {
        $rs.Open()
        $ps.Runspace = $rs
        [void]$ps.AddScript({
            param($rid, $tag, $op, $azureProfile)
            Update-AzTag -ResourceId $rid -Tag $tag -Operation $op -DefaultProfile $azureProfile -ErrorAction Stop | Out-Null
        }).AddArgument($ResourceId).AddArgument($Tag).AddArgument($Operation).AddArgument($DefaultProfile)

        $asyncResult = $ps.BeginInvoke()
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        while (-not $asyncResult.IsCompleted -and (Get-Date) -lt $deadline) {
            Flush-UI
            Start-Sleep -Milliseconds 100
        }
        if ($asyncResult.IsCompleted) {
            $ps.EndInvoke($asyncResult) | Out-Null
            if ($ps.Streams.Error.Count -gt 0) { throw $ps.Streams.Error[0].Exception }
        } else {
            $ps.Stop()
            throw [System.TimeoutException]::new("Update-AzTag timed out after $($TimeoutSeconds)s")
        }
    } finally {
        $ps.Dispose()
        $rs.Dispose()
    }
}

function Search-AzGraphSafe {
    param(
        [string]$Query,
        [string[]]$Subscriptions,
        [Parameter(Mandatory)]$DefaultProfile
    )
    $all   = [System.Collections.Generic.List[object]]::new()
    $skip  = $null
    $first = 1000

    do {
        $params = @{
            Query        = $Query
            Subscription = $Subscriptions
            First        = $first
            DefaultProfile = $DefaultProfile
            ErrorAction  = 'Stop'
        }
        if ($skip) { $params['SkipToken'] = $skip }

        Flush-UI
        $result = Search-AzGraph @params
        Flush-UI

        if ($null -eq $result) { break }
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

    return $all.ToArray()
}

# ─────────────────────────────────────────────────────────────────
# Shared preview and execution for every tagging entry point
# ─────────────────────────────────────────────────────────────────
function New-TagPreviewWindow {
    param([object[]]$Plan, $Scope, [bool]$DryRun)
    [xml]$previewXaml = Get-Content (Join-Path $script:ScriptRoot 'gui\TagPreview.xaml') -Raw -Encoding UTF8
    $previewReader = [System.Xml.XmlNodeReader]::new($previewXaml)
    try { $dialog = [Windows.Markup.XamlReader]::Load($previewReader) }
    finally { $previewReader.Dispose() }
    $dialog.Owner = $window
    $mode = if ($DryRun) { 'DRY RUN - no changes will be written' } else { 'LIVE - review before applying' }
    $group = if ($Scope.ResourceGroup) { $Scope.ResourceGroup } else { '(all resource groups)' }
    $dialog.FindName('ModeText').Text = $mode
    $dialog.FindName('ScopeText').Text = "$($Scope.Environment) | Tenant: $($Scope.TenantId)`nSubscription: $($Scope.SubscriptionId) | RG: $group`nAccount: $($Scope.AccountId) | Scanned: $($script:ScannedAt.ToString('u'))"
    $rows = @(Get-TagPlanRow -Plan $Plan)
    $dialog.FindName('PlanGrid').ItemsSource = $rows
    $changes = @($rows | Where-Object { $_.Status -eq 'Planned' }).Count
    $skipped = @($rows | Where-Object { $_.Status -eq 'Skipped' }).Count
    $errors = @($Plan | Where-Object { $_.Status -eq 'Error' }).Count
    $dialog.FindName('SummaryText').Text = "$changes tag change(s) planned, $skipped skipped across $($Plan.Count) target(s); $errors target(s) could not be read."
    $apply = $dialog.FindName('ApplyButton')
    $close = $dialog.FindName('CloseButton')
    $apply.Content = "Apply $changes change(s)"
    $apply.IsEnabled = -not $DryRun -and $changes -gt 0
    if ($DryRun) { $apply.Visibility = 'Collapsed'; $close.Content = 'Close preview' }
    $apply.Add_Click({ $dialog.DialogResult = $true }.GetNewClosure())
    $close.Add_Click({ $dialog.DialogResult = $false }.GetNewClosure())
    return $dialog
}

function Show-TagPlanPreview {
    param([object[]]$Plan, $Scope, [bool]$DryRun)
    $dialog = New-TagPreviewWindow -Plan $Plan -Scope $Scope -DryRun $DryRun
    return $dialog.ShowDialog() -eq $true
}

function Invoke-TagWorkflow {
    param(
        [object[]]$Targets,
        [hashtable]$Tags,
        [ValidateSet('Merge','Delete')][string]$Operation = 'Merge',
        [bool]$DryRun = $true,
        [bool]$Overwrite = $false,
        [bool]$MatchValue = $false,
        [ValidateSet('All','Untagged','MissingTag')][string]$Selection = 'All',
        [string]$SelectionTagName = '',
        $ResultsGrid = $ui.ApplyResultsGrid,
        $StatusControl = $ui.ApplyStatusText
    )
    if ($script:Busy) { return }
    $writeAttempted = $false
    try {
        Assert-TagScope -Expected $script:ScanScope -Actual (Get-CurrentTagScope)
        $scope = $script:ScanScope
        $azureProfile = $script:ScanContext
        Set-TaggerBusy $true
        $targetsToPlan = @($Targets | Sort-Object Id -Unique)
        if ($targetsToPlan.Count -eq 0) {
            [System.Windows.MessageBox]::Show('There are no targets in this scan or selection.', 'No Targets', 'OK', 'Information') | Out-Null
            return
        }
        $plan = @(Get-TagOperationPlan -Targets $targetsToPlan -Scope $scope -DefaultProfile $azureProfile `
            -Tags $Tags -Operation $Operation -Overwrite:$Overwrite -MatchValue:$MatchValue `
            -Selection $Selection -SelectionTagName $SelectionTagName -OnProgress {
                param($done, $total, $name)
                Update-Status "Reading current tags $done / $total - $name" ([math]::Round(100 * $done / $total))
            })
        $ResultsGrid.ItemsSource = @(Get-TagPlanRow -Plan $plan)
        $approved = Show-TagPlanPreview -Plan $plan -Scope $scope -DryRun $DryRun
        if ($DryRun -or -not $approved) {
            $planned = @($plan | Where-Object Status -eq Planned).Count
            $skipped = @($plan | Where-Object Status -eq Skipped).Count
            $errors = @($plan | Where-Object Status -eq Error).Count
            $StatusControl.Text = "No writes - $planned targets planned, $skipped skipped, $errors errors."
            Update-Status 'Preview closed. No changes written.' 100
            return
        }

        Assert-TagScope -Expected $scope -Actual (Get-CurrentTagScope)
        $writeAttempted = $true
        $results = @(Invoke-TagOperationPlan -Plan $plan -Scope $scope -DefaultProfile $azureProfile -Confirm:$false `
            -UpdateAction {
                param($id, $delta, $operation, $context)
                Invoke-AzTagUpdateSafe -ResourceId $id -Tag $delta -Operation $operation -DefaultProfile $context
            } -OnProgress {
                param($done, $total, $name)
                Update-Status "Validating and applying $done / $total - $name" ([math]::Round(100 * $done / $total))
            })
        $ResultsGrid.ItemsSource = @(Get-TagPlanRow -Plan $results)
        $success = @($results | Where-Object Status -eq Success).Count
        $skipped = @($results | Where-Object Status -eq Skipped).Count
        $conflicts = @($results | Where-Object Status -eq Conflict).Count
        $errors = @($results | Where-Object { $_.Status -in 'Error','Unverified' }).Count
        $StatusControl.Text = "$success verified, $skipped skipped, $conflicts conflicts, $errors errors/unverified. Re-scan before continuing."
        Update-Status $StatusControl.Text 100
    } catch {
        Clear-TagScan -Reason 'Operation stopped. Re-scan before continuing.'
        $StatusControl.Text = "Operation stopped: $($_.Exception.Message)"
        Update-Status $StatusControl.Text 0
        Show-TaggerError -Message $StatusControl.Text -Title 'Tag Operation Error'
    } finally {
        if ($writeAttempted) { Clear-TagScan -Reason 'Tags may have changed. Re-scan before another operation or export.' }
        Set-TaggerBusy $false
    }
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

    if ($script:Busy) { return }
    Set-TaggerBusy $true
    $script:Connected = $false
    $script:ActiveContext = $null
    $script:Subscriptions = @()
    Clear-TagScan -Reason 'Connection changed. Select a subscription and re-scan.'
    $ui.SubscriptionSelector.Items.Clear()
    $ui.RGSelector.Items.Clear()
    $ui.TenantLabel.Text = ''
    $ui.CommercialButton.Content = 'Commercial Tenant'
    $ui.GovButton.Content = 'Gov Tenant'

    $envLabel = if ($AzureEnvironment -eq 'AzureUSGovernment') { 'Gov' } else { 'Commercial' }
    $btn      = if ($AzureEnvironment -eq 'AzureUSGovernment') { $ui.GovButton } else { $ui.CommercialButton }
    $btn.Content = "$($script:LockOpen) $envLabel Tenant"

    try {
        Update-Status "Connecting to Azure $envLabel..." 10

        # Disable Az 12+ interactive subscription picker
        $env:AZURE_LOGIN_EXPERIENCE_V2 = 'Off'

        $ctx = Get-AzContext -ErrorAction Stop
        if (-not $ctx -or $ctx.Environment.Name -ne $AzureEnvironment) {
            $window.WindowState = 'Minimized'
            try {
                Connect-AzAccount -Environment $AzureEnvironment -Scope Process -ErrorAction Stop | Out-Null
            } finally {
                $window.WindowState = 'Normal'
                [void]$window.Activate()
            }
            $ctx = Get-AzContext -ErrorAction Stop
        }

        # List accessible tenants and show picker
        Update-Status 'Loading accessible tenants...' 20
        $tenants = @(Get-AzTenant -DefaultProfile $ctx -ErrorAction Stop)

        if ($tenants.Count -eq 0) {
            throw 'No accessible tenants found.'
        }

        $selectedTenantId = Show-TenantPicker -Tenants $tenants
        if (-not $selectedTenantId) {
            Update-Status 'Tenant selection cancelled.' 0
            $btn.Content = "$envLabel Tenant"
            return
        }

        # Switch tenant if needed
        if ($selectedTenantId -ne $ctx.Tenant.Id) {
            Update-Status "Switching to tenant $selectedTenantId..." 25
            $window.WindowState = 'Minimized'
            try {
                Connect-AzAccount -Environment $AzureEnvironment -TenantId $selectedTenantId -Scope Process -ErrorAction Stop | Out-Null
            } finally {
                $window.WindowState = 'Normal'
                [void]$window.Activate()
            }
            $ctx = Get-AzContext -ErrorAction Stop
        }

        if ($ctx.Environment.Name -ne $AzureEnvironment -or $ctx.Tenant.Id -ne $selectedTenantId) {
            throw 'The connected Azure context does not match the selected cloud and tenant.'
        }
        $tenantId = $ctx.Tenant.Id
        $script:ActiveContext = $ctx
        $script:Environment = $AzureEnvironment

        Update-Status 'Listing subscriptions...' 30

        $script:Subscriptions = @(Get-AzSubscription -TenantId $tenantId -DefaultProfile $ctx -ErrorAction Stop |
            Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)
        if ($script:Subscriptions.Count -eq 0) { throw 'No enabled subscriptions were found in this tenant.' }

        $ui.SubscriptionSelector.Items.Clear()
        foreach ($sub in $script:Subscriptions) {
            $item = "$($sub.Name)  ($($sub.Id))"
            $ui.SubscriptionSelector.Items.Add($item) | Out-Null
        }
        $ui.SubscriptionSelector.SelectedIndex = 0
        Select-TaggerSubscription

        $ui.TenantLabel.Text = "Tenant: $tenantId  |  $($ctx.Account.Id)  |  $AzureEnvironment"
        $script:Connected = $true

        $btn.Content = "$($script:LockClosed) $envLabel Tenant"
        $subCount = @($script:Subscriptions).Count
        Update-Status "Connected to $envLabel - $subCount subscriptions found" 100
    }
    catch {
        $script:Connected = $false
        $script:ActiveContext = $null
        $script:Subscriptions = @()
        $ui.SubscriptionSelector.Items.Clear()
        $ui.RGSelector.Items.Clear()
        Update-Status "Connection failed: $($_.Exception.Message)" 0
        $btn.Content = "$envLabel Tenant"
        Show-TaggerError -Message "Failed to connect:`n$($_.Exception.Message)" -Title 'Connection Error'
    } finally {
        Set-TaggerBusy $false
    }
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
function Select-TaggerSubscription {
    $idx = $ui.SubscriptionSelector.SelectedIndex
    if ($idx -lt 0 -or $null -eq $script:ActiveContext) { throw 'Select a connected subscription.' }
    $sub = $script:Subscriptions[$idx]
    Clear-TagScan -Reason 'Subscription changed. Re-scan before making changes.'
    $script:ActiveContext = Set-AzContext -Subscription $sub.Id -Tenant $script:ActiveContext.Tenant.Id `
        -DefaultProfile $script:ActiveContext -Scope Process -ErrorAction Stop
    if ($script:ActiveContext.Subscription.Id -ne $sub.Id) { throw 'Azure did not select the requested subscription.' }
    Flush-UI

    $ui.RGSelector.Items.Clear()
    $ui.RGSelector.Items.Add('(All Resource Groups)') | Out-Null
    $rgs = @(Get-AzResourceGroup -DefaultProfile $script:ActiveContext -ErrorAction Stop | Sort-Object ResourceGroupName)
    Flush-UI
    foreach ($rg in $rgs) { $ui.RGSelector.Items.Add($rg.ResourceGroupName) | Out-Null }
    $ui.RGSelector.SelectedIndex = 0
}

$ui.SubscriptionSelector.Add_SelectionChanged({
    if ($script:Busy) { return }
    try {
        Set-TaggerBusy $true
        Select-TaggerSubscription
        Update-Status 'Subscription selected. Scan to load the inventory.' 0
    } catch {
        $script:Connected = $false
        $script:ActiveContext = $null
        Clear-TagScan -Reason 'Subscription selection failed. Reconnect before scanning.'
        Update-Status "Subscription selection failed: $($_.Exception.Message)" 0
        Show-TaggerError -Message $_.Exception.Message -Title 'Subscription Error'
    } finally {
        Set-TaggerBusy $false
    }
})

# ─────────────────────────────────────────────────────────────────
# Scope level change → enable/disable RG selector
# ─────────────────────────────────────────────────────────────────
$ui.ScopeLevel.Add_SelectionChanged({
    if ($script:Busy) { return }
    Clear-TagScan -Reason 'Scope changed. Re-scan before making changes.'
})
$ui.RGSelector.Add_SelectionChanged({
    if ($script:Busy) { return }
    Clear-TagScan -Reason 'Resource-group selection changed. Re-scan before making changes.'
})

# ─────────────────────────────────────────────────────────────────
# SCAN
# ─────────────────────────────────────────────────────────────────
$ui.ScanButton.Add_Click({
    if ($script:Busy) { return }
    $subIdx = $ui.SubscriptionSelector.SelectedIndex
    if ($subIdx -lt 0) { return }
    $sub = $script:Subscriptions[$subIdx]
    $subId = $sub.Id

    try {
        $script:Scanning = $true
        Set-TaggerBusy $true
        $scope = Get-CurrentTagScope
        $azureProfile = $script:ActiveContext
        Clear-TagScan -Reason 'Scanning the selected Azure scope...'

        Update-Status 'Scanning resource groups...' 10
        Flush-UI

        # --- Resource Groups ---
        $rgFilter = $scope.ResourceGroup

        $rgQuery = "resourcecontainers | where type == 'microsoft.resources/subscriptions/resourcegroups' | project name, id, location, tags, subscriptionId"
        if ($rgFilter) {
            $safeRGName = $rgFilter -replace "[`'`"]", ''
            $rgQuery = "resourcecontainers | where type == 'microsoft.resources/subscriptions/resourcegroups' | where name =~ '$safeRGName' | project name, id, location, tags, subscriptionId"
        }
        $allRGs = @(Search-AzGraphSafe -Query $rgQuery -Subscriptions @($subId) -DefaultProfile $azureProfile)
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
        $allResources = @(Search-AzGraphSafe -Query $resQuery -Subscriptions @($subId) -DefaultProfile $azureProfile)
        Flush-UI
        Assert-TagScope -Expected $scope -Actual (Get-CurrentTagScope)
        Assert-TagScope -Expected $scope -Actual (New-TagScope -Context $azureProfile -ResourceGroup $rgFilter)
        foreach ($target in @($allRGs) + @($allResources)) {
            Assert-TagTargetScope -ResourceId $target.id -Scope $scope
        }

        $script:AllRGs       = $allRGs
        $script:AllResources = $allResources

        Update-Status 'Building tag summary...' 70
        Flush-UI

        # --- Tag key inventory ---
        $tagKeyCount = @{}
        $untaggedRGs = 0

        foreach ($rg in $allRGs) {
            $tagMap = ConvertTo-TagHashtable (Get-SafeTags $rg)

            foreach ($k in $tagMap.Keys) {
                if (-not $tagKeyCount.ContainsKey($k)) { $tagKeyCount[$k] = 0 }
                $tagKeyCount[$k]++
            }

            if ($tagMap.Count -eq 0) { $untaggedRGs++ }
        }

        # --- Resource rows ---
        $taggedRes = 0
        foreach ($res in $allResources) {
            $tagMap = ConvertTo-TagHashtable (Get-SafeTags $res)
            if ($tagMap.Count -gt 0) { $taggedRes++ }
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
        Update-ResourceGroupView
        Flush-UI
        Update-ResourceView
        Flush-UI

        # --- Summary cards ---
        $ui.RGCountText.Text       = @($allRGs).Count.ToString()
        $ui.ResourceCountText.Text = @($allResources).Count.ToString()
        $ui.TagCoverageText.Text   = "$coveragePct%"
        $ui.UntaggedRGText.Text    = $untaggedRGs.ToString()
        $ui.UniqueTagsText.Text    = $uniqueTags.ToString()

        # Auto-populate Remove Tags dropdown
        Update-Status 'Discovering tag keys...' 90
        Flush-UI
        $script:ScanScope = $scope
        $script:ScanContext = $azureProfile
        $script:ScannedAt = Get-Date
        $groupLabel = if ($rgFilter) { $rgFilter } else { '(all resource groups)' }
        $ui.ScanScopeText.Text = "$($scope.Environment) | Tenant: $($scope.TenantId) | Subscription: $($scope.SubscriptionId)`nRG: $groupLabel | Account: $($scope.AccountId) | Scanned: $($script:ScannedAt.ToString('u'))"
        Update-RemoveTagList

        Update-Status "Scan complete - $(@($allRGs).Count) RGs, $(@($allResources).Count) resources" 100
        Flush-UI
    }
    catch {
        Clear-TagScan -Reason 'Scan failed. Correct the error and re-scan before continuing.'
        Update-Status "Scan error: $($_.Exception.Message)" 0
        Show-TaggerError -Message "Scan failed:`n$($_.Exception.Message)" -Title 'Scan Error'
    }
    finally {
        $script:Scanning = $false
        Set-TaggerBusy $false
    }
})

# ─────────────────────────────────────────────────────────────────
# RG FILTER change
# ─────────────────────────────────────────────────────────────────
$ui.RGFilterTag.Add_SelectionChanged({
    if ($script:Busy) { return }
    Update-ResourceGroupView
})
$ui.RequiredTagsInput.Add_TextChanged({
    if ($script:Busy) { return }
    Update-ResourceGroupView
})

# ─────────────────────────────────────────────────────────────────
# Resource filter: enable tag name textbox when 'Missing Specific Tag'
# ─────────────────────────────────────────────────────────────────
$ui.ResFilterTag.Add_SelectionChanged({
    if ($script:Busy) { return }
    Update-ControlState
    Update-ResourceView
})
$ui.ResFilterTagName.Add_TextChanged({
    if ($script:Busy) { return }
    Update-ResourceView
})

# ─────────────────────────────────────────────────────────────────
# ADD TAG to queue
# ─────────────────────────────────────────────────────────────────
$ui.AddTagButton.Add_Click({
    if ($script:Busy) { return }
    $tagName  = $ui.ApplyTagName.Text.Trim()
    $tagValue = $ui.ApplyTagValue.Text

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
    if ($script:Busy) { return }
    $script:TagQueue.Clear()
})

$ui.RemoveTagButton.Add_Click({
    if ($script:Busy) { return }
    $sel = $ui.TagQueueGrid.SelectedItem
    if ($sel) { $script:TagQueue.Remove($sel) }
})

# ─────────────────────────────────────────────────────────────────
# RESOURCES TAB - Enable/disable Tag Selected button on selection
# ─────────────────────────────────────────────────────────────────
$ui.ResourceGrid.Add_SelectionChanged({
    if ($script:Busy) { return }
    Update-ControlState
    $ui.ResTagStatusText.Text = "$($ui.ResourceGrid.SelectedItems.Count) selected"
})

# ─────────────────────────────────────────────────────────────────
# RESOURCES TAB - Apply tag to selected resources inline
# ─────────────────────────────────────────────────────────────────
function Read-ResourceTagInput {
    param([int]$ResourceCount)
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
            <Button Name="ApplyBtn" Content="Preview" Width="90" Height="32" FontSize="13" FontWeight="SemiBold"
                    Background="#107C10" Foreground="White" BorderThickness="0" Margin="0,0,8,0"/>
            <Button Name="CancelBtn" Content="Cancel" Width="90" Height="32" FontSize="13"
                    Background="White" Foreground="#333" BorderBrush="#CCC" BorderThickness="1"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $rdr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($tagDlgXaml))
    $tagDlg = [System.Windows.Markup.XamlReader]::Load($rdr)
    $rdr.Dispose()
    $tagDlg.Owner = $window

    $tagNameBox  = $tagDlg.FindName('TagNameBox')
    $tagValueBox = $tagDlg.FindName('TagValueBox')
    $infoLabel   = $tagDlg.FindName('InfoLabel')
    $applyBtn    = $tagDlg.FindName('ApplyBtn')
    $cancelDlgBtn = $tagDlg.FindName('CancelBtn')

    $infoLabel.Text = "Previewing $ResourceCount resource(s)"

    $applyBtn.Add_Click({ $tagDlg.DialogResult = $true; $tagDlg.Close() }.GetNewClosure())
    $cancelDlgBtn.Add_Click({ $tagDlg.DialogResult = $false; $tagDlg.Close() }.GetNewClosure())

    Set-TaggerBusy $true
    try { $result = $tagDlg.ShowDialog() }
    finally { Set-TaggerBusy $false }
    if (-not $result) { return }

    $tagName  = $tagNameBox.Text.Trim()
    $tagValue = $tagValueBox.Text
    if ([string]::IsNullOrWhiteSpace($tagName)) {
        [System.Windows.MessageBox]::Show('Tag name cannot be empty.', 'Validation', 'OK', 'Warning') | Out-Null
        return
    }

    return [pscustomobject]@{ TagName = $tagName; TagValue = $tagValue }
}

$ui.ResTagSelectedButton.Add_Click({
    if ($script:Busy -or $null -eq $script:ScanScope) { return }
    $selected = @($ui.ResourceGrid.SelectedItems)
    if ($selected.Count -eq 0) { return }
    $tag = Read-ResourceTagInput -ResourceCount $selected.Count
    if ($null -eq $tag) { return }
    $targets = @($selected | ForEach-Object {
        [pscustomobject]@{ Id = $_.ResourceId; Name = $_.Name; Kind = $_.Type }
    })
    Invoke-TagWorkflow -Targets $targets -Tags @{ $tag.TagName = $tag.TagValue } `
        -DryRun $ui.ResDryRunCheck.IsChecked -Overwrite $ui.ResOverwriteCheck.IsChecked
    $ui.ResTagStatusText.Text = $ui.ApplyStatusText.Text
    $ui.MainTabs.SelectedIndex = 3
})

# ─────────────────────────────────────────────────────────────────
# APPLY TAGS
# ─────────────────────────────────────────────────────────────────
function Select-TagResourceGroups {
    param([object[]]$Groups)
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
            <Button Name="OkBtn" Content="Preview" Width="90" Height="32" FontSize="13" FontWeight="SemiBold"
                    Background="#107C10" Foreground="White" BorderThickness="0" Margin="0,0,8,0" IsEnabled="False"/>
            <Button Name="CancelBtn" Content="Cancel" Width="90" Height="32" FontSize="13"
                    Background="White" Foreground="#333" BorderBrush="#CCC" BorderThickness="1"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $rdr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($pickerXaml))
    $dlg = [System.Windows.Markup.XamlReader]::Load($rdr)
    $rdr.Dispose()
    $dlg.Owner = $window

    $rgList       = $dlg.FindName('RGList')
    $okBtn        = $dlg.FindName('OkBtn')
    $cancelBtn    = $dlg.FindName('CancelBtn')
    $selectAllBtn = $dlg.FindName('SelectAllBtn')
    $selectNoneBtn = $dlg.FindName('SelectNoneBtn')
    $countLabel   = $dlg.FindName('CountLabel')

    foreach ($rg in ($Groups | Sort-Object { $_.name })) {
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

    Set-TaggerBusy $true
    try { $picked = $dlg.ShowDialog() }
    finally { Set-TaggerBusy $false }
    if (-not $picked -or $rgList.SelectedItems.Count -eq 0) { return }

    return @($rgList.SelectedItems | ForEach-Object { $_.Tag })
}

function Get-ScanTarget {
    param([ValidateSet('ResourceGroups','Resources','All')][string]$Kind)
    if ($Kind -in 'ResourceGroups','All') {
        foreach ($rg in $script:AllRGs) {
            [pscustomobject]@{ Id = $rg.id; Name = $rg.name; Kind = 'ResourceGroup' }
        }
    }
    if ($Kind -in 'Resources','All') {
        foreach ($resource in $script:AllResources) {
            [pscustomobject]@{ Id = $resource.id; Name = $resource.name; Kind = ($resource.type -split '/')[-1] }
        }
    }
}

$ui.ApplyTagsButton.Add_Click({
    if ($script:Busy -or $null -eq $script:ScanScope) { return }
    if ($script:TagQueue.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Add at least one tag to the queue first.', 'No Tags', 'OK', 'Warning') | Out-Null
        return
    }

    # Build tag hashtable from queue
    $tagsToApply = @{}
    foreach ($t in $script:TagQueue) {
        $tagsToApply[$t.TagName] = $t.TagValue
    }

    $selection = 'All'
    $targets = @()
    switch ($ui.ApplyScope.SelectedIndex) {
        0 { $targets = @(Get-ScanTarget -Kind ResourceGroups) }
        1 {
            $targets = @(Get-ScanTarget -Kind ResourceGroups)
            $selection = 'MissingTag'
        }
        2 { $targets = @(Get-ScanTarget -Kind Resources) }
        3 {
            $targets = @(Get-ScanTarget -Kind Resources)
            $selection = 'Untagged'
        }
        4 {
            $groups = @(Select-TagResourceGroups -Groups $script:AllRGs)
            if ($groups.Count -eq 0) { return }
            $targets = @($groups | ForEach-Object {
                [pscustomobject]@{ Id = $_.id; Name = $_.name; Kind = 'ResourceGroup' }
            })
        }
    }
    Invoke-TagWorkflow -Targets $targets -Tags $tagsToApply -Selection $selection `
        -SelectionTagName $script:TagQueue[0].TagName -DryRun $ui.DryRunCheck.IsChecked `
        -Overwrite $ui.OverwriteCheck.IsChecked
})

# ─────────────────────────────────────────────────────────────────
# REMOVE TAGS - Placeholder text behavior
# ─────────────────────────────────────────────────────────────────
$ui.RemoveTagValueFilter.Add_GotFocus({
    $ui.RemoveTagValuePlaceholder.Visibility = 'Collapsed'
})
$ui.RemoveTagValueFilter.Add_LostFocus({
    if ([string]::IsNullOrEmpty($ui.RemoveTagValueFilter.Text)) {
        $ui.RemoveTagValuePlaceholder.Visibility = 'Visible'
    }
})
$ui.RemoveMatchValueCheck.Add_Click({ Update-ControlState })

# ─────────────────────────────────────────────────────────────────
# REMOVE TAGS - Refresh tag list from scan data
# ─────────────────────────────────────────────────────────────────
$ui.RefreshTagListButton.Add_Click({
    if ($script:Busy -or $null -eq $script:ScanScope) { return }
    Update-RemoveTagList
    Update-Status "Tag list refreshed from the scan - $($ui.RemoveTagSelector.Items.Count) unique keys found" 100
})

# ─────────────────────────────────────────────────────────────────
# REMOVE TAGS - Execute removal
# ─────────────────────────────────────────────────────────────────
$ui.RemoveTagsButton.Add_Click({
    if ($script:Busy -or $null -eq $script:ScanScope) { return }
    $tagToRemove = $ui.RemoveTagSelector.Text.Trim()
    if (-not $tagToRemove) {
        [System.Windows.MessageBox]::Show('Select or enter a tag name to remove.', 'No Tag Selected', 'OK', 'Warning') | Out-Null
        return
    }

    $kind = @('ResourceGroups','Resources','All')[$ui.RemoveScope.SelectedIndex]
    $targets = @(Get-ScanTarget -Kind $kind)
    Invoke-TagWorkflow -Targets $targets -Tags @{ $tagToRemove = $ui.RemoveTagValueFilter.Text } `
        -Operation Delete -MatchValue $ui.RemoveMatchValueCheck.IsChecked -DryRun $ui.RemoveDryRunCheck.IsChecked `
        -ResultsGrid $ui.RemoveResultsGrid -StatusControl $ui.RemoveStatusText
})

# ─────────────────────────────────────────────────────────────────
# EXPORT TO CSV
# ─────────────────────────────────────────────────────────────────
$ui.ExportButton.Add_Click({
    if ($script:Busy -or $null -eq $script:ScanScope) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter   = 'CSV Files (*.csv)|*.csv'
    $dlg.FileName = "AzureTagReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog()) {
        try {
            Assert-TagScope -Expected $script:ScanScope -Actual (Get-CurrentTagScope)
            Set-TaggerBusy $true
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
            $export | ConvertTo-SpreadsheetSafeRow | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            Update-Status "Exported $($export.Count) rows to $($dlg.FileName). Formula-like cells are prefixed with [text]." 100
        }
        catch {
            Show-TaggerError -Message "Export failed: $($_.Exception.Message)"
        } finally {
            Set-TaggerBusy $false
        }
    }
})

# ─────────────────────────────────────────────────────────────────
# Show window
# ─────────────────────────────────────────────────────────────────
$window.Add_Closing({
    param($sourceWindow, [System.ComponentModel.CancelEventArgs]$closingEvent)
    if ($script:Busy) {
        $closingEvent.Cancel = $true
        $sourceWindow.FindName('StatusText').Text = 'An operation is active. Wait for it to finish before closing.'
    }
})
Update-ControlState
$window.ShowDialog() | Out-Null

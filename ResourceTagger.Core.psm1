#Requires -Version 5.1
Set-StrictMode -Version Latest

function ConvertTo-TagHashtable {
    param($Tags)
    $map = @{}
    if ($null -eq $Tags) { return $map }
    if ($Tags -is [System.Collections.IDictionary]) {
        foreach ($key in $Tags.Keys) { $map[[string]$key] = [string]$Tags[$key] }
    } elseif ($Tags -is [PSCustomObject]) {
        foreach ($property in $Tags.PSObject.Properties) { $map[$property.Name] = [string]$property.Value }
    } else {
        throw [System.ArgumentException]::new('Tags must be a dictionary or a tag property object.')
    }
    return $map
}

function Get-SafeTags {
    param($Obj)
    if ($null -eq $Obj) { return $null }
    if ($Obj.PSObject.Properties.Match('tags').Count -gt 0) { return $Obj.tags }
    return $null
}

function New-TagScope {
    param(
        [Parameter(Mandatory)]$Context,
        [string]$ResourceGroup = ''
    )
    $scope = [pscustomobject]@{
        Environment   = [string]$Context.Environment.Name
        TenantId      = [string]$Context.Tenant.Id
        SubscriptionId = [string]$Context.Subscription.Id
        AccountId     = [string]$Context.Account.Id
        ResourceGroup = $ResourceGroup
    }
    foreach ($name in @('Environment', 'TenantId', 'SubscriptionId', 'AccountId')) {
        if ([string]::IsNullOrWhiteSpace($scope.$name)) {
            throw "Azure context is missing $name. Reconnect and re-scan before continuing."
        }
    }
    return $scope
}

function Test-TagScopeMatch {
    param($Expected, $Actual)
    if ($null -eq $Expected -or $null -eq $Actual) { return $false }
    foreach ($name in @('Environment', 'TenantId', 'SubscriptionId', 'AccountId', 'ResourceGroup')) {
        if ($Expected.PSObject.Properties.Match($name).Count -eq 0 -or
            $Actual.PSObject.Properties.Match($name).Count -eq 0 -or
            -not [string]::Equals($Expected.$name, $Actual.$name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }
    }
    return $true
}

function Assert-TagScope {
    param($Expected, $Actual)
    if (-not (Test-TagScopeMatch -Expected $Expected -Actual $Actual)) {
        throw 'The cloud, account, tenant, subscription, or resource-group scope has changed. Re-scan before continuing.'
    }
}

function Assert-TagTargetScope {
    param(
        [Parameter(Mandatory)][string]$ResourceId,
        [Parameter(Mandatory)]$Scope
    )
    $subscriptionPrefix = "/subscriptions/$($Scope.SubscriptionId)/"
    if (-not $ResourceId.StartsWith($subscriptionPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Target '$ResourceId' is outside the scanned subscription."
    }
    if ($Scope.ResourceGroup) {
        $groupId = "$($subscriptionPrefix)resourceGroups/$($Scope.ResourceGroup)"
        if (-not [string]::Equals($ResourceId.TrimEnd('/'), $groupId, [System.StringComparison]::OrdinalIgnoreCase) -and
            -not $ResourceId.StartsWith("$groupId/", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Target '$ResourceId' is outside the scanned resource group."
        }
    }
}

function Get-RequiredTagName {
    param([string]$Text = '')
    $seen = @{}
    foreach ($part in ($Text -split ',')) {
        $name = $part.Trim()
        if ($name -and -not $seen.ContainsKey($name)) {
            $seen[$name] = $true
            $name
        }
    }
}

function Get-ResourceGroupRow {
    param(
        [object[]]$Resources = @(),
        [string[]]$RequiredTags = @(),
        [switch]$OnlyMissing
    )
    foreach ($resource in $Resources) {
        $tags = ConvertTo-TagHashtable (Get-SafeTags $resource)
        $missing = @($RequiredTags | Where-Object { -not $tags.ContainsKey($_) })
        if ($OnlyMissing -and $RequiredTags.Count -gt 0 -and $missing.Count -eq 0) { continue }
        [pscustomobject]@{
            Name        = $resource.name
            Location    = $resource.location
            TagCount    = $tags.Count
            MissingTags = if ($missing.Count) { $missing -join ', ' } else { '-' }
            Tags        = ($tags.GetEnumerator() | Sort-Object Key | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
        }
    }
}

function Get-ResourceRow {
    param(
        [object[]]$Resources = @(),
        [ValidateSet('All', 'Untagged', 'MissingTag')][string]$Filter = 'All',
        [string]$TagName = ''
    )
    foreach ($resource in $Resources) {
        $tags = ConvertTo-TagHashtable (Get-SafeTags $resource)
        if ($Filter -eq 'Untagged' -and $tags.Count -gt 0) { continue }
        if ($Filter -eq 'MissingTag' -and $TagName -and $tags.ContainsKey($TagName)) { continue }
        [pscustomobject]@{
            Name          = $resource.name
            Type          = ($resource.type -split '/')[-1]
            ResourceGroup = $resource.resourceGroup
            TagCount      = $tags.Count
            Tags          = ($tags.GetEnumerator() | Sort-Object Key | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
            ResourceId    = $resource.id
        }
    }
}

function New-TagPlanEntry {
    param($Target, $Scope, [string]$Operation, [string]$Selection)
    [pscustomobject]@{
        ResourceId = [string]$Target.Id
        Resource   = [string]$Target.Name
        Kind       = [string]$Target.Kind
        Scope      = $Scope.PSObject.Copy()
        Operation  = $Operation
        Selection  = $Selection
        Changes    = @()
        Payload    = @{}
        Status     = 'Skipped'
        Detail     = ''
    }
}

function New-TagChangePlan {
    param(
        [Parameter(Mandatory)]$Target,
        [Parameter(Mandatory)]$Scope,
        [Parameter(Mandatory)][hashtable]$CurrentTags,
        [Parameter(Mandatory)][hashtable]$Tags,
        [ValidateSet('Merge', 'Delete')][string]$Operation = 'Merge',
        [switch]$Overwrite,
        [switch]$MatchValue,
        [ValidateSet('All', 'Untagged', 'MissingTag')][string]$Selection = 'All',
        [string]$SelectionTagName = ''
    )
    Assert-TagTargetScope -ResourceId $Target.Id -Scope $Scope
    if ($Tags.Count -eq 0) { throw 'At least one tag must be specified.' }
    if ($Selection -eq 'MissingTag' -and [string]::IsNullOrWhiteSpace($SelectionTagName)) {
        throw 'The missing-tag scope requires a tag name.'
    }

    $entry = New-TagPlanEntry -Target $Target -Scope $Scope -Operation $Operation -Selection $Selection
    $changes = [System.Collections.Generic.List[object]]::new()
    $selectionReason = ''
    if ($Selection -eq 'Untagged' -and $CurrentTags.Count -gt 0) { $selectionReason = 'Resource is no longer untagged.' }
    if ($Selection -eq 'MissingTag' -and $CurrentTags.ContainsKey($SelectionTagName)) {
        $selectionReason = "The first queued tag '$SelectionTagName' already exists."
    }

    foreach ($key in ($Tags.Keys | Sort-Object)) {
        if ([string]::IsNullOrWhiteSpace($key)) { throw 'Tag names cannot be empty.' }
        $exists = $CurrentTags.ContainsKey($key)
        $previous = if ($exists) { [string]$CurrentTags[$key] } else { $null }
        $next = if ($Operation -eq 'Merge') { [string]$Tags[$key] } else { $null }
        $action = 'Skip'
        $reason = $selectionReason

        if (-not $reason) {
            if ($Operation -eq 'Delete') {
                if (-not $exists) { $reason = 'Tag does not exist.' }
                elseif ($MatchValue -and $previous -cne [string]$Tags[$key]) { $reason = 'Value does not match exactly.' }
                else { $action = 'Remove'; $entry.Payload[$key] = $previous }
            } elseif ($exists -and $previous -ceq $next) {
                $reason = 'Value is unchanged.'
            } elseif ($exists -and -not $Overwrite) {
                $reason = 'Tag exists; overwrite is disabled.'
            } else {
                $action = if ($exists) { 'Change' } else { 'Add' }
                $entry.Payload[$key] = $next
            }
        }

        $changes.Add([pscustomobject]@{
            TagName        = [string]$key
            PreviousExists = $exists
            PreviousValue  = $previous
            NewValue       = $next
            Action         = $action
            Reason         = $reason
        })
    }
    $entry.Changes = $changes.ToArray()
    if ($entry.Payload.Count -gt 0) {
        $entry.Status = 'Planned'
        $entry.Detail = "$($entry.Payload.Count) tag change(s) planned."
    } else {
        $entry.Detail = 'No changes required.'
    }
    return $entry
}

function Get-ResourceTagMap {
    param([string]$ResourceId, $DefaultProfile)
    $resource = Get-AzTag -ResourceId $ResourceId -DefaultProfile $DefaultProfile -ErrorAction Stop
    if ($null -eq $resource -or $resource.PSObject.Properties.Match('Properties').Count -eq 0 -or
        $null -eq $resource.Properties -or $resource.Properties.PSObject.Properties.Match('TagsProperty').Count -eq 0) {
        throw "Azure returned an invalid tag response for '$ResourceId'."
    }
    return ConvertTo-TagHashtable $resource.Properties.TagsProperty
}

function Get-TagOperationPlan {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Targets,
        [Parameter(Mandatory)]$Scope,
        [Parameter(Mandatory)]$DefaultProfile,
        [Parameter(Mandatory)][hashtable]$Tags,
        [ValidateSet('Merge', 'Delete')][string]$Operation = 'Merge',
        [switch]$Overwrite,
        [switch]$MatchValue,
        [ValidateSet('All', 'Untagged', 'MissingTag')][string]$Selection = 'All',
        [string]$SelectionTagName = '',
        [scriptblock]$OnProgress
    )
    Assert-TagScope -Expected $Scope -Actual (New-TagScope -Context $DefaultProfile -ResourceGroup $Scope.ResourceGroup)
    $done = 0
    foreach ($target in $Targets) {
        $done++
        if ($OnProgress) { & $OnProgress $done $Targets.Count $target.Name | Out-Null }
        try {
            Assert-TagScope -Expected $Scope -Actual (New-TagScope -Context $DefaultProfile -ResourceGroup $Scope.ResourceGroup)
            Assert-TagTargetScope -ResourceId $target.Id -Scope $Scope
            $current = Get-ResourceTagMap -ResourceId $target.Id -DefaultProfile $DefaultProfile
            New-TagChangePlan -Target $target -Scope $Scope -CurrentTags $current -Tags $Tags `
                -Operation $Operation -Overwrite:$Overwrite -MatchValue:$MatchValue `
                -Selection $Selection -SelectionTagName $SelectionTagName
        } catch {
            $entry = New-TagPlanEntry -Target $target -Scope $Scope -Operation $Operation -Selection $Selection
            $entry.Status = 'Error'
            $entry.Detail = $_.Exception.Message
            $entry
        }
    }
}

function Invoke-TagOperationPlan {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Plan,
        [Parameter(Mandatory)]$Scope,
        [Parameter(Mandatory)]$DefaultProfile,
        [scriptblock]$UpdateAction,
        [scriptblock]$OnProgress
    )
    Assert-TagScope -Expected $Scope -Actual (New-TagScope -Context $DefaultProfile -ResourceGroup $Scope.ResourceGroup)
    $done = 0
    foreach ($planned in $Plan) {
        $done++
        $entry = $planned.PSObject.Copy()
        if ($OnProgress) { & $OnProgress $done $Plan.Count $entry.Resource | Out-Null }
        if ($entry.Status -ne 'Planned') { $entry; continue }
        $writeStarted = $false
        try {
            Assert-TagScope -Expected $entry.Scope -Actual $Scope
            Assert-TagScope -Expected $Scope -Actual (New-TagScope -Context $DefaultProfile -ResourceGroup $Scope.ResourceGroup)
            Assert-TagTargetScope -ResourceId $entry.ResourceId -Scope $Scope
            if (-not $PSCmdlet.ShouldProcess($entry.ResourceId, "$($entry.Operation) $($entry.Payload.Count) tag(s)")) {
                $entry.Status = 'Skipped'
                $entry.Detail = 'Execution was not approved (WhatIf or confirmation declined).'
                $entry
                continue
            }

            $current = Get-ResourceTagMap -ResourceId $entry.ResourceId -DefaultProfile $DefaultProfile
            $changed = @($entry.Changes | Where-Object { $_.Action -ne 'Skip' })
            $conflicts = @($changed | Where-Object {
                $present = $current.ContainsKey($_.TagName)
                $present -ne $_.PreviousExists -or ($present -and [string]$current[$_.TagName] -cne $_.PreviousValue)
            })
            if ($conflicts.Count -gt 0 -or ($entry.Selection -eq 'Untagged' -and $current.Count -gt 0)) {
                $entry.Status = 'Conflict'
                $entry.Detail = 'Tags changed after preview. No write attempted; re-scan and preview again.'
                $entry
                continue
            }

            # Merge only the planned delta; unrelated tags may have changed since the read.
            $writeStarted = $true
            if ($UpdateAction) {
                & $UpdateAction $entry.ResourceId $entry.Payload $entry.Operation $DefaultProfile | Out-Null
            } else {
                Update-AzTag -ResourceId $entry.ResourceId -Tag $entry.Payload -Operation $entry.Operation `
                    -DefaultProfile $DefaultProfile -ErrorAction Stop | Out-Null
            }
            $verified = Get-ResourceTagMap -ResourceId $entry.ResourceId -DefaultProfile $DefaultProfile
            $unverified = @($changed | Where-Object {
                if ($_.Action -eq 'Remove') { $verified.ContainsKey($_.TagName) }
                else { -not $verified.ContainsKey($_.TagName) -or [string]$verified[$_.TagName] -cne $_.NewValue }
            })
            if ($unverified.Count -gt 0) {
                $entry.Status = 'Unverified'
                $entry.Detail = 'The requested result could not be verified. Re-scan before retrying.'
            } else {
                $entry.Status = 'Success'
                $entry.Detail = 'Tag changes verified.'
            }
        } catch {
            $entry.Status = if ($writeStarted) { 'Unverified' } else { 'Error' }
            $entry.Detail = $_.Exception.Message
            if ($writeStarted) { $entry.Detail += ' The write outcome is uncertain; re-scan before retrying.' }
        }
        $entry
    }
}

function Get-TagPlanRow {
    param([object[]]$Plan = @())
    foreach ($entry in $Plan) {
        if ($entry.Changes.Count -eq 0) {
            [pscustomobject]@{
                Resource = $entry.Resource; Kind = $entry.Kind; TagName = ''; Before = ''; After = ''
                Action = 'None'; Status = $entry.Status; Detail = $entry.Detail; ResourceId = $entry.ResourceId
            }
        }
        foreach ($change in $entry.Changes) {
            $before = if (-not $change.PreviousExists) { '(absent)' }
                elseif ($change.PreviousValue -ceq '') { '(empty)' } else { $change.PreviousValue }
            $after = if ($change.Action -eq 'Skip') { $before }
                elseif ($change.Action -eq 'Remove') { '(absent)' }
                elseif ($change.NewValue -ceq '') { '(empty)' } else { $change.NewValue }
            [pscustomobject]@{
                Resource = $entry.Resource
                Kind = $entry.Kind
                TagName = $change.TagName
                Before = $before
                After = $after
                Action = $change.Action
                Status = if ($change.Action -eq 'Skip') { 'Skipped' } else { $entry.Status }
                Detail = if ($change.Action -eq 'Skip') { $change.Reason } else { $entry.Detail }
                ResourceId = $entry.ResourceId
            }
        }
    }
}

function ConvertTo-SpreadsheetSafeRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)][psobject]$InputObject
    )
    process {
        $row = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $value = $property.Value
            if ($value -is [string] -and
                ($value -match '^[\s\p{C}'']*[=+\-@\uFF1D\uFF0B\uFF0D\uFF20]' -or $value -match '^[\t\r\n]')) {
                # A visible text prefix survives spreadsheet CSV save/reopen, unlike quote-only escaping.
                $value = '[text] ' + $value
            }
            $row[$property.Name] = $value
        }
        [pscustomobject]$row
    }
}

Export-ModuleMember -Function ConvertTo-TagHashtable, Get-SafeTags, New-TagScope, Test-TagScopeMatch,
    Assert-TagScope, Assert-TagTargetScope, Get-RequiredTagName, Get-ResourceGroupRow, Get-ResourceRow,
    New-TagChangePlan, Get-TagOperationPlan, Invoke-TagOperationPlan, Get-TagPlanRow, ConvertTo-SpreadsheetSafeRow

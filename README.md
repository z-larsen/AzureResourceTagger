# Azure Resource Tagger

![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)
![Azure Az Modules](https://img.shields.io/badge/Azure-Az%20Modules-0078D4?logo=microsoftazure&logoColor=white)
![License MIT](https://img.shields.io/badge/License-MIT-green)
![Version 1.3.0](https://img.shields.io/badge/Version-1.3.0-brightgreen)

Find tagging gaps, review proposed changes, and apply or remove tags across Azure
resource groups and resources.

Azure Resource Tagger is a Windows desktop tool built with PowerShell and Windows
Presentation Foundation (WPF). The workflow is straightforward: connect, scan,
preview, then confirm. Dry run is on by default for every tagging workflow.

Start with the [quick start](#quick-start), check the [known limits](#know-the-limits),
or read the [release notes](CHANGELOG.md).

## Why this exists

Azure Policy can enforce tags, and applicable **modify** policies can update
existing resources through [remediation tasks](https://learn.microsoft.com/azure/governance/policy/how-to/remediate-resources).
It's great for ongoing governance once the tag foundation has been set. 

But what if you don't have IaC fully implemented yet? What if your environment is large and your tagging spans years of deployments where some resources were tagged, some weren't or some were tagged using older standards you no longer follow?

This tool was made for that interactive cleanup. Use Azure Policy for ongoing governance.
Tag cleanup is easier when you can see what needs attention before making a bulk
change. This tool gives you that view, whether you're filling gaps, updating
existing values, or preparing a subscription for Azure Policy tag enforcement.


## Prerequisites

- Windows with Windows PowerShell 5.1 or PowerShell 7. WPF doesn't run on macOS or Linux.
- The `Az.Accounts`, `Az.Resources`, and `Az.ResourceGraph` PowerShell modules.
- An Azure account with access to the subscription you want to scan.

You don't need local administrator rights. The app runs in your user context.

### Required permissions

| Task | Azure role |
|------|------------|
| Scan a subscription | **Reader** on the subscription |
| Apply or remove tags | **Tag Contributor** on the target scope, in addition to the read access needed for scanning |

These roles are a starting point, not a way around resource locks or policy.
Use the narrowest scope that covers your work.

### Cloud support

| Cloud | Azure environment |
|-------|-------------------|
| Azure Commercial | `AzureCloud` |
| Azure Government | `AzureUSGovernment` |

## Quick start

Download or clone this repository, then open PowerShell in the repository folder.
Keep the folder structure intact so the startup script can find the core module
and XAML files.

If the required Azure modules aren't installed, install them in the PowerShell
environment you'll use to run the app:

```powershell
Install-Module Az.Accounts, Az.Resources, Az.ResourceGraph -Scope CurrentUser
```

Start the app in single-threaded apartment (STA) mode, which WPF requires.

**Windows PowerShell 5.1:**

```powershell
powershell -NoProfile -STA -File .\Start-ResourceTagger.ps1
```

**PowerShell 7:**

```powershell
pwsh -NoProfile -STA -File .\Start-ResourceTagger.ps1
```

If Windows blocks downloaded files, see [troubleshooting](#troubleshoot-common-problems).

## Scan and review tags

Start with a resource group you know before running a subscription-wide operation.

1. Select **Commercial Tenant** or **Gov Tenant**. Sign in if prompted, then choose a tenant.
2. On **Scope & Scan**, select a subscription. To narrow the scan, set the scope to **Resource Group** and choose a resource group.
3. Select **Scan Tags**.
4. Check the recorded cloud, tenant, subscription, account, resource group, and scan time.
5. Review the summary, then use **Resource Groups** or **Resources** to inspect the inventory.

The scan reads from Azure Resource Graph. It shows the resources returned for
your selected scope and permissions, not a guaranteed inventory of every Azure
resource type.

The coverage percentage means **resources with at least one tag**. It doesn't
mean those resources have every required tag. The tag-key summary and unique-key
count cover resource groups only.

### Find missing tags

| Tab | What to do |
|-----|------------|
| **Resource Groups** | Enter a comma-separated list of required tag names, then select **Missing Any Required Tag**. |
| **Resources** | Select **Untagged Resources**, or select **Missing Specific Tag** and enter one tag name. |

Required-tag names are case-insensitive, and duplicates are ignored. The
required-tag list applies to resource groups; it isn't a compliance check across
all resources.

Filters use the saved scan data. Changing a filter or required-tag name doesn't
make another Azure request. Select **All Resource Groups** or **All Resources**
to restore the full scanned list.

## Apply tags

Preview first. A dry run reads current tags from Azure, but it doesn't write any
changes and has no enabled Apply action in the preview.

### Preview a bulk update

1. On **Apply Tags**, enter a tag name and value, then select **Add Tag**. Repeat for each tag you want to apply.
2. Choose a target scope from the table below.
3. Leave dry run selected. Select **Overwrite existing tag values** only if you intend to replace existing values.
4. Select **Apply Tags**. If you chose **Selected Resource Groups**, select the groups and then select **Preview**.
5. Review the scope, before-and-after values, skipped tags, and any read errors.

| Target scope | What it includes |
|--------------|------------------|
| **All Resource Groups in Scope** | All scanned resource groups, subject to your overwrite setting |
| **RGs Missing the First Queued Tag** | Scanned resource groups that don't have the first tag name in the queue |
| **All Resources in Scope** | All scanned resources, subject to your overwrite setting |
| **All Untagged Resources** | Scanned resources that currently have no tags |
| **Selected Resource Groups** | Resource groups you choose from the scanned inventory |

Here, RG means resource group. The missing-tag scope checks the **first queued
tag**, not whether a resource group is missing any tag in the queue.

The app checks eligibility and tag values against fresh Azure Resource Manager
(ARM) tag reads. It sends only new or changed tag keys in a merge, so unrelated
tags aren't sent back. Unchanged values are skipped, even with overwrite enabled.
Empty values and whitespace are preserved.

### Preview an update to selected resources

1. On **Resources**, select the resources you want to tag. Use Ctrl or Shift to select more than one.
2. Leave **Dry run** selected. Select **Overwrite** only if you intend to replace existing values.
3. Select **Preview Tag**, enter a tag name and value, then select **Preview**.
4. Review the proposed changes. Results appear on **Apply Tags**.

The dry run and overwrite settings on **Resources** are separate from those on
**Apply Tags**.

## Remove tags

1. On **Remove Tags**, select or enter a tag name.
2. Choose whether to match an exact value, as described below.
3. Choose all resource groups, all resources, or both within the scanned scope.
4. Leave dry run selected, then select **Remove Tags**.
5. Review the proposed removals before starting a live run.

**Match exact value** is off by default. With it off, the selected key is eligible
for removal regardless of its value. With it on, matching is case-sensitive and
preserves whitespace. An empty input matches only an empty tag value.

For example, `Production` and `production` aren't the same value. Tag names are
still matched without regard to case.

**Refresh List** rebuilds the tag-name list from the saved scan. It doesn't query
Azure again. You can also enter a tag name manually.

## Run a live operation

The preview is the confirmation step, not a promise that Azure has changed.

1. Close the dry run preview.
2. Clear the dry run checkbox for the operation you want to run.
3. Start the operation again and review the new preview. Confirm only if the scope and proposed changes are right.
4. Read the results, then scan again before another write or export.

The app binds each scan to its cloud, tenant, subscription, account, and resource
group. Changing the connection or scope clears the old scan. During an operation,
controls that could change the scope or start conflicting work are disabled.
Wait for the operation to finish before closing the window.

An approved live run clears the saved inventory, even if some changes couldn't
be completed. Operation results stay visible.

## Understand the results

| Status | What it means | What to do |
|--------|---------------|------------|
| **Planned** | A change is in the preview. It hasn't been written. | Review it before confirming a live run. |
| **Skipped** | No write is planned or performed for this tag under the selected options. | Read the detail, such as an unchanged value, a value mismatch, or overwrite being off. |
| **Success** | The app wrote the change and read it back to verify the result. | Review the result and scan again before continuing. |
| **Conflict** | An affected tag or the target's eligibility changed after preview. No write was attempted for that target. | Scan again and review a new preview. |
| **Error** | Planning or validation failed before a write started. | Read the error and correct the cause. |
| **Unverified** | A write failed or timed out, or the app couldn't confirm the result. | Inspect the resource and scan again before retrying. |

**Unverified doesn't mean nothing changed.** The app doesn't automatically retry
an uncertain write.

## Export a report

After a successful scan, select **Export to CSV** and choose a file location.
The report includes the scanned resource groups and resources, their locations,
resource types, tag counts, and tag text.

Formula-like text cells get a visible `[text] ` prefix. This includes formula
markers hidden behind whitespace or control characters, and full-width variants.
The prefix is an export-only safety measure; it doesn't change the inventory or
your Azure tags.

Leave that prefix in place when viewing or saving the report in a spreadsheet.
Removing it can turn metadata into an active formula. Treat the CSV as a report,
not a file for importing unchanged tag data back into Azure.

Review reports before sharing them. Resource names and tag values can contain
sensitive information.

## Know the limits

- **Scans can lag behind Azure.** Resource Graph is eventually consistent. New resources and recent tag changes might not appear immediately.
- **Previews refresh known targets only.** Fresh ARM reads update tag information for resources already in the scan. They don't discover additional resources.
- **Conflict checks aren't a lock.** The app checks affected tags before writing, but another tool can still change them between that read and the update. Azure tag updates aren't atomic compare-and-swap operations.
- **Azure still enforces its rules.** Permissions, resource locks, policy, supported resource types, and tag limits can prevent changes.
- **Reads can pause the UI.** Inventory and planning reads aren't fully asynchronous. This release doesn't provide a cancellation workflow.
- **Tests aren't production approval.** Validate the tool in an approved test subscription before using it in production.

## Troubleshoot common problems

| Problem | What to check |
|---------|---------------|
| The app reports a missing module. | Install the required module in the same PowerShell environment you use to launch the app. Windows PowerShell and PowerShell 7 can use different module paths. |
| The WPF window won't open. | Use Windows and launch PowerShell with `-STA`. |
| Apply or export is disabled. | Connect and complete a new scan. Scope changes and approved live runs clear the previous inventory. |
| Recent tags aren't in the scan. | Allow time for Resource Graph to update, then scan again. Previews read current tags for known targets. |
| A write is denied. | Check the error details, target-scope permissions, resource locks, and Azure Policy. |
| A result is **Conflict** or **Unverified**. | Inspect the resource and scan again. Don't assume a retry is safe. |

If Windows marks downloaded files as blocked, review the source and files first.
Then, from this repository folder, unblock only the project files you trust:

```powershell
Get-ChildItem -Path . -Recurse -File -Include *.ps1, *.psm1, *.xaml | Unblock-File
```

Unblocking doesn't override an execution policy that requires signed scripts.
Follow your organization's [PowerShell execution policy](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_execution_policies)
rather than bypassing it.

## Run the tests

Run the suite on Windows in both Windows PowerShell 5.1 and PowerShell 7. Start
an STA session with `powershell -NoProfile -STA` or `pwsh -NoProfile -STA`, then
change to the repository folder.

If Pester 5.7.1 isn't installed in that environment, install it:

```powershell
Install-Module Pester -RequiredVersion 5.7.1 -Scope CurrentUser
```

Import that version explicitly and run the tests:

```powershell
Import-Module Pester -RequiredVersion 5.7.1 -ErrorAction Stop
Invoke-Pester -Path .\tests -Output Detailed
```

Version 1.3.0 passed all 102 tests on Windows PowerShell 5.1 and PowerShell 7.6.6.
Coverage includes scope checks, filters, tag planning, dry runs, conflicts, write
verification, CSV formula protection, UI events, and background worker failures.

The tests create WPF windows without displaying them and replace Azure calls
with test doubles. They don't sign in to Azure or change resources. They aren't
a live Azure integration test or a comprehensive dependency vulnerability audit.

## File structure

| File | Purpose |
|------|---------|
| [Start-ResourceTagger.ps1](Start-ResourceTagger.ps1) | Starts the app and connects the UI to Azure operations |
| [ResourceTagger.Core.psm1](ResourceTagger.Core.psm1) | Handles scope checks, filters, tag plans, verified writes, and CSV protection |
| [MainWindow.xaml](gui/MainWindow.xaml) | Defines the main window |
| [TagPreview.xaml](gui/TagPreview.xaml) | Defines the shared preview and confirmation dialog |
| [Core tests](tests/ResourceTagger.Core.Tests.ps1) | Tests tagging logic and CSV protection |
| [UI tests](tests/ResourceTagger.UI.Tests.ps1) | Tests UI events and isolated background workers |
| [Changelog](CHANGELOG.md) | Records release changes |

## Author

**Zac Larsen**. This is a personal project, not an official Microsoft product.

## Support and responsible use

Issues and pull requests are welcome. Include the steps to reproduce a problem,
your PowerShell and Azure module versions, and the error text. Remove tenant and
subscription IDs, credentials, internal URLs, and other confidential information
before posting. Keep secrets out of tag values, too.

The app uses Azure sign-in and management APIs to work with resources your
account can access. It displays results locally and can save a CSV report. It
doesn't send inventories to a separate reporting service.

For Azure platform issues or outages, contact [Azure Support](https://azure.microsoft.com/support/).
Use this repository for issues with the tool itself.

## License and disclaimer

This project is licensed under the [MIT license](LICENSE). It contains sample
tooling developed by a Microsoft employee for informational and educational use.

**This isn't an official Microsoft product, service, or supported offering.**

The project is provided **as is**, without express or implied warranties. It
doesn't guarantee production readiness, security hardening, tenant compatibility,
governance alignment, successful tag updates, or policy compliance.

Microsoft support agreements, Premier or Unified Support plans, and Azure support
contracts don't cover this project. No Microsoft service-level agreements,
warranties, or product commitments apply to the project or its derivatives.

Live operations can change Azure tags within the permissions you've granted.
You're responsible for validating the tool and its changes before using it in
production.

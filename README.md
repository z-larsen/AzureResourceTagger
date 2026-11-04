# Azure Resource Tagger

![PowerShell 7.6 LTS and 5.1](https://img.shields.io/badge/PowerShell-7.6%20LTS%20%7C%205.1-blue?logo=powershell&logoColor=white)
![Azure Az Modules](https://img.shields.io/badge/Azure-Az%20Modules-0078D4?logo=microsoftazure&logoColor=white)
![License MIT](https://img.shields.io/badge/License-MIT-green)
![Version 1.3.0](https://img.shields.io/badge/Version-1.3.0-brightgreen)

Find the tagging gaps in your Azure environment, preview exactly what you'd change,
and then apply or remove tags across your resource groups and resources.

Azure Resource Tagger is a Windows desktop tool I built with PowerShell and Windows
Presentation Foundation (WPF). The flow is simple: connect, scan, preview, then
confirm. Dry run is on by default for every tagging workflow, so you see what would
change before anything gets written.

Ready to try it? Head to the [quick start](#quick-start). Want the fine print first?
Check the [limits and gotchas](#limits-and-gotchas) or the [release notes](CHANGELOG.md).

## Why I made this

Azure Policy can enforce tags, and applicable **modify** policies can update
existing resources through [remediation tasks](https://learn.microsoft.com/azure/governance/policy/how-to/remediate-resources).
It's great for ongoing governance once the tag foundation has been set.

But what if you don't have IaC fully implemented yet? What if your environment is large and your tagging spans years of deployments where some resources were tagged, some weren't, or some were tagged using older standards you no longer follow?

I made this tool for that interactive cleanup. It shows you what needs attention
before you make a bulk change, whether you're filling gaps, updating existing
values, or getting a subscription ready for Azure Policy tag enforcement. Once the
foundation's in place, use Azure Policy for ongoing governance.

## What you need

- Windows, with PowerShell 7.6 (LTS) or Windows PowerShell 5.1. WPF doesn't run on macOS or Linux. PowerShell 7.4 and 7.5 reach end of support on November 10, 2026. On Windows PowerShell 5.1, the Az modules need .NET Framework 4.7.2 or later.
- The `Az.Accounts` (5.2.0 or later), `Az.Resources`, and `Az.ResourceGraph` PowerShell modules. Az 16 is the current major version, and Az.Accounts 5.2.0 or later returns clear multifactor authentication (MFA) errors.
- An Azure account with access to the subscription you want to scan.

You don't need local administrator rights. The app runs in your user context.

### Permissions you need

| Task | Azure role |
|------|------------|
| Scan a subscription | **Reader** on the subscription |
| Apply or remove tags | **Tag Contributor** on the target scope, in addition to the read access needed for scanning |

Think of these roles as a starting point. They won't get you around resource locks
or policy. Use the narrowest scope that covers your work.

### Which clouds it works with

| Cloud | Azure environment |
|-------|-------------------|
| Azure Commercial | `AzureCloud` |
| Azure Government | `AzureUSGovernment` |

## Quick start

Download or clone this repo, then open PowerShell in the repo folder. Keep the
folder structure as is so the startup script can find the core module and XAML
files.

If you don't have the Azure modules yet, install them in the PowerShell environment
you'll run the app from:

```powershell
Install-Module Az.Accounts -Scope CurrentUser -Repository PSGallery -MinimumVersion 5.2.0
Install-Module Az.Resources, Az.ResourceGraph -Scope CurrentUser -Repository PSGallery
```

If you use PSResourceGet, `Install-PSResource Az.Accounts, Az.Resources, Az.ResourceGraph -Scope CurrentUser` works too.

Start the app in single-threaded apartment (STA) mode. WPF needs it.

**Windows PowerShell 5.1:**

```powershell
powershell -NoProfile -STA -File .\Start-ResourceTagger.ps1
```

**PowerShell 7:**

```powershell
pwsh -NoProfile -STA -File .\Start-ResourceTagger.ps1
```

If Windows blocks the files you downloaded, see [troubleshooting](#if-something-goes-wrong).

## Scan and review your tags

I'd start with a resource group you know before running anything subscription-wide.

1. Select **Commercial Tenant** or **Gov Tenant**. Sign in if prompted, then choose a tenant.
2. On **Scope & Scan**, select a subscription. To narrow the scan, set the scope to **Resource Group** and choose a resource group.
3. Select **Scan Tags**.
4. Check the recorded cloud, tenant, subscription, account, resource group, and scan time.
5. Review the summary, then use **Resource Groups** or **Resources** to inspect the inventory.

The scan reads from Azure Resource Graph, so you see what comes back for your
selected scope and permissions. That's not a guaranteed inventory of every Azure
resource type.

The coverage percentage means **resources with at least one tag**. It doesn't tell
you those resources have every required tag. The tag-key summary and unique-key
count only cover resource groups.

### Find missing tags

| Tab | What to do |
|-----|------------|
| **Resource Groups** | Enter a comma-separated list of required tag names, then select **Missing Any Required Tag**. |
| **Resources** | Select **Untagged Resources**, or select **Missing Specific Tag** and enter one tag name. |

Required-tag names aren't case-sensitive, and duplicates are ignored. The
required-tag list applies to resource groups only. It isn't a compliance check
across all your resources.

Filters work off the saved scan data, so changing a filter or a required-tag name
doesn't make another Azure request. Select **All Resource Groups** or **All
Resources** to get the full scanned list back.

## Apply tags

Preview first, always. A dry run reads current tags from Azure but doesn't write
any changes, and the preview has no enabled Apply action.

### Preview a bulk update

1. On **Apply Tags**, enter a tag name and value, then select **Add Tag**. Repeat for each tag you want to apply.
2. Choose a target scope from the table below.
3. Leave dry run selected. Select **Overwrite existing tag values** only if you mean to replace what's already there.
4. Select **Apply Tags**. If you chose **Selected Resource Groups**, select the groups and then select **Preview**.
5. Review the scope, before-and-after values, skipped tags, and any read errors.

| Target scope | What it includes |
|--------------|------------------|
| **All Resource Groups in Scope** | All scanned resource groups, subject to your overwrite setting |
| **RGs Missing the First Queued Tag** | Scanned resource groups that don't have the first tag name in the queue |
| **All Resources in Scope** | All scanned resources, subject to your overwrite setting |
| **All Untagged Resources** | Scanned resources that currently have no tags |
| **Selected Resource Groups** | Resource groups you choose from the scanned inventory |

RG means resource group. Heads up: the missing-tag scope checks only the **first
queued tag**. It doesn't check whether a resource group is missing any tag in the
queue.

The app checks eligibility and tag values against fresh Azure Resource Manager
(ARM) tag reads. In a merge, it sends only new or changed tag keys, so your
unrelated tags don't get sent back. Unchanged values are skipped, even with
overwrite on. Empty values and whitespace are preserved.

### Preview an update to selected resources

1. On **Resources**, select the resources you want to tag. Use Ctrl or Shift to select more than one.
2. Leave **Dry run** selected. Select **Overwrite** only if you mean to replace what's already there.
3. Select **Preview Tag**, enter a tag name and value, then select **Preview**.
4. Review the proposed changes. Results appear on **Apply Tags**.

The dry run and overwrite settings on **Resources** are separate from the ones on
**Apply Tags**.

## Remove tags

1. On **Remove Tags**, select or enter a tag name.
2. Choose whether to match an exact value, as described below.
3. Choose all resource groups, all resources, or both within the scanned scope.
4. Leave dry run selected, then select **Remove Tags**.
5. Review the proposed removals before starting a live run.

**Match exact value** is off by default. With it off, the tag you pick is eligible
for removal no matter what its value is. With it on, matching is case-sensitive and
whitespace counts. An empty input matches only an empty tag value.

For example, `Production` and `production` aren't the same value. Tag names are
still matched without regard to case.

**Refresh List** rebuilds the tag-name list from the saved scan, so it doesn't
query Azure again. You can also type a tag name yourself.

## Run it for real

The preview is the confirmation step. It's not a promise that Azure has changed.

1. Close the dry run preview.
2. Clear the dry run checkbox for the operation you want to run.
3. Start the operation again and look over the new preview. Confirm only if the scope and proposed changes are right.
4. Read the results, then scan again before you write or export anything else.

The app ties each scan to its cloud, tenant, subscription, account, and resource
group. If you change the connection or scope, the old scan gets cleared. While an
operation is running, the app disables any controls that could change the scope or
start conflicting work. Wait for the operation to finish before you close the window.

An approved live run clears the saved inventory, even if some changes didn't go
through. The operation results stay visible.

## What the results mean

| Status | What it means | What to do |
|--------|---------------|------------|
| **Planned** | A change is in the preview. Nothing's been written yet. | Review it before you confirm a live run. |
| **Skipped** | Nothing's planned or written for this tag with the options you picked. | Read the detail. Common reasons are an unchanged value, a value mismatch, or overwrite being off. |
| **Success** | The app wrote the change and read it back to verify it. | Review the result, and scan again before you keep going. |
| **Conflict** | An affected tag, or the target's eligibility, changed after the preview. The app didn't attempt a write for that target. | Scan again and review a new preview. |
| **Error** | Planning or validation failed before a write started. | Read the error and fix the cause. |
| **Unverified** | A write failed or timed out, or the app couldn't confirm the result. | Inspect the resource and scan again before you retry. |

**Unverified doesn't mean nothing changed.** The app doesn't retry an uncertain
write for you.

## Export a report

After a successful scan, select **Export to CSV** and pick a location. The report
has the scanned resource groups and resources, along with their locations,
resource types, tag counts, and tag text.

Any text cell that looks like a formula gets a visible `[text] ` prefix, both in the
CSV and when you copy rows from a grid. That includes formula markers hidden behind
whitespace or control characters, and full-width variants. The prefix is a safety
measure for what leaves the app. It doesn't change the inventory or your Azure tags.

Leave the prefix in place when you open or save the report in a spreadsheet. If you
remove it, metadata can turn into an active formula. Treat the CSV as a report, not
something to import back into Azure unchanged.

Look over a report before you share it. Resource names and tag values can contain
sensitive information, and tags are plain text, so don't put secrets or other
sensitive values in them.

## Limits and gotchas

- **Scans can lag behind Azure.** Resource Graph is eventually consistent, so new resources and recent tag changes might not show up right away.
- **Previews refresh known targets only.** Fresh ARM reads update tag info for resources that are already in the scan. They don't find new ones.
- **Conflict checks aren't a lock.** The app checks affected tags before it writes, but another tool can still change them between that read and the update. Azure tag updates aren't atomic compare-and-swap operations.
- **Azure still enforces its rules.** Permissions, resource locks, policy, supported resource types, and tag limits can all block a change. Resources allow up to 50 tags, and some types allow only 15. Tag names can't contain `< > % & \ ? /`, and can be up to 512 characters (128 for storage accounts). Values can be up to 256 characters. The app checks the name and value rules and the 50-tag cap before it plans a write, but it can't know every resource-specific limit, so Azure has the final say.
- **Resource group scope needs a specific group.** If you set the scope to **Resource Group**, pick one. The app won't quietly fall back to the whole subscription.
- **User accounts need MFA for writes.** Azure requires MFA for create, update, and delete operations from user accounts. Reads don't need it, and managed and workload identities are exempt.
- **Reads can pause the UI.** Inventory and planning reads aren't fully asynchronous, and this release has no way to cancel one.
- **Tests aren't production approval.** Try the tool in an approved test subscription before you use it in production.

## If something goes wrong

| Problem | What to check |
|---------|---------------|
| The app reports a missing module. | Install it in the same PowerShell environment you launch the app from. Windows PowerShell and PowerShell 7 can use different module paths. |
| The WPF window won't open. | Make sure you're on Windows and that you started PowerShell with `-STA`. |
| Apply or export is disabled. | Connect and run a new scan. Scope changes and approved live runs clear the previous inventory. |
| Recent tags aren't in the scan. | Give Resource Graph some time to catch up, then scan again. Previews read current tags for known targets. |
| A write is denied. | Check the error details, your permissions on the target scope, resource locks, and Azure Policy. |
| A write fails with "Users must use MFA". | Select **Commercial Tenant** or **Gov Tenant** to sign in again, complete the MFA prompt, then scan again and retry. Update Az.Accounts to 5.2.0 or later so the error is clear. |
| A result is **Conflict** or **Unverified**. | Inspect the resource and scan again. Don't assume a retry is safe. |

If Windows marks the files you downloaded as blocked, review the source and the
files first. Then, from the repo folder, unblock only the project files you trust:

```powershell
Get-Item -LiteralPath .\Start-ResourceTagger.ps1, .\ResourceTagger.Core.psm1, .\gui\MainWindow.xaml, .\gui\TagPreview.xaml | Unblock-File
```

Unblocking won't override an execution policy that requires signed scripts. Follow
your organization's [PowerShell execution policy](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_execution_policies)
instead of working around it.

The scripts and XAML files aren't signed. If you put the tool on shared machines,
sign the scripts with Authenticode and restrict write access to the folder so no
one else can change the code you run.

## Run the tests

Run the suite on Windows in both Windows PowerShell 5.1 and PowerShell 7. Start
an STA session with `powershell -NoProfile -STA` or `pwsh -NoProfile -STA`, then
change to the repo folder.

If you don't have Pester 5.7.1 in that environment, install it:

```powershell
Install-Module Pester -RequiredVersion 5.7.1 -Scope CurrentUser
```

Import that exact version and run the tests:

```powershell
Import-Module Pester -RequiredVersion 5.7.1 -ErrorAction Stop
Invoke-Pester -Path .\tests -Output Detailed
```

The suite has 143 tests, and they all pass on Windows PowerShell 5.1 and PowerShell
7.6.9 with Pester 5.7.1 and 6.0.0. They cover scope checks, filters, tag planning,
dry runs, conflicts, write verification, CSV formula protection, UI events, and
background worker failures.

The tests create WPF windows without showing them and swap the Azure calls for test
doubles. They don't sign in to Azure or change any resources. They also aren't a
live Azure integration test or a full dependency vulnerability audit.

## What's in the repo

| File | Purpose |
|------|---------|
| [Start-ResourceTagger.ps1](Start-ResourceTagger.ps1) | Starts the app and connects the UI to Azure operations |
| [ResourceTagger.Core.psm1](ResourceTagger.Core.psm1) | Handles scope checks, filters, tag plans, verified writes, and CSV protection |
| [MainWindow.xaml](gui/MainWindow.xaml) | Defines the main window |
| [TagPreview.xaml](gui/TagPreview.xaml) | Defines the shared preview and confirmation dialog |
| [Core tests](tests/ResourceTagger.Core.Tests.ps1) | Tests tagging logic and CSV protection |
| [UI tests](tests/ResourceTagger.UI.Tests.ps1) | Tests UI events and isolated background workers |
| [Changelog](CHANGELOG.md) | What changed in each release |

## About me

I'm **Zac Larsen**, and this is a personal project, not an official Microsoft product.

## Help and responsible use

Issues and pull requests are welcome. If you hit a problem, include the steps to
reproduce it, your PowerShell and Azure module versions, and the error text. Before
you post, remove tenant and subscription IDs, credentials, internal URLs, and
anything else confidential. And keep secrets out of your tag values, too.

The app uses Azure sign-in and management APIs to work with the resources your
account can access. Results show up locally, and you can save a CSV report. It
doesn't send inventories to a separate reporting service.

For Azure platform issues or outages, contact [Azure Support](https://azure.microsoft.com/support/).
Use this repo for issues with the tool itself.

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

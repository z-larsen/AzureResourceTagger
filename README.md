# Azure Resource Tagger

![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)
![Azure Az Modules](https://img.shields.io/badge/Azure-Az%20Modules-0078D4?logo=microsoftazure&logoColor=white)
![License MIT](https://img.shields.io/badge/License-MIT-green)
![Version 1.3.0](https://img.shields.io/badge/Version-1.3.0-brightgreen)

A PowerShell WPF application that scans an Azure subscription for existing tags
across resource groups and resources, identifies tagging gaps against a
configurable required-tag list, and lets you bulk-apply tags at scale with a
dry-run-first workflow.

Built for governance and compliance workflows -- especially useful when
preparing a subscription for Azure Policy tag enforcement and backfilling tags
on existing resources before turning on deny policies.

---

## Why This Exists

Azure Policy can enforce tags, and applicable **modify** policies can backfill
existing resources through remediation tasks. Deny policies do not backfill tags:
noncompliant create or update requests can be blocked when enforcement is enabled.

Azure Resource Tagger provides an interactive way to inspect gaps, review exact
changes, and apply desired tags before enforcing policies.

---

## What It Does

| Area | Data Source | What You See |
|------|-----------|--------------|
| **Tag Inventory** | Azure Resource Graph | Every tag name/value on RGs and resources in scope |
| **Gap Analysis** | Resource Graph + required-tag list | Which RGs are missing which required tags |
| **Coverage Metrics** | Resource Graph | Resources with any tag %, untagged RG count, unique RG tag keys |
| **Bulk Tagging** | ARM Tags API (`Update-AzTag -Operation Merge`) | Apply one or more tags to RGs or resources at scale |
| **Selective Tagging** | ARM Tags API + RG picker dialog | Apply tags to hand-picked resource groups |
| **Tag Removal** | ARM Tags API (`Update-AzTag -Operation Delete`) | Remove tags by key with optional exact, case-sensitive value matching |
| **Dry Run** | Current ARM tag reads + shared planner | See additions, changes, removals, skips, and read errors without writes |
| **CSV Export** | Scan results | Full tag inventory for offline analysis, with formula-like text neutralized |

---

## Quick Start

```powershell
# If downloaded from GitHub, unblock the files first:
Get-ChildItem -Path .\AzureResourceTagger -Recurse | Unblock-File

# Set execution policy if needed (current user only):
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

cd AzureResourceTagger
.\Start-ResourceTagger.ps1
```

**Alternative -- run with bypass (no policy change required):**

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-ResourceTagger.ps1
```

> **"Not digitally signed" error?** Windows marks downloaded files as blocked. Run
> `Unblock-File` on the extracted folder, or use the `-ExecutionPolicy Bypass`
> command above, or right-click the `.ps1` file → Properties → check **Unblock**.

1. Click **Commercial Tenant** or **Gov Tenant** to authenticate
2. Select a subscription (and optionally a specific resource group)
3. Click **Scan Tags** to inventory all RGs and resources with their tags
4. Review the **Scope & Scan**, **Resource Groups**, and **Resources** tabs
5. Switch to **Apply Tags** to bulk-tag resources

---

## Prerequisites

- Windows with Windows PowerShell 5.1 or PowerShell 7 (WPF is Windows-only)
- Start PowerShell 7 in STA mode: `pwsh -STA -File .\Start-ResourceTagger.ps1`
- Azure PowerShell modules:

```powershell
Install-Module Az.Accounts, Az.Resources, Az.ResourceGraph -Scope CurrentUser
```

> **No elevated or admin permissions are required on your local machine.** The script
> runs in your normal user context. All it needs is the Azure RBAC roles listed
> under [Required Permissions](#required-permissions) below.

---

## Tabs

### Scope & Scan
- Select subscription and optional RG scope
- The scan identity shows cloud, tenant, subscription, account, RG scope, and timestamp
- Changing the connection or scope clears the inventory and disables writes and export until another successful scan
- Summary cards: RG count, resource count, resources with any tag %, untagged RGs, unique RG tag keys
- Any-tag coverage is **not** required-tag compliance; the tag key summary counts RGs only
- Tag key summary grid showing which tags exist and their coverage across RGs

### Resource Groups
- Full list of resource groups with tag count, missing tags, and all current tags
- Filter to show only RGs missing required tags
- Configurable required-tag list (comma-separated, case-insensitive, duplicate names ignored)
- Filters and required-tag edits rebuild the view locally from the complete scan, without another Azure query

### Resources
- All resources with type, resource group, tag count, and tags
- Filter: all, untagged, or missing a specific tag
- **Preview Tag** button -- select resources (Ctrl+click / Shift+click), enter a tag name and value, and review the shared preview
- Inline **Overwrite** checkbox controls whether existing tag values are replaced on selected resources
- **Dry run** is enabled by default, just like bulk apply and removal
- Results appear on the **Apply Tags** tab

### Apply Tags
- Define tags (name + value) and queue them for application
- Choose target scope:
  - **All Resource Groups in Scope** -- every scanned RG gets the queued tags
  - **RGs Missing the First Queued Tag** -- only RGs that don't have the first queued tag key
  - **All Resources in Scope** -- every scanned resource gets the queued tags
  - **All Untagged Resources** -- only resources with zero tags
  - **Selected Resource Groups** -- opens a multi-select picker dialog where you choose exactly which RGs to tag (supports Ctrl+click, Shift+click, and Select All / Select None)
- **Overwrite** toggle controls whether existing tag values are replaced
- Eligibility and existing values are checked against fresh ARM tag reads, not just the scan snapshot
- Merges send only changed/new keys; unrelated existing tags are not resent
- Unchanged values are skipped, and tag value whitespace and empty strings are preserved
- **Dry Run** mode previews changes without applying (enabled by default)
- The preview is also the explicit confirmation step for live operations
- Results show resource ID, tag, before/after values, action, status, and detail

### Remove Tags
- Tag key dropdown auto-populates from scan data (or type manually)
- **Match exact value** is off by default, meaning any value for the selected key
- When enabled, matching is case-sensitive and preserves whitespace; an empty input matches only an empty tag value
- Choose scope: all RGs, all resources, or both
- **Dry Run** mode previews removals without executing (enabled by default)
- Shared preview and explicit confirmation before live removal
- Results show each tag's previous value and final status

### Safety and result semantics

- Every Azure inventory/tag request uses an explicit context. Planned targets must match the captured subscription and RG boundary.
- Connection, scope, and conflicting operation controls are disabled while an operation is active. Wait for completion before closing the window.
- Previewed tag values are re-read before writing. A changed affected tag produces **Conflict**, with no write to that target.
- **Success** means the requested result was read back and verified. **Skipped** means no change was needed; **Error** means planning or pre-write validation failed.
- **Unverified** means a write failed/timed out or its result could not be confirmed. It is not a success or a guarantee that nothing changed. Inspect and re-scan before retrying.
- After any confirmed live execution, cached scan data is invalidated. Results remain visible, but another scan is required before writing or exporting.
- Dry runs make Azure **read** requests but never write, and their preview has no executable Apply action.
- CSV exports prefix formula-like text cells with `[text] `, including leading whitespace/control characters and full-width formula markers. This visible prefix protects spreadsheet viewing without relying on quote-only escaping. The original inventory and Azure tags are unchanged.

CSV reports are intended for inspection, not lossless tag round-tripping.
Keep the `[text] ` safety prefix when viewing or re-saving reports in spreadsheet
software; removing it can turn untrusted metadata back into an active formula.

Azure tag updates are not atomic compare-and-swap operations: rechecking reduces
concurrency risks but cannot eliminate changes made by another tool between the
read and the write. Resource Graph is eventually consistent, so new resources or
recent tag changes can take time to appear in a scan. The planner refreshes tags
only for targets already in that scan.

Azure RBAC, resource locks, policy, supported resource types, and tag limits still
apply. The app does not automatically retry uncertain writes. Inventory and
planning reads can temporarily pause the UI; a fully asynchronous read/cancellation
framework is outside this release.

---

## Required Permissions

| Action | Minimum Role |
|--------|-------------|
| Scan tags | **Reader** on the subscription |
| Apply/remove tags | **Tag Contributor** on the target scope |

---

## Cloud Support

| Environment | Supported |
|------------|-----------|
| Azure Commercial (`AzureCloud`) | Yes |
| Azure Government (`AzureUSGovernment`) | Yes |

---

## File Structure

```
AzureResourceTagger/
├── Start-ResourceTagger.ps1    # Main script (launch this)
├── ResourceTagger.Core.psm1   # Scope guards, filtering, planning, and verified execution
├── gui/
│   ├── MainWindow.xaml        # Main WPF window
│   └── TagPreview.xaml        # Shared review/confirmation dialog
├── tests/
│   ├── ResourceTagger.Core.Tests.ps1
│   └── ResourceTagger.UI.Tests.ps1
├── CHANGELOG.md
├── LICENSE
└── README.md
```

---

## Tests

Run on Windows, in an STA PowerShell session. Install Pester into the PowerShell
environment you intend to test if it is not already available:

```powershell
Install-Module Pester -RequiredVersion 5.7.1 -Scope CurrentUser
Import-Module Pester -RequiredVersion 5.7.1 -ErrorAction Stop
Invoke-Pester -Path .\tests -Output Detailed -CI
```

Run the suite in both Windows PowerShell 5.1 and PowerShell 7 (`pwsh -STA`).
Tests cover scope isolation, filters, exact value matching, delta payloads,
previews, dry runs, conflicts, verification, CSV formula neutralization, UI event wiring, and background
worker errors/timeouts. WPF windows are instantiated without displaying them.
Azure boundaries are stubbed: these tests do not authenticate or modify Azure
resources and do not replace a live smoke test in an approved test subscription.

---

## Author

**Zac Larsen** — Personal project (not an official Microsoft product)

---

## Support & Responsible Use

This tool queries only public Azure APIs (Resource Graph, ARM Tags) against **your own Azure subscriptions**. It reads resource metadata (such as subscription IDs/names, resource groups, resource types, and tags) and writes results locally (console output and CSV exports); it does **not** transmit this data off your machine except as required to call Azure APIs.

- **Issues & PRs:** Welcome! Please do not include subscription IDs, tenant IDs, internal URLs, or any confidential information.
- **Azure support:** For Azure platform issues or outages, contact [Azure Support](https://azure.microsoft.com/support/) — not this repository.
- **Exported files:** Review CSV exports before sharing externally — they may contain subscription IDs, resource names, and tag values for your environment.

This project may access or process Resource Graph, ARM Tags, Subscription, and Resource Group metadata through Azure APIs.

Execution of this tool may initiate:

- Resource discovery
- Tag inventory and gap analysis
- Tag application (merge) to resource groups and resources
- Tag removal (delete) from resource groups and resources

Ensure that least-privilege access is used when running this utility.

---

## OSS Project Disclaimer

This repository contains sample tooling developed by a Microsoft employee and is provided for informational and educational purposes only.

**This is not an official Microsoft product, service, or supported offering.**

This project is provided "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO:

- Production readiness
- Security hardening
- Tenant compatibility
- Governance alignment
- Tag application outcome guarantees
- Policy compliance assurance

Microsoft does not provide support for this project under any Microsoft support agreement, Premier/Unified Support plan, or Azure support contract.

No Microsoft service level agreements (SLAs), warranties, or product commitments apply to this repository or any derivative use of its contents.

Execution of this tool within an Azure tenant may result in tag modifications to resource groups and resources depending on permissions granted.

Users are solely responsible for validating all scripts and automation prior to execution in production environments.

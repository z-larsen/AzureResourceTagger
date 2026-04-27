# Azure Resource Tagger

![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)
![Azure Az Modules](https://img.shields.io/badge/Azure-Az%20Modules-0078D4?logo=microsoftazure&logoColor=white)
![License MIT](https://img.shields.io/badge/License-MIT-green)
![Version 1.2.0](https://img.shields.io/badge/Version-1.2.0-brightgreen)

A PowerShell WPF application that scans an Azure subscription for existing tags
across resource groups and resources, identifies tagging gaps against a
configurable required-tag list, and lets you bulk-apply tags at scale with a
dry-run-first workflow.

Built for governance and compliance workflows -- especially useful when
preparing a subscription for Azure Policy tag enforcement and backfilling tags
on existing resources before turning on deny policies.

---

## Why This Exists

Azure Policy can enforce tags going forward, but it can't backfill what's
already deployed. When you flip a "require tag" policy from audit to deny,
every resource group that doesn't already have the tag starts failing
deployments.

The Azure Resource Tagger scans what exists, determines what's missing and bulk applies your desired tags.

---

## What It Does

| Area | Data Source | What You See |
|------|-----------|--------------|
| **Tag Inventory** | Azure Resource Graph | Every tag name/value on RGs and resources in scope |
| **Gap Analysis** | Resource Graph + required-tag list | Which RGs are missing which required tags |
| **Coverage Metrics** | Resource Graph | Tag coverage %, untagged RG count, unique tag keys |
| **Bulk Tagging** | ARM Tags API (`Update-AzTag -Operation Merge`) | Apply one or more tags to RGs or resources at scale |
| **Selective Tagging** | ARM Tags API + RG picker dialog | Apply tags to hand-picked resource groups |
| **Tag Removal** | ARM Tags API (`Update-AzTag -Operation Delete`) | Remove tags by key (with optional value filter) at scale |
| **Dry Run** | Local preview | Preview what would be changed before committing |
| **CSV Export** | Scan results | Full tag inventory exported for offline analysis |

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

- PowerShell 5.1 or later (ships with Windows)
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
- Summary cards: RG count, resource count, tag coverage %, untagged RGs, unique tag keys
- Tag key summary grid showing which tags exist and their coverage across RGs

### Resource Groups
- Full list of resource groups with tag count, missing tags, and all current tags
- Filter to show only RGs missing required tags
- Configurable required-tag list (comma-separated)

### Resources
- All resources with type, resource group, tag count, and tags
- Filter: all, untagged, or missing a specific tag
- **Tag Selected** button -- select one or more resources in the grid (Ctrl+click / Shift+click), click the button, enter a tag name and value in the popup, and the tag is applied directly
- Inline **Overwrite** checkbox controls whether existing tag values are replaced on selected resources

### Apply Tags
- Define tags (name + value) and queue them for application
- Choose target scope:
  - **All Resource Groups in Scope** -- every scanned RG gets the queued tags
  - **All Resource Groups Missing This Tag** -- only RGs that don't have the first queued tag key
  - **All Resources in Scope** -- every scanned resource gets the queued tags
  - **All Untagged Resources** -- only resources with zero tags
  - **Selected Resource Groups** -- opens a multi-select picker dialog where you choose exactly which RGs to tag (supports Ctrl+click, Shift+click, and Select All / Select None)
- **Overwrite** toggle controls whether existing tag values are replaced
- **Dry Run** mode previews changes without applying (enabled by default)
- Confirmation dialog before any live operation
- Results grid shows per-resource success/failure detail

### Remove Tags
- Tag key dropdown auto-populates from scan data (or type manually)
- Optional value filter to remove only specific tag values
- Choose scope: all RGs, all resources, or both
- **Dry Run** mode previews removals without executing (enabled by default)
- Confirmation dialog before any live removal
- Results grid shows per-resource removal detail with previous value

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
├── gui/
│   └── MainWindow.xaml         # WPF window definition
├── LICENSE
└── README.md
```

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

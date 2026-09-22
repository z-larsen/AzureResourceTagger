# Changelog

Notable changes to Azure Resource Tagger. For setup, current behavior, and safety
guidance, see the [README](README.md).

## Unreleased

### Changed

- Reworked the README around setup, scanning, previews, live operations, and troubleshooting.
- Applied the Microsoft Writing Style Guide with direct, conversational wording and consistent UI references.
- Clarified result statuses, CSV protections, test coverage, and the limits of automated validation.

## [1.3.0] - 2026-09-21

### Added

- A shared tagging core and before-and-after preview for bulk updates, selected resource groups, individual resources, and tag removal.
- Dry run enabled by default for individual-resource tagging.
- A scan identity display and checks for the selected cloud, tenant, account, subscription, and resource group.
- Case-sensitive, exact-value matching for tag removal, including empty values and whitespace.
- Conflict checks before writes, verification after writes, and an **Unverified** status when the result is uncertain.
- Pester regression tests for tagging logic, WPF events, and isolated background workers.

### Fixed

- Connection and scope changes now clear the saved scan. Users can no longer approve outdated targets by accepting a warning.
- Azure reads and writes now use the captured context, including writes from background workers.
- Merge requests now send only intended changes, not unrelated existing tags.
- Previews now show skipped tags and read errors instead of presenting requested changes as successful writes.
- Filters now restore the full scanned inventory and respond to edits to required-tag and missing-tag names.
- Required-tag parsing now handles empty lists, single names, multiple names, and case-insensitive duplicates.
- Approved live runs now clear scan data before another write or export, while keeping operation results visible.
- The app now blocks overlapping operations and prevents the window from closing while work is active.

### Changed

- Renamed coverage labels to distinguish resources with any tag from required-tag compliance, and to identify resource-group-only key counts.
- Clarified that the missing-tag bulk scope checks the first queued tag.
- Corrected Azure Policy remediation guidance and documented concurrency, scan freshness, and testing limits.

### Security

- Added a visible `[text] ` prefix to formula-like cells in every exported text column. This protects spreadsheet viewing without changing the inventory or Azure tags.
- Added regression tests for formula markers, leading control characters, CSV quoting, and the actual export event.

## [1.2.0] - 2026-04-27

### Added

- A **Selected Resource Groups** scope on **Apply Tags**, with multiple selection and **Select All** and **Select None** controls.
- ARM Tags API (`Get-AzTag`) discovery for the **Remove Tags** list, including tag keys on resource types that Azure Resource Graph doesn't index.

## [1.1.1] - 2026-04-23

### Fixed

- The **Remove Tags** list now populates when a scan completes, without a manual refresh.

## [1.1.0] - 2026-04-23

### Added

- A **Remove Tags** tab for bulk removal:
  - A tag-name list populated from scan data.
  - An optional value filter.
  - Scope selection for resource groups, resources, or both.
  - Dry run enabled by default, with confirmation before live operations.
  - A results grid that shows previous values.
- Tag removal through `Update-AzTag -Operation Delete`, which preserves unrelated tags.

## [1.0.0] - 2026-04-23

### Added

- The initial WPF interface, with an Azure blue theme matching the FinOps Multitool.
- **Commercial** and **Gov** tenant connection buttons.
- Resource group and resource tag scans through Azure Resource Graph.
- Summary cards for resource counts, tag coverage, untagged resource groups, and unique tag keys.
- A tag-key summary with coverage across resource groups.
- A **Resource Groups** tab with a configurable required-tag list.
- A **Resources** tab with filters for untagged resources or a missing tag.
- An **Apply Tags** tab with a tag queue, target scope, overwrite setting, and dry run.
- Bulk updates through `Update-AzTag -Operation Merge`, with confirmation before live operations.
- CSV export of the scanned tag inventory.
- Placeholder text for the required-tag input.
- Windows PowerShell 5.1 compatibility, including UTF-8 BOM encoding, `IDictionary` tag parsing, and array handling.
- The MIT license and personal-project disclaimer.

# Changelog

Here's what's changed in Azure Resource Tagger. For setup, how it behaves today, and
safety guidance, see the [README](README.md).

## Unreleased

### Security

- Copying rows from any grid now neutralizes formula-like text, the same way the CSV export does.
- Tags named `Keys`, `Count`, or other dictionary member names can no longer sneak past the overwrite default, the **Untagged** filter, conflict checks, or write verification.
- The CSV sanitizer now covers non-string values, too.
- KQL literals are escaped instead of stripped. Tenant and subscription names drop control and bidirectional characters.
- Target IDs have to be canonical Azure resource IDs. File paths use `-LiteralPath`, and the XAML here-strings no longer expand variables.

### Fixed

- Swapped the ignored `AZURE_LOGIN_EXPERIENCE_V2` variable for `Update-AzConfig -LoginExperienceV2 Off`.
- Setting the scope to **Resource Group** now requires you to pick a specific group. Before, it quietly scanned the whole subscription.

### Added

- Tag name, value, and 50-tag limit checks before the app plans a write.
- MFA guidance when Azure rejects a write, plus a note if you're on an Az.Accounts version earlier than 5.2.0.
- Retry with backoff when Resource Graph throttles a request.
- 41 new regression tests, for 143 total.

### Changed

- Rewrote the README and this changelog in a more casual, first-person voice, while keeping UI names consistent.
- The README now covers the PowerShell 7.6 LTS floor, Az 16, MFA, tag limits, and signing guidance.
- Reworked the README around setup, scanning, previews, live operations, and troubleshooting.
- Made the result statuses, CSV protections, test coverage, and the limits of automated validation clearer.

## [1.3.0] - 2026-09-21

### Added

- A shared tagging core and a before-and-after preview for bulk updates, selected resource groups, individual resources, and tag removal.
- Dry run is on by default for individual-resource tagging.
- A scan identity display, plus checks for the selected cloud, tenant, account, subscription, and resource group.
- Case-sensitive, exact-value matching for tag removal, including empty values and whitespace.
- Conflict checks before writes, verification after writes, and an **Unverified** status for when the result is uncertain.
- Pester regression tests for tagging logic, WPF events, and isolated background workers.

### Fixed

- Changing the connection or scope now clears the saved scan. You can no longer approve outdated targets just by accepting a warning.
- Azure reads and writes now use the captured context, including writes from background workers.
- Merge requests now send only the changes you intended, not unrelated existing tags.
- Previews now show skipped tags and read errors instead of making requested changes look like successful writes.
- Filters now restore the full scanned inventory and respond when you edit the required-tag and missing-tag names.
- Required-tag parsing now handles empty lists, single names, multiple names, and case-insensitive duplicates.
- Approved live runs now clear the scan data before another write or export, but your operation results stay visible.
- The app now blocks overlapping operations and won't let you close the window while work is running.

### Changed

- Renamed the coverage labels so "resources with any tag" isn't mistaken for required-tag compliance, and so it's clear that key counts cover resource groups only.
- Clarified that the missing-tag bulk scope checks the first queued tag.
- Corrected the Azure Policy remediation guidance, and documented the limits around concurrency, scan freshness, and testing.

### Security

- Added a visible `[text] ` prefix to formula-like cells in every exported text column. It protects you when you open the CSV in a spreadsheet, and it doesn't change the inventory or your Azure tags.
- Added regression tests for formula markers, leading control characters, CSV quoting, and the actual export event.

## [1.2.0] - 2026-04-27

### Added

- A **Selected Resource Groups** scope on **Apply Tags**, with multiple selection plus **Select All** and **Select None** controls.
- The **Remove Tags** list now also finds tags through the ARM Tags API (`Get-AzTag`), including tag keys on resource types that Azure Resource Graph doesn't index.

## [1.1.1] - 2026-04-23

### Fixed

- The **Remove Tags** list now fills in when a scan finishes, so you don't have to refresh it by hand.

## [1.1.0] - 2026-04-23

### Added

- A **Remove Tags** tab for bulk removal:
  - A tag-name list filled in from scan data.
  - An optional value filter.
  - Scope selection for resource groups, resources, or both.
  - Dry run on by default, with a confirmation before live operations.
  - A results grid that shows previous values.
- Tag removal goes through `Update-AzTag -Operation Delete`, so unrelated tags are left alone.

## [1.0.0] - 2026-04-23

### Added

- The first version of the WPF interface, with an Azure blue theme that matches the FinOps Multitool.
- **Commercial** and **Gov** tenant connection buttons.
- Resource group and resource tag scans through Azure Resource Graph.
- Summary cards for resource counts, tag coverage, untagged resource groups, and unique tag keys.
- A tag-key summary with coverage across resource groups.
- A **Resource Groups** tab with a required-tag list you can configure.
- A **Resources** tab with filters for untagged resources or a missing tag.
- An **Apply Tags** tab with a tag queue, target scope, overwrite setting, and dry run.
- Bulk updates through `Update-AzTag -Operation Merge`, with a confirmation before live operations.
- CSV export of the scanned tag inventory.
- Placeholder text for the required-tag input.
- Windows PowerShell 5.1 compatibility, including UTF-8 BOM encoding, `IDictionary` tag parsing, and array handling.
- The MIT license and the personal-project disclaimer.

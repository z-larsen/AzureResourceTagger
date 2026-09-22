# Changelog

All notable changes to the Azure Resource Tagger will be documented in this file.

## [1.3.0] - 2026-09-21

### Added
- Shared, testable tagging core and before/after preview for bulk apply, selected RGs, individual resources, and removal
- Default-on dry run for individual-resource tagging
- Scan identity display and explicit cloud, tenant, account, subscription, and RG scope validation
- Case-sensitive exact-value removal option, including empty values and whitespace
- Pre-write conflict checks, post-write verification, and explicit unverified outcomes
- Pester regression tests covering the core, WPF event wiring, and isolated background write workers

### Fixed
- Connection/scope changes invalidate cached scans; stale targets cannot be approved through a warning
- All Azure reads/writes use captured contexts, including background workers
- Merge payloads contain only intended changes rather than resending unrelated tags
- Previews reflect actual skips and read errors instead of reporting requested tags as successful changes
- Filters restore the full inventory locally and react to required-tag and missing-tag name edits
- Required-tag parsing handles zero/one/multiple names and case-insensitive duplicates
- Live operations invalidate scan data before subsequent writes/exports, while preserving operation results
- Shared busy-state gating prevents overlapping operations and closing the window during active work

### Changed
- Renamed coverage labels to clarify any-tag coverage and RG-only key counts
- Clarified that the missing-tag bulk scope uses the first queued tag
- Corrected Azure Policy remediation guidance and documented concurrency, consistency, and testing limits

### Security
- Neutralized spreadsheet-formula injection in every exported text column using a visible `[text] ` prefix, without changing inventory or Azure tag data
- Added regression coverage for formula markers, leading control characters, CSV quoting, and the actual export event

## [1.2.0] - 2026-04-27

### Added
- **Selected Resource Groups** scope in Apply Tags tab -- opens a multi-select picker dialog so you can choose exactly which RGs receive tags (Ctrl+click, Shift+click, Select All / Select None)
- ARM Tags API (`Get-AzTag`) discovery for Remove Tags dropdown -- catches tag keys on resource types that Azure Resource Graph doesn't index, matching what the Azure Portal Tags blade shows

## [1.1.1] - 2026-04-23

### Fixed
- Remove Tags dropdown now auto-populates when scan completes (no longer requires manual Refresh click)

## [1.1.0] - 2026-04-23

### Added
- **Remove Tags** tab for bulk tag removal
  - Dropdown populated from scan data with all discovered tag keys
  - Optional value filter to remove only specific tag values
  - Scope selection: all RGs, all resources, or both
  - Dry run mode (on by default) with confirmation dialog for live operations
  - Results grid with previous value column
- Tag removal uses `Update-AzTag -Operation Delete` (surgical key removal, preserves other tags)

## [1.0.0] - 2026-04-23

### Added
- Initial release
- WPF GUI with Azure blue theme matching the FinOps Multitool
- **Commercial** and **Gov** tenant connection buttons
- Tag inventory scan via Azure Resource Graph (resource groups + resources)
- Summary dashboard cards (RG count, resource count, tag coverage %, untagged RGs, unique tag keys)
- Tag key summary grid with per-key coverage across resource groups
- Resource Groups tab with missing-tag analysis against configurable required-tag list
- Resources tab with filter by untagged or missing specific tag
- Apply Tags tab with tag queue, target scope selection, overwrite toggle, and dry-run mode
- Bulk tag application using `Update-AzTag -Operation Merge`
- Confirmation dialog before live operations
- CSV export of full tag inventory
- Placeholder text on required tags input field
- PowerShell 5.1 compatibility (UTF-8 BOM, `IDictionary` tag parsing, `@()` array wrapping)
- MIT license and OSS disclaimer

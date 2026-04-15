=== KO – GF Daily Approved Export ===
Contributors: ko
Requires at least: 5.8
Requires PHP: 7.4
Stable tag: 1.0.8
License: GPLv2 or later

== Description ==
Custom daily XLSX exporter for Gravity Forms warranty approvals.

This plugin was created to replace the fragile "Approved + Since Last Run" export flow that depended on a large third-party plugin. Instead of using a time-window model, this plugin uses a state-based model:

Approved + Not Yet Exported = Include

If an entry is missed one day for any reason, it remains eligible until this plugin actually exports it.

== Features ==
- Daily scheduled export (default 1:30 AM site time)
- XLSX output matching the current daily report structure
- Emails export to configured recipients
- Tracks exported entries using Gravity Forms entry meta
- Manual "Run Export Now" button
- "Reset Export Flags" button to make entries eligible again
- Admin page with last run status, file name, count, and next run time
- Form is selectable in settings for future flexibility

== Approval Logic ==
An entry is exportable if ANY of the following are true:
- workflow_final_status = complete
- workflow_final_status = approved
- workflow_step_status_4 = approved

AND the entry has not already been exported by this plugin.

== Stored Data ==
This plugin stores:
- Plugin settings and last run status in WordPress options
- Per-entry export flags in Gravity Forms entry meta:
  - ko_gfde_exported_at
  - ko_gfde_exported_batch
  - ko_gfde_exported_file

== File Name ==
Default:
Daily-Warranty-Approval-Report-{date}.xlsx

{date} is replaced with the local run date (YYYY-MM-DD).

== Recommended Rollout ==
- Leave Entry Automation in place initially
- Send this custom exporter to an internal email address for the first 1-2 days
- Compare the datasets
- Once validated, disable the old export and switch recipients to the client list

== Notes ==
- This plugin does not try to coordinate with Entry Automation
- During testing, both systems may export the same entries
- Within this plugin's own runs, entries are exported only once unless flags are reset

== Changelog ==
= 1.0.1 =
* Added customizable email message template.
* Refreshed the admin screen with a more modern layout.
* Fixed Customer Address (Full) export by building the compound address from address sub-fields.


= 1.0.2 =
* Added per-entry export flag management on the Gravity Forms entry detail screen.
* Added an Exported column to the Gravity Forms entries list.
* Clearing a single entry flag only affects this plugin and makes that entry eligible for the next scheduled export only.
* Updated readme for single-entry flag clearing behavior.

= 1.0.4 =
* Added admin tool to clear export flags for specific entry IDs.
* Accepts a single ID or comma-separated list.
* Entries cleared are included in the next scheduled export only.
* Previous single-entry clear on the Gravity Forms entry screen remains available.


= 1.0.5 =
* Fixed fatal error when clearing individual export flags from the admin screen.
* Added missing helper method used by the single-entry and comma-separated flag clear actions.


= 1.0.6 =
* Moved Reset All Export Flags out of the Run Status actions area.
* Renamed the individual flag section to Reset Entry Flags.
* Placed Reset All Export Flags beneath the Reset Entry Flags controls.


= 1.0.7 =
* Moved Reset Entry Flags directly under the Run Status card.
* Added Reset All Export Flags beneath the individual entry reset control.
* Updated labels and layout to better separate run actions from reset actions.


= 1.0.8 =
* Removed Gravity Forms Entries screen hooks that could cause admin-side fatal errors.
* Kept all export, schedule, and plugin admin reset tools intact.
* Reset Entry Flags and Reset All Export Flags remain available only on the plugin admin screen.

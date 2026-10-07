# Changelog — Meeting Cleanup On-Prem

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the developer guide, appendix D).
Author: Nicolas Fabert.

## [1.0.0] — 2026-10-07

First public release: the Exchange Server counterpart of Meeting Cleanup 1.2.3 (Exchange Online), validated on a
lab Exchange Server 2019 (report, Remove, Restore, Cancel, Transfer, rooms over a period, report replays, paging).

### Added
- **Every room mailbox** for `-SearchIn Rooms` (the default): listed by Exchange Management Shell
  (`Get-Mailbox -RecipientTypeDetails RoomMailbox`), with the rooms of `Search.Rooms` and `Search.RoomFile` (before:
  only these); `Search.AllRooms = $false` keeps the rooms of the configuration only (an impersonation limited to some rooms). A warning says when there is no room to search.
- **An organizer not in the directory** (deleted mailbox: `Get-Recipient` does not find it) is shown as such, and its calendar is not opened (a warning before); its meetings are found in the rooms and the other calendars, from the address typed.
- **Retries** (`Connection.MaxRetries`, 3 by default, read but not used before): a request that Exchange did not
  process — busy (`ErrorServerBusy`, EWS throttling, after the delay it asks for) or unavailable (HTTP 503) — is sent
  again; after a network error, only a read (GetFolder, FindItem, GetItem) is sent again, never a change.
- **User guide** and **developer guide** in English (`docs\MeetingCleanupOnPrem-UserGuide.md`,
  `docs\MeetingCleanupOnPrem-Guide.md`, and their HTML build), in place of the short French notes.
- `tools\Build-Documentation.ps1`, `tools\New-ReadmeImages.ps1`, `tools\New-DocumentationImages.ps1` (images of the
  README and the guides, from the simulated Exchange of the tests) and `tools\New-MeetingCleanupOnPremPackage.ps1`
  (the release zip).

### Removed
- `Restore.ReAccept`: it was never used. Exchange Server restores a copy with the answer it had (an accepted room
  stays accepted): there is nothing to answer again.

## [0.4.1] — 2026-10-07

The Transfers tab of Meeting Cleanup 1.2.3, and the calendars longer than a page.

### Added
- **Transfers tab** in the HTML report of a *Transfer* run, open first (as in Meeting Cleanup 1.2.3): one row per
  meeting of the transfer: old organizer and the state of his mailbox, new organizer, method (*Re-created*), status,
  the new meeting (with the attendees and rooms invited), what became of the old meeting at the old organizer, the
  old copies removed, failed or left, and the notes. Filters by status and method; a row opens the meeting and its
  copies. The same rows in **`MeetingCleanupOnPrem-Transfers.csv`**. The copy of the new organizer has the result
  *Created* (it was *Transferred*), as in Meeting Cleanup.

### Fixed
- **A calendar with more items in the period than `Connection.PageSize` (500) was read only in part**, without a
  word: CalendarView gave the first 500 items and the meetings after them were not found (a room with a daily
  series fills 500 items in two years). The calendar is now read page after page: CalendarView has no offset, so
  each page starts at the start of the last item received, the items seen twice are kept once. When more items
  start at the same time than a page holds, the next page jumps past them and the run ends with a warning (*more
  than 500 items at ..., raise Connection.PageSize*). Measured on the simulated Exchange: an organizer calendar of
  608 items read in 2 pages (603 meetings).
- Restore through EWS (`Restore.Mode = 'Ews'`): Recoverable Items read page after page (1,000 items per page)
  instead of the first 1,000 only.

## [0.4.0] — 2026-10-07

The console, the progress and the performance of Meeting Cleanup 1.2.1 and 1.2.2, for Exchange Server. Same
actions, same files, same configuration.

### Added
- **Console of Meeting Cleanup 1.2.2**: title card, numbered steps, coloured result lines with icons, aligned table
  of the meetings, final card with what to do next (`-FromReport ...`). Colours off when the output is redirected
  or `NO_COLOR` is set (`MCO_FORCE_COLOR=1` forces them); `MCO_ICONS = Emoji | Symbols | Ascii`.
- **Live progress line** in an interactive console, rewritten in place: bar, part done, count and **time left**
  (`1,077/1,858 mailboxes searched · about 40 s left`), told from the speed of the progress once it has run 2 s and
  2 %, rounded as a person would say it. Shown for the mailboxes searched, the attendee mailboxes, the meetings
  cancelled, the copies removed, verified and restored, the meetings re-created.
- A **Report** step lists the files written; the restore shows the meetings of the run before the confirmation.
- `tools\Measure-MeetingCleanupOnPrem.ps1`: measures a whole search on a simulated Exchange Server
  (`tests\FakeEws.ps1`, no server needed), with an optional latency per call, and writes the result to compare two
  versions on the same data.

### Changed
- **Each mailbox is read once per search**: an attendee invited to many meetings was read again for each of them
  (one CalendarView per meeting and attendee); now one CalendarView per mailbox for all its meetings, kept for the
  whole search.
- **GetItem reads 50 items per call** (series masters and single meetings mixed) instead of one; a batch that EWS
  refuses is read again item by item.
- The members of a distribution group are read once per run.
- **Compiled engine** (`src\MeetingCleanupOnPrem.Native.cs`, built by PowerShell, no dependency): the
  calendar items of the EWS answers, the recurrence in words, the dates, the totals, the rows, CSV and JSON of the
  report. The loops of the search, the plan, the backup, the removal, the restore and the transfer no longer use the
  pipeline. The compilation can take tens of seconds on a busy Exchange server (45 s on a lab Exchange 2019 server): it is done once per
  account, PowerShell version and tool version, and kept in `%LOCALAPPDATA%\MeetingCleanupOnPrem` (the module then
  loads in 0.3 s).
- Measured on the simulated Exchange (same computer, 0.3.0 then 0.4.0, identical results):

  | Organization | 0.3.0 | 0.4.0 |
  |---|---|---|
  | 600 meetings, 4,200 copies: EWS calls of the search | 8,424 | 162 |
  | 600 meetings, 4,200 copies: search (no latency) | 405 s | 11 s |
  | 600 meetings, 4,200 copies: report (CSV, JSON, HTML) | 7.1 s | 1.4 s |
  | 200 meetings, 1,610 copies, 20 ms per EWS answer: search | 152 s | 8.5 s |

- Exit code 1 when the run failed (2 stays for a run finished with warnings).
- Log lines in the format of Meeting Cleanup (`2026-10-07T09:12:03.120+02:00 [OK   ] ...`).

### Fixed
- The EWS line of step 1 showed no separator between the URL, the version and the access mode.
- **A report can be replayed**: `-FromReport <folder> -Action Remove | Cancel | Transfer` (with `-MeetingId` to act
  on some meetings only) was refused without `-Organizer`, the backup stopped on the item missing from the report,
  and a transfer found no item to re-create the meeting from (it is now read again through EWS).

## [0.3.0] — 2026-10-06

- Made RPS the delivered configuration default for the Exchange Management Shell adapter.

## [0.2.0] — 2026-10-06

- Added optional Exchange Management Shell directory integration.
- Organizer aliases and X500 addresses are resolved with `Get-Recipient`.
- `AllMailboxes` is enumerated with `Get-Mailbox` when available.
- Nested distribution groups are expanded with `Get-DistributionGroupMember`.
- Added offline adapter coverage.
- Added optional Exchange Server Remote PowerShell (RPS) connection using the `Microsoft.Exchange` endpoint and Kerberos by default.

## [0.1.0] — 2026-10-06

- Initial Exchange Server On-Premises implementation.
- EWS SOAP transport with Windows or Basic authentication, delegate or impersonation access, and Autodiscover/manual endpoint selection.
- Report, silent remove, organizer cancel, EWS Recoverable Items restore, and recreate transfer actions.
- Offline Pester coverage for configuration, request building, EWS XML generation/parsing, and reports.

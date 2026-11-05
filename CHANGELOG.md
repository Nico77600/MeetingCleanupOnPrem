# Changelog — Meeting Cleanup On-Prem

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the developer guide, appendix D).
Author: Nicolas Fabert.

## [1.1.1] — 2026-11-05

### Added
- The release ZIP has one folder, `MeetingCleanupOnPrem-1.1.1`, with the PowerShell files, the HTML guides, `README.md`, `CHANGELOG.md`, `LICENSE` and `THIRD-PARTY-NOTICES.md`.
- The PowerShell files carry the owner's timestamped Authenticode signature (Nicolas Fabert).
- The user guide has a filled configuration example, the command that shows your EWS URL, and a "You should see / If not" check after each setup step.

### Changed
- Get started starts from the release ZIP and says which PowerShell to open (`pwsh`, version 7.4 or later; Windows PowerShell 5.1 is not enough).
- The README in the ZIP links only files that the ZIP contains; EWS, impersonation, RPS and Kerberos / NTLM are explained in plain words.
- Before it starts, the tool checks the signature and SHA256 of every file of the package and stops if one is changed, missing or added. `-DevelopmentMode` starts a source tree that has no `PackageIntegrity.psd1`.
- Keep your own files (credential file, lists of addresses) outside the tool folder: only `config\MeetingCleanupOnPrem.config.psd1`, `reports\` and `logs\` may change in it.
- The start compiles the helper at each run (about 10 seconds, up to a minute on a busy Exchange server); `%LOCALAPPDATA%\MeetingCleanupOnPrem` is unused and can be deleted.
- The runtime README and the script help carry the same disclaimer as the project page.

## [1.1.0] — 2026-10-07

One occurrence of a series, without going through the rooms (Meeting Cleanup 1.3.0).

### Added
- **One occurrence of a series**: `-SeriesScope Occurrences` limits a series to its occurrences in the period; with a period of one day, that is one occurrence. *Cancel* sends one cancellation for that date only; *Remove* takes the occurrence out of the attendees' and rooms' calendars without a message. The series goes on.
- A series whose organizer cannot be read is left as it is; a note says when every occurrence of a series is in the period (`-SeriesScope Whole` acts on it at once).
- The report shows how many occurrences were left out, and a replay (`-FromReport`) leaves out the same ones.

### Changed
- A cancellation that fails holds the copies of that occurrence only, not those of the whole meeting.

## [1.0.0] — 2026-10-07

First public release: the Exchange Server counterpart of Meeting Cleanup 1.2.3 (Exchange Online).

### Added
- **Every room mailbox** is searched by default for `-SearchIn Rooms`, with the rooms you list; `Search.AllRooms = $false` keeps the rooms of the configuration only. A warning says when there is no room to search.
- **An organizer who is not in the directory** (deleted mailbox) is shown as such; its meetings are found in the rooms and the other calendars, from the address typed.
- **Retries**: a request that Exchange did not process (busy or unavailable) is sent again, up to `Connection.MaxRetries` times (3 by default); a change is never sent again after a network error.
- **User guide** and **developer guide** in English, as Markdown and as HTML files.

### Removed
- `Restore.ReAccept`: Exchange Server restores a copy with the answer it had, so there is nothing to answer again.

## [0.4.1] — 2026-10-07

### Added
- **Transfers tab** in the HTML report of a *Transfer* run, with `MeetingCleanupOnPrem-Transfers.csv`: for each meeting, from whom to whom, the new meeting and what became of the old one.

### Fixed
- A calendar with more items in the period than `Connection.PageSize` is read in full, page after page; a warning says when a page cannot be completed.
- Restore through EWS reads all the Recoverable Items, not only the first 1,000.

## [0.4.0] — 2026-10-07

### Added
- Console of Meeting Cleanup 1.2.2: title card, numbered steps, coloured results, a table of the meetings and the next command to run.
- Log lines in the format of Meeting Cleanup (`2026-10-07T09:12:03.120+02:00 [OK   ] ...`).
- A live progress line with the time left, and a *Report* step that lists the files written.

### Changed
- Searches are much faster on large organizations: each mailbox is read once per search and the meetings are read in batches.
- Exit code 1 when the run failed (2 stays for a run finished with warnings).

### Fixed
- A report can be replayed with `-FromReport <folder> -Action Remove | Cancel | Transfer`, with or without `-Organizer`.

## [0.3.0] — 2026-10-06

- Remote PowerShell (RPS) is the delivered default for the Exchange Management Shell connection.

## [0.2.0] — 2026-10-06

- Optional Exchange PowerShell: aliases and X500 addresses, every mailbox, nested groups and remote PowerShell (Kerberos by default).

## [0.1.0] — 2026-10-06

- First versions: EWS access, report, Remove, Cancel, Transfer and Restore.
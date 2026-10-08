# Meeting Cleanup On-Prem

Meeting Cleanup On-Prem finds the meetings of organizers who left or stay, or every meeting of some rooms, in every calendar of Exchange Server where they are, then removes them silently, has the organizer cancel them, or transfers them to a new organizer, even when the old mailbox is gone.

This folder contains everything needed to run the tool: Invoke-MeetingCleanupOnPrem.ps1, the module, the configuration, the report template and the guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements
- Exchange Server 2016, 2019 or Subscription Edition, with EWS reachable over HTTPS.
- PowerShell 7.4 or later.
- Windows 10 / 11 or Windows Server 2016 to 2025.
- Service account with ApplicationImpersonation, limited by a management scope when possible.
- Exchange PowerShell is recommended for aliases, X500 addresses, rooms, mailboxes and groups.
- Mailbox Import Export is needed for Restore.

## Quick start
```powershell
notepad .\config\MeetingCleanupOnPrem.config.psd1     # EWS URL, service account, Exchange PowerShell server, accepted domains

# Always a report first (nothing is changed), then the same command with the action
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com -Action Cancel -Comment 'Megan has left the company.'
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer john.doe@contoso.com -Action Transfer -NewOrganizer jane.roe@contoso.com
.\Invoke-MeetingCleanupOnPrem.ps1 -Room room-paris-01@contoso.com -Start 2026-11-02 -End 2026-11-13 -Action Cancel -Comment 'Closed for works.'
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com -Subject 'Weekly sales review' -SeriesScope Occurrences -Start 2026-11-16 -End 2026-11-16 -Action Cancel -Comment 'No sales review this Monday.'
.\Invoke-MeetingCleanupOnPrem.ps1 -Action Restore -FromReport .\reports\MeetingCleanupOnPrem_Remove_20261105-093000
```

## Content
| Item | Role |
|---|---|
| Invoke-MeetingCleanupOnPrem.ps1 | Entry script for the command line. |
| MeetingCleanupOnPrem.psd1 | Module manifest. |
| MeetingCleanupOnPrem.psm1 | Module loader. |
| config | Delivered configuration template. |
| docs | User and developer guides in Markdown and HTML, with images. |
| src | PowerShell source files and native helper source. |
| templates | HTML report template. |
| LICENSE | MIT license. |
| THIRD-PARTY-NOTICES.md | Third-party notices. |

## Documentation
- [User guide](docs/MeetingCleanupOnPrem-UserGuide.md) - also `docs/MeetingCleanupOnPrem-UserGuide.html`, a single file to open locally
- [Developer guide](docs/MeetingCleanupOnPrem-Guide.md) - also `docs/MeetingCleanupOnPrem-Guide.html`

Project page, releases and change log: https://github.com/Nico77600/MeetingCleanupOnPrem

License: [MIT](LICENSE).

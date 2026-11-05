# Meeting Cleanup On-Prem

Meeting Cleanup On-Prem finds the meetings of organizers who left or stay, or every meeting of some rooms, in every calendar of Exchange Server where they are, then removes them silently, has the organizer cancel them, or transfers them to a new organizer, even when the old mailbox is gone.

This folder is the release ZIP, extracted. It holds everything needed to run the tool: `Invoke-MeetingCleanupOnPrem.ps1`, the module, the configuration, the report template and the two guides (HTML files, open them in your browser).

## Requirements
- Exchange Server 2016, 2019 or Subscription Edition, with EWS reachable over HTTPS.
- PowerShell 7.4 or later. Windows PowerShell 5.1 (the blue window) is not enough.
- Windows 10 / 11 or Windows Server 2016 to 2025.
- A service account with the ApplicationImpersonation role, limited by a management scope when possible.
- Exchange PowerShell is recommended: it resolves aliases and X500 addresses and lists rooms, mailboxes and groups.
- The Mailbox Import Export role is needed for Restore only.

Four terms you will meet:
- **EWS** (Exchange Web Services): the web address of your Exchange server that the tool uses to read and change calendars.
- **Impersonation** (ApplicationImpersonation): an Exchange role that lets the service account open other people's calendars as if it were them.
- **RPS** (remote PowerShell): the Exchange command line, opened by the tool from your computer; it looks up addresses, rooms and groups.
- **Kerberos / NTLM**: two ways Windows proves who you are to a server. Leave the default (`Negotiate`, which picks the right one); change it to `NTLM` only if the tool reports an HTTP 401 error.

## Get started
The user guide has every step with what you should see and what to do if not: open `docs\MeetingCleanupOnPrem-UserGuide.html`. In short:

1. Open PowerShell 7: type `pwsh` in the Start menu, or in a console. Check the version:

   ```powershell
   $PSVersionTable.PSVersion     # Major 7, Minor 4 or higher
   ```

2. Unblock the files once, from the folder where you extracted the ZIP (for example `C:\Tools\MeetingCleanupOnPrem-1.1.1`):

   ```powershell
   cd C:\Tools\MeetingCleanupOnPrem-1.1.1
   Get-ChildItem . -Recurse -File | Unblock-File
   ```

3. Fill in the configuration (EWS URL, service account, Exchange PowerShell server, accepted domains):

   ```powershell
   notepad .\config\MeetingCleanupOnPrem.config.psd1
   ```

   Keep credential files and lists of addresses outside the tool folder, for example in `C:\Tools\MeetingCleanupOnPrem-data`: only the configuration, `reports\` and `logs\` may change in it.

4. Run a report first. It changes nothing:

   ```powershell
   .\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com
   ```

5. Then run the same command with the action:

   ```powershell
   .\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com -Action Cancel -Comment 'Megan has left the company.'
   .\Invoke-MeetingCleanupOnPrem.ps1 -Organizer john.doe@contoso.com -Action Transfer -NewOrganizer jane.roe@contoso.com
   .\Invoke-MeetingCleanupOnPrem.ps1 -Action Restore -FromReport .\reports\MeetingCleanupOnPrem_Remove_20261105-093000
   ```

## Content
| Item | Role |
|---|---|
| Invoke-MeetingCleanupOnPrem.ps1 | Entry script for the command line. |
| MeetingCleanupOnPrem.psd1 | Module manifest. |
| MeetingCleanupOnPrem.psm1 | Module loader. |
| PackageIntegrity.psd1 | Signed list of the files of this folder. Do not edit. |
| README.md | This file. |
| config | Delivered configuration template. |
| docs | The user guide and the developer guide (HTML). |
| src | PowerShell source files and native helper source. |
| templates | HTML report template. |
| CHANGELOG.md | What changed in each version. |
| LICENSE | MIT license. |
| THIRD-PARTY-NOTICES.md | Third-party notices. |

## Documentation
- [User guide](docs/MeetingCleanupOnPrem-UserGuide.html): prerequisites, one-time setup and everyday commands.
- [Developer guide](docs/MeetingCleanupOnPrem-Guide.html): how the tool works, rights, settings, report and troubleshooting.

Project page and releases: https://github.com/Nico77600/MeetingCleanupOnPrem

License: [MIT](LICENSE). Changes: [CHANGELOG](CHANGELOG.md). Notices: [THIRD-PARTY-NOTICES](THIRD-PARTY-NOTICES.md).

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

Removing, cancelling or transferring meetings changes real calendars: always start with a report, and test it in your environment before production use.
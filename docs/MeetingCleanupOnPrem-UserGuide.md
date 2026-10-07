---
title: Meeting Cleanup On-Prem
subtitle: User guide
version: 1.0.0
author: Nicolas Fabert
updated: 2026-10-07
---

# Meeting Cleanup On-Prem — User guide

> What you need before the first run, then one command per everyday question: **which meetings does this person still organize?**, **a person has left: cancel their meetings, or give them to someone else?**, **one series to remove without a message**, **rooms closed for works**, **undo a removal**. How the tool works, the rights in detail, the configuration, the report and the internals are in the [developer guide](MeetingCleanupOnPrem-Guide.md).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.

```cards
checklist | Prerequisites | Chapter 1: PowerShell, the service account and its rights, then the one-time setup.
terminal | Everyday use | Chapter 2: one command per question, in the console or in a scheduled task.
filter | Choose the meetings | Chapter 3: who, where to search, which period, which meetings, which server.
file | Results | Chapter 4: where the report is written, exit codes, the usual messages.
```

<!-- icon: checklist -->
## 1. Prerequisites

| Item | Requirement |
|---|---|
| Exchange | **Exchange Server 2016, 2019 or Subscription Edition**, with EWS reachable over HTTPS. Validated on Exchange Server 2019. For Exchange Online, use [Meeting Cleanup](https://github.com/Nico77600/MeetingCleanup). |
| PowerShell | **7.4 or later** (`pwsh`) — a portable zip is enough. No module to install. |
| Windows | Windows 10 / 11, Windows Server 2016 to 2025: an administration workstation or an Exchange server. The tool runs in a console, or in a scheduled task. |
| Service account | An account with the role **ApplicationImpersonation** (it opens the calendars through EWS), preferably limited to the mailboxes concerned by a management scope ([developer guide, chapter 5](MeetingCleanupOnPrem-Guide.md#5-rights-and-connection)). `Delegate` (full access to each mailbox) works as well. |
| Exchange PowerShell | Recommended: the tool connects to it alone (remote PowerShell, Kerberos). It resolves aliases and X500 addresses, lists every room and every mailbox, expands the groups and runs *Restore*. Rights: [Rights for the directory](MeetingCleanupOnPrem-Guide.md#rights-for-the-directory), [Rights for Restore](MeetingCleanupOnPrem-Guide.md#rights-for-restore). |
| Network | HTTPS to the EWS URL (`https://mail.contoso.com/EWS/Exchange.asmx`); HTTP to `/PowerShell/` of a Mailbox server for remote PowerShell. |

> [!WARNING]
> **ApplicationImpersonation** opens the calendar of every mailbox in its scope. Keep the account like an administrator account, limit it to the mailboxes concerned with a management scope, and store its password only as an encrypted credential file (scheduled task, 1.1).

### 1.1 One-time setup

```steps
Copy the tool | Unblock the files, then copy the folder, for example to `C:\Tools\MeetingCleanupOnPrem`. No installer.
Service account | A user account, for example `svc-meetingcleanup`. Give it ApplicationImpersonation and remote PowerShell (commands below).
Exchange roles | The roles of the directory (read only) and, for *Restore*, Mailbox Import Export: [developer guide, chapter 5](MeetingCleanupOnPrem-Guide.md#5-rights-and-connection).
Configure | `notepad .\config\MeetingCleanupOnPrem.config.psd1`: `Connection.EwsUrl`, `Connection.Mailbox` (the address of the service account), `ManagementShell.ServerFqdn`, `Search.AcceptedDomains`.
Check | `.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer <your address>`: step 1 shows Exchange PowerShell and EWS; a report changes nothing.
Scheduled task | Optional: a credential file of the service account (`Connection.CredentialFile`), written by the account that runs the task.
```

```powershell
# Exchange Management Shell, as an Exchange administrator - once
New-ManagementScope -Name 'Meeting Cleanup scope' -RecipientRestrictionFilter "RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'SharedMailbox'"
New-ManagementRoleAssignment -Name 'Meeting Cleanup impersonation' -Role ApplicationImpersonation -User svc-meetingcleanup -CustomRecipientWriteScope 'Meeting Cleanup scope'
Set-User svc-meetingcleanup -RemotePowerShellEnabled $true
Add-RoleGroupMember 'View-Only Organization Management' -Member svc-meetingcleanup
New-RoleGroup -Name 'Meeting Cleanup Restore' -Roles 'Mailbox Import Export' -Members svc-meetingcleanup   # for Restore only
```

```powershell
# Scheduled task: as the account that RUNS the task (the file can only be read by it, on this computer)
Get-Credential CONTOSO\svc-meetingcleanup | Export-Clixml C:\Tools\MeetingCleanupOnPrem\svc.cred.xml
# then, in the configuration: Connection.CredentialFile = 'C:\Tools\MeetingCleanupOnPrem\svc.cred.xml'
```

<!-- icon: terminal -->
## 2. Everyday use

Run the commands from the tool folder, in PowerShell 7. A command without `-Action` is a **report**: it finds the meetings and every copy of them (organizer, attendees, rooms, members of the groups invited) and **changes nothing**. Read the report, then run the same command with an action: *Remove*, *Cancel*, *Transfer* and *Restore* show exactly what will happen and ask to type **YES**; a backup is written before any change.

> [!IMPORTANT]
> Always start with a report. A silent removal can be undone for 14 days (2.8); a **cancellation cannot be undone**: the attendees received it.

### 2.1 Which meetings does this person still organize?

```powershell
# The coming year (default), in the organizer's calendar and every room
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com

# A given period
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com -Start 2026-11-01 -End 2026-12-31
```

Open the HTML report: the **Meetings** tab lists each meeting (a series once), click one for its copies — who has it, which rooms, how it was found.

### 2.2 A person has left, the mailbox is kept: cancel the meetings

```powershell
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com -Action Cancel -Comment 'Megan Bowen has left: this meeting is cancelled.'
```

Each meeting is cancelled by its organizer, with your message: the attendees receive the cancellation, the rooms release the slot, the copies left are removed.

### 2.3 A person has left, the mailbox is deleted

```powershell
# Every room first, then every mailbox for the meetings without room
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer john.doe@contoso.com -SearchIn Rooms, AllMailboxes -Start 2026-10-01 -End 2027-03-31

# Then remove, without a message, exactly what the report found
.\Invoke-MeetingCleanupOnPrem.ps1 -FromReport .\reports\MeetingCleanupOnPrem_Report_20261006-101500 -Action Remove
```

Use the address of the person, or its **X500 address** when the account is gone from the directory (`/o=Contoso/ou=Exchange Administrative Group (...)/cn=Recipients/cn=...`): it is the organizer address the meetings still carry. A deleted organizer cannot cancel: the copies are removed silently. `AllMailboxes` needs Exchange PowerShell; with a list of the team instead: `-SearchIn Mailboxes -MailboxFile .\team.txt`.

### 2.4 Give the meetings of a person to someone else

```powershell
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer john.doe@contoso.com -Action Transfer -NewOrganizer jane.roe@contoso.com
```

Jane becomes the organizer of every meeting still to come. Exchange Server cannot move a meeting to another organizer: each meeting is **re-created** in Jane's calendar and sent to every attendee and room in **one invitation** (a series from its next occurrence); John's meeting is cancelled with the message *This meeting is now organized by Jane Roe*, or its copies are removed silently when his mailbox is gone. The report opens on its **Transfers** tab: for each meeting, from whom to whom, the new meeting and its invitation, what became of the old one. Details: [developer guide, Transfer to a new organizer](MeetingCleanupOnPrem-Guide.md#transfer-to-a-new-organizer).

### 2.5 One meeting, or one series, without a message

```powershell
.\Invoke-MeetingCleanupOnPrem.ps1 -Organizer megan.bowen@contoso.com -Subject 'Weekly sales review' -Action Remove
```

The copies of the attendees and the rooms are removed, nobody receives anything; a series goes whole. The organizer's own meeting stays (*Kept*) — choose *Cancel* to cancel it with a message.

### 2.6 Rooms closed for works

```powershell
.\Invoke-MeetingCleanupOnPrem.ps1 -Room room-paris-01@contoso.com, room-paris-02@contoso.com -Start 2026-11-02 -End 2026-11-13 `
    -Action Cancel -Comment 'The rooms of the 1st floor are closed for works.'
```

Every meeting of these rooms in the period, whoever organized it. A series loses only its occurrences in the period, and goes on after. The period is required with an action.

### 2.7 The leavers of the month

```powershell
.\Invoke-MeetingCleanupOnPrem.ps1 -OrganizerFile .\leavers-2026-10.txt -SearchIn Organizer, Rooms
```

A text file with one address per line, or a CSV file (`PrimarySmtpAddress`, `UserPrincipalName`, `Address`...): each mailbox is read once for all of them, and the report has an *Organizers* tab.

### 2.8 Undo a Remove

```powershell
.\Invoke-MeetingCleanupOnPrem.ps1 -Action Restore -FromReport .\reports\MeetingCleanupOnPrem_Remove_20261006-001346
```

The copies come back from Recoverable Items, as they were, **without a message**: an accepted room is busy again. It works for the retention of deleted items (**14 days** by default) and needs the role of [Rights for Restore](MeetingCleanupOnPrem-Guide.md#rights-for-restore). A cancelled or transferred meeting, and an occurrence removed in rooms mode, cannot be restored.

### 2.9 In a scheduled task

```powershell
pwsh -NoProfile -File C:\Tools\MeetingCleanupOnPrem\Invoke-MeetingCleanupOnPrem.ps1 -OrganizerFile C:\Tools\leavers.txt -Action Cancel -Comment 'This meeting is cancelled.' -Force
```

`-Force` skips the confirmation (there is no one to type *YES*). The service account signs in from `Connection.CredentialFile` (1.1). The exit code says how it went (chapter 4); the output and the log of the day keep every line.

<!-- icon: filter -->
## 3. Choose the meetings

| Parameter | Values | Example |
|---|---|---|
| `-Organizer` | One or more addresses (any alias), or the X500 address of a deleted mailbox | `-Organizer megan.bowen@contoso.com` |
| `-OrganizerFile` | A list of organizers: text (one per line) or CSV | `-OrganizerFile .\leavers.txt` |
| `-Room` · `-RoomFile` | Rooms mode: every meeting of these rooms, whatever its organizer | `-Room room-paris-01@contoso.com` |
| `-Start` · `-End` | The period (an end date without a time is included). Default: the coming year | `-Start 2026-11-01 -End 2026-12-31` |
| `-Subject` | The subject contains this text (`*` and `?` allowed) | `-Subject 'Weekly*'` |
| `-MeetingId` | Only these meetings (column *MeetingId* of a report) | `-MeetingId 040000008200E0...` |
| `-SearchIn` | Where to search (below). Default: `Organizer`, `Rooms` | `-SearchIn Rooms, AllMailboxes` |
| `-Mailbox` · `-MailboxFile` | The mailboxes of `-SearchIn Mailboxes` | `-MailboxFile .\team.txt` |
| `-FromReport` | Act on exactly the meetings of a reviewed report | `-FromReport .\reports\MeetingCleanupOnPrem_Report_...` |
| `-Action` | `Report` (default), `Remove`, `Cancel`, `Transfer`, `Restore` | `-Action Cancel` |
| `-Comment` | The message of a cancellation (or of the old organizer in a transfer) | `-Comment 'This meeting is cancelled.'` |
| `-NewOrganizer` | *Transfer*: the new organizer (a mailbox of the organization) | `-NewOrganizer jane.roe@contoso.com` |
| `-Force` | No confirmation: scheduled task, script | `-Force` |
| `-NoReport` | No CSV and no HTML: an action writes its backup and `Summary.json` only, a report writes nothing | `-NoReport` |

| `-SearchIn` | Searches | When |
|---|---|---|
| `Organizer` | The organizer's calendar | The mailbox still exists (skipped when Exchange PowerShell says it is gone). |
| `Rooms` | Every room mailbox (Exchange PowerShell; `Search.AllRooms = $false` to search only the rooms given), and the rooms of `Search.Rooms` / `Search.RoomFile` | Always useful: a meeting with a room is found even when its organizer is gone. |
| `Mailboxes` | The mailboxes of `-Mailbox` / `-MailboxFile` | A deleted organizer, meetings without room: his team. |
| `AllMailboxes` | Every mailbox of the organization (Exchange PowerShell) | A deleted organizer, meetings without room, team unknown. |

A series is found when one of its occurrences falls in the period, and is handled as a whole (in rooms mode: its occurrences in the period only).

The connection comes from the configuration; these parameters override it for one run: `-EwsUrl`, `-EwsMailbox` (the account that signs in), `-Authentication Windows | Basic`, `-AccessMode Impersonation | Delegate | Self`, `-CredentialUser` (asks for the password), `-ManagementShellMode Existing | Auto | Rps`, `-ManagementShellServer`, `-ManagementShellUri`, `-ConfigPath`.

<!-- icon: file -->
## 4. Results

Each run writes a new folder under `reports\`, named after the action and the time (`MeetingCleanupOnPrem_Remove_20261006-001346`). The console shows it at the end, with the next command to run:

- **`MeetingCleanupOnPrem.html`** — the report: self-contained, it can be sent alone. Tiles, the search, and the *Meetings*, *Copies* and *Organizers* tabs, searchable and sortable; after a *Transfer*, the **Transfers** tab first: per meeting, from whom to whom, the new meeting and what became of the old one.
- **`MeetingCleanupOnPrem-Meetings.csv`**, **`-Copies.csv`**, **`-Organizers.csv`** (and **`-Transfers.csv`** after a transfer) — the same data, separator `;`, open directly in Excel.
- **`MeetingCleanupOnPrem-Summary.json`** — the whole result, used by `-FromReport` and *Restore*.
- **`MeetingCleanupOnPrem-Backup.json`** — Remove, Cancel, Transfer: every meeting in full (subject, body, attendees, rooms, recurrence), written **before** any change.

| Meeting status | Meaning |
|---|---|
| **Found** | Report only: nothing was changed. |
| **Removed** · **Cancelled** | The copies of the attendees and rooms removed (the organizer's meeting *Kept*) · cancelled by its organizer, the copies left removed. |
| **Transferred** | Re-created by the new organizer, the old meeting cancelled or removed. |
| **Restored** | Every copy removed by the run is back. |
| **Skipped** | Left as it is (not selected, already over, cancelled before...): the reason is in the report. |
| **Partial** · **Failed** | Some · every request failed: each copy gives the Exchange error. |

| Exit code | Meaning |
|---|---|
| `0` | Completed. |
| `2` | Finished with warnings: a copy not removed, a mailbox not read, a calendar read in part. The report says which. |
| `1` | Failed: read the error in red, and the log of the day in `logs\`. |

| Message | What to do |
|---|---|
| *EWS could not open the calendar: HTTP 401* | Wrong password, or the authentication of the EWS virtual directory: `Connection.WindowsPackage = 'NTLM'` when Kerberos to the EWS name fails. |
| *ErrorImpersonateUserDenied* · *ErrorAccessDenied* | ApplicationImpersonation missing for this mailbox (outside the management scope), or not applied yet. |
| *ErrorNonExistentMailbox* | No mailbox with this address (deleted, or an alias of another domain). |
| *Exchange RPS connection failed* | `ManagementShell.ServerFqdn` / `ConnectionUri`, Kerberos to `http://<server>/PowerShell/`, or `RemotePowerShellEnabled` of the account. |
| *Get-RecoverableItems is not available* | The role Mailbox Import Export is missing, or not applied yet (open a new session). |
| *more than 500 items at ..., raise Connection.PageSize* | A calendar holds more items starting at the same time than a page: raise `Connection.PageSize` (up to 1,000). |
| *Confirmation needed: run interactively, or add -Force* | An action without a console (scheduled task): add `-Force`. |

Anything else: [developer guide, Appendix A — Troubleshooting](MeetingCleanupOnPrem-Guide.md#appendix-a---troubleshooting); every status and column of the report: [developer guide, chapter 10](MeetingCleanupOnPrem-Guide.md#10-reading-the-report).

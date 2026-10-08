<#
.SYNOPSIS
    Meeting Cleanup for Exchange Server On-Premises.

.DESCRIPTION
    Searches organizer, attendee and room copies through EWS (Exchange Server 2016, 2019, SE), with the
    directory, groups and Recoverable Items through Exchange PowerShell (local session or remote PowerShell).
      Report    read only.
      Remove    removes the attendee and room copies silently (DeleteItem SoftDelete, no message); the meeting
                stays in the organizer's calendar.
      Cancel    the organizer cancels the meeting (message to the attendees, rooms released), then the copies
                left are removed silently.
      Restore   -FromReport <Remove run>: the copies removed come back from Recoverable Items.
      Transfer  -NewOrganizer: the meeting is re-created and sent by the new organizer (a series from its next
                occurrence), the old one is cancelled by its organizer or removed silently.
    Remove, Cancel and Transfer write a backup first and ask to type YES (-Force skips it).
    A series is acted on whole; -SeriesScope Occurrences (Search.SeriesScope) limits it to its occurrences in the
    period (a period of one day: one occurrence; give -Start and -End with an action). Rooms mode always does.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
    Console : a live progress line (bar, part done, time left) in an interactive console; colours off when the
              output is redirected or NO_COLOR is set (MCO_FORCE_COLOR=1 forces them); MCO_ICONS = Emoji |
              Symbols | Ascii forces the style of the icons.
    Exit codes : 0 = completed, 1 = failed, 2 = finished with warnings (a copy not removed or restored, a mailbox not read...).
    Documentation : docs\MeetingCleanupOnPrem-UserGuide.html (user guide: prerequisites, everyday commands) and
                    docs\MeetingCleanupOnPrem-Guide.html (developer guide); sources: docs\*.md
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string[]]$Organizer,
    [string]$OrganizerFile,
    [string[]]$Room,
    [string]$RoomFile,
    [string[]]$Mailbox,
    [string]$MailboxFile,
    [datetime]$Start,
    [datetime]$End,
    [string]$Subject,
    [ValidateSet('Whole', 'Occurrences')][string]$SeriesScope,
    [string[]]$MeetingId,
    [ValidateSet('Organizer', 'Rooms', 'Mailboxes', 'AllMailboxes')][string[]]$SearchIn,
    [ValidateSet('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')][string]$Action = 'Report',
    [string]$Comment,
    [string]$NewOrganizer,
    [string]$FromReport,
    [switch]$Force,
    [switch]$NoReport,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\MeetingCleanupOnPrem.config.psd1'),
    [string]$EwsUrl,
    [string]$EwsMailbox,
    [ValidateSet('Windows', 'Basic')][string]$Authentication,
    [ValidateSet('Self', 'Delegate', 'Impersonation')][string]$AccessMode,
    [string]$CredentialUser,
    [ValidateSet('Existing', 'Auto', 'Rps')][string]$ManagementShellMode,
    [string]$ManagementShellServer,
    [string]$ManagementShellUri
)

$ErrorActionPreference = 'Stop'
# Numbers and dates of the console and the report in one format (1,254 - 2026-10-06), whatever the Windows culture.
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$exitCode = 1
$loaded = $false
$dot = [char]0x00B7

function Confirm-McoAction {
    <# Shows the plan and asks to type YES; throws when the answer is not YES or the console cannot ask. #>
    param([string]$Title, [string[]]$Lines)
    Write-Host ''
    Write-Host "  $Title - $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Yellow
    foreach ($line in $Lines) { Write-Host "    - $line" -ForegroundColor Yellow }
    if ([Console]::IsInputRedirected -or -not [Environment]::UserInteractive) { throw 'Confirmation needed: run interactively, or add -Force.' }
    if ((Read-Host '  Type YES to continue') -cne 'YES') { throw 'Cancelled: nothing was changed.' }
    Write-McoLog 'INFO' "Confirmed by $([Environment]::UserName): $($Lines -join ' | ')"
}

try {
    Import-Module (Join-Path $PSScriptRoot 'MeetingCleanupOnPrem.psd1') -Force
    $loaded = $true
    $settings = Import-McoConfiguration -Path $ConfigPath -Root $PSScriptRoot
    foreach ($pair in @(
        @('EwsUrl', $EwsUrl), @('Mailbox', $EwsMailbox), @('Authentication', $Authentication), @('AccessMode', $AccessMode), @('CredentialUser', $CredentialUser),
        @('ManagementShellMode', $ManagementShellMode), @('ManagementShellServer', $ManagementShellServer), @('ManagementShellUri', $ManagementShellUri)
    )) { if ($pair[1]) { $settings[$pair[0]] = $pair[1] } }
    $check = Test-McoConfiguration $settings
    if (-not $check.IsValid) { throw ("Invalid value:`n - " + ($check.Problems -join "`n - ")) }
    $logPath = Start-McoLog -Directory $settings.LogPath -RetentionDays $settings.LogRetentionDays
    $request = New-McoRequest -Settings $settings -Organizer $Organizer -OrganizerFile $OrganizerFile -Room $Room -RoomFile $RoomFile -Mailbox $Mailbox -MailboxFile $MailboxFile -Start $(if ($PSBoundParameters.ContainsKey('Start')) { $Start } else { $null }) -End $(if ($PSBoundParameters.ContainsKey('End')) { $End } else { $null }) -Subject $Subject -MeetingId $MeetingId -SearchIn $SearchIn -Action $Action -Comment $Comment -FromReport $FromReport -NewOrganizer $NewOrganizer -SeriesScope $SeriesScope
    $requestCheck = Test-McoRequest $request
    if (-not $requestCheck.IsValid) { throw ("Cannot run:`n - " + ($requestCheck.Problems -join "`n - ")) }
    Write-McoRunBanner -Settings $settings -Request $request -LogPath $logPath -NoReport:$NoReport

    $credential = $null
    if ($settings.CredentialFile) { $credential = Get-McoSavedCredential -Settings $settings }
    elseif ($settings.Authentication -eq 'Basic' -or $settings.CredentialUser) {
        $user = if ($settings.CredentialUser) { $settings.CredentialUser } else { $settings.Mailbox }
        $credential = Get-Credential -UserName $user -Message 'Credentials used for the Exchange Server EWS endpoint'
    }
    $rpsCredential = $credential
    if ($settings.ManagementShellCredentialUser -and (-not $rpsCredential -or $rpsCredential.UserName -ne $settings.ManagementShellCredentialUser)) {
        $rpsCredential = Get-Credential -UserName $settings.ManagementShellCredentialUser -Message 'Credentials used for Exchange Server Remote PowerShell'
    }
    elseif ($settings.ManagementShellMode -eq 'Rps' -and $settings.ManagementShellAuthentication -eq 'Basic' -and -not $rpsCredential) {
        $rpsCredential = Get-Credential -Message 'Credentials used for Exchange Server Remote PowerShell'
    }
    # Exchange Server, then: Restore (copies of the run, restore, verify) | search (organizers, mailboxes, search,
    # attendees) or the meetings of a report | the action and its verification | the report.
    $verify = [int][bool]$settings.Verify
    $steps = switch ($Action) {
        'Restore' { 3 + $verify }
        'Report' { 5 }
        default { $(if ($FromReport) { 2 } else { 5 }) + 1 + $verify }
    }
    if (-not $NoReport -or $Action -ne 'Report') { $steps++ }
    Initialize-McoSteps -Total $steps
    Write-McoNextStep 'Exchange Server' 'Server'
    $shell = Connect-McoExchangeShell -Settings $settings -Credential $rpsCredential
    $hasCmdlets = [bool](Get-Command Get-Recipient -ErrorAction SilentlyContinue)
    Write-McoItem $(if ($shell.Connected -or $hasCmdlets) { 'Ok' } else { 'Warn' }) ("Exchange PowerShell: {0}" -f $(if ($shell.Connected) { $shell.Detail } elseif ($hasCmdlets) { 'cmdlets of the current session' } else { 'not available (aliases, groups, rooms and restore need it)' })) -Icon Shield
    $ews = Connect-McoEws -Settings $settings -Credential $credential
    Write-McoItem Ok ("EWS {0} {1} {2} {1} {3}{4}" -f $ews.EwsUrl, $dot, $(if ($ews.ServerVersion) { $ews.ServerVersion } else { 'connected' }), $settings.AccessMode, $(if ($settings.AccessMode -ne 'Self') { " ($($settings.Mailbox) signs in)" } else { '' })) -Icon Server
    if ($Action -eq 'Restore') {
        Write-McoNextStep 'Copies removed by the run' 'File'
        $result = Import-McoRestoreSource -Path $FromReport -MeetingId $MeetingId
        $plan = Get-McoRestorePlan $result
        Write-McoItem Ok ('{0} meeting(s) of the {1} run {2} {3} cop{4} to restore' -f $result.Meetings.Count, $result.SourceAction, $dot, $plan.Restore.Count, $(if ($plan.Restore.Count -eq 1) { 'y' } else { 'ies' })) -Icon File
        Write-McoMeetingTable -Meetings @($result.Meetings)
        if ($plan.Restore.Count -and -not $Force) { Confirm-McoAction 'Restore from Recoverable Items' $plan.Lines }
        $result = Invoke-McoRestore -Settings $settings -Result $result
        Write-McoMeetingTable -Meetings @($result.Meetings)
        $runPath = New-McoRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action 'Restore'
    }
    else {
        if ($FromReport) {
            Write-McoNextStep 'Meetings of the report' 'File'
            $result = Import-McoRestoreSource -Path $FromReport -MeetingId $MeetingId
            $result.Action = 'Report'
            Write-McoItem Ok ('{0} meeting(s), {1} cop{2} from {3}' -f $result.Counts.Meetings, $result.Counts.Copies, $(if ($result.Counts.Copies -eq 1) { 'y' } else { 'ies' }), $result.FromReport) -Icon File
        }
        else { $result = Find-McoMeetings -Settings $settings -Request $request }
        Write-McoMeetingTable -Meetings @($result.Meetings)
        $runPath = $null
        if ($Action -eq 'Transfer' -and @($result.Meetings).Count) {
            $target = Resolve-McoNewOrganizer -Address $NewOrganizer
            $plan = Get-McoTransferPlan -Result $result -NewOrganizer $target -Method $settings.TransferMethod -Comment $request.Comment
            if (-not $Force) { Confirm-McoAction 'Transfer' (@($plan.Lines) + 'A backup of the meetings is written first.') }
            $runPath = New-McoRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action 'Transfer'
            $result = Invoke-McoTransfer -Settings $settings -Result $result -Plan $plan -Comment $request.Comment -BackupPath (Join-Path $runPath "$($settings.ReportPrefix)-Backup.json")
            Write-McoMeetingTable -Meetings @($result.Meetings)
        }
        elseif ($Action -in 'Remove', 'Cancel' -and @($result.Meetings).Count) {
            $plan = Get-McoCleanupPlan -Result $result -Action $Action
            if (-not $Force) {
                $lines = @($plan.Lines)
                if ($Action -eq 'Cancel') { $lines += "Message: $($request.Comment)" }
                $lines += 'A backup of the meetings is written first; the copies removed can be restored for the retention of deleted items (-Action Restore -FromReport <report folder>).'
                Confirm-McoAction $Action $lines
            }
            $runPath = New-McoRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action $Action
            $result = Invoke-McoCleanup -Settings $settings -Result $result -Action $Action -Comment $request.Comment -BackupPath (Join-Path $runPath "$($settings.ReportPrefix)-Backup.json")
            Write-McoMeetingTable -Meetings @($result.Meetings)
        }
        elseif ($Action -ne 'Report') { $result.Action = $Action }
    }
    $reportText = 'none (-NoReport)'
    if (-not $NoReport -or $runPath) {
        Write-McoNextStep 'Report' 'Report'
        $report = Export-McoReport -Result $result -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats $settings.ReportFormats -Delimiter $settings.CsvDelimiter -Directory $runPath -SummaryOnly:$NoReport
        foreach ($f in $report.Files.Values) { Write-McoItem Ok $f -Icon File }
        $reportText = if ($report.Files.Contains('Html')) { $report.Files.Html } else { $report.Directory }
    }
    Write-McoRunSummary -Result $result -ReportText $reportText -LogPath $logPath
    $exitCode = switch ($result.Status) { 'Completed' { 0 } 'Failed' { 1 } default { 2 } }
}
catch {
    if ($loaded) {
        $errorText = [string]$_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($errorText)) { $errorText = 'Unknown error.' }
        Write-McoItem Fail $errorText
        $stack = [string]$_.ScriptStackTrace
        if ($stack) { Write-McoLog 'ERROR' ($stack -replace '\r?\n', ' | ') }
        Write-Host ''
    }
    else { Write-Host "Meeting Cleanup On-Prem: $($_.Exception.Message)" -ForegroundColor Red }
    $exitCode = 1
}
finally {
    if ($loaded) { Disconnect-McoEws; Disconnect-McoExchangeShell; Stop-McoLog }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode

<#
.SYNOPSIS
    Meeting Cleanup On-Prem - report files (dot-sourced by MeetingCleanupOnPrem.psm1).

.DESCRIPTION
    One folder per execution, <FilePrefix>_<Action>_<yyyyMMdd-HHmmss>, with:
      <prefix>-Organizers.csv one row per organizer: address, state of the mailbox, meetings and copies, results
      <prefix>-Meetings.csv   one row per meeting: subject, organizer, start, series, organizer copy, copies, status
      <prefix>-Copies.csv     one row per mailbox: role, how it was found, action, result, EWS status, verified
      <prefix>-Transfers.csv  Transfer: one row per meeting: old and new organizer, new meeting, old copies
      <prefix>-Summary.json   the whole result, for scripts and for -FromReport (replay, restore)
      <prefix>-Backup.json    Remove, Cancel and Transfer: the meetings and copies as they were, written before any change
      <prefix>.html           self-contained dashboard (templates\Report.template.html)
    CSV files: UTF-8 with BOM, configurable delimiter, text cells starting with = + - @ are prefixed with an
    apostrophe (no formula injection when opened in Excel). No token or secret is ever part of the result.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

$script:ReportColumns = [ordered]@{
    Meetings = @('MeetingId', 'Subject', 'Organizer', 'OrganizerName', 'Kind', 'Scope', 'Occurrences', 'OccurrencesSkipped', 'StartText', 'EndText', 'NextInPeriod', 'Recurrence', 'Location', 'OrganizerCopy', 'Copies', 'RoomCopies', 'AttendeeCopies', 'NotProcessed', 'Cancelled', 'Selected', 'Status', 'NewOrganizer', 'NewMeetingId', 'TransferMethod', 'Notes')
    Copies   = @('MeetingId', 'MeetingSubject', 'Organizer', 'Mailbox', 'Role', 'Via', 'Occurrence', 'Response', 'ShowAs', 'Cancelled', 'Action', 'Result', 'HttpStatus', 'Verified', 'ActionUtc', 'Detail', 'EventId')
    Organizers = @('Input', 'DisplayName', 'PrimaryAddress', 'State', 'Detail', 'Meetings', 'Series', 'Copies', 'Removed', 'Cancelled', 'Restored', 'Transferred', 'Failed')
    Transfers  = @('MeetingId', 'Subject', 'StartText', 'Kind', 'Recurrence', 'OldOrganizer', 'OldOrganizerName', 'OldOrganizerState', 'OldOrganizerDetail', 'NewOrganizer', 'Method', 'Status', 'NewMeetingId', 'NewMeeting', 'NewMeetingDetail', 'Invited', 'Rooms', 'OldOrganizerCopy', 'OldCopiesRemoved', 'OldCopiesFailed', 'OldCopiesLeft', 'Selected', 'Notes')
}

function Format-McoCsvCell {
    <# A CSV cell: text starting with = + - @ (or tab, CR) prefixed with an apostrophe (no formula injection in Excel), quoted when needed (compiled). #>
    param([AllowNull()][object]$Value, [Parameter(Mandatory = $true)][string]$Delimiter)
    [MeetingCleanupOnPremNative.Fast]::CsvCell($Value, $Delimiter)
}

function Write-McoCsv {
    <# A CSV file: UTF-8 with BOM, the columns given, one line per row (compiled). #>
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Delimiter = ';'
    )
    [MeetingCleanupOnPremNative.Fast]::WriteCsv($Rows, $Columns, $Path, $Delimiter)
}

function ConvertTo-McoEmbeddedJson {
    <# JSON safe inside a <script type="application/json"> block. #>
    param([AllowNull()][object]$Value)
    $json = ConvertTo-Json -InputObject $Value -Depth 10 -Compress
    if ([string]::IsNullOrEmpty($json)) { $json = 'null' }
    return $json.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function Get-McoMeetingRows {
    <# The meetings flattened for the CSV and the HTML: one copy per mailbox (an occurrence copy is counted once for its mailbox); the new organizer is not a copy (compiled). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    [MeetingCleanupOnPremNative.Fast]::MeetingTable($Result.Meetings).ToObjects()
}

function Get-McoCopyRows {
    <# One row per copy (compiled). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    [MeetingCleanupOnPremNative.Fast]::CopyTable($Result.Meetings).ToObjects()
}

function Get-McoOrganizerRows {
    <# One row per organizer of the run: its mailbox and what was found and done for its meetings (compiled). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    [MeetingCleanupOnPremNative.Fast]::OrganizerTable($Result.Organizers, $Result.Meetings).ToObjects()
}

function Get-McoTransferRows {
    <# One row per meeting of a transfer: old and new organizer, status, new meeting, old copies (compiled). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    [MeetingCleanupOnPremNative.Fast]::TransferTable($Result.Organizers, $Result.Meetings).ToObjects()
}

function New-McoRunFolder {
    <# New folder of a run under OutputPath: <Prefix>_<Action>_<yyyyMMdd-HHmmss>. #>
    param([Parameter(Mandatory = $true)][string]$OutputPath, [string]$Prefix = 'MeetingCleanup', [string]$Action = 'Report')
    $base = Join-Path $OutputPath ('{0}_{1}_{2}' -f $Prefix, $Action, (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $path = $base
    $n = 2
    while (Test-Path -LiteralPath $path) { $path = "$base-$n"; $n++ }
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

function Export-McoReport {
    <#
    .SYNOPSIS
        Writes the CSV, JSON and HTML files of one result in a new folder under OutputPath (or in Directory, the
        folder created before the action, where the backup already is).
    .PARAMETER SummaryOnly
        Only the Summary.json file (-NoReport of an action: the restore needs it).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$Prefix = 'MeetingCleanup',
        [ValidateSet('Csv', 'Html')][string[]]$Formats = @('Csv', 'Html'),
        [ValidateSet(';', ',', "`t")][string]$Delimiter = ';',
        [string]$Directory,
        [switch]$SummaryOnly
    )

    $runPath = if ($Directory) { [void][IO.Directory]::CreateDirectory($Directory); $Directory } else { New-McoRunFolder -OutputPath $OutputPath -Prefix $Prefix -Action ([string]$Result.Action) }
    # The rows (thousands for a large search) as compiled tables, written to CSV and JSON without PowerShell objects.
    $fast = [MeetingCleanupOnPremNative.Fast]
    $meetings = $fast::MeetingTable($Result.Meetings)
    $copies = $fast::CopyTable($Result.Meetings)
    $organizers = $fast::OrganizerTable($Result.Organizers, $Result.Meetings)
    # Transfer: one more table, the organizer change of each meeting (its own CSV and tab of the HTML report).
    $transfers = if ([string]$Result.Action -eq 'Transfer') { $fast::TransferTable($Result.Organizers, $Result.Meetings) } else { $null }
    $files = [ordered]@{}
    if ($Formats -contains 'Csv' -and -not $SummaryOnly) {
        $files.Meetings = Join-Path $runPath "$Prefix-Meetings.csv"
        $fast::WriteTableCsv($meetings, [string[]]$script:ReportColumns.Meetings, $files.Meetings, $Delimiter)
        $files.Copies = Join-Path $runPath "$Prefix-Copies.csv"
        $fast::WriteTableCsv($copies, [string[]]$script:ReportColumns.Copies, $files.Copies, $Delimiter)
        $files.Organizers = Join-Path $runPath "$Prefix-Organizers.csv"
        $fast::WriteTableCsv($organizers, [string[]]$script:ReportColumns.Organizers, $files.Organizers, $Delimiter)
        if ($transfers) {
            $files.Transfers = Join-Path $runPath "$Prefix-Transfers.csv"
            $fast::WriteTableCsv($transfers, [string[]]$script:ReportColumns.Transfers, $files.Transfers, $Delimiter)
        }
    }
    $files.Summary = Join-Path $runPath "$Prefix-Summary.json"
    # RunRemoved (restore: the copies of the source run, also in Meetings) is not written again, nor the full item
    # read by EWS for each copy (body, attendees...: Backup.json keeps it for the meetings acted on).
    $data = [ordered]@{}
    foreach ($prop in $Result.PSObject.Properties) {
        if ($prop.Name -in 'RunRemoved', 'SourceRunUtc') { continue }
        if ($prop.Name -eq 'Meetings') {
            $data.Meetings = @(foreach ($m in @($Result.Meetings)) {
                    $row = [ordered]@{}
                    foreach ($mp in $m.PSObject.Properties) {
                        if ($mp.Name -ne 'Copies') { $row[$mp.Name] = $mp.Value; continue }
                        $row.Copies = @(foreach ($c in $m.Copies) {
                                $copy = [ordered]@{}
                                foreach ($cp in $c.PSObject.Properties) { if ($cp.Name -notin 'Event', 'RemovedUtc') { $copy[$cp.Name] = $cp.Value } }
                                $copy
                            })
                    }
                    $row
                })
            continue
        }
        $data[$prop.Name] = $prop.Value
    }
    [IO.File]::WriteAllText($files.Summary, (ConvertTo-Json -InputObject $data -Depth 10), [Text.UTF8Encoding]::new($false))
    $backup = Join-Path $runPath "$Prefix-Backup.json"
    if (Test-Path -LiteralPath $backup) { $files.Backup = $backup }

    if ($Formats -contains 'Html' -and -not $SummaryOnly) {
        $summary = [ordered]@{}
        foreach ($key in 'Tool', 'Version', 'Action', 'Status', 'Error', 'StartedUtc', 'CompletedUtc', 'DurationSeconds', 'Request', 'Tenant', 'Organization', 'AppId', 'AppName', 'Organizers', 'Searched', 'Warnings', 'Counts', 'CleanupComment', 'FromReport', 'BackupFile', 'SourceAction', 'NewOrganizer') {
            # Direct assignment: a list of one item stays a list (a function output would unroll it).
            $summary[$key] = $null
            $prop = $Result.PSObject.Properties[$key]
            if ($prop) { $summary[$key] = $prop.Value }
        }
        $summary.ActionText = Get-McoActionText ([string]$Result.Action)
        $summary.GeneratedText = Format-McoDate ([datetime]::UtcNow) ([string](Get-McoProperty $Result.Request 'TimeZone'))
        $html = [IO.File]::ReadAllText((Join-Path $script:ToolRoot 'templates\Report.template.html'))
        $names = @(foreach ($o in @($Result.Organizers)) { if ($o.DisplayName) { $o.DisplayName } else { $o.Input } })
        $title = "Meeting Cleanup On-Prem | $($Result.Action) | $(if ($names.Count -le 3) { $names -join ', ' } else { "$($names.Count) organizers" })"
        $html = $html.Replace('{{TITLE}}', [Net.WebUtility]::HtmlEncode($title))
        $html = $html.Replace('{{SUMMARY_JSON}}', (ConvertTo-McoEmbeddedJson $summary))
        # The rows (thousands for a large search): compiled JSON writer, HTML-safe as well.
        $html = $html.Replace('{{MEETINGS_JSON}}', $fast::TableJson($meetings))
        $html = $html.Replace('{{COPIES_JSON}}', $fast::TableJson($copies))
        $html = $html.Replace('{{ORGANIZERS_JSON}}', $fast::TableJson($organizers))
        $html = $html.Replace('{{TRANSFERS_JSON}}', $(if ($transfers) { $fast::TableJson($transfers) } else { '[]' }))
        if ($html -match '\{\{[A-Z_]+\}\}') { throw "Report template marker not replaced: $($Matches[0])" }
        $files.Html = Join-Path $runPath "$Prefix.html"
        [IO.File]::WriteAllText($files.Html, $html, [Text.UTF8Encoding]::new($true))
    }
    [pscustomobject]@{ Directory = $runPath; Files = $files }
}
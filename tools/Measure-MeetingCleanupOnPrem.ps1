<#
.SYNOPSIS
    Measures Meeting Cleanup On-Prem on a simulated Exchange Server (tests\FakeEws.ps1): no server needed.

.DESCRIPTION
    Builds an organization of -Meetings meetings of org@contoso.test (each with -Attendees attendees out of 40 and
    one of 5 rooms, plus the meetings of the tests), then runs a whole search (organizer and rooms, then every
    attendee mailbox) and writes the report. Gives the time of each step and the EWS calls by operation, and writes
    the result (without timings) to -Out, to compare two versions on the same data:

        .\tools\Measure-MeetingCleanupOnPrem.ps1 -Out C:\Temp\new
        .\tools\Measure-MeetingCleanupOnPrem.ps1 -Root <folder of another version> -Out C:\Temp\old
        Compare-Object (Get-Content C:\Temp\old\search.json) (Get-Content C:\Temp\new\search.json)

    -Scenario All also removes, restores, cancels, transfers and searches a room (small organization: a removal
    leaves 3 seconds between copies with the same subject in one mailbox). -LatencyMs adds the time of an Exchange
    answer to each call. -Show: the console of a run (banner, steps, live progress line with the time left).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$Root = (Split-Path $PSScriptRoot -Parent),
    [string]$Out = (Join-Path ([IO.Path]::GetTempPath()) 'MeetingCleanupOnPrem-Measure'),
    [int]$Meetings = 600,
    [int]$Attendees = 6,
    [ValidateSet('Search', 'All')][string]$Scenario = 'Search',
    [int]$LatencyMs = 0,
    [switch]$Show
)

$ErrorActionPreference = 'Stop'
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
[void][IO.Directory]::CreateDirectory($Out)
Import-Module (Join-Path $Root 'MeetingCleanupOnPrem.psd1') -Force
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'tests\FakeEws.ps1')
$module = Get-Module MeetingCleanupOnPrem
& $module { param($quiet) $script:Quiet = $quiet } (-not $Show)

$settings = Import-McoConfiguration -Path (Join-Path $Root 'config\MeetingCleanupOnPrem.config.psd1') -Root $Root
$settings.DirectoryMode = 'None'; $settings.TimeZone = 'Romance Standard Time'; $settings.RestoreMode = 'Ews'
$settings.Rooms = @(1..7 | ForEach-Object { "room$_@contoso.test" })
$settings.OutputPath = Join-Path $Out 'reports'; $settings.LogPath = Join-Path $Out 'logs'
$install = {
    param([hashtable]$Store)
    $Store.LatencyMs = $LatencyMs
    Install-FakeEws -Store $Store -Settings $settings
}
# The result without what changes from one run to the other (times, item IDs of the simulated server).
$clean = {
    param($Result)
    $json = ConvertTo-Json -InputObject $Result -Depth 12
    $json = $json -replace '"(StartedUtc|CompletedUtc|ActionUtc|RestoredUtc|RemovedUtc|CreatedUtc|DurationSeconds|BackupFile|FromReport|Version|DateTimeCreated|LastModifiedUtc)":\s*("[^"]*"|[\d.]+|null)', '"$1": "-"'
    $json -replace '"(EventId|SeriesId|NewMeetingId|EntryId|ChangeKey)":\s*"[^"]*"', '"$1": "-"'
}
$times = [ordered]@{ Version = (Get-Module MeetingCleanupOnPrem).Version.ToString(); Meetings = $Meetings; LatencyMs = $LatencyMs }
$sw = [Diagnostics.Stopwatch]::new()

$store = New-FakeTenant -Extra $Meetings -ExtraAttendees $Attendees
& $install $store
$request = New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -Start ([datetime]'2030-01-01') -End ([datetime]'2030-12-31')
if ($Show) { Write-McoRunBanner -Settings $settings -Request $request }
$sw.Restart()
$result = Find-McoMeetings -Settings $settings -Request $request
$times.Search = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
$times.Copies = $result.Counts.Copies
$times.EwsCalls = [ordered]@{}
foreach ($k in $store.Calls.Keys) { $times.EwsCalls[$k] = $store.Calls[$k] }
& $clean $result | Set-Content (Join-Path $Out 'search.json')
if ($Show) { Write-McoMeetingTable -Meetings @($result.Meetings | Select-Object -First 12) }
$sw.Restart()
$report = Export-McoReport -Result $result -OutputPath $settings.OutputPath -Prefix 'Measure' -Formats Csv, Html
$times.Report = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
foreach ($f in 'Measure-Meetings.csv', 'Measure-Copies.csv', 'Measure-Organizers.csv') { Copy-Item (Join-Path $report.Directory $f) (Join-Path $Out $f) -Force }
$sw.Restart()
$plan = Get-McoCleanupPlan -Result $result -Action Remove
$times.Plan = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
$plan.Lines | Set-Content (Join-Path $Out 'plan.txt')

if ($Scenario -eq 'All') {
    # A small organization near today (the restore checks the calendars from 30 days ago to 400 days ahead).
    $near = [datetime]::UtcNow.Date.AddDays(3).AddHours(9)
    $period = @{ Start = [datetime]::UtcNow.Date; End = [datetime]::UtcNow.Date.AddDays(60) }
    $store = New-FakeTenant -From $near; & $install $store
    $found = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' @period)
    $folder = Join-Path $Out 'remove-run'
    $sw.Restart()
    $removed = Invoke-McoCleanup -Settings $settings -Result $found -Action Remove -BackupPath (Join-Path $folder 'MeetingCleanupOnPrem-Backup.json')
    $times.Remove = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
    & $clean $removed | Set-Content (Join-Path $Out 'remove.json')
    $null = Export-McoReport -Result $removed -OutputPath $settings.OutputPath -Prefix 'MeetingCleanupOnPrem' -Formats Csv -Directory $folder
    $sw.Restart()
    $restored = Invoke-McoRestore -Settings $settings -Result (Import-McoRestoreSource -Path $folder)
    $times.Restore = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
    & $clean $restored | Set-Content (Join-Path $Out 'restore.json')

    $store = New-FakeTenant -From $near; & $install $store
    $found = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' @period)
    $cancelled = Invoke-McoCleanup -Settings $settings -Result $found -Action Cancel -Comment 'Cancelled.'
    & $clean $cancelled | Set-Content (Join-Path $Out 'cancel.json')

    $store = New-FakeTenant -From $near; & $install $store
    $found = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' @period)
    $tplan = Get-McoTransferPlan -Result $found -NewOrganizer (Resolve-McoNewOrganizer -Address 'new@contoso.test') -Comment 'Now organized by {0}.'
    $sw.Restart()
    $moved = Invoke-McoTransfer -Settings $settings -Result $found -Plan $tplan -Comment 'Now organized by {0}.'
    $times.Transfer = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
    & $clean $moved | Set-Content (Join-Path $Out 'transfer.json')

    $store = New-FakeTenant; & $install $store
    $rooms = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Room 'room2@contoso.test' -Start ([datetime]'2030-01-01') -End ([datetime]'2030-01-20'))
    & $clean $rooms | Set-Content (Join-Path $Out 'rooms.json')
}
$times | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $Out 'timings.json')
[pscustomobject]$times

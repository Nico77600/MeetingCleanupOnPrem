function Get-McoCleanupPlan {
    <#
        What an action would do on the selected meetings, without doing it (confirmation, banner). One pass, no pipeline.
        A series is handled as a whole, unless it was limited to its occurrences in the period (rooms mode,
        -SeriesScope Occurrences): then each occurrence is acted on, and the occurrences left out in a reviewed
        report (SkippedOccurrences) are left as they are.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Result, [Parameter(Mandatory)][ValidateSet('Remove', 'Cancel')][string]$Action)
    $fast = [MeetingCleanupOnPremNative.Fast]
    $cancel = [Collections.Generic.List[object]]::new()
    $remove = [Collections.Generic.List[object]]::new()
    $keep = [Collections.Generic.List[object]]::new()
    $held = [Collections.Generic.List[object]]::new()
    $acted = [Collections.Generic.List[object]]::new()
    $notSelected = [Collections.Generic.List[object]]::new()
    $series = 0; $removeRooms = 0; $occMeetings = 0; $occurrences = 0; $cancelOccurrences = 0; $skippedOccurrences = 0
    $removeMailboxes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $keptMeetings = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($meeting in $Result.Meetings) {
        if (-not $meeting.Selected) { continue }
        # Cancel needs the organizer's own copy: removing only the attendees' copies would leave his meeting live.
        $orgItem = $false
        foreach ($c in $meeting.Copies) { if ($c.Role -eq 'Organizer' -and $c.EventId) { $orgItem = $true; break } }
        $why = if ($fast::Text($meeting, 'TransferMethod') -and $fast::Text($meeting, 'NewMeetingId')) { "transferred to $($fast::Text($meeting, 'NewOrganizer')) by the run of this report" }
            elseif ($Action -eq 'Cancel' -and [string]$meeting.OrganizerCopy -eq 'Present' -and -not $orgItem) { 'the copy of its organizer could not be read' }
            else { '' }
        if ($why) { $held.Add([pscustomobject]@{ Meeting = $meeting; Reason = $why }); continue }
        $acted.Add($meeting)
        if ($meeting.Kind -eq 'Series') { $series++ }
        # The occurrences left out: their copies are left as they are.
        $skip = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($k in @($fast::Prop($meeting, 'SkippedOccurrences'))) { $key = $fast::KeyText($k); if ($key) { [void]$skip.Add($key) } }
        if ($fast::Text($meeting, 'Scope') -eq 'Occurrences') { $occMeetings++; $occurrences += [Math]::Max(0, [int]$fast::Prop($meeting, 'Occurrences') - $skip.Count); $skippedOccurrences += $skip.Count }
        foreach ($copy in $meeting.Copies) {
            if (-not $copy.EventId -or $copy.Role -eq 'New organizer') { continue }
            if ($skip.Count -and $skip.Contains($fast::OccurrenceKeyOf($copy))) { $notSelected.Add($copy); continue }
            if ($copy.Role -eq 'Organizer') {
                if ($Action -eq 'Cancel') { $cancel.Add($copy); if ($fast::Text($copy, 'Occurrence')) { $cancelOccurrences++ } } else { $keep.Add($copy); [void]$keptMeetings.Add([string]$copy.MeetingId) }
            }
            elseif ($copy.Role -in 'Attendee', 'Room') {
                $remove.Add($copy); [void]$removeMailboxes.Add([string]$copy.Mailbox)
                if ($copy.Role -eq 'Room') { $removeRooms++ }
            }
        }
    }
    $selected = $acted.ToArray()
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add(('{0} meeting(s) selected ({1} series)' -f $selected.Count, $series))
    if ($cancel.Count) { $lines.Add(('{0} cancelled by their organizer: Exchange sends the cancellation to the attendees and releases the rooms' -f $(if ($cancelOccurrences) { '{0} meeting(s) and {1} occurrence(s)' -f ($cancel.Count - $cancelOccurrences), $cancelOccurrences } else { $cancel.Count }))) }
    $lines.Add(('{0} cop{1} removed without any message in {2} mailbox(es): {3} attendee(s), {4} room(s)' -f $remove.Count, $(if ($remove.Count -eq 1) { 'y' } else { 'ies' }), $removeMailboxes.Count, ($remove.Count - $removeRooms), $removeRooms))
    if ($keep.Count) { $lines.Add(('{0} meeting(s) stay in the calendar of their organizer (choose Cancel to cancel them with a message)' -f $keptMeetings.Count)) }
    if ($occMeetings) { $lines.Add(('{0} series limited to their occurrences in the period ({1} occurrence(s)): an occurrence removed is not kept in Recoverable Items, it cannot be restored' -f $occMeetings, $occurrences)) }
    if ($skippedOccurrences) { $lines.Add(('{0} occurrence(s) left out in the report: left as they are' -f $skippedOccurrences)) }
    if ($held.Count) {
        $reasons = [ordered]@{}
        foreach ($h in $held) { $reasons[$h.Reason] = 1 + [int]$reasons[$h.Reason] }
        $lines.Add(('{0} meeting(s) left as they are: {1}' -f $held.Count, ((@($reasons.Keys | ForEach-Object { "$($reasons[$_]) $_" })) -join '; ')))
    }
    [pscustomobject]@{ Action = $Action; Meetings = $selected; Cancel = $cancel.ToArray(); Remove = $remove.ToArray(); Keep = $keep.ToArray(); Held = $held.ToArray(); NotSelected = $notSelected.ToArray(); Lines = $lines.ToArray(); Text = ($lines -join " $($script:Dot) ") }
}

function Set-McoCopyResult {
    param([Parameter(Mandatory)]$Copy, [Parameter(Mandatory)][string]$Action, [Parameter(Mandatory)]$Response, [Parameter(Mandatory)][string]$Success)
    $Copy.Action = $Action
    $Copy.HttpStatus = [int]$Response.HttpStatus
    if ($Response.ResponseClass -eq 'Success') { $Copy.Result = $Success; $Copy.Detail = '' }
    elseif ($Response.ResponseCode -eq 'ErrorItemNotFound') { $Copy.Result = 'Already gone'; $Copy.Detail = 'not in the calendar any more' }
    else { $Copy.Result = 'Failed'; $Copy.Detail = Get-McoEwsFailureText $Response }
}

function Save-McoBackup {
    <#
        Before any change: the full content of each meeting acted on (organizer copy when there is one, else an
        attendee copy: subject, body, attendees, rooms, recurrence...) and the state of each copy. Nothing is changed
        when the file cannot be written.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Result, [Parameter(Mandatory)][pscustomobject]$Plan, [Parameter(Mandatory)][string]$Path)
    $fast = [MeetingCleanupOnPremNative.Fast]
    $saved = 0
    $items = foreach ($meeting in @($Plan.Meetings)) {
        $reference = $null; $rank = 9
        foreach ($c in $meeting.Copies) {
            if (-not $c.EventId) { continue }
            $rk = switch ($c.Role) { 'Organizer' { 0 } 'Attendee' { 1 } 'Room' { 2 } default { 9 } }
            if ($rk -lt $rank) { $rank = $rk; $reference = $c }
        }
        $event = if ($reference) { $fast::Prop($reference, 'Event') } else { $null }
        $copies = [Collections.Generic.List[object]]::new()
        foreach ($c in $meeting.Copies) {
            if (-not $c.EventId) { continue }
            $copies.Add([ordered]@{ Mailbox = $c.Mailbox; Role = $c.Role; EventId = $c.EventId; Subject = $c.Subject; Response = $c.Response; Occurrence = $fast::Text($c, 'Occurrence'); SeriesId = $fast::Text($c, 'SeriesId') })
        }
        $saved += $copies.Count
        [ordered]@{
            MeetingId = $meeting.MeetingId; Subject = $meeting.Subject; Organizer = $meeting.Organizer; OrganizerName = $meeting.OrganizerName; Kind = $meeting.Kind
            StartText = $meeting.StartText; EndText = $meeting.EndText; Scope = $meeting.Scope; ReferenceMailbox = if ($reference) { $reference.Mailbox } else { '' }
            Event = if ($event) {
                [ordered]@{
                    Subject = $event.Subject; Start = ([datetime]$event.Start).ToString('o'); End = ([datetime]$event.End).ToString('o'); Organizer = $event.Organizer; Location = $event.Location
                    RequiredAttendees = @($event.RequiredAttendees); OptionalAttendees = @($event.OptionalAttendees); Resources = @($event.Resources)
                    Body = $event.Body; RecurrenceXml = $event.RecurrenceXml; TimeZoneId = $event.TimeZoneId; AppointmentType = $event.AppointmentType
                }
            } else { $null }
            Copies = $copies.ToArray()
        }
    }
    $data = [ordered]@{ Tool = 'Meeting Cleanup On-Prem'; Version = $script:ToolVersion; Kind = 'Backup'; CreatedUtc = [datetime]::UtcNow.ToString('o'); Action = $Plan.Action; Tenant = $Result.Tenant; Meetings = @($items) }
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $Path))
    [IO.File]::WriteAllText($Path, (ConvertTo-Json $data -Depth 20), [Text.UTF8Encoding]::new($false))
    Write-McoItem Ok ('Backup of {0} meeting(s) and {1} cop{2} before any change: {3}' -f @($items).Count, $saved, $(if ($saved -eq 1) { 'y' } else { 'ies' }), $Path) -Icon File
    $Path
}

function Remove-McoCopyWaves {
    <#
        Silent removal (DeleteItem SoftDelete, SendMeetingCancellations SendToNone) with the time of each removal.
        Copies with the same subject in one mailbox (a room shows the organizer's name) are removed one after the
        other, at least 3 seconds apart: their order is then certain in Recoverable Items, for the restore.
    #>
    param([Parameter(Mandatory)][object[]]$Copies, [string]$DeleteType = 'SoftDelete', [string]$ProgressText = 'copies removed', [switch]$NoProgress)
    # Waves: the n-th copy of each (mailbox, subject) group in wave n. An occurrence is alone in its group.
    $groups = [ordered]@{}
    foreach ($c in $Copies) {
        $key = if ([MeetingCleanupOnPremNative.Fast]::Text($c, 'Occurrence')) { "occ|$($c.Mailbox)|$($c.EventId)" } else { '{0}|{1}' -f $c.Mailbox, ([string]$c.Subject).Trim().ToLowerInvariant() }
        $list = $groups[$key]
        if (-not $list) { $list = [Collections.Generic.List[object]]::new(); $groups[$key] = $list }
        $list.Add($c)
    }
    $waves = [Collections.Generic.List[object]]::new()
    foreach ($list in $groups.Values) {
        for ($k = 0; $k -lt $list.Count; $k++) {
            while ($waves.Count -le $k) { $waves.Add([Collections.Generic.List[object]]::new()) }
            $waves[$k].Add($list[$k])
        }
    }
    $lastDone = [datetime]::MinValue
    $done = 0
    foreach ($wave in $waves) {
        while (([datetime]::UtcNow - $lastDone).TotalSeconds -lt 3) { Start-Sleep -Milliseconds 250 }
        foreach ($copy in $wave) {
            $response = Invoke-McoDeleteItem -Mailbox $copy.Mailbox -EventId $copy.EventId -DeleteType $DeleteType
            $copy.ActionUtc = [datetime]::UtcNow.ToString('o')
            Set-McoCopyResult -Copy $copy -Action 'Remove' -Response $response -Success 'Removed'
            $done++
            if (-not $NoProgress) { Write-McoProgress ($done / [Math]::Max(1, $Copies.Count)) ('{0:N0}/{1:N0} {2}' -f $done, $Copies.Count, $ProgressText) }
        }
        $lastDone = [datetime]::UtcNow
    }
    if ($waves.Count -gt 1) { Write-McoLog 'INFO' "Removal in $($waves.Count) waves: copies with the same subject in one mailbox removed 3 s apart." }
}

function Invoke-McoCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Settings, [Parameter(Mandatory)][pscustomobject]$Result,
        [Parameter(Mandatory)][ValidateSet('Remove', 'Cancel')][string]$Action, [string]$Comment, [string]$BackupPath
    )
    $started = [datetime]::UtcNow
    $dot = $script:Dot
    $plan = Get-McoCleanupPlan -Result $Result -Action $Action
    $Result.Action = $Action
    $Result | Add-Member -NotePropertyName CleanupComment -NotePropertyValue $(if ($Action -eq 'Cancel') { $Comment } else { '' }) -Force
    foreach ($m in $Result.Meetings) { if (-not $m.Selected) { $m.Status = 'Skipped' } }
    foreach ($h in $plan.Held) {
        $h.Meeting.Status = 'Skipped'; $h.Meeting.Notes.Add("Not acted on: $($h.Reason).")
        foreach ($c in $h.Meeting.Copies) { if ($c.EventId) { $c.Action = 'None'; $c.Result = 'Not processed'; $c.Detail = "left as it is: $($h.Reason)" } }
    }
    foreach ($c in $plan.NotSelected) { $c.Action = 'None'; $c.Result = 'Skipped'; $c.Detail = 'occurrence left out in the report: left as it is' }
    Write-McoNextStep $(if ($Action -eq 'Cancel') { 'Cancel and clean' } else { 'Remove silently' }) $(if ($Action -eq 'Cancel') { 'Cancel' } else { 'Trash' })
    foreach ($line in $plan.Lines) { Write-McoItem Info $line }
    if ($BackupPath -and $plan.Meetings.Count) { $Result | Add-Member -NotePropertyName BackupFile -NotePropertyValue (Save-McoBackup $Result $plan $BackupPath) -Force }

    # ---- organizer: cancel (Cancel) or keep (Remove) -------------------------------------------------------
    # A cancellation that failed holds the copies of that meeting (of that occurrence, for a series by occurrences).
    $fast = [MeetingCleanupOnPremNative.Fast]
    $failedMeetings = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $sent = 0; $tried = 0; $occurrenceCancels = 0
    foreach ($copy in $plan.Cancel) {
        $response = Invoke-McoCancelItem -Mailbox $copy.Mailbox -EventId $copy.EventId -Comment $Comment
        $copy.ActionUtc = [datetime]::UtcNow.ToString('o')
        Set-McoCopyResult -Copy $copy -Action 'Cancel' -Response $response -Success 'Cancelled'
        $tried++
        if ($fast::Text($copy, 'Occurrence')) { $occurrenceCancels++ }
        if ($copy.Result -eq 'Cancelled') { $sent++ }
        if ($copy.Result -eq 'Failed') { [void]$failedMeetings.Add(('{0}|{1}' -f $copy.MeetingId, $fast::OccurrenceKeyOf($copy))); Write-McoItem Fail "Cancel $($copy.Mailbox): $($copy.Detail)" }
        Write-McoProgress ($tried / [Math]::Max(1, $plan.Cancel.Count)) ('{0:N0}/{1:N0} cancellations sent' -f $tried, $plan.Cancel.Count)
    }
    if ($plan.Cancel.Count) { Write-McoItem $(if ($failedMeetings.Count) { 'Warn' } else { 'Ok' }) ('{0} {1} cancelled by their organizer {2} {3} failed' -f $sent, $(if ($occurrenceCancels) { 'meeting(s) or occurrence(s)' } else { 'meeting(s)' }), $dot, $failedMeetings.Count) -Icon Cancel }
    foreach ($copy in $plan.Keep) { $copy.Action = 'Keep'; $copy.Result = 'Kept'; $copy.Detail = "left in the organizer's calendar (Remove is silent; Cancel cancels it with a message)" }

    # ---- attendee and room copies: silent removal ------------------------------------------------------------
    if ($plan.Cancel.Count) { Start-Sleep -Seconds 5 }
    $toRemove = [Collections.Generic.List[object]]::new()
    foreach ($c in $plan.Remove) {
        if ($failedMeetings.Contains(('{0}|{1}' -f $c.MeetingId, $fast::OccurrenceKeyOf($c))) -or $failedMeetings.Contains("$($c.MeetingId)|")) { $c.Action = 'None'; $c.Result = 'Not done'; $c.Detail = 'the cancellation by the organizer failed: copy left as it was' }
        else { $toRemove.Add($c) }
    }
    if ($toRemove.Count) {
        Remove-McoCopyWaves -Copies $toRemove.ToArray()
        $removed = 0; $gone = 0; $failed = [Collections.Generic.List[object]]::new()
        foreach ($c in $toRemove) { switch ($c.Result) { 'Removed' { $removed++ } 'Already gone' { $gone++ } 'Failed' { $failed.Add($c) } } }
        Write-McoItem $(if ($failed.Count) { 'Warn' } else { 'Ok' }) ('{0} cop{1} removed without a message {2} {3} already gone {2} {4} failed' -f $removed, $(if ($removed -eq 1) { 'y' } else { 'ies' }), $dot, $gone, $failed.Count) -Icon Trash
        foreach ($c in ($failed | Select-Object -First 5)) { Write-McoItem Fail "$($c.Mailbox): $($c.Detail)" }
    }
    elseif (-not $plan.Cancel.Count) { Write-McoItem Skip 'No copy to remove in the attendees and the rooms.' }

    # ---- verify -------------------------------------------------------------------------------------------------
    $done = [Collections.Generic.List[object]]::new()
    $cancelled = [Collections.Generic.List[object]]::new()
    foreach ($c in $plan.Remove) { if ($c.Result -eq 'Removed') { $done.Add($c) } }
    foreach ($c in $plan.Cancel) { if ($c.Result -eq 'Cancelled') { $cancelled.Add($c) } }
    if ($Settings.Verify -and ($done.Count -or $cancelled.Count)) {
        Write-McoNextStep 'Verify' 'Search'
        $still = 0; $checked = 0; $all = $done.Count + $cancelled.Count; $confirmed = 0
        foreach ($copy in $done) {
            $exists = Test-McoItemExists -Mailbox $copy.Mailbox -EventId $copy.EventId
            if ($exists -eq $false) { $copy.Verified = 'Yes' }
            elseif ($exists) { $copy.Verified = 'No'; $copy.Result = 'Failed'; $copy.Detail = 'still in the calendar after the request'; $still++ }
            else { $copy.Verified = 'Unknown' }
            $checked++
            Write-McoProgress ($checked / [Math]::Max(1, $all)) ('{0:N0}/{1:N0} copies verified' -f $checked, $all)
        }
        foreach ($copy in $cancelled) {
            $item = Get-McoItem -Mailbox $copy.Mailbox -EventId $copy.EventId
            $copy.Verified = if ($null -eq $item -or $item.IsCancelled) { 'Yes' } else { 'No' }
            if ($copy.Verified -eq 'Yes') { $confirmed++ }
            $checked++
            Write-McoProgress ($checked / [Math]::Max(1, $all)) ('{0:N0}/{1:N0} copies verified' -f $checked, $all)
        }
        Write-McoItem $(if ($still) { 'Warn' } else { 'Ok' }) ('{0} of {1} removed copies verified gone{2}{3}' -f ($done.Count - $still), $done.Count, $(if ($still) { " $dot $still still present" } else { '' }), $(if ($cancelled.Count) { " $dot $confirmed of $($cancelled.Count) organizer meetings verified cancelled" } else { '' })) -Icon Search
    }

    # ---- status of each meeting and of the run (one pass over its copies) ----------------------------------------
    foreach ($meeting in @($plan.Meetings)) {
        $ok = 0; $bad = 0; $isCancelled = $false
        foreach ($c in $meeting.Copies) {
            if ($c.Action -in 'Remove', 'Cancel' -and $c.Result -in 'Removed', 'Cancelled', 'Already gone') { $ok++ }
            if ($c.Result -in 'Failed', 'Not done') { $bad++ }
            if ($c.Result -eq 'Cancelled') { $isCancelled = $true }
        }
        $meeting.Status = if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } elseif ($Action -eq 'Cancel' -and $isCancelled) { 'Cancelled' } else { 'Removed' }
    }
    Update-McoResultCounts $Result
    $Result.CompletedUtc = [datetime]::UtcNow.ToString('o')
    $Result.DurationSeconds = [Math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
    $sel = @($plan.Meetings)
    $failedCount = 0; $warned = $false
    foreach ($m in $sel) { if ($m.Status -eq 'Failed') { $failedCount++ }; if ($m.Status -in 'Partial', 'Failed') { $warned = $true } }
    $Result.Status = if ($sel.Count -and $failedCount -eq $sel.Count) { 'Failed' } elseif ($warned -or $plan.Held.Count) { 'Warning' } else { 'Completed' }
    $Result
}

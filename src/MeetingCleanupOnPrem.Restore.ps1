function Import-McoRestoreSource {
    <# Reads the report of a Remove, Cancel or Transfer run: the copies it removed, with the time of each removal. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string[]]$MeetingId)
    $file = if (Test-Path -LiteralPath $Path -PathType Container) { Get-ChildItem -LiteralPath $Path -Filter '*-Summary.json' -File | Select-Object -First 1 -ExpandProperty FullName } else { $Path }
    if (-not $file -or -not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "No *-Summary.json in ${Path}: give the folder of the report of a Remove, Cancel or Transfer run." }
    $data = [IO.File]::ReadAllText($file) | ConvertFrom-Json -Depth 32
    if ($data.Tool -ne 'Meeting Cleanup On-Prem' -or -not $data.PSObject.Properties['Meetings']) { throw "$file is not a Meeting Cleanup On-Prem report." }
    $ids = @($MeetingId | Where-Object { $_ } | ForEach-Object { ([string]$_).ToUpperInvariant() })
    $meetings = [Collections.Generic.List[object]]::new()
    $runRemoved = [Collections.Generic.List[object]]::new()
    foreach ($source in @($data.Meetings)) {
        $copies = [Collections.Generic.List[object]]::new()
        foreach ($copy in @($source.Copies)) {
            foreach ($name in 'ActionUtc', 'Occurrence', 'SeriesId', 'Verified', 'Detail') { if (-not $copy.PSObject.Properties[$name]) { $copy | Add-Member -NotePropertyName $name -NotePropertyValue '' } }
            $copy | Add-Member -NotePropertyName PreviousResult -NotePropertyValue ([string]$copy.Result) -Force
            $copy | Add-Member -NotePropertyName RemovedUtc -NotePropertyValue $(if ($copy.ActionUtc) { ConvertTo-McoRecoverableTime $copy.ActionUtc } else { $null }) -Force
            $copy | Add-Member -NotePropertyName RestoredUtc -NotePropertyValue '' -Force
            if ($copy.EventId -and [string]$copy.Action -eq 'Remove' -and [string]$copy.Result -eq 'Removed' -and -not $copy.Occurrence) { $runRemoved.Add($copy) }
            $copies.Add($copy)
        }
        if ($ids.Count -and $ids -notcontains ([string]$source.MeetingId).ToUpperInvariant()) { continue }
        if (-not $source.Selected) { continue }
        $notes = [Collections.Generic.List[string]]::new()
        foreach ($n in @($source.Notes)) { if ($n) { $notes.Add([string]$n) } }
        $source.Copies = $copies
        $source.Notes = $notes
        Add-McoMeetingDefaults $source
        $meetings.Add($source)
    }
    $result = [pscustomobject]@{
        Tool = 'Meeting Cleanup On-Prem'; Version = $script:ToolVersion; Action = 'Restore'; Status = 'Completed'; Error = ''
        StartedUtc = [datetime]::UtcNow.ToString('o'); CompletedUtc = ''; DurationSeconds = 0
        Request = $data.Request; Tenant = $data.Tenant; Organization = $data.Organization; AppId = ''; AppName = $data.AppName
        Organizers = @($data.Organizers); Searched = $data.Searched; Meetings = $meetings; Warnings = [Collections.Generic.List[string]]::new()
        Counts = $null; FromReport = $file; SourceAction = [string]$data.Action; RunRemoved = $runRemoved.ToArray()
        SourceRunUtc = [pscustomobject]@{ Start = ConvertTo-McoRecoverableTime $data.StartedUtc; End = ConvertTo-McoRecoverableTime $data.CompletedUtc }
    }
    Update-McoResultCounts $result
    $result
}

function Get-McoRestorePlan {
    <#
        Restorable: the copies removed silently (Result Removed). Not restorable: a meeting cancelled by its
        organizer (the attendees received the cancellation), a meeting transferred to a new organizer, an occurrence
        of a series (rooms mode).
    #>
    param([Parameter(Mandatory)][pscustomobject]$Result)
    $restore = [Collections.Generic.List[object]]::new()
    $cancelled = [Collections.Generic.List[object]]::new()
    $transferred = [Collections.Generic.List[object]]::new()
    $occurrences = [Collections.Generic.List[object]]::new()
    $fast = [MeetingCleanupOnPremNative.Fast]
    foreach ($meeting in $Result.Meetings) {
        if (-not $meeting.Selected) { continue }
        $removed = [Collections.Generic.List[object]]::new()
        $wasCancelled = $false
        foreach ($c in $meeting.Copies) {
            $previous = $fast::Text($c, 'PreviousResult')
            if ($c.EventId -and [string]$c.Action -eq 'Remove' -and ([string]$c.Result -eq 'Removed' -or $previous -eq 'Removed')) { $removed.Add($c) }
            if ([string]$c.Action -eq 'Cancel' -and ([string]$c.Result -eq 'Cancelled' -or $previous -eq 'Cancelled')) { $wasCancelled = $true }
        }
        if (-not $removed.Count) { continue }
        if ($fast::Text($meeting, 'NewMeetingId')) { $transferred.Add($meeting); continue }
        if ($wasCancelled) { $cancelled.Add($meeting); continue }
        foreach ($c in $removed) { if ($fast::Text($c, 'Occurrence')) { $occurrences.Add($c) } else { $restore.Add($c) } }
    }
    $lines = [Collections.Generic.List[string]]::new()
    if ($restore.Count) {
        $mailboxes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $rooms = 0
        foreach ($c in $restore) { [void]$mailboxes.Add([string]$c.Mailbox); if ($c.Role -eq 'Room') { $rooms++ } }
        $lines.Add(('{0} cop{1} to put back in {2} mailbox(es): {3} attendee(s), {4} room(s), from Recoverable Items (retention of deleted items)' -f $restore.Count, $(if ($restore.Count -eq 1) { 'y' } else { 'ies' }), $mailboxes.Count, ($restore.Count - $rooms), $rooms))
        $lines.Add('No message is sent: each copy comes back to its calendar as it was')
    }
    else { $lines.Add('No copy removed silently by this run') }
    if ($cancelled.Count) { $lines.Add(('{0} meeting(s) cancelled by their organizer are not restorable: the attendees received the cancellation' -f $cancelled.Count)) }
    if ($transferred.Count) { $lines.Add(('{0} meeting(s) transferred to a new organizer are not restorable: they are in the calendars again, with the new organizer' -f $transferred.Count)) }
    if ($occurrences.Count) { $lines.Add(('{0} occurrence(s) of series are not restorable (Backup.json lists them)' -f $occurrences.Count)) }
    [pscustomobject]@{ Restore = $restore.ToArray(); Cancelled = $cancelled.ToArray(); Transferred = $transferred.ToArray(); Occurrences = $occurrences.ToArray(); Lines = $lines.ToArray(); Text = $lines -join " $($script:Dot) " }
}

function Invoke-McoRestore {
    <#
        Puts back the copies removed by a run. A copy is restored only when its item is certain: the items with its
        subject removed around the time of the run are exactly the copies of the run that are not back, in the order
        of the removals (Remove leaves 3 seconds between the copies with the same subject in one mailbox). Otherwise
        nothing is restored for that subject in that mailbox: a wrong item is never put back.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Settings, [Parameter(Mandatory)][pscustomobject]$Result)
    $started = [datetime]::UtcNow
    $dot = $script:Dot
    $plan = Get-McoRestorePlan $Result
    $window = [TimeSpan]::FromMinutes([int]$Settings.RestoreWindowMinutes)
    $tolerance = [TimeSpan]::FromMinutes(2)
    $todo = @($plan.Restore)
    $todoSet = [Collections.Generic.HashSet[object]]::new([Collections.Generic.ReferenceEqualityComparer]::Instance)
    foreach ($c in $todo) { [void]$todoSet.Add($c) }
    foreach ($m in $Result.Meetings) {
        foreach ($c in $m.Copies) {
            if ($todoSet.Contains($c)) { $c.Action = 'Restore'; $c.Result = ''; $c.Detail = ''; $c.Verified = ''; $c.HttpStatus = 0 }
            else { $c.Action = ''; if ($c.PreviousResult) { $c.Detail = "not restored: $($c.PreviousResult.ToLowerInvariant()) by the run" }; $c.Result = '' }
        }
    }
    foreach ($m in $plan.Cancelled) { $m.Status = 'Not restorable'; $m.Notes.Add('Cancelled by its organizer: the attendees received the cancellation. The organizer has to send the meeting again.') }
    foreach ($m in $plan.Transferred) { $m.Status = 'Not restorable'; $m.Notes.Add("Transferred to $($m.NewOrganizer): the meeting is in the calendars with its new organizer.") }
    foreach ($c in $plan.Occurrences) { $c.Result = 'Not restorable'; $c.Detail = 'occurrence of a series: see Backup.json' }

    Write-McoNextStep 'Restore' 'Refresh'
    foreach ($line in $plan.Lines) { Write-McoItem Info $line }
    if (-not $todo.Count) { Write-McoItem Skip 'No copy to restore.' }

    # ---- which copies of the run are already back in their calendar (each mailbox read once) -------------------
    $keyOf = { param($c) '{0}|{1}' -f $c.Mailbox, ([string]$c.Subject).Trim().ToLowerInvariant() }
    $keys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($c in $todo) { [void]$keys.Add((& $keyOf $c)) }
    # The copies removed by the run, by mailbox and subject (in the order of the run).
    $membersOf = @{}
    $uidsOf = [ordered]@{}
    foreach ($c in @($Result.RunRemoved)) {
        $k = & $keyOf $c
        if (-not $keys.Contains($k)) { continue }
        $list = $membersOf[$k]
        if (-not $list) { $list = [Collections.Generic.List[object]]::new(); $membersOf[$k] = $list }
        $list.Add($c)
        $mailbox = ([string]$c.Mailbox).ToLowerInvariant()
        if (-not $uidsOf.Contains($mailbox)) { $uidsOf[$mailbox] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
        [void]$uidsOf[$mailbox].Add([string]$c.MeetingId)
    }
    $present = [Collections.Generic.Dictionary[object, object]]::new([Collections.Generic.ReferenceEqualityComparer]::Instance)
    $searchStart = [datetime]::UtcNow.AddDays(-30); $searchEnd = [datetime]::UtcNow.AddDays(400)
    $inCalendar = @{}
    $read = 0
    foreach ($mailbox in @($uidsOf.Keys)) {
        try {
            $set = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($s in (Get-McoCalendarView -Mailbox $mailbox -StartUtc $searchStart -EndUtc $searchEnd)) { if ($s.MeetingId) { [void]$set.Add([string]$s.MeetingId) } }
            $inCalendar[$mailbox] = $set
        }
        catch { $inCalendar[$mailbox] = $null }
        $read++
        Write-McoProgress ($read / [Math]::Max(1, $uidsOf.Count)) ('{0:N0}/{1:N0} calendars checked' -f $read, $uidsOf.Count)
    }
    foreach ($list in $membersOf.Values) {
        foreach ($c in $list) {
            $set = $inCalendar[([string]$c.Mailbox).ToLowerInvariant()]
            $present[$c] = if ($null -eq $set) { $null } else { $set.Contains([string]$c.MeetingId) }
        }
    }
    foreach ($c in $todo) {
        if ($present.ContainsKey($c) -and $present[$c]) { $c.Result = 'Already present'; $c.Detail = 'already in the calendar: nothing to restore' }
        elseif ($present.ContainsKey($c) -and $null -eq $present[$c]) { $c.Result = 'Failed'; $c.Detail = 'mailbox not readable through EWS' }
    }

    # ---- each mailbox and subject: items of Recoverable Items matched to the copies of the run ------------------
    $restored = [Collections.Generic.List[object]]::new()
    $groups = [ordered]@{}
    foreach ($c in $todo) {
        if ($c.Result) { continue }
        $k = & $keyOf $c
        $list = $groups[$k]
        if (-not $list) { $list = [Collections.Generic.List[object]]::new(); $groups[$k] = $list }
        $list.Add($c)
    }
    $groupDone = 0
    foreach ($groupKey in @($groups.Keys)) {
        $groupDone++
        Write-McoProgress ($groupDone / [Math]::Max(1, $groups.Count)) ('{0:N0}/{1:N0} mailboxes and subjects restored' -f $groupDone, $groups.Count)
        $pend = @($groups[$groupKey])
        $mailbox = $pend[0].Mailbox
        $subject = ([string]$pend[0].Subject).Trim()
        $want = @(@($membersOf[$groupKey]) | Where-Object { $_ -and $present.ContainsKey($_) -and $present[$_] -eq $false } | Sort-Object { if ($_.RemovedUtc) { $_.RemovedUtc } else { [datetime]::MinValue } })
        $known = @($want | Where-Object RemovedUtc)
        $lo = if ($known.Count) { ($known | ForEach-Object RemovedUtc | Measure-Object -Minimum).Minimum } else { $Result.SourceRunUtc.Start }
        $hi = if ($known.Count) { ($known | ForEach-Object RemovedUtc | Measure-Object -Maximum).Maximum } else { $Result.SourceRunUtc.End }
        if (-not $lo -or -not $hi) { foreach ($c in $pend) { $c.Result = 'Failed'; $c.Detail = 'time of the removal unknown in the report' }; continue }
        try { $candidates = @(Get-McoPurgedItems -Mailbox $mailbox -StartUtc ($lo - $window) -EndUtc ($hi + $window) -Mode $Settings.RestoreMode) }
        catch { foreach ($c in $pend) { $c.Result = 'Failed'; $c.Detail = "Get-RecoverableItems: $($_.Exception.Message)" }; continue }
        $found = @($candidates | Where-Object { ([string]$_.Subject).Trim() -eq $subject -and (-not $_.LastModifiedUtc -or ($_.LastModifiedUtc -ge ($lo - $tolerance) -and $_.LastModifiedUtc -le ($hi + $tolerance))) } | Sort-Object LastModifiedUtc, EntryId)
        if (-not $found.Count) {
            foreach ($c in $pend) { $c.Result = 'Not found'; $c.Detail = 'not in Recoverable Items: retention over, restored before, or removed by someone else' }
            continue
        }
        $distinct = { param($values) $list = @($values); @($list | Select-Object -Unique).Count -eq $list.Count }
        $ordered = $known.Count -eq $want.Count -and (& $distinct @($found | ForEach-Object { if ($_.LastModifiedUtc) { $_.LastModifiedUtc.ToString('yyyyMMddHHmmss') } })) -and (& $distinct @($want | ForEach-Object { $_.RemovedUtc.ToString('yyyyMMddHHmmss') }))
        if ($found.Count -ne $want.Count -or -not ($ordered -or $want.Count -eq 1)) {
            foreach ($c in $pend) {
                $c.Result = 'Failed'
                $c.Detail = "ambiguous, nothing restored ($($found.Count) item(s) in Recoverable Items for $($want.Count) copies removed by the run). By hand: Get-RecoverableItems -Identity $mailbox -FilterItemType IPM.Appointment -SubjectContains '$($subject.Replace("'", "''"))', then Restore-RecoverableItems -Identity $mailbox -EntryID <item>"
            }
            continue
        }
        for ($j = 0; $j -lt $want.Count; $j++) {
            $c = $want[$j]
            if ($pend -notcontains $c) { continue }
            $r = Invoke-McoRestoreItem -Mailbox $c.Mailbox -EntryId $found[$j].EntryId -Mode $Settings.RestoreMode
            $c.RestoredUtc = [datetime]::UtcNow.ToString('o')
            if ($r.Ok) { $c.Result = 'Restored'; $c.Detail = "back in $(if ($r.Folder) { $r.Folder } else { 'its calendar' }) (from $($found[$j].Folder))"; $restored.Add($c) }
            else { $c.Result = 'Failed'; $c.Detail = "Restore-RecoverableItems: $($r.Error)" }
        }
    }
    if ($todo.Count) {
        $back = 0; $already = 0; $notFound = 0; $failed = 0
        foreach ($c in $todo) { switch ($c.Result) { 'Restored' { $back++ } 'Already present' { $already++ } 'Not found' { $notFound++ } 'Failed' { $failed++ } } }
        Write-McoItem $(if ($notFound + $failed) { 'Warn' } else { 'Ok' }) ('{0} cop{1} put back {2} {3} already present {2} {4} not found {2} {5} failed' -f $back, $(if ($back -eq 1) { 'y' } else { 'ies' }), $dot, $already, $notFound, $failed) -Icon Refresh
    }

    # ---- verify: each copy back in its calendar, with the state of its answer (each mailbox read once) ------------
    if ($restored.Count) {
        Write-McoNextStep 'Verify' 'Search'
        Start-Sleep -Seconds 3
        $back = 0
        $byMailbox = [ordered]@{}
        foreach ($c in $restored) {
            $mailbox = ([string]$c.Mailbox).ToLowerInvariant()
            if (-not $byMailbox.Contains($mailbox)) { $byMailbox[$mailbox] = [Collections.Generic.List[object]]::new() }
            $byMailbox[$mailbox].Add($c)
        }
        $read = 0
        foreach ($mailbox in @($byMailbox.Keys)) {
            $copies = $byMailbox[$mailbox]
            $hitOf = @{}
            try {
                $uids = [string[]]@(foreach ($c in $copies) { [string]$c.MeetingId })
                foreach ($hit in @(Get-McoMailboxEvents -Mailbox $mailbox -StartUtc $searchStart -EndUtc $searchEnd -Uid $uids)) { if ($hit -and -not $hitOf.ContainsKey($hit.MeetingId)) { $hitOf[$hit.MeetingId] = $hit } }
            }
            catch { }
            foreach ($c in $copies) {
                $hit = $hitOf[[string]$c.MeetingId]
                if ($hit) { $back++; $c.Verified = 'Yes'; $c.EventId = $hit.EventId; $c.Detail = "$($c.Detail); answer $($hit.Response)$(if ($hit.IsCancelled) { ', cancelled' })" }
                else { $c.Verified = 'No'; $c.Detail = "$($c.Detail); not seen in the calendar yet" }
            }
            $read++
            Write-McoProgress ($read / [Math]::Max(1, $byMailbox.Count)) ('{0:N0}/{1:N0} calendars verified' -f $read, $byMailbox.Count)
        }
        Write-McoItem $(if ($back -lt $restored.Count) { 'Warn' } else { 'Ok' }) ('{0} of {1} verified back in their calendar' -f $back, $restored.Count) -Icon Search
    }

    foreach ($m in $Result.Meetings) {
        if ($m.Status -eq 'Not restorable') { continue }
        $acted = 0; $ok = 0
        foreach ($c in $m.Copies) { if ($c.Action -eq 'Restore') { $acted++; if ($c.Result -in 'Restored', 'Already present') { $ok++ } } }
        $m.Status = if (-not $acted) { 'Nothing to do' } elseif ($ok -eq $acted) { 'Restored' } elseif ($ok) { 'Partial' } else { 'Failed' }
    }
    Update-McoResultCounts $Result
    $all = @($Result.Meetings)
    $failedCount = 0; $warned = $false
    foreach ($m in $all) { if ($m.Status -eq 'Failed') { $failedCount++ }; if ($m.Status -in 'Partial', 'Failed', 'Not restorable') { $warned = $true } }
    $Result.Status = if ($all.Count -and $failedCount -eq $all.Count) { 'Failed' } elseif ($warned -or $Result.Warnings.Count) { 'Warning' } else { 'Completed' }
    $Result.CompletedUtc = [datetime]::UtcNow.ToString('o')
    $Result.DurationSeconds = [Math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
    $Result
}

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
            foreach ($name in 'ActionUtc', 'Occurrence', 'OccurrenceStart', 'OccurrenceKey', 'SeriesId', 'Verified', 'Detail') { if (-not $copy.PSObject.Properties[$name]) { $copy | Add-Member -NotePropertyName $name -NotePropertyValue '' } }
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
        of a series (rooms mode, -SeriesScope Occurrences).
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

# SIG # Begin signature block
# MIImIQYJKoZIhvcNAQcCoIImEjCCJg4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBj0Xs9B5EFas1T
# ab9FNqVEBTV6SrUVQE1hjlBEQV5s0aCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYCDCCGAQCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC8Ve2GCbqx
# GkDqY2sVVq15eCubI8MTUzvhpgMtYqDgJTANBgkqhkiG9w0BAQEFAASCAgChKnzF
# hRM3sA384QbRd7Tx+ju0q7LdTc33S/aqcBTECN3r1V5qbYm8ez21P/i6jIlrn1uI
# gA8e6O7AgOCqhmHwNoTVF5fCOw43xhfMup5fzsr+DCpvfXtCGRHvxaADhKUIVLfH
# RxpHQYiz8GwR37fW4ZXQ+O7/36oVB9fdybQKNsn3S3I4Jj9AB8Fd9PXtaJsCKYOJ
# SzlI+stUW6Zxww8h9ld1Tt2J4L2BAeau2OFRjBUvM7sIstNDz4FgFBLEyRIhPxyB
# wgWPs4FpSRPbTKODqYzd58Eb/hZUvJHI1wf7VGLVk6H0k/NW8faYUX6G+0yV2OnS
# H3YuBlo7eN4ZawN90a7tOFEq8iEk4cDutzYvfQI176gP3Krd7IY2ZzQrAMvROk/t
# 7XjJ8X0CyZm/LFzgsAzWEIFEAkLIBFZXRJ4zqezw74CVRcxOPHDZa4KAs+q8cTjX
# Ci4Y+2hJP+I8I6pv1ng8JwZq46hKLk/ggr8nNHNPZPelxf743T0kZSv4Kysof1JH
# vdeZ4NNMMhweGi9w1VntuSI+21fs6oNV86/y7lYf49ctfT07lhBLGXTMdGbGWafT
# mZMv2K6EcCJE9aqmE05RO8LbGG+sYrzX3p/OlLZVmoSorJmlzOEZ+C51bYr8m0cu
# NFCVTX+UREN0cETkRspkEqwmNb16vvXjyO3+I6GCFMUwghTBBgorBgEEAYI3AwMB
# MYIUsTCCFK0GCSqGSIb3DQEHAqCCFJ4wghSaAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIArIiFiTYVPRIXcIi5mHcI5prcDe2h+pWrOFom+QyDGlAggIuQhqLVV8
# txgPMjAyNjExMDUxNzIyNDBaMAMCAQECCC3ybLXyHwGGoIIRpzCCBPMwggLboAMC
# AQICEFRD+De9Wu4YAYRyuFzajuMwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhT
# U0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNzdWluZyBS
# U0EgQ0EgMjAeFw0yNjA2MTIxNzIyNDZaFw0zNjA2MDkxNzIyNDVaMG4xCzAJBgNV
# BAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UE
# CgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1waW5nIFVuaXQg
# MjAyNiBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABOgli7hmEy5IkjjzCzbJ
# DgbYV7CIUUgOOn3jkML0i3A2jYer8VMwShMdSnRQURbOzh65x/ytFQOjGw0CBaVI
# ND2jggFSMIIBTjAfBgNVHSMEGDAWgBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDBNBggr
# BgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5jb20vU1NM
# LmNvbS10aW1lU3RhbXBpbmctSS1SMi5jZXIwUQYDVR0gBEowSDA8BgwrBgEEAYKp
# MAEDBgEwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0
# b3J5MAgGBmeBDAEEAjAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5
# MDegNaAzhjFodHRwOi8vY3Jscy5zc2wuY29tL1NTTC5jb20tdGltZVN0YW1waW5n
# LUktUjIuY3JsMB0GA1UdDgQWBBQnrANWuihquhc7gz1Af3moJOzDPDAOBgNVHQ8B
# Af8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE92bs78A57CGU1wZ/oY5rgV4Mwd
# dn22MU/bmHumb4b6W2VinMg0IGnBE5D8PJOsnqp2AmMhWdOt5TDxKgIQA9pzBute
# GYez7R1vQ1qn7IIr8Wz2AXjfre1MYJGE93saW1jHTxKgbVlchzry3CrzMYCVnyYw
# DcFZMFQ35Mv4+3H9vPoopjfDynO0v7vPy3p0UZ1dltmH53EZ5M/Qn7OZ0Q82mYLo
# h666dhscIJzUEh+BgYxzpdz8y7J1WxqGSouGRI4F7qbOJhhjiu/hgdhwJUJ/jkve
# OHB6pZB9IBJbdlns6AjUaRC6YfamsbgZ40hKl1ln4BOshgcDtLsSjaqVlUx5gMK+
# krwT6H8bm1yLhJ/Yr0eAizrgMVEqyeNWQYP6qhtKCYxehpESm8au3MlkpBGWCYZi
# goNh4OX1x2oDsYsG5p+ST2db4Vgq25N7/3U8doL0yufnqwuR/f603JY1CO++sndq
# 34Y35Kb8gLG7V7OA+5658PIvcKK+fn3bUUiQVcasy5spBX7aC9FKdkPvrciglHa4
# m3XJ6weBykt7PrdwI3XGwNOgwi/Bon0IvXy9MkQUTZQOKOgT/kq++eTGgjY1EWGQ
# 7w8sSRwQU+3Ko2SMLIkfLPuvOtJ6BWwQbXYHtiMOwjwv9UN9SNu4VaJUuwdMmqN7
# /F9jLVst5rGZV03BMIIF3TCCA8WgAwIBAgIIeyyb0xaAMpkwDQYJKoZIhvcNAQEL
# BQAwfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3Vz
# dG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20g
# Um9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0EwHhcNMTYwMjEyMTczOTM5
# WhcNNDEwMjEyMTczOTM5WjB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAPkP3aMrfcvQKv7sZ4Wm5y4b
# unfh4/WvpOz6Sl2RxFdHaxh3a3by/ZPkPQ/CFp4LZsNWlJ4Xg4XOVu/yFv0AYvUi
# CVToZRdOQbngT0aXqhvIuG5iXmmxX9sqAn78bMrzQdjt0Oj8P2FI7bADFB0QDksZ
# 4LtO7IZl/zbzXmcCC52GVWH9ejjt/uIZALdvoVBidXQ8oPrIJZK0bnoix/geoeOy
# 3ZExqysdBP+lSgQ36YWkMyv94tZVNHwZpEpox7Ko07fKoZOI68GXvIz5HdkihCR0
# xwQ9aqkpk8zruFvh/l8lqjRYyMEjVJ0bmBHDOJx+PYZspQ9AhnwC9FwCTyjLrnGf
# DzrIM/4RJTXq/LrFYD3ZfBjVsqnTdXgDciLKOsMf7yzlLqn6niy2UUb9rwPW6mBo
# 6oUWNmuF6R7As93EJNyAKoFBbZQ+yODJgUEAnl6/f8UImKIYLEJAs/lvOCdLToD0
# PYFH4Ih86hzOtXVcUS4cK38acijnALXRdMbX5J+tB5O2UzU1/Dfkw/ZdFr4hc96S
# CvigY2q8lpJqPvi8ZVWb3vUNiSYE/CUapiVpy8JtynziWV+XrOvvLsi81xtZPCvM
# 8hnIk2snYxnP/Okm+Mpxm3+T/jRnhE6Z6/yzeAkzcLpmpnbtG3PrGqUNxCITIJRW
# Ck4sbE6x/c+cCbqiM+2HAgMBAAGjYzBhMB0GA1UdDgQWBBTdBAkHovV6fVJTEpKV
# 7jiAJQ2mWTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMS
# kpXuOIAlDaZZMA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG9w0BAQsFAAOCAgEAIBgR
# lCn7Jp0cHh5wYfGVcpNxJK1ok1iOMq8bs3AD/CUrdIWQPXhq9LmLpZc7tRiRux6n
# +UBbkflVma8eEdBcHadm47GUBwwyOabqG7B52B2ccETjit3E+ZUfijhDPwGFpUen
# PUayvOUiaPd7nNgsPgohyC0zrL/FgZkxdMF1ccW+sfAjRfSda/wZY52jvATGGAsl
# u1OJD7OAUN5F7kR/q5R4ZJjT9ijdh9hwZXT7DrkT66cPYakylszeu+1jTBi7qUD3
# oFRuIIhxdRjqerQ0cuAjJ3dctpDqhiVAq+8zD8ufgr6iIPv2tS0a5sKFsXQP+8hl
# AqRSAUfdSSLBv9jra6x+3uxjMxW3IwiPxg+NQVrdjsW5j+VFP3jbutIbQLH+cU0/
# 4IGiul607BXgk90IH37hVZkLId6Tngr75qNJvTYw/ud3sqB1l7UtgYgXZSD32pAA
# n8lSzDLKNXz1PQ/YK9f1JmzJBjSWFupwWRoyeXkLtoh/D1JIPb9s2KJELtFOt3JY
# 04kTlf5Eq/jXixtunLwsoFvVagCvXzfh1foQC5ichucmj87w7G6KVwuA406ywKBj
# YZC6VWg3dGq2ktufoYYitmUnDuy2n0Jg5GfCtdpBC8TTi2EbvPofkSvXRAdeuims
# 2cXp71NIWuuA8ShYIc2wBlX7Jz9TkHCpBB5XJ7kwggbLMIIEs6ADAgECAhBSnbvC
# uImH9YOA4LDrogY2MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTI2MDYxMjE2NDIxNVoXDTQxMDIxMjE3MzkzOVowcjELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgMjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJ8v
# q+ZzTZC8MUabNgd5WlBjVPcu71oiedNd8n9Il8+uW0HcGvW44/Jf0pHl6hxqn0od
# haGnvvNBPJgxH/d6CMI7Mtgl9nF5isarRfByRmAeH+skYShA+dMJGzsrC9NiF3/v
# xGZ2uOI334YVc+DlP5OdsUPz5yIP7hYXNlpPknPwh6gK/pWFZD8+y8uUagfX9USh
# q/SQt+1tS4FXRwudHTSHbGt0Fz4ZbWGbm0d56ROAW34tI6FFuZ+TtGssHjeOUf2F
# S8QNGBpadTh1X1Z7Kl6vG4Z8NS3OJIyelzoF5TpcP2vM3ue0723bmZd3oXsE5rzt
# AJMv56jlg0vSDz8XcXRur6awW0rh5cTQDQQGx04hYJH2Zn9HWWJU/wJ2SKBcQKtB
# yCz6JX+DoFoBE++kolkpZtl2ESF8FY6XqhtKZb18ESEAc0ohCY+0HkBZwrx8eIED
# CDbzd/shv9a8YRG2oxr3tLUz95V+Dzzh1QBNECaeGczyNWGCZIVUa4lVcVUVOBbI
# OAT8D+iPmC5h29bWeES8DhTS4f4CWA4D5QoQqMTqIBcV4MD2Waw8hEWAdGnjnsVt
# DimnDeEUbvXKhdiirvwG+QC+3mAy1RlD6xBMmGyDAKVSqqXm3NP4w0J4gvt65Grk
# QI4FXPJtYqfwkIBtQ8OGfbpVUQGPVUemWtui7E21AgMBAAGjggFRMIIBTTASBgNV
# HRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMSkpXuOIAlDaZZ
# ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3NsLmNv
# bS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNlcjA/BgNVHSAEODA2MDQGBFUd
# IAAwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0b3J5
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9j
# cmxzLnNzbC5jb20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jcmwwHQYDVR0O
# BBYEFHCaZcFDgKP86bFkiW7xMJwJIcl8MA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG
# 9w0BAQsFAAOCAgEAZRRbzTHpgOogYgWKOeiCkI1oVrkGmCrqOKqX9CiTREQLdYM/
# h7ANki98QP1fhAPEkTzT/Ur19IrPuZg7KaWL6SsmBqmD9+mBVtJmxdRdGbV8bpgl
# FDv2VeZgTSg3+UOD2O0G9HlgHqVHNKzSa/zL5+WYl3yT8Q29sZnIPGkQ+udv56/j
# q5NlAT9BON4iXSOCxe1bNvUd0hkS2GTsTNeQyOEHm7CkZdCg5X+DpmBginO7pk8Y
# m5GbTdQtSloJ3/Y2nVPycJ8kmmBAYMev7/hVQukhjjc4mO03LyPTFfqmMR1D5Tpu
# gGkD9nUuG41l55vOpXQm1JMbTgArpkn0/e9O+yTYlWuz2dSAZHf9jxNQAvILO1Pv
# 2JNQTfV0TyIteZP7gNcRfaHnnuSho8ohipKiE+1vEJIFE+c3gopxSpHCmn2THETk
# lWKV7ZrUKhhkdoOLj2KidCtYZrOoCV6x7f9+KfDPT61fKDVntGZdRKwzLV53VIN6
# oSS+KQN2VJaO8OZTOZpfgO4o6/9xpNr+EEaA3hNkoZk13R4LaqPC/309dfZmHBE8
# nBNjWBkV7AZdzdK8w7B0a/D9066a6oUa19OIktD+3kPpQ8YEt8oRp9zs1hoQ7fRn
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJV
# MIICUQIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDUxNzIyNDBaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCCi9O36iRM4BQ8i
# 4dqgAKSm1l1qnPqS2J6G/b+UfWkASTCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEYw
# RAIgKAcEN7Hzz5cLle78Li3F4a2TbLv0NF0MN1AeIv16LxICIAehcazkZnyLOgu8
# a6OgeTXpWnALC9b3mJOOgZVNJYq9
# SIG # End signature block

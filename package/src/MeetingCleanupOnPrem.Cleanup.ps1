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

# SIG # Begin signature block
# MIImIwYJKoZIhvcNAQcCoIImFDCCJhACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDXTUlmC3Y4mwxi
# TCglepbQI9+x3zSceeEvKjr5GspbN6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYCjCCGAYCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBciv8+aPAu
# RBfdqTeTYb7NuXjwv+ryxQ2poQJCd47fCDANBgkqhkiG9w0BAQEFAASCAgBTGBa1
# saz/pE2EXMvt7AOxSRKVFKY4V4aGFXqj5gZojaARSCIk7tXyBVL0vapmwQS4vNEb
# i02SF0AqtQtpd96XubVaifz1NZUwqy/bEAUkq9sgDYr4pE2RbNXiaz7UVSt6w0AM
# K2TttBe/mNRdlDJVMTbEXsHrZvMCTQ214jZ8UP8eYTGLy8tYh5M//pxlYWb2N3jH
# CTPnDVP/8nS2nU/2BZKDO+/ETbcFwUzuGUPA36a0zJL6ur17CAKKXVMdahlUYPqW
# JkJeU2RwLkZeyI8BTqtdnHSz349iA2wctNEAVPQqwS1PrcWY8YhkkL+ayDLxRxrA
# 9D6espvKHxsYpoQgH5XN+JHwmxr3ZOOcCC9LQS1HVXZ+DYNRmR6cegb0nOtBGX/L
# 40C/3vbn2RTcCaxTWtz0KrxwvqX5XdU78eC1f8OrMDJOim0gfGrlmJjtDrcl4QTv
# LCWUsHqEJMKC15Irr77yCMfZawNhnA2c4+KXIjhguLsnNqkyI0+JQn0Dw0tmNEuz
# FDUphQDPT7avF8jSs/Wl3Z/1L4MiFbcUoFKmHn2PF4z86YiqHE96BQgUayP5YA1Q
# ZavNYuOoaDAXa9guuT1ULDHSNxzq+ZvIbWegw5vwXV7j0NlRqTA5SOjs3XBS0BO5
# 5eMR+GxpV7+5OA4kNLLMSnbCTfj3aFNwwb7tNaGCFMcwghTDBgorBgEEAYI3AwMB
# MYIUszCCFK8GCSqGSIb3DQEHAqCCFKAwghScAgEDMQ0wCwYJYIZIAWUDBAIBMIGC
# BgsqhkiG9w0BCRABBKBzBHEwbwIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIGH+vN9IMVa5AWrM+HKBN1OMtCLbXHtzgKQDs+4JoVZFAggDoU16deEY
# VBgPMjAyNjExMDUxNzIyMzZaMAMCAQECCQD2UBTCRa0aF6CCEacwggTzMIIC26AD
# AgECAhBUQ/g3vVruGAGEcrhc2o7jMA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYT
# AlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwI
# U1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0YW1waW5nIElzc3Vpbmcg
# UlNBIENBIDIwHhcNMjYwNjEyMTcyMjQ2WhcNMzYwNjA5MTcyMjQ1WjBuMQswCQYD
# VQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAPBgNV
# BAoMCFNTTCBDb3JwMSowKAYDVQQDDCFTU0wuY29tIFRpbWVzdGFtcGluZyBVbml0
# IDIwMjYgRTEwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAAToJYu4ZhMuSJI48ws2
# yQ4G2FewiFFIDjp945DC9ItwNo2Hq/FTMEoTHUp0UFEWzs4eucf8rRUDoxsNAgWl
# SDQ9o4IBUjCCAU4wHwYDVR0jBBgwFoAUcJplwUOAo/zpsWSJbvEwnAkhyXwwTQYI
# KwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2VydC5zc2wuY29tL1NT
# TC5jb20tdGltZVN0YW1waW5nLUktUjIuY2VyMFEGA1UdIARKMEgwPAYMKwYBBAGC
# qTABAwYBMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNzbC5jb20vcmVwb3Np
# dG9yeTAIBgZngQwBBAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDsw
# OTA3oDWgM4YxaHR0cDovL2NybHMuc3NsLmNvbS9TU0wuY29tLXRpbWVTdGFtcGlu
# Zy1JLVIyLmNybDAdBgNVHQ4EFgQUJ6wDVrooaroXO4M9QH95qCTswzwwDgYDVR0P
# AQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBPdm7O/AOewhlNcGf6GOa4FeDM
# HXZ9tjFP25h7pm+G+ltlYpzINCBpwROQ/DyTrJ6qdgJjIVnTreUw8SoCEAPacwbr
# XhmHs+0db0Nap+yCK/Fs9gF4363tTGCRhPd7GltYx08SoG1ZXIc68twq8zGAlZ8m
# MA3BWTBUN+TL+Ptx/bz6KKY3w8pztL+7z8t6dFGdXZbZh+dxGeTP0J+zmdEPNpmC
# 6IeuunYbHCCc1BIfgYGMc6Xc/MuydVsahkqLhkSOBe6mziYYY4rv4YHYcCVCf45L
# 3jhweqWQfSASW3ZZ7OgI1GkQumH2prG4GeNISpdZZ+ATrIYHA7S7Eo2qlZVMeYDC
# vpK8E+h/G5tci4Sf2K9HgIs64DFRKsnjVkGD+qobSgmMXoaREpvGrtzJZKQRlgmG
# YoKDYeDl9cdqA7GLBuafkk9nW+FYKtuTe/91PHaC9Mrn56sLkf3+tNyWNQjvvrJ3
# at+GN+Sm/ICxu1ezgPueufDyL3Civn5921FIkFXGrMubKQV+2gvRSnZD763IoJR2
# uJt1yesHgcpLez63cCN1xsDToMIvwaJ9CL18vTJEFE2UDijoE/5KvvnkxoI2NRFh
# kO8PLEkcEFPtyqNkjCyJHyz7rzrSegVsEG12B7YjDsI8L/VDfUjbuFWiVLsHTJqj
# e/xfYy1bLeaxmVdNwTCCBd0wggPFoAMCAQICCHssm9MWgDKZMA0GCSqGSIb3DQEB
# CwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0aW9uMTEwLwYDVQQDDChTU0wuY29t
# IFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUlNBMB4XDTE2MDIxMjE3Mzkz
# OVoXDTQxMDIxMjE3MzkzOVowfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFz
# MRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAv
# BgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0Ew
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQD5D92jK33L0Cr+7GeFpucu
# G7p34eP1r6Ts+kpdkcRXR2sYd2t28v2T5D0PwhaeC2bDVpSeF4OFzlbv8hb9AGL1
# IglU6GUXTkG54E9Gl6obyLhuYl5psV/bKgJ+/GzK80HY7dDo/D9hSO2wAxQdEA5L
# GeC7TuyGZf82815nAgudhlVh/Xo47f7iGQC3b6FQYnV0PKD6yCWStG56Isf4HqHj
# st2RMasrHQT/pUoEN+mFpDMr/eLWVTR8GaRKaMeyqNO3yqGTiOvBl7yM+R3ZIoQk
# dMcEPWqpKZPM67hb4f5fJao0WMjBI1SdG5gRwzicfj2GbKUPQIZ8AvRcAk8oy65x
# nw86yDP+ESU16vy6xWA92XwY1bKp03V4A3IiyjrDH+8s5S6p+p4stlFG/a8D1upg
# aOqFFjZrhekewLPdxCTcgCqBQW2UPsjgyYFBAJ5ev3/FCJiiGCxCQLP5bzgnS06A
# 9D2BR+CIfOoczrV1XFEuHCt/GnIo5wC10XTG1+SfrQeTtlM1Nfw35MP2XRa+IXPe
# kgr4oGNqvJaSaj74vGVVm971DYkmBPwlGqYlacvCbcp84llfl6zr7y7IvNcbWTwr
# zPIZyJNrJ2MZz/zpJvjKcZt/k/40Z4ROmev8s3gJM3C6ZqZ27Rtz6xqlDcQiEyCU
# VgpOLGxOsf3PnAm6ojPthwIDAQABo2MwYTAdBgNVHQ4EFgQU3QQJB6L1en1SUxKS
# le44gCUNplkwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTdBAkHovV6fVJT
# EpKV7jiAJQ2mWTAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZIhvcNAQELBQADggIBACAY
# EZQp+yadHB4ecGHxlXKTcSStaJNYjjKvG7NwA/wlK3SFkD14avS5i6WXO7UYkbse
# p/lAW5H5VZmvHhHQXB2nZuOxlAcMMjmm6huwedgdnHBE44rdxPmVH4o4Qz8BhaVH
# pz1GsrzlImj3e5zYLD4KIcgtM6y/xYGZMXTBdXHFvrHwI0X0nWv8GWOdo7wExhgL
# JbtTiQ+zgFDeRe5Ef6uUeGSY0/Yo3YfYcGV0+w65E+unD2GpMpbM3rvtY0wYu6lA
# 96BUbiCIcXUY6nq0NHLgIyd3XLaQ6oYlQKvvMw/Ln4K+oiD79rUtGubChbF0D/vI
# ZQKkUgFH3Ukiwb/Y62usft7sYzMVtyMIj8YPjUFa3Y7FuY/lRT9427rSG0Cx/nFN
# P+CBorpetOwV4JPdCB9+4VWZCyHek54K++ajSb02MP7nd7KgdZe1LYGIF2Ug99qQ
# AJ/JUswyyjV89T0P2CvX9SZsyQY0lhbqcFkaMnl5C7aIfw9SSD2/bNiiRC7RTrdy
# WNOJE5X+RKv414sbbpy8LKBb1WoAr1834dX6EAuYnIbnJo/O8OxuilcLgONOssCg
# Y2GQulVoN3RqtpLbn6GGIrZlJw7stp9CYORnwrXaQQvE04thG7z6H5Er10QHXrop
# rNnF6e9TSFrrgPEoWCHNsAZV+yc/U5BwqQQeVye5MIIGyzCCBLOgAwIBAgIQUp27
# wriJh/WDgOCw66IGNjANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwG
# A1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jw
# b3JhdGlvbjExMC8GA1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0
# aG9yaXR5IFJTQTAeFw0yNjA2MTIxNjQyMTVaFw00MTAyMTIxNzM5MzlaMHIxCzAJ
# BgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8G
# A1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0YW1waW5nIElz
# c3VpbmcgUlNBIENBIDIwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCf
# L6vmc02QvDFGmzYHeVpQY1T3Lu9aInnTXfJ/SJfPrltB3Br1uOPyX9KR5eocap9K
# HYWhp77zQTyYMR/3egjCOzLYJfZxeYrGq0XwckZgHh/rJGEoQPnTCRs7KwvTYhd/
# 78RmdrjiN9+GFXPg5T+TnbFD8+ciD+4WFzZaT5Jz8IeoCv6VhWQ/PsvLlGoH1/VE
# oav0kLftbUuBV0cLnR00h2xrdBc+GW1hm5tHeekTgFt+LSOhRbmfk7RrLB43jlH9
# hUvEDRgaWnU4dV9WeyperxuGfDUtziSMnpc6BeU6XD9rzN7ntO9t25mXd6F7BOa8
# 7QCTL+eo5YNL0g8/F3F0bq+msFtK4eXE0A0EBsdOIWCR9mZ/R1liVP8CdkigXECr
# Qcgs+iV/g6BaARPvpKJZKWbZdhEhfBWOl6obSmW9fBEhAHNKIQmPtB5AWcK8fHiB
# Awg283f7Ib/WvGERtqMa97S1M/eVfg884dUATRAmnhnM8jVhgmSFVGuJVXFVFTgW
# yDgE/A/oj5guYdvW1nhEvA4U0uH+AlgOA+UKEKjE6iAXFeDA9lmsPIRFgHRp457F
# bQ4ppw3hFG71yoXYoq78BvkAvt5gMtUZQ+sQTJhsgwClUqql5tzT+MNCeIL7euRq
# 5ECOBVzybWKn8JCAbUPDhn26VVEBj1VHplrbouxNtQIDAQABo4IBUTCCAU0wEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2m
# WTBNBggrBgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5j
# b20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jZXIwPwYDVR0gBDgwNjA0BgRV
# HSAAMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNzbC5jb20vcmVwb3NpdG9y
# eTATBgNVHSUEDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5MDegNaAzhjFodHRwOi8v
# Y3Jscy5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY3JsMB0GA1Ud
# DgQWBBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZI
# hvcNAQELBQADggIBAGUUW80x6YDqIGIFijnogpCNaFa5Bpgq6jiql/Qok0REC3WD
# P4ewDZIvfED9X4QDxJE80/1K9fSKz7mYOymli+krJgapg/fpgVbSZsXUXRm1fG6Y
# JRQ79lXmYE0oN/lDg9jtBvR5YB6lRzSs0mv8y+flmJd8k/ENvbGZyDxpEPrnb+ev
# 46uTZQE/QTjeIl0jgsXtWzb1HdIZEthk7EzXkMjhB5uwpGXQoOV/g6ZgYIpzu6ZP
# GJuRm03ULUpaCd/2Np1T8nCfJJpgQGDHr+/4VULpIY43OJjtNy8j0xX6pjEdQ+U6
# boBpA/Z1LhuNZeebzqV0JtSTG04AK6ZJ9P3vTvsk2JVrs9nUgGR3/Y8TUALyCztT
# 79iTUE31dE8iLXmT+4DXEX2h557koaPKIYqSohPtbxCSBRPnN4KKcUqRwpp9kxxE
# 5JVile2a1CoYZHaDi49ionQrWGazqAlese3/finwz0+tXyg1Z7RmXUSsMy1ed1SD
# eqEkvikDdlSWjvDmUzmaX4DuKOv/caTa/hBGgN4TZKGZNd0eC2qjwv99PXX2ZhwR
# PJwTY1gZFewGXc3SvMOwdGvw/dOumuqFGtfTiJLQ/t5D6UPGBLfKEafc7NYaEO30
# Z/bHWp7T+p2EUBU6RfHMxnPhRq3rGwB/NfT7+bwYQpbk6PD8dj6Qu9XDNI87MYIC
# VjCCAlICAQEwgYYwcjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO
# 4zALBglghkgBZQMEAgGgggFgMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAc
# BgkqhkiG9w0BCQUxDxcNMjYxMTA1MTcyMjM2WjAoBgkqhkiG9w0BCTQxGzAZMAsG
# CWCGSAFlAwQCAaEKBggqhkjOPQQDAjAvBgkqhkiG9w0BCQQxIgQgGTbI4Y1M5BJ/
# dsS3DJWhL9ocJQerYShStcWLiIuswbYwgcgGCyqGSIb3DQEJEAIvMYG4MIG1MIGy
# MIGvBCBXT+oXqvpefGormwjJmLFzXkULOHPKROaiTKCTDmPzsDCBijB2pHQwcjEL
# MAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREw
# DwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcg
# SXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO4zAKBggqhkjOPQQDAgRH
# MEUCIQDUTl1FSaDvPdBm+1+WoVmM1v2K+eD98DiLHvBEH/WJCwIgFZxFm8keT4Cg
# 1UZ8nv64VU/EmkNT4iiT+rImP7G8muU=
# SIG # End signature block

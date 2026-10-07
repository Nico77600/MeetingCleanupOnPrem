function Resolve-McoNewOrganizer {
    <# The new organizer: a mailbox of the organization whose calendar EWS can open. #>
    param([Parameter(Mandatory)][string]$Address)
    $address = $Address.Trim().ToLowerInvariant()
    if ($address -notmatch $script:SmtpPattern) { throw "New organizer '$Address' is not an SMTP address." }
    $probe = Invoke-McoEws -Operation 'GetFolder' -Body (New-McoGetFolderBody -Distinguished 'calendar' -Mailbox $address) -Mailbox $address
    if ($probe.HttpStatus -ne 200 -or $probe.ResponseClass -ne 'Success') { throw "New organizer $Address is not reachable through EWS: $(Get-McoEwsFailureText $probe)" }
    $recipient = Resolve-McoExchangeRecipient -Identity $address
    $primary = if ($recipient -and -not $recipient.Error -and $recipient.PrimaryAddress) { $recipient.PrimaryAddress } else { $address }
    [pscustomobject]@{ Address = $primary; Input = $address; Name = $(if ($recipient -and -not $recipient.Error) { $recipient.DisplayName } else { '' }); UserId = '' }
}

function Get-McoTransferPlan {
    <#
        Exchange Server has no supported organizer transfer: each meeting is re-created by the new organizer
        (one invitation to every attendee and room), then the old meeting goes (cancelled by its old organizer when
        his copy exists, else removed silently from the attendees). A series moves from its next occurrence.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Result, [Parameter(Mandatory)][pscustomobject]$NewOrganizer, [string]$Method = 'Recreate', [datetime]$From = [datetime]::UtcNow, [string]$Comment)
    if ($Method -notin 'Recreate', 'Auto') { throw 'Exchange Server has no supported native organizer transfer: use -TransferMethod Recreate.' }
    if ([string](Get-McoProperty $Result.Request 'Mode') -eq 'Rooms') { throw 'Transfer moves the meetings of organizers: search them with -Organizer (a rooms search keeps occurrences).' }
    $now = [datetime]::UtcNow
    $todo = [Collections.Generic.List[object]]::new()
    $skipped = [Collections.Generic.List[object]]::new()
    $new = @($NewOrganizer.Address, $NewOrganizer.Input) | Where-Object { $_ }
    foreach ($meeting in @($Result.Meetings | Where-Object Selected)) {
        $ref = Get-McoTransferReference $meeting $Result
        $occ = @(if ($ref) { @($ref.Event.Occurrences) | Where-Object { $_ -and ([datetime]$_.Start).ToUniversalTime() -gt $now } })
        $reason = if ($new -contains [string]$meeting.Organizer -or $new -contains [string]$meeting.OrganizerKey) { "already organized by $($NewOrganizer.Address)" }
            elseif ([string](Get-McoProperty $meeting 'NewMeetingId')) { "already transferred by the run of this report (to $(Get-McoProperty $meeting 'NewOrganizer'))" }
            elseif ($meeting.Cancelled) { 'cancelled meeting' }
            elseif (-not $ref) { 'no copy to read the meeting from' }
            elseif ($meeting.Kind -ne 'Series' -and ([datetime]$meeting.End).ToUniversalTime() -le $now) { 'already over' }
            elseif ($meeting.Kind -eq 'Series' -and -not $occ.Count) { 'no occurrence to come in the period' }
            else { '' }
        if ($reason) { $skipped.Add([pscustomobject]@{ Meeting = $meeting; Reason = $reason }); continue }
        $todo.Add($meeting)
    }
    $who = if ($NewOrganizer.Name) { "$($NewOrganizer.Name) <$($NewOrganizer.Address)>" } else { $NewOrganizer.Address }
    $active = @($todo | Where-Object OrganizerCopy -eq 'Present').Count
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add(('{0} meeting(s) re-created in the calendar of {1}: every attendee and room receives ONE invitation from him and answers again' -f $todo.Count, $who))
    if (@($todo | Where-Object Kind -eq 'Series').Count) { $lines.Add('A series moves from its next occurrence (the occurrences before stay in Backup.json); exceptions of the old series are not carried over') }
    $lines.Add('The old room copies are removed first (silently) so that the rooms accept the new meeting without a conflict')
    if ($active) { $lines.Add(('{0} of them still in the calendar of their old organizer: he cancels it, with the message "{1}"' -f $active, ($Comment -f $who))) }
    if ($todo.Count - $active) { $lines.Add(('{0} without an organizer copy: the old attendee copies are removed without a message' -f ($todo.Count - $active))) }
    if ($skipped.Count) { $lines.Add(('{0} meeting(s) not transferred: {1}' -f $skipped.Count, ((@($skipped | Group-Object Reason | ForEach-Object { "$($_.Count) $($_.Name)" })) -join '; '))) }
    [pscustomobject]@{ Native = @(); Recreate = $todo.ToArray(); Skipped = $skipped.ToArray(); NewOrganizer = $NewOrganizer; Lines = $lines.ToArray(); Text = $lines -join " $($script:Dot) " }
}

function Get-McoTransferReference {
    <#
        The copy the meeting is re-created from: the organizer's, else an attendee's (a room shows the organizer's name).
        A meeting of a report (-FromReport: the items are not in it) is read again through EWS, with its occurrences in
        the period of the report.
    #>
    param([Parameter(Mandatory)]$Meeting, $Result)
    $ref = Get-McoBestCopy $Meeting -WithEvent
    if ($ref -or -not $Result -or -not [MeetingCleanupOnPremNative.Fast]::Text($Result, 'FromReport')) { return $ref }
    $copy = Get-McoBestCopy $Meeting
    if (-not $copy) { return $null }
    $request = [MeetingCleanupOnPremNative.Fast]::Prop($Result, 'Request')
    $start = [MeetingCleanupOnPremNative.Fast]::Prop($request, 'Start'); $end = [MeetingCleanupOnPremNative.Fast]::Prop($request, 'End')
    $start = if ($start) { ([datetime]$start).ToUniversalTime() } else { [datetime]::UtcNow.Date }
    $end = if ($end) { ([datetime]$end).ToUniversalTime() } else { $start.AddDays(365) }
    $hit = $null
    try { $hit = @(Get-McoMailboxEvents -Mailbox $copy.Mailbox -StartUtc $start -EndUtc $end -Uid $Meeting.MeetingId) | Select-Object -First 1 } catch { }
    if (-not $hit) { return $null }
    $copy | Add-Member -NotePropertyName Event -NotePropertyValue $hit -Force
    $copy
}

function Get-McoShiftedRecurrence {
    <#
        The recurrence of the new series: same pattern, starting on the date of its first occurrence. A numbered
        range keeps the occurrences still to come.
    #>
    param([Parameter(Mandatory)][string]$Xml, [Parameter(Mandatory)][datetime]$FirstStartUtc, [string]$TimeZoneId, [int]$Before = 0)
    $doc = [xml]$Xml
    $zone = try { if ($TimeZoneId) { [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId) } else { [TimeZoneInfo]::Utc } } catch { [TimeZoneInfo]::Utc }
    $localDate = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::SpecifyKind($FirstStartUtc, [DateTimeKind]::Utc), $zone).ToString('yyyy-MM-dd')
    foreach ($node in $doc.SelectNodes("//*[local-name()='StartDate']")) { $node.InnerText = $localDate }
    foreach ($node in $doc.SelectNodes("//*[local-name()='NumberOfOccurrences']")) { $node.InnerText = [string][Math]::Max(1, [int]$node.InnerText - $Before) }
    $doc.DocumentElement.OuterXml
}

function Add-McoNewOrganizerCopy {
    param([Parameter(Mandatory)]$Meeting, [Parameter(Mandatory)][string]$Mailbox, [string]$EventId, [string]$MeetingId, [string]$Detail)
    $Meeting.Copies.Add([pscustomobject]@{
            MeetingId = $MeetingId; Mailbox = $Mailbox; Role = 'New organizer'; Via = 'EWS CreateItem'; EventId = $EventId; Subject = $Meeting.Subject; Response = 'Organizer'; ShowAs = ''
            Cancelled = $false; Action = 'Transfer'; Result = 'Created'; HttpStatus = 200; Detail = $Detail; Verified = ''; ActionUtc = [datetime]::UtcNow.ToString('o'); Occurrence = ''; OccurrenceStart = ''; SeriesId = ''; Event = $null
        })
}

function Invoke-McoTransfer {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Settings, [Parameter(Mandatory)][pscustomobject]$Result, [Parameter(Mandatory)][pscustomobject]$Plan, [string]$Comment, [string]$BackupPath, [datetime]$From = [datetime]::UtcNow)
    $started = [datetime]::UtcNow
    $dot = $script:Dot
    $new = $Plan.NewOrganizer
    $who = if ($new.Name) { "$($new.Name) <$($new.Address)>" } else { $new.Address }
    $message = $Comment -f $who
    $Result.Action = 'Transfer'
    $Result | Add-Member -NotePropertyName NewOrganizer -NotePropertyValue $new.Address -Force
    foreach ($m in @($Result.Meetings | Where-Object { -not $_.Selected })) { $m.Status = 'Skipped' }
    foreach ($skip in $Plan.Skipped) { $skip.Meeting.Status = 'Skipped'; $skip.Meeting.Notes.Add("Not transferred: $($skip.Reason).") }
    Write-McoNextStep "Transfer to $who" 'User'
    foreach ($line in $Plan.Lines) { Write-McoItem Info $line }
    if ($BackupPath -and $Plan.Recreate.Count) { $Result | Add-Member -NotePropertyName BackupFile -NotePropertyValue (Save-McoBackup $Result ([pscustomobject]@{ Meetings = $Plan.Recreate; Action = 'Transfer' }) $BackupPath) -Force }
    $now = [datetime]::UtcNow
    $index = 0
    foreach ($meeting in $Plan.Recreate) {
        $index++
        Write-McoProgress (($index - 1) / [Math]::Max(1, $Plan.Recreate.Count)) ('{0:N0}/{1:N0} meetings re-created' -f ($index - 1), $Plan.Recreate.Count)
        $reference = Get-McoTransferReference $meeting
        $source = $reference.Event
        $meeting.TransferMethod = 'Recreate'; $meeting.NewOrganizer = $new.Address
        # ---- the new meeting: the next occurrence for a series, the same slot for a single meeting -----------------
        $first = $source
        $recurrence = ''
        if ($meeting.Kind -eq 'Series' -and $source.RecurrenceXml) {
            $next = @($source.Occurrences | Where-Object { ([datetime]$_.Start).ToUniversalTime() -gt $now } | Sort-Object Start)[0]
            $before = 0
            if ($source.RecurrenceXml -match 'NumberedRecurrence') {
                $past = @(Get-McoMailboxEvents -Mailbox $reference.Mailbox -StartUtc ([datetime]$source.Start).AddMinutes(-1) -EndUtc ([datetime]$next.Start) -Uid $meeting.MeetingId -Occurrences)
                $before = @($past | Where-Object { ([datetime]$_.Start) -lt ([datetime]$next.Start) }).Count
            }
            $recurrence = Get-McoShiftedRecurrence -Xml $source.RecurrenceXml -FirstStartUtc ([datetime]$next.Start) -TimeZoneId $source.TimeZoneId -Before $before
            $first = [pscustomobject]@{ Start = $next.Start; End = $next.End }
        }
        $attendees = @(@($source.RequiredAttendees) | Where-Object { $_ -and $_ -ne $new.Address -and $_ -ne $new.Input })
        $optional = @(@($source.OptionalAttendees) | Where-Object { $_ -and $_ -ne $new.Address -and $_ -ne $new.Input })
        $rooms = @($source.Resources | Where-Object { $_ })
        $event = [pscustomobject]@{
            Subject = $meeting.Subject; Start = $first.Start; End = $first.End; Location = $source.Location
            RequiredAttendees = $attendees; OptionalAttendees = $optional; Resources = $rooms; RecurrenceXml = $recurrence; TimeZoneId = $source.TimeZoneId
        }
        $body = (@($message, '', [string]$source.Body) -join "`r`n").Trim()

        # ---- 1. old room copies removed silently: the slot is free for the new invitation -------------------------
        $oldRooms = @(foreach ($c in $meeting.Copies) { if ($c.EventId -and $c.Role -eq 'Room') { $c } })
        if ($oldRooms.Count) { Remove-McoCopyWaves -Copies $oldRooms -NoProgress }

        # ---- 2. the new meeting, sent to every attendee and room ----------------------------------------------------
        $created = Invoke-McoCreateAppointment -Mailbox $new.Address -Event $event -Comment $body
        if (-not $created.Ok) {
            $meeting.Status = 'Failed'
            $meeting.Notes.Add("EWS could not create the meeting of $($new.Address): $($created.Error).$(if ($oldRooms.Count) { ' The old room copies were removed: -Action Restore -FromReport <this report> puts them back.' })")
            Write-McoItem Fail "$($meeting.Subject): $($created.Error)"
            continue
        }
        $meeting.NewMeetingId = $created.MeetingId
        Add-McoNewOrganizerCopy -Meeting $meeting -Mailbox $new.Address -EventId $created.EventId -MeetingId $created.MeetingId -Detail "re-created and sent ($(if ($recurrence) { 'series from ' + (Format-McoDate $first.Start $Settings.TimeZone) } else { 'single meeting' }))"

        # ---- 3. the old meeting: cancelled by its organizer, then the copies left removed silently ------------------
        $org = @(foreach ($c in $meeting.Copies) { if ($c.EventId -and $c.Role -eq 'Organizer') { $c } })
        foreach ($copy in $org) {
            $response = Invoke-McoCancelItem -Mailbox $copy.Mailbox -EventId $copy.EventId -Comment $message
            $copy.ActionUtc = [datetime]::UtcNow.ToString('o')
            Set-McoCopyResult -Copy $copy -Action 'Cancel' -Response $response -Success 'Cancelled'
        }
        if ($org.Count) { Start-Sleep -Seconds 5 }
        $left = @(foreach ($c in $meeting.Copies) { if ($c.EventId -and $c.Role -eq 'Attendee') { $c } })
        if ($left.Count) { Remove-McoCopyWaves -Copies $left -NoProgress }
        $bad = 0; foreach ($c in $meeting.Copies) { if ($c.Result -eq 'Failed') { $bad++ } }
        $meeting.Status = if ($bad) { 'Partial' } else { 'Transferred' }
        if ($bad) { $meeting.Notes.Add("Re-created, but $bad old cop$(if ($bad -eq 1) { 'y' } else { 'ies' }) could not be removed (see the copies).") }
        Write-McoProgress ($index / [Math]::Max(1, $Plan.Recreate.Count)) ('{0:N0}/{1:N0} meetings re-created' -f $index, $Plan.Recreate.Count)
    }
    $doneCount = 0; foreach ($m in $Plan.Recreate) { if ($m.Status -in 'Transferred', 'Partial') { $doneCount++ } }
    Write-McoItem $(if ($doneCount -lt $Plan.Recreate.Count) { 'Warn' } else { 'Ok' }) ('{0} of {1} re-created with {2} and sent {3} {4}' -f $doneCount, $Plan.Recreate.Count, $new.Address, $dot, (Format-McoDuration ([datetime]::UtcNow - $started).TotalSeconds)) -Icon Calendar

    # ---- verify: old copies gone, new meeting in the calendar of the new organizer --------------------------------
    $gone = [Collections.Generic.List[object]]::new()
    $created = [Collections.Generic.List[object]]::new()
    foreach ($m in $Result.Meetings) { foreach ($c in $m.Copies) { if ($c.Result -eq 'Removed' -and $c.EventId) { $gone.Add($c) } elseif ($c.Role -eq 'New organizer') { $created.Add($c) } } }
    if ($Settings.Verify -and ($gone.Count -or $doneCount)) {
        Write-McoNextStep 'Verify' 'Search'
        $still = 0; $checked = 0; $all = $gone.Count + $created.Count
        foreach ($c in $gone) {
            $exists = Test-McoItemExists -Mailbox $c.Mailbox -EventId $c.EventId
            if ($exists -eq $false) { $c.Verified = 'Yes' } elseif ($exists) { $c.Verified = 'No'; $still++ } else { $c.Verified = 'Unknown' }
            $checked++
            Write-McoProgress ($checked / [Math]::Max(1, $all)) ('{0:N0}/{1:N0} copies verified' -f $checked, $all)
        }
        foreach ($c in $created) {
            $c.Verified = if (Test-McoItemExists -Mailbox $c.Mailbox -EventId $c.EventId) { 'Yes' } else { 'No' }
            $checked++
            Write-McoProgress ($checked / [Math]::Max(1, $all)) ('{0:N0}/{1:N0} copies verified' -f $checked, $all)
        }
        Write-McoItem $(if ($still) { 'Warn' } else { 'Ok' }) ('{0} of {1} old copies verified gone' -f ($gone.Count - $still), $gone.Count) -Icon Search
    }
    Update-McoResultCounts $Result
    $selected = 0; $failedCount = 0; $warned = $false
    foreach ($m in $Result.Meetings) {
        if ($m.Status -in 'Failed', 'Partial') { $warned = $true }
        if (-not $m.Selected -or $m.Status -eq 'Skipped') { continue }
        $selected++
        if ($m.Status -eq 'Failed') { $failedCount++ }
    }
    $Result.Status = if ($selected -and $failedCount -eq $selected) { 'Failed' } elseif ($warned -or $Plan.Skipped.Count) { 'Warning' } else { 'Completed' }
    $Result.CompletedUtc = [datetime]::UtcNow.ToString('o')
    $Result.DurationSeconds = [Math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
    $Result
}

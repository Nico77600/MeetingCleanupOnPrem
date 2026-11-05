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
            elseif ([string](Get-McoProperty $meeting 'Scope') -eq 'Occurrences') { 'occurrences of a series: transfer the whole series (search its organizer with -SeriesScope Whole)' }
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

# SIG # Begin signature block
# MIImIwYJKoZIhvcNAQcCoIImFDCCJhACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASnYx9u+BRoVY1
# FbsCOaUGFMq/idnDjRhI4Y12v2Xka6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA1sPnbEsK8
# F8SqzW8t7A/t0TFG0rq14T5KO8nrDrXbmTANBgkqhkiG9w0BAQEFAASCAgAbILUT
# AiAHwFfgb/4nTyZ9bnoSfUBXRo6HPO/bj2cNvJpm1V5JitlwxRsrwsAECXNzOy68
# ZcyFyBpRpB4e13O7QxO6LD6evPrVnzUZn5E4iirJQ1Q5hOHq9/aKOGOKIFloT8qI
# x7P/PSMjYC3elgJHosVqvAHm5sqw6y2c5sCByh//2wU6NXExgy1k7RmjSjxjlNHO
# C0ziyppn6vG3KdQ38dTPlFc2oieL00ZI98yF+nH3KAFNLkCc1ESAttj/fjtgeQmb
# hzeYb1pVCPBSZJFDHyLzRS34yNstutEHwH6LML5kAyZtIVdcjEvAvJDZIsLhJv4w
# yd8HU74GGg7cFiLcsqCsb/oPszJLpeUGHQVSU/gy3UGPhhC5vgSYIIUQCCsS0PK+
# 5Ef5FpgeBJ2JnlnAhddDW997NNUFMD7pvNYdAFdX2viI9yhqP4HTdX/KbZPTC4bO
# 7Hn9ZtHp9Uo7bXICaN4/2FEOTJthvvbJ1l961DZXFwCRgB5Pi8JXM6NY+k9leroJ
# mJI4EzxXipthAyDqe7KslKZgaiFqV7rV+PtqgpCK8QlLXY1l3rz42akbEemAL6iF
# nTvpXWT0v8bDX/Vy8kEy1bgDra3OM7+Im3E6U0xWIvpzN0FRm7h7FvcTRoVI/pl3
# juXHwwLD1/rQyVKhirwaF4LQrHiCEX490Q9rmqGCFMcwghTDBgorBgEEAYI3AwMB
# MYIUszCCFK8GCSqGSIb3DQEHAqCCFKAwghScAgEDMQ0wCwYJYIZIAWUDBAIBMIGC
# BgsqhkiG9w0BCRABBKBzBHEwbwIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEII02VP0TsxtHRiFabTMum8dOywoKUN1djZMwFPhReMmIAggAzZ7MJ/H5
# /xgPMjAyNjExMDUxNzIyNDFaMAMCAQECCQDHAeow4KZT8aCCEacwggTzMIIC26AD
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
# BgkqhkiG9w0BCQUxDxcNMjYxMTA1MTcyMjQxWjAoBgkqhkiG9w0BCTQxGzAZMAsG
# CWCGSAFlAwQCAaEKBggqhkjOPQQDAjAvBgkqhkiG9w0BCQQxIgQg4dr/Z7uPdwhO
# 5AFOvY0jepz0c0yb051rAqRyRZrBcqUwgcgGCyqGSIb3DQEJEAIvMYG4MIG1MIGy
# MIGvBCBXT+oXqvpefGormwjJmLFzXkULOHPKROaiTKCTDmPzsDCBijB2pHQwcjEL
# MAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREw
# DwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcg
# SXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO4zAKBggqhkjOPQQDAgRH
# MEUCIA52H4jR6KXqA5m1KrR/2NbD9ihPjRZthh5PVlJMRspEAiEAvAqdL+F/qLwl
# Mhhz4T8NiVXQ/RCPl6s0R1VI4ttDCn4=
# SIG # End signature block

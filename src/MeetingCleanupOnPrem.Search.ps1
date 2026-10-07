$script:StepIndex = 0
$script:StepTotal = 0

function Initialize-McoSteps { param([int]$Total) $script:StepIndex = 0; $script:StepTotal = $Total }
function Write-McoNextStep {
    param([Parameter(Mandatory)][string]$Title, [string]$Icon = 'Info')
    $script:StepIndex++
    Write-McoStep -Number $script:StepIndex -Total ([Math]::Max($script:StepIndex, $script:StepTotal)) -Title $Title -Icon $Icon
}

function Resolve-McoOrganizer {
    param([Parameter(Mandatory)][string[]]$Identity, [hashtable]$Settings)
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $result = [Collections.Generic.List[object]]::new()
    foreach ($inputAddress in $Identity) {
        $address = ([string]$inputAddress).Trim()
        if (-not $address -or -not $seen.Add($address)) { continue }
        $fallback = [pscustomobject]@{
            Input = $address; DisplayName = ''; PrimaryAddress = $address.ToLowerInvariant(); Addresses = @($address.ToLowerInvariant())
            UserId = ''; Account = 'Unknown'; State = 'Unknown'; Detail = 'EWS does not expose the directory: mailbox reachability is checked during the calendar search.'
        }
        $directoryMode = if ($Settings) { [string]$Settings.DirectoryMode } else { 'Auto' }
        $recipient = if ($directoryMode -ne 'None') { Resolve-McoExchangeRecipient -Identity $address } else { $null }
        if (-not $recipient) { $result.Add($fallback); continue }
        if ($recipient.Error) {
            # Not found by Get-Recipient: a deleted mailbox, searched with the address typed (its calendar is not opened).
            if ($recipient.Error -match "couldn't be found|could not be found|not found") { $fallback.State = 'NotInDirectory'; $fallback.Account = 'Deleted'; $fallback.Detail = 'not in the directory (deleted mailbox?): searched with the address typed' }
            else { $fallback.Detail = "Exchange Management Shell lookup failed: $($recipient.Error). EWS will compare the typed address only." }
            $result.Add($fallback)
            continue
        }
        $isMailbox = [string]$recipient.RecipientTypeDetails -match 'Mailbox'
        $primary = if ($recipient.PrimaryAddress) { $recipient.PrimaryAddress } else { $address.ToLowerInvariant() }
        $result.Add([pscustomobject]@{
            Input = $address; DisplayName = $recipient.DisplayName; PrimaryAddress = $primary; Addresses = @($recipient.Addresses + $address.ToLowerInvariant() | Select-Object -Unique)
            UserId = ''; Account = 'Present'; State = if ($isMailbox) { 'Mailbox' } else { 'NoMailbox' }
            Detail = if ($isMailbox) { 'resolved by Exchange Management Shell' } else { "resolved as $($recipient.RecipientTypeDetails), not a mailbox" }
        })
    }
    $result.ToArray()
}

function Get-McoSearchMailboxes {
    param([Parameter(Mandatory)][hashtable]$Settings, [Parameter(Mandatory)][pscustomobject]$Request, [object[]]$Organizers = @())
    $list = [Collections.Generic.List[object]]::new()
    $warnings = [Collections.Generic.List[string]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $rooms = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in @($Settings.Rooms) + @($Request.Room)) { if ($a) { [void]$rooms.Add(([string]$a).ToLowerInvariant()) } }
    if ($Settings.RoomFile) { foreach ($a in Read-McoAddressFile $Settings.RoomFile) { [void]$rooms.Add($a.ToLowerInvariant()) } }
    if ($Request.RoomFile) { foreach ($a in Read-McoAddressFile $Request.RoomFile) { [void]$rooms.Add($a.ToLowerInvariant()) } }
    $add = {
        param([string]$Address, [string]$Scope)
        $a = ([string]$Address).Trim().ToLowerInvariant()
        if ($a -and $seen.Add($a)) { $list.Add([pscustomobject]@{ Address = $a; Scope = $Scope; IsRoom = $rooms.Contains($a) }) }
    }
    if ($Request.Mode -eq 'Rooms') {
        foreach ($a in @($Request.Room)) { & $add $a 'Rooms' }
    }
    else {
        if (@($Request.SearchIn) -contains 'Organizer') { foreach ($o in $Organizers) { if ($o.PrimaryAddress -match $script:SmtpPattern -and $o.State -ne 'NotInDirectory') { & $add $o.PrimaryAddress 'Organizer' } } }
        if (@($Request.SearchIn) -contains 'Rooms') {
            # Every room mailbox of the organization (Exchange Management Shell), with the rooms of the configuration.
            if ([string]$Settings.DirectoryMode -ne 'None' -and $Settings.AllRooms -ne $false) {
                $directory = Get-McoExchangeMailboxAddresses -RecipientTypeDetails 'RoomMailbox'
                foreach ($a in $directory.Addresses) { [void]$rooms.Add($a) }
                if ($directory.Available -and $directory.Error) { $warnings.Add("The room mailboxes could not be listed: $($directory.Error). Only the rooms of Search.Rooms and Search.RoomFile are searched.") }
            }
            if (-not $rooms.Count) { $warnings.Add('No room to search: Exchange PowerShell lists them (ManagementShell), or give them in Search.Rooms / Search.RoomFile.') }
            foreach ($a in $rooms) { & $add $a 'Rooms' }
        }
        if (@($Request.SearchIn) -contains 'Mailboxes') {
            foreach ($a in @($Settings.Mailboxes) + @($Request.Mailboxes)) { & $add $a 'Mailboxes' }
            if ($Settings.MailboxFile) { foreach ($a in Read-McoAddressFile $Settings.MailboxFile) { & $add $a 'Mailboxes' } }
            if ($Request.MailboxFile) { foreach ($a in Read-McoAddressFile $Request.MailboxFile) { & $add $a 'Mailboxes' } }
        }
        if (@($Request.SearchIn) -contains 'AllMailboxes') {
            if ([string]$Settings.DirectoryMode -eq 'None') {
                $warnings.Add('AllMailboxes needs Exchange Management Shell directory access. Set Search.DirectoryMode to Auto or ExchangePowerShell, or supply -MailboxFile.')
            }
            else {
                $directory = Get-McoExchangeMailboxAddresses
                if ($directory.Addresses.Count) { foreach ($a in $directory.Addresses) { & $add $a 'AllMailboxes' } }
                elseif ($directory.Error) { $warnings.Add("AllMailboxes could not be enumerated: $($directory.Error)") }
                else { $warnings.Add('AllMailboxes needs Exchange Management Shell cmdlets in the current PowerShell session. Supply -MailboxFile as a fallback.') }
            }
        }
    }
    [pscustomobject]@{ Mailboxes = @($list | Where-Object Address); Rooms = $rooms; Warnings = $warnings.ToArray() }
}

function New-McoCopy {
    param([Parameter(Mandatory)]$Event, [Parameter(Mandatory)][string]$Role, [string]$Via, [string]$Result = '', [string]$Detail = '')
    $isOccurrence = [string]$Event.AppointmentType -in 'Occurrence', 'Exception'
    [pscustomobject]@{
        MeetingId = $Event.MeetingId; Mailbox = ([string]$Event.Mailbox).ToLowerInvariant(); Role = $Role; Via = $Via
        EventId = $Event.EventId; Subject = $Event.Subject; Response = $Event.Response; ShowAs = ''
        Cancelled = [bool]$Event.IsCancelled; Action = ''; Result = $Result; HttpStatus = 0; Detail = $Detail; Verified = ''; ActionUtc = ''
        Occurrence = if ($isOccurrence) { ([datetime]$Event.Start).ToString('o') } else { '' }
        OccurrenceStart = if ($isOccurrence) { ([datetime]$Event.Start).ToString('o') } else { '' }
        SeriesId = [string](Get-McoProperty $Event 'SeriesId'); Event = $Event
    }
}

function New-McoPlaceholderCopy {
    <# A row of the report for an attendee that has no copy to act on (external, not found, group expanded...). #>
    param([Parameter(Mandatory)]$Meeting, [Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$Role, [Parameter(Mandatory)][string]$Via, [Parameter(Mandatory)][string]$Result, [Parameter(Mandatory)][string]$Detail)
    [pscustomobject]@{
        MeetingId = $Meeting.MeetingId; Mailbox = $Mailbox; Role = $Role; Via = $Via; EventId = ''; Subject = $Meeting.Subject; Response = ''; ShowAs = ''
        Cancelled = $false; Action = ''; Result = $Result; HttpStatus = 0; Detail = $Detail; Verified = ''; ActionUtc = ''; Occurrence = ''; OccurrenceStart = ''; SeriesId = ''; Event = $null
    }
}

function New-McoMeeting {
    param([Parameter(Mandatory)]$Event, [Parameter(Mandatory)][hashtable]$Settings)
    $fast = [MeetingCleanupOnPremNative.Fast]
    $occ = @(Get-McoProperty $Event 'Occurrences')
    $series = [string]$Event.AppointmentType -in 'RecurringMaster', 'Occurrence', 'Exception' -or $occ.Count
    $next = $null
    foreach ($o in $occ) { if ($null -eq $next -or $o.Start -lt $next) { $next = $o.Start } }
    $zone = Get-McoTimeZone $Settings.TimeZone
    $attendees = [Collections.Generic.List[object]]::new()
    foreach ($a in @($Event.RequiredAttendees)) { $attendees.Add([pscustomobject]@{ Type = 'required'; Address = $a }) }
    [pscustomobject]@{
        MeetingId = $Event.MeetingId; Subject = $Event.Subject; Organizer = $Event.Organizer; OrganizerName = $Event.OrganizerName; OrganizerKey = $Event.Organizer
        Kind = if ($series) { 'Series' } else { 'Single' }
        Start = $Event.Start; End = $Event.End; StartText = $fast::FormatDate($Event.Start, $zone, $false, $false); EndText = $fast::FormatDate($Event.End, $zone, $false, $false)
        NextInPeriod = if ($next) { $fast::FormatDate($next, $zone, $false, $false) } else { '' }
        Recurrence = $fast::Recurrence($fast::Text($Event, 'RecurrenceXml'))
        Location = $Event.Location; Cancelled = $Event.IsCancelled; OrganizerCopy = 'Not checked'
        Attendees = $attendees.ToArray()
        Copies = [Collections.Generic.List[object]]::new(); Selected = $true; Status = 'Found'; Notes = [Collections.Generic.List[string]]::new()
        SubjectFromRoom = $false; Scope = 'Whole'; Occurrences = $occ.Count; RecurrenceData = $null; TimeZone = $fast::Text($Event, 'TimeZoneId'); NewOrganizer = ''; NewMeetingId = ''; TransferMethod = ''
    }
}

function Format-McoEwsRecurrence {
    <# An EWS Recurrence element in words: Weekly (Monday), 4 occurrences from 2026-10-12 (compiled). #>
    param([AllowEmptyString()][string]$Xml)
    [MeetingCleanupOnPremNative.Fast]::Recurrence($Xml)
}

function Get-McoBestCopy {
    <# The copy a meeting is read from: the organizer's, else an attendee's, else a room's (one with an item ID). #>
    param([Parameter(Mandatory)]$Meeting, [switch]$WithEvent)
    $best = $null; $bestRank = 9
    foreach ($c in $Meeting.Copies) {
        if (-not $c.EventId) { continue }
        if ($WithEvent -and -not [MeetingCleanupOnPremNative.Fast]::Prop($c, 'Event')) { continue }
        $rk = switch ($c.Role) { 'Organizer' { 0 } 'Attendee' { 1 } 'Room' { 2 } default { 9 } }
        if ($rk -lt $bestRank) { $bestRank = $rk; $best = $c }
    }
    $best
}

function Get-McoCopyRole {
    <# Organizer (the copy says so), Room (a room or equipment mailbox), else Attendee. #>
    param([Parameter(Mandatory)]$Event, $RoomSet, [hashtable]$Settings)
    if ([string]$Event.Response -eq 'Organizer') { return 'Organizer' }
    if ($RoomSet -and $RoomSet.Contains([string]$Event.Mailbox)) { return 'Room' }
    if ($Settings -and [string]$Settings.DirectoryMode -ne 'None' -and (Get-McoRecipientKind -Address $Event.Mailbox) -in 'Room', 'Equipment') { return 'Room' }
    'Attendee'
}

function Find-McoMeetings {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Settings, [Parameter(Mandatory)][pscustomobject]$Request)
    if (-not $script:Ews) { throw 'EWS is not connected.' }
    # The CalendarView of each mailbox is read once for the whole search (organizers, rooms, then attendees).
    $script:CalendarViewCache = @{}
    $script:CalendarWarnings = [Collections.Generic.List[string]]::new()
    try {
        $result = Find-McoMeetingsCore -Settings $Settings -Request $Request
        # Calendars read only in part: in the warnings of the run (the report shows them).
        foreach ($w in $script:CalendarWarnings) { $result.Warnings.Add($w) }
        if ($script:CalendarWarnings.Count -and $result.Status -eq 'Completed') { $result.Status = 'Warning' }
        $result
    }
    finally { $script:CalendarViewCache = $null; $script:CalendarWarnings = $null }
}

function Find-McoMeetingsCore {
    param([Parameter(Mandatory)][hashtable]$Settings, [Parameter(Mandatory)][pscustomobject]$Request)
    $started = [datetime]::UtcNow
    $dot = $script:Dot
    $warnings = [Collections.Generic.List[string]]::new()
    $roomsMode = $Request.Mode -eq 'Rooms'
    $directory = [string]$Settings.DirectoryMode -ne 'None'
    Write-McoNextStep $(if ($roomsMode) { "Rooms ($(@($Request.Room).Count))" } else { 'Organizers' }) $(if ($roomsMode) { 'Room' } else { 'User' })
    $organizers = @(if (-not $roomsMode) { Resolve-McoOrganizer -Identity $Request.Organizer -Settings $Settings })
    foreach ($o in $organizers) { Write-McoItem $(if ($o.State -eq 'Mailbox') { 'Ok' } else { 'Info' }) ("{0} {1} {2}{3}" -f $(if ($o.DisplayName) { "$($o.DisplayName) <$($o.PrimaryAddress)>" } else { $o.Input }), $dot, $o.Detail, $(if (@($o.Addresses).Count -gt 1) { " $dot $(@($o.Addresses).Count) addresses compared" } else { '' })) -Icon User }
    if ($roomsMode) { foreach ($r in @($Request.Room)) { Write-McoItem Info $r -Icon Room } }
    $addressMap = @{}
    foreach ($o in $organizers) { foreach ($a in $o.Addresses) { $addressMap[([string]$a).ToLowerInvariant()] = $o } }
    $orgMailboxes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($o in $organizers) { foreach ($a in @($o.Addresses) + [string]$o.PrimaryAddress) { if ($a) { [void]$orgMailboxes.Add($a) } } }
    $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($id in @($Request.MeetingId)) { if ($id) { [void]$ids.Add([string]$id) } }

    Write-McoNextStep 'Mailboxes to search' 'Search'
    $search = Get-McoSearchMailboxes -Settings $Settings -Request $Request -Organizers $organizers
    foreach ($w in $search.Warnings) { $warnings.Add($w); Write-McoItem Warn $w }
    $scopes = [ordered]@{}
    foreach ($mb in $search.Mailboxes) { $scopes[$mb.Scope] = 1 + [int]$scopes[$mb.Scope] }
    foreach ($scope in $scopes.Keys) { Write-McoItem Info ('{0}: {1:N0} mailbox(es)' -f $scope, $scopes[$scope]) -Icon $(if ($scope -eq 'Rooms') { 'Room' } else { 'Mail' }) }

    Write-McoNextStep 'Search' 'Calendar'
    $meetings = [ordered]@{}
    # Per meeting: the items already listed (mailbox|item ID) and the mailboxes holding a copy of it.
    $itemsOf = @{}
    $withCopy = @{}
    $addCopy = {
        param($meeting, $event, [string]$role, [string]$via)
        $item = '{0}|{1}' -f $event.Mailbox, $event.EventId
        if (-not $itemsOf[$meeting.MeetingId].Add($item)) { return }
        $meeting.Copies.Add((New-McoCopy -Event $event -Role $role -Via $via))
        if ($event.EventId) { [void]$withCopy[$meeting.MeetingId].Add([string]$event.Mailbox) }
    }
    $mailboxes = @($search.Mailboxes)
    $searched = [ordered]@{ Mailboxes = $mailboxes.Count; Read = 0; Events = 0; NoMailbox = 0; Denied = 0; Errors = 0 }
    $roomSet = $search.Rooms
    $done = 0
    foreach ($mb in $mailboxes) {
        $events = $null
        try { $events = Get-McoMailboxEvents -Mailbox $mb.Address -StartUtc $Request.Start -EndUtc $Request.End -PageSize ([int]$Settings.PageSize) -Occurrences:$roomsMode; $searched.Read++ }
        catch {
            $text = $_.Exception.Message
            if ($text -match 'ErrorNonExistentMailbox|ErrorMailboxNotFound|ErrorInvalidSmtpAddress') { $searched.NoMailbox++ } elseif ($text -match 'ErrorImpersonat|ErrorAccessDenied') { $searched.Denied++ } else { $searched.Errors++ }
            $warnings.Add("$($mb.Address): $text"); Write-McoItem Warn "$($mb.Address): $text"
        }
        foreach ($event in @($events)) {
            if ($null -eq $event) { continue }
            $searched.Events++
            if (-not $event.MeetingId) { continue }
            if ($event.ItemClass -and [string]$event.ItemClass -notlike 'IPM.Appointment*') { continue }
            $isOwn = [string]$event.Response -eq 'Organizer' -and $orgMailboxes.Contains($mb.Address)
            $org = if ($event.Organizer) { $addressMap[[string]$event.Organizer] } else { $null }
            if (-not ($roomsMode -or $isOwn -or [bool]$org)) { continue }
            # A plain appointment (no attendee) is not a meeting.
            if ($isOwn -and -not @($event.RequiredAttendees).Count -and -not @($event.OptionalAttendees).Count -and -not @($event.Resources).Count) { continue }
            if ($ids.Count -and -not $ids.Contains($event.MeetingId)) { continue }
            if (-not $meetings.Contains($event.MeetingId)) {
                $meetings[$event.MeetingId] = New-McoMeeting -Event $event -Settings $Settings
                $itemsOf[$event.MeetingId] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $withCopy[$event.MeetingId] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            }
            $meeting = $meetings[$event.MeetingId]
            if ($event.Organizer -and -not $meeting.Organizer) { $meeting.Organizer = $event.Organizer; $meeting.OrganizerKey = $event.Organizer }
            if ($isOwn) { $meeting.OrganizerKey = $mb.Address }
            & $addCopy $meeting $event (Get-McoCopyRole -Event $event -RoomSet $roomSet -Settings $Settings) "$($mb.Scope) calendar"
        }
        $done++
        Write-McoProgress ($done / [Math]::Max(1, $mailboxes.Count)) ('{0:N0}/{1:N0} mailboxes searched' -f $done, $mailboxes.Count)
    }
    Write-McoItem Ok ('{0:N0} mailbox(es) read {1} {2:N0} calendar item(s) {1} {3:N0} meeting(s)' -f $searched.Read, $dot, $searched.Events, $meetings.Count) -Icon Calendar

    Write-McoNextStep 'Attendees, rooms and groups' 'People'
    if (-not $meetings.Count) { Write-McoItem Skip 'No matching meeting: attendee and room lookup skipped.' }
    $accepted = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($d in @($Settings.AcceptedDomains)) { if ($d) { [void]$accepted.Add(([string]$d).ToLowerInvariant()) } }
    # ---- 1. what to look for, meeting by meeting, in the order of their attendees (no EWS call) ------------------
    $actions = [Collections.Generic.List[object]]::new()
    $actionKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $uidsOf = [ordered]@{}
    $expanded = 0
    foreach ($meeting in @($meetings.Values)) {
        $reference = Get-McoBestCopy $meeting
        if (-not $reference) { continue }
        $referenceEvent = $reference.Event
        # A room shows the organizer's name as subject: the subject of an organizer or attendee copy is kept.
        if ($reference.Role -ne 'Room' -and $referenceEvent.Subject) { $meeting.Subject = $referenceEvent.Subject }
        elseif ($reference.Role -eq 'Room') { $meeting.SubjectFromRoom = $true }
        $attendees = [Collections.Generic.List[object]]::new()
        foreach ($a in @($referenceEvent.RequiredAttendees)) { $attendees.Add([pscustomobject]@{ Type = 'required'; Address = $a }) }
        foreach ($a in @($referenceEvent.OptionalAttendees)) { $attendees.Add([pscustomobject]@{ Type = 'optional'; Address = $a }) }
        foreach ($a in @($referenceEvent.Resources)) { $attendees.Add([pscustomobject]@{ Type = 'resource'; Address = $a }) }
        $meeting.Attendees = $attendees.ToArray()
        foreach ($r in @($referenceEvent.Resources)) { if ($r) { [void]$roomSet.Add(([string]$r).ToLowerInvariant()) } }
        $occurrenceStarts = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($c in $meeting.Copies) { if ($c.Occurrence) { [void]$occurrenceStarts.Add([string]$c.Occurrence) } }
        $targets = [Collections.Generic.List[object]]::new()
        $seenAddress = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($address in @($referenceEvent.RequiredAttendees) + @($referenceEvent.OptionalAttendees) + @($referenceEvent.Resources)) {
            if (-not $address -or -not $seenAddress.Add([string]$address)) { continue }
            $address = ([string]$address).ToLowerInvariant()
            $members = @(if ($directory) { Get-McoExchangeGroupMembers -Identity $address })
            if ($members.Count) {
                $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $address -Role 'Attendee' -Via 'Invitation list' -Result 'Expanded' -Detail "distribution group expanded ($($members.Count) member(s))"))
                $expanded++
                foreach ($member in $members) { $targets.Add([pscustomobject]@{ Address = $member.Address; Via = "Group $address" }) }
            }
            else { $targets.Add([pscustomobject]@{ Address = $address; Via = 'Invitation list' }) }
        }
        if ($meeting.Organizer) { $targets.Add([pscustomobject]@{ Address = [string]$meeting.Organizer; Via = 'Organizer of the meeting' }) }
        foreach ($targetInfo in $targets) {
            $target = ([string]$targetInfo.Address).ToLowerInvariant()
            if ($withCopy[$meeting.MeetingId].Contains($target)) { continue }
            $role = if ($roomSet.Contains($target)) { 'Room' } elseif ($targetInfo.Via -eq 'Organizer of the meeting') { 'Organizer' } else { 'Attendee' }
            # The same mailbox reached twice for the same meeting and role (a group and the invitation list): once.
            if (-not $actionKeys.Add("$($meeting.MeetingId)|$target|$role")) { continue }
            $domain = if ($target.Contains('@')) { $target.Split('@')[-1] } else { '' }
            $internal = $target -match '^/o=' -or -not $accepted.Count -or $accepted.Contains($domain)
            $kind = if (-not $internal) { 'External' } elseif ($target -match '^/o=') { 'X500' } else { 'Lookup' }
            $actions.Add([pscustomobject]@{ Meeting = $meeting; Target = $target; Role = $role; Via = $targetInfo.Via; Kind = $kind; Occurrences = $occurrenceStarts })
            if ($kind -eq 'Lookup') {
                $uids = $uidsOf[$target]
                if (-not $uids) { $uids = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal); $uidsOf[$target] = $uids }
                [void]$uids.Add($meeting.MeetingId)
            }
        }
    }
    # ---- 2. each attendee mailbox read once, for all its meetings -------------------------------------------------
    $found = @{}
    $done = 0
    foreach ($target in @($uidsOf.Keys)) {
        $byUid = @{}
        try {
            foreach ($hit in @(Get-McoMailboxEvents -Mailbox $target -StartUtc $Request.Start -EndUtc $Request.End -PageSize ([int]$Settings.PageSize) -Uid @($uidsOf[$target]) -Occurrences:$roomsMode)) {
                if ($null -eq $hit) { continue }
                $list = $byUid[$hit.MeetingId]
                if (-not $list) { $list = [Collections.Generic.List[object]]::new(); $byUid[$hit.MeetingId] = $list }
                $list.Add($hit)
            }
            $found[$target] = @{ Events = $byUid; Error = '' }
        }
        catch { $found[$target] = @{ Events = $byUid; Error = $_.Exception.Message } }
        $done++
        Write-McoProgress ($done / [Math]::Max(1, $uidsOf.Count)) ('{0:N0}/{1:N0} attendee mailboxes searched' -f $done, $uidsOf.Count)
    }
    # ---- 3. the copies of each meeting, in the order of its attendees -----------------------------------------------
    foreach ($a in $actions) {
        $meeting = $a.Meeting; $target = $a.Target; $role = $a.Role
        if ($withCopy[$meeting.MeetingId].Contains($target)) { continue }
        if ($a.Kind -eq 'External') { $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $target -Role $role -Via $a.Via -Result 'Not processed' -Detail 'not in an accepted domain of the organization (external)')); continue }
        if ($a.Kind -eq 'X500') { $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $target -Role $role -Via $a.Via -Result 'Not processed' -Detail 'X500 address: the mailbox no longer exists')); continue }
        $f = $found[$target]
        if ($f.Error) {
            $why = if ($f.Error -match 'ErrorNonExistentMailbox|ErrorMailboxNotFound|ErrorInvalidSmtpAddress') { 'no mailbox with this address (deleted or not a mailbox)' } else { $f.Error }
            $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $target -Role $role -Via $a.Via -Result 'Not processed' -Detail $why))
            continue
        }
        $hits = [Collections.Generic.List[object]]::new()
        foreach ($hit in @($f.Events[$meeting.MeetingId])) {
            if ($null -eq $hit) { continue }
            if ($roomsMode -and $a.Occurrences.Count -and $hit.AppointmentType -and $hit.AppointmentType -ne 'Single' -and -not $a.Occurrences.Contains(([datetime]$hit.Start).ToString('o'))) { continue }
            $hits.Add($hit)
        }
        foreach ($hit in $hits) { & $addCopy $meeting $hit (Get-McoCopyRole -Event $hit -RoomSet $roomSet -Settings $Settings) $a.Via }
        if (-not $hits.Count) { $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $target -Role $role -Via $a.Via -Result 'Not found' -Detail $(if ($role -eq 'Organizer') { 'not in the calendar of its organizer' } else { 'no copy of this meeting in the mailbox' }))) }
    }
    # ---- 4. each meeting: organizer copy, subject, occurrences (one pass over its copies) ---------------------------
    $total = 0; $roomCopies = 0; $skipped = 0; $present = 0
    foreach ($meeting in @($meetings.Values)) {
        $orgCopy = $false; $orgDeleted = $false; $better = $null; $betterRank = 9
        $occ = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($c in $meeting.Copies) {
            if ($c.EventId) {
                $total++
                if ($c.Role -eq 'Organizer') { $orgCopy = $true }
                if ($c.Role -eq 'Room') { $roomCopies++ }
                if ($c.Subject -and $c.Role -in 'Organizer', 'Attendee') { $rk = if ($c.Role -eq 'Organizer') { 0 } else { 1 }; if ($rk -lt $betterRank) { $betterRank = $rk; $better = $c } }
            }
            elseif ($c.Role -eq 'Organizer' -and $c.Detail -match 'no mailbox') { $orgDeleted = $true }
            if ($c.Occurrence) { [void]$occ.Add([string]$c.Occurrence) }
            if ($c.Result -eq 'Not processed') { $skipped++ }
        }
        # A room shows the organizer's name as subject: once the organizer or an attendee copy is found, its subject wins.
        if ($better -and $meeting.SubjectFromRoom) { $meeting.Subject = $better.Subject; $meeting.SubjectFromRoom = $false }
        $meeting.OrganizerCopy = if ($orgCopy) { 'Present' } elseif ($orgDeleted) { 'Mailbox deleted' } else { 'Absent' }
        if ($orgCopy) { $present++ }
        if ($roomsMode -and $meeting.Kind -eq 'Series') { $meeting.Scope = 'Occurrences' }
        if ($roomsMode) { $meeting.Occurrences = $occ.Count }
    }
    if ($meetings.Count) {
        Write-McoItem Ok ('{0:N0} cop{1} of {2:N0} meeting(s) {3} {4:N0} in rooms {3} organizer copy present for {5:N0} {3} {6:N0} attendee mailbox(es) searched' -f $total, $(if ($total -eq 1) { 'y' } else { 'ies' }), $meetings.Count, $dot, $roomCopies, $present, $uidsOf.Count) -Icon People
        if ($expanded) { Write-McoItem Info ('{0:N0} distribution group(s) expanded' -f $expanded) -Icon People }
        if ($skipped) { Write-McoItem Skip ('{0:N0} attendee(s) not processed: external, deleted or not reachable (listed in the report)' -f $skipped) }
    }
    $list = @($meetings.Values)
    if ($Request.Subject) {
        $pattern = if ($Request.Subject -match '[*?]') { $Request.Subject } else { "*$($Request.Subject)*" }
        $before = $list.Count
        $list = @($list | Where-Object { $_.Subject -like $pattern })
        Write-McoItem Info ("Subject '{0}': {1} of {2} meeting(s) kept" -f $Request.Subject, $list.Count, $before)
    }
    $list = @($list | Sort-Object Start, Subject)
    if ($roomsMode) {
        $organizers = @($list | Group-Object Organizer | ForEach-Object {
                $first = $_.Group[0]
                [pscustomobject]@{ Input = $_.Name; DisplayName = [string]$first.OrganizerName; PrimaryAddress = $_.Name; Addresses = @($_.Name); UserId = ''; Account = 'Unknown'
                    State = switch ($first.OrganizerCopy) { 'Present' { 'Mailbox' } 'Mailbox deleted' { 'NotInDirectory' } default { 'Unknown' } }
                    Detail = '{0} meeting(s) in the rooms {1} organizer copy {2}' -f $_.Count, $dot, ([string]$first.OrganizerCopy).ToLowerInvariant() }
            })
    }
    $result = [pscustomobject]@{
        Tool = 'Meeting Cleanup On-Prem'; Version = $script:ToolVersion; Action = 'Report'; Status = 'Completed'; Error = ''
        StartedUtc = $started.ToString('o'); CompletedUtc = [datetime]::UtcNow.ToString('o'); DurationSeconds = [Math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
        Request = [pscustomobject]@{ Mode = $Request.Mode; Organizer = @($Request.Organizer); Room = @($Request.Room); Start = $Request.Start.ToString('o'); End = $Request.End.ToString('o'); StartText = Format-McoDate $Request.Start $Settings.TimeZone; EndText = Format-McoDate $Request.End $Settings.TimeZone -PeriodEnd; Subject = $Request.Subject; MeetingId = @($Request.MeetingId); SearchIn = @($Request.SearchIn); TimeZone = (Get-McoTimeZone $Settings.TimeZone).Id }
        Tenant = 'Exchange Server On-Premises'; Organization = [string]$Settings.Mailbox.Split('@')[-1]; AppId = ''; AppName = "EWS $($script:Ews.Url)"; Organizers = @($organizers); Searched = [pscustomobject]$searched; Meetings = [Collections.Generic.List[object]]::new(); Warnings = $warnings; Counts = $null
    }
    foreach ($m in $list) { $result.Meetings.Add($m) }
    Update-McoResultCounts $result
    if ($warnings.Count) { $result.Status = 'Warning' }
    $result
}

function Add-McoMeetingDefaults {
    param([Parameter(Mandatory)]$Meeting)
    foreach ($name in 'Notes', 'Copies', 'Attendees') {
        if (-not $Meeting.PSObject.Properties[$name]) { $Meeting | Add-Member -NotePropertyName $name -NotePropertyValue ([Collections.Generic.List[object]]::new()) }
    }
    if (-not $Meeting.PSObject.Properties['Selected']) { $Meeting | Add-Member -NotePropertyName Selected -NotePropertyValue $true }
    if (-not $Meeting.PSObject.Properties['Status']) { $Meeting | Add-Member -NotePropertyName Status -NotePropertyValue 'Found' }
}

function Update-McoResultCounts {
    <# The counters of the console summary and of the HTML report (same names as Meeting Cleanup; compiled, one pass over the copies). #>
    param([Parameter(Mandatory)]$Result)
    $Result.Counts = [MeetingCleanupOnPremNative.Fast]::Counts($Result.Meetings)
}

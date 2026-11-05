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
    $when = if ($isOccurrence) { ([datetime]$Event.Start).ToString('o') } else { '' }
    [pscustomobject]@{
        MeetingId = $Event.MeetingId; Mailbox = ([string]$Event.Mailbox).ToLowerInvariant(); Role = $Role; Via = $Via
        EventId = $Event.EventId; Subject = $Event.Subject; Response = $Event.Response; ShowAs = ''
        Cancelled = [bool]$Event.IsCancelled; Action = ''; Result = $Result; HttpStatus = 0; Detail = $Detail; Verified = ''; ActionUtc = ''
        # OccurrenceKey: the slot of the occurrence (its start, UTC), the same in every copy; the key of SkippedOccurrences.
        Occurrence = $when; OccurrenceStart = $when; OccurrenceKey = $when
        SeriesId = [string](Get-McoProperty $Event 'SeriesId'); Event = $Event
    }
}

function New-McoPlaceholderCopy {
    <# A row of the report for an attendee that has no copy to act on (external, not found, group expanded...). #>
    param([Parameter(Mandatory)]$Meeting, [Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$Role, [Parameter(Mandatory)][string]$Via, [Parameter(Mandatory)][string]$Result, [Parameter(Mandatory)][string]$Detail)
    [pscustomobject]@{
        MeetingId = $Meeting.MeetingId; Mailbox = $Mailbox; Role = $Role; Via = $Via; EventId = ''; Subject = $Meeting.Subject; Response = ''; ShowAs = ''
        Cancelled = $false; Action = ''; Result = $Result; HttpStatus = 0; Detail = $Detail; Verified = ''; ActionUtc = ''; Occurrence = ''; OccurrenceStart = ''; OccurrenceKey = ''; SeriesId = ''; Event = $null
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
        SkippedOccurrences = @()
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

function Test-McoSeriesInPeriod {
    <#
        Whether every occurrence of a series falls in the period: a numbered series by its count of occurrences, a
        series with an end date by its dates; a series without end never.
    #>
    param([AllowEmptyString()][string]$RecurrenceXml, [int]$Count, [datetime]$Start, [datetime]$End, [AllowEmptyString()][string]$TimeZone)
    if (-not $RecurrenceXml) { return $false }
    $numbered = [regex]::Match($RecurrenceXml, 'NumberOfOccurrences>\s*(\d+)\s*<')
    if ($numbered.Success) { return $Count -ge [int]$numbered.Groups[1].Value }
    if ($RecurrenceXml -notmatch 'EndDateRecurrence') { return $false }
    $first = [regex]::Match($RecurrenceXml, 'StartDate>\s*(\d{4}-\d{2}-\d{2})')
    $last = [regex]::Match($RecurrenceXml, 'EndDate>\s*(\d{4}-\d{2}-\d{2})')
    if (-not $first.Success -or -not $last.Success) { return $false }
    $zone = Get-McoTimeZone $TimeZone
    $from = [TimeZoneInfo]::ConvertTimeFromUtc($Start.ToUniversalTime(), $zone).Date
    $to = [TimeZoneInfo]::ConvertTimeFromUtc($End.ToUniversalTime().AddSeconds(-1), $zone).Date
    $inv = [Globalization.CultureInfo]::InvariantCulture
    [datetime]::ParseExact($first.Groups[1].Value, 'yyyy-MM-dd', $inv) -ge $from -and [datetime]::ParseExact($last.Groups[1].Value, 'yyyy-MM-dd', $inv) -le $to
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
    # A series acted on by its occurrences in the period: always in rooms mode, on demand for organizers.
    $byOccurrence = -not $roomsMode -and [string](Get-McoProperty $Request 'SeriesScope') -eq 'Occurrences'
    $occurrencesMode = $roomsMode -or $byOccurrence
    $directory = [string]$Settings.DirectoryMode -ne 'None'
    Write-McoNextStep $(if ($roomsMode) { "Rooms ($(@($Request.Room).Count))" } else { 'Organizers' }) $(if ($roomsMode) { 'Room' } else { 'User' })
    $organizers = @(if (-not $roomsMode) { Resolve-McoOrganizer -Identity $Request.Organizer -Settings $Settings })
    foreach ($o in $organizers) { Write-McoItem $(if ($o.State -eq 'Mailbox') { 'Ok' } else { 'Info' }) ("{0} {1} {2}{3}" -f $(if ($o.DisplayName) { "$($o.DisplayName) <$($o.PrimaryAddress)>" } else { $o.Input }), $dot, $o.Detail, $(if (@($o.Addresses).Count -gt 1) { " $dot $(@($o.Addresses).Count) addresses compared" } else { '' })) -Icon User }
    if ($roomsMode) { foreach ($r in @($Request.Room)) { Write-McoItem Info $r -Icon Room } }
    if ($byOccurrence) { Write-McoItem Info 'A series: only its occurrences in the period are acted on (a period of one day: one occurrence); the series goes on outside the period.' }
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
        try { $events = Get-McoMailboxEvents -Mailbox $mb.Address -StartUtc $Request.Start -EndUtc $Request.End -PageSize ([int]$Settings.PageSize) -Occurrences:$occurrencesMode; $searched.Read++ }
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
            foreach ($hit in @(Get-McoMailboxEvents -Mailbox $target -StartUtc $Request.Start -EndUtc $Request.End -PageSize ([int]$Settings.PageSize) -Uid @($uidsOf[$target]) -Occurrences:$occurrencesMode)) {
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
    $organizerUnread = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in $actions) {
        $meeting = $a.Meeting; $target = $a.Target; $role = $a.Role
        if ($withCopy[$meeting.MeetingId].Contains($target)) { continue }
        if ($a.Kind -eq 'External') { $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $target -Role $role -Via $a.Via -Result 'Not processed' -Detail 'not in an accepted domain of the organization (external)')); continue }
        if ($a.Kind -eq 'X500') { $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $target -Role $role -Via $a.Via -Result 'Not processed' -Detail 'X500 address: the mailbox no longer exists')); continue }
        $f = $found[$target]
        if ($f.Error) {
            $why = if ($f.Error -match 'ErrorNonExistentMailbox|ErrorMailboxNotFound|ErrorInvalidSmtpAddress') { 'no mailbox with this address (deleted or not a mailbox)' } else { $f.Error }
            $meeting.Copies.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $target -Role $role -Via $a.Via -Result 'Not processed' -Detail $why))
            if ($role -eq 'Organizer' -and $why -ne 'no mailbox with this address (deleted or not a mailbox)') { [void]$organizerUnread.Add($meeting.MeetingId) }
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
    # ---- 3b. series by occurrences (organizers): the occurrences of the period in the organizer's calendar, or in
    # every copy when the organizer has none (deleted mailbox); the copies of other occurrences are left out. ----------
    $leftSeries = @{}
    if ($byOccurrence) {
        foreach ($meeting in @($meetings.Values)) {
            if ($meeting.Kind -ne 'Series') { continue }
            if ($organizerUnread.Contains($meeting.MeetingId)) { $leftSeries[$meeting.MeetingId] = 'the occurrences of its organizer could not be read'; continue }
            $fromOrganizer = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            $fromAll = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            foreach ($c in $meeting.Copies) {
                if (-not $c.EventId -or -not $c.OccurrenceKey) { continue }
                [void]$fromAll.Add([string]$c.OccurrenceKey)
                if ($c.Role -eq 'Organizer') { [void]$fromOrganizer.Add([string]$c.OccurrenceKey) }
            }
            $keys = $fromAll
            if ($fromOrganizer.Count) { $keys = $fromOrganizer }
            if (-not $keys.Count) { $leftSeries[$meeting.MeetingId] = 'no copy whose occurrences could be read'; continue }
            # A mailbox keeps the occurrences that count; one left with none of them is said (Not found), once.
            $holding = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($c in $meeting.Copies) { if ($c.EventId -and (-not $c.OccurrenceKey -or $keys.Contains([string]$c.OccurrenceKey))) { [void]$holding.Add("$($c.Mailbox)|$($c.Role)") } }
            $list = [Collections.Generic.List[object]]::new()
            $said = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($c in $meeting.Copies) {
                if (-not $c.EventId -or -not $c.OccurrenceKey -or $keys.Contains([string]$c.OccurrenceKey)) { $list.Add($c); continue }
                $who = "$($c.Mailbox)|$($c.Role)"
                if (-not $holding.Contains($who) -and $said.Add($who)) { $list.Add((New-McoPlaceholderCopy -Meeting $meeting -Mailbox $c.Mailbox -Role $c.Role -Via $c.Via -Result 'Not found' -Detail 'no occurrence of the period in this mailbox (declined, removed or moved)')) }
            }
            $meeting.Copies.Clear()
            foreach ($c in $list) { $meeting.Copies.Add($c) }
        }
        # Never the attendees without their organizer, never the whole series: a series whose occurrences cannot be
        # known is left as it is.
        foreach ($id in $leftSeries.Keys) {
            foreach ($c in $meetings[$id].Copies) { if ($c.EventId) { $c.EventId = ''; $c.Result = 'Not processed'; $c.Detail = "the series is left as it is: $($leftSeries[$id])" } }
            $meetings[$id].Notes.Add("Not acted on: $($leftSeries[$id]). Acting on the attendees only, or on the whole series, would not be right.")
        }
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
            if ($c.EventId -and $c.Occurrence) { [void]$occ.Add([string]$c.Occurrence) }
            if ($c.Result -eq 'Not processed') { $skipped++ }
        }
        # A room shows the organizer's name as subject: once the organizer or an attendee copy is found, its subject wins.
        if ($better -and $meeting.SubjectFromRoom) { $meeting.Subject = $better.Subject; $meeting.SubjectFromRoom = $false }
        $meeting.OrganizerCopy = if ($orgCopy) { 'Present' } elseif ($orgDeleted) { 'Mailbox deleted' } elseif ($organizerUnread.Contains($meeting.MeetingId) -and $leftSeries.Contains($meeting.MeetingId)) { 'Not read' } else { 'Absent' }
        if ($orgCopy) { $present++ }
        if ($occurrencesMode -and $meeting.Kind -eq 'Series') {
            $meeting.Scope = 'Occurrences'
            if (-not $leftSeries.Contains($meeting.MeetingId)) {
                # Asked for by occurrence, every occurrence in the period: each one is still acted on (said, not changed).
                if ($byOccurrence -and $occ.Count) {
                    $reference = Get-McoBestCopy $meeting -WithEvent
                    if ($reference -and (Test-McoSeriesInPeriod -RecurrenceXml ([MeetingCleanupOnPremNative.Fast]::Text($reference.Event, 'RecurrenceXml')) -Count $occ.Count -Start $Request.Start -End $Request.End -TimeZone $Settings.TimeZone)) {
                        $meeting.Notes.Add('Every occurrence of the series is in the period: each one is acted on separately (one cancellation each with Cancel); -SeriesScope Whole acts on the series at once.')
                    }
                }
                $meeting.Notes.Add(('{0} occurrence(s) in the period{1}: only they are acted on, the series goes on outside the period.' -f $occ.Count, $(if ($roomsMode) { ' in the rooms searched' } else { '' })))
            }
        }
        if ($occurrencesMode) { $meeting.Occurrences = $occ.Count }
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
    if ($occurrencesMode) {
        $limited = 0; $occurrenceCopies = 0; $leftCount = 0
        foreach ($m in $list) {
            if ($m.Scope -ne 'Occurrences') { continue }
            if ($leftSeries.Contains($m.MeetingId)) { $leftCount++; continue }
            $limited++
            foreach ($c in $m.Copies) { if ($c.EventId -and $c.Occurrence) { $occurrenceCopies++ } }
        }
        if ($limited -or $leftCount) { Write-McoItem Info ('{0} series limited to their occurrences in the period ({1:N0} occurrence copies){2}' -f $limited, $occurrenceCopies, $(if ($leftCount) { " $dot $leftCount left as they are (occurrences not read)" } else { '' })) -Icon Calendar }
        if ($leftCount) { $text = '{0} series left as they are: the occurrences of their organizer could not be read (see the notes).' -f $leftCount; $warnings.Add($text); Write-McoItem Warn $text }
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
        Request = [pscustomobject]@{ Mode = $Request.Mode; Organizer = @($Request.Organizer); Room = @($Request.Room); Start = $Request.Start.ToString('o'); End = $Request.End.ToString('o'); StartText = Format-McoDate $Request.Start $Settings.TimeZone; EndText = Format-McoDate $Request.End $Settings.TimeZone -PeriodEnd; Subject = $Request.Subject; MeetingId = @($Request.MeetingId); SearchIn = @($Request.SearchIn); SeriesScope = $(if ($occurrencesMode) { 'Occurrences' } else { 'Whole' }); TimeZone = (Get-McoTimeZone $Settings.TimeZone).Id }
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
    # Reports of 1.0.0: the occurrences left out of a series (SkippedOccurrences), none.
    if (-not $Meeting.PSObject.Properties['SkippedOccurrences']) { $Meeting | Add-Member -NotePropertyName SkippedOccurrences -NotePropertyValue @() }
}

function Update-McoResultCounts {
    <# The counters of the console summary and of the HTML report (same names as Meeting Cleanup; compiled, one pass over the copies). #>
    param([Parameter(Mandatory)]$Result)
    $Result.Counts = [MeetingCleanupOnPremNative.Fast]::Counts($Result.Meetings)
}

# SIG # Begin signature block
# MIImIwYJKoZIhvcNAQcCoIImFDCCJhACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAn5MLaozoOdMTk
# Q8wHBjKVcdQSsK7UQf0QS9Bf15odSKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDL/V5f8zYp
# 2Q5VktRH0tRe8zhWx+DkFt5dG7+l8BoSvzANBgkqhkiG9w0BAQEFAASCAgAs5QqQ
# T9Okat8i2w03WVRhdZnn4VOcSH3kj2i8zzDZgKukuePgv/DV8IL1lU7u+DpGVeeJ
# B9Qtw2RJl8zqouwAC0CurHqBy31kzDpDkIKd29sKulvYNQ4zfsdJBAL50SgYba+O
# VsmUjI+5zsFXQkzgXb2F+xlOqXJbdcdmEIUCAa0B8lTPb6nXK6xdqp3CmaXrRWZc
# EScs2FETo9bYVwAWBRLHFdqQ2+vUeT90xPz9SEr4wB0gIJKUDhW1tlS8IuRQvAta
# 2Dx/fjVgHQ6ZmT1clB6hr0nDgVNMReJiWbogBNlPQBarxKM/TqCwu/aoFUNrqy8L
# F6A/M3jGslHnwgqMwpb/8e9PEIu551ICNM6qiHs+S4ypoRrdE0x3GwBZw94zpzDf
# c6wrIECZPMUWyjY6NyC0Cmp9rv8D8Xcup2oZ6BxCo/eNITmxti0daChkQTUxsucx
# 3CgyTS/AeayoZZ3Kio55j9OJvdbEWdXEeLldcm+5J0DvkzMNfPX4VYjLKbSgRcP7
# Njzfn3doeSRnvPFC05NhMpWbU6518orv23V6PAdl8s++XiKAtiuHsmnWbyMIChEy
# 4BzNNqpOp2mFKIpf/w48qNq/dPhIlTHH8ilvTeNcT1Vpmoq1avAe62xPtOIets77
# F5pDdVzqJUpZ9/565biI869kXzyIdfGeVxZpcKGCFMcwghTDBgorBgEEAYI3AwMB
# MYIUszCCFK8GCSqGSIb3DQEHAqCCFKAwghScAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEICGw/XdIt9ixmq3Sbhqkp0r1+8XJeV+m0pbwP2GAwAhnAghqEAuwVIVe
# ghgPMjAyNjExMDUxNzIyNDFaMAMCAQECCAWYizRhi1x3oIIRpzCCBPMwggLboAMC
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
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJX
# MIICUwIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDUxNzIyNDFaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCBNITfIcHhRu6hp
# oBHrK8Qa6Iczg2zVtJVJfBry/3J+KzCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEgw
# RgIhAPX6rhTb/BhhB6H/jtD+EhyNzDvmsIKZ0WanZN6LSJNkAiEAkxIs72FpioRQ
# HKXt6CclrydStwWks2RTngZ6op4A51U=
# SIG # End signature block

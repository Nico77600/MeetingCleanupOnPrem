<#
    Meeting Cleanup On-Prem - simulated Exchange Server for the tests (no server needed).

    An in-memory store of calendars answers the EWS operations the tool sends (GetFolder, FindItem CalendarView and
    Recoverable Items, GetItem one or many items, DeleteItem, CreateItem cancel / new meeting, MoveItem) with the
    SOAP XML of Exchange Server. Install-FakeEws replaces Invoke-McoEws in the module loaded (any version), so a
    whole search, cleanup, restore or transfer runs offline; Store.Calls counts the calls by operation.
#>

$script:FakeT = 'http://schemas.microsoft.com/exchange/services/2006/types'
$script:FakeM = 'http://schemas.microsoft.com/exchange/services/2006/messages'

function New-FakeStore {
    @{
        Mailboxes = @{}; Deleted = @{}; Calls = [Collections.Generic.Dictionary[string, int]]::new(); Next = 0
        Names = @{}; LatencyMs = 0; RecoverablePage = 0; Busy = $null; DenyMailbox = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    }
}

function Get-FakeCalendar {
    param([hashtable]$Store, [string]$Mailbox)
    $key = $Mailbox.ToLowerInvariant()
    if (-not $Store.Mailboxes.ContainsKey($key)) { $Store.Mailboxes[$key] = [Collections.Generic.List[hashtable]]::new() }
    , $Store.Mailboxes[$key]
}

function New-FakeId { param([hashtable]$Store, [string]$Prefix = 'AAMk') $Store.Next++; '{0}{1:D6}=' -f $Prefix, $Store.Next }

function Add-FakeMeeting {
    <#
        A meeting in the calendars of its organizer, attendees and rooms (each its own item ID, same UID). -Weeks: a
        weekly series of that many occurrences (a master and its occurrences in each calendar). A room shows the
        organizer's name as subject, as Exchange does by default.
    #>
    param(
        [hashtable]$Store, [string]$Subject, [string]$Organizer, [string[]]$Attendees = @(), [string[]]$Optional = @(), [string[]]$Rooms = @(),
        [datetime]$Start, [int]$Minutes = 60, [int]$Weeks = 0, [string[]]$Missing = @(), [switch]$NoOrganizerCopy
    )
    $uid = ('040000008200E00074C5B7101A82E008{0:D8}' -f ($Store.Next + 1)).ToUpperInvariant()
    $start = [datetime]::SpecifyKind($Start, [DateTimeKind]::Utc)
    $recurrence = if ($Weeks) {
        '<t:Recurrence xmlns:t="{0}"><t:WeeklyRecurrence><t:Interval>1</t:Interval><t:DaysOfWeek>{1}</t:DaysOfWeek></t:WeeklyRecurrence><t:NumberedRecurrence><t:StartDate>{2:yyyy-MM-dd}Z</t:StartDate><t:NumberOfOccurrences>{3}</t:NumberOfOccurrences></t:NumberedRecurrence></t:Recurrence>' -f $script:FakeT, $start.DayOfWeek, $start, $Weeks
    } else { '' }
    $holders = @(@{ Address = $Organizer; Response = 'Organizer'; Subject = $Subject })
    if ($NoOrganizerCopy) { $holders = @() }
    foreach ($a in @($Attendees) + @($Optional)) { if ($a) { $holders += @{ Address = $a; Response = 'Accept'; Subject = $Subject } } }
    foreach ($r in $Rooms) { if ($r) { $holders += @{ Address = $r; Response = 'Accept'; Subject = $(if ($Store.Names[$Organizer]) { $Store.Names[$Organizer] } else { $Organizer }) } } }
    foreach ($h in $holders) {
        if ($Missing -contains $h.Address) { continue }
        $calendar = Get-FakeCalendar $Store $h.Address
        $base = @{
            Uid = $uid; Subject = $h.Subject; Class = 'IPM.Appointment'; Org = $Organizer; OrgName = [string]$Store.Names[$Organizer]; Response = $h.Response; Cancelled = $false
            Required = @($Attendees | Where-Object { $_ }); Optional = @($Optional | Where-Object { $_ }); Resources = @($Rooms | Where-Object { $_ }); Location = ($Rooms -join '; '); Body = "Agenda of $Subject"; Zone = 'Romance Standard Time'
            Created = [datetime]::UtcNow.AddDays(-10); Modified = [datetime]::UtcNow.AddDays(-10); Mailbox = $h.Address.ToLowerInvariant()
        }
        if ($Weeks) {
            $master = $base.Clone(); $master.Id = New-FakeId $Store; $master.Type = 'RecurringMaster'; $master.Start = $start; $master.End = $start.AddMinutes($Minutes); $master.Recurrence = $recurrence
            $calendar.Add($master)
            for ($w = 0; $w -lt $Weeks; $w++) {
                $o = $base.Clone(); $o.Id = New-FakeId $Store 'OCC'; $o.Type = 'Occurrence'; $o.MasterId = $master.Id
                $o.Start = $start.AddDays(7 * $w); $o.End = $o.Start.AddMinutes($Minutes); $o.Recurrence = ''
                $calendar.Add($o)
            }
        }
        else {
            $item = $base.Clone(); $item.Id = New-FakeId $Store; $item.Type = 'Single'; $item.Start = $start; $item.End = $start.AddMinutes($Minutes); $item.Recurrence = ''
            $calendar.Add($item)
        }
    }
    $uid
}

function ConvertTo-FakeText { param([string]$Text) [Security.SecurityElement]::Escape([string]$Text) }

function ConvertTo-FakeItemXml {
    param([hashtable]$Item, [switch]$Summary)
    $date = { param($d) ([datetime]$d).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append(('<t:CalendarItem><t:ItemId Id="{0}" ChangeKey="CK{1}"/>' -f $Item.Id, $Item.Id.Length))
    [void]$sb.Append(('<t:Subject>{0}</t:Subject><t:ItemClass>{1}</t:ItemClass>' -f (ConvertTo-FakeText $Item.Subject), $Item.Class))
    [void]$sb.Append(('<t:Start>{0}</t:Start><t:End>{1}</t:End><t:Location>{2}</t:Location>' -f (& $date $Item.Start), (& $date $Item.End), (ConvertTo-FakeText $Item.Location)))
    [void]$sb.Append(('<t:CalendarItemType>{0}</t:CalendarItemType><t:UID>{1}</t:UID>' -f $Item.Type, $Item.Uid))
    if ($Summary) { [void]$sb.Append('</t:CalendarItem>'); return $sb.ToString() }
    [void]$sb.Append(('<t:Body BodyType="Text">{0}</t:Body><t:DateTimeCreated>{1}</t:DateTimeCreated><t:LastModifiedTime>{2}</t:LastModifiedTime>' -f (ConvertTo-FakeText $Item.Body), (& $date $Item.Created), (& $date $Item.Modified)))
    [void]$sb.Append(('<t:IsCancelled>{0}</t:IsCancelled><t:MyResponseType>{1}</t:MyResponseType>' -f $(if ($Item.Cancelled) { 'true' } else { 'false' }), $Item.Response))
    [void]$sb.Append(('<t:Organizer><t:Mailbox><t:Name>{0}</t:Name><t:EmailAddress>{1}</t:EmailAddress><t:RoutingType>SMTP</t:RoutingType></t:Mailbox></t:Organizer>' -f (ConvertTo-FakeText $Item.OrgName), $Item.Org))
    foreach ($kind in @(@('RequiredAttendees', $Item.Required), @('OptionalAttendees', $Item.Optional), @('Resources', $Item.Resources))) {
        if (-not @($kind[1]).Count) { continue }
        [void]$sb.Append("<t:$($kind[0])>")
        foreach ($a in @($kind[1])) { [void]$sb.Append(('<t:Attendee><t:Mailbox><t:Name>{0}</t:Name><t:EmailAddress>{0}</t:EmailAddress></t:Mailbox><t:ResponseType>Unknown</t:ResponseType></t:Attendee>' -f $a)) }
        [void]$sb.Append("</t:$($kind[0])>")
    }
    if ($Item.Recurrence) { [void]$sb.Append(($Item.Recurrence -replace ' xmlns:t="[^"]+"', '')) }
    [void]$sb.Append(('<t:StartTimeZone Id="{0}" Name="(UTC+01:00) Brussels, Copenhagen, Madrid, Paris"/>' -f $Item.Zone))
    [void]$sb.Append('</t:CalendarItem>')
    $sb.ToString()
}

function New-FakeEnvelope {
    param([string]$Operation, [string]$Messages)
    '<?xml version="1.0" encoding="utf-8"?><s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Header><h:ServerVersionInfo xmlns:h="{0}" MajorVersion="15" MinorVersion="2" Version="V2017_07_11"/></s:Header><s:Body><m:{1}Response xmlns:m="{2}" xmlns:t="{0}"><m:ResponseMessages>{3}</m:ResponseMessages></m:{1}Response></s:Body></s:Envelope>' -f $script:FakeT, $Operation, $script:FakeM, $Messages
}

function New-FakeMessage {
    param([string]$Operation, [string]$Code = 'NoError', [string]$Inner = '')
    $class = if ($Code -eq 'NoError') { 'Success' } else { 'Error' }
    $text = if ($Code -eq 'NoError') { '' } else { "<m:MessageText>$Code</m:MessageText>" }
    '<m:{0}ResponseMessage ResponseClass="{1}">{2}<m:ResponseCode>{3}</m:ResponseCode>{4}</m:{0}ResponseMessage>' -f $Operation, $class, $text, $Code, $Inner
}

function Find-FakeItem {
    param([hashtable]$Store, [string]$Mailbox, [string]$Id)
    foreach ($i in (Get-FakeCalendar $Store $Mailbox)) { if ($i.Id -eq $Id) { return $i } }
    $null
}

function Invoke-FakeEwsCore {
    <# The SOAP answer of the simulated Exchange Server to one request. #>
    param([hashtable]$Store, [string]$Operation, [string]$Body, [string]$Mailbox)
    $Store.Calls[$Operation] = 1 + $(if ($Store.Calls.ContainsKey($Operation)) { $Store.Calls[$Operation] } else { 0 })
    if ($Store.LatencyMs) { Start-Sleep -Milliseconds $Store.LatencyMs }
    $doc = [xml]('<x xmlns:m="{0}" xmlns:t="{1}">{2}</x>' -f $script:FakeM, $script:FakeT, $Body)
    $ns = [Xml.XmlNamespaceManager]::new($doc.NameTable); $ns.AddNamespace('m', $script:FakeM); $ns.AddNamespace('t', $script:FakeT)
    $mailbox = ([string]$Mailbox).ToLowerInvariant()
    if ($Store.DenyMailbox.Contains($mailbox)) { return New-FakeEnvelope $Operation (New-FakeMessage $Operation 'ErrorNonExistentMailbox') }
    switch ($Operation) {
        'GetFolder' { return New-FakeEnvelope $Operation (New-FakeMessage $Operation) }
        'FindItem' {
            $view = $doc.SelectSingleNode('//m:CalendarView', $ns)
            if ($view) {
                $from = [datetime]::Parse($view.GetAttribute('StartDate'), [Globalization.CultureInfo]::InvariantCulture, 'AdjustToUniversal')
                $to = [datetime]::Parse($view.GetAttribute('EndDate'), [Globalization.CultureInfo]::InvariantCulture, 'AdjustToUniversal')
                $max = [int]$view.GetAttribute('MaxEntriesReturned')
                # Exchange sorts the view by start and gives the first MaxEntriesReturned items (no offset).
                $all = @((Get-FakeCalendar $Store $mailbox) | Where-Object { $_.Type -ne 'RecurringMaster' -and $_.Start -lt $to -and $_.End -gt $from } | Sort-Object { $_.Start }, { $_.Id })
                $hits = @($all | Select-Object -First $max)
                $items = ($hits | ForEach-Object { ConvertTo-FakeItemXml $_ -Summary }) -join ''
                return New-FakeEnvelope $Operation (New-FakeMessage $Operation -Inner ('<m:RootFolder TotalItemsInView="{0}" IncludesLastItemInRange="{1}"><t:Items>{2}</t:Items></m:RootFolder>' -f $all.Count, $(if ($all.Count -le $max) { 'true' } else { 'false' }), $items))
            }
            # Recoverable Items (IndexedPageItemView): the items removed (SoftDelete: Deletions; Purges stays empty).
            $folder = $doc.SelectSingleNode('//t:DistinguishedFolderId', $ns).GetAttribute('Id')
            $paging = $doc.SelectSingleNode('//m:IndexedPageItemView', $ns)
            $offset = [int]$paging.GetAttribute('Offset'); $max = [int]$paging.GetAttribute('MaxEntriesReturned')
            if ($Store.RecoverablePage) { $max = [Math]::Min($max, [int]$Store.RecoverablePage) }
            $deleted = if ($folder -eq 'recoverableitemsdeletions' -and $Store.Deleted.ContainsKey($mailbox)) { @($Store.Deleted[$mailbox]) } else { @() }
            $hits = @($deleted | Select-Object -Skip $offset -First $max)
            $items = ($hits | ForEach-Object { ConvertTo-FakeItemXml $_ }) -join ''
            return New-FakeEnvelope $Operation (New-FakeMessage $Operation -Inner ('<m:RootFolder IndexedPagingOffset="{0}" TotalItemsInView="{1}" IncludesLastItemInRange="{2}"><t:Items>{3}</t:Items></m:RootFolder>' -f ($offset + $hits.Count), $deleted.Count, $(if ($offset + $hits.Count -ge $deleted.Count) { 'true' } else { 'false' }), $items))
        }
        'GetItem' {
            $messages = [Text.StringBuilder]::new()
            foreach ($node in $doc.SelectNodes('//m:ItemIds/*', $ns)) {
                $item = $null
                if ($node.LocalName -eq 'RecurringMasterItemId') {
                    $occ = Find-FakeItem $Store $mailbox $node.GetAttribute('OccurrenceId')
                    if ($occ -and $occ.MasterId) { $item = Find-FakeItem $Store $mailbox $occ.MasterId }
                }
                else { $item = Find-FakeItem $Store $mailbox $node.GetAttribute('Id') }
                if ($item) { [void]$messages.Append((New-FakeMessage $Operation -Inner ('<m:Items>{0}</m:Items>' -f (ConvertTo-FakeItemXml $item)))) }
                else { [void]$messages.Append((New-FakeMessage $Operation 'ErrorItemNotFound')) }
            }
            return New-FakeEnvelope $Operation $messages.ToString()
        }
        'DeleteItem' {
            $id = $doc.SelectSingleNode('//m:ItemIds/t:ItemId', $ns).GetAttribute('Id')
            $calendar = Get-FakeCalendar $Store $mailbox
            $item = Find-FakeItem $Store $mailbox $id
            if (-not $item) { return New-FakeEnvelope $Operation (New-FakeMessage $Operation 'ErrorItemNotFound') }
            [void]$calendar.Remove($item)
            if ($item.Type -eq 'RecurringMaster') {
                # The occurrences go with their series, and come back with it (MoveItem).
                $item.Occ = @($calendar | Where-Object { $_.MasterId -eq $item.Id })
                foreach ($o in $item.Occ) { [void]$calendar.Remove($o) }
            }
            if ($item.Type -ne 'Occurrence') {
                if (-not $Store.Deleted.ContainsKey($mailbox)) { $Store.Deleted[$mailbox] = [Collections.Generic.List[hashtable]]::new() }
                $item.Modified = [datetime]::UtcNow
                $Store.Deleted[$mailbox].Add($item)
            }
            return New-FakeEnvelope $Operation (New-FakeMessage $Operation)
        }
        'MoveItem' {
            $id = $doc.SelectSingleNode('//m:ItemIds/t:ItemId', $ns).GetAttribute('Id')
            $item = if ($Store.Deleted.ContainsKey($mailbox)) { @($Store.Deleted[$mailbox] | Where-Object { $_.Id -eq $id })[0] } else { $null }
            if (-not $item) { return New-FakeEnvelope $Operation (New-FakeMessage $Operation 'ErrorItemNotFound') }
            [void]$Store.Deleted[$mailbox].Remove($item)
            $item.Id = New-FakeId $Store
            $calendar = Get-FakeCalendar $Store $mailbox
            $calendar.Add($item)
            if ($item.ContainsKey('Occ')) { foreach ($o in @($item.Occ)) { $o.MasterId = $item.Id; $calendar.Add($o) }; $item.Remove('Occ') }
            return New-FakeEnvelope $Operation (New-FakeMessage $Operation -Inner ('<m:Items><t:CalendarItem><t:ItemId Id="{0}"/></t:CalendarItem></m:Items>' -f $item.Id))
        }
        'CreateItem' {
            $cancel = $doc.SelectSingleNode('//t:CancelCalendarItem/t:ReferenceItemId', $ns)
            if ($cancel) {
                $item = Find-FakeItem $Store $mailbox $cancel.GetAttribute('Id')
                if (-not $item) { return New-FakeEnvelope $Operation (New-FakeMessage $Operation 'ErrorItemNotFound') }
                $item.Cancelled = $true
                # The attendees receive the cancellation: their copies are cancelled too.
                foreach ($cal in $Store.Mailboxes.Values) { foreach ($i in $cal) { if ($i.Uid -eq $item.Uid) { $i.Cancelled = $true } } }
                return New-FakeEnvelope $Operation (New-FakeMessage $Operation -Inner '<m:Items/>')
            }
            $new = $doc.SelectSingleNode('//m:Items/t:CalendarItem', $ns)
            if ($new) {
                $get = { param($n) $x = $new.SelectSingleNode("t:$n", $ns); if ($x) { $x.InnerText } else { '' } }
                $addresses = { param($n) @($new.SelectNodes("t:$n/t:Attendee/t:Mailbox/t:EmailAddress", $ns) | ForEach-Object { $_.InnerText.ToLowerInvariant() }) }
                $name = if ($Store.Names[$mailbox]) { $Store.Names[$mailbox] } else { $mailbox }
                $Store.Names[$mailbox] = $name
                $start = [datetime]::Parse((& $get 'Start'), [Globalization.CultureInfo]::InvariantCulture, 'AdjustToUniversal')
                $end = [datetime]::Parse((& $get 'End'), [Globalization.CultureInfo]::InvariantCulture, 'AdjustToUniversal')
                $weeks = 0
                $recurrenceNode = $new.SelectSingleNode('t:Recurrence', $ns)
                if ($recurrenceNode) { $count = $recurrenceNode.SelectSingleNode('.//t:NumberOfOccurrences', $ns); $weeks = if ($count) { [int]$count.InnerText } else { 4 } }
                $before = $Store.Next
                $uid = Add-FakeMeeting -Store $Store -Subject (& $get 'Subject') -Organizer $mailbox -Attendees (& $addresses 'RequiredAttendees') -Optional (& $addresses 'OptionalAttendees') -Rooms (& $addresses 'Resources') -Start $start -Minutes ([int]($end - $start).TotalMinutes) -Weeks $weeks
                $created = @((Get-FakeCalendar $Store $mailbox) | Where-Object { $_.Uid -eq $uid -and $_.Type -ne 'Occurrence' })[0]
                return New-FakeEnvelope $Operation (New-FakeMessage $Operation -Inner ('<m:Items><t:CalendarItem><t:ItemId Id="{0}" ChangeKey="CK"/></t:CalendarItem></m:Items>' -f $created.Id))
            }
            return New-FakeEnvelope $Operation (New-FakeMessage $Operation)
        }
        default { return New-FakeEnvelope $Operation (New-FakeMessage $Operation 'ErrorInvalidRequest') }
    }
}

function Install-FakeEws {
    <# Replaces Invoke-McoEws of the module loaded by the simulated server, and marks EWS as connected. #>
    param([hashtable]$Store, [hashtable]$Settings)
    $m = Get-Module MeetingCleanupOnPrem
    # The simulated server runs where it was loaded (a test script, a session): the module calls it through a script block.
    $core = ${function:Invoke-FakeEwsCore}
    & $m {
        param($s, $settings, $core)
        $script:FakeStore = $s
        $script:FakeCore = $core
        $script:Ews = @{ Client = $null; Url = $(if ($settings.EwsUrl) { $settings.EwsUrl } else { 'https://mail.contoso.test/EWS/Exchange.asmx' }); Credential = $null; Settings = $settings; Mailbox = $settings.Mailbox; ConnectedUtc = [datetime]::UtcNow }
        # Store.Busy: @{ <operation> = <count> } answers ErrorServerBusy (throttling) that many times first.
        $script:FakeAnswer = {
            param([string]$Operation, [string]$Body, [string]$Mailbox)
            $busy = $script:FakeStore['Busy']
            if ($busy -and [int]$busy[$Operation] -gt 0) {
                $busy[$Operation] = [int]$busy[$Operation] - 1
                $script:FakeStore.Calls[$Operation] = 1 + $(if ($script:FakeStore.Calls.ContainsKey($Operation)) { $script:FakeStore.Calls[$Operation] } else { 0 })
                $fault = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><s:Fault><faultcode>a:ErrorServerBusy</faultcode><faultstring>The server cannot service this request right now. Try again later.</faultstring><detail><e:ResponseCode xmlns:e="http://schemas.microsoft.com/exchange/services/2006/errors">ErrorServerBusy</e:ResponseCode><e:Message xmlns:e="http://schemas.microsoft.com/exchange/services/2006/errors">The server cannot service this request right now. Try again later.</e:Message><t:MessageXml xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><t:Value Name="BackOffMilliseconds">20</t:Value></t:MessageXml></detail></s:Fault></s:Body></s:Envelope>'
                return ConvertFrom-McoEwsResponse -HttpStatus 500 -Text $fault
            }
            ConvertFrom-McoEwsResponse -HttpStatus 200 -Text (& $script:FakeCore -Store $script:FakeStore -Operation $Operation -Body $Body -Mailbox $Mailbox)
        }
        if (Get-Command -Name Send-McoEwsRequest -Module MeetingCleanupOnPrem -ErrorAction SilentlyContinue) {
            # The transport only: Invoke-McoEws (its retries) runs as with a real server.
            function script:Send-McoEwsRequest {
                param([Parameter(Mandatory)][string]$Operation, [Parameter(Mandatory)][string]$Body, [Parameter(Mandatory)][string]$Mailbox)
                & $script:FakeAnswer $Operation $Body $Mailbox
            }
        }
        else {
            # Versions without Send-McoEwsRequest (0.4.1 and before, to compare versions).
            function script:Invoke-McoEws {
                param([Parameter(Mandatory)][string]$Operation, [Parameter(Mandatory)][string]$Body, [string]$Mailbox)
                $target = if ($Mailbox) { $Mailbox } else { [string]$script:Ews.Mailbox }
                & $script:FakeAnswer $Operation $Body $target
            }
        }
    } $Store $Settings $core
}

function New-FakeTenant {
    <#
        The organization of the tests: org@contoso.test organizes M1 (single: two attendees, a room, an external),
        S1 (weekly series of 4: an attendee, a room) and M2 (single, an attendee whose mailbox is gone); other@
        organizes O1 (with att1, not searched); org has an appointment without attendees. -Extra: that many more
        meetings of org (for the measures), each with -ExtraAttendees attendees from a pool of 40 and a room.
    #>
    param([int]$Extra = 0, [int]$ExtraAttendees = 6, [datetime]$From = [datetime]'2030-01-07T09:00:00Z')
    $s = New-FakeStore
    $s.Names['org@contoso.test'] = 'Org Anizer'; $s.Names['other@contoso.test'] = 'Other Person'; $s.Names['new@contoso.test'] = 'New Organizer'
    $s.Ids = @{}
    $s.Ids.M1 = Add-FakeMeeting $s 'M1 Budget review' 'org@contoso.test' @('att1@contoso.test', 'att2@contoso.test', 'ext@fabrikam.test') -Rooms @('room1@contoso.test') -Start $From.AddDays(1)
    $s.Ids.S1 = Add-FakeMeeting $s 'S1 Weekly sync' 'org@contoso.test' @('att1@contoso.test') -Rooms @('room2@contoso.test') -Start $From.AddDays(2) -Weeks 4
    $s.Ids.M2 = Add-FakeMeeting $s 'M2 Steering' 'org@contoso.test' @('att3@contoso.test', 'gone@contoso.test') -Rooms @('room1@contoso.test') -Start $From.AddDays(3) -Missing @('gone@contoso.test')
    $s.Ids.O1 = Add-FakeMeeting $s 'O1 Other meeting' 'other@contoso.test' @('att1@contoso.test') -Start $From.AddDays(4)
    [void](Add-FakeMeeting $s 'Dentist' 'org@contoso.test' -Start $From.AddDays(5))
    [void]$s.DenyMailbox.Add('gone@contoso.test')
    for ($i = 0; $i -lt $Extra; $i++) {
        $attendees = @(for ($k = 0; $k -lt $ExtraAttendees; $k++) { 'user{0:D2}@contoso.test' -f (($i * 7 + $k * 3) % 40) }) | Select-Object -Unique
        [void](Add-FakeMeeting $s ('X{0:D4} Project meeting' -f $i) 'org@contoso.test' $attendees -Rooms @('room{0}@contoso.test' -f (3 + $i % 5)) -Start $From.AddDays(1 + ($i % 300)).AddHours($i % 8))
    }
    $s
}

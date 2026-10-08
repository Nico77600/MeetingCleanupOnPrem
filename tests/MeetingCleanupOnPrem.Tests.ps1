#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.1.0' }

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    Import-Module (Join-Path $script:Root 'MeetingCleanupOnPrem.psd1') -Force
    $script:Module = Get-Module MeetingCleanupOnPrem
    & $script:Module { $script:Quiet = $true }
    . (Join-Path $PSScriptRoot 'FakeEws.ps1')

    function New-TestSettings {
        $s = Import-McoConfiguration
        $s.DirectoryMode = 'None'; $s.Rooms = @('room1@contoso.test', 'room2@contoso.test')
        $s.OutputPath = Join-Path $TestDrive 'reports'; $s.LogPath = Join-Path $TestDrive 'logs'; $s.TimeZone = 'Romance Standard Time'; $s.RestoreMode = 'Ews'
        $s
    }

    function Find-Test {
        param([hashtable]$Settings, [hashtable]$Store, [string]$Organizer = 'org@contoso.test', [string[]]$Room)
        Install-FakeEws -Store $Store -Settings $Settings
        $request = if ($Room) { New-McoRequest -Settings $Settings -Room $Room -Start ([datetime]'2030-01-01') -End ([datetime]'2030-01-20') }
            else { New-McoRequest -Settings $Settings -Organizer $Organizer -Start ([datetime]'2030-01-01') -End ([datetime]'2030-12-31') }
        Find-McoMeetings -Settings $Settings -Request $request
    }
}

Describe 'Configuration and request' {
    It 'loads the delivered on-prem configuration without contacting Exchange' {
        $settings = Import-McoConfiguration
        $settings.Authentication | Should -Be 'Windows'
        $settings.AccessMode | Should -Be 'Impersonation'
        $settings.PageSize | Should -Be 500
        $settings.Mailbox | Should -Be 'svc-meetingcleanup@contoso.test'
    }

    It 'rejects EXO native transfer and invalid connection values' {
        $settings = Import-McoConfiguration
        $settings.TransferMethod = 'Native'
        (Test-McoConfiguration $settings).Problems -join ' ' | Should -Match 'native organizer-transfer'
        $settings.TransferMethod = 'Recreate'
        $settings.PageSize = 2
        (Test-McoConfiguration $settings).Problems -join ' ' | Should -Match 'PageSize'
    }

    It 'accepts the optional Exchange Management Shell directory mode' {
        $settings = Import-McoConfiguration
        $settings.DirectoryMode | Should -Be 'Auto'
        $settings.DirectoryMode = 'invalid'
        (Test-McoConfiguration $settings).Problems -join ' ' | Should -Match 'DirectoryMode'
    }

    It 'validates the remote PowerShell connection mode' {
        $settings = Import-McoConfiguration
        $settings.ManagementShellMode = 'Rps'
        $settings.ManagementShellServer = ''
        $settings.ManagementShellUri = ''
        (Test-McoConfiguration $settings).Problems -join ' ' | Should -Match 'ManagementShell.ServerFqdn'
        $settings.ManagementShellMode = 'Existing'
        InModuleScope MeetingCleanupOnPrem -Parameters @{ Settings = $settings } -ScriptBlock {
            param($Settings)
            (Connect-McoExchangeShell -Settings $Settings).Mode | Should -Be 'Existing'
        }
    }

    It 'builds the same request shape for organizer and room modes' {
        $settings = Import-McoConfiguration
        $org = New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -Start ([datetime]'2030-01-01') -End ([datetime]'2030-01-02')
        $org.Mode | Should -Be 'Organizers'
        $org.SearchIn | Should -Contain 'Organizer'
        $rooms = New-McoRequest -Settings $settings -Room 'room@contoso.test' -Start ([datetime]'2030-01-01') -End ([datetime]'2030-01-02')
        $rooms.Mode | Should -Be 'Rooms'
        (Test-McoRequest $rooms).IsValid | Should -BeTrue
    }

    Describe 'Exchange Management Shell directory adapter' {
        It 'resolves a recipient and preserves SMTP and X500 aliases' {
            InModuleScope MeetingCleanupOnPrem {
                function Test-McoFakeGetRecipient {
                    param([string]$Identity)
                    [pscustomobject]@{
                        PrimarySmtpAddress = 'org@contoso.test'; DisplayName = 'Org'; Alias = 'org'
                        RecipientTypeDetails = 'UserMailbox'; LegacyExchangeDN = '/o=Contoso/ou=Exchange/cn=Recipients/cn=org'
                        EmailAddresses = @('SMTP:org@contoso.test', 'smtp:alias@contoso.test', 'x500:/o=Contoso/ou=Exchange/cn=Recipients/cn=org')
                    }
                }
                Mock Get-McoExchangeCommand { [pscustomobject]@{ Name = 'Test-McoFakeGetRecipient' } }
                $recipient = Resolve-McoExchangeRecipient -Identity 'alias@contoso.test'
                $recipient.PrimaryAddress | Should -Be 'org@contoso.test'
                $recipient.Addresses | Should -Contain 'alias@contoso.test'
                $recipient.Addresses | Should -Contain '/o=contoso/ou=exchange/cn=recipients/cn=org'
            }
        }

        It 'expands nested distribution groups without contacting EWS' {
            InModuleScope MeetingCleanupOnPrem {
                function Test-McoFakeGetDistributionGroupMember {
                    param([string]$Identity, [string]$ResultSize)
                    if ($Identity -eq 'dl@contoso.test') {
                        return [pscustomobject]@{ PrimarySmtpAddress = 'nested@contoso.test'; DisplayName = 'Nested'; RecipientTypeDetails = 'MailUniversalDistributionGroup' }
                    }
                    [pscustomobject]@{ PrimarySmtpAddress = 'member@contoso.test'; DisplayName = 'Member'; RecipientTypeDetails = 'UserMailbox' }
                }
                Mock Get-McoExchangeCommand { [pscustomobject]@{ Name = 'Test-McoFakeGetDistributionGroupMember' } }
                @((Get-McoExchangeGroupMembers -Identity 'dl@contoso.test').Address) | Should -Be @('member@contoso.test')
            }
        }

        It 'searches a deleted organizer (not in the directory) without opening its calendar' {
            InModuleScope MeetingCleanupOnPrem {
                function Test-McoFakeGetRecipient { param([string]$Identity) throw "The operation couldn't be performed because object '$Identity' couldn't be found on 'DC1.contoso.test'." }
                Mock Get-McoExchangeCommand { [pscustomobject]@{ Name = 'Test-McoFakeGetRecipient' } }
                $settings = Get-McoDefaultConfiguration
                $o = @(Resolve-McoOrganizer -Identity 'gone@contoso.test' -Settings $settings)
                $o[0].State | Should -Be 'NotInDirectory'
                $o[0].Detail | Should -Match 'not in the directory'
                $request = [pscustomobject]@{ Mode = 'Organizers'; SearchIn = @('Organizer'); Room = @(); RoomFile = ''; Mailboxes = @(); MailboxFile = '' }
                @((Get-McoSearchMailboxes -Settings $settings -Request $request -Organizers $o).Mailboxes).Count | Should -Be 0
            }
        }

        It 'lists every room mailbox with Exchange Management Shell for -SearchIn Rooms' {
            InModuleScope MeetingCleanupOnPrem {
                function Test-McoFakeGetMailbox {
                    param([string]$ResultSize, [string[]]$RecipientTypeDetails)
                    if ($RecipientTypeDetails -contains 'RoomMailbox') { [pscustomobject]@{ PrimarySmtpAddress = 'Room-A@contoso.test' }; [pscustomobject]@{ PrimarySmtpAddress = 'room-b@contoso.test' } }
                    else { [pscustomobject]@{ PrimarySmtpAddress = 'user@contoso.test' } }
                }
                Mock Get-McoExchangeCommand { [pscustomobject]@{ Name = 'Test-McoFakeGetMailbox' } }
                $settings = Get-McoDefaultConfiguration
                $settings.Rooms = @('room-c@contoso.test')
                $request = [pscustomobject]@{ Mode = 'Organizers'; SearchIn = @('Rooms'); Room = @(); RoomFile = ''; Mailboxes = @(); MailboxFile = '' }
                $plan = Get-McoSearchMailboxes -Settings $settings -Request $request
                @($plan.Mailboxes | Sort-Object Address | ForEach-Object Address) | Should -Be @('room-a@contoso.test', 'room-b@contoso.test', 'room-c@contoso.test')
                @($plan.Mailboxes | Where-Object { -not $_.IsRoom }).Count | Should -Be 0
                $plan.Warnings | Should -BeNullOrEmpty
                # Search.AllRooms = $false: the rooms of the configuration only (an impersonation limited to some rooms).
                $settings.AllRooms = $false
                @((Get-McoSearchMailboxes -Settings $settings -Request $request).Mailboxes | ForEach-Object Address) | Should -Be @('room-c@contoso.test')
                $settings.DirectoryMode = 'None'; $settings.Rooms = @()
                (Get-McoSearchMailboxes -Settings $settings -Request $request).Warnings | Should -Match 'No room to search'
            }
        }

        It 'uses Exchange Management Shell to enumerate AllMailboxes when available' {
            InModuleScope MeetingCleanupOnPrem {
                function Test-McoFakeGetMailbox {
                    param([string]$ResultSize)
                    [pscustomobject]@{ PrimarySmtpAddress = 'one@contoso.test' }
                    [pscustomobject]@{ PrimarySmtpAddress = 'two@contoso.test' }
                }
                Mock Get-McoExchangeCommand { [pscustomobject]@{ Name = 'Test-McoFakeGetMailbox' } }
                $settings = Get-McoDefaultConfiguration
                $settings.DirectoryMode = 'ExchangePowerShell'
                $request = [pscustomobject]@{ Mode = 'Organizers'; SearchIn = @('AllMailboxes'); Room = @(); RoomFile = ''; Mailboxes = @(); MailboxFile = '' }
                $plan = Get-McoSearchMailboxes -Settings $settings -Request $request
                @($plan.Mailboxes | ForEach-Object Address) | Should -Be @('one@contoso.test', 'two@contoso.test')
                $plan.Warnings | Should -BeNullOrEmpty
            }
        }
    }
}

Describe 'EWS protocol layer' {
    It 'creates a CalendarView request with the mailbox anchor and calendar properties' {
        InModuleScope MeetingCleanupOnPrem {
            $body = New-McoFindCalendarBody -StartUtc ([datetime]'2030-01-01Z') -EndUtc ([datetime]'2030-01-02Z') -Mailbox 'org@contoso.test' -PageSize 100
            $body | Should -Match 'CalendarView'
            $body | Should -Match 'MaxEntriesReturned="100"'
            $body | Should -Match 'calendar:Start'
            $body | Should -Match 'calendar:CalendarItemType'
            $body | Should -Not -Match 'calendar:AppointmentType'
            (New-McoFindCalendarBody -StartUtc ([datetime]'2030-01-01Z') -EndUtc ([datetime]'2030-01-02Z') -Mailbox 'org@contoso.test' -PageSize 100 -Minimal) | Should -Not -Match 'calendar:UID'
            (New-McoDeleteItemBody -ItemId 'id1') | Should -Match 'DeleteType="SoftDelete" SendMeetingCancellations="SendToNone"'
            $body | Should -Match 'org@contoso.test'
            (New-McoCancelItemBody -ItemId 'id1' -Comment 'Cancelled' -ChangeKey 'ck1') | Should -Match '<t:ReferenceItemId Id="id1" ChangeKey="ck1"/>'
            $event = [pscustomobject]@{ Subject = 'Review'; Body = ''; Start = [datetime]'2030-01-01T10:00:00Z'; End = [datetime]'2030-01-01T11:00:00Z'; Location = 'Room'; RequiredAttendees = @('att@contoso.test'); Resources = @('room@contoso.test'); MeetingId = 'UID1' }
            $create = New-McoCreateAppointmentBody -Event $event -Comment 'Transfer'
            $create | Should -Match 'SendMeetingInvitations="SendToAllAndSaveCopy"'
            $create | Should -Match '<t:EmailAddress>room@contoso.test</t:EmailAddress>'
        }
    }

    It 'does not add a target mailbox to distinguished folders in Self mode' {
        InModuleScope MeetingCleanupOnPrem {
            (New-McoFolderIdXml -Distinguished 'calendar' -Mailbox 'other@contoso.test' -AccessMode Self) | Should -Be '<t:DistinguishedFolderId Id="calendar"/>'
        }
    }

    It 'reads a successful SOAP response without requiring a live Exchange server' {
        $xml = @'
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:GetFolderResponse><m:ResponseMessages><m:GetFolderResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode></m:GetFolderResponseMessage></m:ResponseMessages></m:GetFolderResponse>
  </soap:Body>
</soap:Envelope>
'@
        $answer = InModuleScope MeetingCleanupOnPrem -Parameters @{ Text = $xml } -ScriptBlock {
            param($Text)
            ConvertFrom-McoEwsResponse -HttpStatus 200 -Text $Text
        }
        $answer.ResponseClass | Should -Be 'Success'
        $answer.ResponseCode | Should -Be 'NoError'
    }
}

Describe 'Offline report generation' {
    It 'writes summary and CSV/HTML artifacts with the on-prem tool identity' {
        $settings = Import-McoConfiguration
        $copy = [pscustomobject]@{ MeetingId = 'UID1'; Mailbox = 'org@contoso.test'; Role = 'Organizer'; Via = 'Organizer calendar'; EventId = 'id1'; Subject = 'Review'; Response = 'Organizer'; ShowAs = 'Busy'; Cancelled = $false; Action = ''; Result = ''; HttpStatus = 0; Verified = ''; ActionUtc = ''; Detail = ''; Occurrence = '' }
        $meeting = [pscustomobject]@{ MeetingId = 'UID1'; Subject = 'Review'; Organizer = 'org@contoso.test'; OrganizerName = 'Org'; Kind = 'Single'; Scope = 'Whole'; Occurrences = 0; StartText = '2030-01-01 10:00'; EndText = '2030-01-01 11:00'; NextInPeriod = ''; Recurrence = ''; Location = ''; OrganizerCopy = 'Present'; Copies = @($copy); Attendees = @(); Cancelled = $false; Selected = $true; Status = 'Found'; NewOrganizer = ''; NewMeetingId = ''; TransferMethod = ''; Notes = @() }
        $result = [pscustomobject]@{ Tool = 'Meeting Cleanup On-Prem'; Version = '0.1.0'; Action = 'Report'; Status = 'Completed'; Error = ''; StartedUtc = [datetime]::UtcNow.ToString('o'); CompletedUtc = [datetime]::UtcNow.ToString('o'); DurationSeconds = 0; Request = [pscustomobject]@{ TimeZone = 'UTC' }; Tenant = 'Exchange Server On-Premises'; Organization = ''; AppId = ''; AppName = 'EWS'; Organizers = @([pscustomobject]@{ Input = 'org@contoso.test'; DisplayName = 'Org'; PrimaryAddress = 'org@contoso.test'; Addresses = @('org@contoso.test'); State = 'Unknown'; Detail = '' }); Searched = [pscustomobject]@{}; Meetings = @($meeting); Warnings = @(); Counts = [pscustomobject]@{ Meetings = 1; Copies = 1; Removed = 0; Cancelled = 0; Restored = 0; Failed = 0 } }
        $report = Export-McoReport -Result $result -OutputPath $TestDrive -Prefix 'Test' -Formats @('Csv', 'Html')
        Test-Path (Join-Path $report.Directory 'Test-Summary.json') | Should -BeTrue
        (Get-Content -Raw (Join-Path $report.Directory 'Test-Summary.json')) | Should -Match 'Meeting Cleanup On-Prem'
        Test-Path (Join-Path $report.Directory 'Test.html') | Should -BeTrue
    }
}

Describe 'Progress' {
    It 'tells the time left from the speed of the progress, once it can be told' {
        & $script:Module {
            $script:ProgressEta = $null
            $t0 = [datetime]'2030-01-01T10:00:00Z'
            Get-McoProgressEta -Fraction 0.10 -Text '186/1,858 mailboxes searched' -Now $t0 | Should -Be ''
            # 1 s later: too early to tell.
            Get-McoProgressEta -Fraction 0.20 -Text '372/1,858 mailboxes searched' -Now $t0.AddSeconds(1) | Should -Be ''
            # 4 s for 20 %: 14 s for the 70 % left.
            Get-McoProgressEta -Fraction 0.30 -Text '557/1,858 mailboxes searched' -Now $t0.AddSeconds(4) | Should -Be 'about 15 s left'
            # A slow answer (alone: 26 s): smoothed with the last figure brought forward (10 s), 18 s.
            Get-McoProgressEta -Fraction 0.31 -Text '576/1,858 mailboxes searched' -Now $t0.AddSeconds(8) | Should -Be 'about 20 s left'
            Get-McoProgressEta -Fraction 1.00 -Text '1,858/1,858 mailboxes searched' -Now $t0.AddSeconds(13) | Should -Be ''
            # Another progress (another label, or going back) starts again.
            Get-McoProgressEta -Fraction 0.50 -Text '10/20 attendee mailboxes searched' -Now $t0.AddSeconds(20) | Should -Be ''
            Get-McoProgressEta -Fraction 0.10 -Text '2/20 attendee mailboxes searched' -Now $t0.AddSeconds(30) | Should -Be ''
            $script:ProgressEta.From | Should -Be 0.10
            # Counts written the French way (narrow no-break space) are the same label.
            Get-McoProgressEta -Fraction 0.20 -Text ('1{0}000/5{0}000 copies removed' -f [char]0x202F) -Now $t0.AddSeconds(31) | Should -Be ''
            Get-McoProgressEta -Fraction 0.40 -Text ('2{0}000/5{0}000 copies removed' -f [char]0x202F) -Now $t0.AddSeconds(35) | Should -Be 'about 15 s left'
            # A new step starts without a progress.
            Write-McoStep -Number 2 -Total 6 -Title 'Search'
            $script:ProgressEta | Should -BeNullOrEmpty
        }
    }

    It 'rounds the time left as a person would say it' {
        & $script:Module {
            Format-McoTimeLeft 4 | Should -Be 'a few seconds left'
            Format-McoTimeLeft 21 | Should -Be 'about 25 s left'
            Format-McoTimeLeft 58 | Should -Be 'about 1 min left'
            Format-McoTimeLeft 89 | Should -Be 'about 1 min 30 s left'
            Format-McoTimeLeft 200 | Should -Be 'about 3 min 20 s left'
            Format-McoTimeLeft 600 | Should -Be 'about 10 min left'
            Format-McoTimeLeft 4000 | Should -Be 'about 1 h 07 min left'
            Format-McoDuration 75.4 | Should -Be '1 min 15 s'
        }
    }

    It 'writes the console lines to the log, and the progress to a window queue, without colours' {
        $log = Start-McoLog -Directory (Join-Path $TestDrive 'console-logs')
        try {
            & $script:Module {
                $q = [Collections.Concurrent.ConcurrentQueue[string[]]]::new()
                $script:Ui = @{ Queue = $q }
                try {
                    Write-McoItem Ok 'EWS connected' -Icon Server
                    Write-McoProgress 0.5 '1/2 mailboxes searched'
                }
                finally { $script:Ui = $null }
                $q.Count | Should -Be 2
                $line = $null; [void]$q.TryDequeue([ref]$line); $line[0] | Should -Be 'Ok'
                [void]$q.TryDequeue([ref]$line); $line[1] | Should -Be '0.500|1/2 mailboxes searched|'
            }
        }
        finally { Stop-McoLog }
        $text = Get-Content -Raw $log
        $text | Should -Match '\[OK   \] EWS connected'
        $text | Should -Not -Match ([char]27)
    }
}

Describe 'Compiled helpers' {
    It 'reads an EWS calendar item like the PowerShell parser did, property by property' {
        $xml = Get-Content -Raw (Join-Path $PSScriptRoot 'data\GetItem-Series.xml')
        $item = InModuleScope MeetingCleanupOnPrem -Parameters @{ Text = $xml } -ScriptBlock {
            param($Text)
            $answer = ConvertFrom-McoEwsResponse -HttpStatus 200 -Text $Text
            ConvertFrom-McoCalendarNode $answer.Xml.SelectSingleNode('//m:Items/*', $answer.Ns) $answer.Ns 'Org@Contoso.test'
        }
        @($item.PSObject.Properties.Name) -join ',' | Should -Be 'EventId,ChangeKey,Mailbox,Subject,ItemClass,MeetingId,Organizer,OrganizerName,Start,End,AppointmentType,Location,Response,IsCancelled,RequiredAttendees,OptionalAttendees,Resources,Body,DateTimeCreated,LastModifiedUtc,RecurrenceXml,TimeZoneId,SeriesId,Occurrences'
        $item.EventId | Should -Be 'AAMk1='
        $item.Mailbox | Should -Be 'org@contoso.test'
        $item.Subject | Should -Be 'Weekly & review'
        $item.MeetingId | Should -Be '040000008200E00074C5B7101A82E008'
        $item.Organizer | Should -Be 'org@contoso.test'
        $item.Start | Should -Be ([datetime]'2030-01-07T09:00:00Z').ToUniversalTime()
        $item.Start.Kind | Should -Be 'Utc'
        @($item.RequiredAttendees) | Should -Be @('a1@contoso.test')
        @($item.OptionalAttendees).Count | Should -Be 0
        $item.Location | Should -Be ''
        $item.ItemClass | Should -Be 'IPM.Appointment'
        $item.DateTimeCreated | Should -BeNullOrEmpty
        $item.LastModifiedUtc | Should -BeNullOrEmpty
        $item.IsCancelled | Should -BeFalse
        $item.RecurrenceXml | Should -Match '^<t:Recurrence xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">'
        InModuleScope MeetingCleanupOnPrem -Parameters @{ Xml = $item.RecurrenceXml } -ScriptBlock { param($Xml) Format-McoEwsRecurrence $Xml } | Should -Be 'every 2 weeks (Monday), from 2030-01-07 until 2030-06-30'
        # A copy of the object (an occurrence of a series in rooms mode) is independent of it.
        $copy = $item.PSObject.Copy(); $copy.EventId = 'OCC1'
        $item.EventId | Should -Be 'AAMk1='
    }

    It 'counts a result as the PowerShell code of 0.3.0 did' {
        $settings = New-TestSettings
        $result = Find-Test $settings (New-FakeTenant)
        $result.Meetings[0].Copies[0].Result = 'Removed'; $result.Meetings[0].Copies[0].Action = 'Remove'
        $result.Meetings[1].Copies[0].Result = 'Not found'; $result.Meetings[1].Copies[0].Action = 'Restore'
        $result.Meetings[2].Copies[0].Result = 'Not done'
        $result.Meetings[2].Status = 'Transferred'
        $result.Meetings[1].Selected = $false
        InModuleScope MeetingCleanupOnPrem -Parameters @{ Result = $result } -ScriptBlock {
            param($Result)
            Update-McoResultCounts $Result
            $meetings = @($Result.Meetings)
            $all = @($meetings | ForEach-Object { @($_.Copies) })
            $copies = @($all | Where-Object { $_.EventId -and $_.Role -in 'Organizer', 'Attendee', 'Room' })
            $expected = [pscustomobject]@{
                Meetings = $meetings.Count; Series = @($meetings | Where-Object Kind -eq 'Series').Count; Selected = @($meetings | Where-Object Selected).Count
                Organizers = @($meetings | ForEach-Object { [string]$_.Organizer } | Where-Object { $_ } | Select-Object -Unique).Count
                Copies = $copies.Count; Mailboxes = @($copies | ForEach-Object Mailbox | Select-Object -Unique).Count
                RoomCopies = @($copies | Where-Object Role -eq 'Room').Count; OrganizerCopies = @($copies | Where-Object Role -eq 'Organizer').Count
                AttendeeCopies = @($copies | Where-Object Role -eq 'Attendee').Count; OccurrenceCopies = @($copies | Where-Object Occurrence).Count
                Rooms = @($copies | Where-Object Role -eq 'Room').Count; Attendees = @($copies | Where-Object Role -eq 'Attendee').Count
                NotProcessed = @($all | Where-Object Result -eq 'Not processed').Count
                Removed = @($all | Where-Object Result -eq 'Removed').Count; Cancelled = @($all | Where-Object Result -eq 'Cancelled').Count
                AlreadyGone = @($all | Where-Object Result -eq 'Already gone').Count; Kept = @($all | Where-Object Result -eq 'Kept').Count
                Restored = @($all | Where-Object Result -eq 'Restored').Count; AlreadyPresent = @($all | Where-Object Result -eq 'Already present').Count
                NotFound = @($all | Where-Object { $_.Result -eq 'Not found' -and $_.Action -eq 'Restore' }).Count; NotRestorable = @($all | Where-Object Result -eq 'Not restorable').Count
                Transferred = @($meetings | Where-Object Status -eq 'Transferred').Count
                Failed = @($all | Where-Object { $_.Result -in 'Failed', 'Not done' }).Count
            }
            ($Result.Counts | ConvertTo-Json -Compress) | Should -Be ($expected | ConvertTo-Json -Compress)
            $Result.Counts.Removed | Should -Be 1
            $Result.Counts.NotFound | Should -Be 1
            $Result.Counts.Failed | Should -Be 1
        }
    }

    It 'neutralises formulas in CSV cells and writes HTML-safe JSON' {
        $fast = [MeetingCleanupOnPremNative.Fast]
        $fast::CsvCell('=HYPERLINK("x")', ';') | Should -Be '"''=HYPERLINK(""x"")"'
        $fast::CsvCell('-1', ';') | Should -Be "'-1"
        $fast::CsvCell(@('a', 'b'), ';') | Should -Be 'a | b'
        $fast::CsvCell($true, ';') | Should -Be 'True'
        $fast::CsvCell('a;b', ';') | Should -Be '"a;b"'
        $table = [MeetingCleanupOnPremNative.Table]::new([string[]]@('Subject'))
        $table.Rows.Add([object[]]@('</script><b>'))
        $fast::TableJson($table) | Should -Not -Match '<'
        $zone = [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time')
        $fast::FormatDate([datetime]'2030-01-07T09:00:00Z', $zone, $false, $false) | Should -Be '2030-01-07 10:00'
        $fast::FormatDate('2030-01-07T23:00:00Z', $zone, $false, $true) | Should -Be '2030-01-07'
    }
}

Describe 'Simulated Exchange Server' {
    It 'finds the meetings of an organizer with every copy, reading each mailbox once and the items 50 at a time' {
        $settings = New-TestSettings
        $store = New-FakeTenant
        $r = Find-Test $settings $store
        @($r.Meetings.Subject) | Should -Be @('M1 Budget review', 'S1 Weekly sync', 'M2 Steering')
        $m1 = $r.Meetings[0]
        $m1.OrganizerCopy | Should -Be 'Present'
        @($m1.Copies | Where-Object EventId).Count | Should -Be 4
        ($m1.Copies | Where-Object Mailbox -eq 'room1@contoso.test').Role | Should -Be 'Room'
        ($m1.Copies | Where-Object Mailbox -eq 'ext@fabrikam.test').Result | Should -Be 'Not processed'
        $r.Meetings[1].Kind | Should -Be 'Series'
        $r.Meetings[1].Recurrence | Should -Match '^Weekly \(\w+\), 4 occurrences from 2030-01-09'
        ($r.Meetings[2].Copies | Where-Object Mailbox -eq 'gone@contoso.test').Detail | Should -Match 'no mailbox'
        $r.Counts.Copies | Should -Be 10
        # Organizer, 2 rooms, then each attendee mailbox once (att1 has M1 and S1): one CalendarView each.
        $store.Calls.FindItem | Should -Be 7
        $store.Calls.GetItem | Should -BeLessOrEqual 7
    }

    It 'reads 50 items per GetItem call, and an item at a time when the batch is refused' {
        $settings = New-TestSettings
        $store = New-FakeTenant -Extra 120
        Install-FakeEws -Store $store -Settings $settings
        $events = InModuleScope MeetingCleanupOnPrem { Get-McoMailboxEvents -Mailbox 'org@contoso.test' -StartUtc ([datetime]'2030-01-01Z') -EndUtc ([datetime]'2030-12-31Z') }
        @($events).Count | Should -Be 124
        $store.Calls.GetItem | Should -Be 3
        @($events | Where-Object { -not $_.Organizer }).Count | Should -Be 0
        InModuleScope MeetingCleanupOnPrem {
            $items = Get-McoItems -Mailbox 'org@contoso.test' -Requests @([pscustomobject]@{ EventId = 'missing'; Master = $false }, [pscustomobject]@{ EventId = 'missing2'; Master = $true })
            $items.Count | Should -Be 2
            $items[0] | Should -BeNullOrEmpty
        }
    }

    It 'sends a request again when Exchange is busy (throttling), up to Connection.MaxRetries times' {
        $settings = New-TestSettings
        $settings.MaxRetries = 2
        $store = New-FakeTenant
        $store.Busy = @{ FindItem = 2 }
        Install-FakeEws -Store $store -Settings $settings
        $view = InModuleScope MeetingCleanupOnPrem { , (Get-McoCalendarView -Mailbox 'org@contoso.test' -StartUtc ([datetime]'2030-01-01Z') -EndUtc ([datetime]'2030-12-31Z')) }
        $view.Count | Should -BeGreaterThan 0
        $store.Calls.FindItem | Should -Be 3
        # Busy once more than the retries: the error of Exchange.
        $store.Busy = @{ FindItem = 3 }
        { InModuleScope MeetingCleanupOnPrem { Get-McoCalendarView -Mailbox 'att1@contoso.test' -StartUtc ([datetime]'2030-01-01Z') -EndUtc ([datetime]'2030-12-31Z') } } | Should -Throw '*ErrorServerBusy*'
    }

    It 'reads a calendar page after page, and says when items starting at the same time do not fit in a page' {
        $settings = New-TestSettings
        $settings.PageSize = 10
        $store = New-FakeTenant
        $t0 = [datetime]'2030-02-04T09:00:00Z'
        for ($i = 0; $i -lt 35; $i++) { [void](Add-FakeMeeting $store ("P{0:D2} Planning" -f $i) 'other@contoso.test' @('att9@contoso.test') -Rooms @('busy@contoso.test') -Start $t0.AddHours($i)) }
        Install-FakeEws -Store $store -Settings $settings
        $view = InModuleScope MeetingCleanupOnPrem { , (Get-McoCalendarView -Mailbox 'busy@contoso.test' -StartUtc ([datetime]'2030-01-01Z') -EndUtc ([datetime]'2030-12-31Z') -PageSize 10) }
        $view.Count | Should -Be 35
        @($view | ForEach-Object EventId | Select-Object -Unique).Count | Should -Be 35
        $store.Calls.FindItem | Should -BeGreaterOrEqual 4

        # 15 meetings at the same time and a page of 10: the 5 later ones are still read, with a warning.
        $store = New-FakeTenant
        for ($i = 0; $i -lt 15; $i++) { [void](Add-FakeMeeting $store ("Q{0:D2} Same time" -f $i) 'other@contoso.test' @('att9@contoso.test') -Rooms @('busy@contoso.test') -Start $t0) }
        for ($i = 0; $i -lt 5; $i++) { [void](Add-FakeMeeting $store ("R{0:D2} Later" -f $i) 'other@contoso.test' @('att9@contoso.test') -Rooms @('busy@contoso.test') -Start $t0.AddDays(1).AddHours($i)) }
        Install-FakeEws -Store $store -Settings $settings
        $request = New-McoRequest -Settings $settings -Room 'busy@contoso.test' -Start ([datetime]'2030-01-01') -End ([datetime]'2030-12-31')
        $r = Find-McoMeetings -Settings $settings -Request $request
        @($r.Meetings | Where-Object Subject -like 'R*').Count | Should -Be 5
        @($r.Meetings | Where-Object Subject -like 'Q*').Count | Should -Be 10
        $r.Status | Should -Be 'Warning'
        ($r.Warnings -join ' ') | Should -Match 'busy@contoso\.test: more than 10 items at .*raise Connection\.PageSize'
    }

    It 'removes the attendee and room copies silently, then restores them from Recoverable Items' {
        $settings = New-TestSettings
        $store = New-FakeTenant -From ([datetime]::UtcNow.Date.AddDays(3).AddHours(9))
        # Recoverable Items one item per page: room1 holds two removed copies, read in two pages.
        $store.RecoverablePage = 1
        Install-FakeEws -Store $store -Settings $settings
        $r = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -Start ([datetime]::UtcNow.Date) -End ([datetime]::UtcNow.Date.AddDays(60)))
        $folder = Join-Path $TestDrive 'remove-run'
        $removed = Invoke-McoCleanup -Settings $settings -Result $r -Action Remove -BackupPath (Join-Path $folder 'MeetingCleanupOnPrem-Backup.json')
        $removed.Status | Should -Be 'Completed'
        $removed.Counts.Removed | Should -Be 7
        $removed.Counts.Kept | Should -Be 3
        @((Get-FakeCalendar $store 'att1@contoso.test')).Count | Should -Be 1
        (Get-Content -Raw (Join-Path $folder 'MeetingCleanupOnPrem-Backup.json') | ConvertFrom-Json).Meetings[0].Event.Body | Should -Be 'Agenda of M1 Budget review'
        $null = Export-McoReport -Result $removed -OutputPath $settings.OutputPath -Prefix 'MeetingCleanupOnPrem' -Formats Csv -Directory $folder
        $restored = Invoke-McoRestore -Settings $settings -Result (Import-McoRestoreSource -Path $folder)
        $restored.Counts.Restored | Should -Be 7
        $restored.Status | Should -Be 'Completed'
        @($restored.Meetings | ForEach-Object { $_.Copies } | Where-Object Verified -eq 'Yes').Count | Should -Be 7
        @((Get-FakeCalendar $store 'att1@contoso.test')).Count | Should -Be 7
    }

    It 'replays a report: Remove of one meeting and Transfer, without a new search' {
        $settings = New-TestSettings
        $store = New-FakeTenant -From ([datetime]::UtcNow.Date.AddDays(3).AddHours(9))
        Install-FakeEws -Store $store -Settings $settings
        $period = @{ Start = [datetime]::UtcNow.Date; End = [datetime]::UtcNow.Date.AddDays(60) }
        $found = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' @period)
        $report = Export-McoReport -Result $found -OutputPath $settings.OutputPath -Prefix 'MeetingCleanupOnPrem' -Formats Csv
        $replay = New-McoRequest -Settings $settings -FromReport $report.Directory -Action Remove
        (Test-McoRequest $replay).IsValid | Should -BeTrue
        $m1 = ($found.Meetings | Where-Object Subject -like 'M1*').MeetingId
        $source = Import-McoRestoreSource -Path $report.Directory -MeetingId $m1
        $source.Meetings.Count | Should -Be 1
        $removed = Invoke-McoCleanup -Settings $settings -Result $source -Action Remove -BackupPath (Join-Path $TestDrive 'replay\MeetingCleanupOnPrem-Backup.json')
        $removed.Counts.Removed | Should -Be 3
        (Get-Content -Raw (Join-Path $TestDrive 'replay\MeetingCleanupOnPrem-Backup.json') | ConvertFrom-Json).Meetings[0].Copies.Count | Should -Be 4

        $s1 = ($found.Meetings | Where-Object Subject -like 'S1*').MeetingId
        $source = Import-McoRestoreSource -Path $report.Directory -MeetingId $s1
        $plan = Get-McoTransferPlan -Result $source -NewOrganizer (Resolve-McoNewOrganizer -Address 'new@contoso.test') -Comment 'Now organized by {0}.'
        $plan.Recreate.Count | Should -Be 1
        $moved = Invoke-McoTransfer -Settings $settings -Result $source -Plan $plan -Comment 'Now organized by {0}.'
        $moved.Counts.Transferred | Should -Be 1
        @((Get-FakeCalendar $store 'new@contoso.test') | Where-Object Type -eq 'RecurringMaster').Count | Should -Be 1
    }

    It 'cancels with the organizer, transfers to a new organizer, and keeps the occurrences of a room' {
        $settings = New-TestSettings
        $store = New-FakeTenant
        $cancelled = Invoke-McoCleanup -Settings $settings -Result (Find-Test $settings $store) -Action Cancel -Comment 'Cancelled.'
        $cancelled.Counts.Cancelled | Should -Be 3
        @($cancelled.Meetings.Status) | Should -Be @('Cancelled', 'Cancelled', 'Cancelled')

        $store = New-FakeTenant -From ([datetime]::UtcNow.Date.AddDays(3).AddHours(9))
        Install-FakeEws -Store $store -Settings $settings
        $request = New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -Start ([datetime]::UtcNow.Date) -End ([datetime]::UtcNow.Date.AddDays(60))
        $found = Find-McoMeetings -Settings $settings -Request $request
        $plan = Get-McoTransferPlan -Result $found -NewOrganizer (Resolve-McoNewOrganizer -Address 'new@contoso.test') -Comment 'Now organized by {0}.'
        $moved = Invoke-McoTransfer -Settings $settings -Result $found -Plan $plan -Comment 'Now organized by {0}.'
        $moved.Counts.Transferred | Should -Be 3
        @((Get-FakeCalendar $store 'new@contoso.test') | Where-Object { $_.Type -ne 'Occurrence' }).Count | Should -Be 3
        # The Transfers file and tab: one row per meeting, the organizer change at a glance (and only for a transfer).
        $report = Export-McoReport -Result $moved -OutputPath $settings.OutputPath -Prefix 'MeetingCleanupOnPrem' -Formats Csv, Html
        $rows = @(Import-Csv $report.Files.Transfers -Delimiter ';')
        $rows.Count | Should -Be 3
        $row = $rows | Where-Object Subject -like 'M1*'
        $row.OldOrganizer | Should -Be 'org@contoso.test'
        $row.OldOrganizerState | Should -Be 'Not checked'
        $row.NewOrganizer | Should -Be 'new@contoso.test'
        $row.Method | Should -Be 'Re-created'
        $row.Status | Should -Be 'Transferred'
        $row.NewMeeting | Should -Be 'Created'
        $row.NewMeetingId | Should -Not -BeNullOrEmpty
        $row.Invited | Should -Be '2'
        $row.Rooms | Should -Be '1'
        $row.OldOrganizerCopy | Should -Be 'Cancelled'
        $row.OldCopiesRemoved | Should -Be '3'
        $row.OldCopiesFailed | Should -Be '0'
        $html = Get-Content $report.Files.Html -Raw
        $html | Should -Match '<script type="application/json" id="data-transfers">\[\{"MeetingId"'
        $html | Should -Match 'data-tab="transfers"'

        $rooms = Find-Test $settings (New-FakeTenant) -Room 'room2@contoso.test'
        $rooms.Meetings[0].Scope | Should -Be 'Occurrences'
        $rooms.Meetings[0].Occurrences | Should -Be 2
        $rooms.Counts.OccurrenceCopies | Should -Be 6
        $search = Export-McoReport -Result $rooms -OutputPath $settings.OutputPath -Prefix 'Search' -Formats Csv, Html
        $search.Files.Contains('Transfers') | Should -BeFalse
        Get-Content $search.Files.Html -Raw | Should -Match 'id="data-transfers">\[\]</script>'
    }
}

Describe 'Series by occurrences' {
    It 'checks -SeriesScope: a period with an action, never a transfer, always in rooms mode' {
        $settings = New-TestSettings
        $settings.SeriesScope | Should -Be 'Whole'
        $day = @{ Start = [datetime]'2030-01-16'; End = [datetime]'2030-01-16' }
        $r = New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Occurrences -Action Cancel
        ((Test-McoRequest $r).Problems -join ' ') | Should -Match 'Series by occurrences: give the period of the action'
        $r = New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Occurrences -Action Cancel @day
        (Test-McoRequest $r).IsValid | Should -BeTrue
        $r.SeriesScope | Should -Be 'Occurrences'
        (Test-McoRequest (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Occurrences)).IsValid | Should -BeTrue
        $r = New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Occurrences -Action Transfer -NewOrganizer 'new@contoso.test' @day
        ((Test-McoRequest $r).Problems -join ' ') | Should -Match 'Transfer moves whole series'
        (New-McoRequest -Settings $settings -Room 'room1@contoso.test').SeriesScope | Should -Be 'Occurrences'
        $settings.SeriesScope = 'Occurrences'
        (New-McoRequest -Settings $settings -Organizer 'org@contoso.test').SeriesScope | Should -Be 'Occurrences'
        (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Whole).SeriesScope | Should -Be 'Whole'
        $settings.SeriesScope = 'Some'
        ((Test-McoConfiguration $settings).Problems -join ' ') | Should -Match "Search.SeriesScope must be 'Whole' or 'Occurrences'"
    }

    It 'cancels one occurrence of a series: one cancellation for that date, the other occurrences intact' {
        $settings = New-TestSettings
        $store = New-FakeTenant
        Install-FakeEws -Store $store -Settings $settings
        $request = New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Occurrences -Start ([datetime]'2030-01-16') -End ([datetime]'2030-01-16') -Action Cancel
        $found = Find-McoMeetings -Settings $settings -Request $request
        @($found.Meetings.Subject) | Should -Be @('S1 Weekly sync')
        $found.Request.SeriesScope | Should -Be 'Occurrences'
        $s1 = $found.Meetings[0]
        $s1.Scope | Should -Be 'Occurrences'
        $s1.Occurrences | Should -Be 1
        [MeetingCleanupOnPremNative.Fast]::KindText($s1) | Should -Be '1 occ.'
        $copies = @($s1.Copies | Where-Object EventId)
        @($copies.Role | Sort-Object) | Should -Be @('Attendee', 'Organizer', 'Room')
        @($copies | Where-Object { $_.OccurrenceKey -like '2030-01-16T*' -and $_.SeriesId }).Count | Should -Be 3
        $found.Counts.OccurrenceCopies | Should -Be 3

        $plan = Get-McoCleanupPlan -Result $found -Action Cancel
        $plan.Cancel.Count | Should -Be 1
        $plan.Remove.Count | Should -Be 2
        ($plan.Lines -join ' ') | Should -Match '0 meeting\(s\) and 1 occurrence\(s\) cancelled'
        $done = Invoke-McoCleanup -Settings $settings -Result $found -Action Cancel -Comment 'No sync this week.'
        $done.Status | Should -Be 'Completed'
        $done.Meetings[0].Status | Should -Be 'Cancelled'
        $done.Counts.Removed | Should -Be 2
        # The organizer's occurrence of 16 January cancelled, the attendee's and the room's gone; the series goes on.
        $series = { param($mailbox) @((Get-FakeCalendar $store $mailbox) | Where-Object { $_.Uid -eq $store.Ids.S1 }) }
        $org = & $series 'org@contoso.test'
        @($org | Where-Object Type -eq 'RecurringMaster').Count | Should -Be 1
        @($org | Where-Object Cancelled | ForEach-Object { $_.Start.ToString('yyyy-MM-dd') }) | Should -Be @('2030-01-16')
        foreach ($mailbox in 'att1@contoso.test', 'room2@contoso.test') {
            $left = & $series $mailbox
            @($left | Where-Object Type -eq 'Occurrence' | ForEach-Object { $_.Start.ToString('MM-dd') }) | Should -Be @('01-09', '01-23', '01-30')
            @($left | Where-Object Cancelled).Count | Should -Be 0
        }
        @($done.Meetings[0].Copies | Where-Object Verified -eq 'Yes').Count | Should -Be 3
    }

    It 'reads the occurrences from the attendees when the organizer mailbox is gone, and leaves a series whose organizer cannot be read' {
        $settings = New-TestSettings
        $store = New-FakeTenant
        [void](Add-FakeMeeting $store 'L1 Leaver sync' 'left@contoso.test' @('att1@contoso.test') -Rooms @('room2@contoso.test') -Start ([datetime]'2030-01-14T10:00:00Z') -Weeks 3 -NoOrganizerCopy)
        [void]$store.DenyMailbox.Add('left@contoso.test')
        Install-FakeEws -Store $store -Settings $settings
        $r = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'left@contoso.test' -SeriesScope Occurrences -Start ([datetime]'2030-01-21') -End ([datetime]'2030-01-21'))
        $l1 = @($r.Meetings | Where-Object Subject -eq 'L1 Leaver sync')
        $l1.Count | Should -Be 1
        $l1[0].OrganizerCopy | Should -Be 'Mailbox deleted'
        $l1[0].Occurrences | Should -Be 1
        @($l1[0].Copies | Where-Object { $_.EventId -and $_.OccurrenceKey -like '2030-01-21T*' }).Count | Should -Be 2

        # The organizer's mailbox answers with an error: the series is left as it is, never the attendees alone.
        $store = New-FakeTenant
        $store.MailboxError['org@contoso.test'] = 'ErrorMailboxMoveInProgress'
        Install-FakeEws -Store $store -Settings $settings
        $r = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Occurrences -Start ([datetime]'2030-01-16') -End ([datetime]'2030-01-16'))
        $s1 = @($r.Meetings | Where-Object Kind -eq 'Series')
        $s1.Count | Should -Be 1
        $s1[0].OrganizerCopy | Should -Be 'Not read'
        @($s1[0].Copies | Where-Object EventId).Count | Should -Be 0
        ($s1[0].Notes -join ' ') | Should -Match 'Not acted on: the occurrences of its organizer could not be read'
        $r.Status | Should -Be 'Warning'
        ($r.Warnings -join ' ') | Should -Match '1 series left as they are'
        (Get-McoCleanupPlan -Result $r -Action Cancel).Remove.Count | Should -Be 0
    }

    It 'leaves out the occurrences skipped in a reviewed report, says when every occurrence is in the period, and keeps a transfer for whole series' {
        $settings = New-TestSettings
        $store = New-FakeTenant
        Install-FakeEws -Store $store -Settings $settings
        $found = Find-McoMeetings -Settings $settings -Request (New-McoRequest -Settings $settings -Organizer 'org@contoso.test' -SeriesScope Occurrences -Start ([datetime]'2030-01-01') -End ([datetime]'2030-01-31'))
        $s1 = $found.Meetings | Where-Object Subject -like 'S1*'
        $s1.Occurrences | Should -Be 4
        ($s1.Notes -join ' ') | Should -Match 'Every occurrence of the series is in the period: each one is acted on separately'
        $report = Export-McoReport -Result $found -OutputPath $settings.OutputPath -Prefix 'MeetingCleanupOnPrem' -Formats Csv, Html
        (Import-Csv $report.Files.Meetings -Delimiter ';' | Where-Object Subject -like 'S1*').OccurrencesSkipped | Should -Be '0'
        Get-Content $report.Files.Html -Raw | Should -Match '"SeriesScope":"Occurrences"'

        # The report reviewed: the occurrence of 23 January left out (SkippedOccurrences, as in Meeting Cleanup).
        $summary = Get-Content $report.Files.Summary -Raw | ConvertFrom-Json -Depth 32
        $key = ($s1.Copies | Where-Object { $_.Role -eq 'Organizer' -and $_.OccurrenceKey -like '2030-01-23*' }).OccurrenceKey
        ($summary.Meetings | Where-Object Subject -like 'S1*').SkippedOccurrences = @($key)
        [IO.File]::WriteAllText($report.Files.Summary, ($summary | ConvertTo-Json -Depth 32), [Text.UTF8Encoding]::new($false))
        $source = Import-McoRestoreSource -Path $report.Directory -MeetingId $s1.MeetingId
        [MeetingCleanupOnPremNative.Fast]::KindText($source.Meetings[0]) | Should -Be '3/4 occ.'
        $plan = Get-McoCleanupPlan -Result $source -Action Remove
        $plan.Remove.Count | Should -Be 6
        $plan.NotSelected.Count | Should -Be 3
        ($plan.Lines -join ' ') | Should -Match '1 occurrence\(s\) left out in the report'
        $removed = Invoke-McoCleanup -Settings $settings -Result $source -Action Remove
        $removed.Counts.Removed | Should -Be 6
        @($removed.Meetings[0].Copies | Where-Object Result -eq 'Skipped').Count | Should -Be 3
        @((Get-FakeCalendar $store 'att1@contoso.test') | Where-Object { $_.Uid -eq $store.Ids.S1 -and $_.Type -eq 'Occurrence' } | ForEach-Object { $_.Start.ToString('MM-dd') }) | Should -Be @('01-23')
        $rows = @([MeetingCleanupOnPremNative.Fast]::MeetingTable($removed.Meetings).ToObjects())
        $rows[0].OccurrencesSkipped | Should -Be 1

        $plan = Get-McoTransferPlan -Result $source -NewOrganizer (Resolve-McoNewOrganizer -Address 'new@contoso.test') -Comment 'Now organized by {0}.'
        @($plan.Recreate).Count | Should -Be 0
        ($plan.Lines -join ' ') + ' ' + (@($plan.Skipped | ForEach-Object Reason) -join ' ') | Should -Match 'occurrences of a series: transfer the whole series'
    }
}
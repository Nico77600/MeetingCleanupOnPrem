$script:ConfigSchema = [ordered]@{
    Connection = [ordered]@{
        EwsUrl = 'EwsUrl'; Discovery = 'Discovery'; Mailbox = 'Mailbox'; AccessMode = 'AccessMode'
        Authentication = 'Authentication'; CredentialUser = 'CredentialUser'; RequestServerVersion = 'RequestServerVersion'
        MaxRetries = 'MaxRetries'; TimeoutSeconds = 'TimeoutSeconds'; PageSize = 'PageSize'; EwsServer = 'EwsServer'
        WindowsPackage = 'WindowsPackage'; CredentialFile = 'CredentialFile'
    }
    Search = [ordered]@{
        SearchIn = 'SearchIn'; PastDays = 'PastDays'; FutureDays = 'FutureDays'; Rooms = 'Rooms'
        RoomFile = 'RoomFile'; Mailboxes = 'Mailboxes'; MailboxFile = 'MailboxFile'; AcceptedDomains = 'AcceptedDomains'; DirectoryMode = 'DirectoryMode'; AllRooms = 'AllRooms'
    }
    ManagementShell = [ordered]@{
        Mode = 'ManagementShellMode'; ServerFqdn = 'ManagementShellServer'; ConnectionUri = 'ManagementShellUri'
        Authentication = 'ManagementShellAuthentication'; CredentialUser = 'ManagementShellCredentialUser'
    }
    Cleanup = [ordered]@{ CancelComment = 'CancelComment'; Verify = 'Verify' }
    Restore = [ordered]@{ Mode = 'RestoreMode'; WindowMinutes = 'RestoreWindowMinutes' }
    Transfer = [ordered]@{ Method = 'TransferMethod'; Comment = 'TransferComment' }
    Report = [ordered]@{ OutputPath = 'OutputPath'; FilePrefix = 'ReportPrefix'; Formats = 'ReportFormats'; CsvDelimiter = 'CsvDelimiter'; TimeZone = 'TimeZone' }
    Logging = [ordered]@{ Path = 'LogPath'; RetentionDays = 'LogRetentionDays' }
}
$script:SearchScopes = @('Organizer', 'Rooms', 'Mailboxes', 'AllMailboxes')
$script:Actions = @('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')
$script:SmtpPattern = '^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$'
$script:X500Pattern = '^(?i)(x500:)?/o=[^\r\n]+/cn=[^\r\n]+$'

function Get-McoDefaultConfiguration {
    @{
        EwsUrl = ''; Discovery = 'Manual'; Mailbox = ''; AccessMode = 'Impersonation'; Authentication = 'Windows'
        CredentialUser = ''; RequestServerVersion = 'Exchange2016'; MaxRetries = 3; TimeoutSeconds = 120; PageSize = 500; EwsServer = ''
        WindowsPackage = 'Negotiate'; CredentialFile = ''
        SearchIn = @('Organizer', 'Rooms'); PastDays = 0; FutureDays = 365; Rooms = @(); RoomFile = ''
        Mailboxes = @(); MailboxFile = ''; AcceptedDomains = @(); DirectoryMode = 'Auto'; AllRooms = $true
        ManagementShellMode = 'Rps'; ManagementShellServer = ''; ManagementShellUri = ''
        ManagementShellAuthentication = 'Kerberos'; ManagementShellCredentialUser = ''
        CancelComment = 'This meeting has been cancelled by the IT department.'; Verify = $true
        RestoreMode = 'Auto'; RestoreWindowMinutes = 10; TransferMethod = 'Recreate'
        TransferComment = 'This meeting is now organized by {0}.'
        OutputPath = '.\reports'; ReportPrefix = 'MeetingCleanupOnPrem'; ReportFormats = @('Csv', 'Html')
        CsvDelimiter = ';'; TimeZone = ''; LogPath = '.\logs'; LogRetentionDays = 30
    }
}

function Get-McoProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Resolve-McoPath {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root)
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    [IO.Path]::GetFullPath($Path, $Root)
}

$script:ZoneCache = @{}

function Get-McoTimeZone {
    <# Time zone of the dates: Report.TimeZone (IANA or Windows ID), the one of Windows when empty. Looked up once per ID. #>
    param([AllowNull()][AllowEmptyString()][string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return [TimeZoneInfo]::Local }
    $zone = $script:ZoneCache[$Id]
    if (-not $zone) { $zone = [TimeZoneInfo]::FindSystemTimeZoneById($Id); $script:ZoneCache[$Id] = $zone }
    $zone
}

function Format-McoDate {
    <# A UTC date shown in the time zone of the report: yyyy-MM-dd HH:mm (or yyyy-MM-dd). -PeriodEnd: an end at 00:00 shows the day before (included). #>
    param([AllowNull()][object]$Utc, [AllowNull()][AllowEmptyString()][string]$TimeZone, [switch]$DateOnly, [switch]$PeriodEnd)
    [MeetingCleanupOnPremNative.Fast]::FormatDate($Utc, (Get-McoTimeZone $TimeZone), [bool]$DateOnly, [bool]$PeriodEnd)
}

function ConvertTo-McoUtc {
    param([Parameter(Mandatory)][datetime]$Date, [AllowNull()][AllowEmptyString()][string]$TimeZone, [switch]$EndOfDay)
    if ($Date.Kind -eq [DateTimeKind]::Utc) { return $Date }
    $d = [datetime]::SpecifyKind($Date, [DateTimeKind]::Unspecified)
    if ($EndOfDay -and $d.TimeOfDay -eq [TimeSpan]::Zero) { $d = $d.AddDays(1) }
    [TimeZoneInfo]::ConvertTimeToUtc($d, (Get-McoTimeZone $TimeZone))
}

function Get-McoDefaultPeriod {
    param([Parameter(Mandatory)][hashtable]$Settings)
    $zone = Get-McoTimeZone $Settings.TimeZone
    $today = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $zone).Date
    [pscustomobject]@{
        Start = ConvertTo-McoUtc $today.AddDays(-[int]$Settings.PastDays) $Settings.TimeZone
        End = ConvertTo-McoUtc $today.AddDays([int]$Settings.FutureDays) $Settings.TimeZone -EndOfDay
    }
}

function Test-McoConfiguration {
    param([Parameter(Mandatory)][hashtable]$Configuration)
    $p = [Collections.Generic.List[string]]::new()
    $c = $Configuration
    if ([string]$c.EwsUrl -notmatch '^https://') { $p.Add('Connection.EwsUrl must be an HTTPS EWS endpoint.') }
    if ([string]$c.Discovery -notin 'Manual', 'Autodiscover') { $p.Add("Connection.Discovery must be 'Manual' or 'Autodiscover'.") }
    if ([string]$c.AccessMode -notin 'Self', 'Delegate', 'Impersonation') { $p.Add("Connection.AccessMode must be 'Self', 'Delegate' or 'Impersonation'.") }
    if ([string]$c.AccessMode -ne 'Self' -and -not [string]$c.Mailbox) { $p.Add('Connection.Mailbox is required for Delegate or Impersonation access.') }
    if ([string]$c.Authentication -notin 'Windows', 'Basic') { $p.Add("Connection.Authentication must be 'Windows' or 'Basic'.") }
    if ([string]$c.RequestServerVersion -notin 'Exchange2013_SP1', 'Exchange2016', 'Exchange2019') { $p.Add('Connection.RequestServerVersion must be Exchange2013_SP1, Exchange2016 or Exchange2019.') }
    if ([int]$c.MaxRetries -lt 0 -or [int]$c.MaxRetries -gt 10) { $p.Add('Connection.MaxRetries must be between 0 and 10.') }
    if ([int]$c.TimeoutSeconds -lt 10 -or [int]$c.TimeoutSeconds -gt 900) { $p.Add('Connection.TimeoutSeconds must be between 10 and 900.') }
    if ([int]$c.PageSize -lt 10 -or [int]$c.PageSize -gt 1000) { $p.Add('Connection.PageSize must be between 10 and 1000.') }
    if ([string]$c.EwsServer -and [string]$c.EwsServer -notmatch '^[A-Za-z0-9.-]+$') { $p.Add('Connection.EwsServer must be empty, a server name or an IP address.') }
    if ([string]$c.WindowsPackage -notin 'Negotiate', 'NTLM', 'Kerberos') { $p.Add("Connection.WindowsPackage must be 'Negotiate', 'NTLM' or 'Kerberos'.") }
    if ([string]$c.CredentialFile -and -not (Test-Path -LiteralPath ([string]$c.CredentialFile) -PathType Leaf)) { $p.Add("Connection.CredentialFile '$($c.CredentialFile)' not found (create it with Get-Credential | Export-Clixml, as the account that runs the tool).") }
    if (@($c.SearchIn).Count -eq 0 -or @($c.SearchIn | Where-Object { $_ -notin $script:SearchScopes }).Count) { $p.Add("Search.SearchIn must contain one or more of: $($script:SearchScopes -join ', ').") }
    foreach ($a in @($c.Rooms) + @($c.Mailboxes)) { if ($a -and [string]$a -notmatch $script:SmtpPattern) { $p.Add("Search address '$a' is not an SMTP address.") } }
    foreach ($d in @($c.AcceptedDomains)) { if ($d -notmatch '^[A-Za-z0-9.-]+$') { $p.Add("Search.AcceptedDomains '$d' is invalid.") } }
    if ([string]$c.DirectoryMode -notin 'Auto', 'ExchangePowerShell', 'None') { $p.Add("Search.DirectoryMode must be 'Auto', 'ExchangePowerShell' or 'None'.") }
    if ([string]$c.ManagementShellMode -notin 'Existing', 'Auto', 'Rps') { $p.Add("ManagementShell.Mode must be 'Existing', 'Auto' or 'Rps'.") }
    if ([string]$c.ManagementShellAuthentication -notin 'Kerberos', 'Negotiate', 'Basic') { $p.Add("ManagementShell.Authentication must be 'Kerberos', 'Negotiate' or 'Basic'.") }
    if ([string]$c.ManagementShellMode -eq 'Rps' -and -not $c.ManagementShellServer -and -not $c.ManagementShellUri) { $p.Add('ManagementShell.ServerFqdn or ManagementShell.ConnectionUri is required in Rps mode.') }
    if ($c.ManagementShellUri -and [string]$c.ManagementShellUri -notmatch '^https?://') { $p.Add('ManagementShell.ConnectionUri must start with http:// or https://.') }
    foreach ($n in 'PastDays', 'FutureDays') { if ([int]$c[$n] -lt 0 -or [int]$c[$n] -gt 3650) { $p.Add("Search.$n must be between 0 and 3650.") } }
    if ($c.Verify -isnot [bool] -or $c.AllRooms -isnot [bool]) { $p.Add('Cleanup.Verify and Search.AllRooms must be $true or $false.') }
    if ([int]$c.RestoreWindowMinutes -lt 1 -or [int]$c.RestoreWindowMinutes -gt 240) { $p.Add('Restore.WindowMinutes must be between 1 and 240.') }
    if ([string]$c.RestoreMode -notin 'Auto', 'ExchangePowerShell', 'Ews') { $p.Add("Restore.Mode must be 'Auto', 'ExchangePowerShell' or 'Ews'.") }
    if ([string]$c.TransferMethod -ne 'Recreate') { $p.Add('Transfer.Method must be Recreate: Exchange Server has no EXO native organizer-transfer cmdlet.') }
    if ([string]$c.OutputPath -eq '' -or [string]$c.LogPath -eq '') { $p.Add('Report.OutputPath and Logging.Path are required.') }
    if ([string]$c.ReportPrefix -match '[\\/:*?"<>|\s]') { $p.Add('Report.FilePrefix must be a file name without spaces.') }
    if (@($c.ReportFormats | Where-Object { $_ -notin 'Csv', 'Html' }).Count -or -not @($c.ReportFormats).Count) { $p.Add("Report.Formats must contain 'Csv', 'Html' or both.") }
    if ([string]$c.CsvDelimiter -notin ';', ',', "`t") { $p.Add("Report.CsvDelimiter must be ';', ',' or a tab.") }
    if ([string]$c.TimeZone) { try { $null = Get-McoTimeZone $c.TimeZone } catch { $p.Add("Report.TimeZone '$($c.TimeZone)' is not available on this computer.") } }
    [pscustomobject]@{ IsValid = $p.Count -eq 0; Problems = @($p) }
}

function Import-McoConfiguration {
    param([string]$Path = (Join-Path $script:ToolRoot 'config\MeetingCleanupOnPrem.config.psd1'), [string]$Root = $script:ToolRoot)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration file not found: $Path" }
    $s = Get-McoDefaultConfiguration
    $data = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($section in $data.Keys) {
        if (-not $script:ConfigSchema.Contains($section)) { throw "Unknown configuration section '$section'." }
        foreach ($key in $data[$section].Keys) {
            if (-not $script:ConfigSchema[$section].Contains($key)) { throw "Unknown configuration key '$section.$key'." }
            $s[$script:ConfigSchema[$section][$key]] = $data[$section][$key]
        }
    }
    foreach ($key in 'SearchIn', 'Rooms', 'Mailboxes', 'AcceptedDomains', 'ReportFormats') { $s[$key] = @($s[$key] | Where-Object { "$_" }) }
    foreach ($key in 'OutputPath', 'LogPath', 'RoomFile', 'MailboxFile') { if ($s[$key]) { $s[$key] = Resolve-McoPath $s[$key] $Root } }
    $s.ConfigPath = [IO.Path]::GetFullPath($Path)
    $check = Test-McoConfiguration $s
    if (-not $check.IsValid) { throw "Invalid configuration:`n - $($check.Problems -join "`n - ")" }
    $s
}

function Read-McoAddressFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Address file not found: $Path" }
    $lines = @([IO.File]::ReadAllLines($Path) | ForEach-Object Trim | Where-Object { $_ -and -not $_.StartsWith('#') })
    if (-not $lines.Count) { return @() }
    $header = @($lines[0] -split '[;,]' | ForEach-Object { $_.Trim().Trim('"') })
    $columns = @('PrimarySmtpAddress', 'EmailAddress', 'Mail', 'WindowsEmailAddress', 'UserPrincipalName', 'Address', 'Organizer')
    $found = @($columns | Where-Object { $header -contains $_ })
    if ($found.Count) {
        $delimiter = if ($lines[0].Contains(';')) { ';' } else { ',' }
        return @($lines | ConvertFrom-Csv -Delimiter $delimiter | ForEach-Object { $found | ForEach-Object { ([string]$_.$_).Trim() } | Where-Object { $_ } | Select-Object -First 1 })
    }
    @($lines | ForEach-Object { ($_ -split '[;,\s]+')[0].Trim('"') } | Where-Object { $_ -match $script:SmtpPattern })
}

function Split-McoAddressList {
    param([AllowNull()][string[]]$Text)
    @($Text | ForEach-Object { [string]$_ -split '[;,\s]+' } | Where-Object { $_ } | Select-Object -Unique)
}

function New-McoRequest {
    param(
        [Parameter(Mandatory)][hashtable]$Settings, [string[]]$Organizer, [string]$OrganizerFile,
        [string[]]$Room, [string]$RoomFile, [Nullable[datetime]]$Start, [Nullable[datetime]]$End,
        [string]$Subject, [string[]]$MeetingId, [string[]]$SearchIn, [string[]]$Mailbox, [string]$MailboxFile,
        [ValidateSet('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')][string]$Action = 'Report',
        [string]$Comment, [string]$FromReport, [string]$NewOrganizer
    )
    $period = Get-McoDefaultPeriod $Settings
    $orgFile = if ($OrganizerFile) { [IO.Path]::GetFullPath($OrganizerFile, (Get-Location).Path) } else { '' }
    $org = @(Split-McoAddressList $Organizer)
    if ($orgFile) { $org += Read-McoAddressFile $orgFile }
    $roomFilePath = if ($RoomFile) { [IO.Path]::GetFullPath($RoomFile, (Get-Location).Path) } else { '' }
    $rooms = @(Split-McoAddressList $Room | ForEach-Object ToLowerInvariant)
    if ($roomFilePath) { $rooms += @(Read-McoAddressFile $roomFilePath | ForEach-Object ToLowerInvariant) }
    $mb = @(Split-McoAddressList $Mailbox | ForEach-Object ToLowerInvariant)
    $mbFile = if ($MailboxFile) { [IO.Path]::GetFullPath($MailboxFile, (Get-Location).Path) } elseif ($Settings.MailboxFile) { $Settings.MailboxFile } else { '' }
    if ($mbFile) { $mb += @(Read-McoAddressFile $mbFile | ForEach-Object ToLowerInvariant) }
    [pscustomobject]@{
        Mode = if (-not $org.Count -and $rooms.Count) { 'Rooms' } else { 'Organizers' }
        Organizer = @($org | Select-Object -Unique); OrganizerFile = $orgFile; Room = @($rooms | Select-Object -Unique); RoomFile = $roomFilePath
        Start = if ($null -ne $Start) { ConvertTo-McoUtc $Start $Settings.TimeZone } else { $period.Start }
        End = if ($null -ne $End) { ConvertTo-McoUtc $End $Settings.TimeZone -EndOfDay } else { $period.End }
        Subject = [string]$Subject; MeetingId = @($MeetingId | ForEach-Object { $_ -split '[;,\s]+' } | Where-Object { $_ } | ForEach-Object ToUpperInvariant | Select-Object -Unique)
        SearchIn = if (-not $org.Count -and $rooms.Count) { @('Rooms') } elseif ($SearchIn) { @($SearchIn | Select-Object -Unique) } else { @($Settings.SearchIn) }
        Mailboxes = @($mb | Select-Object -Unique); MailboxFile = $mbFile; Action = $Action; Comment = if ($Comment) { $Comment } elseif ($Action -eq 'Transfer') { $Settings.TransferComment } else { $Settings.CancelComment }
        FromReport = $FromReport; NewOrganizer = ([string]$NewOrganizer).Trim().ToLowerInvariant()
    }
}

function Test-McoRequest {
    param([Parameter(Mandatory)][pscustomobject]$Request)
    $p = [Collections.Generic.List[string]]::new()
    if ($Request.Action -ne 'Restore' -and -not $Request.FromReport -and $Request.Mode -eq 'Organizers' -and -not @($Request.Organizer).Count) { $p.Add('Give the organizer with -Organizer or -OrganizerFile, or use -Room.') }
    if ($Request.Mode -eq 'Rooms' -and -not @($Request.Room).Count) { $p.Add('Rooms mode needs at least one room.') }
    foreach ($a in @($Request.Organizer) + @($Request.Room) + @($Request.Mailboxes)) { if ($a -notmatch $script:SmtpPattern -and $a -notmatch $script:X500Pattern) { $p.Add("'$a' is neither an SMTP nor an X500 address.") } }
    if ($Request.End -le $Request.Start) { $p.Add('The end of the period must be after its start.') }
    if ($Request.Action -eq 'Restore' -and -not $Request.FromReport) { $p.Add('Restore needs -FromReport.') }
    if ($Request.FromReport -and $Request.Action -eq 'Report') { $p.Add('A report replay needs -Action Remove, Cancel, Transfer or Restore.') }
    if ($Request.Action -eq 'Transfer' -and -not $Request.NewOrganizer) { $p.Add('Transfer needs -NewOrganizer.') }
    [pscustomobject]@{ IsValid = $p.Count -eq 0; Problems = @($p) }
}

$script:ConfigSchema = [ordered]@{
    Connection = [ordered]@{
        EwsUrl = 'EwsUrl'; Discovery = 'Discovery'; Mailbox = 'Mailbox'; AccessMode = 'AccessMode'
        Authentication = 'Authentication'; CredentialUser = 'CredentialUser'; RequestServerVersion = 'RequestServerVersion'
        MaxRetries = 'MaxRetries'; TimeoutSeconds = 'TimeoutSeconds'; PageSize = 'PageSize'; EwsServer = 'EwsServer'
        WindowsPackage = 'WindowsPackage'; CredentialFile = 'CredentialFile'
    }
    Search = [ordered]@{
        SearchIn = 'SearchIn'; PastDays = 'PastDays'; FutureDays = 'FutureDays'; SeriesScope = 'SeriesScope'; Rooms = 'Rooms'
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
# A series: the whole series (every occurrence, past ones included), or only its occurrences in the period.
$script:SeriesScopes = @('Whole', 'Occurrences')
$script:Actions = @('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')
$script:SmtpPattern = '^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$'
$script:X500Pattern = '^(?i)(x500:)?/o=[^\r\n]+/cn=[^\r\n]+$'

function Get-McoDefaultConfiguration {
    @{
        EwsUrl = ''; Discovery = 'Manual'; Mailbox = ''; AccessMode = 'Impersonation'; Authentication = 'Windows'
        CredentialUser = ''; RequestServerVersion = 'Exchange2016'; MaxRetries = 3; TimeoutSeconds = 120; PageSize = 500; EwsServer = ''
        WindowsPackage = 'Negotiate'; CredentialFile = ''
        SearchIn = @('Organizer', 'Rooms'); PastDays = 0; FutureDays = 365; SeriesScope = 'Whole'; Rooms = @(); RoomFile = ''
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
    if ([string]$c.SeriesScope -notin $script:SeriesScopes) { $p.Add("Search.SeriesScope must be 'Whole' or 'Occurrences'.") }
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
    <#
        What one run does, with the defaults of the configuration. Mode: 'Organizers' or 'Rooms' (every meeting of
        the rooms given, whatever its organizer). SeriesScope: 'Whole' (a series is acted on whole) or 'Occurrences'
        (only its occurrences in the period); always 'Occurrences' in rooms mode.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Settings, [string[]]$Organizer, [string]$OrganizerFile,
        [string[]]$Room, [string]$RoomFile, [Nullable[datetime]]$Start, [Nullable[datetime]]$End,
        [string]$Subject, [string[]]$MeetingId, [string[]]$SearchIn, [string[]]$Mailbox, [string]$MailboxFile,
        [ValidateSet('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')][string]$Action = 'Report',
        [string]$Comment, [string]$FromReport, [string]$NewOrganizer, [string]$SeriesScope
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
        PeriodGiven = $null -ne $Start -and $null -ne $End
        Subject = [string]$Subject; MeetingId = @($MeetingId | ForEach-Object { $_ -split '[;,\s]+' } | Where-Object { $_ } | ForEach-Object ToUpperInvariant | Select-Object -Unique)
        SearchIn = if (-not $org.Count -and $rooms.Count) { @('Rooms') } elseif ($SearchIn) { @($SearchIn | Select-Object -Unique) } else { @($Settings.SearchIn) }
        SeriesScope = if (-not $org.Count -and $rooms.Count) { 'Occurrences' } elseif ($SeriesScope) { $SeriesScope } elseif ($Settings.ContainsKey('SeriesScope') -and $Settings.SeriesScope) { [string]$Settings.SeriesScope } else { 'Whole' }
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
    $scope = [string](Get-McoProperty $Request 'SeriesScope')
    if ($scope -and $scope -notin $script:SeriesScopes) { $p.Add("Series scope must be 'Whole' or 'Occurrences'.") }
    if ($Request.Mode -eq 'Organizers' -and $scope -eq 'Occurrences' -and -not $Request.FromReport -and $Request.Action -in 'Remove', 'Cancel' -and -not (Get-McoProperty $Request 'PeriodGiven')) { $p.Add('Series by occurrences: give the period of the action (-Start and -End): the occurrences of the series in it are acted on.') }
    if ($Request.Action -eq 'Restore' -and -not $Request.FromReport) { $p.Add('Restore needs -FromReport.') }
    if ($Request.FromReport -and $Request.Action -eq 'Report') { $p.Add('A report replay needs -Action Remove, Cancel, Transfer or Restore.') }
    if ($Request.Action -eq 'Transfer' -and -not $Request.NewOrganizer) { $p.Add('Transfer needs -NewOrganizer.') }
    if ($Request.Action -eq 'Transfer' -and $Request.Mode -eq 'Organizers' -and $scope -eq 'Occurrences' -and -not $Request.FromReport) { $p.Add('Transfer moves whole series: -SeriesScope Occurrences cannot be transferred (use -SeriesScope Whole).') }
    [pscustomobject]@{ IsValid = $p.Count -eq 0; Problems = @($p) }
}

# SIG # Begin signature block
# MIImJAYJKoZIhvcNAQcCoIImFTCCJhECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAeINwZIs4bmrHO
# KuN/vKNJTO1i8MBHEt1A7I9ibqk+jKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYCzCCGAcCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDhVkJZK+2p
# r9jLAuz79pisTxZY3RXoKRU8m5vq30SLHDANBgkqhkiG9w0BAQEFAASCAgADrG6k
# WqQVwFB8zrMskpalBSe9a7sl+SUFjHGR9G8RlcJ96GpWP60TZvjMJwRx0BRPCrVn
# S4VE9zlZztWH6xmD6BsH/uwYmRI6zhDzbaWM/swhN68PbxEqvJFidIiG/KR/q4UH
# 1Xq/0U3ZzaHV/N/sx7mArcPIJMeR7FoJX3/oG/fc8vppr/PPuU9Dg09Ifled8haf
# c0LAbMkauLvuTSmAjJRDWDUJB2mKXKMSvgg9AKFdPSdC56g0oIwqB8l/PZhq/g4e
# YQkP5V8iktjYnoPnle7ZPS2EliqR6xymYTk9GXXsTW7DnK7xp9GSPIoXwk2CqaUF
# vrS+Uaz3sXHfxRgV2wJMeyxSRAfhEVhzl0AQvubUW6MYPzU3c7q/S+X8fnH4vK+2
# yRgk+t5yH/5eK0Q8eqTg9ppCNlHwlO39Ch0brt4yy/+8O6KgO7UOfkbxljbp8XYn
# Wg1BgpSjbnBJoAnMRRqo19mK2w6edXT9LE++Y/uGb6qHd3vPQWv8JbNVxtjm+Hub
# Ya0oOQpZoBrPD4IYTZ56POd28Ivngyv/cbyGqG9lgZhiMd6nn1D2nDAlidzDdXMC
# zc3CVui5BTqKR1RPhzhQuRo0UgORkoqq+pQUeNYQBiUB8IDealbjAkGBGgCZFHDp
# Dz9mf3QWtauc8UfAVaM1X+2T1EGIt576ilAyRaGCFMgwghTEBgorBgEEAYI3AwMB
# MYIUtDCCFLAGCSqGSIb3DQEHAqCCFKEwghSdAgEDMQ0wCwYJYIZIAWUDBAIBMIGC
# BgsqhkiG9w0BCRABBKBzBHEwbwIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEINN47ZR46iv2/qHK6vwFQNALUwfBI6OvnW8d/rn/Mi2GAghkX3dT+xX8
# 0xgPMjAyNjExMDUxNzIyMzZaMAMCAQECCQCwgXal58f8nKCCEacwggTzMIIC26AD
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
# VzCCAlMCAQEwgYYwcjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO
# 4zALBglghkgBZQMEAgGgggFgMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAc
# BgkqhkiG9w0BCQUxDxcNMjYxMTA1MTcyMjM2WjAoBgkqhkiG9w0BCTQxGzAZMAsG
# CWCGSAFlAwQCAaEKBggqhkjOPQQDAjAvBgkqhkiG9w0BCQQxIgQgtHxTcm1SGlcn
# o3xKFmxaD+D6hGUOlYEl1vskbxRwykcwgcgGCyqGSIb3DQEJEAIvMYG4MIG1MIGy
# MIGvBCBXT+oXqvpefGormwjJmLFzXkULOHPKROaiTKCTDmPzsDCBijB2pHQwcjEL
# MAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREw
# DwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcg
# SXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO4zAKBggqhkjOPQQDAgRI
# MEYCIQCN+Y3jlFANyu8XeNKkz3eYE+JZ5GMxwAs+W2i/r3YRSwIhAJ0C8G5VHezE
# x+N7ORDuiYhKFDX7cR038QzBPsZl2y1H
# SIG # End signature block

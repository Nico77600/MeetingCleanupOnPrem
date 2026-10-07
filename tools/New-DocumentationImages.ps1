<#
.SYNOPSIS
    Renders the images of the guides and the README: the console of a run and of a search in progress, and the HTML
    report (light, dark, Transfers tab), from fictitious data.

.DESCRIPTION
    No Exchange server and no real data: the meetings come from the simulated Exchange of the tests
    (tests\FakeEws.ps1), with contoso.com names; the cmdlets of Exchange PowerShell are replaced by fictitious
    answers. The console is captured with its colours (MCO_FORCE_COLOR=1, MCO_ICONS=Emoji) in a second PowerShell
    process, turned into a page with the colours and the font of Windows Terminal; the pages are opened by Microsoft
    Edge headless.

    Writes docs\images\console-run.png, console-progress.png, report-overview.png, report-dark.png and
    report-transfers.png. Needs Microsoft Edge.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
    Part of : Meeting Cleanup On-Prem (repository tool, not in the package)
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$Destination = (Join-Path $PSScriptRoot '..\docs\images'),
    # Internal: the console of a run (Run) or of a search in progress (Progress), written to -OutFile.
    [ValidateSet('', 'Run', 'Progress')][string]$Console = '',
    [string]$OutFile,
    [string]$Work
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
# Paths shown in the console and the reports: those of a real installation.
$shownRoot = 'C:\Tools\MeetingCleanupOnPrem'
# The simulated Exchange, at the level of the script: it answers for as long as the script runs.
. (Join-Path $root 'tests\FakeEws.ps1')

function Initialize-DocOrganization {
    <# The module, the simulated Exchange and the fictitious organization of the images; returns the settings. #>
    param([string]$Work, [datetime]$From)
    Import-Module (Join-Path $root 'MeetingCleanupOnPrem.psd1') -Force
    $d = 'contoso.com'
    $s = New-FakeStore
    $people = [ordered]@{ 'megan.bowen' = 'Megan Bowen'; 'alex.wilber' = 'Alex Wilber'; 'lidia.holloway' = 'Lidia Holloway'; 'adele.vance' = 'Adele Vance'; 'joni.sherman' = 'Joni Sherman'; 'lee.gu' = 'Lee Gu'; 'nestor.wilke' = 'Nestor Wilke'; 'lynne.robbins' = 'Lynne Robbins' }
    foreach ($p in $people.Keys) { $s.Names["$p@$d"] = $people[$p] }
    $rooms = @("room-paris-01@$d", "room-paris-02@$d", "room-lyon-01@$d")
    [void](Add-FakeMeeting $s 'Weekly sales review' "megan.bowen@$d" @("alex.wilber@$d", "joni.sherman@$d", "lee.gu@$d") -Rooms @("room-paris-01@$d") -Start $From.AddDays(4).AddHours(8.5) -Minutes 60 -Weeks 12)
    [void](Add-FakeMeeting $s 'Q1 budget workshop' "megan.bowen@$d" @("lidia.holloway@$d", "adele.vance@$d", "partner@fabrikam.com") -Rooms @("room-paris-02@$d", "room-lyon-01@$d") -Start $From.AddDays(13).AddHours(13) -Minutes 120)
    [void](Add-FakeMeeting $s 'Project Atlas kick-off' "megan.bowen@$d" @("nestor.wilke@$d", "adele.vance@$d") -Rooms @("room-lyon-01@$d") -Start $From.AddDays(19).AddHours(9) -Minutes 60)
    [void](Add-FakeMeeting $s '1:1 Alex / Megan' "megan.bowen@$d" @("alex.wilber@$d") -Start $From.AddDays(7).AddHours(16) -Minutes 30)
    # Lynne Robbins has left, her mailbox is deleted: her meeting is found in a room and in the attendees' calendars.
    [void](Add-FakeMeeting $s 'Supplier quarterly review' "lynne.robbins@$d" @("adele.vance@$d", "lee.gu@$d") -Rooms @("room-paris-02@$d") -Start $From.AddDays(33).AddHours(14) -Minutes 60 -NoOrganizerCopy)
    [void]$s.DenyMailbox.Add("lynne.robbins@$d")
    [void](Add-FakeMeeting $s 'Not Megan''s meeting' "alex.wilber@$d" @("megan.bowen@$d") -Rooms @("room-paris-01@$d") -Start $From.AddDays(7).AddHours(10) -Minutes 60)

    $settings = Import-McoConfiguration -Path (Join-Path $root 'config\MeetingCleanupOnPrem.config.psd1') -Root $root
    $settings.EwsUrl = "https://mail.$d/EWS/Exchange.asmx"; $settings.Mailbox = "svc-meetingcleanup@$d"; $settings.ManagementShellServer = "mail.$d"
    $settings.AcceptedDomains = @($d); $settings.Rooms = @(); $settings.TimeZone = 'Romance Standard Time'; $settings.CredentialFile = 'C:\Tools\MeetingCleanupOnPrem\svc.cred.xml'
    $settings.OutputPath = Join-Path $Work 'reports'; $settings.LogPath = Join-Path $Work 'logs'
    Install-FakeEws -Store $s -Settings $settings
    # Exchange PowerShell: fictitious answers for the directory (names, rooms, the group Sales team).
    & (Get-Module MeetingCleanupOnPrem) {
        param($people, $rooms, $d)
        $script:DocPeople = $people; $script:DocRooms = $rooms; $script:DocDomain = $d
        function script:Resolve-McoExchangeRecipient {
            param([string]$Identity)
            $alias = $Identity.Split('@')[0].ToLowerInvariant()
            $name = $script:DocPeople[$alias]
            if (-not $name) { return $null }
            if ($alias -eq 'lynne.robbins') { return [pscustomobject]@{ Input = $Identity; PrimaryAddress = ''; DisplayName = ''; Addresses = @(); RecipientTypeDetails = ''; Alias = ''; Error = "The operation couldn't be performed because object '$Identity' couldn't be found on 'DC1.contoso.com'." } }
            $primary = "$alias@$($script:DocDomain)"
            [pscustomobject]@{ Input = $Identity; PrimaryAddress = $primary; DisplayName = $name; Addresses = @($primary, "/o=contoso/ou=exchange administrative group (fydibohf23spdlt)/cn=recipients/cn=$alias"); RecipientTypeDetails = 'UserMailbox'; Alias = $alias; Error = '' }
        }
        function script:Get-McoExchangeMailboxAddresses { param([string[]]$RecipientTypeDetails) [pscustomobject]@{ Available = $true; Addresses = @($script:DocRooms); Error = '' } }
        function script:Get-McoRecipientKind { param([string]$Address) if ($script:DocRooms -contains $Address.ToLowerInvariant()) { 'Room' } else { 'Mailbox' } }
        function script:Get-McoExchangeGroupMembers { param([string]$Identity) @() }
        $script:Ews.ServerVersion = 'V2017_07_11'
    } $people $rooms $d
    $s.Ids = @{}
    [pscustomobject]@{ Settings = $settings; Store = $s; Domain = $d }
}

$year = (Get-Date).Year + 1
$from = [datetime]::new($year, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
$period = @{ Start = [datetime]::new($year, 1, 1); End = [datetime]::new($year, 3, 31) }

# ---- internal: the console of a run, in its own process (colours forced, emoji) -----------------------------
if ($Console) {
    [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
    $org = Initialize-DocOrganization -Work $Work -From $from
    $settings = $org.Settings
    $m = Get-Module MeetingCleanupOnPrem
    & $m { $script:Quiet = $false }
    $request = New-McoRequest -Settings $settings -Organizer "megan.bowen@$($org.Domain)" @period
    Write-McoRunBanner -Settings $settings -Request $request -LogPath "$shownRoot\logs\MeetingCleanupOnPrem_$(Get-Date -Format yyyyMMdd).log"
    Initialize-McoSteps -Total 6
    Write-McoNextStep 'Exchange Server' 'Server'
    Write-McoItem Ok "Exchange PowerShell: Exchange cmdlets imported from http://mail.$($org.Domain)/PowerShell/." -Icon Shield
    Write-McoItem Ok "EWS https://mail.$($org.Domain)/EWS/Exchange.asmx $([char]0x00B7) V2017_07_11 $([char]0x00B7) Impersonation (svc-meetingcleanup@$($org.Domain) signs in)" -Icon Server
    if ($Console -eq 'Progress') {
        # A search of a large organization stopped at two thirds: the live line, as an interactive console shows it.
        Write-McoNextStep 'Organizers' 'User'
        Write-McoItem Ok "Megan Bowen <megan.bowen@$($org.Domain)> $([char]0x00B7) resolved by Exchange Management Shell $([char]0x00B7) 2 addresses compared" -Icon User
        Write-McoNextStep 'Mailboxes to search' 'Search'
        Write-McoItem Info 'Organizer: 1 mailbox(es)' -Icon Mail
        Write-McoItem Info 'Rooms: 1,860 mailbox(es)' -Icon Room
        Write-McoNextStep 'Search' 'Calendar'
        & $m {
            $C = $script:C; $percent = 67; $filled = [int][Math]::Round(12 * $percent / 100.0)
            $bar = $C.Accent + [string]::new([char]0x2588, $filled) + $C.Dim + [string]::new([char]0x2591, 12 - $filled) + $C.Reset
            Write-Host ("      {0}{1} {2,3}%  {3}{4}{5}" -f (Get-McoIcon 'Clock'), $bar, $percent, $C.Dim, "1,254/1,861 mailboxes searched $([char]0x00B7) about 20 s left", $C.Reset)
        }
        return
    }
    $result = Find-McoMeetings -Settings $settings -Request $request
    Write-McoMeetingTable -Meetings @($result.Meetings)
    Write-McoNextStep 'Report' 'Report'
    $report = Export-McoReport -Result $result -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats $settings.ReportFormats
    foreach ($f in $report.Files.Values) { Write-McoItem Ok $f -Icon File }
    Write-McoRunSummary -Result $result -ReportText $report.Files.Html -LogPath "$shownRoot\logs\MeetingCleanupOnPrem_$(Get-Date -Format yyyyMMdd).log"
    return
}

# ---- the console: ANSI colours to a page with the look of Windows Terminal -----------------------------------
function ConvertTo-ConsoleHtml {
    param([string]$Text)
    $esc = [char]27
    $sb = [Text.StringBuilder]::new()
    $state = @{ Fg = ''; Bg = ''; Bold = $false }
    $open = $false
    $emit = {
        param([string]$s)
        foreach ($rune in $s.EnumerateRunes()) {
            $v = $rune.Value
            $t = [Net.WebUtility]::HtmlEncode($rune.ToString())
            # Emoji: two cells wide, as in the terminal (box drawing and blocks stay one cell).
            if ($v -ge 0x1F000 -or ($v -ge 0x2300 -and $v -le 0x23FF) -or ($v -ge 0x2600 -and $v -le 0x27BF)) { [void]$sb.Append("<span class=""e"">$t</span>") }
            elseif ($v -eq 0xFE0F) { }
            else { [void]$sb.Append($t) }
        }
    }
    foreach ($part in [regex]::Split($Text, "($esc\[[\d;]*m)")) {
        if ($part -match "^$esc\[([\d;]*)m$") {
            $codes = @($Matches[1].Split(';') | Where-Object { $_ -ne '' } | ForEach-Object { [int]$_ })
            if (-not $codes.Count) { $codes = @(0) }
            for ($i = 0; $i -lt $codes.Count; $i++) {
                switch ($codes[$i]) {
                    0 { $state = @{ Fg = ''; Bg = ''; Bold = $false } }
                    1 { $state.Bold = $true }
                    90 { $state.Fg = '#8b949e' }
                    97 { $state.Fg = '#f2f2f2' }
                    38 { if ($codes[$i + 1] -eq 2) { $state.Fg = '#{0:x2}{1:x2}{2:x2}' -f $codes[$i + 2], $codes[$i + 3], $codes[$i + 4]; $i += 4 } }
                    48 { if ($codes[$i + 1] -eq 2) { $state.Bg = '#{0:x2}{1:x2}{2:x2}' -f $codes[$i + 2], $codes[$i + 3], $codes[$i + 4]; $i += 4 } }
                }
            }
            if ($open) { [void]$sb.Append('</span>'); $open = $false }
            $style = @()
            if ($state.Fg) { $style += "color:$($state.Fg)" }
            if ($state.Bg) { $style += "background:$($state.Bg)" }
            if ($state.Bold) { $style += 'font-weight:700' }
            if ($style.Count) { [void]$sb.Append("<span style=""$($style -join ';')"">"); $open = $true }
            continue
        }
        & $emit $part
    }
    if ($open) { [void]$sb.Append('</span>') }
    $sb.ToString()
}

function Save-EdgeScreenshot {
    param([string]$Url, [string]$Png, [string]$Size)
    $profile = Join-Path $work "edge-$([guid]::NewGuid().ToString('N'))"
    $before = if (Test-Path -LiteralPath $Png) { (Get-Item -LiteralPath $Png).LastWriteTimeUtc } else { [datetime]::MinValue }
    # Edge headless sometimes stays open after writing its screenshot: waited for 60 s at most, then stopped.
    $p = Start-Process -FilePath $edge -PassThru -WindowStyle Hidden -ArgumentList @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profile`"", "--window-size=$Size", "--screenshot=`"$Png`"", "`"$Url`"")
    if (-not $p.WaitForExit(60000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path -LiteralPath $Png) -or (Get-Item -LiteralPath $Png).LastWriteTimeUtc -le $before) { Write-Warning "Edge did not write $Png." }
    Start-Sleep -Milliseconds 500
}

$Destination = [IO.Path]::GetFullPath($Destination)
[void][IO.Directory]::CreateDirectory($Destination)
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge is needed for the images.' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('mco-docimages-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
[void][IO.Directory]::CreateDirectory($work)
try {
    # ---- console images: a second process, its output redirected, colours and emoji forced ----------------------
    foreach ($kind in 'Run', 'Progress') {
        $out = Join-Path $work "console-$kind.txt"
        $env:MCO_FORCE_COLOR = '1'; $env:MCO_ICONS = 'Emoji'
        try { & pwsh -NoProfile -File $PSCommandPath -Console $kind -Work $work *> $out }
        finally { Remove-Item Env:\MCO_FORCE_COLOR, Env:\MCO_ICONS -ErrorAction SilentlyContinue }
        $text = [IO.File]::ReadAllText($out).Replace((Join-Path $work 'reports'), "$shownRoot\reports").TrimEnd()
        # Lines of a console 150 columns wide: a longer line wraps, as in the terminal.
        $lines = 0; foreach ($l in ($text -split "`r?`n")) { $lines += [Math]::Max(1, [Math]::Ceiling(([regex]::Replace($l, "$([char]27)\[[\d;]*m", '')).Length / 150)) }
        $html = "<!doctype html><html><head><meta charset=""utf-8""><style>html,body{margin:0;background:#0c0c0c}" +
            "pre{margin:0;padding:18px 22px;width:150ch;color:#cccccc;font:14px/19px 'Cascadia Mono','Cascadia Code',Consolas,'Segoe UI Emoji',monospace;white-space:pre-wrap;overflow-wrap:anywhere}" +
            ".e{display:inline-block;width:2ch;font-family:'Segoe UI Emoji';line-height:1}</style></head><body><pre>$(ConvertTo-ConsoleHtml $text)</pre></body></html>"
        $page = Join-Path $work "console-$kind.html"
        [IO.File]::WriteAllText($page, $html, [Text.UTF8Encoding]::new($false))
        Save-EdgeScreenshot -Url ([Uri]$page).AbsoluteUri -Png (Join-Path $Destination "console-$($kind.ToLowerInvariant()).png") -Size "1290,$(44 + 19 * $lines)"
    }

    # ---- reports: a search, a cancellation, a transfer ----------------------------------------------------------
    $org = Initialize-DocOrganization -Work $work -From $from
    $settings = $org.Settings
    & (Get-Module MeetingCleanupOnPrem) { $script:Quiet = $true }
    $d = $org.Domain
    $request = New-McoRequest -Settings $settings -Organizer "megan.bowen@$d", "lynne.robbins@$d" @period
    $found = Find-McoMeetings -Settings $settings -Request $request
    $search = Export-McoReport -Result $found -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats Html

    foreach ($m in $found.Meetings) { if ($m.Subject -like '1:1*') { $m.Selected = $false } }
    $cancelled = Invoke-McoCleanup -Settings $settings -Result $found -Action Cancel -Comment 'Megan Bowen has left Contoso: this meeting is cancelled. Contact Alex Wilber for the follow-up.'
    $cancel = Export-McoReport -Result $cancelled -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats Html

    # The transfer: meetings to come (the simulated clock is today).
    $org = Initialize-DocOrganization -Work $work -From ([datetime]::UtcNow.Date.AddDays(3))
    $settings = $org.Settings
    & (Get-Module MeetingCleanupOnPrem) { $script:Quiet = $true }
    $request = New-McoRequest -Settings $settings -Organizer "megan.bowen@$d", "lynne.robbins@$d" -Start ([datetime]::UtcNow.Date) -End ([datetime]::UtcNow.Date.AddDays(90))
    $found = Find-McoMeetings -Settings $settings -Request $request
    foreach ($m in $found.Meetings) { $m.Selected = $m.Subject -in 'Weekly sales review', 'Q1 budget workshop', 'Supplier quarterly review' }
    $target = Resolve-McoNewOrganizer -Address "alex.wilber@$d"
    $plan = Get-McoTransferPlan -Result $found -NewOrganizer $target -Comment 'This meeting is now organized by {0}.'
    $moved = Invoke-McoTransfer -Settings $settings -Result $found -Plan $plan -Comment 'This meeting is now organized by {0}.'
    $transfer = Export-McoReport -Result $moved -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats Html

    foreach ($shot in @(
            @{ Html = $search.Files.Html; Query = '?scoutTheme=light'; File = 'report-overview.png'; Size = '1360,1180' }
            @{ Html = $cancel.Files.Html; Query = '?scoutTheme=dark'; File = 'report-dark.png'; Size = '1360,1180' }
            @{ Html = $transfer.Files.Html; Query = '?scoutTheme=light&scoutFocus=tables'; File = 'report-transfers.png'; Size = '1360,600' })) {
        Save-EdgeScreenshot -Url (([Uri]$shot.Html).AbsoluteUri + $shot.Query) -Png (Join-Path $Destination $shot.File) -Size $shot.Size
    }
}
finally {
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}
Get-ChildItem -LiteralPath $Destination -Filter '*.png' | Where-Object Name -match '^(console|report)-' | Sort-Object Name | ForEach-Object { '{0,10:N0}  {1}' -f $_.Length, $_.Name }

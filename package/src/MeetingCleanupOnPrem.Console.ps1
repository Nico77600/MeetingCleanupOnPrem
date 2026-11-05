<#
.SYNOPSIS
    Meeting Cleanup On-Prem - console output and log file (dot-sourced by MeetingCleanupOnPrem.psm1).

.DESCRIPTION
    Same rules as Meeting Cleanup 1.2.2:
      - ANSI colours are disabled when the output is redirected or NO_COLOR is set; MCO_FORCE_COLOR=1 forces them.
      - Icons: emoji in Windows Terminal / VS Code, symbols of the classic console fonts elsewhere.
        MCO_ICONS = Emoji | Symbols | Ascii forces a style.
      - A live progress line (bar, part done, count, time left) is rewritten in place in an interactive console.
      - Every line shown is also written to the daily log file, without colours or icons.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.1
#>

$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Green = ''; Yellow = ''; Red = ''; Blue = ''; White = '' }
if ($env:MCO_FORCE_COLOR -eq '1' -or (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR)) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"; Blue = "$e[38;2;110;170;240m"
    }
}
$script:IconStyle = if ($env:MCO_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:MCO_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }
$script:Dot = [char]0x00B7
$script:ProgressShown = $false
# The progress in course (Get-McoProgressEta): its label, when and where it began.
$script:ProgressEta = $null

function Get-McoIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' {
            return @{
                Logo = & $u 0x1F4C5; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C; Info = & $u 0x1F539
                Skip = & $u 0x23E9; Key = & $u 0x1F511; Server = & $u 0x1F5A5; Shield = & $u 0x1F512; Room = & $u 0x1F3E2
                Mail = & $u 0x1F4E8; People = & $u 0x1F465; User = & $u 0x1F464; File = & $u 0x1F4C4; Log = & $u 0x1F4DD
                Report = & $u 0x1F4CA; Done = & $u 0x1F389; Target = & $u 0x1F3AF; Search = & $u 0x1F50E; Clock = & $u 0x23F3
                Calendar = & $u 0x1F4C6; Trash = (& $u 0x1F5D1) + [char]0xFE0F; Cancel = & $u 0x1F6AB; Cloud = (& $u 0x2601) + [char]0xFE0F; Refresh = & $u 0x1F504
            }
        }
        'Symbols' {
            return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Key = & $u 0x00A7; Server = & $u 0x2261; Shield = & $u 0x25CA; Room = & $u 0x2302
                Mail = '@'; People = & $u 0x2192; User = & $u 0x263A; File = & $u 0x25AC; Log = & $u 0x00B6
                Report = & $u 0x2261; Done = & $u 0x221A; Target = & $u 0x25D9; Search = & $u 0x25BA; Clock = & $u 0x25CB
                Calendar = & $u 0x25A1; Trash = & $u 0x00D7; Cancel = & $u 0x00F8; Cloud = & $u 0x2248; Refresh = & $u 0x00AB
            }
        }
        default {
            return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Key = 'k'; Server = '='; Shield = 'o'; Room = '#'
                Mail = '@'; People = '&'; User = 'u'; File = '-'; Log = '='; Report = '='; Done = '*'; Target = 'o'; Search = '?'
                Clock = '~'; Calendar = '#'; Trash = 'x'; Cancel = '/'; Cloud = '~'; Refresh = 'r'
            }
        }
    }
}

function Get-McoFrameSet {
    <# Rounded corners in modern terminals (emoji style), square corners elsewhere (present in every console font). #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    if ($Style -eq 'Ascii') {
        return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' }
    }
    if ($Style -eq 'Symbols') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

$script:Icons = Get-McoIconSet $script:IconStyle
$script:Frame = Get-McoFrameSet $script:IconStyle
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
$script:IconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }

function Get-McoIcon { param([Parameter(Mandatory = $true)][string]$Name) return $script:Icons[$Name] + $script:IconPad }

function Format-McoDuration {
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalHours -ge 1) { return [string]::Format($inv, '{0} h {1:00} min', [int][Math]::Floor($t.TotalHours), $t.Minutes) }
    if ($t.TotalMinutes -ge 1) { return [string]::Format($inv, '{0} min {1:00} s', $t.Minutes, $t.Seconds) }
    return [string]::Format($inv, '{0:0.0} s', $t.TotalSeconds)
}

function Format-McoText {
    <# Text cut to a width with an ellipsis, padded to the width. #>
    param([AllowEmptyString()][AllowNull()][string]$Text, [int]$Width)
    $t = [string]$Text -replace '[\r\n\t]+', ' '
    if ($Width -le 0) { return $t }
    if ($t.Length -gt $Width) { return $t.Substring(0, [Math]::Max(0, $Width - 1)) + [char]0x2026 }
    return $t.PadRight($Width)
}

function Send-McoUi {
    <#
        Forwards a console line to a window run, when one is in progress (none yet in On-Prem): into the queue the
        window reads (Ui.Queue) or to its sink (Ui.Sink).
    #>
    param([string]$Status, [string]$Text)
    $u = $script:Ui
    if (-not $u) { return }
    if ($u.Queue) { $u.Queue.Enqueue([string[]]@($Status, $Text)) }
    elseif ($u.Sink) { & $u.Sink $Status $Text }
}

function Start-McoLog {
    <# Opens (or continues) today's log file and deletes the log files older than the retention. #>
    param([Parameter(Mandatory = $true)][string]$Directory, [int]$RetentionDays = 30)

    Stop-McoLog
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('MeetingCleanupOnPrem_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    $writer.AutoFlush = $true
    $script:LogWriter = [IO.TextWriter]::Synchronized($writer)
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'MeetingCleanupOnPrem_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-McoLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-McoLog {
    <# One line in the log file only. The log never contains colours, icons, passwords or secrets. #>
    param(
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message
    )
    if ($script:LogWriter) { $script:LogWriter.WriteLine(('{0:yyyy-MM-ddTHH:mm:ss.fffzzz} [{1,-5}] {2}' -f (Get-Date), $Level, $Message)) }
}

function Get-McoConsoleWidth {
    try { $w = [Console]::WindowWidth; if ($w -ge 40) { return $w } } catch { }
    return 120
}

function Clear-McoProgress {
    <# Ends the live progress line (if any) so that the next line starts on a new row. #>
    if ($script:ProgressShown) {
        [Console]::Write("`r" + (' ' * [Math]::Max(10, (Get-McoConsoleWidth) - 1)) + "`r")
        $script:ProgressShown = $false
    }
}

function Write-McoBanner {
    <# Title card at the start of an execution, followed by the context rows (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Subtitle,
        [System.Collections.Specialized.OrderedDictionary]$Details
    )

    Write-McoLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $v = $Details[$key]
            Write-McoLog 'INFO' ('{0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
        }
    }
    if ($script:Quiet) { return }
    $C = $script:C; $F = $script:Frame; $width = 78
    $right = "v$($script:ToolVersion) $($script:Dot) Nicolas Fabert"
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $script:IconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.TopLeft, [string]::new($F.Horizontal, $width), $F.TopRight, $C.Reset)
    Write-Host ('  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, [string]::new(' ', $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        Write-Host ('  {0}{1}{2}{3}{4}{5}{0}{6}{2}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, (Format-McoText "     $Subtitle" $width), $C.Reset, $F.Vertical)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon, $text = if ($value -is [array]) { (Get-McoIcon $value[0]), $value[1] } else { '   ', $value }
            Write-Host ('     {0}{1}{2,-11}{3} {4}' -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
}

function Write-McoStep {
    <# Step header with a coloured number pill and an icon:  ─ 3/6 ─ 🔎  Search #>
    param(
        [Parameter(Mandatory = $true)][int]$Number,
        [Parameter(Mandatory = $true)][int]$Total,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Icon = 'Info'
    )

    Write-McoLog 'STEP' "[$Number/$Total] $Title"
    $script:ProgressEta = $null
    Send-McoUi 'Step' "[$Number/$Total] $Title"
    if ($script:Quiet) { return }
    Clear-McoProgress
    $C = $script:C
    Write-Host ''
    Write-Host ('  {0} {1}/{2} {3} {4}{5}{6}{3}' -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-McoIcon $Icon), $C.Bold, $Title)
}

function Write-McoItem {
    <# One indented result line with a status icon, also written to the log (and to a window run). #>
    param(
        [ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip', 'Block')][string]$Status = 'Info',
        [Parameter(Mandatory = $true)][Alias('Message')][AllowEmptyString()][string]$Text,
        [string]$Icon
    )

    if ($Status -eq 'Block') { $Status = 'Warn' }
    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO' }[$Status]
    Write-McoLog $level $Text
    Send-McoUi $Status $Text
    if ($script:Quiet) { return }
    Clear-McoProgress
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim }[$Status]
    $symbol = Get-McoIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip') { $color } else { '' }
    Write-Host ('      {0}{1}{2}{3}{4}{2}' -f $color, $symbol, $script:C.Reset, $textColor, $Text)
}

function Format-McoTimeLeft {
    <#
        The time left of a progress, rounded as a person would say it: a few seconds, about 25 s (5 s steps under a
        minute), about 1 min 30 s (10 s steps under 5 minutes), about 12 min, about 1 h 05 min.
    #>
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    if ($Seconds -lt 10) { return 'a few seconds left' }
    $away = [MidpointRounding]::AwayFromZero
    $r = [int]$(if ($Seconds -lt 60) { [Math]::Ceiling($Seconds / 5) * 5 } elseif ($Seconds -lt 300) { [Math]::Round($Seconds / 10, $away) * 10 } else { [Math]::Round($Seconds / 60, $away) * 60 })
    if ($r -lt 60) { return [string]::Format($inv, 'about {0} s left', $r) }
    $h = [int][Math]::Floor($r / 3600); $m = [int][Math]::Floor(($r % 3600) / 60); $s = $r % 60
    if ($h) { return [string]::Format($inv, 'about {0} h {1:00} min left', $h, $m) }
    if ($s) { return [string]::Format($inv, 'about {0} min {1:00} s left', $m, $s) }
    return [string]::Format($inv, 'about {0} min left', $m)
}

function Get-McoProgressEta {
    <#
        Time left of the progress in course, from its speed since it began; empty until it can be told (2 s and
        2 % of progress since its first value). Another label (the counts aside), or a value going back, is a
        new progress. A step starts with none (Write-McoStep). -Now: tests.
    #>
    param([Parameter(Mandatory = $true)][double]$Fraction, [AllowEmptyString()][string]$Text, [datetime]$Now = [datetime]::UtcNow)

    $key = $Text -replace '[\d\s,.\u00A0\u202F/]+', ''
    $s = $script:ProgressEta
    if (-not $s -or $s.Key -ne $key -or $Fraction -lt $s.Last) {
        $script:ProgressEta = @{ Key = $key; Start = $Now; From = $Fraction; Last = $Fraction; Left = -1.0; At = $Now }
        return ''
    }
    $s.Last = $Fraction
    $done = $Fraction - $s.From
    $elapsed = ($Now - $s.Start).TotalSeconds
    if ($Fraction -ge 1 -or $done -lt 0.02 -or $elapsed -lt 2) { return '' }
    $left = $elapsed / $done * (1 - $Fraction)
    # EWS answers do not all take the same time (a mailbox with many meetings, a GetItem of 50 items, the 3 s
    # between the removal waves): half the new figure, half the last one brought forward.
    if ($s.Left -ge 0) { $left = 0.5 * $left + 0.5 * [Math]::Max(0.0, $s.Left - ($Now - $s.At).TotalSeconds) }
    $s.Left = $left; $s.At = $Now
    return Format-McoTimeLeft $left
}

function Write-McoProgress {
    <#
        Live progress line, rewritten in place (interactive console); sent to a window run (fraction|text|time left).
              ⏳  ███████░░░░░  58%  1,077/1,858 mailboxes searched · about 40 s left
    #>
    param([Parameter(Mandatory = $true)][double]$Fraction, [Parameter(Mandatory = $true)][string]$Text)

    $left = Get-McoProgressEta -Fraction $Fraction -Text $Text
    Send-McoUi 'Progress' ('{0}|{1}|{2}' -f $Fraction.ToString('0.000', [Globalization.CultureInfo]::InvariantCulture), $Text, $left)
    if ($script:Quiet -or [Console]::IsOutputRedirected) { return }
    $C = $script:C
    $percent = [int][Math]::Floor(100 * [Math]::Min(1.0, [Math]::Max(0.0, $Fraction)))
    $filled = [int][Math]::Round(12 * $percent / 100.0)
    $bar = $C.Accent + [string]::new([char]0x2588, $filled) + $C.Dim + [string]::new([char]0x2591, 12 - $filled) + $C.Reset
    $line = if ($left) { "$Text $($script:Dot) $left" } else { $Text }
    $plain = Format-McoText $line ([Math]::Max(10, (Get-McoConsoleWidth) - 30))
    [Console]::Write(("`r      {0}{1} {2,3}%  {3}{4}{5}" -f (Get-McoIcon 'Clock'), $bar, $percent, $C.Dim, $plain.TrimEnd(), $C.Reset))
    $script:ProgressShown = $true
}

function Write-McoTable {
    <#
        Aligned table, one row per object, with a status icon in front of each row.
        Columns: @{ Name = 'Header'; Property = 'PropertyName'; Width = 20; Align = 'Right' } - Width 0 = the rest of the console.
        StatusProperty: Ok | Warn | Fail | Info | Skip (colour and icon of the row).
    #>
    param(
        [Parameter(Mandatory = $true)][object[]]$Columns,
        [AllowEmptyCollection()][AllowNull()][object[]]$Rows,
        [string]$StatusProperty = 'Status',
        [int]$Indent = 6,
        [int]$MaxWidth = 170
    )

    if (-not $Rows -or -not $Rows.Count) { return }
    foreach ($row in $Rows) { Write-McoLog 'INFO' (($Columns | ForEach-Object { "$($_.Name)=$([string]$row.($_.Property))" }) -join ' | ') }
    if ($script:Quiet) { return }
    Clear-McoProgress
    $C = $script:C
    $consoleWidth = [Math]::Min($MaxWidth, (Get-McoConsoleWidth) - 1)
    if ($consoleWidth -lt 80) { $consoleWidth = 120 }
    $fixed = [int](($Columns | ForEach-Object { [int]$_['Width'] } | Measure-Object -Sum).Sum) + 2 * $Columns.Count
    $last = [Math]::Max(20, $consoleWidth - $Indent - 3 - $fixed)
    $pad = ' ' * $Indent
    $cell = {
        param($col, $text)
        $w = if ([int]$col['Width']) { [int]$col['Width'] } else { $last }
        if ($col['Align'] -eq 'Right') { (Format-McoText $text $w).Trim().PadLeft($w) } else { Format-McoText $text $w }
    }
    $header = ($Columns | ForEach-Object { & $cell $_ $_['Name'] }) -join '  '
    Write-Host ('{0}{1}{2}{3}{4}' -f $pad, $C.Dim, (' ' * ($script:IconWidth + $script:IconPad.Length)), $header.TrimEnd(), $C.Reset)
    foreach ($row in $Rows) {
        $status = [string]$row.$StatusProperty
        if ($status -notin 'Ok', 'Warn', 'Fail', 'Info', 'Skip') { $status = 'Info' }
        $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red; Info = $C.Blue; Skip = $C.Dim }[$status]
        $cells = foreach ($col in $Columns) { & $cell $col ([string]$row.($col.Property)) }
        $textColor = if ($status -eq 'Skip') { $C.Dim } else { '' }
        Write-Host ('{0}{1}{2}{3}{4}{5}{3}' -f $pad, $color, (Get-McoIcon $status), $C.Reset, $textColor, (($cells -join '  ').TrimEnd()))
    }
}

function Write-McoSummary {
    <# Final summary card (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][System.Collections.Specialized.OrderedDictionary]$Values,
        [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok'
    )

    foreach ($key in $Values.Keys) {
        $v = $Values[$key]
        Write-McoLog 'INFO' ('Summary - {0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
    }
    if ($script:Quiet) { return }
    Clear-McoProgress
    $C = $script:C; $F = $script:Frame; $width = 78
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $script:IconWidth))
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}{0}{5}{6}{7}' -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), ([string]::new($F.Horizontal, $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon, $text = if ($value -is [array]) { (Get-McoIcon $value[0]), $value[1] } else { '   ', $value }
        Write-Host ('    {0}{1}{2,-10}{3} {4}' -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $color, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}

function Get-McoActionText {
    <# The action in words, for the banner, the summary and the report. #>
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Action)
    switch ($Action) {
        'Remove' { 'Remove silently: the copies of the attendees and the rooms, no message' }
        'Cancel' { 'Cancel and clean: the organizer cancels (message to the attendees), then the copies left are removed' }
        'Restore' { 'Restore: the copies removed by a run come back from Recoverable Items, no message' }
        'Transfer' { 'Transfer: the meetings are re-created and sent by a new organizer, the old ones are cancelled' }
        default { 'Report only: nothing is changed' }
    }
}

function Get-McoScopeText {
    param([Parameter(Mandatory = $true)][string]$Scope)
    switch ($Scope) {
        'Organizer' { "organizer's calendar" }
        'Rooms' { 'room mailboxes' }
        'Mailboxes' { 'mailboxes of the list' }
        'AllMailboxes' { 'every mailbox' }
        default { $Scope }
    }
}

function Format-McoOrganizerList {
    <# The organizers in one line: all of them up to 3, else the first ones and the count (and the file). #>
    param([string[]]$Organizer, [string]$File)
    $list = @($Organizer)
    $text = if ($list.Count -le 3) { $list -join ', ' } else { '{0} organizers ({1}, ...)' -f $list.Count, (($list | Select-Object -First 2) -join ', ') }
    if ($File) { $text += " $($script:Dot) file $([IO.Path]::GetFileName($File))" }
    return $text
}

function Get-McoConnectionText {
    <# EWS, sign-in and Exchange PowerShell of the configuration, in words (banner). #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings)
    $dot = $script:Dot
    $ews = if ($Settings.Discovery -eq 'Autodiscover') { "Autodiscover for $($Settings.Mailbox)" } else { [string]$Settings.EwsUrl }
    if ($Settings.EwsServer) { $ews += " (sent to $($Settings.EwsServer))" }
    $ews += " $dot $($Settings.AccessMode)"
    if ($Settings.AccessMode -ne 'Self' -and $Settings.Mailbox) { $ews += " as $($Settings.Mailbox)" }
    $who = if ($Settings.CredentialFile) { "credential file $([IO.Path]::GetFileName([string]$Settings.CredentialFile))" }
        elseif ($Settings.CredentialUser) { "$($Settings.CredentialUser) (prompt)" }
        elseif ($Settings.Authentication -eq 'Basic') { "$($Settings.Mailbox) (prompt)" }
        else { "account running the tool ($([Environment]::UserDomainName)\$([Environment]::UserName))" }
    $sign = if ($Settings.Authentication -eq 'Windows') { "Windows ($($Settings.WindowsPackage)) $dot $who" } else { "Basic $dot $who" }
    $shell = if ([string]$Settings.DirectoryMode -eq 'None') { 'not used (Search.DirectoryMode None): rooms, lists and restore by EWS only' }
        else {
            $uri = if ($Settings.ManagementShellUri) { $Settings.ManagementShellUri } elseif ($Settings.ManagementShellServer) { '{0}://{1}/PowerShell/' -f $(if ($Settings.ManagementShellAuthentication -eq 'Kerberos') { 'http' } else { 'https' }), $Settings.ManagementShellServer } else { '' }
            switch ([string]$Settings.ManagementShellMode) {
                'Existing' { 'Exchange cmdlets of this PowerShell session' }
                'Auto' { "Exchange cmdlets of this session, else remote PowerShell $uri ($($Settings.ManagementShellAuthentication))" }
                default { "remote PowerShell $uri ($($Settings.ManagementShellAuthentication))" }
            }
        }
    [pscustomobject]@{ Ews = $ews; SignIn = $sign; Shell = $shell }
}

function Write-McoRunBanner {
    <# Title card of a command-line run: organizer, meetings, where to search, action, EWS, report and log. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][pscustomobject]$Request,
        [string]$LogPath,
        [switch]$NoReport
    )

    $dot = $script:Dot
    $banner = [ordered]@{}
    if ($Request.FromReport) {
        $banner['Plan'] = @('File', $(if ($Request.Action -eq 'Restore') { "copies removed by the run of $($Request.FromReport)" } else { "meetings reviewed in $($Request.FromReport)" }))
    }
    else {
        if ($Request.Mode -eq 'Rooms') {
            $list = @($Request.Room)
            $text = if ($list.Count -le 3) { $list -join ', ' } else { '{0} rooms ({1}, ...)' -f $list.Count, (($list | Select-Object -First 2) -join ', ') }
            if ($Request.RoomFile) { $text += " $dot file $([IO.Path]::GetFileName($Request.RoomFile))" }
            $banner['Rooms'] = @('Room', "$text $dot every organizer")
        }
        else { $banner['Organizer'] = @('User', (Format-McoOrganizerList -Organizer $Request.Organizer -File $Request.OrganizerFile)) }
        $what = "from $(Format-McoDate $Request.Start $Settings.TimeZone -DateOnly) to $(Format-McoDate $Request.End $Settings.TimeZone -DateOnly -PeriodEnd)"
        if ($Request.Subject) { $what += " $dot subject contains '$($Request.Subject)'" }
        if (@($Request.MeetingId).Count) { $what += " $dot $(@($Request.MeetingId).Count) meeting ID(s)" }
        if ($Request.Mode -ne 'Rooms' -and [string](Get-McoProperty $Request 'SeriesScope') -eq 'Occurrences') { $what += " $dot a series: its occurrences in the period" }
        $banner['Meetings'] = @('Calendar', $what)
        if ($Request.Mode -eq 'Rooms') { $banner['Search in'] = @('Search', 'these rooms only (a series: its occurrences in the period)') }
        else { $banner['Search in'] = @('Search', ((@($Request.SearchIn) | ForEach-Object { Get-McoScopeText $_ }) -join " $dot ")) }
    }
    $banner['Action'] = @($(switch ($Request.Action) { 'Remove' { 'Trash' } 'Cancel' { 'Cancel' } 'Restore' { 'Refresh' } 'Transfer' { 'People' } default { 'Report' } }), (Get-McoActionText $Request.Action))
    if ($Request.Action -eq 'Transfer') { $banner['New organizer'] = @('User', "$($Request.NewOrganizer) $dot re-created by the new organizer (Exchange Server has no organizer transfer)") }
    $connection = Get-McoConnectionText $Settings
    $banner['EWS'] = @('Server', $connection.Ews)
    $banner['Sign-in'] = @('Key', $connection.SignIn)
    $banner['Exchange PS'] = @('Shield', $connection.Shell)
    $banner['Report'] = @('Report', $(if ($NoReport) { 'backup and Summary.json only (-NoReport)' } else { $Settings.OutputPath }))
    if ($LogPath) { $banner['Log'] = @('Log', $LogPath) }
    Write-McoBanner -Title 'Meeting Cleanup On-Prem' -Subtitle "Exchange Server $dot meetings of organizers or rooms, in every calendar" -Details $banner
}

function Write-McoMeetingTable {
    <# The meetings of a result, one line each: start, kind, organizer copy, copies, result, (organizer), subject. #>
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Meetings)

    if (-not $Meetings.Count) { return }
    $fast = [MeetingCleanupOnPremNative.Fast]
    $organizers = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $Meetings) { [void]$organizers.Add([string]$m.Organizer) }
    $rows = foreach ($m in $Meetings) {
        # One line per mailbox: an occurrence copy is counted once for its mailbox.
        $copies = $fast::RealCopies($m)
        $rooms = 0; foreach ($c in $copies) { if ($c.Role -eq 'Room') { $rooms++ } }
        [pscustomobject]@{
            Status    = switch ($m.Status) { { $_ -in 'Removed', 'Cancelled', 'Restored', 'Transferred' } { 'Ok' } { $_ -in 'Partial', 'Not restorable' } { 'Warn' } 'Failed' { 'Fail' } { $_ -in 'Skipped', 'Nothing to do' } { 'Skip' } default { 'Info' } }
            Start     = $m.StartText
            Kind      = [MeetingCleanupOnPremNative.Fast]::KindText($m)
            Subject   = $m.Subject
            Organizer = $m.OrganizerCopy
            Who       = if ($m.OrganizerName) { $m.OrganizerName } else { $m.Organizer }
            Copies    = '{0} ({1} room{2})' -f $copies.Count, $rooms, $(if ($rooms -eq 1) { '' } else { 's' })
            Result    = $m.Status
        }
    }
    $columns = @(
        @{ Name = 'Start'; Property = 'Start'; Width = 16 }
        @{ Name = 'Kind'; Property = 'Kind'; Width = 7 }
        @{ Name = 'Organizer copy'; Property = 'Organizer'; Width = 15 }
        @{ Name = 'Copies'; Property = 'Copies'; Width = 13 }
        @{ Name = 'Result'; Property = 'Result'; Width = 14 }
    )
    if ($organizers.Count -gt 1) { $columns += @{ Name = 'Organizer'; Property = 'Who'; Width = 22 } }
    $columns += @{ Name = 'Subject'; Property = 'Subject'; Width = 0 }
    Write-McoTable -Rows @($rows) -Columns $columns
}

function Write-McoRunSummary {
    <# Final card of a command-line run: status, meetings, copies, report, log and what to do next. #>
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [string]$ReportText = 'none (-NoReport)',
        [string]$LogPath
    )

    $dot = $script:Dot
    $fast = [MeetingCleanupOnPremNative.Fast]
    $n = $Result.Counts
    $values = [ordered]@{}
    $values['Status'] = @($(switch ($Result.Status) { 'Completed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }), "$($Result.Status) $dot $(Get-McoActionText ([string]$Result.Action))")
    if ($n) {
        if ($n.Organizers -gt 1) { $values['Organizers'] = @('People', ('{0} with meetings' -f $n.Organizers)) }
        $values['Meetings'] = @('Calendar', ('{0} found {1} {2} series {1} {3} selected' -f $n.Meetings, $dot, $n.Series, $n.Selected))
        if ($Result.Action -eq 'Restore') {
            $values['Restored'] = @('Refresh', ('{0} restored {1} {2} already present {1} {3} not found {1} {4} not restorable {1} {5} failed' -f $n.Restored, $dot, $n.AlreadyPresent, $n.NotFound, $n.NotRestorable, $n.Failed))
        }
        else {
            $values['Copies'] = @('People', ('{0} in {1} mailbox(es) {2} {3} organizer {2} {4} attendee {2} {5} room {2} {6} not processed' -f $n.Copies, $n.Mailboxes, $dot, $n.OrganizerCopies, $n.AttendeeCopies, $n.RoomCopies, $n.NotProcessed))
            if ($Result.Action -eq 'Transfer') {
                $skipped = 0; $failed = 0
                foreach ($m in $Result.Meetings) { if ($m.Selected -and $m.Status -eq 'Skipped') { $skipped++ } elseif ($m.Status -eq 'Failed') { $failed++ } }
                $values['Transferred'] = @('People', ('{0} re-created by {1} {2} {3} not transferred {2} {4} failed' -f $n.Transferred, $fast::Text($Result, 'NewOrganizer'), $dot, $skipped, $failed))
                $values['Done'] = @('Trash', ('old copies: {0} removed {1} {2} cancelled by their old organizer {1} {3} failed' -f $n.Removed, $dot, $n.Cancelled, $n.Failed))
            }
            elseif ($Result.Action -ne 'Report') {
                $values['Done'] = @('Trash', ('{0} removed {1} {2} cancelled {1} {3} already gone {1} {4} kept {1} {5} failed' -f $n.Removed, $dot, $n.Cancelled, $n.AlreadyGone, $n.Kept, $n.Failed))
            }
            if ($n.OccurrenceCopies) {
                $series = 0; foreach ($m in $Result.Meetings) { if ($fast::Text($m, 'Scope') -eq 'Occurrences') { $series++ } }
                $values['Occurrences'] = @('Calendar', ('{0} series limited to the period: {1} occurrence copies (not restorable once removed)' -f $series, $n.OccurrenceCopies))
            }
        }
    }
    $warnings = @($Result.Warnings)
    if ($warnings.Count) { $values['Warnings'] = @('Warn', $(if ($warnings.Count -eq 1) { [string]$warnings[0] } else { '{0} (first: {1})' -f $warnings.Count, $warnings[0] })) }
    if ($fast::Text($Result, 'BackupFile')) { $values['Backup'] = @('Shield', $Result.BackupFile) }
    if ($fast::Text($Result, 'Error')) { $values['First issue'] = @('Fail', $Result.Error) }
    $values['Duration'] = @('Clock', (Format-McoDuration ([double]$Result.DurationSeconds)))
    $values['Report'] = @('Report', $ReportText)
    if ($LogPath) { $values['Log'] = @('Log', $LogPath) }
    # The folder of this run, for the command to give next (-FromReport).
    $folder = if ($fast::Text($Result, 'BackupFile')) { Split-Path $Result.BackupFile -Parent } elseif ($ReportText -match '[\\/]') { Split-Path $ReportText -Parent } else { '<report folder>' }
    $values['Next'] = @('Info', $(switch ($Result.Status) {
                'Completed' {
                    if ($Result.Action -eq 'Report') {
                        if ($n -and $n.Meetings) { "Review the report, then: -FromReport '$folder' -Action Remove (or Cancel, or Transfer -NewOrganizer <address>), with -MeetingId <id> to act on some of them only." }
                        else { 'Nothing found: widen the period or search more mailboxes (-SearchIn Rooms, Mailboxes, AllMailboxes).' }
                    }
                    elseif ($Result.Action -eq 'Restore') { 'Nothing to do. The copies are back in their calendars; no message was sent.' }
                    elseif ($Result.Action -eq 'Transfer') { "Nothing to do. The attendees answer the invitation of $($fast::Text($Result, 'NewOrganizer')) again (re-created meetings)." }
                    elseif ($Result.Action -eq 'Cancel') { 'Nothing to do. A cancellation cannot be undone: the attendees received it.' }
                    else { "Nothing to do. To undo it, within the retention of deleted items (14 days by default): -Action Restore -FromReport '$folder'." }
                }
                'Failed' { 'Read the first issue above and the log; nothing more was changed after it.' }
                default {
                    if ($Result.Action -eq 'Restore' -and $n -and $n.NotRestorable -and -not ($n.Failed + $n.NotFound)) { 'Nothing more to do: a cancelled or transferred meeting, or an occurrence, cannot be restored.' }
                    else { 'Open the report: each copy gives its result. Running the same command again retries what is left.' }
                }
            }))
    $title = switch ($Result.Status) { 'Completed' { switch ($Result.Action) { 'Report' { 'Search finished' } 'Restore' { 'Restore finished' } 'Transfer' { 'Transfer finished' } default { 'Cleanup finished' } } } 'Failed' { 'Run failed' } default { 'Finished with warnings' } }
    $card = switch ($Result.Status) { 'Completed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }
    Write-McoSummary -Title $title -Values $values -Status $card
}

# SIG # Begin signature block
# MIImIwYJKoZIhvcNAQcCoIImFDCCJhACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBRjveOgBQioAmC
# uZYtc5VwixeO/MXToh2T0GgPrcQwXqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCADRPZLZEeq
# jMUQjnlfrBDPtnxRWMI9+TBeFRdaUvxDTDANBgkqhkiG9w0BAQEFAASCAgBYSNSc
# djX8w1h167tQonDzjm10GfjyA/7tSKbtLE5MTBSp1NFn4QNoFqQ9Jnr3/l8nwfQM
# fvznbsr5CL2LUIhTziJ5OJwTz+Eqtep1VpY7nukzVUMzWcINRUnZzBzgaWb3GDYK
# wPGLpZDc8hqE5Znp1rTPHrplg8927dTe7cqmh5z7CylZQiRxEKvkhBDOKtiPNa+c
# RVsSP2W2DoLVORFXtYBu7PyPXiF5JXeD28RE3DUmDFsyujYXvFxj3dGjVhsf/A9l
# +9vq31I9PGG2njgTGpM291WqP9tIB2vkFpiqn/bA0OZPC4cs+hhgNJq5FsINJssq
# v7K7dewZVwioE1LdY+lzkXSdG/NMpCMO3eMVRv0D/kBAS9Ix7hZZt3eV1PP0rbzk
# Sxzx1nmi5oQtdcctjOg92Fr0BVKEm17nVWojlYag7U40gzERrWS+b5t690ezxOJ/
# b3Uz6FrxvDau3J+4h+t8AR7X9qj9uFMKa2+bzS9uT7DNxmOMbt23Wnyj5yEth0aU
# 51Ib9UAJ9PfhuxafY5z6mFcUCaT+S4cV8ff8+SJ2M/II44tmM2bbvGxLu9YnQpr9
# fbmY5O2/bYCAiJkiNC1BLgzWREr4Pz9ECj+Jj4uHwvWZsg28nxSvFm8aY+GnifbO
# TzuPeB4CO1cUNiuJhDBspeiSSXNFw0UWxbrqCKGCFMcwghTDBgorBgEEAYI3AwMB
# MYIUszCCFK8GCSqGSIb3DQEHAqCCFKAwghScAgEDMQ0wCwYJYIZIAWUDBAIBMIGC
# BgsqhkiG9w0BCRABBKBzBHEwbwIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIJgRzheh0MNIxI+nChsvNSl2MPtrnhAd9bYZh8Fxt/34AghzF/6lCU7G
# ihgPMjAyNjExMDUxNzIyMzdaMAMCAQECCQDZLJ9KTWChXaCCEacwggTzMIIC26AD
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
# BgkqhkiG9w0BCQUxDxcNMjYxMTA1MTcyMjM3WjAoBgkqhkiG9w0BCTQxGzAZMAsG
# CWCGSAFlAwQCAaEKBggqhkjOPQQDAjAvBgkqhkiG9w0BCQQxIgQgWdZw1OEpep6j
# 8BFfroPTlXyKW2wjdyetzgN6Z5ANhwgwgcgGCyqGSIb3DQEJEAIvMYG4MIG1MIGy
# MIGvBCBXT+oXqvpefGormwjJmLFzXkULOHPKROaiTKCTDmPzsDCBijB2pHQwcjEL
# MAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREw
# DwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcg
# SXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO4zAKBggqhkjOPQQDAgRH
# MEUCIQDnauhRXJ+61NvVL4J6j6cLwRcNKCthruI4Cz92p5zOJAIgfk5y447RffvB
# +hhcHgudlly7Qo1WFn/S1Hn05Hkk6PM=
# SIG # End signature block

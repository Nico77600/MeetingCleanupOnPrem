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
    Version : 1.1.0
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

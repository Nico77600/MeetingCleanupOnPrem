#Requires -Version 7.4

<#
.SYNOPSIS
    Renders the graphics of the GitHub README from the Meeting Cleanup On-Prem developer guide, in a light
    and a dark version: banner, principles, how it works, the transfer of meetings to a new organizer.

.DESCRIPTION
    GitHub renders Markdown only: the custom blocks of the guide (cards, flow) and its theme are lost. This
    tool renders them as images with the CSS of the built HTML guide and the icons of
    tools\Build-Documentation.ps1, so that the README and the guide always look the same. The README shows
    them with <picture>, which picks the light or dark image from the theme of the reader.

    Sources:
      docs\MeetingCleanupOnPrem-Guide.md     the cards block of the introduction, the version
      docs\MeetingCleanupOnPrem-Guide.html   the CSS (run tools\Build-Documentation.ps1 first)
      tools\Build-Documentation.ps1          the icons

    Screenshots: Microsoft Edge in headless mode, with a temporary profile, 2x resolution. Only local files
    are opened. Output: docs\images\readme-<name>-light.png and readme-<name>-dark.png.

.PARAMETER OutputFolder
    Default: docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (the HTML pages of the graphics) and shows its path.

.EXAMPLE
    .\tools\Build-Documentation.ps1; .\tools\New-ReadmeImages.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0  (from Meeting Cleanup 1.3.0)
    Part of : Meeting Cleanup On-Prem (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'docs\images' }

#region Assets of the guide ------------------------------------------------------------------------
function ConvertTo-ReadmeInline([string]$Text) {
    # Inline Markdown of a guide block (code, bold, italic) -> HTML.
    $h = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $h = [regex]::Replace($h, '`([^`]+)`', '<code>$1</code>')
    $h = [regex]::Replace($h, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    return [regex]::Replace($h, '(?<![\w*])\*([^*\s][^*]*)\*(?![\w*])', '<em>$1</em>')
}

function Get-ReadmeAssets {
    param([string]$Root)
    $builder = Join-Path $Root 'tools\Build-Documentation.ps1'
    $guideHtml = Join-Path $Root 'docs\MeetingCleanupOnPrem-Guide.html'
    $guideMd = Join-Path $Root 'docs\MeetingCleanupOnPrem-Guide.md'
    if (-not (Test-Path $guideHtml)) { throw 'docs\MeetingCleanupOnPrem-Guide.html not found: run tools\Build-Documentation.ps1 first (it holds the CSS of the graphics).' }
    # Icons: the $Icons table of the documentation builder, read without running the builder.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($builder, [ref]$null, [ref]$null)
    $assign = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$Icons' }, $true)
    if (-not $assign) { throw "Icon table not found in $builder." }
    $md = [IO.File]::ReadAllText($guideMd) -replace "`r`n", "`n"
    $blocks = foreach ($m in [regex]::Matches($md, '(?s)```(flow|cards)\n(.*?)\n```')) {
        [pscustomobject]@{ Kind = $m.Groups[1].Value; Lines = @($m.Groups[2].Value -split "`n" | Where-Object { $_.Trim() }) }
    }
    [pscustomobject]@{
        Icons   = & ([scriptblock]::Create($assign.Right.Extent.Text))
        Css     = [regex]::Match([IO.File]::ReadAllText($guideHtml), '(?s)<style>(.*?)</style>').Groups[1].Value
        Version = [regex]::Match($md, '(?m)^version:\s*(\S+)').Groups[1].Value
        Flows   = @($blocks | Where-Object Kind -eq 'flow')
        Cards   = @($blocks | Where-Object Kind -eq 'cards')
    }
}

function Get-ReadmeIcon([string]$Name, [string]$Class = 'icon') {
    $path = $assets.Icons[$Name]; if (-not $path) { $path = $assets.Icons['info'] }
    "<svg class=""$Class"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"">$path</svg>"
}

function ConvertTo-ReadmeFlow([string[]]$Lines, [switch]$Vertical) {
    # Vertical: the nodes are stacked, icon on the left, with a downward arrow and its label.
    $items = foreach ($l in $Lines) {
        $icon, $title, $sub = $l.Split('|', 3).ForEach({ $_.Trim() })
        $title = [System.Net.WebUtility]::HtmlEncode($title); $sub = [System.Net.WebUtility]::HtmlEncode($sub)
        if ($Vertical) {
            if ($icon -eq 'arrow') {
                $note = if ($sub) { "<span class=""flow-sub"">$sub</span>" } else { '' }
                "<div class=""rb-varrow""><svg viewBox=""0 0 12 30""><path d=""M6 1v26M1 21l5 6 5-6"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-label"">$title</span>$note</div>"
            } else {
                "<div class=""rb-vnode""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div></div>"
            }
        } elseif ($icon -eq 'arrow') {
            $class = if ($title -or $sub) { 'flow-arrow' } else { 'flow-arrow rb-bare' }
            "<div class=""$class""><span class=""flow-label"">$title</span><svg viewBox=""0 0 40 12""><path d=""M0 6h36M31 1l6 5-6 5"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-sub"">$sub</span></div>"
        } else {
            "<div class=""flow-node""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div>"
        }
    }
    $class = if ($Vertical) { 'flow rb-vflow' } else { 'flow rb-flow' }
    "<div class=""$class"">$($items -join '')</div>"
}

function ConvertTo-ReadmeCards([string[]]$Lines, [string]$Class = '') {
    $items = foreach ($l in $Lines) {
        $icon, $title, $text = $l.Split('|', 3).ForEach({ $_.Trim() })
        "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""card-title"">$(ConvertTo-ReadmeInline $title)</div><div class=""card-text"">$(ConvertTo-ReadmeInline $text)</div></div></div>"
    }
    "<div class=""cards $Class"">$($items -join '')</div>"
}

function Get-ReadmePill([string]$Text, [string]$Tone) { "<span class=""rb-pill"" style=""--tone: var(--cp-$Tone)"">$Text</span>" }
#endregion

#region Styles of the graphics, on top of the CSS of the guide -------------------------------------
$Script:ReadmeCss = @'
html, body { background: #ffffff; }
html[data-theme="dark"], html[data-theme="dark"] body { background: #0d1117; }
:root { --cp-info: #0078d4; --cp-violet: #7c3aed; --cp-teal: #0d9488; }
html[data-theme="dark"] { --cp-info: #4da6ff; --cp-violet: #a78bfa; --cp-teal: #2dd4bf; }
body { display: block; margin: 0; padding: 0; }
.canvas { padding: 6px; }
.rb-pill { display: inline-block; padding: 1px 10px; margin: 8px 6px 0 0; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.6;
  color: var(--tone); background: color-mix(in srgb, var(--tone) 11%, transparent); border: 1px solid color-mix(in srgb, var(--tone) 38%, transparent); }
.rb-caption { font-size: 11.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); margin: 0 0 8px 4px; }
.rb-caption span { color: var(--cp-text-muted); font-weight: 600; letter-spacing: 0.04em; text-transform: none; font-size: 12.5px; }
/* Before / after */
.rb-bench { display: grid; gap: 14px; padding: 20px 22px; border-radius: 16px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-row { display: grid; grid-template-columns: 250px minmax(0, 1fr); gap: 18px; align-items: center; }
.rb-row .label { font-size: 13.5px; font-weight: 650; color: var(--cp-text); } .rb-row .label span { display: block; font-weight: 500; font-size: 12px; color: var(--cp-text-muted); margin-top: 2px; }
.rb-bars { display: grid; gap: 6px; }
.rb-bar { display: flex; align-items: center; gap: 10px; font-size: 12.5px; color: var(--cp-text-muted); }
.rb-bar i { display: block; height: 18px; border-radius: 6px; min-width: 6px; }
.rb-bar.old i { background: color-mix(in srgb, var(--cp-text-muted) 45%, transparent); } .rb-bar.new i { background: var(--cp-accent); }
.rb-bar b { color: var(--cp-text); font-weight: 650; } .rb-bar .bad { color: var(--cp-danger); font-weight: 600; }
.rb-legend { display: flex; gap: 18px; font-size: 12px; color: var(--cp-text-muted); margin: 0 0 4px 4px; } .rb-legend i { display: inline-block; width: 12px; height: 12px; border-radius: 4px; margin-right: 6px; vertical-align: -1px; }
/* Banner */
.rb-hero { margin: 0; padding: 32px 36px 30px; }
.rb-hero-grid { position: relative; display: grid; grid-template-columns: minmax(0, 1fr) 240px; gap: 34px; align-items: center; }
.rb-hero h1 { font-size: 35px; }
.rb-hero .lead { margin: 18px 0 0; font-size: 17px; max-width: none; }
.rb-hero .badges { margin: 20px 0 0; }
.rb-stats { position: relative; display: grid; gap: 10px; }
.rb-stat { display: flex; align-items: center; gap: 14px; padding: 12px 16px; border-radius: 14px; background: var(--cp-panel-strong); border: 1px solid var(--cp-border); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08); }
.rb-stat b { font-size: 30px; line-height: 1; color: var(--cp-accent); font-weight: 750; min-width: 40px; text-align: center; }
.rb-stat span { font-size: 13px; color: var(--cp-text-muted); line-height: 1.35; }
.rb-stat strong { display: block; color: var(--cp-text); font-size: 14px; }
/* Cards and flows */
.cards { margin: 0; }
.rb-cards2 { grid-template-columns: 1fr 1fr; }
.rb-flow { margin: 0; flex-wrap: nowrap; padding: 18px; gap: 4px; }
.rb-flow .flow-node { flex: 1 1 0; min-width: 0; padding: 14px 10px; }
.rb-flow .flow-title { font-size: 13.5px; overflow-wrap: anywhere; }
.rb-flow .flow-arrow { min-width: 0; width: 84px; flex: 0 0 84px; }
.rb-flow .flow-arrow.rb-bare { width: 46px; flex-basis: 46px; }
.rb-flow .flow-sub { max-width: 84px; }
.rb-space { height: 18px; }
/* How it works: vertical pipeline and the three modes */
.rb-hiw { display: grid; grid-template-columns: minmax(0, 1.08fr) minmax(0, 1fr); gap: 16px; align-items: stretch; }
.rb-col { display: flex; flex-direction: column; }
.rb-vflow { flex: 1; flex-direction: column; flex-wrap: nowrap; align-items: stretch; justify-content: center; gap: 0; margin: 0; padding: 16px 18px; }
.rb-vnode { display: flex; align-items: center; gap: 14px; padding: 11px 16px; border-radius: 12px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-vnode .flow-icon { margin: 0; flex-shrink: 0; }
.rb-vnode .flow-text { margin-top: 1px; }
.rb-varrow { display: flex; align-items: center; gap: 10px; min-height: 36px; padding-left: 31px; }
.rb-varrow svg { width: 12px; height: 28px; color: var(--cp-accent); flex-shrink: 0; }
.rb-varrow .flow-sub { max-width: none; font-size: 12px; }
.rb-modes { flex: 1; display: flex; flex-direction: column; gap: 10px; }
.rb-modes .card-item { flex: 1; align-items: center; }
.rb-modes .card-title { display: flex; align-items: center; gap: 8px; }
.rb-chip { font-size: 11px; font-weight: 600; padding: 0 8px; border-radius: 999px; border: 1px solid var(--cp-border); color: var(--cp-text-muted); }
.rb-chip.hot { color: var(--cp-accent-fg); background: var(--cp-accent); border-color: var(--cp-accent); }
'@
#endregion

#region Rendering (Microsoft Edge, headless) -------------------------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height, [int]$Scale = 1) {
    $url = 'file:///' + ($Html -replace '\\', '/')
    $profilePath = Join-Path $work 'edge-profile'
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profilePath`"", "--window-size=$Width,$Height", "--force-device-scale-factor=$Scale", "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # Height of the .canvas element: the page writes it in body[data-h], read with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = [IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

function New-ReadmeGraphic {
    # One graphic, light and dark: HTML page -> height measured by Edge -> 2x screenshot.
    param([string]$Name, [string]$Body, [int]$Width)
    $pages = @{}
    foreach ($theme in 'light', 'dark') {
        $html = "<!doctype html><html lang=""en"" data-theme=""$theme""><head><meta charset=""utf-8""><style>$($assets.Css)`n$($Script:ReadmeCss)</style></head>" +
            "<body><div class=""canvas"" style=""width:$($Width)px"">$Body</div><script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.canvas').getBoundingClientRect().height));</script></body></html>"
        $pages[$theme] = Join-Path $work "readme-$Name-$theme.html"
        [IO.File]::WriteAllText($pages[$theme], $html, [Text.UTF8Encoding]::new($false))
    }
    $height = Get-PageHeight $pages['light'] $Width
    foreach ($theme in 'light', 'dark') { Save-Screenshot $pages[$theme] (Join-Path $OutputFolder "readme-$Name-$theme.png") $Width $height 2 }
}
#endregion

#region Main ---------------------------------------------------------------------------------------
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('mco-readme-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$Script:assets = Get-ReadmeAssets -Root $root
if ($assets.Cards.Count -lt 1) { throw 'The guide must hold the cards block of the introduction.' }
$mid = '&middot;'

try {
    Write-Host 'Rendering the README graphics (light and dark, 2x)...'

    # Banner: the hero of the guide, with the key figures.
    $badges = @(
        "<span class=""badge badge-accent"">Version $($assets.Version)</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'terminal' 'icon-sm')PowerShell 7.4+</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'server' 'icon-sm')Exchange Server 2016 &middot; 2019 &middot; SE</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'key' 'icon-sm')EWS &middot; Exchange PowerShell</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'tag' 'icon-sm')MIT license</span>"
    ) -join ''
    $banner = "<header class=""hero rb-hero""><div class=""rb-hero-grid""><div>" +
        "<div class=""hero-top""><div class=""hero-logo"">$(Get-ReadmeIcon 'calendar')</div><div><div class=""eyebrow"">Exchange Server $mid Calendars $mid Organizers $mid Rooms</div><h1>Meeting Cleanup On-Prem</h1></div></div>" +
        "<p class=""lead"">Finds the meetings of <strong>organizers who left or stay</strong>, or <strong>every meeting of some rooms</strong>, in <strong>every calendar</strong> of Exchange Server where they are &mdash; then removes them silently, has the organizer cancel them, or <strong>transfers them to a new organizer</strong>, even when the old mailbox is gone.</p>" +
        "<div class=""badges"">$badges</div></div>" +
        "<div class=""rb-stats"">" +
        "<div class=""rb-stat""><b>1</b><span><strong>UID</strong>finds every copy of a meeting</span></div>" +
        "<div class=""rb-stat""><b>0</b><span><strong>message</strong>for a silent removal, restorable</span></div>" +
        "<div class=""rb-stat""><b>4</b><span><strong>actions</strong>remove, cancel, transfer, restore</span></div>" +
        "</div></div></header>"
    New-ReadmeGraphic -Name 'banner' -Body $banner -Width 1080

    # Principles: the cards block of the introduction of the guide.
    New-ReadmeGraphic -Name 'principles' -Body (ConvertTo-ReadmeCards $assets.Cards[0].Lines 'rb-cards2') -Width 1080

    # How it works: the search as a vertical pipeline, and the four actions.
    $stages = @(
        'user | Organizers or rooms | addresses, aliases, X500 of a deleted mailbox'
        'arrow | where | organizer, every room, a list, every mailbox'
        'search | Search | CalendarView of the period, page after page'
        'arrow | UID | the same in every copy'
        'people | Every copy | attendees, rooms, members of the groups invited'
        'arrow | report first | nothing changed until you confirm'
        'check | Action | backup first, then every copy read again'
    )
    $uses = @(
        [pscustomobject]@{ Icon = 'trash'; Name = 'Remove'; Chip = '<span class="rb-chip">silent</span>'; Text = 'The copies of the attendees and the rooms go, <strong>without any message</strong>; the organizer keeps his meeting.'; Pills = (Get-ReadmePill 'Rooms freed' 'success') + (Get-ReadmePill 'Restorable' 'info') }
        [pscustomobject]@{ Icon = 'mail'; Name = 'Cancel'; Chip = '<span class="rb-chip hot">organizer</span>'; Text = 'The organizer <strong>cancels with your message</strong>, then the copies left are removed. A room closed: only its occurrences in the period.'; Pills = (Get-ReadmePill 'Your message' 'warning') + (Get-ReadmePill 'Rooms mode' 'teal') }
        [pscustomobject]@{ Icon = 'people'; Name = 'Transfer'; Chip = '<span class="rb-chip">new organizer</span>'; Text = '<strong>Re-created by the new organizer</strong> with one invitation; the old meeting cancelled by its organizer, or removed when he is gone.'; Pills = (Get-ReadmePill 'One invitation' 'violet') + (Get-ReadmePill 'Series from now' 'info') }
        [pscustomobject]@{ Icon = 'refresh'; Name = 'Restore'; Chip = '<span class="rb-chip">undo</span>'; Text = 'The copies of a Remove come back from <strong>Recoverable Items</strong>, as they were, without a message.'; Pills = (Get-ReadmePill '14 days by default' 'teal') + (Get-ReadmePill 'Never deletes' 'success') }
    )
    $useHtml = ($uses | ForEach-Object { "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $_.Icon)</div><div><div class=""card-title"">$($_.Name) $($_.Chip)</div><div class=""card-text"">$($_.Text)</div><div>$($_.Pills)</div></div></div>" }) -join ''
    $howItWorks = "<div class=""rb-hiw""><div class=""rb-col""><div class=""rb-caption"">The search <span>$mid the same for every case</span></div>$(ConvertTo-ReadmeFlow $stages -Vertical)</div>" +
        "<div class=""rb-col""><div class=""rb-caption"">Four actions <span>$mid the report is the default</span></div><div class=""rb-modes"">$useHtml</div></div></div>"
    New-ReadmeGraphic -Name 'how-it-works' -Body $howItWorks -Width 1080

    # Transfer: the re-creation by the new organizer, step by step.
    $transfer = @(
        'search | Read | the meeting, its occurrences to come'
        'arrow | | '
        'building | Rooms | old copies removed'
        'arrow | | '
        'mail | Create and send | one invitation'
        'arrow | | '
        'calendar | Old meeting | cancelled by its organizer'
        'arrow | | '
        'trash | Old copies | removed silently'
        'arrow | | '
        'check | Verify | old gone, new there'
    )
    $transferHtml = "<div class=""rb-caption"">Transfer to a new organizer <span>$mid re-created by him, measured on a lab Exchange Server 2019</span></div>$(ConvertTo-ReadmeFlow $transfer)"
    New-ReadmeGraphic -Name 'transfer' -Body $transferHtml -Width 1080

    Get-ChildItem $OutputFolder -Filter 'readme-*.png' | Select-Object Name, @{ n = 'KB'; e = { [math]::Round($_.Length / 1KB) } } | Format-Table -AutoSize | Out-String | Write-Host
} finally {
    # Edge helper processes of the temporary profile, if any are left.
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
}
#endregion
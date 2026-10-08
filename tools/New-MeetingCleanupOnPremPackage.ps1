<#
.SYNOPSIS
    Copies the files needed to run Meeting Cleanup On-Prem into a separate folder, and zips it: the release package.

.DESCRIPTION
    The package contains only what Invoke-MeetingCleanupOnPrem.ps1 needs at run time, plus the HTML guides:
        Invoke-MeetingCleanupOnPrem.ps1, MeetingCleanupOnPrem.psd1, MeetingCleanupOnPrem.psm1, src\, config\,
        templates\, docs\MeetingCleanupOnPrem-UserGuide.html, docs\MeetingCleanupOnPrem-Guide.html, README.md,
        CHANGELOG.md, LICENSE, THIRD-PARTY-NOTICES.md
    The HTML guides are rebuilt first from their Markdown source (tools\Build-Documentation.ps1): they are
    self-contained (images inline), so the Markdown sources and the images are not copied.
    It never copies reports\, logs\, artifacts\, lab\, tests\ (with the simulated Exchange) or tools\.

    The configuration is copied as delivered (sample contoso values, no credential). The script checks the content
    of the package and that the module loads from it, then writes <Destination>.zip.

.PARAMETER Destination
    Package folder. Default: package\MeetingCleanupOnPrem-<version>, next to the repository folder.

.PARAMETER Force
    Replace the destination folder (and its zip) if it already contains a package. A folder that contains reports\
    or logs\ (a package that has been run) is never replaced.

.EXAMPLE
    .\tools\New-MeetingCleanupOnPremPackage.ps1
    Creates ..\package\MeetingCleanupOnPrem-<version> and ..\package\MeetingCleanupOnPrem-<version>.zip.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0  (from Meeting Cleanup 1.3.0)
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$root = Join-Path $repoRoot 'package'
$version = (Import-PowerShellDataFile -LiteralPath (Join-Path $root 'MeetingCleanupOnPrem.psd1')).ModuleVersion
if (-not $Destination) { $Destination = Join-Path (Split-Path $repoRoot -Parent) "package\MeetingCleanupOnPrem-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')
$zip = "$Destination.zip"
$rootPrefix = [IO.Path]::GetFullPath($repoRoot).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the repository folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-MeetingCleanupOnPrem.ps1'))) { throw "The destination is not a Meeting Cleanup On-Prem package, it is not replaced: $Destination" }
    foreach ($used in 'reports', 'logs') {
        if (Test-Path -LiteralPath (Join-Path $Destination $used)) { throw "The destination contains a $used folder (a package that has been run), it is not replaced: $Destination" }
    }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}
if (Test-Path -LiteralPath $zip) {
    if (-not $Force) { throw "The zip already exists: $zip. Use -Force to replace it." }
    Remove-Item -LiteralPath $zip -Force
}

# ---- HTML guides, rebuilt from the Markdown sources ------------------------------------------------
& (Join-Path $PSScriptRoot 'Build-Documentation.ps1') | Out-Null

# ---- Files needed at run time -------------------------------------------------------------------
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-MeetingCleanupOnPrem.ps1', 'MeetingCleanupOnPrem.psd1', 'MeetingCleanupOnPrem.psm1', 'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md',
    'config\MeetingCleanupOnPrem.config.psd1', 'templates\Report.template.html', 'docs\MeetingCleanupOnPrem-UserGuide.html', 'docs\MeetingCleanupOnPrem-Guide.html') { $files.Add($f) }
Get-ChildItem -LiteralPath (Join-Path $root 'src') -File | Where-Object Extension -in '.ps1', '.cs' | ForEach-Object { $files.Add("src\$($_.Name)") }
foreach ($f in $files) {
    $sourceRoot = if ($f -eq 'CHANGELOG.md') { $repoRoot } else { $root }
    $source = Join-Path $sourceRoot $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Checks ---------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'reports', 'logs', 'tests', 'artifacts', 'tools', 'lab', 'docs\images') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File -Include '*.log', '*.csv', '*.json', '*.png', '*.xml', '*.Tests.ps1', '*FakeEws*' |
    ForEach-Object { $problems.Add("Not a run-time file: $($_.Name)") }
foreach ($part in 'Console.ps1', 'Config.ps1', 'Ews.ps1', 'Rps.ps1', 'Exchange.ps1', 'Search.ps1', 'Cleanup.ps1', 'Restore.ps1', 'Transfer.ps1', 'Report.ps1', 'Native.cs') {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination "src\MeetingCleanupOnPrem.$part"))) { $problems.Add("Missing in the package: src\MeetingCleanupOnPrem.$part") }
}
foreach ($guide in 'MeetingCleanupOnPrem-UserGuide.html', 'MeetingCleanupOnPrem-Guide.html') {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination "docs\$guide"))) { $problems.Add("Missing in the package: docs\$guide") }
}
$config = Import-PowerShellDataFile -LiteralPath (Join-Path $Destination 'config\MeetingCleanupOnPrem.config.psd1')
if ($config.Connection.CredentialFile -or $config.Connection.CredentialUser -or $config.ManagementShell.CredentialUser -or $config.Connection.EwsServer) { $problems.Add('The configuration of the package must not hold a credential, an account or a server of a real organization.') }
if ($config.Connection.EwsUrl -notmatch 'contoso\.') { $problems.Add('The configuration of the package must hold the sample EWS URL (contoso).') }
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

# The module loads from the package, and its reports go under the package folder.
$loaded = & pwsh -NoProfile -Command "Import-Module '$Destination\MeetingCleanupOnPrem.psd1'; (Get-Command -Module MeetingCleanupOnPrem).Count; (Import-McoConfiguration).OutputPath"
if ($LASTEXITCODE -ne 0 -or [int]$loaded[0] -lt 20 -or -not ([string]$loaded[1]).StartsWith($Destination)) { throw "The module does not load correctly from the package: $loaded" }

# The zip: the folder MeetingCleanupOnPrem-<version> inside, as in the releases.
Compress-Archive -Path $Destination -DestinationPath $zip -CompressionLevel Optimal

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Meeting Cleanup On-Prem $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host ("  Zip      : {0} ({1:N1} MB)" -f $zip, ((Get-Item -LiteralPath $zip).Length / 1MB))
Write-Host "  Check    : module loads ($($loaded[0]) commands), reports written under the package folder"
Write-Host "  Config   : sample contoso values - fill them in before the first run (user guide, chapter 1)"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,10:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }

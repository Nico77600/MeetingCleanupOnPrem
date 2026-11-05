<#
.SYNOPSIS
    Builds the release package of Meeting Cleanup On-Prem in three phases, so that the owner signs between them.

.DESCRIPTION
    The package contains only what Invoke-MeetingCleanupOnPrem.ps1 needs at run time, plus the HTML guides:
        Invoke-MeetingCleanupOnPrem.ps1, MeetingCleanupOnPrem.psd1, MeetingCleanupOnPrem.psm1, src\, config\,
        templates\, docs\MeetingCleanupOnPrem-UserGuide.html, docs\MeetingCleanupOnPrem-Guide.html, README.md,
        CHANGELOG.md, LICENSE, THIRD-PARTY-NOTICES.md, PackageIntegrity.psd1
    The HTML guides are the committed files of package\docs\ (built from their Markdown source with
    tools\Build-Documentation.ps1, which is run by hand or with -RebuildDocs): they are self-contained (images inline),
    so the Markdown sources and the images are not copied.
    It never copies reports\, logs\, artifacts\, lab\, tests\ (with the simulated Exchange) or tools\.

    Build (default): copies the files into a new folder and refuses to continue unless the PowerShell files of
    package\ carry the timestamped owner signature. No manifest is written and no ZIP.
    -Unsigned (continuous integration): the same build without the signature requirement. The committed
    package\PackageIntegrity.psd1 is copied when it exists and the folder is verified with -DevelopmentMode;
    an '-unsigned-ci' ZIP is written. It is never a release.
    -Manifest <package folder>: hashes the built folder and writes an unsigned PackageIntegrity.psd1. The owner signs it.
    -Seal <package folder>: verifies the signed manifest on the built folder and on package\ (the bytes of the
    protected files must be the ones in the repository), then writes the release ZIP, with the package folder as its
    single root folder, and prints its SHA-256.
    Signing key material is never read by this script.

    The configuration is copied as delivered (sample contoso values, no credential).

.PARAMETER Destination
    Package folder. Default: package\MeetingCleanupOnPrem-<version>, next to the repository folder.

.PARAMETER Force
    Replace the destination folder (and its zip) if it already contains a package. A folder that contains reports\
    or logs\ (a package that has been run) is never replaced.

.PARAMETER RebuildDocs
    Build only: runs tools\Build-Documentation.ps1 first. Refused with -Unsigned, because the committed
    PackageIntegrity.psd1 lists the committed HTML guides. After a rebuild, the manifest must be regenerated.

.PARAMETER Unsigned
    Continuous integration: builds without the owner signature requirement.

.PARAMETER Manifest
    Built package folder: writes its unsigned PackageIntegrity.psd1.

.PARAMETER Seal
    Built package folder: verifies it against the signed manifest and writes the release ZIP.

.EXAMPLE
    .\tools\New-MeetingCleanupOnPremPackage.ps1 -Destination C:\work\out\MeetingCleanupOnPrem-1.1.1
    Builds the package folder from signed sources. Then -Manifest, sign the manifest, copy it to package\, -Seal.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.1  (from Meeting Cleanup 1.3.0)
#>
#Requires -Version 7.4
[CmdletBinding(DefaultParameterSetName = 'Build')]
param(
    [Parameter(ParameterSetName = 'Build')][string]$Destination,
    [Parameter(ParameterSetName = 'Build')][switch]$Force,
    [Parameter(ParameterSetName = 'Build')][switch]$Unsigned,
    [Parameter(ParameterSetName = 'Build')][switch]$RebuildDocs,
    [Parameter(ParameterSetName = 'Manifest', Mandatory)][string]$Manifest,
    [Parameter(ParameterSetName = 'Seal', Mandatory)][string]$Seal
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$root = Join-Path $repoRoot 'package'
$version = (Import-PowerShellDataFile -LiteralPath (Join-Path $root 'MeetingCleanupOnPrem.psd1')).ModuleVersion
$signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
# Files of the package that carry the owner signature, relative to the package folder.
$signedScripts = @('Invoke-MeetingCleanupOnPrem.ps1', 'MeetingCleanupOnPrem.psm1', 'MeetingCleanupOnPrem.psd1', 'src\Test-PackageIntegrity.ps1') +
    @('Cleanup', 'Config', 'Console', 'Ews', 'Exchange', 'Report', 'Restore', 'Rps', 'Search', 'Transfer' | ForEach-Object { "src\MeetingCleanupOnPrem.$_.ps1" })

function Get-RelativeFile([string]$Base) {
    Get-ChildItem -LiteralPath $Base -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($Base.Length + 1) } | Sort-Object { $_ } -CaseSensitive
}
function Assert-OwnerSigned([string]$Path) {
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne $signer -or $null -eq $signature.TimeStamperCertificate) {
        throw "Not signed with the trusted, timestamped owner certificate ($($signature.Status)): $Path"
    }
}
function Write-Manifest([string]$Package) {
    $entries = foreach ($relative in Get-RelativeFile $Package) {
        if ($relative -ceq 'PackageIntegrity.psd1' -or $relative -ceq 'config\MeetingCleanupOnPrem.config.psd1') { continue }
        if ($relative.Split('\')[0] -cin @('reports', 'logs')) { continue }
        [pscustomobject]@{ Path = $relative; Sha256 = (Get-FileHash -LiteralPath (Join-Path $Package $relative) -Algorithm SHA256).Hash }
    }
    $text = [Text.StringBuilder]::new()
    [void]$text.AppendLine('@{')
    [void]$text.AppendLine('    SchemaVersion = 1')
    [void]$text.AppendLine("    Algorithm = 'SHA256'")
    [void]$text.AppendLine("    SignerThumbprint = '$signer'")
    [void]$text.AppendLine("    Version = '$version'")
    [void]$text.AppendLine('    Files = @(')
    foreach ($entry in $entries) {
        if ($entry.Path.Contains("'")) { throw "A packaged path contains a quote: $($entry.Path)" }
        [void]$text.AppendLine("        @{ Path = '$($entry.Path)'; Sha256 = '$($entry.Sha256)' }")
    }
    [void]$text.AppendLine('    )')
    [void]$text.AppendLine('}')
    $target = Join-Path $Package 'PackageIntegrity.psd1'
    [IO.File]::WriteAllText($target, ($text.ToString() -replace "`r?`n", "`r`n"), [Text.UTF8Encoding]::new($false))
    $target
}

if ($PSCmdlet.ParameterSetName -eq 'Manifest') {
    $Manifest = [IO.Path]::GetFullPath($Manifest).TrimEnd('\')
    foreach ($relative in $signedScripts) { Assert-OwnerSigned (Join-Path $Manifest $relative) }
    $path = Write-Manifest $Manifest
    Write-Output "Unsigned manifest written: $path"
    Write-Output "Sign it, then copy the signed file to $(Join-Path $root 'PackageIntegrity.psd1') and run: .\tools\New-MeetingCleanupOnPremPackage.ps1 -Seal `"$Manifest`""
    return
}

if ($PSCmdlet.ParameterSetName -eq 'Seal') {
    $Seal = [IO.Path]::GetFullPath($Seal).TrimEnd('\')
    $signedManifest = Join-Path $root 'PackageIntegrity.psd1'
    if (-not (Test-Path -LiteralPath $signedManifest)) { throw "Copy the signed manifest to $signedManifest first." }
    Assert-OwnerSigned $signedManifest
    $builtManifest = Join-Path $Seal 'PackageIntegrity.psd1'
    if (-not (Test-Path -LiteralPath $builtManifest) -or (Get-FileHash -LiteralPath $signedManifest -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $builtManifest -Algorithm SHA256).Hash) {
        Copy-Item -LiteralPath $signedManifest -Destination $builtManifest -Force
    }
    $result = & (Join-Path $Seal 'src\Test-PackageIntegrity.ps1') -Root $Seal -Distribution
    if ($result.Status -cne 'Valid') { throw 'The built package is not Valid.' }
    $sourceResult = & (Join-Path $root 'src\Test-PackageIntegrity.ps1') -Root $root
    if ($sourceResult.Status -notlike 'Valid*') { throw 'The repository package folder is not Valid.' }
    foreach ($relative in Get-RelativeFile $Seal) {
        $sourceFile = Join-Path $root $relative
        if ($relative -ceq 'config\MeetingCleanupOnPrem.config.psd1' -or -not (Test-Path -LiteralPath $sourceFile)) { continue }
        if ((Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath (Join-Path $Seal $relative) -Algorithm SHA256).Hash) {
            throw "The packaged file differs from the repository file: $relative"
        }
    }
    $changelog = Join-Path $repoRoot 'CHANGELOG.md'
    if ((Get-FileHash -LiteralPath $changelog -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath (Join-Path $Seal 'CHANGELOG.md') -Algorithm SHA256).Hash) {
        throw 'The packaged CHANGELOG.md differs from the repository CHANGELOG.md.'
    }
    foreach ($used in 'data', 'reports', 'logs', 'bin') {
        if (Test-Path -LiteralPath (Join-Path $Seal $used)) { throw "The built folder contains a $used folder (a package that has been run); it is not sealed: $Seal" }
    }
    $zip = "$Seal.zip"
    if (Test-Path -LiteralPath $zip) { throw "The ZIP already exists: $zip" }
    Compress-Archive -LiteralPath $Seal -DestinationPath $zip -CompressionLevel Optimal
    [pscustomobject]@{
        Package = $Seal; Zip = $zip; ZipBytes = (Get-Item -LiteralPath $zip).Length
        ZipSha256 = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
        Files = $result.Files; Signed = $true; Published = $false
    } | ConvertTo-Json
    return
}

# ---- Build -----------------------------------------------------------------------------------------------
if (-not $Destination) { $Destination = Join-Path (Split-Path $repoRoot -Parent) "package\MeetingCleanupOnPrem-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')
$zip = if ($Unsigned) { "$Destination-unsigned-ci.zip" } else { "$Destination.zip" }
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
if (-not $Unsigned) { foreach ($relative in $signedScripts) { Assert-OwnerSigned (Join-Path $root $relative) } }

# ---- HTML guides: the committed files, rebuilt only on request ----------------------------------------
if ($RebuildDocs) {
    if ($Unsigned) { throw '-RebuildDocs cannot be combined with -Unsigned: the committed manifest lists the committed HTML guides.' }
    & (Join-Path $PSScriptRoot 'Build-Documentation.ps1') | Out-Null
}

# ---- Files needed at run time -------------------------------------------------------------------
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-MeetingCleanupOnPrem.ps1', 'MeetingCleanupOnPrem.psd1', 'MeetingCleanupOnPrem.psm1', 'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md',
    'config\MeetingCleanupOnPrem.config.psd1', 'templates\Report.template.html', 'docs\MeetingCleanupOnPrem-UserGuide.html', 'docs\MeetingCleanupOnPrem-Guide.html', 'src\Test-PackageIntegrity.ps1') { $files.Add($f) }
Get-ChildItem -LiteralPath (Join-Path $root 'src') -File | Where-Object { $_.Extension -in '.ps1', '.cs' -and $_.Name -ne 'Test-PackageIntegrity.ps1' } | ForEach-Object { $files.Add("src\$($_.Name)") }
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

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
if ($Unsigned) {
    $committed = Join-Path $root 'PackageIntegrity.psd1'
    if (Test-Path -LiteralPath $committed) {
        Copy-Item -LiteralPath $committed -Destination (Join-Path $Destination 'PackageIntegrity.psd1')
        & (Join-Path $Destination 'src\Test-PackageIntegrity.ps1') -Root $Destination -DevelopmentMode | Out-Null
    } else {
        Write-Warning 'package\PackageIntegrity.psd1 is not committed: this unsigned build carries no manifest.'
    }
    # The module loads from the package, and its reports go under the package folder.
    $loaded = & pwsh -NoProfile -Command "Import-Module '$Destination\MeetingCleanupOnPrem.psd1' 3>`$null; (Get-Command -Module MeetingCleanupOnPrem).Count; (Import-McoConfiguration).OutputPath"
    if ($LASTEXITCODE -ne 0 -or [int]$loaded[0] -lt 20 -or -not ([string]$loaded[1]).StartsWith($Destination)) { throw "The module does not load correctly from the package: $loaded" }
    Compress-Archive -LiteralPath $Destination -DestinationPath $zip -CompressionLevel Optimal
    Write-Host ''
    Write-Host "  Meeting Cleanup On-Prem $version - unsigned CI package ready (never a release)" -ForegroundColor Yellow
    Write-Host "  Folder   : $Destination"
    Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
    Write-Host ("  Zip      : {0} ({1:N1} MB)" -f $zip, ((Get-Item -LiteralPath $zip).Length / 1MB))
    Write-Host ''
    return
}
Write-Output "Package built (no manifest yet): $Destination ($($all.Count) files)"
Write-Output "Next: .\tools\New-MeetingCleanupOnPremPackage.ps1 -Manifest `"$Destination`""
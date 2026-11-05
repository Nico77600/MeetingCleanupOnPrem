<#
    Meeting Cleanup On-Prem - Exchange Server meeting cleanup.
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolRoot = $PSScriptRoot
$script:ToolVersion = '1.1.1'
$script:LogWriter = $null
$script:LogPath = $null
$script:Quiet = $false
$script:Ews = $null
$script:Ui = $null

# Package integrity. A package carries the owner-signed PackageIntegrity.psd1 manifest; when it is present, the signed
# verifier (src\Test-PackageIntegrity.ps1) must accept the tree before anything is compiled or loaded. A source tree
# without a manifest is an unverified development tree: the entry script refuses it unless -DevelopmentMode is given.
function Assert-McoPackageIntegrity {
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot)
    $packageRoot = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Root))
    $manifest = Join-Path $packageRoot 'PackageIntegrity.psd1'
    if ((Test-Path -LiteralPath $manifest) -and -not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        throw 'Package integrity: PackageIntegrity.psd1 must be a file.'
    }
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        # Without a manifest nothing is verified: say so once per process.
        if (-not [AppDomain]::CurrentDomain.GetData('MeetingCleanupOnPrem.UnverifiedTreeWarned')) {
            [AppDomain]::CurrentDomain.SetData('MeetingCleanupOnPrem.UnverifiedTreeWarned', $true)
            Write-Warning 'Package integrity: no PackageIntegrity.psd1; this is an unverified source tree (package byte integrity and publisher authentication are not established).'
        }
        return
    }
    $verifier = Join-Path $packageRoot 'src\Test-PackageIntegrity.ps1'
    $cursor = [IO.Path]::GetFullPath($verifier)
    while ($cursor) {
        if (-not (Test-Path -LiteralPath $cursor)) { throw 'Package integrity: the verifier is missing; nothing was compiled or loaded.' }
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature.'
    }
    $result = @(& $verifier -Root $packageRoot)
    if ($result.Count -ne 1 -or -not $result[0].PSObject.Properties['Status'] -or $result[0].Status -notlike 'Valid*') {
        throw 'Package integrity: the verifier did not return an explicit Valid result; nothing was loaded.'
    }
}
Assert-McoPackageIntegrity
# Compiled helpers (src\MeetingCleanupOnPrem.Native.cs): compiled in memory, once per PowerShell process, from the source
# that the package integrity check has just verified. No assembly is written to disk or loaded from disk.
$native = 'MeetingCleanupOnPremNative.Fast' -as [type]
if (-not $native) {
    Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'src\MeetingCleanupOnPrem.Native.cs')))
}
elseif ($native::Version -ne $script:ToolVersion) { throw "Meeting Cleanup On-Prem $($native::Version) is already loaded in this PowerShell session: open a new PowerShell window to use $($script:ToolVersion)." }
foreach ($part in 'Console', 'Config', 'Ews', 'Rps', 'Exchange', 'Search', 'Cleanup', 'Restore', 'Transfer', 'Report') {
    . (Join-Path $PSScriptRoot "src\MeetingCleanupOnPrem.$part.ps1")
}

Export-ModuleMember -Function @(
    'Import-McoConfiguration', 'Test-McoConfiguration', 'New-McoRequest', 'Test-McoRequest',
    'Connect-McoEws', 'Disconnect-McoEws', 'Get-McoSavedCredential', 'Connect-McoExchangeShell', 'Disconnect-McoExchangeShell', 'Resolve-McoOrganizer', 'Find-McoMeetings',
    'Get-McoCleanupPlan', 'Invoke-McoCleanup', 'Import-McoRestoreSource', 'Get-McoRestorePlan',
    'Invoke-McoRestore', 'Resolve-McoNewOrganizer', 'Get-McoTransferPlan', 'Invoke-McoTransfer',
    'Export-McoReport', 'New-McoRunFolder',
    'Start-McoLog', 'Stop-McoLog', 'Write-McoLog', 'Write-McoBanner', 'Write-McoStep',
    'Write-McoItem', 'Write-McoSummary', 'Initialize-McoSteps', 'Write-McoNextStep',
    'Write-McoRunBanner', 'Write-McoMeetingTable', 'Write-McoRunSummary', 'Format-McoDuration'
)

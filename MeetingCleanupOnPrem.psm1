<#
    Meeting Cleanup On-Prem - Exchange Server meeting cleanup.
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolRoot = $PSScriptRoot
$script:ToolVersion = '1.1.0'
$script:LogWriter = $null
$script:LogPath = $null
$script:Quiet = $false
$script:Ews = $null
$script:Ui = $null

# Compiled helpers (src\MeetingCleanupOnPrem.Native.cs): once per PowerShell process. The compilation can take tens
# of seconds on a busy server: the assembly is kept in the profile (LocalAppData) by version, content and PowerShell
# version, and only loaded on the next runs.
$native = 'MeetingCleanupOnPremNative.Fast' -as [type]
if (-not $native) {
    $source = Join-Path $PSScriptRoot 'src\MeetingCleanupOnPrem.Native.cs'
    $code = [IO.File]::ReadAllText($source)
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes("$code|$($PSVersionTable.PSVersion)"))).Substring(0, 16)
    $cache = [Environment]::GetFolderPath('LocalApplicationData')
    if (-not $cache) { $cache = [IO.Path]::GetTempPath() }
    $cache = Join-Path $cache 'MeetingCleanupOnPrem'
    $dll = Join-Path $cache "MeetingCleanupOnPrem.Native-$($script:ToolVersion)-$hash.dll"
    $loaded = $false
    try {
        if (-not (Test-Path -LiteralPath $dll)) {
            [void][IO.Directory]::CreateDirectory($cache)
            $temp = "$dll.$PID.tmp"
            Add-Type -TypeDefinition $code -OutputAssembly $temp -OutputType Library
            try { [IO.File]::Move($temp, $dll) } catch { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
            # Assemblies of other versions are not needed any more (one in use stays).
            Get-ChildItem -LiteralPath $cache -Filter 'MeetingCleanupOnPrem.Native-*.dll' -File -ErrorAction SilentlyContinue | Where-Object FullName -ne $dll | Remove-Item -Force -ErrorAction SilentlyContinue
        }
        Add-Type -Path $dll
        $loaded = $true
    }
    catch { Remove-Item -LiteralPath $dll -Force -ErrorAction SilentlyContinue }
    if (-not $loaded) { Add-Type -TypeDefinition $code }
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

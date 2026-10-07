@{
    RootModule        = 'MeetingCleanupOnPrem.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '5a657b89-9c11-4ba4-9d2e-7d2130df90a1'
    Author            = 'Nicolas Fabert'
    Copyright         = '(c) 2026 Nicolas Fabert. MIT License.'
    Description       = 'Meeting Cleanup On-Prem: finds the meetings of organizers or rooms in every calendar of Exchange Server (EWS, Exchange PowerShell), then removes them silently, has the organizer cancel them, transfers them to a new organizer, or restores them; CSV, JSON and HTML reports.'
    PowerShellVersion = '7.4'
    FunctionsToExport = @(
        'Import-McoConfiguration', 'Test-McoConfiguration', 'New-McoRequest', 'Test-McoRequest',
        'Connect-McoEws', 'Disconnect-McoEws', 'Get-McoSavedCredential', 'Connect-McoExchangeShell', 'Disconnect-McoExchangeShell', 'Resolve-McoOrganizer', 'Find-McoMeetings',
        'Get-McoCleanupPlan', 'Invoke-McoCleanup', 'Import-McoRestoreSource', 'Get-McoRestorePlan',
        'Invoke-McoRestore', 'Resolve-McoNewOrganizer', 'Get-McoTransferPlan', 'Invoke-McoTransfer',
        'Export-McoReport', 'New-McoRunFolder',
        'Start-McoLog', 'Stop-McoLog', 'Write-McoLog', 'Write-McoBanner', 'Write-McoStep',
        'Write-McoItem', 'Write-McoSummary', 'Initialize-McoSteps', 'Write-McoNextStep',
        'Write-McoRunBanner', 'Write-McoMeetingTable', 'Write-McoRunSummary', 'Format-McoDuration'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags = @('ExchangeServer', 'OnPremises', 'EWS', 'Calendar', 'Meeting', 'Cleanup')
            LicenseUri = 'https://opensource.org/licenses/MIT'
            ProjectUri = 'https://github.com/Nico77600/MeetingCleanupOnPrem'
        }
    }
}

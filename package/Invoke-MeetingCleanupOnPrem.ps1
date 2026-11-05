<#
.SYNOPSIS
    Meeting Cleanup for Exchange Server On-Premises.

.DESCRIPTION
    Searches organizer, attendee and room copies through EWS (Exchange Server 2016, 2019, SE), with the
    directory, groups and Recoverable Items through Exchange PowerShell (local session or remote PowerShell).
      Report    read only.
      Remove    removes the attendee and room copies silently (DeleteItem SoftDelete, no message); the meeting
                stays in the organizer's calendar.
      Cancel    the organizer cancels the meeting (message to the attendees, rooms released), then the copies
                left are removed silently.
      Restore   -FromReport <Remove run>: the copies removed come back from Recoverable Items.
      Transfer  -NewOrganizer: the meeting is re-created and sent by the new organizer (a series from its next
                occurrence), the old one is cancelled by its organizer or removed silently.
    Remove, Cancel and Transfer write a backup first and ask to type YES (-Force skips it).
    A series is acted on whole; -SeriesScope Occurrences (Search.SeriesScope) limits it to its occurrences in the
    period (a period of one day: one occurrence; give -Start and -End with an action). Rooms mode always does.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.1
    Console : a live progress line (bar, part done, time left) in an interactive console; colours off when the
              output is redirected or NO_COLOR is set (MCO_FORCE_COLOR=1 forces them); MCO_ICONS = Emoji |
              Symbols | Ascii forces the style of the icons.
    Exit codes : 0 = completed, 1 = failed, 2 = finished with warnings (a copy not removed or restored, a mailbox not read...).
    Integrity : a release package is verified (owner signature, SHA256 of every file) before anything loads;
                -DevelopmentMode starts a source tree that has no PackageIntegrity.psd1 (a present manifest is always verified).
    Documentation : docs\MeetingCleanupOnPrem-UserGuide.html (user guide: prerequisites, everyday commands) and
                    docs\MeetingCleanupOnPrem-Guide.html (developer guide); sources: docs\*.md

    Disclaimer:
    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

    Removing, cancelling or transferring meetings changes real calendars: always start with a report, and test it in your environment before production use.
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string[]]$Organizer,
    [string]$OrganizerFile,
    [string[]]$Room,
    [string]$RoomFile,
    [string[]]$Mailbox,
    [string]$MailboxFile,
    [datetime]$Start,
    [datetime]$End,
    [string]$Subject,
    [ValidateSet('Whole', 'Occurrences')][string]$SeriesScope,
    [string[]]$MeetingId,
    [ValidateSet('Organizer', 'Rooms', 'Mailboxes', 'AllMailboxes')][string[]]$SearchIn,
    [ValidateSet('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')][string]$Action = 'Report',
    [string]$Comment,
    [string]$NewOrganizer,
    [string]$FromReport,
    [switch]$Force,
    [switch]$NoReport,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\MeetingCleanupOnPrem.config.psd1'),
    [string]$EwsUrl,
    [string]$EwsMailbox,
    [ValidateSet('Windows', 'Basic')][string]$Authentication,
    [ValidateSet('Self', 'Delegate', 'Impersonation')][string]$AccessMode,
    [string]$CredentialUser,
    [ValidateSet('Existing', 'Auto', 'Rps')][string]$ManagementShellMode,
    [string]$ManagementShellServer,
    [string]$ManagementShellUri,
    [switch]$DevelopmentMode
)

$ErrorActionPreference = 'Stop'
# Numbers and dates of the console and the report in one format (1,254 - 2026-10-06), whatever the Windows culture.
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$exitCode = 1
$loaded = $false
$dot = [char]0x00B7

function Confirm-McoAction {
    <# Shows the plan and asks to type YES; throws when the answer is not YES or the console cannot ask. #>
    param([string]$Title, [string[]]$Lines)
    Write-Host ''
    Write-Host "  $Title - $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Yellow
    foreach ($line in $Lines) { Write-Host "    - $line" -ForegroundColor Yellow }
    if ([Console]::IsInputRedirected -or -not [Environment]::UserInteractive) { throw 'Confirmation needed: run interactively, or add -Force.' }
    if ((Read-Host '  Type YES to continue') -cne 'YES') { throw 'Cancelled: nothing was changed.' }
    Write-McoLog 'INFO' "Confirmed by $([Environment]::UserName): $($Lines -join ' | ')"
}

try {
    # The package is verified before any of its code runs: the verifier must carry the owner signature, then it
    # checks the signed manifest. A tree without a manifest is accepted only with -DevelopmentMode.
    $integrityManifest = Join-Path $PSScriptRoot 'PackageIntegrity.psd1'
    $integrityHint = 'Use the complete release ZIP (https://github.com/Nico77600/MeetingCleanupOnPrem/releases/latest): extract it again and unblock it. A development or CI build is started from its folder with pwsh -NoProfile -File .\Invoke-MeetingCleanupOnPrem.ps1 <parameters> -DevelopmentMode.'
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $integrityManifest) {
        $integrityVerifier = Join-Path $PSScriptRoot 'src\Test-PackageIntegrity.ps1'
        $integrityCursor = [IO.Path]::GetFullPath($integrityVerifier)
        while ($integrityCursor) {
            $integrityItem = Microsoft.PowerShell.Management\Get-Item -LiteralPath $integrityCursor -Force -ErrorAction Stop
            if ($integrityItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $integrityCursor" }
            $integrityCursor = [IO.Path]::GetDirectoryName($integrityCursor)
        }
        $integritySignature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $integrityVerifier -ErrorAction Stop
        if ([string]$integritySignature.Status -cne 'Valid' -or $null -eq $integritySignature.SignerCertificate -or
            $integritySignature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
            $null -eq $integritySignature.TimeStamperCertificate) {
            throw "Package integrity: verifier must have the trusted, timestamped owner signature. $integrityHint"
        }
        if ($DevelopmentMode) { & $integrityVerifier -Root $PSScriptRoot -DevelopmentMode | Out-Null } else { & $integrityVerifier -Root $PSScriptRoot -Distribution | Out-Null }
    }
    elseif (-not $DevelopmentMode) {
        throw "Package integrity: PackageIntegrity.psd1 is missing. $integrityHint"
    }
    else {
        Write-Warning 'DevelopmentMode: no package manifest; package byte integrity and publisher authentication are not established.'
        # The module does not repeat this warning in the same process.
        [AppDomain]::CurrentDomain.SetData('MeetingCleanupOnPrem.UnverifiedTreeWarned', $true)
    }
    Import-Module (Join-Path $PSScriptRoot 'MeetingCleanupOnPrem.psd1') -Force
    $loaded = $true
    $settings = Import-McoConfiguration -Path $ConfigPath -Root $PSScriptRoot
    foreach ($pair in @(
        @('EwsUrl', $EwsUrl), @('Mailbox', $EwsMailbox), @('Authentication', $Authentication), @('AccessMode', $AccessMode), @('CredentialUser', $CredentialUser),
        @('ManagementShellMode', $ManagementShellMode), @('ManagementShellServer', $ManagementShellServer), @('ManagementShellUri', $ManagementShellUri)
    )) { if ($pair[1]) { $settings[$pair[0]] = $pair[1] } }
    $check = Test-McoConfiguration $settings
    if (-not $check.IsValid) { throw ("Invalid value:`n - " + ($check.Problems -join "`n - ")) }
    $logPath = Start-McoLog -Directory $settings.LogPath -RetentionDays $settings.LogRetentionDays
    $request = New-McoRequest -Settings $settings -Organizer $Organizer -OrganizerFile $OrganizerFile -Room $Room -RoomFile $RoomFile -Mailbox $Mailbox -MailboxFile $MailboxFile -Start $(if ($PSBoundParameters.ContainsKey('Start')) { $Start } else { $null }) -End $(if ($PSBoundParameters.ContainsKey('End')) { $End } else { $null }) -Subject $Subject -MeetingId $MeetingId -SearchIn $SearchIn -Action $Action -Comment $Comment -FromReport $FromReport -NewOrganizer $NewOrganizer -SeriesScope $SeriesScope
    $requestCheck = Test-McoRequest $request
    if (-not $requestCheck.IsValid) { throw ("Cannot run:`n - " + ($requestCheck.Problems -join "`n - ")) }
    Write-McoRunBanner -Settings $settings -Request $request -LogPath $logPath -NoReport:$NoReport

    $credential = $null
    if ($settings.CredentialFile) { $credential = Get-McoSavedCredential -Settings $settings }
    elseif ($settings.Authentication -eq 'Basic' -or $settings.CredentialUser) {
        $user = if ($settings.CredentialUser) { $settings.CredentialUser } else { $settings.Mailbox }
        $credential = Get-Credential -UserName $user -Message 'Credentials used for the Exchange Server EWS endpoint'
    }
    $rpsCredential = $credential
    if ($settings.ManagementShellCredentialUser -and (-not $rpsCredential -or $rpsCredential.UserName -ne $settings.ManagementShellCredentialUser)) {
        $rpsCredential = Get-Credential -UserName $settings.ManagementShellCredentialUser -Message 'Credentials used for Exchange Server Remote PowerShell'
    }
    elseif ($settings.ManagementShellMode -eq 'Rps' -and $settings.ManagementShellAuthentication -eq 'Basic' -and -not $rpsCredential) {
        $rpsCredential = Get-Credential -Message 'Credentials used for Exchange Server Remote PowerShell'
    }
    # Exchange Server, then: Restore (copies of the run, restore, verify) | search (organizers, mailboxes, search,
    # attendees) or the meetings of a report | the action and its verification | the report.
    $verify = [int][bool]$settings.Verify
    $steps = switch ($Action) {
        'Restore' { 3 + $verify }
        'Report' { 5 }
        default { $(if ($FromReport) { 2 } else { 5 }) + 1 + $verify }
    }
    if (-not $NoReport -or $Action -ne 'Report') { $steps++ }
    Initialize-McoSteps -Total $steps
    Write-McoNextStep 'Exchange Server' 'Server'
    $shell = Connect-McoExchangeShell -Settings $settings -Credential $rpsCredential
    $hasCmdlets = [bool](Get-Command Get-Recipient -ErrorAction SilentlyContinue)
    Write-McoItem $(if ($shell.Connected -or $hasCmdlets) { 'Ok' } else { 'Warn' }) ("Exchange PowerShell: {0}" -f $(if ($shell.Connected) { $shell.Detail } elseif ($hasCmdlets) { 'cmdlets of the current session' } else { 'not available (aliases, groups, rooms and restore need it)' })) -Icon Shield
    $ews = Connect-McoEws -Settings $settings -Credential $credential
    Write-McoItem Ok ("EWS {0} {1} {2} {1} {3}{4}" -f $ews.EwsUrl, $dot, $(if ($ews.ServerVersion) { $ews.ServerVersion } else { 'connected' }), $settings.AccessMode, $(if ($settings.AccessMode -ne 'Self') { " ($($settings.Mailbox) signs in)" } else { '' })) -Icon Server
    if ($Action -eq 'Restore') {
        Write-McoNextStep 'Copies removed by the run' 'File'
        $result = Import-McoRestoreSource -Path $FromReport -MeetingId $MeetingId
        $plan = Get-McoRestorePlan $result
        Write-McoItem Ok ('{0} meeting(s) of the {1} run {2} {3} cop{4} to restore' -f $result.Meetings.Count, $result.SourceAction, $dot, $plan.Restore.Count, $(if ($plan.Restore.Count -eq 1) { 'y' } else { 'ies' })) -Icon File
        Write-McoMeetingTable -Meetings @($result.Meetings)
        if ($plan.Restore.Count -and -not $Force) { Confirm-McoAction 'Restore from Recoverable Items' $plan.Lines }
        $result = Invoke-McoRestore -Settings $settings -Result $result
        Write-McoMeetingTable -Meetings @($result.Meetings)
        $runPath = New-McoRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action 'Restore'
    }
    else {
        if ($FromReport) {
            Write-McoNextStep 'Meetings of the report' 'File'
            $result = Import-McoRestoreSource -Path $FromReport -MeetingId $MeetingId
            $result.Action = 'Report'
            Write-McoItem Ok ('{0} meeting(s), {1} cop{2} from {3}' -f $result.Counts.Meetings, $result.Counts.Copies, $(if ($result.Counts.Copies -eq 1) { 'y' } else { 'ies' }), $result.FromReport) -Icon File
        }
        else { $result = Find-McoMeetings -Settings $settings -Request $request }
        Write-McoMeetingTable -Meetings @($result.Meetings)
        $runPath = $null
        if ($Action -eq 'Transfer' -and @($result.Meetings).Count) {
            $target = Resolve-McoNewOrganizer -Address $NewOrganizer
            $plan = Get-McoTransferPlan -Result $result -NewOrganizer $target -Method $settings.TransferMethod -Comment $request.Comment
            if (-not $Force) { Confirm-McoAction 'Transfer' (@($plan.Lines) + 'A backup of the meetings is written first.') }
            $runPath = New-McoRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action 'Transfer'
            $result = Invoke-McoTransfer -Settings $settings -Result $result -Plan $plan -Comment $request.Comment -BackupPath (Join-Path $runPath "$($settings.ReportPrefix)-Backup.json")
            Write-McoMeetingTable -Meetings @($result.Meetings)
        }
        elseif ($Action -in 'Remove', 'Cancel' -and @($result.Meetings).Count) {
            $plan = Get-McoCleanupPlan -Result $result -Action $Action
            if (-not $Force) {
                $lines = @($plan.Lines)
                if ($Action -eq 'Cancel') { $lines += "Message: $($request.Comment)" }
                $lines += 'A backup of the meetings is written first; the copies removed can be restored for the retention of deleted items (-Action Restore -FromReport <report folder>).'
                Confirm-McoAction $Action $lines
            }
            $runPath = New-McoRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action $Action
            $result = Invoke-McoCleanup -Settings $settings -Result $result -Action $Action -Comment $request.Comment -BackupPath (Join-Path $runPath "$($settings.ReportPrefix)-Backup.json")
            Write-McoMeetingTable -Meetings @($result.Meetings)
        }
        elseif ($Action -ne 'Report') { $result.Action = $Action }
    }
    $reportText = 'none (-NoReport)'
    if (-not $NoReport -or $runPath) {
        Write-McoNextStep 'Report' 'Report'
        $report = Export-McoReport -Result $result -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats $settings.ReportFormats -Delimiter $settings.CsvDelimiter -Directory $runPath -SummaryOnly:$NoReport
        foreach ($f in $report.Files.Values) { Write-McoItem Ok $f -Icon File }
        $reportText = if ($report.Files.Contains('Html')) { $report.Files.Html } else { $report.Directory }
    }
    Write-McoRunSummary -Result $result -ReportText $reportText -LogPath $logPath
    $exitCode = switch ($result.Status) { 'Completed' { 0 } 'Failed' { 1 } default { 2 } }
}
catch {
    if ($loaded) {
        $errorText = [string]$_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($errorText)) { $errorText = 'Unknown error.' }
        Write-McoItem Fail $errorText
        $stack = [string]$_.ScriptStackTrace
        if ($stack) { Write-McoLog 'ERROR' ($stack -replace '\r?\n', ' | ') }
        Write-Host ''
    }
    else { Write-Host "Meeting Cleanup On-Prem: $($_.Exception.Message)" -ForegroundColor Red }
    $exitCode = 1
}
finally {
    if ($loaded) { Disconnect-McoEws; Disconnect-McoExchangeShell; Stop-McoLog }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode

# SIG # Begin signature block
# MIImIgYJKoZIhvcNAQcCoIImEzCCJg8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCAV8fzmQYFCVYb
# vJB506uh169tw9DbLMWfQ8j2xEKNUaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYCTCCGAUCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAARTAwGBWN
# GSTuHnG+DS7rQcihwc70CzW9hqIibExgZjANBgkqhkiG9w0BAQEFAASCAgAAAsnj
# GVvHCR6kNT3RcCRPK57M8Zui6WFiFZDFRpqnRSdyeS0YVVPqrjeJL1WuUgdkT5bX
# DodJbaEMTq3hzbOoJfUZfPYhuVPiAd4ztoybn+a9b3Km+tu4Nx9alsafSpCPMH+f
# x/wXboxRKaFC9L5wAivzX99zov2M0KxA9e6IpM90YjAGXIYu6ZTDpUCtO5foNmOD
# RKUeZAnSGQGQNHMByM7FayM4pORDbORW1UXvxnvLTwcTf5Fu0wB7chUe2YqRAc8o
# r5UKXoW81I2/IleuXIzCTgVLjfwIr5Rn4xCBrBWDWOGI6khzhO9UwQ2wDgtjU12F
# V40vT9G9AEtt8vHi09lJfJ895+Ra4wc8f2olZYEsm0B1XHl4xHTrWyBS4J+RWoB4
# Z++hO690E5A2piLWA659JdkZeqNsAj0nzt4wX2HybzoSt9VcWuHVVep/g4L9zRMG
# r6edpKHpDV0J6Hyl55em82ccJ6Lqj3yiKGy39sEtj6fmBuq2FxG+hfJfzOG4PWmK
# FEd/FLm0ruMzYKSDD5gv00MAgCO1xcLOYS0Rft7N2xRRAdvxRqOtHvkKR1ePi7kD
# WYtbATRkXx6RKoUUqyPYZA10ovV2EQ766JQ6Sp+KV6xW8ed+dnE1ddK6mijhHzcN
# pLJ/T5k8Aco5Q+L+xmXJmVl5EYHiDW9rfMRyAKGCFMYwghTCBgorBgEEAYI3AwMB
# MYIUsjCCFK4GCSqGSIb3DQEHAqCCFJ8wghSbAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIO41N4RBspcMW48ElRa98OHOdhT8HOvdh89BOhc7tSKuAgh/3rWCLNc5
# QxgPMjAyNjExMDUxNzIyMzNaMAMCAQECCBB6DP7rjS+xoIIRpzCCBPMwggLboAMC
# AQICEFRD+De9Wu4YAYRyuFzajuMwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhT
# U0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNzdWluZyBS
# U0EgQ0EgMjAeFw0yNjA2MTIxNzIyNDZaFw0zNjA2MDkxNzIyNDVaMG4xCzAJBgNV
# BAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UE
# CgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1waW5nIFVuaXQg
# MjAyNiBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABOgli7hmEy5IkjjzCzbJ
# DgbYV7CIUUgOOn3jkML0i3A2jYer8VMwShMdSnRQURbOzh65x/ytFQOjGw0CBaVI
# ND2jggFSMIIBTjAfBgNVHSMEGDAWgBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDBNBggr
# BgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5jb20vU1NM
# LmNvbS10aW1lU3RhbXBpbmctSS1SMi5jZXIwUQYDVR0gBEowSDA8BgwrBgEEAYKp
# MAEDBgEwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0
# b3J5MAgGBmeBDAEEAjAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5
# MDegNaAzhjFodHRwOi8vY3Jscy5zc2wuY29tL1NTTC5jb20tdGltZVN0YW1waW5n
# LUktUjIuY3JsMB0GA1UdDgQWBBQnrANWuihquhc7gz1Af3moJOzDPDAOBgNVHQ8B
# Af8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE92bs78A57CGU1wZ/oY5rgV4Mwd
# dn22MU/bmHumb4b6W2VinMg0IGnBE5D8PJOsnqp2AmMhWdOt5TDxKgIQA9pzBute
# GYez7R1vQ1qn7IIr8Wz2AXjfre1MYJGE93saW1jHTxKgbVlchzry3CrzMYCVnyYw
# DcFZMFQ35Mv4+3H9vPoopjfDynO0v7vPy3p0UZ1dltmH53EZ5M/Qn7OZ0Q82mYLo
# h666dhscIJzUEh+BgYxzpdz8y7J1WxqGSouGRI4F7qbOJhhjiu/hgdhwJUJ/jkve
# OHB6pZB9IBJbdlns6AjUaRC6YfamsbgZ40hKl1ln4BOshgcDtLsSjaqVlUx5gMK+
# krwT6H8bm1yLhJ/Yr0eAizrgMVEqyeNWQYP6qhtKCYxehpESm8au3MlkpBGWCYZi
# goNh4OX1x2oDsYsG5p+ST2db4Vgq25N7/3U8doL0yufnqwuR/f603JY1CO++sndq
# 34Y35Kb8gLG7V7OA+5658PIvcKK+fn3bUUiQVcasy5spBX7aC9FKdkPvrciglHa4
# m3XJ6weBykt7PrdwI3XGwNOgwi/Bon0IvXy9MkQUTZQOKOgT/kq++eTGgjY1EWGQ
# 7w8sSRwQU+3Ko2SMLIkfLPuvOtJ6BWwQbXYHtiMOwjwv9UN9SNu4VaJUuwdMmqN7
# /F9jLVst5rGZV03BMIIF3TCCA8WgAwIBAgIIeyyb0xaAMpkwDQYJKoZIhvcNAQEL
# BQAwfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3Vz
# dG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20g
# Um9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0EwHhcNMTYwMjEyMTczOTM5
# WhcNNDEwMjEyMTczOTM5WjB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAPkP3aMrfcvQKv7sZ4Wm5y4b
# unfh4/WvpOz6Sl2RxFdHaxh3a3by/ZPkPQ/CFp4LZsNWlJ4Xg4XOVu/yFv0AYvUi
# CVToZRdOQbngT0aXqhvIuG5iXmmxX9sqAn78bMrzQdjt0Oj8P2FI7bADFB0QDksZ
# 4LtO7IZl/zbzXmcCC52GVWH9ejjt/uIZALdvoVBidXQ8oPrIJZK0bnoix/geoeOy
# 3ZExqysdBP+lSgQ36YWkMyv94tZVNHwZpEpox7Ko07fKoZOI68GXvIz5HdkihCR0
# xwQ9aqkpk8zruFvh/l8lqjRYyMEjVJ0bmBHDOJx+PYZspQ9AhnwC9FwCTyjLrnGf
# DzrIM/4RJTXq/LrFYD3ZfBjVsqnTdXgDciLKOsMf7yzlLqn6niy2UUb9rwPW6mBo
# 6oUWNmuF6R7As93EJNyAKoFBbZQ+yODJgUEAnl6/f8UImKIYLEJAs/lvOCdLToD0
# PYFH4Ih86hzOtXVcUS4cK38acijnALXRdMbX5J+tB5O2UzU1/Dfkw/ZdFr4hc96S
# CvigY2q8lpJqPvi8ZVWb3vUNiSYE/CUapiVpy8JtynziWV+XrOvvLsi81xtZPCvM
# 8hnIk2snYxnP/Okm+Mpxm3+T/jRnhE6Z6/yzeAkzcLpmpnbtG3PrGqUNxCITIJRW
# Ck4sbE6x/c+cCbqiM+2HAgMBAAGjYzBhMB0GA1UdDgQWBBTdBAkHovV6fVJTEpKV
# 7jiAJQ2mWTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMS
# kpXuOIAlDaZZMA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG9w0BAQsFAAOCAgEAIBgR
# lCn7Jp0cHh5wYfGVcpNxJK1ok1iOMq8bs3AD/CUrdIWQPXhq9LmLpZc7tRiRux6n
# +UBbkflVma8eEdBcHadm47GUBwwyOabqG7B52B2ccETjit3E+ZUfijhDPwGFpUen
# PUayvOUiaPd7nNgsPgohyC0zrL/FgZkxdMF1ccW+sfAjRfSda/wZY52jvATGGAsl
# u1OJD7OAUN5F7kR/q5R4ZJjT9ijdh9hwZXT7DrkT66cPYakylszeu+1jTBi7qUD3
# oFRuIIhxdRjqerQ0cuAjJ3dctpDqhiVAq+8zD8ufgr6iIPv2tS0a5sKFsXQP+8hl
# AqRSAUfdSSLBv9jra6x+3uxjMxW3IwiPxg+NQVrdjsW5j+VFP3jbutIbQLH+cU0/
# 4IGiul607BXgk90IH37hVZkLId6Tngr75qNJvTYw/ud3sqB1l7UtgYgXZSD32pAA
# n8lSzDLKNXz1PQ/YK9f1JmzJBjSWFupwWRoyeXkLtoh/D1JIPb9s2KJELtFOt3JY
# 04kTlf5Eq/jXixtunLwsoFvVagCvXzfh1foQC5ichucmj87w7G6KVwuA406ywKBj
# YZC6VWg3dGq2ktufoYYitmUnDuy2n0Jg5GfCtdpBC8TTi2EbvPofkSvXRAdeuims
# 2cXp71NIWuuA8ShYIc2wBlX7Jz9TkHCpBB5XJ7kwggbLMIIEs6ADAgECAhBSnbvC
# uImH9YOA4LDrogY2MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTI2MDYxMjE2NDIxNVoXDTQxMDIxMjE3MzkzOVowcjELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgMjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJ8v
# q+ZzTZC8MUabNgd5WlBjVPcu71oiedNd8n9Il8+uW0HcGvW44/Jf0pHl6hxqn0od
# haGnvvNBPJgxH/d6CMI7Mtgl9nF5isarRfByRmAeH+skYShA+dMJGzsrC9NiF3/v
# xGZ2uOI334YVc+DlP5OdsUPz5yIP7hYXNlpPknPwh6gK/pWFZD8+y8uUagfX9USh
# q/SQt+1tS4FXRwudHTSHbGt0Fz4ZbWGbm0d56ROAW34tI6FFuZ+TtGssHjeOUf2F
# S8QNGBpadTh1X1Z7Kl6vG4Z8NS3OJIyelzoF5TpcP2vM3ue0723bmZd3oXsE5rzt
# AJMv56jlg0vSDz8XcXRur6awW0rh5cTQDQQGx04hYJH2Zn9HWWJU/wJ2SKBcQKtB
# yCz6JX+DoFoBE++kolkpZtl2ESF8FY6XqhtKZb18ESEAc0ohCY+0HkBZwrx8eIED
# CDbzd/shv9a8YRG2oxr3tLUz95V+Dzzh1QBNECaeGczyNWGCZIVUa4lVcVUVOBbI
# OAT8D+iPmC5h29bWeES8DhTS4f4CWA4D5QoQqMTqIBcV4MD2Waw8hEWAdGnjnsVt
# DimnDeEUbvXKhdiirvwG+QC+3mAy1RlD6xBMmGyDAKVSqqXm3NP4w0J4gvt65Grk
# QI4FXPJtYqfwkIBtQ8OGfbpVUQGPVUemWtui7E21AgMBAAGjggFRMIIBTTASBgNV
# HRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMSkpXuOIAlDaZZ
# ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3NsLmNv
# bS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNlcjA/BgNVHSAEODA2MDQGBFUd
# IAAwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0b3J5
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9j
# cmxzLnNzbC5jb20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jcmwwHQYDVR0O
# BBYEFHCaZcFDgKP86bFkiW7xMJwJIcl8MA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG
# 9w0BAQsFAAOCAgEAZRRbzTHpgOogYgWKOeiCkI1oVrkGmCrqOKqX9CiTREQLdYM/
# h7ANki98QP1fhAPEkTzT/Ur19IrPuZg7KaWL6SsmBqmD9+mBVtJmxdRdGbV8bpgl
# FDv2VeZgTSg3+UOD2O0G9HlgHqVHNKzSa/zL5+WYl3yT8Q29sZnIPGkQ+udv56/j
# q5NlAT9BON4iXSOCxe1bNvUd0hkS2GTsTNeQyOEHm7CkZdCg5X+DpmBginO7pk8Y
# m5GbTdQtSloJ3/Y2nVPycJ8kmmBAYMev7/hVQukhjjc4mO03LyPTFfqmMR1D5Tpu
# gGkD9nUuG41l55vOpXQm1JMbTgArpkn0/e9O+yTYlWuz2dSAZHf9jxNQAvILO1Pv
# 2JNQTfV0TyIteZP7gNcRfaHnnuSho8ohipKiE+1vEJIFE+c3gopxSpHCmn2THETk
# lWKV7ZrUKhhkdoOLj2KidCtYZrOoCV6x7f9+KfDPT61fKDVntGZdRKwzLV53VIN6
# oSS+KQN2VJaO8OZTOZpfgO4o6/9xpNr+EEaA3hNkoZk13R4LaqPC/309dfZmHBE8
# nBNjWBkV7AZdzdK8w7B0a/D9066a6oUa19OIktD+3kPpQ8YEt8oRp9zs1hoQ7fRn
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJW
# MIICUgIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDUxNzIyMzNaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCBkADD+toqch+tc
# UTwyyVL7xbjSiWYxlkdu98Zoyl6mbzCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEcw
# RQIhALK3EKgrVQ54kaVkijBPpsM2bzEuUUF3I60qQsO31Zi3AiAmyy7NQ/GWgu7p
# QN+4Rs5AY67x5CWq3+MWk1gRTDuWoA==
# SIG # End signature block

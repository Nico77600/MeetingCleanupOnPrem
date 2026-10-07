<#
    Exchange Server Remote PowerShell (RPS) adapter.
    EWS remains independent from this session; RPS only imports Exchange cmdlet
    proxies for directory and Recoverable Items operations.
#>

$script:ExchangeShell = $null

function Get-McoRpsCommand {
    param([Parameter(Mandatory)][string]$Name)
    Get-Command -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1
}

function Connect-McoExchangeShell {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Settings, [pscredential]$Credential)
    $mode = [string]$Settings.ManagementShellMode
    if ($mode -eq 'Existing' -or [string]$Settings.DirectoryMode -eq 'None') {
        return [pscustomobject]@{ Connected = $false; Mode = 'Existing'; Detail = 'The current PowerShell session is used.' }
    }
    $hasRecipient = Get-McoRpsCommand 'Get-Recipient'
    if ($mode -eq 'Auto' -and $hasRecipient) {
        return [pscustomobject]@{ Connected = $false; Mode = 'Existing'; Detail = 'Exchange cmdlets already exist in the current session.' }
    }
    if (-not $Settings.ManagementShellUri) {
        $scheme = if ([string]$Settings.ManagementShellAuthentication -eq 'Kerberos') { 'http' } else { 'https' }
        $Settings.ManagementShellUri = '{0}://{1}/PowerShell/' -f $scheme, $Settings.ManagementShellServer
    }
    if ([string]::IsNullOrWhiteSpace($Settings.ManagementShellServer) -and [string]::IsNullOrWhiteSpace($Settings.ManagementShellUri)) {
        throw 'RPS is enabled but ManagementShell.ServerFqdn or ManagementShell.ConnectionUri is empty.'
    }
    if (-not $Credential -and $Settings.ManagementShellCredentialUser) {
        throw 'The RPS credential must be supplied by the caller when ManagementShell.CredentialUser is configured.'
    }
    $sessionArgs = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri = [string]$Settings.ManagementShellUri
        Authentication = [string]$Settings.ManagementShellAuthentication
        ErrorAction = 'Stop'
    }
    if ($Credential) { $sessionArgs.Credential = $Credential }
    $session = $null
    $imported = $null
    try {
        $session = New-PSSession @sessionArgs
        $imported = Import-PSSession -Session $session -DisableNameChecking -AllowClobber -ErrorAction Stop
        $script:ExchangeShell = @{ Session = $session; Module = $imported; Mode = 'Rps'; Uri = $Settings.ManagementShellUri }
        if (-not (Get-McoRpsCommand 'Get-Recipient')) { throw 'The RPS session connected but Get-Recipient was not imported.' }
        [pscustomobject]@{ Connected = $true; Mode = 'Rps'; Detail = "Exchange cmdlets imported from $($Settings.ManagementShellUri)." }
    }
    catch {
        if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
        $message = [string]$_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($message)) { $message = 'no diagnostic message returned by PowerShell remoting' }
        throw "Exchange RPS connection failed: $message"
    }
}

function Disconnect-McoExchangeShell {
    if (-not $script:ExchangeShell) { return }
    try {
        if ($script:ExchangeShell.Module) { Remove-Module -ModuleInfo $script:ExchangeShell.Module -Force -ErrorAction SilentlyContinue }
        if ($script:ExchangeShell.Session) { Remove-PSSession -Session $script:ExchangeShell.Session -ErrorAction SilentlyContinue }
    }
    finally { $script:ExchangeShell = $null }
}

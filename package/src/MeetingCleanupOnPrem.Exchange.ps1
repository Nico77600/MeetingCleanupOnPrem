<#
    Optional Exchange Management Shell adapter.
    EWS remains the calendar transport; these cmdlets only enrich identities and
    enumerate recipients that EWS cannot discover by itself.
#>

function Get-McoExchangeCommand {
    param([Parameter(Mandatory)][string]$Name)
    Get-Command -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1
}

function Get-McoExchangeProxyAddresses {
    param($Recipient, [string]$Fallback)
    $addresses = [Collections.Generic.List[string]]::new()
    if ($Fallback) { $addresses.Add($Fallback.ToLowerInvariant()) }
    foreach ($value in @(Get-McoProperty $Recipient 'EmailAddresses')) {
        $text = [string]$value
        if ($text -match '^(?i)(smtp|x500):(.+)$') { $addresses.Add($Matches[2].ToLowerInvariant()) }
    }
    $legacy = [string](Get-McoProperty $Recipient 'LegacyExchangeDN')
    if ($legacy) { $addresses.Add($legacy.ToLowerInvariant()) }
    @($addresses | Where-Object { $_ } | Select-Object -Unique)
}

$script:RecipientKinds = @{}

function Get-McoRecipientKind {
    <#
        What an address is in the organization, from Get-Recipient (cached per run):
        Room | Equipment | Group | Mailbox | Other | Unknown (no Exchange cmdlets, or not found / external).
    #>
    param([Parameter(Mandatory)][string]$Address)
    $key = $Address.ToLowerInvariant()
    if ($script:RecipientKinds.ContainsKey($key)) { return $script:RecipientKinds[$key] }
    $kind = 'Unknown'
    $r = Resolve-McoExchangeRecipient -Identity $Address
    if ($r -and -not $r.Error) {
        $kind = switch -Regex ([string]$r.RecipientTypeDetails) {
            '^RoomMailbox$' { 'Room'; break }
            '^EquipmentMailbox$' { 'Equipment'; break }
            'Group' { 'Group'; break }
            'Mailbox$' { 'Mailbox'; break }
            default { 'Other' }
        }
    }
    $script:RecipientKinds[$key] = $kind
    $kind
}

function Resolve-McoExchangeRecipient {
    param([Parameter(Mandatory)][string]$Identity)
    $command = Get-McoExchangeCommand 'Get-Recipient'
    if (-not $command) { return $null }
    try {
        $recipient = @(& $command.Name -Identity $Identity -ErrorAction Stop)[0]
        if (-not $recipient) { return $null }
        $primary = [string](Get-McoProperty $recipient 'PrimarySmtpAddress')
        if (-not $primary) { $primary = [string](Get-McoProperty $recipient 'WindowsEmailAddress') }
        $primary = $primary.ToLowerInvariant()
        [pscustomobject]@{
            Input = $Identity; PrimaryAddress = $primary; DisplayName = [string](Get-McoProperty $recipient 'DisplayName')
            Addresses = @(Get-McoExchangeProxyAddresses -Recipient $recipient -Fallback $primary)
            RecipientTypeDetails = [string](Get-McoProperty $recipient 'RecipientTypeDetails'); Alias = [string](Get-McoProperty $recipient 'Alias'); Error = ''
        }
    }
    catch {
        [pscustomobject]@{ Input = $Identity; PrimaryAddress = ''; DisplayName = ''; Addresses = @(); RecipientTypeDetails = ''; Alias = ''; Error = $_.Exception.Message }
    }
}

function Get-McoExchangeMailboxAddresses {
    <# Every mailbox of the organization (-RecipientTypeDetails: only these kinds, for example RoomMailbox). #>
    param([string[]]$RecipientTypeDetails)
    $command = Get-McoExchangeCommand 'Get-Mailbox'
    if (-not $command) { return [pscustomobject]@{ Available = $false; Addresses = @(); Error = 'Get-Mailbox is not available in this PowerShell session.' } }
    try {
        $mailboxArgs = @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        if ($RecipientTypeDetails) { $mailboxArgs.RecipientTypeDetails = $RecipientTypeDetails }
        $addresses = @(& $command.Name @mailboxArgs | ForEach-Object {
                if ($_.PrimarySmtpAddress) { ([string]$_.PrimarySmtpAddress).ToLowerInvariant() }
            } | Where-Object { $_ } | Select-Object -Unique)
        [pscustomobject]@{ Available = $true; Addresses = $addresses; Error = '' }
    }
    catch {
        [pscustomobject]@{ Available = $true; Addresses = @(); Error = $_.Exception.Message }
    }
}

$script:GroupMembers = @{}

function Get-McoExchangeGroupMembers {
    <# The members of a distribution group, nested groups expanded (cached per run: a group invited to many meetings is read once). #>
    param([Parameter(Mandatory)][string]$Identity, [int]$MaxDepth = 8)
    $cacheKey = '{0}|{1}' -f $Identity.ToLowerInvariant(), $MaxDepth
    if ($script:GroupMembers.ContainsKey($cacheKey)) { return $script:GroupMembers[$cacheKey] }
    $command = Get-McoExchangeCommand 'Get-DistributionGroupMember'
    if (-not $command) { return @() }
    # Only a group is expanded: an attendee that is a mailbox does not cost a failing call.
    if ((Get-McoRecipientKind -Address $Identity) -ne 'Group') { $script:GroupMembers[$cacheKey] = @(); return @() }
    $seenGroups = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $members = [Collections.Generic.List[object]]::new()
    $visit = {
        param([string]$Group, [int]$Depth)
        if ($Depth -gt $MaxDepth -or -not $seenGroups.Add($Group)) { return }
        try { $rows = @(& $command.Name -Identity $Group -ResultSize Unlimited -ErrorAction Stop) }
        catch { return }
        foreach ($row in $rows) {
            $isGroup = [string]$row.RecipientTypeDetails -match 'Group'
            $address = if ($row.PrimarySmtpAddress) { ([string]$row.PrimarySmtpAddress).ToLowerInvariant() } elseif ($row.WindowsEmailAddress) { ([string]$row.WindowsEmailAddress).ToLowerInvariant() } else { '' }
            if ($isGroup -and $address) { & $visit $address ($Depth + 1); continue }
            if ($address) { $members.Add([pscustomobject]@{ Address = $address; DisplayName = [string]$row.DisplayName }) }
        }
    }
    & $visit $Identity 0
    $list = @($members | Sort-Object Address -Unique)
    $script:GroupMembers[$cacheKey] = $list
    $list
}

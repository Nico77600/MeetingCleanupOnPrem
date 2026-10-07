<#
    Small EWS SOAP client used by Meeting Cleanup On-Prem.
    It deliberately uses HTTP EWS instead of the retired/optional EWS Managed API DLL.
#>

$script:EwsNs = @{
    soap = 'http://schemas.xmlsoap.org/soap/envelope/'
    m = 'http://schemas.microsoft.com/exchange/services/2006/messages'
    t = 'http://schemas.microsoft.com/exchange/services/2006/types'
    aut = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006'
}

function ConvertTo-McoXmlText {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ($null -eq $Text) { return '' }
    [Security.SecurityElement]::Escape($Text)
}

function Get-McoNodeText {
    param([AllowNull()][Xml.XmlNode]$Node, [Parameter(Mandatory)][string]$XPath, [Parameter(Mandatory)][Xml.XmlNamespaceManager]$Ns)
    if (-not $Node) { return $null }
    $hit = $Node.SelectSingleNode($XPath, $Ns)
    if ($hit) { return $hit.InnerText }
    $null
}

function New-McoHttpClient {
    <#
        Windows authentication: the credential given (-Credential, Connection.CredentialFile) or the account that runs
        the tool. Connection.WindowsPackage forces NTLM or Kerberos for the EWS address (Negotiate lets Windows choose):
        NTLM is needed when Kerberos to the EWS name fails on the server (alternate service account not deployed), and
        a credential given explicitly is needed where the session has no single sign-on (Credential Guard, batch logon).
    #>
    param([Parameter(Mandatory)][hashtable]$Settings, [pscredential]$Credential, [Uri]$RequestUri)
    $handler = [Net.Http.HttpClientHandler]::new()
    if ($Settings.Authentication -eq 'Windows') {
        $network = if ($Credential) { $Credential.GetNetworkCredential() } else { [Net.CredentialCache]::DefaultNetworkCredentials }
        $package = [string]$Settings.WindowsPackage
        if ($package -and $package -ne 'Negotiate' -and $RequestUri) {
            $cache = [Net.CredentialCache]::new()
            $cache.Add([Uri]$RequestUri.GetLeftPart([UriPartial]::Authority), $package, $network)
            $handler.Credentials = $cache
        }
        elseif ($Credential) { $handler.Credentials = $network; $handler.UseDefaultCredentials = $false }
        else { $handler.UseDefaultCredentials = $true }
    }
    $client = [Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds([int]$Settings.TimeoutSeconds)
    $client.DefaultRequestHeaders.Accept.ParseAdd('text/xml')
    $client.DefaultRequestHeaders.UserAgent.ParseAdd("MeetingCleanupOnPrem/$($script:ToolVersion)")
    $client
}

function Get-McoEwsRequestUri {
    <# The address the EWS requests go to: the URL, or Connection.EwsServer with the path of the URL. #>
    param([Parameter(Mandatory)][string]$Url, [string]$Server)
    if (-not $Server) { return [Uri]$Url }
    $b = [UriBuilder]::new([Uri]$Url); $b.Host = $Server; $b.Uri
}

function Resolve-McoAutodiscoverUrl {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][hashtable]$Settings, [Parameter(Mandatory)][Net.Http.HttpClient]$HttpClient)
    $domain = $Mailbox.Split('@')[-1]
    $uri = "https://autodiscover.$domain/autodiscover/autodiscover.xml"
    $xml = '<?xml version="1.0" encoding="utf-8"?><Autodiscover xmlns="{0}"><Request><EMailAddress>{1}</EMailAddress><AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema></Request></Autodiscover>' -f $script:EwsNs.aut, (ConvertTo-McoXmlText $Mailbox)
    $content = [Net.Http.StringContent]::new($xml, [Text.Encoding]::UTF8, 'text/xml')
    try {
        $response = $HttpClient.PostAsync($uri, $content).GetAwaiter().GetResult()
        $text = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $doc = [Xml.XmlDocument]::new(); $doc.LoadXml($text)
        $node = $doc.SelectSingleNode("//*[local-name()='EwsUrl']")
        if ($node -and $node.InnerText -match '^https://') { return $node.InnerText }
        throw "Autodiscover did not return an HTTPS EWS URL for $Mailbox."
    }
    catch { throw "Autodiscover for $Mailbox failed: $($_.Exception.Message)" }
    finally { $content.Dispose() }
}

function Get-McoSavedCredential {
    <# Connection.CredentialFile: a PSCredential saved with Export-Clixml by the account that runs the tool (DPAPI). #>
    param([Parameter(Mandatory)][hashtable]$Settings)
    if (-not [string]$Settings.CredentialFile) { return $null }
    $c = Import-Clixml -LiteralPath $Settings.CredentialFile
    if ($c -isnot [pscredential]) { throw "Connection.CredentialFile $($Settings.CredentialFile) does not hold a credential." }
    $c
}

function Connect-McoEws {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Settings, [pscredential]$Credential)
    if (-not $Credential) { $Credential = Get-McoSavedCredential -Settings $Settings }
    if ($Settings.Authentication -eq 'Basic' -and -not $Credential) {
        throw 'Basic authentication needs -Credential or Connection.CredentialUser with a password prompt.'
    }
    if ($Settings.AccessMode -eq 'Self' -and -not $Settings.Mailbox) {
        throw 'Connection.Mailbox is required when AccessMode is Self.'
    }
    $url = [string]$Settings.EwsUrl
    if ($Settings.Discovery -eq 'Autodiscover') {
        if (-not $Settings.Mailbox) { throw 'Connection.Mailbox is required for Autodiscover.' }
        $discovery = New-McoHttpClient -Settings $Settings -Credential $Credential
        try { $url = Resolve-McoAutodiscoverUrl -Mailbox $Settings.Mailbox -Settings $Settings -HttpClient $discovery } finally { $discovery.Dispose() }
    }
    $client = New-McoHttpClient -Settings $Settings -Credential $Credential -RequestUri (Get-McoEwsRequestUri -Url $url -Server ([string]$Settings.EwsServer))
    $script:Ews = @{
        Client = $client; Url = $url; Credential = $Credential
        Settings = $Settings; Mailbox = [string]$Settings.Mailbox; ConnectedUtc = [datetime]::UtcNow
    }
    $probe = Invoke-McoEws -Operation 'GetFolder' -Body (New-McoGetFolderBody -Distinguished 'calendar' -Mailbox $Settings.Mailbox)
    if ($probe.HttpStatus -ne 200 -or $probe.ResponseClass -ne 'Success') {
        Disconnect-McoEws
        throw "EWS could not open the calendar: $(Get-McoEwsFailureText $probe)"
    }
    [pscustomobject]@{ EwsUrl = $url; Mailbox = $Settings.Mailbox; Authentication = $Settings.Authentication; AccessMode = $Settings.AccessMode; ServerVersion = $probe.ServerVersion }
}

function Disconnect-McoEws {
    if ($script:Ews) {
        if ($script:Ews.Client) { $script:Ews.Client.Dispose() }
        $script:Ews = $null
    }
}

function New-McoFolderIdXml {
    param([Parameter(Mandatory)][string]$Distinguished, [string]$Mailbox, [ValidateSet('Self', 'Delegate', 'Impersonation')][string]$AccessMode)
    $mode = if ($AccessMode) { $AccessMode } elseif ($script:Ews) { [string]$script:Ews.Settings.AccessMode } else { '' }
    if ($mode -eq 'Self') { $Mailbox = '' }
    if ($Mailbox) { return '<t:DistinguishedFolderId Id="{0}"><t:Mailbox><t:EmailAddress>{1}</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId>' -f $Distinguished, (ConvertTo-McoXmlText $Mailbox) }
    '<t:DistinguishedFolderId Id="{0}"/>' -f $Distinguished
}

function New-McoSoapEnvelope {
    param([Parameter(Mandatory)][string]$Body, [string]$Mailbox)
    $header = '<t:RequestServerVersion Version="{0}"/>' -f (ConvertTo-McoXmlText ([string]$script:Ews.Settings.RequestServerVersion))
    if ($script:Ews.Settings.AccessMode -eq 'Impersonation' -and $Mailbox) {
        $header += '<t:ExchangeImpersonation><t:ConnectingSID><t:SmtpAddress>{0}</t:SmtpAddress></t:ConnectingSID></t:ExchangeImpersonation>' -f (ConvertTo-McoXmlText $Mailbox)
    }
    '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="{0}" xmlns:m="{1}" xmlns:t="{2}"><soap:Header>{3}</soap:Header><soap:Body>{4}</soap:Body></soap:Envelope>' -f $script:EwsNs.soap, $script:EwsNs.m, $script:EwsNs.t, $header, $Body
}

function ConvertFrom-McoEwsResponse {
    param([Parameter(Mandatory)][int]$HttpStatus, [Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    $doc = $null
    if ($Text.TrimStart().StartsWith('<')) { try { $doc = [Xml.XmlDocument]::new(); $doc.LoadXml($Text) } catch { $doc = $null } }
    $ns = $null
    if ($doc) {
        $ns = [Xml.XmlNamespaceManager]::new($doc.NameTable)
        foreach ($k in @('soap', 'm', 't')) { $ns.AddNamespace($k, $script:EwsNs[$k]) }
    }
    $fault = if ($doc) { $doc.SelectSingleNode('//soap:Fault', $ns) } else { $null }
    $messages = @(if ($doc) { $doc.SelectNodes('//m:ResponseMessages/*', $ns) })
    $first = $messages | Where-Object { $_.GetAttribute('ResponseClass') -ne 'Success' } | Select-Object -First 1
    if (-not $first) { $first = $messages | Select-Object -First 1 }
    $faultCode = if ($fault) { $fault.SelectSingleNode('detail/*[local-name()="ResponseCode"]') } else { $null }
    $faultMessage = if ($fault) { $fault.SelectSingleNode('detail/*[local-name()="Message"]') } else { $null }
    $responseCode = if ($first) { $first.SelectSingleNode('m:ResponseCode', $ns) } else { $null }
    $responseMessage = if ($first) { $first.SelectSingleNode('m:MessageText', $ns) } else { $null }
    $code = if ($faultCode) { $faultCode.InnerText } elseif ($responseCode) { $responseCode.InnerText } else { $null }
    $message = if ($faultMessage) { $faultMessage.InnerText } elseif ($responseMessage) { $responseMessage.InnerText } else { $null }
    [pscustomobject]@{
        HttpStatus = $HttpStatus; Xml = $doc; Ns = $ns; Fault = $fault
        ResponseClass = if ($fault) { 'Error' } elseif ($first) { $first.GetAttribute('ResponseClass') } elseif ($HttpStatus -eq 200) { 'Success' } else { 'Error' }
        ResponseCode = $code; MessageText = $message; ServerVersion = if ($doc) { $sv = $doc.SelectSingleNode('//t:ServerVersionInfo', $ns); if ($sv) { [string]$sv.GetAttribute('Version') } else { '' } } else { '' }
    }
}

function Send-McoEwsRequest {
    <# One EWS request, one answer (the transport of Invoke-McoEws; the tests replace it with a simulated server). #>
    param([Parameter(Mandatory)][string]$Operation, [Parameter(Mandatory)][string]$Body, [Parameter(Mandatory)][string]$Mailbox)
    $soap = New-McoSoapEnvelope -Body $Body -Mailbox $Mailbox
    # Connection.EwsServer: the request goes to that server, with the name of the URL kept for TLS and Host
    # (a load balancer cannot send a server back to itself: Azure internal load balancer, hairpin).
    $uri = [Uri]$script:Ews.Url
    $server = [string]$script:Ews.Settings.EwsServer
    $requestUri = Get-McoEwsRequestUri -Url $script:Ews.Url -Server $server
    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $requestUri)
    if ($server) { $request.Headers.Host = $uri.Authority }
    $request.Content = [Net.Http.StringContent]::new($soap, [Text.Encoding]::UTF8, 'text/xml')
    $request.Headers.Add('SOAPAction', "`"http://schemas.microsoft.com/exchange/services/2006/messages/$Operation`"")
    $request.Headers.Add('X-AnchorMailbox', $Mailbox)
    if ($script:Ews.Settings.Authentication -eq 'Basic' -and $script:Ews.Credential) {
        $pair = '{0}:{1}' -f $script:Ews.Credential.UserName, $script:Ews.Credential.GetNetworkCredential().Password
        $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Basic', [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair)))
    }
    try {
        $response = $script:Ews.Client.SendAsync($request).GetAwaiter().GetResult()
        $text = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        ConvertFrom-McoEwsResponse -HttpStatus ([int]$response.StatusCode) -Text $text
    }
    finally { $request.Dispose() }
}

function Invoke-McoEws {
    <#
        An EWS request, sent again up to Connection.MaxRetries times when Exchange did not process it: busy
        (ErrorServerBusy: throttling, after the delay it asks for, BackOffMilliseconds) or unavailable (HTTP 503).
        After a network error only a read (GetFolder, FindItem, GetItem) is sent again: a change may have been done.
    #>
    param([Parameter(Mandatory)][string]$Operation, [Parameter(Mandatory)][string]$Body, [string]$Mailbox)
    if (-not $script:Ews) { throw 'EWS is not connected.' }
    $target = if ($Mailbox) { $Mailbox } else { [string]$script:Ews.Mailbox }
    $retries = [Math]::Max(0, [int]$script:Ews.Settings.MaxRetries)
    $read = $Operation -in 'GetFolder', 'FindItem', 'GetItem'
    for ($attempt = 0; ; $attempt++) {
        $wait = [Math]::Min(60000, 2000 * [Math]::Pow(2, $attempt))
        try { $answer = Send-McoEwsRequest -Operation $Operation -Body $Body -Mailbox $target }
        catch {
            if (-not $read -or $attempt -ge $retries) { throw "EWS $Operation failed: $($_.Exception.Message)" }
            Write-McoLog 'WARN' ("EWS {0} {1}: {2} - sent again in {3:N0} s ({4}/{5})" -f $Operation, $target, $_.Exception.Message, ($wait / 1000), ($attempt + 1), $retries)
            Start-Sleep -Milliseconds ([int]$wait)
            continue
        }
        if (($answer.ResponseCode -ne 'ErrorServerBusy' -and $answer.HttpStatus -ne 503) -or $attempt -ge $retries) { return $answer }
        $backOff = if ($answer.Xml) { $answer.Xml.SelectSingleNode("//*[local-name()='Value'][@Name='BackOffMilliseconds']") } else { $null }
        $ms = 0
        if ($backOff -and [int]::TryParse($backOff.InnerText, [ref]$ms) -and $ms -gt 0) { $wait = [Math]::Min(300000, $ms) }
        Write-McoLog 'WARN' ("EWS {0} {1}: {2} - sent again in {3:N1} s ({4}/{5})" -f $Operation, $target, $(if ($answer.ResponseCode) { $answer.ResponseCode } else { "HTTP $($answer.HttpStatus)" }), ($wait / 1000), ($attempt + 1), $retries)
        Start-Sleep -Milliseconds ([int]$wait)
    }
}

function Get-McoEwsFailureText {
    param([Parameter(Mandatory)]$Answer)
    if ($Answer.MessageText) { return "$($Answer.ResponseCode): $($Answer.MessageText)" }
    "HTTP $($Answer.HttpStatus) $($Answer.ResponseCode)"
}

function New-McoGetFolderBody {
    param([ValidateSet('calendar', 'recoverableitemspurges', 'recoverableitemsdeletions')][string]$Distinguished = 'calendar', [string]$Mailbox)
    '<m:GetFolder><m:FolderShape><t:BaseShape>Default</t:BaseShape></m:FolderShape><m:FolderIds>{0}</m:FolderIds></m:GetFolder>' -f (New-McoFolderIdXml $Distinguished $Mailbox)
}

function Get-McoEwsCalendarProperties {
    @(
        'item:Subject', 'item:ItemClass', 'item:DateTimeCreated', 'item:LastModifiedTime', 'item:Body',
        'calendar:UID', 'calendar:Start', 'calendar:End', 'calendar:CalendarItemType', 'calendar:Location',
        'calendar:Organizer', 'calendar:MyResponseType', 'calendar:IsCancelled', 'calendar:RequiredAttendees',
        'calendar:OptionalAttendees', 'calendar:Resources', 'calendar:Recurrence', 'calendar:StartTimeZone'
    )
}

# CalendarView answers fast only with a few properties; UID and the item type let the caller group the
# occurrences of a series before one GetItem per meeting. Organizer and attendees come from GetItem only.
function Get-McoEwsCalendarViewProperties {
    param([switch]$Minimal)
    $base = @('item:Subject', 'item:ItemClass', 'calendar:Start', 'calendar:End', 'calendar:Location')
    if ($Minimal) { return $base }
    $base + @('calendar:CalendarItemType', 'calendar:UID')
}

function New-McoFindCalendarBody {
    param([Parameter(Mandatory)][datetime]$StartUtc, [Parameter(Mandatory)][datetime]$EndUtc, [string]$Mailbox, [int]$PageSize = 500, [string]$Folder = 'calendar', [switch]$Minimal)
    $props = (Get-McoEwsCalendarViewProperties -Minimal:$Minimal | ForEach-Object { '<t:FieldURI FieldURI="{0}"/>' -f $_ }) -join ''
    $view = '<m:CalendarView MaxEntriesReturned="{0}" StartDate="{1}" EndDate="{2}"/>' -f $PageSize, $StartUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'), $EndUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    '<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>{0}</t:AdditionalProperties></m:ItemShape>{1}<m:ParentFolderIds>{2}</m:ParentFolderIds></m:FindItem>' -f $props, $view, (New-McoFolderIdXml $Folder $Mailbox)
}

function New-McoGetItemBody {
    <# -Master: the series master of an occurrence (RecurringMasterItemId), else the item itself. #>
    param([Parameter(Mandatory)][string]$ItemId, [string]$Mailbox, [switch]$Master)
    $props = (Get-McoEwsCalendarProperties | ForEach-Object { '<t:FieldURI FieldURI="{0}"/>' -f $_ }) -join ''
    $id = if ($Master) { '<t:RecurringMasterItemId OccurrenceId="{0}"/>' -f (ConvertTo-McoXmlText $ItemId) } else { '<t:ItemId Id="{0}"/>' -f (ConvertTo-McoXmlText $ItemId) }
    '<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:BodyType>Text</t:BodyType><t:AdditionalProperties>{0}</t:AdditionalProperties></m:ItemShape><m:ItemIds>{1}</m:ItemIds></m:GetItem>' -f $props, $id
}

function New-McoDeleteItemBody {
    <# A calendar item needs SendMeetingCancellations: SendToNone keeps the removal silent. #>
    param([Parameter(Mandatory)][string]$ItemId, [ValidateSet('MoveToDeletedItems', 'SoftDelete', 'HardDelete')][string]$DeleteType = 'SoftDelete')
    '<m:DeleteItem DeleteType="{0}" SendMeetingCancellations="SendToNone"><m:ItemIds><t:ItemId Id="{1}"/></m:ItemIds></m:DeleteItem>' -f $DeleteType, (ConvertTo-McoXmlText $ItemId)
}

function New-McoCancelItemBody {
    param([Parameter(Mandatory)][string]$ItemId, [string]$Comment, [string]$ChangeKey)
    $ck = if ($ChangeKey) { ' ChangeKey="{0}"' -f (ConvertTo-McoXmlText $ChangeKey) } else { '' }
    '<m:CreateItem MessageDisposition="SendAndSaveCopy"><m:Items><t:CancelCalendarItem><t:ReferenceItemId Id="{0}"{1}/><t:NewBodyContent BodyType="Text">{2}</t:NewBodyContent></t:CancelCalendarItem></m:Items></m:CreateItem>' -f (ConvertTo-McoXmlText $ItemId), $ck, (ConvertTo-McoXmlText $Comment)
}

function New-McoAcceptItemBody {
    param([Parameter(Mandatory)][string]$ItemId)
    '<m:CreateItem MessageDisposition="SendOnly"><m:Items><t:AcceptItem><t:ReferenceItemId Id="{0}"/></t:AcceptItem></m:Items></m:CreateItem>' -f (ConvertTo-McoXmlText $ItemId)
}

function New-McoMoveItemBody {
    param([Parameter(Mandatory)][string]$ItemId, [Parameter(Mandatory)][string]$Mailbox)
    '<m:MoveItem><m:ToFolderId>{0}</m:ToFolderId><m:ItemIds><t:ItemId Id="{1}"/></m:ItemIds></m:MoveItem>' -f (New-McoFolderIdXml 'calendar' $Mailbox), (ConvertTo-McoXmlText $ItemId)
}

function ConvertTo-McoAttendeeXml {
    <# RequiredAttendees / OptionalAttendees / Resources: an empty list is left out (the schema refuses an empty array). #>
    param([Parameter(Mandatory)][string]$Element, [AllowNull()][AllowEmptyCollection()][string[]]$Address)
    $list = @($Address | Where-Object { $_ })
    if (-not $list.Count) { return '' }
    '<t:{0}>{1}</t:{0}>' -f $Element, (($list | ForEach-Object { '<t:Attendee><t:Mailbox><t:EmailAddress>{0}</t:EmailAddress></t:Mailbox></t:Attendee>' -f (ConvertTo-McoXmlText $_) }) -join '')
}

function New-McoCreateAppointmentBody {
    <#
        A meeting created and sent to its attendees and rooms. A series keeps the recurrence read from the source
        (RecurrenceXml) and its time zone (TimeZoneId); Start / End are then the first occurrence of the new series.
        Elements follow the order of CalendarItemType in the schema.
    #>
    param([Parameter(Mandatory)]$Event, [string]$Comment, [ValidateSet('SendToAllAndSaveCopy', 'SendToNone', 'SendOnlyToAll')][string]$Send = 'SendToAllAndSaveCopy')
    $recurrence = [string](Get-McoProperty $Event 'RecurrenceXml')
    $zone = [string](Get-McoProperty $Event 'TimeZoneId')
    $location = if ([string]$Event.Location) { '<t:Location>{0}</t:Location>' -f (ConvertTo-McoXmlText ([string]$Event.Location)) } else { '' }
    $tz = if ($recurrence -and $zone) { '<t:StartTimeZone Id="{0}"/><t:EndTimeZone Id="{0}"/>' -f (ConvertTo-McoXmlText $zone) } else { '' }
    ('<m:CreateItem SendMeetingInvitations="{0}"><m:SavedItemFolderId>{1}</m:SavedItemFolderId><m:Items><t:CalendarItem>' -f $Send, (New-McoFolderIdXml 'calendar')) +
    ('<t:Subject>{0}</t:Subject><t:Body BodyType="Text">{1}</t:Body>' -f (ConvertTo-McoXmlText ([string]$Event.Subject)), (ConvertTo-McoXmlText ([string]$Comment))) +
    ('<t:Start>{0}</t:Start><t:End>{1}</t:End>' -f ([datetime]$Event.Start).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'), ([datetime]$Event.End).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')) +
    $location +
    (ConvertTo-McoAttendeeXml 'RequiredAttendees' @(Get-McoProperty $Event 'RequiredAttendees')) +
    (ConvertTo-McoAttendeeXml 'OptionalAttendees' @(Get-McoProperty $Event 'OptionalAttendees')) +
    (ConvertTo-McoAttendeeXml 'Resources' @(Get-McoProperty $Event 'Resources')) +
    $recurrence + $tz +
    '</t:CalendarItem></m:Items></m:CreateItem>'
}

function Invoke-McoCreateAppointment {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)]$Event, [string]$Comment, [string]$Send = 'SendToAllAndSaveCopy')
    $answer = Invoke-McoEws -Operation 'CreateItem' -Body (New-McoCreateAppointmentBody -Event $Event -Comment $Comment -Send $Send) -Mailbox $Mailbox
    $node = if ($answer.Xml) { $answer.Xml.SelectSingleNode('//m:Items/t:CalendarItem/t:ItemId', $answer.Ns) } else { $null }
    $newId = if ($node) { $node.GetAttribute('Id') } else { '' }
    $created = if ($newId) { Get-McoItem -Mailbox $Mailbox -EventId $newId } else { $null }
    [pscustomobject]@{
        Ok = $answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success'
        EventId = $newId
        MeetingId = if ($created) { $created.MeetingId } else { [string]$Event.MeetingId }
        Error = if ($answer.ResponseClass -ne 'Success' -and $answer.ResponseCode) { "$($answer.ResponseCode): $($answer.MessageText)" } elseif ($answer.HttpStatus -ne 200) { "HTTP $($answer.HttpStatus)" } else { '' }
    }
}

function ConvertFrom-McoEwsDate {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if (-not $Text) { return $null }
    [datetime]::Parse($Text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
}

function ConvertFrom-McoCalendarNode {
    <# A calendar item of an EWS answer as the object of the tool (compiled: MeetingCleanupOnPremNative.Fast.CalendarNode). #>
    param([Parameter(Mandatory)][Xml.XmlNode]$Node, [Parameter(Mandatory)][Xml.XmlNamespaceManager]$Ns, [string]$Mailbox)
    [MeetingCleanupOnPremNative.Fast]::CalendarNode($Node, $Ns, $Mailbox)
}

# The CalendarView of each mailbox read during a search (Find-McoMeetings sets it, then clears it): an attendee
# mailbox shared by many meetings is read once, not once per meeting.
$script:CalendarViewCache = $null
# Calendars read only in part (more items starting at the same time than a page holds): set by Find-McoMeetings,
# the search adds them to its warnings.
$script:CalendarWarnings = $null
# GetItem: items read per call (a series master, an occurrence, an item ID mixed).
$script:GetItemBatchSize = 50

function Get-McoCalendarView {
    <#
        The items of a calendar in a period (CalendarView: few properties, every occurrence of a series), page after
        page: CalendarView has no offset, so while Exchange says the last item of the period is not in the answer
        (IncludesLastItemInRange="false"), the next page starts at the start of the last item received; the items
        seen twice (they overlap the new start) are kept once. A page that brings nothing later (more items starting
        at that time than Connection.PageSize) jumps past them, and the calendar is reported as read only in part.
        When the server refuses UID in CalendarView, each item is read to know it (slower, in batches).
    #>
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][datetime]$StartUtc, [Parameter(Mandatory)][datetime]$EndUtc, [int]$PageSize = 500)
    $key = '{0}|{1}|{2}|{3}' -f $Mailbox.ToLowerInvariant(), $StartUtc.ToUniversalTime().Ticks, $EndUtc.ToUniversalTime().Ticks, $PageSize
    if ($null -ne $script:CalendarViewCache -and $script:CalendarViewCache.ContainsKey($key)) { return , $script:CalendarViewCache[$key] }
    $summaries = [Collections.Generic.List[psobject]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $from = $StartUtc.ToUniversalTime(); $end = $EndUtc.ToUniversalTime()
    $minimal = $false
    $pages = 0
    while ($from -lt $end -and $pages -lt 5000) {
        $answer = Invoke-McoEws -Operation 'FindItem' -Body (New-McoFindCalendarBody $from $end $Mailbox $PageSize -Minimal:$minimal) -Mailbox $Mailbox
        if (-not $minimal -and $answer.HttpStatus -eq 200 -and $answer.ResponseCode -match '^ErrorInvalid(Request|PropertyRequest)$') { $minimal = $true; continue }
        if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') { throw "Mailbox ${Mailbox}: $(Get-McoEwsFailureText $answer)" }
        $pages++
        $page = [MeetingCleanupOnPremNative.Fast]::CalendarNodes($answer.Xml, '//m:RootFolder/t:Items/*', $answer.Ns, $Mailbox)
        foreach ($s in $page) { if ($seen.Add([string]$s.EventId)) { $summaries.Add($s) } }
        $root = $answer.Xml.SelectSingleNode('//m:RootFolder', $answer.Ns)
        if (-not $page.Count -or -not $root -or $root.GetAttribute('IncludesLastItemInRange') -ne 'false') { break }
        # The next page: from the start of the last item (CalendarView sorts by start).
        $next = $from
        foreach ($s in $page) { if ($s.Start -gt $next) { $next = $s.Start } }
        if ($next -le $from) {
            # Every item of the page starts at (or before) the start of the page: past the first of them to end.
            $next = $end
            foreach ($s in $page) { if ($s.End -gt $from -and $s.End -lt $next) { $next = $s.End } }
            $text = "Calendar of ${Mailbox}: more than $PageSize items at $(Format-McoDate $from $script:Ews.Settings.TimeZone), some of them may be missing (raise Connection.PageSize)."
            Write-McoItem Warn $text
            if ($null -ne $script:CalendarWarnings) { $script:CalendarWarnings.Add($text) }
            if ($next -le $from) { break }
        }
        $from = $next
    }
    if ($pages -gt 1) { Write-McoLog 'INFO' "Calendar of ${Mailbox}: $($summaries.Count) items read in $pages pages of $PageSize." }
    if ($minimal -and $summaries.Count) {
        $full = Get-McoItems -Mailbox $Mailbox -Requests @(foreach ($s in $summaries) { [pscustomobject]@{ EventId = $s.EventId; Master = $false } })
        for ($i = 0; $i -lt $summaries.Count; $i++) {
            if ($full[$i]) { $summaries[$i].MeetingId = $full[$i].MeetingId; $summaries[$i].AppointmentType = $full[$i].AppointmentType }
        }
    }
    if ($null -ne $script:CalendarViewCache) { $script:CalendarViewCache[$key] = $summaries }
    , $summaries
}

function Get-McoMailboxEvents {
    <#
        The calendar items of a mailbox in a period. CalendarView gives every occurrence of a series; they are
        grouped by UID and read with GetItem (the series master for an occurrence, 50 items per call), so a series
        is ONE item with its occurrences in the period (Occurrences). -Occurrences: one item per occurrence instead
        (rooms mode, a series is acted on in the period only). -Uid: only these meetings (attendee lookup).
    #>
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][datetime]$StartUtc, [Parameter(Mandatory)][datetime]$EndUtc, [int]$PageSize = 500, [switch]$Occurrences, [string[]]$Uid)
    $summaries = Get-McoCalendarView -Mailbox $Mailbox -StartUtc $StartUtc -EndUtc $EndUtc -PageSize $PageSize
    $wanted = $null
    if ($Uid) {
        $wanted = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($u in $Uid) { if ($u) { [void]$wanted.Add(([string]$u).ToUpperInvariant()) } }
        if (-not $wanted.Count) { $wanted = $null }
    }
    # Grouped by UID, in the order of the calendar (Group-Object without the pipeline).
    $groups = [ordered]@{}
    foreach ($s in $summaries) {
        if ($wanted -and -not $wanted.Contains($s.MeetingId)) { continue }
        $list = $groups[$s.MeetingId]
        if (-not $list) { $list = [Collections.Generic.List[object]]::new(); $groups[$s.MeetingId] = $list }
        $list.Add($s)
    }
    $requests = [Collections.Generic.List[object]]::new()
    foreach ($list in $groups.Values) { $requests.Add([pscustomobject]@{ EventId = $list[0].EventId; Master = ($list[0].AppointmentType -in 'Occurrence', 'Exception') }) }
    $details = if ($requests.Count) { Get-McoItems -Mailbox $Mailbox -Requests $requests.ToArray() } else { @() }
    $items = [Collections.Generic.List[object]]::new()
    $k = 0
    foreach ($list in $groups.Values) {
        $first = $list[0]
        $isSeries = $first.AppointmentType -in 'Occurrence', 'Exception'
        $detail = $details[$k]; $k++
        # Not read: the summary of CalendarView stands for it (a copy: the summaries are kept for the search).
        if (-not $detail) { $detail = $first.PSObject.Copy() }
        $occ = [Collections.Generic.List[object]]::new()
        foreach ($s in $list) { if ($s.AppointmentType -in 'Occurrence', 'Exception') { $occ.Add([pscustomobject]@{ EventId = $s.EventId; Start = $s.Start; End = $s.End; Type = $s.AppointmentType }) } }
        $occ = @($occ | Sort-Object Start)
        if ($isSeries -and $Occurrences) {
            foreach ($o in $occ) {
                $copy = $detail.PSObject.Copy()
                $copy.SeriesId = $detail.EventId; $copy.EventId = $o.EventId; $copy.Start = $o.Start; $copy.End = $o.End; $copy.AppointmentType = $o.Type; $copy.Occurrences = @($o)
                $items.Add($copy)
            }
        }
        else {
            $detail.Occurrences = $occ
            $items.Add($detail)
        }
    }
    $items.ToArray()
}

function Get-McoItems {
    <#
        Several items of one mailbox read with as few GetItem calls as possible ($script:GetItemBatchSize per call).
        Requests: objects with EventId and Master (the series master of an occurrence). The answer has one entry
        per request, in the same order: the item, or $null when it could not be read. A batch that EWS refuses as
        a whole is read again item by item.
    #>
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Requests)
    $results = [object[]]::new($Requests.Count)
    $size = [Math]::Max(1, [int]$script:GetItemBatchSize)
    $props = (Get-McoEwsCalendarProperties | ForEach-Object { '<t:FieldURI FieldURI="{0}"/>' -f $_ }) -join ''
    for ($i = 0; $i -lt $Requests.Count; $i += $size) {
        $n = [Math]::Min($size, $Requests.Count - $i)
        $ids = [Text.StringBuilder]::new()
        for ($j = 0; $j -lt $n; $j++) {
            $r = $Requests[$i + $j]
            $format = if ($r.Master) { '<t:RecurringMasterItemId OccurrenceId="{0}"/>' } else { '<t:ItemId Id="{0}"/>' }
            [void]$ids.Append(($format -f (ConvertTo-McoXmlText ([string]$r.EventId))))
        }
        $body = '<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:BodyType>Text</t:BodyType><t:AdditionalProperties>{0}</t:AdditionalProperties></m:ItemShape><m:ItemIds>{1}</m:ItemIds></m:GetItem>' -f $props, $ids.ToString()
        $answer = Invoke-McoEws -Operation 'GetItem' -Body $body -Mailbox $Mailbox
        $messages = @(if ($answer.HttpStatus -eq 200 -and $answer.Xml -and -not $answer.Fault) { $answer.Xml.SelectNodes('//m:ResponseMessages/*', $answer.Ns) })
        if ($messages.Count -ne $n) {
            if ($n -eq 1) { continue }
            for ($j = 0; $j -lt $n; $j++) { $results[$i + $j] = Get-McoItem -Mailbox $Mailbox -EventId ([string]$Requests[$i + $j].EventId) -Master:([bool]$Requests[$i + $j].Master) }
            continue
        }
        for ($j = 0; $j -lt $n; $j++) {
            $message = $messages[$j]
            if ($message.GetAttribute('ResponseClass') -ne 'Success') { continue }
            $node = $message.SelectSingleNode('m:Items/*', $answer.Ns)
            if ($node) { $results[$i + $j] = [MeetingCleanupOnPremNative.Fast]::CalendarNode($node, $answer.Ns, $Mailbox) }
        }
    }
    , $results
}

function Get-McoItem {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$EventId, [switch]$Master)
    $answer = Invoke-McoEws -Operation 'GetItem' -Body (New-McoGetItemBody -ItemId $EventId -Mailbox $Mailbox -Master:$Master) -Mailbox $Mailbox
    if ($answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success') {
        $node = $answer.Xml.SelectSingleNode('//m:Items/*', $answer.Ns)
        if ($node) { return ConvertFrom-McoCalendarNode $node $answer.Ns $Mailbox }
    }
    $null
}

function Test-McoItemExists {
    <# $true / $false, or $null when EWS did not answer clearly (the verification then reports Unknown). #>
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$EventId)
    $answer = Invoke-McoEws -Operation 'GetItem' -Body ('<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties><t:FieldURI FieldURI="item:ParentFolderId"/><t:FieldURI FieldURI="calendar:IsCancelled"/></t:AdditionalProperties></m:ItemShape><m:ItemIds><t:ItemId Id="{0}"/></m:ItemIds></m:GetItem>' -f (ConvertTo-McoXmlText $EventId)) -Mailbox $Mailbox
    if ($answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success') { return $true }
    if ($answer.ResponseCode -eq 'ErrorItemNotFound') { return $false }
    $null
}

function Invoke-McoDeleteItem {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$EventId, [ValidateSet('SoftDelete', 'HardDelete', 'MoveToDeletedItems')][string]$DeleteType = 'SoftDelete')
    Invoke-McoEws -Operation 'DeleteItem' -Body (New-McoDeleteItemBody -ItemId $EventId -DeleteType $DeleteType) -Mailbox $Mailbox
}

function Invoke-McoCancelItem {
    <# The organizer cancels the meeting (message to the attendees). The ChangeKey is read again just before. #>
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$EventId, [string]$Comment)
    $current = Get-McoItem -Mailbox $Mailbox -EventId $EventId
    $changeKey = if ($current) { [string]$current.ChangeKey } else { '' }
    Invoke-McoEws -Operation 'CreateItem' -Body (New-McoCancelItemBody -ItemId $EventId -Comment $Comment -ChangeKey $changeKey) -Mailbox $Mailbox
}

function ConvertTo-McoRecoverableTime {
    <# LastModifiedTime of Get-RecoverableItems: a DateTime, or text MM/dd/yyyy HH:mm:ss in UTC. #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) } return $Value.ToUniversalTime() }
    $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact([string]$Value, [string[]]@('MM/dd/yyyy HH:mm:ss', 'M/d/yyyy h:mm:ss tt', 'o', 'yyyy-MM-ddTHH:mm:ss'), [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$d)) { return $d }
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$d)) { return $d }
    $null
}

function Get-McoPurgedItems {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][datetime]$StartUtc, [Parameter(Mandatory)][datetime]$EndUtc, [ValidateSet('Auto', 'ExchangePowerShell', 'Ews')][string]$Mode = 'Auto')
    $getRecoverable = Get-Command Get-RecoverableItems -ErrorAction SilentlyContinue
    if ($Mode -eq 'ExchangePowerShell' -and -not $getRecoverable) { throw 'Restore.Mode is ExchangePowerShell, but Get-RecoverableItems is not available. Run from Exchange Management Shell or import the on-premises Exchange session.' }
    if ($Mode -ne 'Ews' -and $getRecoverable) {
        # No -SourceFolder: Deleted Items, Recoverable Items\Deletions (SoftDelete) and Purges (HardDelete) are searched.
        # The filter times are read by the Exchange server in its own time zone: the window is widened by a day and
        # the exact window is applied on LastModifiedTime below.
        $items = @(Get-RecoverableItems -Identity $Mailbox -FilterItemType IPM.Appointment -ResultSize Unlimited -FilterStartTime $StartUtc.AddDays(-1) -FilterEndTime $EndUtc.AddDays(1) -ErrorAction Stop -WarningAction SilentlyContinue)
        foreach ($item in $items) {
            $modified = ConvertTo-McoRecoverableTime $item.LastModifiedTime
            if ($modified -and ($modified -lt $StartUtc -or $modified -gt $EndUtc)) { continue }
            [pscustomobject]@{
                Mailbox = $Mailbox; Subject = [string]$item.Subject; EntryId = [string]$item.EntryID
                LastModifiedUtc = $modified
                Folder = [string]$(if ($item.PSObject.Properties['LastParentPath']) { $item.LastParentPath } else { 'Recoverable Items' })
            }
        }
        return
    }
    foreach ($folder in 'recoverableitemspurges', 'recoverableitemsdeletions') {
        $props = @('item:Subject', 'item:LastModifiedTime', 'calendar:UID', 'item:ItemClass')
        $fields = ($props | ForEach-Object { '<t:FieldURI FieldURI="{0}"/>' -f $_ }) -join ''
        # IndexedPageItemView: 1,000 items per page, the next page from the offset Exchange gives back.
        $offset = 0
        for ($page = 0; $page -lt 200; $page++) {
            $body = '<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>{0}</t:AdditionalProperties></m:ItemShape><m:IndexedPageItemView MaxEntriesReturned="1000" Offset="{1}" BasePoint="Beginning"/><m:ParentFolderIds>{2}</m:ParentFolderIds></m:FindItem>' -f $fields, $offset, (New-McoFolderIdXml $folder $Mailbox)
            $answer = Invoke-McoEws -Operation 'FindItem' -Body $body -Mailbox $Mailbox
            if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') { break }
            $items = [MeetingCleanupOnPremNative.Fast]::CalendarNodes($answer.Xml, '//m:RootFolder/t:Items/*', $answer.Ns, $Mailbox)
            foreach ($item in $items) {
                if ($item.LastModifiedUtc -and $item.LastModifiedUtc -ge $StartUtc -and $item.LastModifiedUtc -le $EndUtc) {
                    [pscustomobject]@{ Mailbox = $Mailbox; Subject = $item.Subject; EntryId = $item.EventId; LastModifiedUtc = $item.LastModifiedUtc; Folder = $folder }
                }
            }
            $root = $answer.Xml.SelectSingleNode('//m:RootFolder', $answer.Ns)
            if (-not $items.Count -or -not $root -or $root.GetAttribute('IncludesLastItemInRange') -ne 'false') { break }
            $nextOffset = 0
            if (-not [int]::TryParse($root.GetAttribute('IndexedPagingOffset'), [ref]$nextOffset) -or $nextOffset -le $offset) { $nextOffset = $offset + $items.Count }
            $offset = $nextOffset
        }
    }
}

function Invoke-McoRestoreItem {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$EntryId, [ValidateSet('Auto', 'ExchangePowerShell', 'Ews')][string]$Mode = 'Auto')
    $restoreRecoverable = Get-Command Restore-RecoverableItems -ErrorAction SilentlyContinue
    if ($Mode -eq 'ExchangePowerShell' -and -not $restoreRecoverable) { throw 'Restore.Mode is ExchangePowerShell, but Restore-RecoverableItems is not available. Run from Exchange Management Shell or import the on-premises Exchange session.' }
    if ($Mode -ne 'Ews' -and $restoreRecoverable) {
        try {
            $r = @(Restore-RecoverableItems -Identity $Mailbox -EntryID $EntryId -ErrorAction Stop -WarningAction SilentlyContinue)[0]
            $ok = if ($r -and $r.PSObject.Properties['WasRestoredSuccessfully']) { [bool]$r.WasRestoredSuccessfully } else { $true }
            $folder = if ($r -and $r.PSObject.Properties['RestoredToFolderPath']) { [string]$r.RestoredToFolderPath } else { '' }
            return [pscustomobject]@{ Ok = $ok; Error = $(if ($ok) { '' } else { 'Restore-RecoverableItems did not restore the item' }); Folder = $folder }
        }
        catch { return [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message; Folder = '' } }
    }
    $answer = Invoke-McoEws -Operation 'MoveItem' -Body (New-McoMoveItemBody $EntryId $Mailbox) -Mailbox $Mailbox
    [pscustomobject]@{ Ok = $answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success'; Error = if ($answer.ResponseCode) { "$($answer.ResponseCode): $($answer.MessageText)" } else { '' }; Folder = 'Calendar' }
}

function Invoke-McoAcceptItem {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$EventId)
    Invoke-McoEws -Operation 'CreateItem' -Body (New-McoAcceptItemBody $EventId) -Mailbox $Mailbox
}

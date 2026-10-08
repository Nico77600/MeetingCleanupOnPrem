#
#  Meeting Cleanup On-Prem - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 1.1.0
#
#  Read by Invoke-MeetingCleanupOnPrem.ps1. It is a PowerShell data file: text between quotes, $true / $false,
#  numbers, @( ) for lists and @{ } for groups of settings. Lines starting with # are comments. Relative paths
#  (.\reports, .\logs) are relative to the tool folder. Every value is checked at start; all the problems are
#  listed at once. The parameters of the command line override it for one run.
#
#  No password here: the tool signs in as the account that runs it, asks for the password (CredentialUser), or
#  reads a credential file encrypted for the account that runs it (CredentialFile, Export-Clixml).
#
@{
    # ---------------------------------------------------------------------
    # EWS (developer guide, chapter 5). The service account opens the calendars with ApplicationImpersonation
    # (limited by a management scope), or Full Access with Delegate.
    # ---------------------------------------------------------------------
    Connection = @{
        EwsUrl               = 'https://mail.contoso.test/EWS/Exchange.asmx'   # the EWS URL of the organization
        Discovery            = 'Manual'          # Manual | Autodiscover (from Mailbox)
        Mailbox              = 'svc-meetingcleanup@contoso.test'   # the address of the service account
        AccessMode           = 'Impersonation'   # Impersonation | Delegate | Self
        Authentication       = 'Windows'         # Windows | Basic (where the EWS virtual directory still accepts it)
        WindowsPackage       = 'Negotiate'       # Negotiate | NTLM (when Kerberos to the name of the URL fails) | Kerberos
        CredentialUser       = ''                # an account whose password is asked at each run (empty: the account that runs the tool)
        CredentialFile       = ''                # scheduled task: Get-Credential | Export-Clixml <file>, as the account of the task
        RequestServerVersion = 'Exchange2016'    # Exchange2013_SP1 | Exchange2016 | Exchange2019
        MaxRetries           = 3                 # a request Exchange did not process (busy, unavailable) is sent again
        TimeoutSeconds       = 120               # timeout of one request
        PageSize             = 500               # items of a calendar page (10-1000); a calendar is read page after page
        EwsServer            = ''                # optional: one server (name or IP) to send the requests to, the name of the URL kept (load balancer)
    }

    # ---------------------------------------------------------------------
    # Search. The rooms come from Exchange PowerShell (every room mailbox, AllRooms), plus Rooms and RoomFile. A room outside
    # the scope of the impersonation is not read (warning): set AllRooms = $false and list the rooms concerned.
    # A series is kept when one of its occurrences falls in the period. SeriesScope: 'Whole' acts on the whole
    # series (every occurrence, past ones included); 'Occurrences' only on its occurrences in the period
    # (-SeriesScope). Rooms mode always acts on the occurrences of the period.
    # ---------------------------------------------------------------------
    Search = @{
        SearchIn        = @('Organizer', 'Rooms')   # Organizer | Rooms | Mailboxes | AllMailboxes
        PastDays        = 0                          # default period: today minus PastDays ...
        FutureDays      = 365                        # ... to today plus FutureDays
        SeriesScope     = 'Whole'                    # Whole | Occurrences
        Rooms           = @()                        # rooms added to those of Exchange PowerShell
        RoomFile        = ''                         # a file of rooms (one address per line, or CSV)
        Mailboxes       = @()                        # the mailboxes of -SearchIn Mailboxes
        MailboxFile     = ''                         # a file of mailboxes
        AcceptedDomains = @('contoso.test')          # the domains of the organization: another one is external (listed, not processed)
        DirectoryMode   = 'Auto'                     # Auto | ExchangePowerShell | None (never the Exchange cmdlets)
        AllRooms        = $true                      # $true: every room mailbox of Exchange PowerShell, with Rooms and RoomFile; $false: Rooms and RoomFile only
    }

    # ---------------------------------------------------------------------
    # Exchange PowerShell: aliases and X500 addresses, every room and mailbox, the groups, Restore.
    # ---------------------------------------------------------------------
    ManagementShell = @{
        Mode           = 'Rps'                 # Rps (remote PowerShell) | Auto (the cmdlets of this session, else Rps) | Existing
        ServerFqdn     = 'mail.contoso.test'   # a Mailbox server (its FQDN for Kerberos)
        ConnectionUri  = ''                    # empty = http://<ServerFqdn>/PowerShell/ with Kerberos
        Authentication = 'Kerberos'            # Kerberos | Negotiate | Basic
        CredentialUser = ''                    # another account for remote PowerShell (its password is asked)
    }

    Cleanup = @{
        CancelComment = 'This meeting has been cancelled by the IT department.'   # message of -Action Cancel
        Verify        = $true                                                      # read each removed copy again
    }

    # ---------------------------------------------------------------------
    # Restore (-Action Restore -FromReport <report of a Remove run>): Get-RecoverableItems and
    # Restore-RecoverableItems, role Mailbox Import Export (developer guide, chapter 5).
    # ---------------------------------------------------------------------
    Restore = @{
        Mode          = 'Auto'   # Auto (Exchange PowerShell when present, else EWS) | ExchangePowerShell | Ews
        WindowMinutes = 10       # tolerance around the time of each removal
    }

    Transfer = @{
        Method  = 'Recreate'                                 # Exchange Server: the meeting is re-created by the new organizer
        Comment = 'This meeting is now organized by {0}.'    # message of the old organizer ({0} = the new organizer)
    }

    Report = @{
        OutputPath   = '.\reports'
        FilePrefix   = 'MeetingCleanupOnPrem'
        Formats      = @('Csv', 'Html')   # a Summary.json is always written
        CsvDelimiter = ';'
        TimeZone     = ''                 # time zone of the dates typed and shown (empty: the one of Windows)
    }

    Logging = @{
        Path          = '.\logs'
        RetentionDays = 30
    }
}

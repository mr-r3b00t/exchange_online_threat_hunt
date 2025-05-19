# Function to check if connected to Exchange Online
function Test-ExchangeOnlineConnection {
    try {
        # Attempt a simple Exchange cmdlet to verify connection
        Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Check if already connected to Exchange Online
$IsConnected = Test-ExchangeOnlineConnection

if ($IsConnected) {
    Write-Host "An active Exchange Online session is detected."
    $Response = Read-Host "Do you want to disconnect the existing session and reconnect? (Y/N)"
    if ($Response -eq 'Y' -or $Response -eq 'y') {
        # Disconnect existing sessions
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from existing Exchange Online session."
        # Reconnect
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false
    } else {
        Write-Host "Continuing with the existing session."
    }
} else {
    # Connect to Exchange Online
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited

# Array to store results
$Results = @()

# Loop through each mailbox
foreach ($Mailbox in $Mailboxes) {
    # Get Full Access permissions
    $FullAccessPerms = Get-MailboxPermission -Identity $Mailbox.UserPrincipalName | 
        Where-Object { $_.AccessRights -contains "FullAccess" -and $_.User -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false } |
        Select-Object @{Name="Delegate";Expression={$_.User}}, @{Name="Permission";Expression={"FullAccess"}}

    # Get Send As permissions
    $SendAsPerms = Get-RecipientPermission -Identity $Mailbox.UserPrincipalName |
        Where-Object { $_.AccessRights -contains "SendAs" -and $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false } |
        Select-Object @{Name="Delegate";Expression={$_.Trustee}}, @{Name="Permission";Expression={"SendAs"}}

    # Get Send on Behalf permissions
    $SendOnBehalfPerms = Get-Mailbox -Identity $Mailbox.UserPrincipalName | 
        Where-Object { $_.GrantSendOnBehalfTo } |
        Select-Object @{Name="Delegate";Expression={$_.GrantSendOnBehalfTo -join ", "}}, @{Name="Permission";Expression={"SendOnBehalf"}}

    # Combine permissions for the mailbox
    if ($FullAccessPerms -or $SendAsPerms -or $SendOnBehalfPerms) {
        $FullAccessDelegates = if ($FullAccessPerms) { ($FullAccessPerms.Delegate -join ", ") } else { "None" }
        $SendAsDelegates = if ($SendAsPerms) { ($SendAsPerms.Delegate -join ", ") } else { "None" }
        $SendOnBehalfDelegates = if ($SendOnBehalfPerms) { $SendOnBehalfPerms.Delegate } else { "None" }

        $Results += [PSCustomObject]@{
            Mailbox              = $Mailbox.UserPrincipalName
            FullAccessDelegates  = $FullAccessDelegates
            SendAsDelegates      = $SendAsDelegates
            SendOnBehalfDelegates = $SendOnBehalfDelegates
        }
    }
}

# Display results
if ($Results) {
    Write-Host "Mailboxes with delegation permissions:"
    $Results | Format-Table Mailbox, FullAccessDelegates, SendAsDelegates, SendOnBehalfDelegates -AutoSize
} else {
    Write-Host "No mailboxes with delegation permissions found."
}

# Optionally, export to CSV
$Results | Export-Csv -Path "MailboxDelegationPermissions.csv" -NoTypeInformation
Write-Host "Results exported to MailboxDelegationPermissions.csv"

# Prompt to disconnect from Exchange Online
$DisconnectResponse = Read-Host "Do you want to disconnect from Exchange Online? (Y/N)"
if ($DisconnectResponse -eq 'Y' -or $DisconnectResponse -eq 'y') {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected from Exchange Online."
} else {
    Write-Host "Remaining connected to Exchange Online."
}

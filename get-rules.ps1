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
    # Get inbox rules for the mailbox
    $Rules = Get-InboxRule -Mailbox $Mailbox.UserPrincipalName -ErrorAction SilentlyContinue
    
    # Check each rule for forwarding
    foreach ($Rule in $Rules) {
        if ($Rule.ForwardTo -or $Rule.ForwardAsAttachmentTo) {
            # Get forwarding addresses
            $ForwardAddresses = @($Rule.ForwardTo + $Rule.ForwardAsAttachmentTo)
            foreach ($Address in $ForwardAddresses) {
                # Skip if address is empty
                if ($Address) {
                    # Check if the address is external (does not contain the mailbox's domain)
                    $IsExternal = $Address -notlike "*@$($Mailbox.PrimarySmtpAddress.Split('@')[1])"
                    $Results += [PSCustomObject]@{
                        Mailbox            = $Mailbox.UserPrincipalName
                        RuleName           = $Rule.Name
                        RuleDescription    = $Rule.Description
                        ForwardTo          = $Address
                        IsExternal         = $IsExternal
                    }
                }
            }
        }
    }
}

# Filter for external forwarding rules
$ExternalResults = $Results | Where-Object { $_.IsExternal }

# Display results
if ($ExternalResults) {
    Write-Host "Mailboxes with external forwarding rules:"
    $ExternalResults | Format-Table Mailbox, RuleName, RuleDescription, ForwardTo -AutoSize
} else {
    Write-Host "No mailboxes with external forwarding rules found."
}

# Optionally, export to CSV
$ExternalResults | Export-Csv -Path "ExternalForwardingRules.csv" -NoTypeInformation
Write-Host "Results exported to ExternalForwardingRules.csv"

# Prompt to disconnect from Exchange Online
$DisconnectResponse = Read-Host "Do you want to disconnect from Exchange Online? (Y/N)"
if ($DisconnectResponse -eq 'Y' -or $DisconnectResponse -eq 'y') {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected from Exchange Online."
} else {
    Write-Host "Remaining connected to Exchange Online."
}

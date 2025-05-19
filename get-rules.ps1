# Check if already connected to Exchange Online
$ExistingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }

if ($ExistingSession) {
    Write-Host "An active Exchange Online session is detected."
    $Response = Read-Host "Do ''

System: you want to disconnect the existing session and reconnect? (Y/N)"
    if ($Response -eq 'Y' -or $Response -eq 'y') {
        # Disconnect existing sessions
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from existing Exchange Online session."
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
    
    # Check each rule for external forwarding
    foreach ($Rule in $Rules) {
        if ($Rule.ForwardTo -or $Rule.ForwardAsAttachmentTo -or $Rule.RedirectTo) {
            # Check if the rule forwards to an external email
            $ForwardAddresses = @($Rule.ForwardTo + $Rule.ForwardAsAttachmentTo + $Rule.RedirectTo)
            foreach ($Address in $ForwardAddresses) {
                # Skip if address is empty or internal (contains domain)
                if ($Address -and $Address -notlike "*@$($Mailbox.PrimarySmtpAddress.Split('@')[1])") {
                    $Results += [PSCustomObject]@{
                        Mailbox            = $Mailbox.UserPrincipalName
                        RuleName           = $Rule.Name
                        RuleDescription    = $Rule.Description
                        ForwardTo          = $Address
                    }
                }
            }
        }
    }
}

# Display results
$Results | Format-Table -AutoSize

# Optionally, export to CSV
$Results | Export-Csv -Path "ExternalForwardingRules.csv" -NoTypeInformation

# Prompt to disconnect from Exchange Online
$DisconnectResponse = Read-Host "Do you want to disconnect from Exchange Online? (Y/N)"
if ($DisconnectResponse -eq 'Y' -or $DisconnectResponse -eq 'y') {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected from Exchange Online."
} else {
    Write-Host "Remaining connected to Exchange Online."
}

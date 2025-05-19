# Get current timestamp for logging
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Display script start message
Write-Host "===== Mailbox Delegation Permissions Checker =====" -ForegroundColor Cyan
Write-Host "Script started at $Timestamp" -ForegroundColor Green
Write-Host "This script checks all mailboxes for Full Access, Send As, and Send on Behalf delegation permissions."
Write-Host "Verbose mode enabled: Detailed output will be displayed for each step."
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Function to check if connected to Exchange Online
function Test-ExchangeOnlineConnection {
    Write-Host "Checking for active Exchange Online session..." -ForegroundColor Yellow
    try {
        Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
        Write-Host "SUCCESS: Active Exchange Online session detected." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "INFO: No active Exchange Online session found. Error: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# Check if already connected to Exchange Online
Write-Host "===== Step 1: Verifying Exchange Online Connection =====" -ForegroundColor Cyan
$IsConnected = Test-ExchangeOnlineConnection

if ($IsConnected) {
    Write-Host "An active Exchange Online session is already established."
    Write-Host "Prompting user to decide whether to disconnect and reconnect..." -ForegroundColor Yellow
    $Response = Read-Host "Do you want to disconnect the existing session and reconnect? (Y/N)"
    if ($Response -eq 'Y' -or $Response -eq 'y') {
        Write-Host "User chose to disconnect. Disconnecting from Exchange Online..." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "SUCCESS: Disconnected from existing Exchange Online session at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
        Write-Host "Reconnecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "SUCCESS: Reconnected to Exchange Online at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
    } else {
        Write-Host "User chose to continue with the existing session." -ForegroundColor Green
    }
} else {
    Write-Host "No active session found. Initiating connection to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "SUCCESS: Connected to Exchange Online at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
}
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Get all mailboxes
Write-Host "===== Step 2: Retrieving Mailboxes =====" -ForegroundColor Cyan
Write-Host "Fetching all mailboxes from Exchange Online..." -ForegroundColor Yellow
$Mailboxes = Get-Mailbox -ResultSize Unlimited
$MailboxCount = $Mailboxes.Count
Write-Host "SUCCESS: Retrieved $MailboxCount mailboxes at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Array to store results
$Results = @()
$CurrentMailbox = 0

# Loop through each mailbox with progress bar
Write-Host "===== Step 3: Processing Mailboxes for Delegation Permissions =====" -ForegroundColor Cyan
Write-Host "Starting to process $MailboxCount mailboxes to check for delegation permissions."
Write-Host "A progress bar will display the current mailbox being processed and the overall progress."
Write-Host "Processing may take time depending on the number of mailboxes." -ForegroundColor Yellow
Write-Host ""

foreach ($Mailbox in $Mailboxes) {
    $CurrentMailbox++
    $PercentComplete = [math]::Round(($CurrentMailbox / $MailboxCount) * 100, 2)
    $CurrentMailboxName = $Mailbox.UserPrincipalName
    Write-Progress -Activity "Processing Mailboxes for Delegation Permissions" `
                  -Status "Processing mailbox: $CurrentMailboxName ($CurrentMailbox of $MailboxCount)" `
                  -PercentComplete $PercentComplete

    Write-Host "Processing mailbox: $CurrentMailboxName ($CurrentMailbox of $MailboxCount) at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow

    # Get Full Access permissions
    Write-Host "  Checking Full Access permissions for $CurrentMailboxName..." -ForegroundColor White
    $FullAccessPerms = Get-MailboxPermission -Identity $Mailbox.UserPrincipalName | 
        Where-Object { $_.AccessRights -contains "FullAccess" -and $_.User -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false } |
        Select-Object @{Name="Delegate";Expression={$_.User}}, @{Name="Permission";Expression={"FullAccess"}}
    if ($FullAccessPerms) {
        Write-Host "    FOUND: Full Access delegates: $($FullAccessPerms.Delegate -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "    NONE: No Full Access delegates found." -ForegroundColor Gray
    }

    # Get Send As permissions
    Write-Host "  Checking Send As permissions for $CurrentMailboxName..." -ForegroundColor White
    $SendAsPerms = Get-RecipientPermission -Identity $Mailbox.UserPrincipalName |
        Where-Object { $_.AccessRights -contains "SendAs" -and $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false } |
        Select-Object @{Name="Delegate";Expression={$_.Trustee}}, @{Name="Permission";Expression={"SendAs"}}
    if ($SendAsPerms) {
        Write-Host "    FOUND: Send As delegates: $($SendAsPerms.Delegate -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "    NONE: No Send As delegates found." -ForegroundColor Gray
    }

    # Get Send on Behalf permissions
    Write-Host "  Checking Send on Behalf permissions for $CurrentMailboxName..." -ForegroundColor White
    $SendOnBehalfPerms = Get-Mailbox -Identity $Mailbox.UserPrincipalName | 
        Where-Object { $_.GrantSendOnBehalfTo } |
        Select-Object @{Name="Delegate";Expression={$_.GrantSendOnBehalfTo -join ", "}}, @{Name="Permission";Expression={"SendOnBehalf"}}
    if ($SendOnBehalfPerms) {
        Write-Host "    FOUND: Send on Behalf delegates: $($SendOnBehalfPerms.Delegate)" -ForegroundColor Green
    } else {
        Write-Host "    NONE: No Send on Behalf delegates found." -ForegroundColor Gray
    }

    # Combine permissions for the mailbox
    if ($FullAccessPerms -or $SendAsPerms -or $SendOnBehalfPerms) {
        $FullAccessDelegates = if ($FullAccessPerms) { ($FullAccessPerms.Delegate -join ", ") } else { "None" }
        $SendAsDelegates = if ($SendAsPerms) { ($SendAsPerms.Delegate -join ", ") } else { "None" }
        $SendOnBehalfDelegates = if ($SendOnBehalfPerms) { $SendOnBehalfPerms.Delegate } else { "None" }

        Write-Host "  Adding mailbox $CurrentMailboxName to results with delegation permissions." -ForegroundColor Green
        $Results += [PSCustomObject]@{
            Mailbox              = $Mailbox.UserPrincipalName
            FullAccessDelegates  = $FullAccessDelegates
            SendAsDelegates      = $SendAsDelegates
            SendOnBehalfDelegates = $SendOnBehalfDelegates
        }
    } else {
        Write-Host "  No delegation permissions found for $CurrentMailboxName." -ForegroundColor Gray
    }
    Write-Host ""
}

# Complete the progress bar
Write-Progress -Activity "Processing Mailboxes for Delegation Permissions" -Completed
Write-Host "Completed processing $MailboxCount mailboxes at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Display results
Write-Host "===== Step 4: Displaying Results =====" -ForegroundColor Cyan
if ($Results) {
    Write-Host "Found $($Results.Count) mailboxes with delegation permissions:"
    Write-Host "Displaying results in a formatted table..." -ForegroundColor Yellow
    $Results | Format-Table Mailbox, FullAccessDelegates, SendAsDelegates, SendOnBehalfDelegates -AutoSize
} else {
    Write-Host "No mailboxes with delegation permissions found." -ForegroundColor Yellow
}
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Optionally, export to CSV
Write-Host "===== Step 5: Exporting Results =====" -ForegroundColor Cyan
Write-Host "Exporting results to MailboxDelegationPermissions.csv..." -ForegroundColor Yellow
$Results | Export-Csv -Path "MailboxDelegationPermissions.csv" -NoTypeInformation
Write-Host "SUCCESS: Results exported to MailboxDelegationPermissions.csv at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Prompt to disconnect from Exchange Online
Write-Host "===== Step 6: Session Cleanup =====" -ForegroundColor Cyan
Write-Host "Prompting user to disconnect from Exchange Online..." -ForegroundColor Yellow
$DisconnectResponse = Read-Host "Do you want to disconnect from Exchange Online? (Y/N)"
if ($DisconnectResponse -eq 'Y' -or $DisconnectResponse -eq 'y') {
    Write-Host "User chose to disconnect. Disconnecting from Exchange Online..." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "SUCCESS: Disconnected from Exchange Online at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
} else {
    Write-Host "User chose to remain connected to Exchange Online." -ForegroundColor Green
}

# Display script completion message
Write-Host "===== Script Completed =====" -ForegroundColor Cyan
Write-Host "Script finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
Write-Host "Check MailboxDelegationPermissions.csv for detailed results."
Write-Host "============================================" -ForegroundColor Cyan

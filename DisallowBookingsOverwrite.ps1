# Import the Exchange Online Management module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName mail@example.com

# Retrieve all meeting rooms and equipment mailboxes
$resourceMailboxes = Get-Mailbox -ResultSize Unlimited -Filter { RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'EquipmentMailbox' } | 
                     Select-Object DisplayName, UserPrincipalName, RecipientTypeDetails

# Check if any resource mailboxes were retrieved
if ($resourceMailboxes.Count -eq 0) {
    Write-Output "No meeting rooms or equipment found."
} else {
    # Iterate through each resource mailbox
    foreach ($resource in $resourceMailboxes) {
        try {
            # Get the current calendar processing settings
            $calendarProcessing = Get-CalendarProcessing -Identity $resource.UserPrincipalName
            
            # Set the calendar processing to allow requests out of policy
            Set-CalendarProcessing -Identity $resource.UserPrincipalName -AllRequestOutOfPolicy $False

            # Output the change
            Write-Output "Successfully updated calendar processing for: $($resource.DisplayName), Type: $($resource.RecipientTypeDetails)"
        } catch {
            # Handle errors and output the error message
            Write-Output "Failed to update calendar processing for: $($resource.DisplayName). Error: $_"
        }
    }
}



# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

<#
# Check if users were retrieved
if ($users.Count -eq 0) {
    Write-Output "No users found."
} else {
    # Iterate through each user and print their details to the console
    foreach ($user in $users) {
        Write-Output "Display Name: $($user.DisplayName), User Principal Name: $($user.UserPrincipalName)"
    }
}
#>

# Import the Microsoft Graph module
#Import-Module Microsoft.Graph

# Connect to Microsoft Graph
#Connect-MgGraph -Scopes "User.Read.All"

# Check if the connection was successful
#if ($null -eq (Get-MgUser -Top 1)) {
 #   Write-Host "Failed to connect to Microsoft Graph. Check your permissions." -ForegroundColor Red
 #   exit
#}

# Get all users in the Azure tenancy
#$users = Get-MgUser -Top 999

# Output user details
#$users | Select-Object Id, DisplayName, UserPrincipalName | Format-Table -AutoSize

# Disconnect from Microsoft Graph
#Disconnect-MgGraph

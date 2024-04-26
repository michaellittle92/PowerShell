# Function to check if a module is installed
function IsModuleInstalled {
    param (
        [string]$moduleName
    )

    $installedModule = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName }
    return [bool]($installedModule)
}

# Function to install a module
function InstallRequiredModule {
    param (
        [string]$moduleName
    )

    Write-Host "Installing module $moduleName..."
    Install-Module -Name $moduleName -Force -Scope CurrentUser -AllowClobber -Verbose
}

# Check if ExchangeOnlineManagement module is installed
$exchangeOnlineModule = "ExchangeOnlineManagement"
if (-not (IsModuleInstalled $exchangeOnlineModule)) {
    InstallRequiredModule $exchangeOnlineModule
}
# Check if Exchange module is installed
$exchangeModule = "ExchangeOnlineManagement"
if (-not (IsModuleInstalled $exchangeModule)) {
    InstallRequiredModule $exchangeModule
}

# Connect to Exchange Online
Connect-ExchangeOnline

# Get all distribution groups
$distributionLists = Get-DistributionGroup | Select-Object DisplayName, PrimarySmtpAddress

# Display the distribution lists
$distributionLists

# Display a menu for the user to select a distribution list
Write-Host "Select a Distribution List:`n"
for ($i = 0; $i -lt $distributionLists.Count; $i++) {
    Write-Host "$($i+1). $($distributionLists[$i].DisplayName)"
}
$selectedIndex = Read-Host "Enter the number of the distribution list"


# Validate the selected index
try {
    $selectedDistributionList = $distributionLists[$selectedIndex - 1]  # Adjust index here
} catch {
    Write-Host "Invalid selection. Please enter a valid number."
    exit
}

# Prompt the user to remove current members
$removeMembers = Read-Host "Do you want to remove all current members before adding new ones? (yes/no)"
if ($removeMembers -eq "yes") {
    Write-Host "Deletion in progress. Please wait."
    # Remove all current members
    $members = Get-DistributionGroupMember -Identity $selectedDistributionList.PrimarySmtpAddress
    foreach ($member in $members) {
        Remove-DistributionGroupMember -Identity $selectedDistributionList.PrimarySmtpAddress -Member $member.PrimarySmtpAddress -Confirm:$false
    }
    Write-Host "All current members removed from $($selectedDistributionList.DisplayName)."
}

# Prompt the user for the CSV file path
$csvFilePath = Read-Host "Enter the path to the CSV file"

# Check if the CSV file exists
if (-not (Test-Path $csvFilePath -PathType Leaf)) {
    Write-Host "The specified CSV file does not exist. Please provide a valid path."
    exit
}

# Read data from the CSV file
$csvData = Import-Csv -Path $csvFilePath

# Iterate over each row in the CSV file
foreach ($row in $csvData) {
    $contactDisplayName = $row.DisplayName
    $contactEmailAddress = $row.EmailAddress
    
    # Check if the contact already exists
    $existingContact = Get-MailContact -Filter "ExternalEmailAddress -eq '$contactEmailAddress'"

    if (-not $existingContact) {
        # Create the contact if it doesn't exist
        New-MailContact -ExternalEmailAddress $contactEmailAddress -Name $contactDisplayName
    }

# Try adding the contact to the selected distribution list
try {
    Add-DistributionGroupMember -Identity $selectedDistributionList.PrimarySmtpAddress -Member $contactEmailAddress -ErrorAction Stop
    Write-Host "Added $contactDisplayName to $($selectedDistributionList.DisplayName)"
} catch {
    if ($_.Exception.Message -match "The recipient '(.+)' is already a member of the group") {
        $existingMember = $matches[1]
        Write-Host "Skipping $($contactDisplayName): The recipient $($existingMember) is already a member of the group $($selectedDistributionList.DisplayName)."
    } else {
        Write-Host "Error adding $contactDisplayName to $($selectedDistributionList.DisplayName): $_"
    }
}

}
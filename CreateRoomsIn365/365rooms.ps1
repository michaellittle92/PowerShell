# creds prompt - update email address
Connect-ExchangeOnline -UserPrincipalName ictadmin@example.com

#room details - this should probably be some sort of a 2d array. with a loop 
$room1Name = "Room 1"
$room1Email = "room1@example.com"

$room2Name = "Room 2"
$room2Email = "room2@example.com"


# Define room list details
$roomListName = "MeetingRooms"
$roomListEmail = "meeting.rooms@example.com"

# Create the first room
New-Mailbox -Name $room1Name -Room -PrimarySmtpAddress $room1Email

# Create the second room
New-Mailbox -Name $room2Name -Room -PrimarySmtpAddress $room2Email

# Create or update the room list
# Check if the room list already exists
$roomList = Get-DistributionGroup -Identity $roomListEmail -ErrorAction SilentlyContinue

if ($null -eq $roomList) {
    # Create the room list if it does not exist
    New-DistributionGroup -Name $roomListName -PrimarySmtpAddress $roomListEmail -RoomList
} else {
    # Update the room list if it already exists
    Set-DistributionGroup -Identity $roomListEmail -DisplayName $roomListName -RoomList
}

# Add the rooms to the room list
Add-DistributionGroupMember -Identity $roomListEmail -Member $room1Email
Add-DistributionGroupMember -Identity $roomListEmail -Member $room2Email

# Verify room list members
Get-DistributionGroupMember -Identity $roomListEmail

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

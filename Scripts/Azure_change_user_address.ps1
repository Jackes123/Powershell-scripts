#Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All"

#Define the old data you wanna search for
$oldCity = ""

# Fetch all users
Write-Host "Fetching all users from Azure AD..."
$allUsers = Get-MgUser -All -Property "Id,UserPrincipalName,StreetAddress,City,postalCode"

$Oldaddress = $allUsers|Where-Object { $_.City -eq $oldCity}

# Log the number of users fetched
Write-Host "Total users fetched: $($allUsers.Count)"



#And the new data you would like to add/change:
$newStreetAddress = ""
$newCity = ""
$newpostalCode = "" 

# Loop through each CSV record to update addresses
foreach ($record in $Oldaddress) {
    # Check if there are users to update
    if ($Oldaddress.Count -gt 0) {

    # Loop through each user and update their address
        try {
            # Update the user's address details
            Update-MgUser -UserId $record.Id -StreetAddress $newStreetAddress -City $newCity -PostalCode $newPostalCode
            
            # Output the updated user's UPN for confirmation
            Write-Host "Updated address for user: $($record.UserPrincipalName)"
        }
        catch {
            Write-Host "Failed to update address for user: $($record.UserPrincipalName) - $($_.Exception.Message)"
        }
    
    }
}
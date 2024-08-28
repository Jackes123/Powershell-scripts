Import-Module Microsoft.Graph.Users
 
#This code is used for connecting to mggraph through a service principal
<#
$ClientSecretCredential = Get-Credential -Credential "(Client secret here)"
connect-mggraph -TenantId "Tenant id HERE" -ClientSecretCredential $ClientSecretCredential -nowelcome
Disconnect-Graph
#>



#Define the manager old and new manager - should be UPN since the manager field is not a text field
$oldManagerUPN = ""
$newManagerUPN = ""

# Get the old manager's object ID
$oldManager = Get-MgUser -UserId $oldManagerUPN
$oldManagerId = $oldManager.Id

# Get the new manager's object ID
$newManager = Get-MgUser -UserId $newManagerUPN
$newManagerId = $newManager.Id

# Fetch all users
Write-Host "Fetching all users from Azure AD..."
$allUsers = Get-MgUser -All -ExpandProperty "Manager" -Property "Id,UserPrincipalName"

# Filter users based on the old manager ID
$filteredUsers = $allUsers | Where-Object { 
    $_.Manager.Id -eq $oldManagerId
}

# Log the number of users fetched
Write-Host "Total users fetched: $($allUsers.Count)"

# Log the number of users filtered
Write-Host "Total users filtered: $($filteredUsers.Count)"

# Loop through each filtered user
foreach ($user in $filteredUsers) {
    # Loop through each user and update their address
    try {
        # Create the body parameter with the correct format
        $bodyParameter = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$newManagerId"
        } | ConvertTo-Json

        # Update the user's manager
        Set-MgUserManagerByRef -UserId $user.ID -BodyParameter $bodyParameter
        
        # Output the updated user's UPN for confirmation
        Write-Host "Updated manger for user: $($user.UserPrincipalName)"
    }
    catch {
        Write-Host "Failed to update manager for user: $($user.UserPrincipalName) - $($_.Exception.Message)"
    }
}
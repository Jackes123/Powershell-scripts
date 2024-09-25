# Import modules
Import-Module ImportExcel
Import-Module Microsoft.Graph.Groups

# Connect to Microsoft Graph with required permissions
Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All"

# Initialize counter
$counter = 0

# Path to excel file and output csv file
$ExcelFilePath = "C:\Users\JacobKastbjerg\OneDrive - DanBred\Desktop\test.xlsx"
$InvitedUsersCSV = "C:\Users\JacobKastbjerg\OneDrive - DanBred\Desktop\InvitedUsers.csv"

# Import excel data
$UsersFromExcel = Import-Excel -Path $ExcelFilePath

# Remove duplicates from excel data based on email
$UsersFromExcel = $UsersFromExcel | Sort-Object 'E-mail' -Unique

# Function to check if user already exists in Azure AD
function CheckUserExistsOrInvited {
    param (
        [string]$email
    )
    try {
        # Adjust filter to account for UPN format
        $user = Get-MgUser -Filter "mail eq '$email' or proxyAddresses/any(c:c eq 'smtp:$email')" -ErrorAction Stop

        if ($user) {
            if ($user.UserType -eq "Guest" -and $user.ExternalUserState -eq "PendingAcceptance") {
                return "PendingAcceptance"
            }
            return $user
        }
        return $null
    } catch {
        # Handle cases where user is not found
        return $null
    }
}

# Function to send user invitation using Microsoft Graph API
function Send-UserInvitation {
    param (
        [string]$email,
        [string]$name
    )

    $Uri = "https://graph.microsoft.com/v1.0/invitations"
    $Body = @{
        invitedUserEmailAddress = $email
        invitedUserDisplayName = $name
        sendInvitationMessage = $false
        inviteRedirectUrl = "https://danbred.com/" # Not used, but necessary
    } | ConvertTo-Json

    try {
        $response = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $Body -ContentType "application/json"
        return $response.id
    } catch {
        Write-Host "Failed to send invitation for $name ($email): $_"
        return $null
    }
}

# Check if the CSV file exists, if not create it with headers
if (-not (Test-Path $InvitedUsersCSV)) {
    "Email,Name,InvitationId" | Out-File -FilePath $InvitedUsersCSV
}

# Iterate through each email in the excel
foreach ($User in $UsersFromExcel) {   
    $email = $User.'E-mail'
    $name = $User.'Navn'
        
    if ($email -like "*@danbred.com"){
        Write-Host "Skipping user $name ($email) as they belong to the danbred.com domain."
        continue
    }

    # Check if user already exists or has a pending invitation in Azure AD
    $existingUser = CheckUserExistsOrInvited -email $email
        
    if ($existingUser -eq "PendingAcceptance") {
        Write-Host "User $name ($email) has a pending invitation. Skipping."
        continue
    } elseif ($existingUser) {
        Write-Host "User $name ($email) already exists in Azure AD. Skipping invitation."
        continue
    }

    try {
        # Send invite to external user without notification
        $InvitationId = Send-UserInvitation -email $email -name $name

        if ($InvitationId) {
            # Increment the counter only if invitation was successful
            $counter += 1

            # Append user info to the CSV file
            "$email, $name, $InvitationId" | Out-File -FilePath $InvitedUsersCSV -Append

            Write-Host "Successfully invited $name ($email)."
        }
        
    } catch {
        Write-Host "Failed to process user $name ($email): $_"
    }
}

# Display success message
Write-Host "$counter users have been invited."

# Import modules
Import-Module ImportExcel
Import-Module Microsoft.Graph.Groups

# Authenticate and connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"

$counter = 0

# Read email addresses from a file
$ExcelFilePath = "C:\Users\JacobKastbjerg\OneDrive - DanBred\Desktop\Genepro agreements.xlsx"

# Import Excel data
$UsersFromExcel = Import-Excel -Path $ExcelFilePath

# Remove duplicates from Excel data based on email
$UsersFromExcel = $UsersFromExcel | Sort-Object 'E-mail' -Unique

# File to store the results
$OutputFile = "C:\Users\JacobKastbjerg\OneDrive - DanBred\Desktop\test.txt"

# Write the header to the output file
"version:v1.0" | Out-File -FilePath $OutputFile
"Member object ID or user principal name [memberObjectIdOrUpn] Required:" | Out-File -FilePath $OutputFile -Append

# Iterate over each email address and get Object ID
foreach ($userObject in $UsersFromExcel) {
    # Extract the email address from the custom object
    $email = $userObject.'E-mail'
    
    try {
        # Try fetching the user by their email
        $user = Get-MgUser -Filter "mail eq '$email'"


        if ($user) {
            # Write the Object ID or UPN to the output file
            Write-Host "User $email with object ID: $($user.Id)"
            $user.Id | Out-File -FilePath $OutputFile -Append
            $counter += 1
            
        }
        else {
            Write-Host "User not found for email $email"
        }
    }
    catch {
        Write-Host "Error retrieving user for email ${email}: $_"
    }
}

Write-Host "Results have been written to $OutputFile"
Write-Host $counter

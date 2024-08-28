Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser


#This code is used for connecting to mggraph through a service principal
<#
$ClientSecretCredential = Get-Credential -Credential "(Client secret here)"
connect-mggraph -TenantId "Tenant id HERE" -ClientSecretCredential $ClientSecretCredential -nowelcome
Disconnect-Graph
#>


# Connect to SharePoint Online
$adminUrl = "https://(domain)-admin.sharepoint.com" #Input correct URL
Connect-SPOService -Url $adminUrl

#Manager UPN
$managerUPN = "(manager upn )" #Input the correct manager/user UPN that needs the access
#User UPN that is no longer employeed.
$UPN = "(user upn)" #Input the user/employee

#Replace @ and .
$sanitizedUPN = $UPN -replace '@', '_' -replace '\.', '_'

#URL on users sharepoint
$URL = "https://(domain)-my.sharepoint.com/personal/$sanitizedUPN"

#Set manger/user as site connection admin.
Set-SPOUser -Site $URL -LoginName $managerUPN -IsSiteCollectionAdmin: $true

#Print statement for user
Write-host "User $managerUPN has been granted Site Collection Admin rights on $UPN.`n"

Write-Host "URL where the manager can find the users onedrive folder is: `n$URL`n"
**# Powershell-scripts**
Different scripts i have created in my time working so far - with proper comments so they are easy to customize to different companies:
If you wanna connect through service principal that is done in a few of these scripts. You need to input your own client secret and tenant ID.

Short description of each script:



**Azure_change_manager_of_employees:**
This script can take a large group of people with an manager in azure called x and switch it to y. Only changes that needs to be made in the script is to input the old and new manager user principal name.

Use case: Manager either quits or gets fired.


**Azure_change_user_address:**
Script can change address from one adress to another for a large group of people. You need to input the old city name or something like it. The script uses the old city to find all users with the old address.
You also need to input information on the new location, so the new street address, new city name and new postal code.

Use case: Company changes location and a group of users need a new address.


**Change_TLS_servers:**
Script disables ssl and enables tls. Each function can easily be commented out for customization.

Use case: Servers needs TLS enabled


**Give_manager_access_to_employee_onedrive**
Script gives manager/user access to another users onedrive. Multiple things you need to change. Make sure you fix the URL. Then you need a user UPN and manager UPN. 

Use case: As part of an automatic offboarding process, the goal was to have this automated in a script and then further used in the offboarding process.

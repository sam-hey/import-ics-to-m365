# Importing ICS Files into Microsoft 365 Group Calendars

This project facilitates importing `.ics` calendar files into Microsoft 365 (M365) group calendars using PowerShell and the Microsoft Graph API. This approach addresses the challenge that direct copying of `.ics` files to M365 group calendars is not supported.



## Background

The original script was derived from [this Spiceworks community post](https://community.spiceworks.com/t/importing-ics-into-office-365-group-calendar/761145). While M365 might seem to allow copying `.ics` files to group calendars, these entries do not persist and are eventually removed, as the functionality has effectively been patched out.

This script leverages the Microsoft Graph API to overcome these limitations, providing a reliable way to import `.ics` files.



## Prerequisites

Before using the script, you must configure your Azure AD application and gather necessary credentials.

### Step 1: Registering an App in Azure AD
1. Navigate to the [Azure AD App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).
2. Create a new app registration:
   - Provide a name.
   - Select **Accounts in this organizational directory only (Single tenant)**.
   - Click **Register**.
3. Add a redirect URI:
   - Click **Add a platform** and select **Web**.
   - Set `http://localhost` as the Redirect URL.
   - Set `https://localhost` as the Logout URI.
4. Add API permissions:
   - Select **Microsoft Graph** > **Delegated permissions**.
   - Search for `group.readwrite.all` and enable it.
   - Grant admin consent for the permission.
5. Modify the app manifest:
   - Set `"allowPublicClient": null` to `"allowPublicClient": true`.
6. Generate client credentials:
   - Create a **Client Secret** and copy its value immediately.

### Step 2: Finding the Group ID
1. Navigate to [Azure AD Groups](https://portal.azure.com/#blade/Microsoft_AAD_IAM/GroupsManagementMenuBlade/AllGroups).
2. Locate the desired M365 group.
3. Copy the **Object ID** (this will serve as the `GroupId`).



## Configuration

Before running the script, update the following variables:

```powershell
#############
# Variables #
#############

# Path to the ICS file
$ics = "C:\temp\calendar\calendar.ics"

# Graph API token credentials
$TenantName = "<tenant>.onmicrosoft.com"
$ClientId = "<your-client-id>"
$ClientSecret = "<your-client-secret>"

# Office 365 Group ID
$GroupId = "<your-group-id>"
```


## Usage
Ensure you have PowerShell installed with the required modules (e.g., Microsoft.Graph).
Save the script to a .ps1 file.
Open PowerShell and execute the script:

    `.\Import-M365GroupCalendar.ps1

## References

[Original Spiceworks Post](https://community.spiceworks.com/t/importing-ics-into-office-365-group-calendar/761145)

[Guide to Authenticating the API with PowerShell](https://www.thelazyadministrator.com/2019/07/22/connect-and-navigate-the-microsoft-graph-api-with-powershell/)

## Notes

Directly copying .ics files to M365 group calendars is not supported.
This script provides a workaround by programmatically interacting with the M365 group calendar via the Graph API.

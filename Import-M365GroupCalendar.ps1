<#

    ##################################################################
    # Creating an App in your Office 365 tenancy to access Graph API #
    ##################################################################

    1. Go to https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
    2. Click 'New registration'
    3. Give it a name and select 'Accounts in this organizational directory only (<TenantName> only - Single tenant)' and click 'Register'
    4. On the 'Overview' page click 'Add a Redirect URI'
    5. Now click 'Add a platform', then 'Web'
    6. Now enter 'http://localhost' in 'Redirect URL and 'https://localhost' in 'Logout URI' and click 'Configure'
    7. Now go to 'API permissions'
    8. Click 'Add a permission' and select 'Microsoft Graph'
    9. Click 'Delegated permissions'
    10. Search for 'group.readwrite.all' and tick the box, then click 'Add permission'
    11. Now click 'Grant admin consent for <TenantName>' and click 'Yes'
    12. Now go to 'Manifest'
    13. Replace "allowPublicClient": null with "allowPublicClient": true
    14. Click 'Save'.
    15. Go to 'Overview'
    16. Copy the 'Application (client) ID' into the $ClientId below.
    17. Go to 'Certificates & secrets'
    18. Create 'New client secret'
    19. Give the secret a name and select an expiry date.
    20. Copy the 'Value' into the $ClientSecret below, IMPORTANT once you click off the value you cannot view it again, make sure you copy it right away.
    21. Finally enter your Tenant Name in the $TenantName at top. E.g. $TenantName = "<tenantname>.onmicrosoft.com"
    
    ##############################################################################
    # Finding the Group Id of the Office 365 group you want to import the ICS to #
    ##############################################################################

    1. Go to https://portal.azure.com/#blade/Microsoft_AAD_IAM/GroupsManagementMenuBlade/AllGroups
    2. Find the Office 365 group you wish to import the calendar to
    3. In the 'Overview' of the group copy the 'Object Id'
    4. Paste the object ID into $GroupId below.

    ##########
    # Credit #
    ##########

    # Guide to authenticating the API with powershell 
    # https://www.thelazyadministrator.com/2019/07/22/connect-and-navigate-the-microsoft-graph-api-with-powershell/
#>

#############
# Variables #
#############

# ICS file
$ics = "C:\temp\calendar\calendar.ics"

# Graph API token
$TenantName = "<tenant>.onmicrosoft.com"
$ClientId = ""
$ClientSecret = ""

# Office 365 Group ID
$GroupId = ""

#########
# Setup #
#########

# Ensure system is using TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;

#####################
# ICS to Data Table #
#####################

# Create Table
$tbl  = New-Object System.Data.DataTable "Events"
$col1 = New-Object System.Data.DataColumn "Subject"
$col2 = New-Object System.Data.DataColumn "Body"
$col3 = New-Object System.Data.DataColumn "Start"
$col4 = New-Object System.Data.DataColumn "End"
$col5 = New-Object System.Data.DataColumn "Location"
$tbl.Columns.Add($col1)
$tbl.Columns.Add($col2)
$tbl.Columns.Add($col3)
$tbl.Columns.Add($col4)
$tbl.Columns.Add($col5)      

# Process ICS into Powershell table
Get-Content $ics -Encoding UTF8 | foreach-Object {

    # Split key:value
    if($_.Contains(':')){
        
        $z = @{ $_.split( ':')[0] =( $_.split( ':')[1]).Trim() }
    

        # Begin
        if ($z.keys -eq "BEGIN"){         
        
        }

        # Get start date
        if ($z.keys -eq "DTSTART;VALUE=DATE") {
            $Start = $z.values -replace "r\n\s"
            $Start = [datetime]::ParseExact($Start,"yyyyMMdd" ,$null)
            $StartDate = $Start.ToShortDateString()           
            $StartDate = get-date $StartDate -Format yyyy-MM-dd
            $StartTime = "00:00:00"           
        }

        if ($z.keys -eq "DTSTART") {
            $Start = $z.values -replace "r\n\s"           
            $Start = $Start -replace "T"           
            $Start = $Start -replace "Z"           
            $Start = [datetime]::ParseExact($Start,"yyyyMMddHHmmss" ,$null)           
            $StartDate = $Start.ToShortDateString()           
            $StartTime = $Start.ToLongTimeString()
            $StartDate = get-date $StartDate -Format yyyy-MM-dd
        }

        # Get end date
        if ($z.keys -eq "DTEND") {
            $End = $z.values -replace "\r\n\s"           
            $End = $End -replace "T"           
            $End = $End -replace "Z"           
            $End = [datetime]::ParseExact($End,"yyyyMMddHHmmss" ,$null)           
            $EndDate = $End.ToShortDateString()           
            $EndTime = $End.ToLongTimeString()                      
            $EndDate = get-date $EndDate -Format yyyy-MM-dd                      
        }

        if ($z.keys -eq "DTEND;VALUE=DATE") {
            $End = $z.values -replace "r\n\s"
            $End = [datetime]::ParseExact($End,"yyyyMMdd" ,$null)
            $EndDate = $End.ToShortDateString()
            $EndDate = get-date $EndDate -Format yyyy-MM-dd
            $EndTime = "00:00:00"      
        }
      
        # Get summary
        if ($z.keys -eq "SUMMARY") {           
            $Title = $z.values -replace "\r\n\s"           
            $Title = $z.values -replace ",","-"        
        }

        # Get description
        if ($z.keys -eq "DESCRIPTION") {
            $Description = $z.values -replace "\r\n\s"           
            $Description = $Description -replace "<p>"
            $Description = $Description -replace "</p>"        
            $Description = $Description -replace "<div> </div>"  
            $Description = $Description -replace "<div>"
            $Description = $Description -replace "</div>"
        }

        # Get location
        if ($z.keys -eq "LOCATION") {           
            $Location = $z.values -replace "\r\n\s"           
            $Location = $z.values -replace ",","-"           
        }    

        # End of event
        if ($z.keys -eq "END") {

            # Check Subject exists
            if($Title -ne ""){
            
                # Add to Table
                $row = $tbl.NewRow()
                $row.Subject = "$Title"
                $row.Body = "$Description"
                $row.Start = "$($StartDate)T$($StartTime)"
                $row.End = "$($EndDate)T$($EndTime)"
                $row.Location = "$Location"
                $tbl.Rows.Add($row)
            }

            # Clear variables
            $Start = ""
            $End = ""
            $Title = ""  
            $Description = ""
            $Location = ""
            $EndDate = ""
            $EndTime = ""         
            $StartDate = ""
            $StartTime = ""
        }
    }
}

###########################
# Data Table to Graph API #
###########################

# UrlEncode the ClientID and ClientSecret and URL's for special characters 
Add-Type -AssemblyName System.Web
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($ClientId)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecret)
$redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode("http://localhost")
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com")
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/group.readwrite.all")

# Function to popup Auth Dialog Windows Form
Function Get-AuthCode {
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }

    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri        
        if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }
}

# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode

# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUriEncoded&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop

# Count rows
$i = 1

# Add each event
foreach($event in $tbl){
    
    # Increment counter
    $i++
    
    # Check if HTTP 429, maximum of 10000 requests, per user, per 10 minutes, per application
    if($i -eq 10000){
        $timer = New-TimeSpan -Minutes 10
        Write-Host "Sleeping for 10 minutes whilst waiting for the API throttling to reset. Please wait until $((Get-Date) + $timer)"
        Start-Sleep -Seconds (60*10) # Sleep 
        $i = 1
    }
    
    # Generate request
    $hash = @{
                subject = $event.Subject; 
                body = @{
                    contentType = "html";
                    content = $event.Body;
                };
                start = @{
                    dateTime = $event.Start;
                    timeZone = "GMT Standard Time";
                };
                end = @{
                    dateTime = $event.End;
                    timeZone = "GMT Standard Time";
                };
                location = @{
                    displayName = $event.Location;
                };
            }

    $JSON = $hash | ConvertTo-Json

    # Perform query
    $apiUrl = "https://graph.microsoft.com/v1.0/groups/$($GroupId)/calendar/events"
    $Data   = Invoke-RestMethod -Headers @{Authorization = "Bearer $($TokenResponse.access_token)"} -Uri $apiUrl -Method Post -Body $JSON -ContentType "application/json"
}
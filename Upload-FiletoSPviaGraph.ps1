# This script call the Microsoft Graph API to upload a file to a sharepoint site
# You need a Service Principal in the AAD with application permissions
# Files.ReadWrite.all
# Sites.Read.All
# Sites.Selected
#
# Then you need to change the secrets.json file and fill in your secrets


function GetAccessToken {
    # Create a JSON with all the secrets 
    $file = Get-Content -Path .\Secrets.json | ConvertFrom-Json
    # AppID
    $AppClientId= $file.AppClientId
    # Secret of the App. Never commit it to a code repository!
    $Secret = $file.Secret
    $TenantID = $file.TenantID
    # Because we are in a machine 2 machine scenario we direct call the token endpoint
    # We don't need to access to authoritazion endpoint
    $uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    # Client_credentials is the auth flow type
    # scope is always .default because. AAD will provide us all scopes in the token that is configured in the API permissions of the App
    $body = @{
        grant_type = "client_credentials"
        client_id = $AppClientId
        client_secret = $Secret
        scope = "https://graph.microsoft.com/.default"
    }
    # Call the endpoint
    $response1 = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType "application/x-www-form-urlencoded"
    # Copy token to another variable
    $AccessToken = $response1.access_token
    # Prepare the next request 
    return $AccessToken
}

function GetSiteID {
    param (
        [string]$Searchterm,
        [string]$AccessToken
    )
    $headers = @{'Content-Type'="application\json";'Authorization'="Bearer $AccessToken"}
    # Get site id from site. Please change the search term
    $apiurl3 = "https://graph.microsoft.com/v1.0/sites"
    $response3 = Invoke-RestMethod -Headers $headers -Uri $apiurl3 -Method Get
    $array = $response3.value
    $index = [Array]::IndexOf($array.name, $Searchterm)
    #$index
    $siteid = $array[$index].id
    return $siteid
}

# Get content from site id
function GetSiteContent {
    param (
        [string]$sitename
        )
    $AccessToken = GetAccessToken
    $siteid = GetSiteID -Searchterm $sitename -AccessToken $AccessToken
    $apiurl5 = "https://graph.microsoft.com/v1.0/sites/$siteid/drive/items/root/children"
    $response5 = Invoke-RestMethod -Headers $headers -Uri $apiurl5 -Method Get
    return $response5.value.name
}

# Upload a file to  site
function UploadFile {
    param (
        [string]$Path,
        [string]$SiteName
    )
    $AccessToken = GetAccessToken
    $headers6 = @{'Content-Type'="text/plain";'Authorization'="Bearer $AccessToken"}
    $siteid = GetSiteID -Searchterm $SiteName -AccessToken $AccessToken 
    $filename = $Path.Substring(3)
    $apiurl6 = "https://graph.microsoft.com/v1.0/sites/$siteid/drive/items/root:/${filename}:/content"
    $response6 = Invoke-RestMethod -Headers $headers6 -Uri $apiurl6 -Method Put -InFile $Path
    return $response6.createdDateTime
}

UploadFile -Path "C:\Henrik-NonProfit-Engel.png" -SiteName "FileUpload"
GetSiteContent -sitename "FileUpload"




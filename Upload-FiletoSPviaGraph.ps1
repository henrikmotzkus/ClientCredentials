# This demo script calls the Microsoft Graph API to upload a file to a sharepoint site
# You need a Service Principal in the AAD with application permissions
#
# Sites.Selected
#
# Then you need to change the secrets.json file and fill in your secrets
#
# First prepare your app registration in order to upload files
# https://techcommunity.microsoft.com/t5/microsoft-sharepoint-blog/develop-applications-that-use-sites-selected-permissions-for-spo/ba-p/3790476
#
#
# You need the site ID first
#
#
#



function GetAccessToken {
    # Create a JSON with all the secrets 
    $file = Get-Content -Path .\Secrets.json | ConvertFrom-Json
    # AppID
    $AppClientId = $file.AppClientId
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

# Not used anymore
<#
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
#>

# Get content from site id
function GetSiteContent {
    param (
        [string]$SiteId
        )
    $AccessToken = GetAccessToken
    #$sitename = "FileUpload"
    #$siteid = GetSiteID -Searchterm $sitename -AccessToken $AccessToken
    $siteid = "motzkus.sharepoint.com,3a3fbf16-19c9-4938-b37c-4ae52431c9e5,193021de-64a7-40dd-91f9-8e8f0effd11d"
    $apiurl5 = "https://graph.microsoft.com/v1.0/sites/$siteid/drive/items/root/children"
    $headers = @{'Content-Type'="application\json";'Authorization'="Bearer $AccessToken"}
    $response5 = Invoke-RestMethod -Headers $headers -Uri $apiurl5 -Method Get
    $response5.value
    return $response5.value.name
}

# Upload a file to  site
function UploadFile {
    param (
        [string]$Path,
        [string]$SiteId
    )
    $AccessToken = GetAccessToken
    $headers6 = @{'Content-Type'="text/plain";'Authorization'="Bearer $AccessToken"}
    $filename = $Path.Substring(3)
    $apiurl6 = "https://graph.microsoft.com/v1.0/sites/$siteid/drive/items/root:/${filename}:/content"
    $response6 = Invoke-RestMethod -Headers $headers6 -Uri $apiurl6 -Method Put -InFile $Path
    return $response6.createdDateTime
}

$file = Get-Content -Path .\Secrets.json | ConvertFrom-Json
$siteid = $file.siteid

GetSiteContent -SiteId $siteid
UploadFile -Path "C:\Users\hemotzku\OneDrive - Microsoft\Desktop\Template.xlsx" -SiteId $siteid

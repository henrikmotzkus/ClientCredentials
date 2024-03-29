# This script call the Microsoft Graph API to get the sharepoint root site

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

# Print out the token
Write-Host  $response1.access_token

# Copy token to another variable
$AccessToken = $response1.access_token

# Prepare the next request 
$headers = @{'Content-Type'="application\json";'Authorization'="Bearer $AccessToken"}
 
# Get Sharepoint Root site
$apiurl = "https://graph.microsoft.com/v1.0/sites/root"
$response2 = Invoke-RestMethod -Headers $headers -Uri $ApiUrl -Method Get
$response2.value
$response2.webUrl
$response2
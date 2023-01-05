# This script will create a SharePoint site at the specified URL and will then groupify the site and then teamify the site.

$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #Change this to your own AdminService App ID
$TenantID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #Change this to your own tenant ID
$SPHostUrl = "https://mytenant.sharepoint.com" #Change this to your SP host URL
$newSiteUrl = ("{0}/teams/new-site" -f $SPHostUrl)
$scopes = ("{0}/AllSites.FullControl" -f $SPHostUrl)
$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

$authResponse = Get-MsalToken -ClientId $clientId -TenantId $TenantID -Scopes $scopes -RedirectUri $redirectUri -Interactive

Write-Host "Access Token Successful"

$spHeader = @{"Accept"="application/json;odata=verbose"; "Content-Type"="application/json;odata=verbose"; "Authorization" = "Bearer $($authResponse.AccessToken)"}

$requestBody = @{
    "request" = @{
        "Title" = "New Site Name";
        "Url" = $newSiteUrl;
        "Lcid" = 1033;
        "ShareByEmailEnabled" = $false;
        "Description" = "Description";
        "WebTemplate" = "STS#3";
        "Owner" = "user@domain.com";
    }
}

try {
    Write-Host "Creating Site"
    $createResponse = Invoke-RestMethod -Method Post -Uri ("{0}/_api/SPSiteManager/create" -f $SPHostUrl) -Headers $spHeader -Body (ConvertTo-Json $requestBody)
    $digestResponse = Invoke-RestMethod -Method Post -Uri ("{0}/_api/contextinfo" -f $createResponse.d.Create.newSiteUrl) -Headers $spHeader
    $spHeader.Add("X-RequestDigest",$digestResponse.d.GetContextWebInformation.FormDigestValue)
    if ($createResponse) {
        Write-Host "Get Groupify Date"
        $getconversionDataResponse = Invoke-RestMethod -Method Post -Uri ("{0}/_api/GroupSiteManager/GetGroupSiteConversionData" -f $createResponse.d.Create.newSiteUrl) -Headers $spHeader
        if ($getconversionDataResponse) {
            Write-Host "Groupify Site"
            $groupifyResponse = Invoke-RestMethod -Method Post -Uri ("{0}/_api/GroupSiteManager/CreateGroupForSite" -f  $createResponse.d.Create.newSiteUrl) -Headers $spHeader -Body (ConvertTo-Json @{
                "alias" = $newSiteUrl.Substring($newSiteUrl.LastIndexOf("/")+1)
                "displayName" = $requestBody.request.Title;
                "isPublic" = $false;
                "optionalParams" = @{
                    "Description" = $requestBody.request.Description;
                    "CreationOptions" = @{
                        "results" = ([string[]]@("SPSiteLanguage:10330","test"));
                    };
                    "Classification" = "";
                    "Owners" = $getconversionDataResponse.d.GetGroupSiteConversionData.SuggestedOwners.results
                }
            } -Depth 100)

            $teamifyResponse = Invoke-RestMethod -Method Post -Uri ("{0}/_api/groupsitemanager/EnsureTeamForGroup" -f $createResponse.d.Create.newSiteUrl) -Headers $spHeader
        }
        
    }
}
catch {
    $_
}
##########################################################################################################
# Disclaimer
# This sample code, scripts, and other resources are not supported under any Microsoft standard support 
# program or service and are meant for illustrative purposes only.
#
# The sample code, scripts, and resources are provided AS IS without warranty of any kind. Microsoft 
# further disclaims all implied warranties including, without limitation, any implied warranties of 
# merchantability or of fitness for a particular purpose. The entire risk arising out of the use or 
# performance of this material and documentation remains with you. In no event shall Microsoft, its 
# authors, or anyone else involved in the creation, production, or delivery of the sample be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of the 
# use of or inability to use the samples or documentation, even if Microsoft has been advised of 
# the possibility of such damages.
##########################################################################################################

Import-Module MSAL.PS
Import-Module SharePointPnPPowerShellOnline

$inboxUpn = "user@domain.com" # user or shared mailbox UPN
$spSiteId = "https://tenant.sharepoint.us/sites/siteUrl" # SharePoint site URL where the queue list is stored
$spListName = "Queue" # processing queue list name
$folderName = "Inbox" # mailbox folder for messages

$config = @{
    "auth" = @{
        "appId" = ""; # app id registration
        "tenantId" = "tenant.onmicrosoft.com"; # tenant id or domain name
        "redirectUri" = "http://localhost"; # used for interactive login
        "loginBaseUrl" = "https://login.microsoftonline.com"; # Login url for tenant type
    };
    "locations" = @{
        "inboxUpn" = "user@domain.com"; # user or shared mailbox UPN
        "spSiteId" = "https://tenant.sharepoint.com/sites/siteUrl"; # SharePoint site URL where the queue list is stored
        "spListName" = "Queue"; # "Queue"; # processing queue list name
        "folderName" = "Inbox"; # mailbox folder for messages
        "graphBaseUrl" = "https://graph.microsoft.com"; # Base URL for graph API
    };
}

# Azure AD app API Permissions
#   Graph API
#       Delegated
#           Mail.Read.Shared
#   SharePoint API
#       Delegated
#           
### App Registration Manifest ###
#
# "requiredResourceAccess": [
# 		{
# 			"resourceAppId": "00000003-0000-0000-c000-000000000000",
# 			"resourceAccess": [
# 				{
# 					"id": "7b9103a5-4610-446b-9670-80643382c1fa",
# 					"type": "Scope"
# 				}
# 			]
# 		},
# 		{
# 			"resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
# 			"resourceAccess": [
# 				{
# 					"id": "640ddd16-e5b7-4d71-9690-3f4022699ee7",
# 					"type": "Scope"
# 				}
# 			]
# 		}
# 	]

try {
    # Here we initially login using delegated (user) permissions for the registered app id using an interactive login process 
    # for Graph API resources. Then, use implicit auth to silently get an additional access token for SharePoint API delegated 
    # permissions to write to items. The SharePoint access token is then used by the PnP PowerShell cmdlets.

    # Prompt for Graph access 
    $auth = Get-MSALToken -ClientId $config.auth.appId -Authority "$($config.auth.loginBaseUrl)/$($config.auth.tenantId)" -Scopes "https://graph.microsoft.us/.default" -RedirectUri $config.auth.redirectUri

    # Silently aquire SharePoint token - Implicit auth
    [Uri]$spSiteIdUri = New-Object -TypeName Uri -ArgumentList $config.locations.spSiteId
    $spToken = Get-MSALToken -Scopes "$($spSiteIdUri.Scheme)://$($spSiteIdUri.Authority)/AllSites.Write" -ClientId $config.auth.appId -Authority "$($config.auth.loginBaseUrl)/$($config.auth.tenantId)" -RedirectUri $config.auth.redirectUri -LoginHint $auth.User.DisplayableId -Silent
    Connect-PnPOnline -Url $config.locations.spSiteId -AccessToken $spToken.AccessToken
    
    # Get mailbox folder info
    $folderId = $null
    [bool]$continue = $true
    $folderInfo = Invoke-RestMethod -Method Get -Uri "$($config.locations.graphBaseUrl)/v1.0/users/$($config.locations.inboxUpn)/mailfolders" -Headers @{ "Authorization" = "Bearer $($auth.AccessToken)" }
    do {
        if ($folderInfo.'@odata.nextLink') {
            $continue = $true
        } else {
            $continue = $false
        }
        $folderId = ($folderInfo.value |? {$_.displayName -eq $folderName }).id
        if (-not $folderId) {
            $folderInfo = Invoke-RestMethod -Method Get -Uri $folderInfo.'@odata.nextLink' -Headers @{ "Authorization" = "Bearer $($auth.AccessToken)" }
        }
    } until (-not $continue)

    # Get first batch of messages from the mailbox folder
    $messages = Invoke-RestMethod -Method Get -Uri "$($config.locations.graphBaseUrl)/v1.0/users/$($config.locations.inboxUpn)/mailFolders/$($folderId)/messages" -Headers @{ "Authorization" = "Bearer $($auth.AccessToken)" }

    # Process items in first batch
    foreach ($msg in $messages.value) {
        $body = $msg.body.content | ConvertFrom-Json
        Add-PnPListItem -List $spListName -Values @{
            "Title" = $msg.id;
            "status" = "not started";
            "Message_Date" = $msg.receivedDateTime;
            "Sender" = $msg.from;
            "DestinationUrl" = ("{0}, {1}" -f $body.destinationFolderURL, $body.destinationFolderURL);
        }
    }

    # Evaluate if there are additional messages to retreive
    if ($messages.'@odata.nextLink') {
        do {
            # keep getting items until no nextLink
            if ($messages.'@odata.nextLink') {
                $messages = Invoke-RestMethod -Method Get -Uri $messages.'@odata.nextLink' -Headers @{ "Authorization" = "Bearer $($auth.AccessToken)" }
            }

            # process additional batches
            foreach ($msg in $messages.value) {
                $body = $msg.body.content | ConvertFrom-Json
                Add-PnPListItem -List $spListName -Values @{
                    "Title" = $msg.id;
                    "status" = "not started";
                    "Message_Date" = $msg.receivedDateTime;
                    "Sender" = $msg.from;
                    "DestinationUrl" = ("{0}, {1}" -f $body.destinationFolderURL, $body.destinationFolderURL);
                }
            }
        } until (-not $messages.'@odata.nextLink')
    }
}
catch {
    $_
}
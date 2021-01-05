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
# the possibility of such damages.
##########################################################################################################

Import-Module MSAL.PS
Import-Module SharePointPnPPowerShellOnline

$inboxUpn = "user@domain.com" # shared mailbox UPN
$spSiteIdUri = "https://tenant.sharepoint.us"
$mailboxQueuFolderName = "Inbox" # mailbox folder for messages
$mailboxProcessedFolderName = "processed"
$mailboxFailedFolderName = "failed"

$config = @{
    "appId" = ""; # app id registration
    "tenantId" = "tenant.onmicrosoft.com"; # tenant id
    "redirectUri" = "http://localhost"; # used for interactive login
}


function Save-AttachmentToSharePoint($destinationFolderUrl, $bytes, $filename) {
    $idx = $destinationFolderUrl.IndexOf("/Forms/AllItems.aspx")
    if ($idx -gt 0) {
        $correctedSiteUrl = $destinationFolderUrl.Substring(0,$idx).Replace('https://','')
    } else {
        $correctedSiteUrl = $destinationFolderUrl.Replace('https://','')
    }
    $urlParts = $correctedSiteUrl.Replace('https://','') -split '/'

    $continueLoop = $true
    $loopCounter = $urlParts.Length

    do {
        $combineParts = @()
        $pathCounter = 0

        do {
            $combineParts += $urlParts[$pathCounter]
            $pathCounter = $pathCounter + 1
        } until ($pathCounter -ge $loopCounter - 1)

        $siteUrl = [string]::Concat('https://', $combineParts -join '/')
        $folderUrl = $destinationFolderUrl.Replace($siteUrl, '').Replace('%20',' ')

        # attempt to create file here, if success set $continueLoop = $false
        try {
            Write-Output ("Attempting to save file: {0}; to site: {1}; in folder: {2}" -f $filename, $siteUrl, $folderUrl)
            Connect-PnPOnline -Url $siteUrl -UseWebLogin
            $web = Get-PnPWeb -ErrorAction SilentlyContinue
            if ($web -eq $null) {
                throw
            }
            $f = Add-PnPFile -FileName $filename -Folder $folderUrl -Stream $bytes
            if ($f) {
                Write-Output "SUCCESS"
            }
            $continueLoop = $false
            Disconnect-PnPOnline
        }
        catch {
        Write-Output "FAILED"
            $loopCounter = $loopCounter - 1
            if ($loopCounter -le 3) {
                $continueLoop = $false
                throw
            }
        }
    } until ($continueLoop -eq $false) 
}

function Get-MessageAttachments($message) {
    Write-Output "Processing message id: $($message.id)"
    $msgAttachments = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.us/v1.0/users/$($inboxUpn)/mailFolders/$($message.parentFolderId)/messages/$($message.id)/attachments" -Headers $requestHeader
    foreach ($attachmentInfo in $msgAttachments.value) {

        # convert base64 string to memory stream
        $iconimageBytes = [Convert]::FromBase64String($msgAttachments.value[0].contentBytes)
        $ims = New-Object IO.MemoryStream($iconimageBytes, 0, $iconimageBytes.Length)
        $ims.Write($iconimageBytes, 0, $iconimageBytes.Length);
        $ims.Position = 0;

        $destFolderInfo = $message.body.content | ConvertFrom-Json
        Save-AttachmentToSharePoint -destinationFolderUrl $destFolderInfo.destinationFolderURL -bytes $ims -filename $attachmentInfo.name
    }
}

function Move-MsgToArchive($message) {
    $result = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.us/v1.0/users/$($inboxUpn)/mailFolders/$($message.parentFolderId)/messages/$($message.id)/move" -Headers $requestHeader -Body "{ 'destinationId':'$($archiveFolder.id)' }"
}

function Move-MsgToFailed($message) {
    $result = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.us/v1.0/users/$($inboxUpn)/mailFolders/$($message.parentFolderId)/messages/$($message.id)/move" -Headers $requestHeader -Body "{ 'destinationId':'$($failedFolder.id)' }"
}

try {
    $auth = Get-MSALToken -ClientId $config.appId -Authority "https://login.microsoftonline.us/$($config.tenantId)" -Scopes "https://graph.microsoft.us/Mail.ReadWrite.Shared" -RedirectUri $config.redirectUri

    # Try using the below method to login
    # Connect-PnPOnline -Url $spSiteIdUri -CurrentCredentials

    #Connect-PnPOnline -Url $spSiteIdUri
    #$spToken = Get-MSALToken -ClientId $config.appId -Authority "https://login.microsoftonline.us/$($config.tenantId)" -Scopes "$($spSiteIdUri)/AllSites.Write"  -RedirectUri $config.redirectUri

    $requestHeader = @{
        "Authorization" = "Bearer $($auth.AccessToken)";
        "Accept" = "application/json";
        "Content-Type" = "application/json";
    }

    # Get mailbox folder info
    $folderId = $null
    [bool]$continue = $true
    $folderInfo = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.us/v1.0/users/$($inboxUpn)/mailfolders?`$expand=childFolders" -Headers $requestHeader
    do {
        $bobJFolder = $folderInfo.value |? {$_.displayName -eq $mailboxQueuFolderName }
        $archiveFolder = $bobJFolder.childFolders |? {$_.displayName -eq $mailboxProcessedFolderName }
        $failedFolder = $bobJFolder.childFolders |? {$_.displayName -eq $mailboxFailedFolderName }

        if ($folderInfo.'@odata.nextLink') {
            $continue = $true
            $folderInfo = Invoke-RestMethod -Method Get -Uri $folderInfo.'@odata.nextLink' -Headers $requestHeader
        } else {
            $continue = $false
        }
    } until (-not $continue)

    # Get first batch of messages from the mailbox folder
    $messages = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.us/v1.0/users/$($inboxUpn)/mailFolders/$($bobJFolder.id)/messages" -Headers $requestHeader

    $counter = 1
    # Process items in first batch
    do {
        foreach ($msg in $messages.value) {
            Write-Progress -Activity "Processing folder messages" -PercentComplete ( ($counter++ / $bobJFolder.totalItemCount) * 100 )
            if ($msg.hasAttachments) {
                try {
                    Get-MessageAttachments -message $msg
                    Move-MsgToArchive -message $msg
                }
                catch {
                    Write-Error $_
                    Move-MsgToFailed -message $msg
                }
            }
        }
    }
    until (-not $messages.'@odata.nextLink')
}
catch {
    $_
} 

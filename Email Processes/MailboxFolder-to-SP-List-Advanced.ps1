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


# Important
#
# The email body must be a plain text message with the following JSON:
# { 'destinationFolderURL': 'https://tenant.sharepoint.com/sites/siteurl/doclib/folder' }

Import-Module MSAL.PS
# There is currently an issue with any version of SharePointPnPPowerShellOnline above 3.21.2005.2 which does not create a SharePoint context when using an access token.
Import-Module SharePointPnPPowerShellOnline -RequiredVersion 3.21.2005.2 -ErrorAction Stop

#region configuration and variables

$config = @{
    "batch" = @{
        "maxJobTime" = 180;            # Max amount of seconds each message process is allowed to run
        "batchSize" = 4;               # How many items to process at a time - Limit to 4 - Graph API does not allow more than 4 concurrent connections to a mailbox
        "waitTime" = 2;                # Seconds to pause between checking job status
        "LogFolder" = ".\MailLog";     # Folder path where log files are stored
        "LogRetentionDays" = 7;        # How many days log files remain before deletion
    };
    "auth" = @{ # Azure AD Auth
        "appId" = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";  # app id registration
        "tenantId" = "tenant.onmicrosoft.com";             # tenant id
        "redirectUri" = "http://localhost";                # used for interactive login
        "authorityUri" = "https://login.microsoftonline.com/tenant.onmicrosoft.com"; # use login.microsoftonline.com for commercial; login.microsoftonline.us for GCC High
        "spOauthScopes" = "https://tenant.sharepoint.com/AllSites.Write";    # scopes used to connect to SharePoint
        "graphOauthScopes" = "https://graph.microsoft.com/Mail.ReadWrite.Shared";  # scopes used to connect tot the graph API
        "windowsIntegrated" = $false;    # $true = Use windows integrated auth; $false = Prompt for authentication
    };
    "exSource" = @{
        "graphBase" = "https://graph.microsoft.com";           # use https://graph.microsoft.com for commercial; https://graph.microsoft.us for GCC High
        "inboxUpn" = "user@domain.com";                        # shared mailbox UPN
        "queueFolderName" = "Queue";                           # Root email folder where new message will be processed
        "successSubFolder" = "processed";                      # queue sub-folder where successful messages are moved
        "failedSubFolder" = "failed";                          # queue sub-folder where failed messages are moved
    };
    "spDest" = @{
        "spHostUrl" = "https://tenant.sharepoint.com";    # Root SharePoint host and protocol URL i.e. https://tenant.sharepoint.com
    };
}

# application global variables
$global:spAuthToken = $null      # current SharePoint Token
$global:graphAuthToken = $null   # current Exchange Token
$global:queueFolder = $null      # folder of messages to process
$global:successFolder = $null    # folder where messages processed are moved. sub-folder of queueFolder
$global:failedFolder = $null     # folder where failed messages are moved. sub-folder of queueFolder

#endregion

#region Authentication

function Get-SharePointToken() {
    if ($spAuthToken -eq $null -or (($spAuthToken.ExpiresOn.AddMinutes(-10) - (Get-Date)).TotalMinutes -lt 0)) {
        $Global:spAuthToken = Get-MSALToken -ClientId $config.auth.appId -Authority $config.auth.authorityUri -Scopes $config.auth.spOauthScopes -RedirectUri $config.auth.redirectUri -IntegratedWindowsAuth:($config.auth.windowsIntegrated)
    }

    if ($spAuthToken) {
        return $spAuthToken.AccessToken
    }
}


function Get-GraphAuthHeader() {
    $header = @{
        "Authorization" = ("Bearer {0}" -f (Get-GraphToken));
        "Accept" = "application/json";
        "Content-Type" = "application/json";
    }
    return $header
}

function Get-GraphToken() {
    try {
        if ($graphAuthToken -eq $null -or (($graphAuthToken.ExpiresOn.AddMinutes(-10) - (Get-Date)).TotalMinutes -lt 0)) {
            $authResponse = Get-MSALToken -ClientId $config.auth.appId -Authority $config.auth.authorityUri -Scopes $config.auth.graphOauthScopes -RedirectUri $config.auth.redirectUri -IntegratedWindowsAuth:($config.auth.windowsIntegrated)
            $authResponse.AccessToken
            $Global:graphAuthToken = $authResponse
        }

        if ($graphAuthToken) {
            return $graphAuthToken.AccessToken
        } else {

            throw
        }
    }
    catch {
        return $null
    }
}

#endregion

#region Logging

$newLogEntry = @{
    "ReceivedDate" = "";
    "Sender" = "";
    "AttachmentName" = "";
    "Status" = "";
    "DestinationFolder" = "";
    "Notes" = "";
}

function Extract-PropertyValue($inputStr, $key) {
    if ($inputStr -match ("{0}=.+" -f $key)) {
        $splitValue = $Matches[0] -split '='
        return $splitValue[1].ToString().Trim()
    } else {
        return $null
    }
}

function Extract-NonProperty($inputStr) {
    $lines = $inputStr -split "`r`n"
    return $lines -match "^(?!(\w+=))"
}

function Get-LogFileName() {
    $currentDateTime = Get-Date
    $logTime = $currentDateTime.AddMinutes(-($currentDateTime.Minute % 30))
    
    return ("{0}.log" -f $logTime.ToString("MM-dd-yyyyTHH-mm"))
}

function Write-LogFile ($logEntry) {
    $logFormat = '{0},"{1}","{2}","{3}","{4}","{5}"'
    $logFilePath =  Join-Path $config.batch.LogFolder (Get-LogFileName)
    if (!(Test-Path -Path $logFilePath)) {
        Add-Content -Path $logFilePath -Value "ReceivedDate,Sender,Status,AttachmentName,DestinationFolder,Notes"
    }
    $notesContent = ""
    if ($logEntry.Notes) {
        $notesContent = $logEntry.Notes.ToString().Replace("`r`n", " ")
    }
    Add-Content -Path $logFilePath -Value ($logFormat -f $logEntry.ReceivedDate, $logEntry.Sender, $logEntry.Status, $logEntry.AttachmentName, $logEntry.DestinationFolder, $notesContent)
}

function Clean-LogFiles () {
    $logFolder = $config.batch.LogFolder
    if ([string]::IsNullOrWhiteSpace($logFolder) -eq $false -and $logFolder -ne $PSCommandPath) {
        Get-ChildItem -Path $logFolder -Include *.log |? { $_.CreationTime -le (Get-Date).AddDays(-($config.batch.LogRetentionDays)) } | Remove-Item -Confirm:$true -Force
    }
}

#endregion

#region Mailbox Actions

function Get-MailFolders() {
    $allFolders = $null
    [bool]$continue = $true
    $requestHeader = Get-GraphAuthHeader
    $allFolders = Invoke-RestMethod -Method Get -Uri ("{0}/v1.0/users/{1}/mailfolders?`$expand=childFolders" -f $($config.exSource.graphBase), $($config.exSource.inboxUpn)) -Headers $requestHeader 
    do {
        $global:queueFolder = $allFolders.value |? {$_.displayName -eq $config.exSource.queueFolderName }
        $global:successFolder = $queueFolder.childFolders |? {$_.displayName -eq $config.exSource.successSubFolder }
        $global:failedFolder = $queueFolder.childFolders |? {$_.displayName -eq $config.exSource.failedSubFolder }

        if (-not ($queueFolder) -and ($allFolders.'@odata.nextLink')) {
            $continue = $true
            $allFolders = Invoke-RestMethod -Method Get -Uri $allFolders.'@odata.nextLink' -Headers $requestHeader
        } else {
            $continue = $false
        }
    } until (-not $continue)
}

#endregion

#region Thread ScriptBlock
$scriptBlock = {
    param (
        [Parameter(Mandatory=$true)]
        $message,

        [Parameter(Mandatory=$true)]
        $inboxUpn, 

        [Parameter(Mandatory=$true)]
        $requestHeader,

        [Parameter(Mandatory=$true)]
        $spToken,

        [Parameter(Mandatory=$true)]
        $archiveFolder,

        [Parameter(Mandatory=$true)]
        $failedFolder
    )

    # There is currently an issue with any version of SharePointPnPPowerShellOnline above 3.21.2005.2 which does not create a SharePoint context when using an access token.
    Import-Module SharePointPnPPowerShellOnline -RequiredVersion 3.21.2005.2 -WarningAction SilentlyContinue | Out-Null


    if ($message.hasAttachments) {
        try {
            Write-Output ("MessageId={0}" -f $message.id)
            Write-Output ("Sender={0}" -f $message.sender.emailAddress.address)
            Write-Output ("ReceivedDate={0}" -f $message.receivedDateTime)
            $msgAttachments = $message.attachments
            foreach ($attachmentInfo in $msgAttachments) {
                Write-Output ("AttachmentName={0}" -f $attachmentInfo.name)

                # convert attachment base64 string to memory stream which the Add-PnPFile can use
                $iconimageBytes = [Convert]::FromBase64String($attachmentInfo.contentBytes)
                $ims = New-Object IO.MemoryStream($iconimageBytes, 0, $iconimageBytes.Length)
                $ims.Write($iconimageBytes, 0, $iconimageBytes.Length);
                $ims.Position = 0; # must reset position to 0 to work

                # parse message body to extract the destination folder from the JSON body
                $destFolderInfo = $message.body.content | ConvertFrom-Json
                $destinationFolderUrl = $destFolderInfo.destinationFolderURL
                Write-Output ("DestinationFolder={0}" -f $destinationFolderUrl)

                # remove unwated substrings
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
                    # we will not know which part of the URL string represents the site and which represents the library/folder. Split on the '/' 
                    # and start connecting using the most URL parts and shift site url to folder url parts each iteration until there is
                    # success in connecting.

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
                        #Write-Output ("Attempting to save file: {0}; to site: {1}; in folder: {2}" -f $attachmentInfo.name, $siteUrl, $folderUrl)
                        Connect-PnPOnline -Url $siteUrl -AccessToken $spToken -WarningAction SilentlyContinue
                        
                        $web = Get-PnPWeb -ErrorAction SilentlyContinue
                        if ($web -eq $null) {
                            # Write-Error ("Unable to connect to site: {0}" -f $siteUrl)
                            throw New-Object System.Exception
                        }
                        #Write-Output "Connected to site $($siteUrl)"
                        $errorVar = $null
                        $f = Add-PnPFile -FileName $attachmentInfo.name -Folder $folderUrl -Stream $ims -ErrorVariable errorVar
                        if ($f) {
                            #Write-Host "SUCCESS" -ForegroundColor Green
                        } else {
                            $errorObj = New-Object System.EntryPointNotFoundException
                            throw $errorObj
                        }
                        $continueLoop = $false
                        Disconnect-PnPOnline
                    }
                    catch [System.EntryPointNotFoundException] {
                        throw
                    }
                    catch {
                        #Write-Host "FAILED" -ForegroundColor Yellow
                        $loopCounter = $loopCounter - 1
                        if ($loopCounter -le 3) {
                            $continueLoop = $false
                            throw
                        }
                    }
                } until ($continueLoop -eq $false) 

                    
            }
            # move the messages to the archive mailbox folder
            $result = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.us/v1.0/users/$($inboxUpn)/mailFolders/$($message.parentFolderId)/messages/$($message.id)/move" -Headers $requestHeader -Body "{ 'destinationId':'$($archiveFolder.id)' }"
            Write-Output "Result=success"
        }
        catch {
            Write-Output ("Error={0}" -f $_.Exception.Message.Replace("`r`n", " "))
            # move the messages to the failed mailbox folder
            $result = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.us/v1.0/users/$($inboxUpn)/mailFolders/$($message.parentFolderId)/messages/$($message.id)/move" -Headers $requestHeader -Body "{ 'destinationId':'$($failedFolder.id)' }"
            Write-Output "Result=fail"
        }
    }
}

#endregion

#region Main

try {
    # Load Mailbox folders
    Get-MailFolders

    if (-not($queueFolder -and $successFolder -and $failedFolder)) {
        throw ("Unable to detect one of the mailbox folders: queueFolder: {0};, successFolder: {1}, failedFolder {2}" -f $queueFolder, $successFolder, $failedFolder)
    }

    # Get first batch of messages from the mailbox folder
    $messages = Invoke-RestMethod -Method Get -Uri ("{0}/v1.0/users/{1}/mailFolders/{2}/messages?`$select=id,body,sender,parentFolderId,receivedDateTime,hasAttachments,attachments&`$expand=attachments&`$orderby=receivedDateTime asc&`$top={3}" -f $config.exSource.graphBase, $config.exSource.inboxUpn, $queueFolder.id, $config.batch.batchSize) -Headers (Get-GraphAuthHeader)
    Write-Output ("found {0} messages" -f $queueFolder.totalItemCount)
    $counter = 0
    $continue = $false
    # Process items in first batch
    do {
        $Jobs = @()
        foreach ($msg in $messages.value) {
            $counter = $counter + 1
            $counter
            $msg.subject
            $msg.receivedDateTime
            $Jobs += Start-Job -ScriptBlock $ScriptBlock -ArgumentList $msg, $config.exSource.inboxUpn, (Get-GraphAuthHeader), (Get-SharePointToken), $successFolder, $failedFolder
        }

        while ($Jobs -ne $null -and ($Jobs |? { $_.State -eq "Running" }).Count -gt 0) {
            $Jobs |? { $_.State -eq "Running" } |% {
                if ((New-TimeSpan -Start $_.PSBeginTime -End (Get-Date)).TotalSeconds -ge $config.batch.maxJobTime) {
                    Stop-Job -Job $_
                }
            }
            Start-Sleep -Seconds $config.batch.waitTime
        }

        foreach ($job in $Jobs) {
            $logInfo = $newLogEntry.Clone()
            $errorVar = $null
            $jobInfoObj = Receive-Job -Job $job -ErrorVariable errorVar
            $jobInfo = $jobInfoObj -join "`r`n"
            $logInfo.Sender = Extract-PropertyValue -inputStr $jobInfo -key "Sender"
            $logInfo.DestinationFolder = Extract-PropertyValue -inputStr $jobInfo -key "DestinationFolder"
            $logInfo.AttachmentName = Extract-PropertyValue -inputStr $jobInfo -key "AttachmentName"
            $logInfo.ReceivedDate = Extract-PropertyValue -inputStr $jobInfo -key "ReceivedDate"
            $logInfo.Status = Extract-PropertyValue -inputStr $jobInfo -key "Result"
            $logInfo.Notes = [String]::Concat((Extract-PropertyValue -inputStr $jobInfo -key "Error"), (($errorVar |% {$_.Exception.Message.Replace("`r`n", " ") }) -join "; "))
            Write-LogFile -logEntry $logInfo
        }

        if ($messages.'@odata.nextLink') {
            $messages = Invoke-RestMethod -Method Get -Uri $messages.'@odata.nextLink' -Headers (Get-GraphAuthHeader)
            $continue = $true
        } else {
            $continue = $false
        }
    }
    until ($continue -eq $false)
}
catch {
    $logInfo = $newLogEntry.Clone()
    $logInfo.ReceivedDate = (Get-Date).ToString("MM-dd-yyyyTHH-mm-ss")
    $logInfo.Notes = $_
    Write-LogFile -logEntry $logInfo
}
finally {
    Write-Output "Cleaning up old log files"
    Clean-LogFiles
}

#endregion
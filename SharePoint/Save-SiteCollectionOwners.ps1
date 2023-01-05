# This script exports a list of site owners for all sites and uploads it to a SharePoint site

$tenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx"
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx"
$certThumb = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx"
$spUrl = "https://mytenant.sharepoint.com"
$csvSiteUrl = "https://mytenant.sharepoint.com"

Connect-PnPOnline -Url $spUrl -Tenant $tenantId -ClientId $clientId -Thumbprint $certThumb
Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certThumb
$sites = Get-PnPTenantSite

$siteOwners = @()
$counter = 1

foreach ($site in $sites) {
    Write-Progress -Activity $site.Url -PercentComplete (($counter++ / $sites.Count) * 100)
    Connect-PnPOnline -Url $site.Url -Tenant $tenantId -ClientId $clientId -Thumbprint $certThumb
    $siteAdmins = Get-PnPSiteCollectionAdmin

    foreach ($owner in $siteAdmins) {

        switch ($owner.PrincipalType) {
            "User" {
                $siteOwners += @{
                    "SiteName" = $site.Title;
                    "Url" = $site.Url;
                    "DisplayName" = $owner.Title;
                    "LoginName" = $owner.LoginName;
                    "Email" = $owner.Email;
                }
            }

            "SecurityGroup" {
                $loginName = $owner.LoginName.Split("|")
                if ($loginName[0] -eq "c:0o.c") {
                    $groupId = $loginName[2].Split("_")
                    $groupOwners = Get-MgGroupOwner -GroupId $groupId[0] |ForEach-Object { Get-MgUser -UserId $_.Id }
                    foreach ($groupOwner in $groupOwners) {
                        $siteOwners += @{
                            "SiteName" = $site.Title;
                            "Url" = $site.Url;
                            "DisplayName" = $groupOwner.DisplayName;
                            "LoginName" = $owner.LoginName;
                            "Email" = $groupOwner.UserPrincipalName;
                        }                        
                    }
                } else {
                    $siteOwners += @{
                        "SiteName" = $site.Title;
                        "Url" = $site.Url;
                        "DisplayName" = $owner.Title;
                        "LoginName" = $owner.LoginName;
                        "Email" = $owner.Email;
                    }                    
                }
            }

            Default {
                $owner
            }
        }
    }

    Disconnect-PnPOnline
}

Write-Host "Uploading file ..." 
$filename = Join-Path $env:TEMP "siteOwners-$((Get-Date).ToString('MM-dd-yy-hh-mm-ss')).csv"
$siteOwners | ForEach-Object { New-Object PSObject -Property $_ } | Export-Csv -Path $filename -NoTypeInformation -Force

# replace path below where you would like to upload the file
Connect-PnPOnline -Url $csvSiteUrl -Tenant $tenantId -ClientId $clientId -Thumbprint $certThumb
Add-PnPFile -Path $filename -Folder "Shared Documents" -NewFileName SiteOwnersReport.csv
Remove-Item -Path $filename
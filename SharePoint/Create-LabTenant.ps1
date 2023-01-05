
# -------------------------------------------------------------------------------
# Variable Section - IMPORTANT: Only modify values within this section
# -------------------------------------------------------------------------------
$tenantName       = Read-Host "Please enter the tenant name (e.g. MOD123456x)"
$admin            = Read-Host "Please enter the admin account (e.g. admin@MOD123456x.onmicrosoft.com)"
[int]$numOfSites  = Read-Host "Please, enter how many users/sites you'd like to create"
[int]$licenseType = Read-Host "`nWhat type of license would you like to apply to the new accounts?`nPlease enter either: `n`t1 - for E3 Licenses`n`t2 - for E5 Licenses`n->"
# -------------------------------------------------------------------------------
# END of Variable Section
# -------------------------------------------------------------------------------

# Setting initial values based on user input
$currentPath      = $PSScriptRoot
$adminCenterURL   = "https://$tenantName-admin.sharepoint.com"

# Import AzureAD PowerShell Module
Import-Module AzureAD 

# Store credentials
$Global:cred = Get-Credential -UserName $admin -Message "Please, enter the password"

# Connect to Azure to check license information
Connect-AzureAD -Credential $Global:cred | Out-Null

switch ($licenseType){
    1 {$accountSku = "ENTERPRISEPACK"; break}
    2 {$accountSku = "ENTERPRISEPREMIUM"; break}
    default 
    {
        Write-Host "Please, enter a valid license type (1 or 2)" -ForegroundColor Red
        exit
    }
}

$skuLicense = Get-AzureADSubscribedSku | ? {$_.SkuPartNumber -eq $accountSku} | Select -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits
$availableLicenses = $skuLicense.Enabled - $skuLicense.ConsumedUnits

# Script will end if not enough licenses
#if ($availableLicenses -lt $numOfSites)
#{
#    Write-Host "You don't have enough licenses of type '$accountSku' in your tenant. Please remove licenses from other users and try again" -ForegroundColor Red
#    Write-Host "Available licenses: "$availableLicenses -ForegroundColor Red
#    Write-Host ""
#    $skuLicense
#    
#    Exit
#}

# Create array of users and sites to be created
#$index = 1..(1+$numOfSites-1)
#[String[]]$newSiteUrls = @()
[String[]]$newUPNs = @()
#$index | % {$newSiteUrls += "https://$tenantName.sharepoint.com/sites/ContosoElectronics"+$_}
#$index | % {$newUPNs += "user"+$_+"@$tenantName.onmicrosoft.com"}
#Write-Host "`n`n`n`n`n`n`n`n`n`nThe following users will be created:" -ForegroundColor Magenta
$newUPNs

# Create users and assign licenses
$currentUpnCount = 1
Write-Host "`nCreating users in tenant" -ForegroundColor Cyan

foreach ($newUPN in $newUPNs)
{
    $percentUsersComplete = ($currentUpnCount/$numOfSites)*100
    Write-Progress -Activity "Provisioning & Assigning Licenses to users ($currentUpnCount of $numOfSites users)" `
     -Status "Provisioning user $newUPN" -Id 1 -PercentComplete $percentUsersComplete
    
    Write-Host "`tUser '$newUPN'"
    try{
        # Create users
        Write-Host "`tCreating user" -ForegroundColor Gray
        $passwordProfile = New-Object Microsoft.Open.AzureAD.Model.PasswordProfile
        $passwordProfile.Password = "H4ck4th0n"
        $user = New-AzureADUser -DisplayName "User $currentUpnCount" `
         -PasswordProfile $passwordProfile `
         -UserPrincipalName "user$currentUpnCount@$tenantName.onmicrosoft.com" `
         -AccountEnabled $true `
         -MailNickName "user$currentUpnCount" `
         -UsageLocation "US"
        
        # Assign licenses
        Write-Host "`tAssigning License ('"$skuLicense.SkuPartNumber"')" -ForegroundColor Gray
        $license = New-Object Microsoft.Open.AzureAD.Model.AssignedLicense
        $license.SkuId = $skuLicense.SkuId
        $licenseToAssign = New-Object Microsoft.Open.AzureAD.Model.AssignedLicenses
        $licenseToAssign.AddLicenses = $license
        Set-AzureADUserLicense -ObjectId $user.ObjectId -AssignedLicenses $licenseToAssign

        # Assign admin as manager in Azure AD
        Write-Host "`tAssigning $admin as manager in Azure AD" -ForegroundColor Gray
        $manager = Get-AzureADUser -ObjectId $admin
        Set-AzureADUserManager -ObjectId $newUPN -RefObjectId $manager.ObjectId
        Write-Host "`tSuccess!`n" -ForegroundColor Green

    }
    catch
    {
        write-host "`tError: $($_.Exception.Message)`n" -foregroundcolor Red
    }
    finally
    {
        $currentUpnCount++
    }
}

# Connecting to SPO
Write-Host "`nConnecting to SPO" -ForegroundColor Cyan
Connect-SPOService -Url $adminCenterURL -Credential $Global:cred

# Loading Site Script
Write-Host "`nAdding Site Script and Site Design to Tenant" -ForegroundColor Cyan
$siteScriptPath = "$currentPath\SiteScript.txt"
Write-Host "`tFetching Site Script from '$siteScriptPath'" -ForegroundColor Gray
$siteScriptContent = Get-Content -LiteralPath $siteScriptPath -Raw

# Creating Site Script and Site Design
Write-Host "`tAdding Site Script to tenant" -ForegroundColor Gray
$siteScript = Add-SPOSiteScript -Title "Workflow To Power Automate - Hackathon" -Content $siteScriptContent -Description "Workflow To Power Automate - Hackathon"
Write-Host "`tAdding Site Design to tenant" -ForegroundColor Gray
$siteDesign = Add-SPOSiteDesign -Title "Workflow To Power Automate - Hackathon" -WebTemplate 64 -SiteScripts $siteScript.Id
Write-Host "`tSuccess!`n" -ForegroundColor Green

# Create site collections, apply template and set up user as admin
Write-Host "`n`nThe following sites will be created:" -ForegroundColor Magenta
$newSiteUrls

$currentSiteCount = 1
Write-Host "`nCreating site collections, applying site script and adding admin" -ForegroundColor Cyan
foreach ($newSiteUrl in $newSiteUrls) 
{   
    try{
        $percentSitesComplete = ($currentSiteCount/$numOfSites)*100
        $newSiteTitle     = "Contoso Electronics $currentSiteCount"
        Write-Progress -Activity "Provisioning & Configuring Site Collections" -Status "$currentSiteCount of $numOfSites sites" -Id 1 -PercentComplete $percentSitesComplete
        Write-Host "*** Site Collection: '$newSiteTitle' - $newSiteUrl" -ForegroundColor Cyan
    
        # Create site collection
        Write-Host "`tCreating Site Collection" -ForegroundColor Gray
        New-SPOSite -Url $newSiteUrl -Owner $admin -Template STS#3 -StorageQuota 1000 -Title $newSiteTitle
    
        # Apply Site Design
        Write-Host "`tApplying Site Design" -ForegroundColor Gray
        Invoke-SPOSiteDesign -Identity $siteDesign.id -WebUrl $newSiteUrl | Out-Null
        Write-Host "`tSuccess!`n" -ForegroundColor Green
    }
    catch
    {
        write-host "`tError: $($_.Exception.Message)`n" -foregroundcolor Red
    }
    finally
    {
        $currentSiteCount ++
    }
}


Disconnect-AzureAD
Disconnect-SPOService


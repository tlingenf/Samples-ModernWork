Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Users

Select-MgProfile -Name "v1.0"
Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All"

$skuId = Get-MgSubscribedSku |? { $_.SkuPartNumber -eq "ENTERPRISEPREMIUM_FACULTY" }

$foundUsers = Get-MgUser -All:$true -Filter "accountEnabled eq false" -Property "AssignedLicenses","UserPrincipalName" |? { $_.AssignedLicenses.SkuId -eq $skuId.SkuId }

$foundUsers | Select-Object "UserPrincipalName" | Export-Csv -Path 'C:\Users\admin-emmcneal\Desktop\Reports\DisabledLicensedUsers.csv' -NoTypeInformation -Force

#website to get list of licenses https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
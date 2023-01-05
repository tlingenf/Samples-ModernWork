Import-Module -Name Microsoft.PowerApps.Administration.PowerShell  
Import-Module -Name Microsoft.PowerApps.PowerShell 

Add-PowerAppsAccount

$environmentId = "Default-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

$spResources = @()

$environmentFlows = Get-AdminFlow -EnvironmentName $environmentId

foreach ($flow in $environmentFlows) {
    try {
        $flow.Internal.properties.referencedResources `
            | Where-Object { $_.service -eq "sharepoint"} `
                | ForEach-Object {

                    $spResources += ${
                        "FlowId" = $flow.FlowName;
                        "FlowName" = $flow.DisplayName;
                        "EnvironmentId" = $environmentId;
                        "SPSiteUrl" = $_.resource.site;
                        "SPListId" = $_.resource.list;
                    }
                }
    }
    catch {
        $_
    }
}

$spResources | Export-Csv -Path ("SharePointFlows_{0}.csv" -f (Get-Date -Format "yyyy-MM-ddThh-mm-ss"))
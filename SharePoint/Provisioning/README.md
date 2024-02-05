# SharePoint Advanced Provisioning Demo

## Overview

This solution provides an example of a SharePoint site design that will call a Power Automate Flow to configure a site. The solution uses a SharePoint Site Design with Site Script to invoke a Power Automate Flow which will perform actions against the site and will call an Azure Function to apply a PnP provisioning template.

![PnP provisioning customization using Power Automate](https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/images/process-for-triggering-a-custom-flow.png)

## Things You Will Need

1. Create an app registration that will be used to add/remove site collection administrators for the new site.
- Grant the Sites.FullControl.All in the application scope for the SharePoint API.
- Create a certificate and add the .cer file to the app registration.
2. Create an Azure Key Vault
- upload the .pfx certificate that matches the .cer file used in step 1.
- Add the Microsoft Flow Service to an access policy for the Key Vault
3. Create an Azure Function using PowerShell
- Add a queue named *applypnpsitetemplate* to the storage account associated with the function.
- Create a system managed identity for the function app.
- Grant the enterprise app SharePoint application permissions for User.ReadWrite.All, TermStore.ReadWrite.All, and Sites.FullControl.All
4. A Power Platform environment to deploy the deployment tools solution and to create the flow that will be called from the site script.
- Ensure that the Dataverse database will be created.
5. A SharePoint site that will hold the PnP template files.


## Solution Components

- _Folder: spo-mgmt-func_ - source files to deploy to the Azure Function
- _SharePointSiteTools_1_0_0_2.zip_ - Power Platform solution with a set of reusable tasks that can be called from other flows.
- _CustomerProjectSiteCustomizations_1_0_0_1.zip_ - sample flow that will be called by the site script to perform provisioning tasks on the site.
- _SiteScript.json_ - sample site script that will call a Power Automate Flow

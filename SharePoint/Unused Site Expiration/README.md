

# Entra ID App Registration

To integrate Entra ID functionality into your application, you need to register an Entra ID App and configure the necessary API permissions. This documentation will guide you through the process.

### Prerequisites

Before you begin, make sure you have the following:

- A valid Microsoft Azure account with sufficient permissions to register an application.
- Global administrator access to the target Azure AD tenant.

### App Registration

1. Sign in to the [Entra ID portal](https://entra.microsoft.com) using your Administrative account credentials.
2. Navigate to the Applications section.
3. Select "App registrations" from the left-hand menu.
4. Click on the "New registration" button to create a new app registration.
5. Provide a name for your Entra ID App and select the appropriate account type.
6. Click on the "Register" button to create the app registration.

### API Permissions

To enable the necessary functionality, you need to configure the following API permissions for your Entra ID App:

- Microsoft Graph: Group.ReadWrite.All - Application
- Microsoft Graph: GroupMember.Read.All - Application
- SharePoint: Sites.FullControl.All - Application
- SharePoint: User.ReadWrite.All - Application

### Admin Consent

To grant the required permissions, a global administrator needs to provide admin consent for the Entra ID App. This can be done by following these steps:

1. In the Azure portal, go to the "API permissions" section of your Entra ID App registration.
2. Click on the "Grant admin consent" button.
3. A global administrator will be prompted to review and consent to the requested permissions.

### Certificate Provisioning

To enhance the security of your Entra ID App, it is necessary to provide a certificate during the app registration process. This certificate will be used for authentication and authorization purposes.

Please refer to the Azure documentation for detailed instructions on how to generate and upload a certificate for your app registration.

# Setting up Azure Key Vault

To establish the Azure Key Vault and configure the 'Entra App Cert' environment variable, follow these steps:

1. Sign in to the Azure portal using your administrative account credentials.
2. Navigate to the Azure Key Vault service.
3. Click on the "Add" button to create a new Key Vault.
4. Provide a name for your Key Vault and select the appropriate subscription, resource group, and region.
5. Configure the access policies to grant necessary permissions to your Entra ID App.
6. Click on the "Review + Create" button to create the Key Vault.

## Uploading the Certificate

Once the Key Vault is set up, you can upload the certificate for your Entra ID App by following these steps:

1. Open the Key Vault in the Azure portal.
2. Navigate to the "Certificates" section.
3. Click on the "Generate/Import" button to upload a new certificate.
4. Provide a name for the certificate and select the appropriate options for importing or generating the certificate.
5. Follow the instructions to upload the certificate to the Key Vault.

## Configuring the Environment Variable

To configure the 'Entra App Cert' environment variable, perform the following steps:

1. Open the Power Platform environment where your Entra ID App is deployed.
2. Navigate to the environment settings.
3. Locate the 'Entra App Cert' environment variable configuration.
4. Set the value of the environment variable to the secret identifier of the certificate stored in the Azure Key Vault.

Make sure to save the changes and test the configuration to ensure that the certificate is successfully retrieved from the Azure Key Vault.

For more detailed instructions on setting up Azure Key Vault and configuring environment variables, refer to the Azure documentation.

# Importing the Power Platform Solution

To import the Power Platform solution from this repository, follow these steps:

1. Download the solution file from the repository. The solution file is typically a zip file.
2. Open the Power Platform portal and sign in using your administrative account credentials.
3. Navigate to the Solutions section.
4. Click on the "Import" button to start the solution import process.
5. In the import dialog, click on the "Browse" button and select the downloaded solution file.
6. Review the solution details and make any necessary changes or configurations.
7. Click on the "Next" button to proceed with the import.
8. Select the desired options for solution components and dependencies.
9. Click on the "Import" button to start the import process.
10. Wait for the import process to complete. This may take some time depending on the size of the solution.
11. Once the import is successful, you will receive a confirmation message.

Make sure to test the imported solution to ensure that it functions as expected.

For more detailed instructions on importing Power Platform solutions, refer to the Power Platform documentation.

# Issue Turning On Flows

 The flows are not turned on during the import process. Because of the recursive nature of the flows there is currently an issue in power platform preventing flows from turning on when used recursively. 

To enable the flows, you will need to follow these steps:

1. Edit the recursive flows: Open the Power Platform portal and navigate to the flows that are used recursively. Edit each of these flows.

2. Copy the "Call Child Flow" activities: Within each recursive flow, locate the "Call Child Flow" activities and copy them.

3. Remove the "Call Child Flow" activities: After copying the "Call Child Flow" activities, remove them from the recursive flows.

4. Save the flow: Save the changes made to the recursive flows.

5. Go to the details page: Navigate to the details page of each recursive flow.

6. Turn on the flow: On the details page, locate the option to turn on the flow and enable it.

7. Edit the flow: After turning on the flow, go back to the flow editor and open the recursive flow again.

8. Paste the copied "Run a Child Flow" action: Within the recursive flow, paste the previously copied "Run a Child Flow" action back into its original position.

9. Turn on the other flows: Once you have completed the above steps for all the recursive flows, you can proceed to turn on the other flows in the Power Platform.

By following these steps, you should be able to enable the flows in the Power Platform solution even if they are used recursively.

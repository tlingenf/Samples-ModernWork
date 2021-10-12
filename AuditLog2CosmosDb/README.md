# Office 365 Audit Log Extraction

There is not a direct way to connect Power BI reports to Office 365 Audit 
data. To achieve this we can utilize the Office 365 Management API to 
extract audit data and persist it to long term storage. 

## Solution Overview

This solution subscribes to webhook notifications from the Office 365 Management API and saves the information to an Azure Storage Queue. A second function is triggered by the Azure Storage Queue to download the audit data from the Office 365 Management API after authentication and to add each record to Azure CosmosDb.

![solution diagram](Solution%20Diagram.png)

## Available Functions

- **managementApiReceiver** - webhook listener to post notification to a queue.
- **processAuditLogs** - queue trigger to authenticate, download, and insert audit data to CosmosDb.
- **StartAPISubscription** - run manually to start all subscriptions
- **StopAPISubscription** - run manually to stop all subscriptions

# Configuration

## App Settings

```javascript
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "AzureWebJobsStorage": "DefaultEndpointsProtocol=https;AccountName={{StorageAccountName}};AccountKey={{StorageAccountKey}}",
    "FUNCTIONS_EXTENSION_VERSION": "~2",
    "WEBSITE_NODE_DEFAULT_VERSION": "10.14.1",
    "tenantId": "{{TenantId}}",
    "clientId": "{{ClientId}}",
    "clientSecret": "{{ClientSecret}}",
    "Auditdb_DOCUMENTDB": "AccountEndpoint=https://{{CosmosDbName}}.documents.azure.com:443/;AccountKey={{CosmosDbKey}};"
  },
  "ConnectionStrings": {}
}
```

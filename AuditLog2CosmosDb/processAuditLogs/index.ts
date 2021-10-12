import { AzureFunction, Context } from "@azure/functions";
import * as msal from "@azure/msal-node";
import { CosmosClient } from "@azure/cosmos";
import axios from "axios";

const queueTrigger: AzureFunction = async function (context: Context, myQueueItem: any): Promise<void> {
    context.log('Queue trigger function processed work item', myQueueItem);
    try {
        const cca = new msal.ConfidentialClientApplication({
            auth: {
                clientId: process.env.clientId,
                authority: process.env.AadEndpoint + process.env.tenantId,
                clientSecret: process.env.clientSecret        
            }
        });
        const auth = await cca.acquireTokenByClientCredential({scopes: [process.env.ManagementApiUri + "/.default"]});
        
        const headerOptions = {
            'headers': {
                'Content-Type': 'application/json; utf-8',
                'Authorization': 'Bearer ' + auth.accessToken
            }
        };

        const client = new CosmosClient(process.env["Auditdb_DOCUMENTDB"]);
        const { database } = await client.databases.createIfNotExists({ id: "O365AuditData" });
        const { container } = await database.containers.createIfNotExists({ id: "AuditEvents" });

        context.log('Retrieving blob content', myQueueItem.contentUri);
        const getDocumentResponse = await axios.get(`${myQueueItem.contentUri}?PublisherIdentifier=${process.env.tenantId}`, headerOptions);
        context.log('Response code: ', getDocumentResponse.status);

        await getDocumentResponse.data.map(async (itemDef: any) => {
            try {
                context.log('Saving to cosmos db', itemDef)
                await container.items.create(itemDef);
            }
            catch (error) {
                context.log('An error occured writing to cosmos db', error);
            }
        });
    }
    catch (error) {
        console.log(error);
    }
};

export default queueTrigger;
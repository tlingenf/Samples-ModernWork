import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import * as msal from "@azure/msal-node";
import axios from "axios";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
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

        let subscritpionNames: string[] = ['Audit.AzureActiveDirectory', 'Audit.Exchange','Audit.SharePoint','Audit.General','DLP.All'];
        for (let subName of subscritpionNames) {
            context.log('Stopping audit subscription ', subName);
            let createResponse = await axios.post(`${process.env.ManagementApiUri}/api/v1.0/${process.env.tenantId}/activity/feed/subscriptions/stop?contentType=${subName}&PublisherIdentifier=${process.env.tenantId}`, {}, headerOptions);
            console.log(`Stop ${subName} response: ${createResponse.status}`);
        }
    }
    catch (error) {
        console.log(error);
    }
};

export default httpTrigger;

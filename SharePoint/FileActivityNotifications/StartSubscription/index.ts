import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import axios from "axios";
import { env } from "process";
const msal = require('@azure/msal-node');

const msalConfig = {
	auth: {
		clientId: process.env.ClientId,
		authority: process.env.AadEndpoint + process.env.TenantId,
		clientSecret: process.env.ClientSecret,
	}
};
const cca = new msal.ConfidentialClientApplication(msalConfig);
async function getToken(tokenRequest) {
	return await cca.acquireTokenByClientCredential(tokenRequest);
}

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    try {
        const authResponse = await getToken({scopes: ["https://manage.office.com/.default"]});

        const headerOptions = {
            headers: {
                'Content-Type': 'application/json; utf-8',
                'Authorization': 'Bearer ' + authResponse.accessToken
            }
        };
        const bodyData = JSON.stringify({
            'webhook': {
                'address': env.WebhookNotificationUrl,
                'authId': env.ClientId
            }            
        });

        context.log('Starting audit subscription for SharePoint');
        let statusCode = 500;

        const createResponse = await axios.post(`https://manage.office.com/api/v1.0/${process.env.TenantId}/activity/feed/subscriptions/start?contentType=Audit.SharePoint&PublisherIdentifier=${process.env.TenantId}`, bodyData, headerOptions);
        statusCode = createResponse.status;

        context.res = {
            status: statusCode
        };
    }
    catch (error) {
        console.log(error);
        context.res = {
            status: 500
        };
    }
};

export default httpTrigger;
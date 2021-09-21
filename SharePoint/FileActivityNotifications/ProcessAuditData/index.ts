import { AzureFunction, Context } from "@azure/functions";
const msal = require('@azure/msal-node');
import axios from "axios";
import template from './emailTemplate';

const msalConfig = {
	auth: {
		clientId: process.env.ClientId,
		authority: process.env.AadEndpoint + process.env.TenantId,
		clientSecret: process.env.ClientSecret
	}
};
const cca = new msal.ConfidentialClientApplication(msalConfig);
async function getToken(tokenRequest) {
	return await cca.acquireTokenByClientCredential(tokenRequest);
}

const queueTrigger: AzureFunction = async function (context: Context, myQueueItem: any): Promise<void> {
    context.log('Queue trigger function processed work item', myQueueItem);

    const fileOperationsArray = process.env.FileOperations.split(',');
    try {
        const authResponse = await getToken({ scopes: ["https://manage.office.com/.default"] });
        const headerOptions = {
            headers: {
                'Content-Type': 'application/json; utf-8',
                'Authorization': 'Bearer ' + authResponse.accessToken
            }
        };

        context.log('Retrieving blob content', myQueueItem.contentUri);

        const getDocumentResponse = await axios.get(`${myQueueItem.contentUri}?PublisherIdentifier=${process.env.TenantId}`, headerOptions);
        context.log('Response code: ', getDocumentResponse.status);

        //const responsJson = JSON.parse(getDocumentResponse.data);

        const graphAuthResponse = await getToken({ scopes: ["https://graph.microsoft.com/.default"] });
        const graphHeaderOptions = {
            headers: {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + graphAuthResponse.accessToken
            }
        };
        let siteRegEx: RegExp = new RegExp(process.env.SiteUrlRegex);

        await getDocumentResponse.data.map(async (itemDef: any) => {
            if (fileOperationsArray.findIndex(x => x === itemDef.Operation) >= 0) {
                if (siteRegEx.test(itemDef.SiteUrl)) {
                    let msgFromTemplate: string = template;
                    msgFromTemplate = msgFromTemplate.replace("{{creationTime}}", itemDef.CreationTime);
                    msgFromTemplate = msgFromTemplate.replace("{{fileName}}", itemDef.ObjectId);
                    msgFromTemplate = msgFromTemplate.replace("{{Username}}", itemDef.UserId);
                    msgFromTemplate = msgFromTemplate.replace("{{Operation}}", itemDef.Operation);

                    let messageBody = {
                        "message": {
                            "subject": "File Audit Event",
                            "body": {
                                "contentType": "HTML",
                                "content": msgFromTemplate
                            },
                            "toRecipients": [
                                {
                                    "emailAddress": {
                                        "address": process.env.NotificationAddress
                                    }
                                }
                            ]    
                        }
                    };
                    axios.post(`https://graph.microsoft.com/v1.0/users/${process.env.FromMailbox}/sendMail`, messageBody, graphHeaderOptions)
                    .then(messageCreateResponse => {
                        console.log(`Sent message ${messageCreateResponse.data.id}`);
                    })
                    .catch(error => {
                        console.log(error);
                    });
                }
            }
        });
    }
    catch (error) {
        console.log(error);
    }
};

export default queueTrigger;

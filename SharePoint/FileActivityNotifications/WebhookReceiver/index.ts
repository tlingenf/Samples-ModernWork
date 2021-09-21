import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { env } from "process";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    let reqId: string;
    if (req && req.body && req.body.length > 0 && req.body[0].clientId) {
        reqId = req.body[0].clientId;
        context.log('clientId: ' + reqId);
        if (reqId == process.env["clientId"]) {
            context.log('forwarding contentId: ' + req.body[0].contentId);
            context.bindings.outputQueueItem = req.rawBody;
        } else {
            context.log('No clientId match. Ignore.');
        }
    } else {
        context.log('clientId not found.');
        if (req && req.headers && req.headers["webhook-authid"]) {
            if (req.headers["webhook-authid"] == env.ClientId) {
                console.log("Webhook subscription start validation passed.");
                context.res.statusCode = 200;
            } else {
                context.res.statusCode = 500;
            }
        }        
    }
};

export default httpTrigger;
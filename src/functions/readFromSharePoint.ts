import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as ExcelJS from "exceljs";

/**
 * Interface for SharePoint file request
 */
interface SharePointFileRequest {
    siteUrl: string;
    fileName: string;
    folderPath?: string;
    clientId?: string;
    clientSecret?: string;
}

/**
 * Azure Function to read Excel files from SharePoint and transform data
 * 
 * This function accepts SharePoint connection details and file path,
 * retrieves the Excel file, and performs transformations on the data
 */
export async function readFromSharePoint(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log('Processing SharePoint Excel file read request');

    try {
        // Parse request body
        const body = await request.text();
        if (!body) {
            return {
                status: 400,
                body: JSON.stringify({
                    error: 'Request body is required'
                }),
                headers: {
                    'Content-Type': 'application/json'
                }
            };
        }

        const requestData: SharePointFileRequest = JSON.parse(body);

        // Validate required fields
        if (!requestData.siteUrl || !requestData.fileName) {
            return {
                status: 400,
                body: JSON.stringify({
                    error: 'siteUrl and fileName are required fields'
                }),
                headers: {
                    'Content-Type': 'application/json'
                }
            };
        }

        // Get credentials from environment or request
        const clientId = requestData.clientId || process.env.SHAREPOINT_CLIENT_ID;
        const clientSecret = requestData.clientSecret || process.env.SHAREPOINT_CLIENT_SECRET;

        if (!clientId || !clientSecret) {
            return {
                status: 400,
                body: JSON.stringify({
                    error: 'SharePoint credentials not provided. Set SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET environment variables or include in request.'
                }),
                headers: {
                    'Content-Type': 'application/json'
                }
            };
        }

        context.log(`Attempting to read file: ${requestData.fileName} from ${requestData.siteUrl}`);

        // Note: This is a placeholder for SharePoint integration
        // In production, you would use @pnp/sp or similar library to authenticate and download the file
        // For now, we'll return instructions on how to integrate
        
        return {
            status: 200,
            body: JSON.stringify({
                message: 'SharePoint integration endpoint ready',
                instructions: {
                    description: 'To complete SharePoint integration, install @pnp/sp library',
                    steps: [
                        'Run: npm install @pnp/sp @pnp/nodejs',
                        'Implement authentication using client credentials',
                        'Download file from SharePoint document library',
                        'Pass file buffer to ExcelJS for parsing'
                    ],
                    exampleRequest: {
                        siteUrl: 'https://your-tenant.sharepoint.com/sites/your-site',
                        fileName: 'data.xlsx',
                        folderPath: '/Shared Documents',
                        clientId: 'optional-override',
                        clientSecret: 'optional-override'
                    }
                },
                requestReceived: {
                    siteUrl: requestData.siteUrl,
                    fileName: requestData.fileName,
                    folderPath: requestData.folderPath || '/'
                }
            }),
            headers: {
                'Content-Type': 'application/json'
            }
        };

    } catch (error) {
        context.error('Error processing SharePoint request:', error);
        return {
            status: 500,
            body: JSON.stringify({
                error: 'Failed to process SharePoint request',
                message: error instanceof Error ? error.message : 'Unknown error'
            }),
            headers: {
                'Content-Type': 'application/json'
            }
        };
    }
}

// Register the HTTP trigger function
app.http('readFromSharePoint', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: readFromSharePoint
});

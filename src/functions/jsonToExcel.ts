import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getFileFromGoogleDrive, uploadFileToGoogleDrive } from "../helpers/googleDrive";
import * as ExcelJS from 'exceljs';

interface RequestBody {
    fileName: string;
    data: {
        address: `${string}${number}`; // Example: "A1"
        value: string;
    }[];
}

export async function jsonToExcel(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const { fileName, data } = await request.json() as RequestBody;
    try {
        const buffer = await getFileFromGoogleDrive(fileName);
        const workbook = new ExcelJS.Workbook();
        // @ts-ignore
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
            return {
                status: 400,
                jsonBody: { success: false, error: 'No worksheet found in the Excel file.' }
            };
        }
        for (const edit of data) {
            worksheet.getCell(edit.address).value = edit.value;
        }
        const outBuffer = await workbook.xlsx.writeBuffer();
        const base64 = Buffer.from(outBuffer as ArrayBuffer).toString('base64');
        
        // Upload updated file back to Google Drive replacing the old one
        const id = await uploadFileToGoogleDrive(fileName, Buffer.from(outBuffer as ArrayBuffer), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        return { 
            status: 200, 
            jsonBody: { success: true, fileId: id }
        };
    } catch (error: any) {
        context.error(`Error: ${error.message}`);
        return { 
            status: 500, 
            jsonBody: { success: false, error: error.message }
        };
    }
};

app.http('jsonToExcel', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: jsonToExcel
});

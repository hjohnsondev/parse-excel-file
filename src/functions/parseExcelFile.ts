import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as ExcelJS from "exceljs";

/**
 * Azure Function to parse Excel files from SharePoint and transform data
 * 
 * This function accepts an Excel file buffer and performs transformations on the data
 */
export async function parseExcelFile(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log('Processing Excel file parse request');

    try {
        // Get the Excel file from the request body
        const contentType = request.headers.get('content-type');
        
        if (!contentType || !contentType.includes('application/octet-stream')) {
            return {
                status: 400,
                body: JSON.stringify({
                    error: 'Content-Type must be application/octet-stream for Excel file upload'
                }),
                headers: {
                    'Content-Type': 'application/json'
                }
            };
        }

        const fileBuffer = await request.arrayBuffer();
        
        if (!fileBuffer || fileBuffer.byteLength === 0) {
            return {
                status: 400,
                body: JSON.stringify({
                    error: 'No file data provided'
                }),
                headers: {
                    'Content-Type': 'application/json'
                }
            };
        }

        // Parse the Excel file using ExcelJS
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(Buffer.from(fileBuffer) as any);

        // Transform the data
        const transformedData = transformExcelData(workbook);

        return {
            status: 200,
            body: JSON.stringify({
                success: true,
                data: transformedData,
                worksheetCount: workbook.worksheets.length
            }),
            headers: {
                'Content-Type': 'application/json'
            }
        };

    } catch (error) {
        context.error('Error processing Excel file:', error);
        return {
            status: 500,
            body: JSON.stringify({
                error: 'Failed to process Excel file',
                message: error instanceof Error ? error.message : 'Unknown error'
            }),
            headers: {
                'Content-Type': 'application/json'
            }
        };
    }
}

/**
 * Transform Excel data into a structured format
 * @param workbook - ExcelJS Workbook instance
 * @returns Transformed data from all worksheets
 */
function transformExcelData(workbook: ExcelJS.Workbook): any {
    const result: any = {
        worksheets: []
    };

    workbook.eachSheet((worksheet, sheetId) => {
        const sheetData: any = {
            name: worksheet.name,
            rowCount: worksheet.rowCount,
            columnCount: worksheet.columnCount,
            rows: []
        };

        // Extract headers from the first row
        const headerRow = worksheet.getRow(1);
        const headers: string[] = [];
        headerRow.eachCell((cell, colNumber) => {
            headers[colNumber - 1] = cell.value?.toString() || `Column${colNumber}`;
        });

        // Transform each data row into an object
        worksheet.eachRow((row, rowNumber) => {
            // Skip header row
            if (rowNumber === 1) return;

            const rowData: any = {};
            row.eachCell((cell, colNumber) => {
                const header = headers[colNumber - 1];
                rowData[header] = cell.value;
            });

            sheetData.rows.push(rowData);
        });

        result.worksheets.push(sheetData);
    });

    return result;
}

// Register the HTTP trigger function
app.http('parseExcelFile', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: parseExcelFile
});

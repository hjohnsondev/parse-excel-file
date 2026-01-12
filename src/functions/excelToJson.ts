import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions"; 
import * as ExcelJS from 'exceljs';
import { getFileFromGoogleDrive } from "../helpers/googleDrive";

export async function excelToJson(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const { fileName } = await request.json() as any;

    try {
        const buffer = await getFileFromGoogleDrive(fileName);

        // Parse with ExcelJS
        const workbook = new ExcelJS.Workbook();
        // @ts-ignore
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.getWorksheet(1);
        if (!worksheet) {
            throw new Error('No worksheet found');
        }

        const headerMap: { [key: number]: string } = {}; 
        const nullCells: any[] = [];
        let isHeaderRow = true;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (isHeaderRow) {
                // Capture only columns that have header text
                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    headerMap[colNumber] = cell.value ? cell.value.toString() : `Column ${colNumber}`;
                });
                isHeaderRow = false;
            } else {
                const rowDataForContext: string[] = [];
                const pendingNullsInRow: any[] = [];

                // Single pass through the row to collect context and identify nulls
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    const header = headerMap[colNumber];
                    if (!header) return; // Skip columns outside the header range

                    const value = cell.value;

                    if (value === null || value === undefined || value === "") {
                        // Mark this as a target for the LLM
                        pendingNullsInRow.push({
                            address: cell.address,
                            type: cell.type,
                            header: header,
                            value: null
                        });
                    } else {
                        // Add to the context string for other cells in this row
                        rowDataForContext.push(`${header}: ${value}`);
                    }
                });

                // Create the final context string for this row
                const rowContext = rowDataForContext.join(" || ");

                // Attach the context to the null cells found in this specific row
                pendingNullsInRow.forEach(nullItem => {
                    nullCells.push({
                        ...nullItem,
                        rowContext: rowContext
                    });
                });
            }
        });

        return {
            status: 200,
            jsonBody: {
                totalGapsFound: nullCells.length,
                missingData: nullCells
            }
        };
    } catch (error: any) {
        context.error(`Error: ${error.message}`);
        return { 
            status: 500, 
            jsonBody: { success: false, error: error.message }
        };
    }
}

app.http('excelToJson', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: excelToJson
});


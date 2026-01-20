import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as ExcelJS from 'exceljs';
import { getFileFromGoogleDrive } from "../helpers/googleDrive";

interface ExcelToComprehensiveJsonRequest {
    fileName: string;
}

interface CellData {
    value: any;
    formula?: string;
    hyperlink?: string;
    richText?: any[];
    type: string;
}

interface RowData {
    rowNumber: number;
    cells: CellData[];
}

interface WorksheetData {
    name: string;
    rowCount: number;
    columnCount: number;
    rows: RowData[];
}

export async function excelToComprehensiveJson(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const { fileName } = await request.json() as any;
    
    try {
        const buffer = await getFileFromGoogleDrive(fileName);

        // Parse with ExcelJS
        const workbook = new ExcelJS.Workbook();
        // @ts-ignore
        await workbook.xlsx.load(buffer);

        const worksheets: WorksheetData[] = [];

        // Process each worksheet
        workbook.eachSheet((worksheet, sheetId) => {
            const rows: RowData[] = [];
            let maxRow = 0;
            let maxCol = 0;

            // Extract all rows and cells
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                maxRow = Math.max(maxRow, rowNumber);
                
                const rowData: RowData = {
                    rowNumber: rowNumber,
                    cells: []
                };
                
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    maxCol = Math.max(maxCol, colNumber);
                    
                    const cellData: CellData = {
                        value: null,
                        type: cell.type.toString()
                    };

                    // Handle different cell types
                    if (cell.type === ExcelJS.ValueType.Formula) {
                        cellData.formula = (cell.value as any).formula;
                        cellData.value = (cell.value as any).result;
                    } else if (cell.type === ExcelJS.ValueType.Hyperlink) {
                        cellData.hyperlink = (cell.value as any).hyperlink;
                        cellData.value = (cell.value as any).text;
                    } else if (cell.type === ExcelJS.ValueType.RichText) {
                        cellData.richText = (cell.value as any).richText;
                        cellData.value = (cell.value as any).richText
                            .map((rt: any) => rt.text)
                            .join('');
                    } else if (cell.type === ExcelJS.ValueType.Date) {
                        cellData.value = (cell.value as Date).toISOString();
                    } else {
                        cellData.value = cell.value;
                    }

                    rowData.cells.push(cellData);
                });

                rows.push(rowData);
            });

            worksheets.push({
                name: worksheet.name,
                rowCount: maxRow,
                columnCount: maxCol,
                rows: rows
            });
        });

        return {
            status: 200,
            jsonBody: {
                success: true,
                workbook: {
                    fileName: fileName,
                    sheetCount: worksheets.length,
                    worksheets: worksheets
                },
                processedAt: new Date().toISOString()
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

app.http('excelToComprehensiveJson', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: excelToComprehensiveJson
});
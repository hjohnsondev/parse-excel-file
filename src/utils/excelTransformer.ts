import * as ExcelJS from "exceljs";
import { TransformedExcelData, WorksheetData, ExcelParseOptions } from "./types";

/**
 * Utility class for Excel file transformations
 */
export class ExcelTransformer {
    /**
     * Transform an Excel workbook into a structured JSON format
     * @param workbook - ExcelJS Workbook instance
     * @param options - Parse options
     * @returns Transformed data from all worksheets
     */
    static transformWorkbook(workbook: ExcelJS.Workbook, options: ExcelParseOptions = {}): TransformedExcelData {
        const result: TransformedExcelData = {
            worksheets: []
        };

        workbook.eachSheet((worksheet, sheetId) => {
            const sheetData = this.transformWorksheet(worksheet, options);
            result.worksheets.push(sheetData);
        });

        return result;
    }

    /**
     * Transform a single worksheet into structured data
     * @param worksheet - ExcelJS Worksheet instance
     * @param options - Parse options
     * @returns Transformed worksheet data
     */
    static transformWorksheet(worksheet: ExcelJS.Worksheet, options: ExcelParseOptions = {}): WorksheetData {
        const {
            includeEmptyRows = false,
            trimValues = true,
            startRow = 1
        } = options;

        const sheetData: WorksheetData = {
            name: worksheet.name,
            rowCount: worksheet.rowCount,
            columnCount: worksheet.columnCount,
            rows: []
        };

        // Extract headers from the first row
        const headerRow = worksheet.getRow(startRow);
        const headers: string[] = [];
        headerRow.eachCell((cell, colNumber) => {
            let header = cell.value?.toString() || `Column${colNumber}`;
            if (trimValues) {
                header = header.trim();
            }
            headers[colNumber - 1] = header;
        });

        // Transform each data row into an object
        worksheet.eachRow((row, rowNumber) => {
            // Skip header row
            if (rowNumber <= startRow) return;

            const rowData: Record<string, any> = {};
            let hasData = false;

            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const header = headers[colNumber - 1];
                let value = cell.value;

                // Handle different cell value types
                if (value !== null && value !== undefined) {
                    hasData = true;
                    
                    // Handle rich text
                    if (typeof value === 'object' && 'richText' in value) {
                        value = value.richText.map((rt: any) => rt.text).join('');
                    }

                    // Trim string values if option is enabled
                    if (trimValues && typeof value === 'string') {
                        value = value.trim();
                    }
                }

                rowData[header] = value;
            });

            // Add row if it has data or if includeEmptyRows is true
            if (hasData || includeEmptyRows) {
                sheetData.rows.push(rowData);
            }
        });

        return sheetData;
    }

    /**
     * Filter rows based on a condition
     * @param data - Transformed data
     * @param predicate - Filter function
     * @returns Filtered data
     */
    static filterRows(data: TransformedExcelData, predicate: (row: Record<string, any>) => boolean): TransformedExcelData {
        return {
            worksheets: data.worksheets.map(ws => ({
                ...ws,
                rows: ws.rows.filter(predicate)
            }))
        };
    }

    /**
     * Map rows to a new structure
     * @param data - Transformed data
     * @param mapper - Mapping function
     * @returns Mapped data
     */
    static mapRows<T>(data: TransformedExcelData, mapper: (row: Record<string, any>) => T): T[][] {
        return data.worksheets.map(ws => ws.rows.map(mapper));
    }

    /**
     * Get a specific worksheet by name
     * @param data - Transformed data
     * @param worksheetName - Name of the worksheet
     * @returns Worksheet data or undefined
     */
    static getWorksheet(data: TransformedExcelData, worksheetName: string): WorksheetData | undefined {
        return data.worksheets.find(ws => ws.name === worksheetName);
    }
}

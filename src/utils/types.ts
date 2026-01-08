/**
 * Type definitions for Excel transformation
 */

export interface WorksheetData {
    name: string;
    rowCount: number;
    columnCount: number;
    rows: Record<string, any>[];
}

export interface TransformedExcelData {
    worksheets: WorksheetData[];
}

export interface ExcelParseOptions {
    includeEmptyRows?: boolean;
    trimValues?: boolean;
    startRow?: number;
}

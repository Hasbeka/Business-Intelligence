import ExcelJS from 'exceljs';

/**
 * Core utility for exporting data to Excel files
 * Supports multiple sheets, formatting, and formulas
 */

export interface ExcelColumn {
    header: string;
    key: string;
    width?: number;
    style?: Partial<ExcelJS.Style>;
}

export interface ExcelExportOptions {
    filename: string;
    sheets: ExcelSheet[];
}

export interface ExcelSheet {
    name: string;
    columns: ExcelColumn[];
    data: any[];
    summary?: SummaryRow[];
    formatters?: RowFormatter[];
}

export interface SummaryRow {
    label: string;
    value: string | number;
    style?: 'info' | 'success' | 'warning' | 'error';
}

export interface RowFormatter {
    condition: (row: any) => boolean;
    style: Partial<ExcelJS.Style>;
}

/**
 * Main export function - creates and downloads an Excel file
 */
export async function exportToExcel(options: ExcelExportOptions): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    
    // Set workbook properties
    workbook.creator = 'Wine Analytics Dashboard';
    workbook.created = new Date();
    workbook.modified = new Date();

    // Create each sheet
    for (const sheetConfig of options.sheets) {
        await createSheet(workbook, sheetConfig);
    }

    // Generate buffer and download
    const buffer = await workbook.xlsx.writeBuffer();
    downloadFile(buffer, options.filename);
}

/**
 * Creates a single worksheet with data, formatting, and summaries
 */
async function createSheet(workbook: ExcelJS.Workbook, config: ExcelSheet): Promise<void> {
    const worksheet = workbook.addWorksheet(config.name, {
        properties: { defaultRowHeight: 20 }
    });

    // Configure columns
    worksheet.columns = config.columns.map(col => ({
        header: col.header,
        key: col.key,
        width: col.width || 15,
        style: col.style
    }));

    // Style header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F1F1F' } // Black background
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 25;

    // Add data rows
    config.data.forEach((row) => {
        const excelRow = worksheet.addRow(row);
        
        // Apply conditional formatting if specified
        if (config.formatters) {
            config.formatters.forEach(formatter => {
                if (formatter.condition(row)) {
                    excelRow.eachCell((cell) => {
                        Object.assign(cell, formatter.style);
                    });
                }
            });
        }
    });

    // Apply zebra striping for better readability
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1 && rowNumber % 2 === 0) {
            row.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF5F5F5' } // Light gray
                };
            });
        }
    });

    // Add summary section if provided
    if (config.summary && config.summary.length > 0) {
        // Add empty row
        worksheet.addRow([]);
        
        // Add summary title
        const summaryTitleRow = worksheet.addRow(['Summary']);
        summaryTitleRow.font = { bold: true, size: 12 };
        summaryTitleRow.height = 25;
        
        // Add summary rows
        config.summary.forEach(summaryItem => {
            const row = worksheet.addRow([summaryItem.label, summaryItem.value]);
            row.font = { bold: true };
            
            // Apply style based on summary type
            const bgColor = getSummaryColor(summaryItem.style);
            row.eachCell((cell, colNumber) => {
                if (colNumber === 1) {
                    cell.alignment = { horizontal: 'right' };
                }
                if (colNumber === 2) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: bgColor }
                    };
                }
            });
        });
    }

    // Auto-fit columns (approximate)
    worksheet.columns.forEach((column: any) => {
        if (!column.width) {
            let maxLength = 0;
            column.eachCell?.({ includeEmpty: false }, (cell: any) => {
                const length = cell.value ? cell.value.toString().length : 10;
                maxLength = Math.max(maxLength, length);
            });
            column.width = Math.min(Math.max(maxLength + 2, 10), 50);
        }
    });

    // Freeze header row
    worksheet.views = [
        { state: 'frozen', xSplit: 0, ySplit: 1 }
    ];
}

/**
 * Get background color based on summary style
 */
function getSummaryColor(style?: string): string {
    switch (style) {
        case 'success':
            return 'FF90EE90'; // Light green
        case 'warning':
            return 'FFFFD700'; // Gold
        case 'error':
            return 'FFFF6B6B'; // Light red
        case 'info':
        default:
            return 'FF87CEEB'; // Sky blue
    }
}

/**
 * Downloads the Excel file to the user's computer
 */
function downloadFile(buffer: ArrayBuffer, filename: string): void {
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
}

/**
 * Helper: Format date to readable string
 */
export function formatDate(date: Date | string): string {
    if (typeof date === 'string') {
        return date;
    }
    return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    });
}

/**
 * Helper: Format currency
 */
export function formatCurrency(value: number): string {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(value);
}

/**
 * Helper: Format percentage
 */
export function formatPercentage(value: number, decimals: number = 1): string {
    return `${value > 0 ? '+' : ''}${value.toFixed(decimals)}%`;
}

/**
 * Helper: Generate filename with timestamp
 */
export function generateFilename(baseName: string): string {
    const timestamp = new Date().toISOString().split('T')[0];
    return `${baseName}-${timestamp}.xlsx`;
}

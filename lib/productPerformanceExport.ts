// lib/productPerformanceExport.ts
"use client";

import ExcelJS from 'exceljs';

type WineAssociation = {
    wine: string;
    wineCategory: string;
    wineCountry: string;
    totalCustomers: number;
    associations: Array<{
        associatedWine: string;
        associatedCategory: string;
        associatedCountry: string;
        count: number;
        confidence: number;
    }>;
};

type CategoryAssociation = {
    category: string;
    associations: Array<{ category: string; count: number }>;
};

type WinePerformance = {
    wine: string;
    category: string;
    country: string;
    totalSales: number;
    totalRevenue: number;
    totalQuantity: number;
    uniqueCustomers: number;
    avgSaleAmount: number;
};

export async function exportProductPerformanceToExcel(
    wineAssociations: WineAssociation[],
    categoryAssociations: CategoryAssociation[],
    winePerformance: WinePerformance[]
): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Wine Analytics Dashboard';
    workbook.created = new Date();

    // ===== SHEET 1: Overview & Summary =====
    const overviewSheet = workbook.addWorksheet('Overview', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 3 }]
    });

    // Title
    overviewSheet.mergeCells('A1:F1');
    const titleCell = overviewSheet.getCell('A1');
    titleCell.value = '🍷 Product Performance & Associations Analysis';
    titleCell.font = { size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF9333EA' } // Purple
    };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    overviewSheet.getRow(1).height = 35;

    // Date
    overviewSheet.mergeCells('A2:F2');
    const dateCell = overviewSheet.getCell('A2');
    dateCell.value = `Generated: ${new Date().toLocaleString()}`;
    dateCell.font = { size: 10, italic: true };
    dateCell.alignment = { horizontal: 'center' };

    // Key Metrics Header
    overviewSheet.getRow(4).height = 25;
    overviewSheet.getCell('A4').value = 'KEY METRICS';
    overviewSheet.getCell('A4').font = { size: 14, bold: true, color: { argb: 'FF9333EA' } };
    
    // Metrics
    const metrics = [
        ['Total Wines Analyzed', wineAssociations.length],
        ['Avg Associations per Wine', (wineAssociations.reduce((sum, w) => sum + w.associations.length, 0) / wineAssociations.length).toFixed(1)],
        ['Total Wine Performance Records', winePerformance.length],
        ['Total Revenue from Top Wines', `$${winePerformance.reduce((sum, w) => sum + w.totalRevenue, 0).toLocaleString()}`],
        ['Top Performing Wine', winePerformance[0]?.wine || 'N/A'],
        ['Top Wine Revenue', `$${winePerformance[0]?.totalRevenue.toLocaleString() || '0'}`],
    ];

    metrics.forEach((metric, idx) => {
        const row = overviewSheet.getRow(5 + idx);
        row.getCell(1).value = metric[0];
        row.getCell(1).font = { bold: true };
        row.getCell(2).value = metric[1];
        row.getCell(2).font = { size: 12 };
        row.height = 20;
    });

    overviewSheet.getColumn(1).width = 30;
    overviewSheet.getColumn(2).width = 30;

    // ===== SHEET 2: Wine Associations =====
    const associationsSheet = workbook.addWorksheet('Wine Associations', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    // Headers
    const assocHeaders = ['Wine', 'Category', 'Country', 'Total Customers', 
                          'Associated Wine', 'Associated Category', 'Associated Country', 
                          'Co-Purchases', 'Confidence %'];
    
    const assocHeaderRow = associationsSheet.getRow(1);
    assocHeaders.forEach((header, idx) => {
        const cell = assocHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF9333EA' } // Purple
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    assocHeaderRow.height = 25;

    // Data
    let currentRow = 2;
    wineAssociations.forEach((wine) => {
        if (wine.associations.length === 0) {
            // Wine with no associations
            const row = associationsSheet.getRow(currentRow);
            row.getCell(1).value = wine.wine;
            row.getCell(2).value = wine.wineCategory;
            row.getCell(3).value = wine.wineCountry;
            row.getCell(4).value = wine.totalCustomers;
            row.getCell(4).alignment = { horizontal: 'right' };
            row.getCell(5).value = 'No associations found';
            row.getCell(5).font = { italic: true };
            currentRow++;
        } else {
            // Wine with associations
            wine.associations.forEach((assoc, idx) => {
                const row = associationsSheet.getRow(currentRow);
                
                // Base wine info (only on first row)
                if (idx === 0) {
                    row.getCell(1).value = wine.wine;
                    row.getCell(2).value = wine.wineCategory;
                    row.getCell(3).value = wine.wineCountry;
                    row.getCell(4).value = wine.totalCustomers;
                    row.getCell(4).alignment = { horizontal: 'right' };
                    
                    // Bold base wine
                    row.getCell(1).font = { bold: true };
                    row.getCell(2).font = { bold: true };
                }
                
                // Associated wine info
                row.getCell(5).value = assoc.associatedWine;
                row.getCell(6).value = assoc.associatedCategory;
                row.getCell(7).value = assoc.associatedCountry;
                row.getCell(8).value = assoc.count;
                row.getCell(8).alignment = { horizontal: 'right' };
                row.getCell(9).value = assoc.confidence / 100; // As decimal for percentage format
                row.getCell(9).numFmt = '0.0%';
                row.getCell(9).alignment = { horizontal: 'right' };
                
                // Conditional formatting for confidence
                if (assoc.confidence >= 70) {
                    row.getCell(9).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF86EFAC' } // Light green
                    };
                    row.getCell(9).font = { bold: true };
                } else if (assoc.confidence >= 50) {
                    row.getCell(9).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFDE68A' } // Light yellow
                    };
                }
                
                currentRow++;
            });
        }
        
        // Add separator row
        currentRow++;
    });

    // Set column widths
    associationsSheet.getColumn(1).width = 35; // Wine
    associationsSheet.getColumn(2).width = 15; // Category
    associationsSheet.getColumn(3).width = 15; // Country
    associationsSheet.getColumn(4).width = 15; // Total Customers
    associationsSheet.getColumn(5).width = 35; // Associated Wine
    associationsSheet.getColumn(6).width = 18; // Associated Category
    associationsSheet.getColumn(7).width = 15; // Associated Country
    associationsSheet.getColumn(8).width = 15; // Co-Purchases
    associationsSheet.getColumn(9).width = 15; // Confidence %

    // ===== SHEET 3: Category Associations =====
    const categorySheet = workbook.addWorksheet('Category Patterns', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    // Headers
    const catHeaders = ['Base Category', 'Associated Category', 'Co-Purchase Count'];
    const catHeaderRow = categorySheet.getRow(1);
    catHeaders.forEach((header, idx) => {
        const cell = catHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEC4899' } // Pink
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    catHeaderRow.height = 25;

    // Data
    currentRow = 2;
    categoryAssociations.forEach((cat) => {
        cat.associations.forEach((assoc, idx) => {
            const row = categorySheet.getRow(currentRow);
            
            if (idx === 0) {
                row.getCell(1).value = cat.category;
                row.getCell(1).font = { bold: true, color: { argb: 'FFEC4899' } };
            }
            
            row.getCell(2).value = assoc.category;
            row.getCell(3).value = assoc.count;
            row.getCell(3).alignment = { horizontal: 'right' };
            
            currentRow++;
        });
        currentRow++; // Separator
    });

    categorySheet.getColumn(1).width = 25;
    categorySheet.getColumn(2).width = 25;
    categorySheet.getColumn(3).width = 20;

    // ===== SHEET 4: Top Wine Performance =====
    const performanceSheet = workbook.addWorksheet('Top Wine Performance', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    // Headers
    const perfHeaders = ['Rank', 'Wine', 'Category', 'Country', 'Total Sales', 
                         'Total Revenue', 'Total Quantity', 'Unique Customers', 'Avg Sale Amount'];
    
    const perfHeaderRow = performanceSheet.getRow(1);
    perfHeaders.forEach((header, idx) => {
        const cell = perfHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF59E0B' } // Yellow/Orange
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    perfHeaderRow.height = 25;

    // Data
    winePerformance.forEach((wine, idx) => {
        const row = performanceSheet.getRow(idx + 2);
        
        row.getCell(1).value = idx + 1;
        row.getCell(1).alignment = { horizontal: 'center' };
        
        // Medals for top 3
        if (idx === 0) {
            row.getCell(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFBBF24' } // Gold
            };
            row.getCell(1).font = { bold: true, size: 12 };
        } else if (idx === 1) {
            row.getCell(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD1D5DB' } // Silver
            };
            row.getCell(1).font = { bold: true };
        } else if (idx === 2) {
            row.getCell(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFB923C' } // Bronze
            };
            row.getCell(1).font = { bold: true };
        }
        
        row.getCell(2).value = wine.wine;
        row.getCell(2).font = { bold: idx < 3 };
        row.getCell(3).value = wine.category;
        row.getCell(4).value = wine.country;
        row.getCell(5).value = wine.totalSales;
        row.getCell(5).alignment = { horizontal: 'right' };
        row.getCell(6).value = wine.totalRevenue;
        row.getCell(6).numFmt = '$#,##0.00';
        row.getCell(6).alignment = { horizontal: 'right' };
        row.getCell(6).font = { bold: true, color: { argb: 'FF16A34A' } }; // Green
        row.getCell(7).value = wine.totalQuantity;
        row.getCell(7).alignment = { horizontal: 'right' };
        row.getCell(8).value = wine.uniqueCustomers;
        row.getCell(8).alignment = { horizontal: 'right' };
        row.getCell(9).value = wine.avgSaleAmount;
        row.getCell(9).numFmt = '$#,##0.00';
        row.getCell(9).alignment = { horizontal: 'right' };
        
        row.height = 20;
    });

    // Set column widths
    performanceSheet.getColumn(1).width = 8;  // Rank
    performanceSheet.getColumn(2).width = 35; // Wine
    performanceSheet.getColumn(3).width = 15; // Category
    performanceSheet.getColumn(4).width = 15; // Country
    performanceSheet.getColumn(5).width = 12; // Total Sales
    performanceSheet.getColumn(6).width = 15; // Total Revenue
    performanceSheet.getColumn(7).width = 15; // Total Quantity
    performanceSheet.getColumn(8).width = 15; // Unique Customers
    performanceSheet.getColumn(9).width = 15; // Avg Sale Amount

    // Add totals row
    const totalRow = performanceSheet.getRow(winePerformance.length + 3);
    totalRow.getCell(1).value = 'TOTAL';
    totalRow.getCell(1).font = { bold: true };
    totalRow.getCell(5).value = winePerformance.reduce((sum, w) => sum + w.totalSales, 0);
    totalRow.getCell(5).font = { bold: true };
    totalRow.getCell(5).alignment = { horizontal: 'right' };
    totalRow.getCell(6).value = winePerformance.reduce((sum, w) => sum + w.totalRevenue, 0);
    totalRow.getCell(6).numFmt = '$#,##0.00';
    totalRow.getCell(6).font = { bold: true, color: { argb: 'FF16A34A' } };
    totalRow.getCell(6).alignment = { horizontal: 'right' };
    totalRow.getCell(7).value = winePerformance.reduce((sum, w) => sum + w.totalQuantity, 0);
    totalRow.getCell(7).font = { bold: true };
    totalRow.getCell(7).alignment = { horizontal: 'right' };
    
    // Style total row
    totalRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEF3C7' } // Light yellow
    };
    totalRow.height = 25;

    // ===== Export File =====
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Product_Performance_Analysis_${new Date().toISOString().split('T')[0]}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
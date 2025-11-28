import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { generateComposedChartXML, generateDrawingXML, generateDrawingRelsXML } from '@/lib/export-sales-chart/chart-xml-generator';

export async function POST(request: NextRequest) {
    try {
        const body = await request.json();
        const { displayedData, dateRange, correlationInsight } = body;

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Wine Analytics Dashboard';
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet('Price vs Sales');

        // Title
        worksheet.mergeCells('A1:F1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = 'Price Changes vs Sales Evolution';
        titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F1F1F' } };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(1).height = 30;

        // Date range
        worksheet.getCell('A2').value = 'Date Range:';
        worksheet.getCell('B2').value = dateRange;
        worksheet.getCell('A2').font = { bold: true };

        // Data headers
        const dataStartRow = 4;
        const headerRow = worksheet.getRow(dataStartRow);
        headerRow.values = [
            'Month-Year',
            'Avg Price ($)',
            'Total Sales ($)',
            'Price Change %',
            'Sales Change %',
            'Price Change $',
            'Sales Change $'
        ];
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F1F1F' } };
        headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
        headerRow.height = 25;

        // Add data
        displayedData.forEach((item: any, index: number) => {
            const rowNum = dataStartRow + 1 + index;
            const row = worksheet.getRow(rowNum);

            row.values = [
                item.yearMonth,
                item.avgPrice,
                item.totalAmount,
                item.priceChangePercent,
                item.salesChangePercent,
                item.priceChange,
                item.salesChange
            ];

            row.getCell(2).numFmt = '$#,##0.00';
            row.getCell(3).numFmt = '$#,##0.00';
            row.getCell(4).numFmt = '0.0"%"';
            row.getCell(5).numFmt = '0.0"%"';
            row.getCell(6).numFmt = '$#,##0.00';
            row.getCell(7).numFmt = '$#,##0.00';

            if (index % 2 === 0) {
                row.eachCell((cell) => {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
                });
            }
        });

        worksheet.getColumn(1).width = 15;
        worksheet.getColumn(2).width = 14;
        worksheet.getColumn(3).width = 16;
        worksheet.getColumn(4).width = 15;
        worksheet.getColumn(5).width = 15;
        worksheet.getColumn(6).width = 15;
        worksheet.getColumn(7).width = 16;

        // Summary
        const summaryStartRow = dataStartRow + displayedData.length + 18;

        const summaryTitle = worksheet.getCell(`A${summaryStartRow}`);
        summaryTitle.value = 'Correlation Analysis';
        summaryTitle.font = { bold: true, size: 12 };
        summaryTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF87CEEB' } };

        worksheet.getCell(`A${summaryStartRow + 1}`).value = 'Insight:';
        worksheet.getCell(`B${summaryStartRow + 1}`).value = correlationInsight;
        worksheet.getCell(`A${summaryStartRow + 1}`).font = { bold: true };

        worksheet.getColumn(2).width = 60;
        worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: dataStartRow }];

        // Generate base Excel
        const buffer = await workbook.xlsx.writeBuffer();
        const zip = await JSZip.loadAsync(buffer);

        // Add composed chart (Bar + Line)
        const dataEndRow = dataStartRow + displayedData.length;

        const chartXML = generateComposedChartXML(
            'Price vs Sales',
            dataStartRow + 1,
            dataEndRow,
            'A',  // Month-Year
            'D',  // Price Change %
            'E',  // Sales Change %
            'Price Changes vs Sales Evolution',
            'Price Change %',
            'Sales Change %',
            'ef4444',  // Red for bars
            '3b82f6'   // Blue for line
        );

        zip.file('xl/charts/chart1.xml', chartXML);

        const drawingXML = generateDrawingXML('rId1');
        zip.file('xl/drawings/drawing1.xml', drawingXML);

        const drawingRelsXML = generateDrawingRelsXML();
        zip.file('xl/drawings/_rels/drawing1.xml.rels', drawingRelsXML);

        // Update worksheet to reference drawing
        const worksheetXML = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
        if (worksheetXML) {
            const updatedWorksheetXML = worksheetXML.replace(
                '</worksheet>',
                '<drawing r:id="rId99"/></worksheet>'
            );
            zip.file('xl/worksheets/sheet1.xml', updatedWorksheetXML);
        }

        // Update worksheet rels
        let worksheetRels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('string');
        if (!worksheetRels) {
            worksheetRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        }
        const updatedWorksheetRels = worksheetRels.replace(
            '</Relationships>',
            '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'
        );
        zip.file('xl/worksheets/_rels/sheet1.xml.rels', updatedWorksheetRels);

        // Update [Content_Types].xml
        const contentTypes = await zip.file('[Content_Types].xml')?.async('string');
        if (contentTypes) {
            let updated = contentTypes;
            if (!updated.includes('drawingml.chart')) {
                updated = updated.replace(
                    '</Types>',
                    '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'
                );
            }
            if (!updated.includes('spreadsheetDrawing')) {
                updated = updated.replace(
                    '</Types>',
                    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
                );
            }
            zip.file('[Content_Types].xml', updated);
        }

        // Generate final buffer
        const finalBuffer = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 9 }
        });

        return new NextResponse(Uint8Array.from(finalBuffer), {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': `attachment; filename="price-vs-sales-${new Date().toISOString().split('T')[0]}.xlsx"`
            }
        });

    } catch (error) {
        console.error('Export error:', error);
        return NextResponse.json({ error: 'Export failed', details: String(error) }, { status: 500 });
    }
}
import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { generateComposedChartXML, generateMultiLineChartXML, generateDrawingXML, generateDrawingRelsXML, generateMultiBarChartXML } from '@/lib/export-sales-chart/chart-xml-generator';

export async function POST(request: NextRequest) {
    try {
        const body = await request.json();
        const { displayedData, dateRange, selectedCategories, insights, categoryColors } = body;

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Wine Analytics Dashboard';

        // Sheet 1: Price Changes by Category
        const priceSheet = workbook.addWorksheet('Price Changes');

        priceSheet.mergeCells('A1:G1');
        const priceTitleCell = priceSheet.getCell('A1');
        priceTitleCell.value = 'Price Changes by Category';
        priceTitleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        priceTitleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F1F1F' } };
        priceTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        priceSheet.getRow(1).height = 30;

        priceSheet.getCell('A2').value = 'Date Range:';
        priceSheet.getCell('B2').value = dateRange;
        priceSheet.getCell('A2').font = { bold: true };

        // Price change headers
        const priceStartRow = 4;
        const priceHeaderRow = priceSheet.getRow(priceStartRow);
        const priceHeaders = ['Month-Year', ...selectedCategories.map((cat: string) => `${cat} Price %`)];
        priceHeaderRow.values = priceHeaders;
        priceHeaderRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        priceHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F1F1F' } };
        priceHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
        priceHeaderRow.height = 25;

        // Add price data
        displayedData.forEach((item: any, index: number) => {
            const rowNum = priceStartRow + 1 + index;
            const row = priceSheet.getRow(rowNum);
            const values = [item.yearMonth];
            selectedCategories.forEach((cat: string) => {
                values.push(item[`${cat}_price`] || 0);
            });
            row.values = values;

            selectedCategories.forEach((_: any, i: number) => {
                row.getCell(i + 2).numFmt = '0.0"%"';
            });

            if (index % 2 === 0) {
                row.eachCell((cell) => {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
                });
            }
        });

        priceSheet.getColumn(1).width = 15;
        selectedCategories.forEach((_: any, i: number) => {
            priceSheet.getColumn(i + 2).width = 15;
        });

        // Sheet 2: Sales Changes by Category
        const salesSheet = workbook.addWorksheet('Sales Changes');

        salesSheet.mergeCells('A1:G1');
        const salesTitleCell = salesSheet.getCell('A1');
        salesTitleCell.value = 'Sales Changes by Category';
        salesTitleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        salesTitleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F1F1F' } };
        salesTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        salesSheet.getRow(1).height = 30;

        salesSheet.getCell('A2').value = 'Date Range:';
        salesSheet.getCell('B2').value = dateRange;
        salesSheet.getCell('A2').font = { bold: true };

        // Sales change headers
        const salesStartRow = 4;
        const salesHeaderRow = salesSheet.getRow(salesStartRow);
        const salesHeaders = ['Month-Year', ...selectedCategories.map((cat: string) => `${cat} Sales %`)];
        salesHeaderRow.values = salesHeaders;
        salesHeaderRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        salesHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F1F1F' } };
        salesHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
        salesHeaderRow.height = 25;

        // Add sales data
        displayedData.forEach((item: any, index: number) => {
            const rowNum = salesStartRow + 1 + index;
            const row = salesSheet.getRow(rowNum);
            const values = [item.yearMonth];
            selectedCategories.forEach((cat: string) => {
                values.push(item[`${cat}_sales`] || 0);
            });
            row.values = values;

            selectedCategories.forEach((_: any, i: number) => {
                row.getCell(i + 2).numFmt = '0.0"%"';
            });

            if (index % 2 === 0) {
                row.eachCell((cell) => {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
                });
            }
        });

        salesSheet.getColumn(1).width = 15;
        selectedCategories.forEach((_: any, i: number) => {
            salesSheet.getColumn(i + 2).width = 15;
        });

        // Add insights summary to sales sheet
        const insightStartRow = salesStartRow + displayedData.length + 18;
        const insightTitle = salesSheet.getCell(`A${insightStartRow}`);
        insightTitle.value = 'Category Analysis';
        insightTitle.font = { bold: true, size: 12 };
        insightTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF87CEEB' } };

        insights.forEach((insight: any, idx: number) => {
            const row = insightStartRow + idx + 1;
            salesSheet.getCell(`A${row}`).value = insight.category;
            salesSheet.getCell(`B${row}`).value = `Inverse Correlation: ${insight.inverseCorrelation}%`;
            salesSheet.getCell(`A${row}`).font = { bold: true };
        });

        // Generate base Excel
        const buffer = await workbook.xlsx.writeBuffer();
        const zip = await JSZip.loadAsync(buffer);

        // Add charts
        const dataEndRow = priceStartRow + displayedData.length;

        // Chart 1: Multi-bar chart for price changes (on price sheet)
        const priceSeries = selectedCategories.map((cat: string, idx: number) => ({
            name: cat,
            dataColumn: String.fromCharCode(66 + idx), // B, C, D, etc.
            color: categoryColors[cat]?.color?.replace('#', '') || '000000'
        }));

        const priceChartXML = generateMultiBarChartXML(
            'Price Changes',
            priceStartRow + 1,
            dataEndRow,
            'A',
            priceSeries,
            'Price Changes by Category',
            'Price Change %'
        );

        zip.file('xl/charts/chart1.xml', priceChartXML);

        // Chart 2: Multi-line chart for sales (on sales sheet)
        const salesSeries = selectedCategories.map((cat: string, idx: number) => ({
            name: cat,
            dataColumn: String.fromCharCode(66 + idx),
            color: categoryColors[cat]?.color?.replace('#', '') || '000000'
        }));

        const salesChartXML = generateMultiLineChartXML(
            'Sales Changes',
            salesStartRow + 1,
            dataEndRow,
            'A',
            salesSeries,
            'Sales Evolution by Category',
            'Sales Change %'
        );

        zip.file('xl/charts/chart2.xml', salesChartXML);

        // Add drawings for both sheets
        const drawing1XML = generateDrawingXML('rId1');
        zip.file('xl/drawings/drawing1.xml', drawing1XML);

        const drawing2XML = generateDrawingXML('rId1');
        zip.file('xl/drawings/drawing2.xml', drawing2XML);

        const drawingRels1XML = generateDrawingRelsXML();
        zip.file('xl/drawings/_rels/drawing1.xml.rels', drawingRels1XML);

        const drawingRels2XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`;
        zip.file('xl/drawings/_rels/drawing2.xml.rels', drawingRels2XML);

        // Update both worksheets
        for (let sheetNum = 1; sheetNum <= 2; sheetNum++) {
            const worksheetXML = await zip.file(`xl/worksheets/sheet${sheetNum}.xml`)?.async('string');
            if (worksheetXML) {
                const updated = worksheetXML.replace(
                    '</worksheet>',
                    `<drawing r:id="rId99"/></worksheet>`
                );
                zip.file(`xl/worksheets/sheet${sheetNum}.xml`, updated);
            }

            let worksheetRels = await zip.file(`xl/worksheets/_rels/sheet${sheetNum}.xml.rels`)?.async('string');
            if (!worksheetRels) {
                worksheetRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
            }
            const updatedRels = worksheetRels.replace(
                '</Relationships>',
                `<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${sheetNum}.xml"/></Relationships>`
            );
            zip.file(`xl/worksheets/_rels/sheet${sheetNum}.xml.rels`, updatedRels);
        }

        // Update content types
        const contentTypes = await zip.file('[Content_Types].xml')?.async('string');
        if (contentTypes) {
            let updated = contentTypes;
            if (!updated.includes('chart1.xml')) {
                updated = updated.replace(
                    '</Types>',
                    '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'
                );
            }
            if (!updated.includes('chart2.xml')) {
                updated = updated.replace(
                    '</Types>',
                    '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'
                );
            }
            if (!updated.includes('drawing1.xml')) {
                updated = updated.replace(
                    '</Types>',
                    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
                );
            }
            if (!updated.includes('drawing2.xml')) {
                updated = updated.replace(
                    '</Types>',
                    '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
                );
            }
            zip.file('[Content_Types].xml', updated);
        }

        const finalBuffer = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 9 }
        });

        return new NextResponse(Uint8Array.from(finalBuffer), {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': `attachment; filename="price-sales-by-category-${new Date().toISOString().split('T')[0]}.xlsx"`
            }
        });

    } catch (error) {
        console.error('Export error:', error);
        return NextResponse.json({ error: 'Export failed', details: String(error) }, { status: 500 });
    }
}
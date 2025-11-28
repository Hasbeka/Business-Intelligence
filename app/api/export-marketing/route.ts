import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { generateComposedChartXML, generateDrawingXML, generateDrawingRelsXML } from '@/lib/export-sales-chart/chart-xml-generator';

export async function POST(request: NextRequest) {
    try {
        const body = await request.json();
        const {
            monthlyAnalysis,
            recommendations,
            overallAvg,
            bestMonth,
            worstMonth
        } = body;

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Wine Analytics - Marketing Dashboard';

        // ===== SHEET 1: Monthly Sales Performance =====
        const monthlySheet = workbook.addWorksheet('Monthly Performance');

        // Title
        monthlySheet.mergeCells('A1:G1');
        const monthlyTitle = monthlySheet.getCell('A1');
        monthlyTitle.value = 'Monthly Sales Performance Analysis';
        monthlyTitle.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        monthlyTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
        monthlyTitle.alignment = { horizontal: 'center', vertical: 'middle' };
        monthlySheet.getRow(1).height = 35;

        // Key metrics summary
        monthlySheet.getCell('A3').value = 'Overall Average Sales:';
        monthlySheet.getCell('B3').value = overallAvg;
        monthlySheet.getCell('B3').numFmt = '$#,##0';
        monthlySheet.getCell('A3').font = { bold: true };
        monthlySheet.getCell('B3').font = { bold: true, color: { argb: 'FF2563EB' } };

        if (bestMonth) {
            monthlySheet.getCell('D3').value = 'Best Month:';
            monthlySheet.getCell('E3').value = bestMonth.monthName;
            monthlySheet.getCell('F3').value = bestMonth.avgSales;
            monthlySheet.getCell('F3').numFmt = '$#,##0';
            monthlySheet.getCell('D3').font = { bold: true };
            monthlySheet.getCell('E3').font = { color: { argb: 'FF16A34A' } };
            monthlySheet.getCell('F3').font = { bold: true, color: { argb: 'FF16A34A' } };
        }

        if (worstMonth) {
            monthlySheet.getCell('A4').value = 'Needs Attention:';
            monthlySheet.getCell('B4').value = worstMonth.monthName;
            monthlySheet.getCell('C4').value = worstMonth.avgSales;
            monthlySheet.getCell('C4').numFmt = '$#,##0';
            monthlySheet.getCell('A4').font = { bold: true };
            monthlySheet.getCell('B4').font = { color: { argb: 'FFDC2626' } };
            monthlySheet.getCell('C4').font = { bold: true, color: { argb: 'FFDC2626' } };
        }

        // Data headers
        const dataStartRow = 6;
        const headerRow = monthlySheet.getRow(dataStartRow);
        headerRow.values = [
            'Month',
            'Avg Sales ($)',
            'Avg Quantity',
            'Avg Price ($)',
            'Performance vs Avg (%)',
            'Consistency Score',
            'Data Points'
        ];
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
        headerRow.height = 25;

        // Add monthly data
        monthlyAnalysis.forEach((month: any, index: number) => {
            const rowNum = dataStartRow + 1 + index;
            const row = monthlySheet.getRow(rowNum);
            const performance = ((month.avgSales / overallAvg - 1) * 100);

            row.values = [
                month.monthName,
                month.avgSales,
                month.avgQty,
                month.avgPrice,
                performance / 100,
                month.consistency,
                month.dataPoints
            ];

            // Formatting
            row.getCell(2).numFmt = '$#,##0';
            row.getCell(3).numFmt = '#,##0';
            row.getCell(4).numFmt = '$#,##0.00';
            row.getCell(5).numFmt = '0.0%';
            row.getCell(6).numFmt = '0.00';

            // Performance color coding
            if (performance < 0) {
                row.getCell(5).font = { color: { argb: 'FFDC2626' }, bold: true };
            } else {
                row.getCell(5).font = { color: { argb: 'FF16A34A' }, bold: true };
            }

            // Alternating row colors
            if (index % 2 === 0) {
                row.eachCell((cell) => {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
                });
            }
        });

        // Column widths
        monthlySheet.getColumn(1).width = 12;
        monthlySheet.getColumn(2).width = 15;
        monthlySheet.getColumn(3).width = 15;
        monthlySheet.getColumn(4).width = 15;
        monthlySheet.getColumn(5).width = 22;
        monthlySheet.getColumn(6).width = 18;
        monthlySheet.getColumn(7).width = 13;

        // ===== SHEET 2: Campaign Recommendations =====
        const recSheet = workbook.addWorksheet('Campaign Recommendations');

        // Title
        recSheet.mergeCells('A1:F1');
        const recTitle = recSheet.getCell('A1');
        recTitle.value = 'Promotional Campaign Recommendations';
        recTitle.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        recTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEA580C' } };
        recTitle.alignment = { horizontal: 'center', vertical: 'middle' };
        recSheet.getRow(1).height = 35;

        // Subtitle
        recSheet.mergeCells('A2:F2');
        const subtitle = recSheet.getCell('A2');
        subtitle.value = 'Strategic periods identified for launching promotional campaigns';
        subtitle.font = { italic: true, size: 11 };
        subtitle.alignment = { horizontal: 'center' };

        // Headers
        const recStartRow = 4;
        const recHeaderRow = recSheet.getRow(recStartRow);
        recHeaderRow.values = [
            'Month',
            'Campaign Type',
            'Priority',
            'Performance vs Avg',
            'Strategy',
            'Reason'
        ];
        recHeaderRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        recHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        recHeaderRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        recHeaderRow.height = 30;

        // Add recommendations data
        recommendations.forEach((rec: any, index: number) => {
            const rowNum = recStartRow + 1 + index;
            const row = recSheet.getRow(rowNum);

            row.values = [
                rec.month,
                rec.type.charAt(0).toUpperCase() + rec.type.slice(1),
                rec.priority.toUpperCase(),
                ((rec.performance - 1) * 100) / 100,
                rec.strategy,
                rec.reason
            ];

            // Formatting
            row.getCell(4).numFmt = '0.0%';
            row.getCell(5).alignment = { wrapText: true };
            row.getCell(6).alignment = { wrapText: true };

            // Priority color coding
            const priorityCell = row.getCell(3);
            priorityCell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            if (rec.priority === 'high') {
                priorityCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC2626' } };
            } else if (rec.priority === 'medium') {
                priorityCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEA580C' } };
            } else {
                priorityCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } };
            }

            // Campaign type color
            const typeCell = row.getCell(2);
            if (rec.type === 'boost') {
                typeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFECACA' } };
            } else if (rec.type === 'preparation') {
                typeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            } else if (rec.type === 'clearance') {
                typeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFDBFE' } };
            } else if (rec.type === 'premium') {
                typeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBF7D0' } };
            }

            // Performance color
            if (rec.performance < 1) {
                row.getCell(4).font = { color: { argb: 'FFDC2626' }, bold: true };
            } else {
                row.getCell(4).font = { color: { argb: 'FF16A34A' }, bold: true };
            }

            row.height = 50;
        });

        // Column widths
        recSheet.getColumn(1).width = 12;
        recSheet.getColumn(2).width = 15;
        recSheet.getColumn(3).width = 12;
        recSheet.getColumn(4).width = 18;
        recSheet.getColumn(5).width = 50;
        recSheet.getColumn(6).width = 40;

        // Legend section
        const legendRow = recStartRow + recommendations.length + 3;
        recSheet.mergeCells(`A${legendRow}:F${legendRow}`);
        const legendTitle = recSheet.getCell(`A${legendRow}`);
        legendTitle.value = 'Campaign Type Guide';
        legendTitle.font = { bold: true, size: 12 };
        legendTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E7EB' } };

        const legendData = [
            ['🚀 Boost', 'Demand Stimulation - Aggressive promotions for underperforming months'],
            ['📅 Preparation', 'Peak Preparation - Build anticipation before high-sales periods'],
            ['🏷️ Clearance', 'Post-Peak Optimization - Leverage momentum after peak seasons'],
            ['⭐ Premium', 'Premium Testing - Test premium products during strong performance periods']
        ];

        legendData.forEach((item, idx) => {
            const row = recSheet.getRow(legendRow + idx + 1);
            row.getCell(1).value = item[0];
            row.getCell(2).value = item[1];
            recSheet.mergeCells(`B${legendRow + idx + 1}:F${legendRow + idx + 1}`);
            row.getCell(1).font = { bold: true };
            row.height = 25;
        });

        // Generate base Excel
        const buffer = await workbook.xlsx.writeBuffer();
        const zip = await JSZip.loadAsync(buffer);

        // ===== ADD CHART TO MONTHLY PERFORMANCE SHEET =====
        const dataEndRow = dataStartRow + monthlyAnalysis.length;

        // Create composed chart (bars + line)
        const chartXML = generateComposedChartXML(
            'Monthly Performance',
            dataStartRow + 1,
            dataEndRow,
            'A',
            'B',  // Avg Sales for bars
            'E',  // Performance % for line
            'Monthly Sales Performance',
            'Average Sales',
            'Performance vs Average',
            '3B82F6',  // Blue for bars
            'EF4444'   // Red for line
        );

        zip.file('xl/charts/chart1.xml', chartXML);

        // Add drawing
        const drawingXML = generateDrawingXML('rId1');
        zip.file('xl/drawings/drawing1.xml', drawingXML);

        const drawingRelsXML = generateDrawingRelsXML();
        zip.file('xl/drawings/_rels/drawing1.xml.rels', drawingRelsXML);

        // Update worksheet to include drawing
        const worksheetXML = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
        if (worksheetXML) {
            const updated = worksheetXML.replace(
                '</worksheet>',
                '<drawing r:id="rId99"/></worksheet>'
            );
            zip.file('xl/worksheets/sheet1.xml', updated);
        }

        // Update worksheet relationships
        let worksheetRels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('string');
        if (!worksheetRels) {
            worksheetRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        }
        const updatedRels = worksheetRels.replace(
            '</Relationships>',
            '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'
        );
        zip.file('xl/worksheets/_rels/sheet1.xml.rels', updatedRels);

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
            if (!updated.includes('drawing1.xml')) {
                updated = updated.replace(
                    '</Types>',
                    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
                );
            }
            zip.file('[Content_Types].xml', updated);
        }

        // Generate final file
        const finalBuffer = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 9 }
        });

        return new NextResponse(Uint8Array.from(finalBuffer), {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': `attachment; filename="marketing-analytics-${new Date().toISOString().split('T')[0]}.xlsx"`
            }
        });

    } catch (error) {
        console.error('Export error:', error);
        return NextResponse.json({
            error: 'Export failed',
            details: String(error)
        }, { status: 500 });
    }
}
import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export async function POST(request: NextRequest) {
    try {
        const { segments, insights } = await request.json();

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Gender Segments');

        sheet.columns = [
            { header: 'Gender', key: 'label', width: 20 },
            { header: 'Visits (%)', key: 'value', width: 15 }
        ];

        segments.forEach((seg: { label: any; value: any; }) => {
            sheet.addRow({
                label: seg.label,
                value: seg.value
            });
        });

        const insightSheet = workbook.addWorksheet('Insights');
        insightSheet.addRow(['Key Insights']);
        insights.forEach((i: string) => insightSheet.addRow([i]));

        const buffer = await workbook.xlsx.writeBuffer();

        return new NextResponse(buffer, {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': `attachment; filename="gender-segmentation.xlsx"`
            }
        });
    } catch (e) {
        return NextResponse.json({ error: 'Failed to export gender segmentation', details: String(e) }, { status: 500 });
    }
}

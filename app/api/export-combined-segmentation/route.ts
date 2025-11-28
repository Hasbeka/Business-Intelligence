import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export async function POST(request: NextRequest) {
    try {
        const { segments, insights } = await request.json();

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Combined Segments');

        sheet.columns = [
            { header: 'Age Range', key: 'ageLabel', width: 20 },
            { header: 'Gender', key: 'genderLabel', width: 15 },
            { header: 'Visit %', key: 'value', width: 12 }
        ];

        segments.forEach((seg: any) => {
            const values = Object.entries(seg.values || {});

            values.forEach(([gender, value]) => {
                sheet.addRow({
                    ageLabel: seg.label,
                    genderLabel: gender,
                    value
                });
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
                'Content-Disposition': `attachment; filename="combined-segmentation.xlsx"`
            }
        });
    } catch (e) {
        return NextResponse.json({ error: 'Failed to export combined segmentation', details: String(e) }, { status: 500 });
    }
}

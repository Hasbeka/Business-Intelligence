import { NextResponse } from 'next/server';
import { readCsv as readCSV } from '@/lib/CsvReader';

export async function GET(
    request: Request,
    { params }: { params: { filename: string } }
) {
    try {
        const data = await readCSV(`${params.filename}.csv`);

        return NextResponse.json(data);
    } catch (error) {
        return NextResponse.json(
            { error: 'CSV not found' },
            { status: 404 }
        );
    }
}
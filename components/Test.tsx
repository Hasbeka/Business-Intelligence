"use client"

import ExportButton from "@/components/ui/ExportButton"
import { exportToExcel, generateFilename } from "@/lib/ExcelExport"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"

/**
 * Simple test component to verify Excel export works
 * You can use this as a reference for implementing exports in your actual components
 */
export default function ExcelExportTest() {

    const handleExport = async () => {
        // Sample data
        const sampleData = [
            { month: 'Jan 2023', sales: 50000, quantity: 120, avgPrice: 416.67 },
            { month: 'Feb 2023', sales: 55000, quantity: 130, avgPrice: 423.08 },
            { month: 'Mar 2023', sales: 48000, quantity: 110, avgPrice: 436.36 },
            { month: 'Apr 2023', sales: 62000, quantity: 145, avgPrice: 427.59 },
            { month: 'May 2023', sales: 58000, quantity: 135, avgPrice: 429.63 },
        ];

        await exportToExcel({
            filename: generateFilename('test-export'),
            sheets: [
                {
                    name: 'Sales Data',
                    columns: [
                        { header: 'Month', key: 'month', width: 15 },
                        { header: 'Total Sales ($)', key: 'sales', width: 18 },
                        { header: 'Quantity', key: 'quantity', width: 12 },
                        { header: 'Avg Price ($)', key: 'avgPrice', width: 15 },
                    ],
                    data: sampleData,
                    summary: [
                        { label: 'Total Sales', value: '$273,000', style: 'success' },
                        { label: 'Average Sales', value: '$54,600', style: 'info' },
                        { label: 'Total Quantity', value: '640 units', style: 'info' },
                    ]
                }
            ]
        });
    };

    return (
        <Card className="w-full max-w-2xl">
            <CardHeader className="flex flex-row items-center justify-between">
                <CardTitle>Excel Export Test</CardTitle>
                <ExportButton onExport={handleExport} />
            </CardHeader>
            <CardContent>
                <p className="text-sm text-muted-foreground">
                    Click the export button to download a sample Excel file.
                    This demonstrates the Excel export functionality with formatted data and summaries.
                </p>
            </CardContent>
        </Card>
    );
}
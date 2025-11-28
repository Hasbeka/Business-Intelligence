// lib/customerSegmentationExport.ts
"use client";

import ExcelJS from 'exceljs';

type TopItem = { name: string; count: number };

type GenderSegment = {
    gender: string;
    customerCount: number;
    topCategories: TopItem[];
    topVarieties: TopItem[];
    topCountries: TopItem[];
    topPriceRanges: TopItem[];
};

type AgeSegment = {
    ageGroup: string;
    customerCount: number;
    avgAge: number;
    topCategories: TopItem[];
    topVarieties: TopItem[];
};

type CombinedSegment = {
    segment: string;
    customerCount: number;
    topCategories: TopItem[];
    topVarieties: TopItem[];
};

export async function exportCustomerSegmentationToExcel(
    genderSegments: GenderSegment[],
    ageSegments: AgeSegment[],
    combinedSegments: CombinedSegment[]
): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Wine Analytics Dashboard';
    workbook.created = new Date();

    // ===== SHEET 1: Executive Summary =====
    const summarySheet = workbook.addWorksheet('Executive Summary', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 4 }]
    });

    // Title
    summarySheet.mergeCells('A1:G1');
    const titleCell = summarySheet.getCell('A1');
    titleCell.value = '👥 Customer Segmentation Analysis';
    titleCell.font = { size: 20, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF8B5CF6' } // Purple
    };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    summarySheet.getRow(1).height = 40;

    // Subtitle
    summarySheet.mergeCells('A2:G2');
    const subtitleCell = summarySheet.getCell('A2');
    subtitleCell.value = 'Wine Preferences Across Demographics';
    subtitleCell.font = { size: 12, italic: true, color: { argb: 'FF666666' } };
    subtitleCell.alignment = { horizontal: 'center' };
    summarySheet.getRow(2).height = 20;

    // Date
    summarySheet.mergeCells('A3:G3');
    const dateCell = summarySheet.getCell('A3');
    dateCell.value = `Generated: ${new Date().toLocaleString()}`;
    dateCell.font = { size: 10, italic: true };
    dateCell.alignment = { horizontal: 'center' };

    // Key Metrics Section
    summarySheet.getRow(5).height = 30;
    summarySheet.mergeCells('A5:G5');
    const metricsTitle = summarySheet.getCell('A5');
    metricsTitle.value = '📊 KEY METRICS';
    metricsTitle.font = { size: 16, bold: true, color: { argb: 'FF8B5CF6' } };
    metricsTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    metricsTitle.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF3E8FF' } // Light purple
    };

    const totalCustomers = genderSegments.reduce((sum, s) => sum + s.customerCount, 0);
    const totalAgeGroups = ageSegments.length;
    const totalCombinedSegments = combinedSegments.length;

    const metrics = [
        { label: 'Total Customers Analyzed', value: totalCustomers.toLocaleString(), color: 'FF8B5CF6' },
        { label: 'Gender Segments', value: genderSegments.length.toString(), color: 'FFEC4899' },
        { label: 'Age Groups', value: totalAgeGroups.toString(), color: 'FFF59E0B' },
        { label: 'Combined Segments', value: totalCombinedSegments.toString(), color: 'FF10B981' },
        { label: 'Most Popular Category', value: getMostPopularCategory(genderSegments), color: 'FF3B82F6' },
        { label: 'Largest Segment', value: getLargestSegment(genderSegments), color: 'FFEF4444' },
    ];

    let currentRow = 6;
    metrics.forEach((metric, idx) => {
        const row = summarySheet.getRow(currentRow);
        row.height = 25;
        
        // Label
        summarySheet.mergeCells(`A${currentRow}:D${currentRow}`);
        const labelCell = row.getCell(1);
        labelCell.value = metric.label;
        labelCell.font = { bold: true, size: 12 };
        labelCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF8FAFC' }
        };
        labelCell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
        labelCell.alignment = { vertical: 'middle', indent: 1 };
        
        // Value
        summarySheet.mergeCells(`E${currentRow}:G${currentRow}`);
        const valueCell = row.getCell(5);
        valueCell.value = metric.value;
        valueCell.font = { bold: true, size: 14, color: { argb: metric.color } };
        valueCell.alignment = { horizontal: 'center', vertical: 'middle' };
        valueCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' }
        };
        valueCell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
        
        currentRow++;
    });

    // Set column widths
    summarySheet.getColumn(1).width = 30;
    summarySheet.getColumn(5).width = 25;

    // Key Insights Section
    currentRow += 2;
    summarySheet.getRow(currentRow).height = 30;
    summarySheet.mergeCells(`A${currentRow}:G${currentRow}`);
    const insightsTitle = summarySheet.getCell(`A${currentRow}`);
    insightsTitle.value = '💡 KEY INSIGHTS';
    insightsTitle.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    insightsTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    insightsTitle.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF10B981' } // Green
    };

    currentRow++;
    const insights = generateInsights(genderSegments, ageSegments);
    insights.forEach((insight) => {
        const row = summarySheet.getRow(currentRow);
        row.height = 30;
        summarySheet.mergeCells(`A${currentRow}:G${currentRow}`);
        const cell = row.getCell(1);
        cell.value = `• ${insight}`;
        cell.font = { size: 11 };
        cell.alignment = { vertical: 'middle', wrapText: true, indent: 1 };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF0FDF4' } // Light green
        };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
        currentRow++;
    });

    // ===== SHEET 2: Gender Segmentation =====
    const genderSheet = workbook.addWorksheet('Gender Analysis', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    // Headers
    const genderHeaders = [
        'Gender', 'Customer Count', 'Market Share %', 
        'Top Category', 'Category Purchases',
        'Top Variety', 'Variety Count',
        'Top Price Range', 'Price Range Count'
    ];

    const genderHeaderRow = genderSheet.getRow(1);
    genderHeaders.forEach((header, idx) => {
        const cell = genderHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF8B5CF6' } // Purple
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    genderHeaderRow.height = 35;

    // Data
    genderSegments.forEach((segment, idx) => {
        const row = genderSheet.getRow(idx + 2);
        row.height = 25;
        
        const marketShare = (segment.customerCount / totalCustomers) * 100;
        
        row.getCell(1).value = segment.gender;
        row.getCell(1).font = { bold: true, size: 12, color: { argb: idx === 0 ? 'FF3B82F6' : 'FFEC4899' } };
        row.getCell(1).alignment = { vertical: 'middle' };
        
        row.getCell(2).value = segment.customerCount;
        row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        row.getCell(2).font = { bold: true };
        
        row.getCell(3).value = marketShare / 100;
        row.getCell(3).numFmt = '0.0%';
        row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(4).value = segment.topCategories[0]?.name || 'N/A';
        row.getCell(4).alignment = { vertical: 'middle' };
        row.getCell(4).font = { bold: true, color: { argb: 'FF10B981' } };
        
        row.getCell(5).value = segment.topCategories[0]?.count || 0;
        row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(6).value = segment.topVarieties[0]?.name || 'N/A';
        row.getCell(6).alignment = { vertical: 'middle' };
        
        row.getCell(7).value = segment.topVarieties[0]?.count || 0;
        row.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(8).value = segment.topPriceRanges[0]?.name || 'N/A';
        row.getCell(8).alignment = { vertical: 'middle' };
        
        row.getCell(9).value = segment.topPriceRanges[0]?.count || 0;
        row.getCell(9).alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Alternating row colors
        if (idx % 2 === 0) {
            for (let col = 1; col <= 9; col++) {
                row.getCell(col).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF8FAFC' }
                };
            }
        }
    });

    // Detailed preferences section
    let detailRow = genderSegments.length + 4;
    genderSheet.mergeCells(`A${detailRow}:I${detailRow}`);
    const detailTitle = genderSheet.getCell(`A${detailRow}`);
    detailTitle.value = '📋 Detailed Preferences by Gender';
    detailTitle.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    detailTitle.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEC4899' }
    };
    detailTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    genderSheet.getRow(detailRow).height = 30;

    detailRow++;
    genderSegments.forEach((segment) => {
        // Gender header
        genderSheet.mergeCells(`A${detailRow}:I${detailRow}`);
        const genderHeader = genderSheet.getCell(`A${detailRow}`);
        genderHeader.value = `${segment.gender} - Top 5 Categories`;
        genderHeader.font = { bold: true, size: 12, color: { argb: 'FF8B5CF6' } };
        genderHeader.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF3E8FF' }
        };
        genderHeader.alignment = { vertical: 'middle', indent: 1 };
        genderSheet.getRow(detailRow).height = 25;
        detailRow++;

        // Category details
        segment.topCategories.slice(0, 5).forEach((cat, idx) => {
            const row = genderSheet.getRow(detailRow);
            row.height = 20;
            
            row.getCell(1).value = `#${idx + 1}`;
            row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
            row.getCell(1).font = { bold: true };
            
            row.getCell(2).value = cat.name;
            row.getCell(2).alignment = { vertical: 'middle' };
            
            row.getCell(3).value = cat.count;
            row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
            
            // Medal colors for top 3
            if (idx === 0) {
                row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } }; // Gold
            } else if (idx === 1) {
                row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1D5DB' } }; // Silver
            } else if (idx === 2) {
                row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFB923C' } }; // Bronze
            }
            
            detailRow++;
        });
        
        detailRow++; // Separator
    });

    // Set column widths
    genderSheet.getColumn(1).width = 15;
    genderSheet.getColumn(2).width = 15;
    genderSheet.getColumn(3).width = 15;
    genderSheet.getColumn(4).width = 20;
    genderSheet.getColumn(5).width = 18;
    genderSheet.getColumn(6).width = 25;
    genderSheet.getColumn(7).width = 15;
    genderSheet.getColumn(8).width = 18;
    genderSheet.getColumn(9).width = 18;

    // ===== SHEET 3: Age Segmentation =====
    const ageSheet = workbook.addWorksheet('Age Analysis', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    // Headers
    const ageHeaders = [
        'Age Group', 'Customer Count', 'Market Share %', 'Avg Age',
        'Top Category', 'Category Purchases',
        'Top Variety', 'Variety Count'
    ];

    const ageHeaderRow = ageSheet.getRow(1);
    ageHeaders.forEach((header, idx) => {
        const cell = ageHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF59E0B' } // Orange
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    ageHeaderRow.height = 35;

    // Data
    const totalAgeCustomers = ageSegments.reduce((sum, s) => sum + s.customerCount, 0);
    
    ageSegments.forEach((segment, idx) => {
        const row = ageSheet.getRow(idx + 2);
        row.height = 25;
        
        const marketShare = (segment.customerCount / totalAgeCustomers) * 100;
        
        row.getCell(1).value = segment.ageGroup;
        row.getCell(1).font = { bold: true, size: 12 };
        row.getCell(1).alignment = { vertical: 'middle' };
        
        row.getCell(2).value = segment.customerCount;
        row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        row.getCell(2).font = { bold: true };
        
        row.getCell(3).value = marketShare / 100;
        row.getCell(3).numFmt = '0.0%';
        row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(4).value = segment.avgAge;
        row.getCell(4).numFmt = '0.0';
        row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(5).value = segment.topCategories[0]?.name || 'N/A';
        row.getCell(5).alignment = { vertical: 'middle' };
        row.getCell(5).font = { bold: true, color: { argb: 'FF10B981' } };
        
        row.getCell(6).value = segment.topCategories[0]?.count || 0;
        row.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(7).value = segment.topVarieties[0]?.name || 'N/A';
        row.getCell(7).alignment = { vertical: 'middle' };
        
        row.getCell(8).value = segment.topVarieties[0]?.count || 0;
        row.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Alternating row colors
        if (idx % 2 === 0) {
            for (let col = 1; col <= 8; col++) {
                row.getCell(col).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFBEB' } // Light orange
                };
            }
        }
    });

    // Set column widths
    ageSheet.getColumn(1).width = 15;
    ageSheet.getColumn(2).width = 15;
    ageSheet.getColumn(3).width = 15;
    ageSheet.getColumn(4).width = 12;
    ageSheet.getColumn(5).width = 20;
    ageSheet.getColumn(6).width = 18;
    ageSheet.getColumn(7).width = 25;
    ageSheet.getColumn(8).width = 15;

    // ===== SHEET 4: Combined Segmentation =====
    const combinedSheet = workbook.addWorksheet('Combined Analysis', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    // Headers
    const combinedHeaders = [
        'Segment (Gender + Age)', 'Customer Count', 'Market Share %',
        'Rank', 'Top Category', 'Category Purchases',
        'Top Variety', 'Variety Count'
    ];

    const combinedHeaderRow = combinedSheet.getRow(1);
    combinedHeaders.forEach((header, idx) => {
        const cell = combinedHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF10B981' } // Green
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    combinedHeaderRow.height = 35;

    // Data
    const totalCombinedCustomers = combinedSegments.reduce((sum, s) => sum + s.customerCount, 0);
    
    // Sort by customer count
    const sortedCombined = [...combinedSegments].sort((a, b) => b.customerCount - a.customerCount);
    
    sortedCombined.forEach((segment, idx) => {
        const row = combinedSheet.getRow(idx + 2);
        row.height = 25;
        
        const marketShare = (segment.customerCount / totalCombinedCustomers) * 100;
        
        row.getCell(1).value = segment.segment;
        row.getCell(1).font = { bold: true, size: 11 };
        row.getCell(1).alignment = { vertical: 'middle' };
        
        row.getCell(2).value = segment.customerCount;
        row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        row.getCell(2).font = { bold: true };
        
        row.getCell(3).value = marketShare / 100;
        row.getCell(3).numFmt = '0.0%';
        row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(4).value = idx + 1;
        row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        row.getCell(4).font = { bold: true };
        
        // Top 3 ranking colors
        if (idx === 0) {
            row.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } }; // Gold
        } else if (idx === 1) {
            row.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1D5DB' } }; // Silver
        } else if (idx === 2) {
            row.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFB923C' } }; // Bronze
        }
        
        row.getCell(5).value = segment.topCategories[0]?.name || 'N/A';
        row.getCell(5).alignment = { vertical: 'middle' };
        row.getCell(5).font = { color: { argb: 'FF10B981' } };
        
        row.getCell(6).value = segment.topCategories[0]?.count || 0;
        row.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
        
        row.getCell(7).value = segment.topVarieties[0]?.name || 'N/A';
        row.getCell(7).alignment = { vertical: 'middle' };
        
        row.getCell(8).value = segment.topVarieties[0]?.count || 0;
        row.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Alternating row colors
        if (idx % 2 === 0) {
            for (let col = 1; col <= 8; col++) {
                row.getCell(col).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF0FDF4' } // Light green
                };
            }
        }
    });

    // Set column widths
    combinedSheet.getColumn(1).width = 30;
    combinedSheet.getColumn(2).width = 15;
    combinedSheet.getColumn(3).width = 15;
    combinedSheet.getColumn(4).width = 10;
    combinedSheet.getColumn(5).width = 20;
    combinedSheet.getColumn(6).width = 18;
    combinedSheet.getColumn(7).width = 25;
    combinedSheet.getColumn(8).width = 15;

    // ===== Export File =====
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Customer_Segmentation_Analysis_${new Date().toISOString().split('T')[0]}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Helper functions
function getMostPopularCategory(genderSegments: GenderSegment[]): string {
    const categoryCount = new Map<string, number>();
    
    genderSegments.forEach(segment => {
        segment.topCategories.forEach(cat => {
            categoryCount.set(cat.name, (categoryCount.get(cat.name) || 0) + cat.count);
        });
    });
    
    let maxCategory = '';
    let maxCount = 0;
    categoryCount.forEach((count, category) => {
        if (count > maxCount) {
            maxCount = count;
            maxCategory = category;
        }
    });
    
    return maxCategory || 'N/A';
}

function getLargestSegment(genderSegments: GenderSegment[]): string {
    let largest = genderSegments[0];
    genderSegments.forEach(segment => {
        if (segment.customerCount > largest.customerCount) {
            largest = segment;
        }
    });
    return largest?.gender || 'N/A';
}

function generateInsights(genderSegments: GenderSegment[], ageSegments: AgeSegment[]): string[] {
    const insights: string[] = [];

    // Gender insights
    genderSegments.forEach(seg => {
        if (seg.topCategories.length > 0) {
            const topCat = seg.topCategories[0];
            insights.push(
                `${seg.gender} customers show strong preference for ${topCat.name} wines (${topCat.count} purchases)`
            );
        }
    });

    // Age insights
    const olderSegment = ageSegments.find(s => s.ageGroup.includes('55') || s.ageGroup.includes('65'));
    if (olderSegment && olderSegment.topCategories.length > 0) {
        insights.push(
            `Older customers (${olderSegment.ageGroup}) prefer ${olderSegment.topCategories[0].name} wines`
        );
    }

    const youngerSegment = ageSegments.find(s => s.ageGroup.includes('18') || s.ageGroup.includes('25'));
    if (youngerSegment && youngerSegment.topCategories.length > 0) {
        insights.push(
            `Younger customers (${youngerSegment.ageGroup}) favor ${youngerSegment.topCategories[0].name} wines`
        );
    }

    // Market share insight
    const largestGender = genderSegments.reduce((max, seg) => 
        seg.customerCount > max.customerCount ? seg : max
    );
    const genderShare = (largestGender.customerCount / genderSegments.reduce((sum, s) => sum + s.customerCount, 0)) * 100;
    insights.push(
        `${largestGender.gender} customers represent ${genderShare.toFixed(1)}% of total customer base`
    );

    return insights;
}
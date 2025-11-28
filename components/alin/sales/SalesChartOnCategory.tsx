"use client"

import { useState, useMemo } from "react";
import { CartesianGrid, Line, LineChart, XAxis, YAxis, Tooltip, LineProps } from "recharts";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { ChartContainer } from "@/components/ui/chart";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Slider } from "@/components/ui/slider";
import { MonthlySalesF } from "@/app/types";
import { formatNumberEnglishStyle } from "@/lib/utils";
import ExportButton from "@/components/ui/ExportButton";

interface SalesByCategoryChartProps {
    salesData: MonthlySalesF[];
}

const colors = ["#3b82f6", "#ef4444", "#10b981", "#f59e0b", "#8b5cf6", "#ec4899"];
const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const SalesByCategoryChart = ({ salesData }: SalesByCategoryChartProps) => {
    const categories = useMemo(() => Array.from(new Set(salesData.map(d => d.category))), [salesData]);
    const [selectedCategory, setSelectedCategory] = useState<string | "Total">("Total");

    const sortedData = useMemo(() => salesData.sort((a, b) => a.year - b.year || a.month - b.month), [salesData]);

    const completeMonths = useMemo(() => {
        if (sortedData.length <= 2) return [];
        return sortedData.slice(1, sortedData.length - 1);
    }, [sortedData]);

    const aggregatedData = useMemo(() => {
        const map = new Map<string, { year: number; month: number; totalAmount: number; category?: string }>();

        completeMonths.forEach(d => {
            if (selectedCategory === "Total") {
                const key = `${d.year}-${d.month}`;
                const entry = map.get(key);
                if (entry) entry.totalAmount += d.totalAmount;
                else map.set(key, { year: d.year, month: d.month, totalAmount: d.totalAmount });
            } else if (d.category === selectedCategory) {
                const key = `${d.year}-${d.month}`;
                const entry = map.get(key);
                if (entry) entry.totalAmount += d.totalAmount;
                else map.set(key, { year: d.year, month: d.month, totalAmount: d.totalAmount, category: d.category });
            }
        });

        return Array.from(map.values()).sort((a, b) => a.year - b.year || a.month - b.month);
    }, [completeMonths, selectedCategory]);

    const [rangeValues, setRangeValues] = useState<number[]>([0, Math.max(0, aggregatedData.length - 1)]);
    const displayedData = useMemo(() => aggregatedData.slice(rangeValues[0], rangeValues[1] + 1), [aggregatedData, rangeValues]);

    const seasonalTotals = useMemo(() => {
        const seasons = { Winter: 0, Spring: 0, Summer: 0, Fall: 0 };
        const data = displayedData.filter(v => v.year === 2023)

        data.forEach(d => {
            if ([12, 1, 2].includes(d.month)) seasons.Winter += d.totalAmount;
            else if ([3, 4, 5].includes(d.month)) seasons.Spring += d.totalAmount;
            else if ([6, 7, 8].includes(d.month)) seasons.Summer += d.totalAmount;
            else seasons.Fall += d.totalAmount;
        });
        return seasons;
    }, [displayedData]);

    const getSeasonMessage = () => {
        const entries = Object.entries(seasonalTotals).sort((a, b) => b[1] - a[1]);
        return `Peak season: ${entries[0][0]} | Low season: ${entries[entries.length - 1][0]}`;
    };

    const lineColor = useMemo(() => {
        if (selectedCategory === "Total") return colors[0];
        const index = categories.findIndex(c => c === selectedCategory);
        return colors[index % colors.length];
    }, [selectedCategory, categories]);

    const getMaxAndMinSale = () => {
        const data = displayedData.filter(v => {
            return !(v.year === 2022 && v.month === 6 || v.year === 2024 && v.month === 6);
        });
        const maxValue = Math.max(...data.map(obj => obj.totalAmount));
        const minValue = Math.min(...data.map(obj => obj.totalAmount));
        const objOfMax = data.filter(v => v.totalAmount === maxValue);
        const objOfMin = data.filter(v => v.totalAmount === minValue);

        return `The max value was of $${formatNumberEnglishStyle(Math.round(maxValue))} in ${monthNames[objOfMax[0].month - 1]} ${objOfMax[0].year} . The min value was of $${formatNumberEnglishStyle(Math.round(minValue))} in ${monthNames[objOfMin[0].month - 1]} ${objOfMin[0].year}`;
    }

    const handleExport = async () => {
        try {
            const exportData = displayedData.map(item => ({
                monthYear: `${monthNames[item.month - 1]} ${item.year}`,
                year: item.year,
                month: item.month,
                totalAmount: item.totalAmount
            }));

            const dateRange = displayedData.length > 0
                ? `${monthNames[displayedData[0].month - 1]} ${displayedData[0].year} - ${monthNames[displayedData[displayedData.length - 1].month - 1]} ${displayedData[displayedData.length - 1].year}`
                : 'N/A';

            const response = await fetch('/api/export-sales-chart', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    displayedData: exportData,
                    selectedCategory,
                    dateRange,
                    seasonalAnalysis: getSeasonMessage(),
                    maxMinInfo: getMaxAndMinSale(),
                    lineColor
                })
            });

            if (!response.ok) {
                throw new Error('Export failed');
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `sales-trends-${selectedCategory.replace(/\s+/g, '-')}-${new Date().toISOString().split('T')[0]}.xlsx`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Export error:', error);
        }
    };

    return (
        <Card>
            <CardHeader className="flex flex-row items-center justify-between">
                <div className="flex-1">
                    <CardTitle>Sales Trends by Category</CardTitle>
                    <CardDescription>
                        <Select value={selectedCategory} onValueChange={setSelectedCategory}>
                            <SelectTrigger className="w-40">
                                <SelectValue placeholder="Select category" />
                            </SelectTrigger>
                            <SelectContent className="bg-black z-100">
                                <SelectItem value="Total">Total</SelectItem>
                                {categories.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                            </SelectContent>
                        </Select>
                    </CardDescription>
                </div>
                <ExportButton onExport={handleExport} />
            </CardHeader>

            <CardContent className="space-y-4">
                <div className="flex gap-2 mb-2 text-xs text-muted-foreground justify-between">
                    <span>{displayedData[0] ? `${monthNames[displayedData[0].month - 1]} ${displayedData[0].year}` : ""}</span>
                    <span>{displayedData[displayedData.length - 1] ? `${monthNames[displayedData[displayedData.length - 1].month - 1]} ${displayedData[displayedData.length - 1].year}` : ""}</span>
                </div>

                <Slider
                    min={0}
                    max={Math.max(0, aggregatedData.length - 1)}
                    step={1}
                    value={rangeValues}
                    onValueChange={setRangeValues}
                    className="w-full"
                />

                <ChartContainer config={{ desktop: { label: "Sales by Category", color: "from-green-600 to-blue-600" } }}>
                    <LineChart
                        data={displayedData}
                        margin={{ left: 40, right: 24 }}
                    >
                        <CartesianGrid vertical={false} />
                        <XAxis
                            dataKey={d => `${monthNames[d.month - 1]} ${d.year}`}
                            tickLine={false}
                            axisLine={false}
                            tickMargin={8}
                        />
                        <YAxis
                            tickFormatter={(value) => {
                                if (value >= 1_000_000) return `$${(value / 1_000_000).toFixed(1)}M`;
                                if (value >= 1_000) return `$${(value / 1_000).toFixed(1)}K`;
                                return `$${value}`;
                            }}
                        />
                        <Tooltip
                            formatter={(value: any) => `$${Number(value).toFixed(2)}`}
                        />
                        <Line
                            name="Total Amount"
                            dataKey="totalAmount"
                            type="monotone"
                            stroke={lineColor}
                            strokeWidth={2}
                            dot={false}
                        />
                    </LineChart>
                </ChartContainer>
            </CardContent>

            <CardFooter className="flex-col items-start gap-2 text-sm">
                <div>{getSeasonMessage()}</div>
                <div>{getMaxAndMinSale()}</div>
            </CardFooter>
        </Card>
    );
};

export default SalesByCategoryChart;
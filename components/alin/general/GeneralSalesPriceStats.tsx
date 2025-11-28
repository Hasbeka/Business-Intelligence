"use client"

import { CartesianGrid, Line, Tooltip, XAxis, YAxis, Legend, ReferenceLine, Bar, ComposedChart } from "recharts"
import { useState, useMemo } from "react"

import {
    Card,
    CardContent,
    CardDescription,
    CardFooter,
    CardHeader,
    CardTitle,
} from "@/components/ui/card"
import {
    ChartConfig,
    ChartContainer,
} from "@/components/ui/chart"
import { Slider } from "@/components/ui/slider"
import { MonthlyStats } from "@/app/types"
import ExportButton from "@/components/ui/ExportButton"

export const description = "Highlighting the link between price changes and sales evolution";

const chartConfig = {
    priceChange: {
        label: "Price Change %",
        color: "hsl(var(--chart-1))",
    },
    salesChange: {
        label: "Sales Change %",
        color: "hsl(var(--chart-2))",
    },
} satisfies ChartConfig;

interface SalesChartProps {
    chartData: MonthlyStats[],
}

export const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const GeneralSalePriceStats = ({ chartData }: SalesChartProps) => {
    const allFormattedData = useMemo(() => {
        const sorted = chartData
            .map(item => ({
                ...item,
                yearMonth: `${monthNames[item.month - 1]} ${item.year}`
            }))
            .sort((a, b) => (a.year - b.year) || (a.month - b.month));

        if (sorted.length > 2) {
            sorted.shift();
            sorted.pop();
        }

        const withChanges = sorted.map((item, index) => {
            if (index === 0) {
                return {
                    ...item,
                    priceChange: 0,
                    salesChange: 0,
                    priceChangePercent: 0,
                    salesChangePercent: 0
                };
            }

            const prev = sorted[index - 1];
            const priceChange = ((item.avgPrice - prev.avgPrice) / prev.avgPrice) * 100;
            const salesChange = ((item.totalAmount - prev.totalAmount) / prev.totalAmount) * 100;

            return {
                ...item,
                priceChange: item.avgPrice - prev.avgPrice,
                salesChange: item.totalAmount - prev.totalAmount,
                priceChangePercent: priceChange,
                salesChangePercent: salesChange
            };
        });

        return withChanges;
    }, [chartData]);

    const [rangeValues, setRangeValues] = useState<number[]>([0, Math.max(0, allFormattedData.length - 1)]);

    const displayedData = useMemo(() => {
        return allFormattedData.slice(rangeValues[0], rangeValues[1] + 1);
    }, [allFormattedData, rangeValues]);

    const getDateRange = () => {
        if (displayedData.length === 0) return "";
        const first = displayedData[0];
        const last = displayedData[displayedData.length - 1];
        return `${monthNames[first.month - 1]} ${first.year} - ${monthNames[last.month - 1]} ${last.year}`;
    };

    const handleSliderChange = (values: number[]) => {
        setRangeValues(values);
    };

    const correlationInsight = useMemo(() => {
        if (displayedData.length < 2) return null;

        let inverseCases = 0;
        let directCases = 0;
        let totalCases = 0;

        displayedData.slice(1).forEach(item => {
            if (Math.abs(item.priceChangePercent) > 0.5) {
                totalCases++;
                if ((item.priceChangePercent < 0 && item.salesChangePercent > 0) ||
                    (item.priceChangePercent > 0 && item.salesChangePercent < 0)) {
                    inverseCases++;
                } else if ((item.priceChangePercent < 0 && item.salesChangePercent < 0) ||
                    (item.priceChangePercent > 0 && item.salesChangePercent > 0)) {
                    directCases++;
                }
            }
        });

        if (totalCases === 0) return null;

        const inversePercent = (inverseCases / totalCases * 100).toFixed(0);
        return `${inversePercent}% inverse correlation (price changes opposite to sales)`;
    }, [displayedData]);

    const handleExport = async () => {
        try {
            const response = await fetch('/api/export-price-stats', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    displayedData,
                    dateRange: getDateRange(),
                    correlationInsight: correlationInsight || 'No significant correlation detected'
                })
            });

            if (!response.ok) {
                throw new Error('Export failed');
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `price-vs-sales-${new Date().toISOString().split('T')[0]}.xlsx`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Export error:', error);
        }
    };

    return (
        <div>
            <Card>
                <CardHeader className="flex flex-row items-center justify-between">
                    <div className="flex-1">
                        <CardTitle>Price Changes vs Sales Evolution</CardTitle>
                        <CardDescription>{getDateRange()}</CardDescription>
                    </div>
                    <ExportButton onExport={handleExport} />
                </CardHeader>
                <CardContent className="space-y-4">
                    <div className="space-y-2">
                        <div className="flex justify-between text-xs text-muted-foreground">
                            <span>{allFormattedData[rangeValues[0]]?.yearMonth || ''}</span>
                            <span>{allFormattedData[rangeValues[1]]?.yearMonth || ''}</span>
                        </div>
                        <Slider
                            min={0}
                            max={Math.max(0, allFormattedData.length - 1)}
                            step={1}
                            value={rangeValues}
                            onValueChange={handleSliderChange}
                            className="w-full"
                        />
                    </div>

                    <ChartContainer config={chartConfig}>
                        <ComposedChart data={displayedData} margin={{ left: 12, right: 12, top: 20 }}>
                            <CartesianGrid vertical={false} strokeDasharray="3 3" />
                            <XAxis
                                dataKey="yearMonth"
                                tickLine={false}
                                axisLine={false}
                                tickMargin={8}
                                angle={-45}
                                textAnchor="end"
                                height={80}
                            />
                            <YAxis
                                tickFormatter={(value) => `${value > 0 ? '+' : ''}${value.toFixed(0)}%`}
                                label={{ value: 'Change %', angle: -90, position: 'insideLeft' }}
                            />
                            <ReferenceLine y={0} stroke="#666" strokeDasharray="3 3" />

                            <Tooltip
                                content={({ active, payload }) => {
                                    if (!active || !payload?.length) return null;
                                    const data = payload[0].payload;
                                    return (
                                        <div className="bg-neutral-200 z-10 text-black  border rounded-lg shadow-lg p-3">
                                            <p className="font-medium mb-2">{data.yearMonth}</p>
                                            <p className="text-sm" style={{ color: '#ef4444' }}>
                                                Price: {data.priceChangePercent > 0 ? '+' : ''}{data.priceChangePercent.toFixed(1)}%
                                                <span className="text-muted-foreground ml-2">(${data.avgPrice.toFixed(2)})</span>
                                            </p>
                                            <p className="text-sm" style={{ color: '#3b82f6' }}>
                                                Sales: {data.salesChangePercent > 0 ? '+' : ''}{data.salesChangePercent.toFixed(1)}%
                                                <span className="text-muted-foreground ml-2">(${data.totalAmount.toLocaleString()})</span>
                                            </p>
                                        </div>
                                    );
                                }}
                            />
                            <Legend />

                            <Bar
                                name="Price Change %"
                                dataKey="priceChangePercent"
                                fill="#ef4444"
                                opacity={0.7}
                                radius={[4, 4, 0, 0]}
                            />
                            <Line
                                name="Sales Change %"
                                dataKey="salesChangePercent"
                                type="monotone"
                                stroke="#3b82f6"
                                strokeWidth={3}
                                dot={{ fill: '#3b82f6', r: 4 }}
                            />
                        </ComposedChart>
                    </ChartContainer>
                </CardContent>
                <CardFooter className="flex-col items-start gap-2 text-sm">

                    {correlationInsight && (
                        <div className="text-muted-foreground">
                            {correlationInsight}
                        </div>
                    )}
                </CardFooter>
            </Card>
        </div>
    )
}

export default GeneralSalePriceStats
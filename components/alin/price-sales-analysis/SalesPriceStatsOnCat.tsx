"use client"

import { CartesianGrid, Line, LineChart, Tooltip, XAxis, YAxis, Legend, ReferenceLine, Bar, ComposedChart } from "recharts"
import { useState, useMemo } from "react"
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card"
import { Checkbox } from "@/components/ui/checkbox"
import { Label } from "@/components/ui/label"
import { ChartConfig, ChartContainer } from "@/components/ui/chart"
import { Slider } from "@/components/ui/slider"
import { MonthlySalesF } from "@/app/types"
import ExportButton from "@/components/ui/ExportButton"

const chartConfig = {
    Dessert: { label: "Dessert", color: "#8b5cf6" },
    Red: { label: "Red", color: "#dc2626" },
    Rosé: { label: "Rosé", color: "#ec4899" },
    Sparkling: { label: "Sparkling", color: "#fbbf24" },
    White: { label: "White", color: "#22c55e" },
} satisfies ChartConfig;

interface SalesPriceStatsCatProps {
    chartData: MonthlySalesF[]
}

export const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const SalesPriceStatsOnCat = ({ chartData }: SalesPriceStatsCatProps) => {
    const [rangeValues, setRangeValues] = useState<number[]>([0, 0]);
    const [selectedCategories, setSelectedCategories] = useState<string[]>(Object.keys(chartConfig));

    const toggleCategory = (category: string) => {
        setSelectedCategories(prev =>
            prev.includes(category) ? prev.filter(c => c !== category) : [...prev, category]
        );
    };

    const allFormattedData = useMemo(() => {
        const grouped = chartData.reduce((acc, item) => {
            const key = `${item.year}-${item.month}`;
            if (!acc[key]) {
                acc[key] = {
                    year: item.year,
                    month: item.month,
                    yearMonth: `${monthNames[item.month - 1]} ${item.year}`,
                    categories: {}
                };
            }
            acc[key].categories[item.category] = {
                priceChangePercent: 0,
                salesChangePercent: 0,
                avgPrice: item.avgPrice,
                totalAmount: item.totalAmount
            };
            return acc;
        }, {} as any);

        const sorted = Object.values(grouped).sort((a: any, b: any) => (a.year - b.year) || (a.month - b.month));

        if (sorted.length > 2) {
            sorted.shift();
            sorted.pop();
        }

        const withChanges = sorted.map((item: any, index: number) => {
            if (index === 0) return item;
            const prev: any = sorted[index - 1];
            const categories = { ...item.categories };

            Object.keys(categories).forEach(cat => {
                if (prev.categories[cat]) {
                    const priceChange = ((categories[cat].avgPrice - prev.categories[cat].avgPrice) / prev.categories[cat].avgPrice) * 100;
                    const salesChange = ((categories[cat].totalAmount - prev.categories[cat].totalAmount) / prev.categories[cat].totalAmount) * 100;
                    categories[cat] = {
                        ...categories[cat],
                        priceChangePercent: priceChange,
                        salesChangePercent: salesChange
                    };
                }
            });

            return { ...item, categories };
        });

        return withChanges.map((item: any) => {
            const flattened: any = { yearMonth: item.yearMonth, year: item.year, month: item.month };
            Object.keys(item.categories).forEach(cat => {
                flattened[`${cat}_price`] = item.categories[cat].priceChangePercent;
                flattened[`${cat}_sales`] = item.categories[cat].salesChangePercent;
                flattened[`${cat}_avgPrice`] = item.categories[cat].avgPrice;
                flattened[`${cat}_totalAmount`] = item.categories[cat].totalAmount;
            });
            return flattened;
        });
    }, [chartData]);

    useMemo(() => {
        if (allFormattedData.length > 0 && rangeValues[1] === 0) {
            setRangeValues([0, Math.max(0, allFormattedData.length - 1)]);
        }
    }, [allFormattedData]);

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

    const insights = useMemo(() => {
        if (displayedData.length < 2) return [];
        return selectedCategories.map(cat => {
            let inverseCases = 0, totalCases = 0, avgPriceChange = 0, avgSalesChange = 0;
            displayedData.slice(1).forEach(item => {
                const priceChange = item[`${cat}_price`];
                const salesChange = item[`${cat}_sales`];
                if (Math.abs(priceChange) > 0.5) {
                    totalCases++;
                    if ((priceChange < 0 && salesChange > 0) || (priceChange > 0 && salesChange < 0)) {
                        inverseCases++;
                    }
                    avgPriceChange += priceChange;
                    avgSalesChange += salesChange;
                }
            });
            return {
                category: cat,
                color: chartConfig[cat as keyof typeof chartConfig].color,
                inverseCorrelation: totalCases > 0 ? Math.round((inverseCases / totalCases) * 100) : 0,
                avgPriceChange: totalCases > 0 ? avgPriceChange / totalCases : 0,
                avgSalesChange: totalCases > 0 ? avgSalesChange / totalCases : 0,
                priceElasticity: totalCases > 0 && avgPriceChange !== 0 ? (avgSalesChange / avgPriceChange).toFixed(2) : 'N/A'
            };
        });
    }, [displayedData, selectedCategories]);

    const handleExport = async () => {
        try {
            const response = await fetch('/api/export-price-cat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    displayedData,
                    dateRange: getDateRange(),
                    selectedCategories,
                    insights,
                    categoryColors: chartConfig
                })
            });

            if (!response.ok) throw new Error('Export failed');

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `price-sales-by-category-${new Date().toISOString().split('T')[0]}.xlsx`;
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
                        <CardTitle>Price Changes by Category</CardTitle>
                        <CardDescription>{getDateRange()}</CardDescription>
                    </div>
                    <ExportButton onExport={handleExport} />
                </CardHeader>
                <CardContent className="space-y-4">
                    <div className="flex gap-4 flex-wrap">
                        {Object.keys(chartConfig).map(cat => (
                            <div key={cat} className="flex items-center gap-2">
                                <Checkbox
                                    id={cat}
                                    checked={selectedCategories.includes(cat)}
                                    onCheckedChange={() => toggleCategory(cat)}
                                />
                                <Label htmlFor={cat}>{cat}</Label>
                            </div>
                        ))}
                    </div>

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
                            <XAxis dataKey="yearMonth" tickLine={false} axisLine={false} tickMargin={8} angle={-45} textAnchor="end" height={80} />
                            <YAxis tickFormatter={(value) => `${value > 0 ? '+' : ''}${value.toFixed(0)}%`} label={{ value: 'Price Change %', angle: -90, position: 'insideLeft' }} />
                            <ReferenceLine y={0} stroke="#666" strokeDasharray="3 3" />
                            <Tooltip content={({ active, payload }) => {
                                if (!active || !payload?.length) return null;
                                const data = payload[0].payload;
                                return (
                                    <div className="bg-neutral-200 z-10 text-black border rounded-lg shadow-lg p-3 max-w-xs">
                                        <p className="font-medium mb-2">{data.yearMonth}</p>
                                        {selectedCategories.map(cat => {
                                            const priceKey = `${cat}_price`;
                                            if (data[priceKey] !== undefined) {
                                                return <p key={cat} className="text-xs" style={{ color: chartConfig[cat as keyof typeof chartConfig].color }}>
                                                    {cat}: {data[priceKey] > 0 ? '+' : ''}{data[priceKey].toFixed(1)}%
                                                </p>;
                                            }
                                            return null;
                                        })}
                                    </div>
                                );
                            }} />
                            <Legend />
                            {selectedCategories.map(cat => (
                                <Bar key={cat} name={cat} dataKey={`${cat}_price`} fill={chartConfig[cat as keyof typeof chartConfig].color} opacity={0.7} radius={[4, 4, 0, 0]} />
                            ))}
                        </ComposedChart>
                    </ChartContainer>
                </CardContent>
            </Card>

            <Card>
                <CardHeader>
                    <CardTitle>Sales Evolution by Category</CardTitle>
                    <CardDescription>{getDateRange()}</CardDescription>
                </CardHeader>
                <CardContent className="space-y-4">
                    <div className="space-y-2">
                        <div className="flex justify-between text-xs text-muted-foreground">
                            <span>{allFormattedData[rangeValues[0]]?.yearMonth || ''}</span>
                            <span>{allFormattedData[rangeValues[1]]?.yearMonth || ''}</span>
                        </div>
                        <Slider min={0} max={Math.max(0, allFormattedData.length - 1)} step={1} value={rangeValues} onValueChange={handleSliderChange} className="w-full" />
                    </div>

                    <ChartContainer config={chartConfig}>
                        <LineChart data={displayedData} margin={{ left: 12, right: 12, top: 20 }}>
                            <CartesianGrid vertical={false} strokeDasharray="3 3" />
                            <XAxis dataKey="yearMonth" tickLine={false} axisLine={false} tickMargin={8} angle={-45} textAnchor="end" height={80} />
                            <YAxis tickFormatter={(value) => `${value > 0 ? '+' : ''}${value.toFixed(0)}%`} label={{ value: 'Sales Change %', angle: -90, position: 'insideLeft' }} />
                            <ReferenceLine y={0} stroke="#666" strokeDasharray="3 3" />
                            <Tooltip content={({ active, payload }) => {
                                if (!active || !payload?.length) return null;
                                const data = payload[0].payload;
                                return (
                                    <div className="bg-neutral-200 text-black z-10 border rounded-lg shadow-lg p-3 max-w-xs">
                                        <p className="font-medium mb-2">{data.yearMonth}</p>
                                        {selectedCategories.map(cat => {
                                            const salesKey = `${cat}_sales`;
                                            if (data[salesKey] !== undefined) {
                                                return <p key={cat} className="text-xs" style={{ color: chartConfig[cat as keyof typeof chartConfig].color }}>
                                                    {cat}: {data[salesKey] > 0 ? '+' : ''}{data[salesKey].toFixed(1)}%
                                                </p>;
                                            }
                                            return null;
                                        })}
                                    </div>
                                );
                            }} />
                            <Legend />
                            {selectedCategories.map(cat => (
                                <Line key={cat} name={cat} dataKey={`${cat}_sales`} type="monotone" stroke={chartConfig[cat as keyof typeof chartConfig].color} strokeWidth={2.5} dot={{ r: 3 }} />
                            ))}
                        </LineChart>
                    </ChartContainer>
                </CardContent>
                <CardFooter className="flex-col items-start gap-4 text-sm border-t pt-4">
                    <div className="font-semibold text-base">Analysis Results</div>
                    {insights.length === 0 ? (
                        <p className="text-muted-foreground">Select at least one category</p>
                    ) : (
                        <div className="w-full space-y-3">
                            {insights.map((insight: any) => (
                                <div key={insight.category} className="border rounded-lg p-3 space-y-2">
                                    <div className="flex items-center gap-2 font-medium">
                                        <div className="w-3 h-3 rounded" style={{ backgroundColor: insight.color }} />
                                        {insight.category}
                                    </div>
                                    <div className="grid grid-cols-2 gap-2 text-xs text-muted-foreground">
                                        <div><span className="font-medium">Inverse Correlation:</span> {insight.inverseCorrelation}%</div>
                                        <div><span className="font-medium">Avg Price Change:</span> {insight.avgPriceChange > 0 ? '+' : ''}{insight.avgPriceChange.toFixed(1)}%</div>
                                        <div><span className="font-medium">Avg Sales Change:</span> {insight.avgSalesChange > 0 ? '+' : ''}{insight.avgSalesChange.toFixed(1)}%</div>
                                        <div><span className="font-medium">Price Elasticity:</span> {insight.priceElasticity}</div>
                                    </div>
                                </div>
                            ))}
                        </div>
                    )}
                </CardFooter>
            </Card>
        </div>
    )
}

export default SalesPriceStatsOnCat
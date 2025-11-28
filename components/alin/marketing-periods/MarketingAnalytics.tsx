"use client"

import { TrendingUp, TrendingDown, Calendar, AlertCircle, ChartBar, Download } from "lucide-react"
import { Bar, CartesianGrid, XAxis, YAxis, Tooltip, Legend, Line, ComposedChart } from "recharts"
import { useMemo, useState } from "react"

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
import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { MonthlyStats } from "@/app/types"
import { Separator } from "../../ui/separator"
import ExportButton from "@/components/ui/ExportButton"

const chartConfig = {
    avgSales: {
        label: "Average Sales",
        color: "#3b82f6",
    },
    average: {
        label: "Overall Average",
        color: "#ef4444",
    },
} satisfies ChartConfig;

interface PromotionalPeriodsProps {
    chartData: MonthlyStats[]
}

export const exportMarketingAnalytics = async (
    monthlyAnalysis: any[],
    recommendations: any[],
    overallAvg: number,
    bestMonth: any,
    worstMonth: any
) => {
    try {
        const response = await fetch('/api/export-marketing', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                monthlyAnalysis,
                recommendations,
                overallAvg,
                bestMonth,
                worstMonth
            })
        });

        if (!response.ok) {
            throw new Error('Export failed');
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `marketing-analytics-${new Date().toISOString().split('T')[0]}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        return { success: true };
    } catch (error) {
        console.error('Export error:', error);
        return { success: false, error };
    }
};

export const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const MarketingAnalytics = ({ chartData }: PromotionalPeriodsProps) => {
    const [isExporting, setIsExporting] = useState(false);

    const monthlyAnalysis = useMemo(() => {
        const sorted = [...chartData].sort((a, b) =>
            (a.year - b.year) || (a.month - b.month)
        );

        if (sorted.length > 2) {
            sorted.shift();
            sorted.pop();
        }

        const byMonth: { [key: number]: { sales: number[], quantities: number[], prices: number[] } } = {};

        sorted.forEach(item => {
            if (!byMonth[item.month]) {
                byMonth[item.month] = { sales: [], quantities: [], prices: [] };
            }
            byMonth[item.month].sales.push(item.totalAmount);
            byMonth[item.month].quantities.push(item.totalQty);
            byMonth[item.month].prices.push(item.avgPrice);
        });

        const analysis = Object.keys(byMonth)
            .map(month => {
                const monthNum = parseInt(month);
                const data = byMonth[monthNum];

                const avgSales = data.sales.reduce((a, b) => a + b, 0) / data.sales.length;
                const avgQty = data.quantities.reduce((a, b) => a + b, 0) / data.quantities.length;
                const avgPrice = data.prices.reduce((a, b) => a + b, 0) / data.prices.length;

                const salesVariance = data.sales.reduce((sum, val) => sum + Math.pow(val - avgSales, 2), 0) / data.sales.length;
                const consistency = 1 / (1 + Math.sqrt(salesVariance) / avgSales);

                return {
                    month: monthNum,
                    monthName: monthNames[monthNum - 1],
                    avgSales,
                    avgQty,
                    avgPrice,
                    consistency,
                    dataPoints: data.sales.length
                };
            })
            .sort((a, b) => a.month - b.month);

        return analysis;
    }, [chartData]);

    const overallAvg = useMemo(() => {
        if (monthlyAnalysis.length === 0) return 0;
        const total = monthlyAnalysis.reduce((sum, m) => sum + m.avgSales, 0);
        return total / monthlyAnalysis.length;
    }, [monthlyAnalysis]);

    const chartDataFormatted = useMemo(() => {
        return monthlyAnalysis.map(m => ({
            monthName: m.monthName,
            avgSales: m.avgSales,
            average: overallAvg,
            performance: ((m.avgSales / overallAvg - 1) * 100).toFixed(1),
            dataPoints: m.dataPoints
        }));
    }, [monthlyAnalysis, overallAvg]);

    const recommendations = useMemo(() => {
        if (monthlyAnalysis.length === 0) return [];

        const recs: any[] = [];

        monthlyAnalysis.forEach((month, idx) => {
            const performanceRatio = month.avgSales / overallAvg;
            const prevMonth = idx > 0 ? monthlyAnalysis[idx - 1] : null;
            const nextMonth = idx < monthlyAnalysis.length - 1 ? monthlyAnalysis[idx + 1] : null;

            if (performanceRatio < 0.85) {
                recs.push({
                    month: month.monthName,
                    type: 'boost',
                    priority: performanceRatio < 0.7 ? 'high' : 'medium',
                    reason: `Sales ${((1 - performanceRatio) * 100).toFixed(0)}% below average`,
                    strategy: 'Launch aggressive discounts and promotions to stimulate demand',
                    performance: performanceRatio,
                    avgSales: month.avgSales
                });
            }

            if (nextMonth && (nextMonth.avgSales / overallAvg) > 1.15) {
                recs.push({
                    month: month.monthName,
                    type: 'preparation',
                    priority: 'medium',
                    reason: `Prepare for peak season in ${nextMonth.monthName}`,
                    strategy: 'Build anticipation with early-bird offers and inventory optimization',
                    performance: performanceRatio,
                    avgSales: month.avgSales
                });
            }

            if (prevMonth && (prevMonth.avgSales / overallAvg) > 1.15 && performanceRatio < 1.0) {
                recs.push({
                    month: month.monthName,
                    type: 'clearance',
                    priority: 'low',
                    reason: `Leverage post-peak momentum from ${prevMonth.monthName}`,
                    strategy: 'Clearance sales and loyalty rewards to maintain engagement',
                    performance: performanceRatio,
                    avgSales: month.avgSales
                });
            }

            if (performanceRatio >= 1.0 && performanceRatio < 1.2 && month.consistency > 0.7) {
                recs.push({
                    month: month.monthName,
                    type: 'premium',
                    priority: 'low',
                    reason: 'Stable, consistent sales period',
                    strategy: 'Test premium products and upselling strategies',
                    performance: performanceRatio,
                    avgSales: month.avgSales
                });
            }
        });

        const priorityOrder = { high: 0, medium: 1, low: 2 };
        return recs.sort((a, b) => priorityOrder[a.priority as keyof typeof priorityOrder] - priorityOrder[b.priority as keyof typeof priorityOrder]);
    }, [monthlyAnalysis, overallAvg]);

    const getPriorityColor = (priority: string) => {
        switch (priority) {
            case 'high': return 'bg-red-500 text-white';
            case 'medium': return 'bg-yellow-500 text-black';
            case 'low': return 'bg-blue-500 text-white';
            default: return 'bg-gray-500 text-white';
        }
    };

    const getTypeIcon = (type: string) => {
        switch (type) {
            case 'boost': return TrendingUp;
            case 'preparation': return Calendar;
            case 'clearance': return TrendingDown;
            case 'premium': return AlertCircle;
            default: return Calendar;
        }
    };

    const bestMonth = monthlyAnalysis.length > 0
        ? monthlyAnalysis.reduce((max, m) => m.avgSales > max.avgSales ? m : max)
        : null;

    const worstMonth = monthlyAnalysis.length > 0
        ? monthlyAnalysis.reduce((min, m) => m.avgSales < min.avgSales ? m : min)
        : null;

    const handleExport = async () => {
        setIsExporting(true);
        try {
            await exportMarketingAnalytics(
                monthlyAnalysis,
                recommendations,
                overallAvg,
                bestMonth,
                worstMonth
            );
        } finally {
            setIsExporting(false);
        }
    };

    return (
        <div className="space-y-6">
            <Card>
                <CardHeader>
                    <div className="flex items-center justify-between">
                        <div>
                            <CardTitle className="text-2xl font-bold">Monthly Sales Patterns</CardTitle>
                            <Separator className="bg-neutral-300" />
                            <CardDescription>Average sales performance by month across all years</CardDescription>
                        </div>
                        <ExportButton onExport={handleExport} />
                    </div>
                </CardHeader>
                <CardContent>
                    <ChartContainer config={chartConfig}>
                        <ComposedChart data={chartDataFormatted} margin={{ left: 12, right: 12, top: 20 }}>
                            <CartesianGrid vertical={false} strokeDasharray="3 3" />
                            <XAxis
                                dataKey="monthName"
                                tickLine={false}
                                axisLine={false}
                                tickMargin={8}
                            />
                            <YAxis
                                tickFormatter={(value) => {
                                    if (value >= 1_000_000) return `$${(value / 1_000_000).toFixed(1)}M`;
                                    if (value >= 1_000) return `$${(value / 1_000).toFixed(0)}K`;
                                    return `$${value}`;
                                }}
                            />
                            <Tooltip
                                content={({ active, payload }) => {
                                    if (!active || !payload?.length) return null;
                                    const data = payload[0].payload;
                                    return (
                                        <div className="bg-neutral-200 text-black z-10 border rounded-lg shadow-lg p-3">
                                            <p className="font-medium mb-2">{data.monthName}</p>
                                            <p className="text-sm">
                                                Avg Sales: <span className="font-medium">${Number(data.avgSales).toLocaleString(undefined, { maximumFractionDigits: 0 })}</span>
                                            </p>
                                            <p className="text-sm">
                                                Performance: <span className={Number(data.performance) > 0 ? "text-green-600 font-medium" : "text-red-600 font-medium"}>
                                                    {Number(data.performance) > 0 ? '+' : ''}{data.performance}%
                                                </span>
                                            </p>
                                            <p className="text-xs text-muted-foreground mt-1">
                                                Based on {data.dataPoints} data points
                                            </p>
                                        </div>
                                    );
                                }}
                            />
                            <Legend />
                            <Bar
                                name="Average Sales"
                                dataKey="avgSales"
                                fill="#3b82f6"
                                radius={[4, 4, 0, 0]}
                            />
                            <Line
                                name="Overall Average"
                                dataKey="average"
                                stroke="#ef4444"
                                strokeWidth={2}
                                strokeDasharray="5 5"
                                dot={false}
                            />
                        </ComposedChart>
                    </ChartContainer>
                </CardContent>
            </Card>

            <Card>
                <CardHeader>
                    <CardTitle>Promotional Campaign Recommendations</CardTitle>
                    <CardDescription>
                        Strategic periods identified for launching promotional campaigns
                    </CardDescription>
                </CardHeader>
                <CardContent>
                    {recommendations.length === 0 ? (
                        <p className="text-muted-foreground text-center py-8">
                            No specific recommendations available. Need more data points.
                        </p>
                    ) : (
                        <div className="space-y-4">
                            {recommendations.map((rec, idx) => {
                                const Icon = getTypeIcon(rec.type);
                                return (
                                    <div key={idx} className="border rounded-lg p-4 space-y-3 hover:shadow-md transition-shadow">
                                        <div className="flex items-start justify-between">
                                            <div className="flex items-center gap-3">
                                                <Icon className="h-5 w-5 text-muted-foreground shrink-0" />
                                                <div>
                                                    <div className="flex items-center gap-2 flex-wrap">
                                                        <h4 className="font-semibold text-lg">{rec.month}</h4>
                                                        <Badge className={getPriorityColor(rec.priority)}>
                                                            {rec.priority.toUpperCase()}
                                                        </Badge>
                                                    </div>
                                                    <p className="text-sm text-muted-foreground mt-1">
                                                        {rec.reason}
                                                    </p>
                                                </div>
                                            </div>
                                            <div className="text-right shrink-0">
                                                <div className={`text-sm font-medium ${rec.performance < 1 ? 'text-red-600' : 'text-green-600'}`}>
                                                    {((rec.performance - 1) * 100).toFixed(0)}%
                                                </div>
                                                <div className="text-xs text-muted-foreground">vs avg</div>
                                            </div>
                                        </div>

                                        <div className="bg-muted/50 rounded p-3">
                                            <p className="text-sm font-medium mb-1">💡 Recommended Strategy:</p>
                                            <p className="text-sm text-muted-foreground">{rec.strategy}</p>
                                        </div>

                                        <div className="flex gap-2 text-xs flex-wrap">
                                            {rec.type === 'boost' && (
                                                <span className="bg-red-100 text-red-800 px-2 py-1 rounded">🚀 Demand Stimulation</span>
                                            )}
                                            {rec.type === 'preparation' && (
                                                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded">📅 Peak Preparation</span>
                                            )}
                                            {rec.type === 'clearance' && (
                                                <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded">🏷️ Post-Peak Optimization</span>
                                            )}
                                            {rec.type === 'premium' && (
                                                <span className="bg-green-100 text-green-800 px-2 py-1 rounded">⭐ Premium Testing</span>
                                            )}
                                        </div>
                                    </div>
                                );
                            })}
                        </div>
                    )}
                </CardContent>
                <CardFooter className="flex-col items-start gap-3 text-sm border-t pt-4">
                    <div className="font-semibold text-base flex gap-3"><ChartBar /> Key Insights</div>
                    {bestMonth && worstMonth && (
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 w-full">
                            <div className="space-y-2 p-3 bg-green-50 rounded-lg border border-green-200">
                                <div className="flex items-center gap-2">
                                    <TrendingUp className="h-4 w-4 text-green-600" />
                                    <span className="font-medium text-green-900">Best Performing Month</span>
                                </div>
                                <p className="text-sm text-green-700">
                                    <span className="font-semibold">{bestMonth.monthName}</span> -
                                    ${bestMonth.avgSales.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                                </p>
                            </div>
                            <div className="space-y-2 p-3 bg-red-50 rounded-lg border border-red-200">
                                <div className="flex items-center gap-2">
                                    <TrendingDown className="h-4 w-4 text-red-600" />
                                    <span className="font-medium text-red-900">Needs Most Attention</span>
                                </div>
                                <p className="text-sm text-red-700">
                                    <span className="font-semibold">{worstMonth.monthName}</span> -
                                    ${worstMonth.avgSales.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                                </p>
                            </div>
                        </div>
                    )}
                </CardFooter>
            </Card>
        </div>
    )
}

export default MarketingAnalytics
"use client"

import { TrendingUp, TrendingDown, Minus } from "lucide-react"
import { CartesianGrid, Line, LineChart, Tooltip, XAxis, YAxis } from "recharts"
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
import { MonthlySales, SalesComparison } from "@/app/types"


export const description = "The sales trends over time for the wine products.";

const chartConfig = {
    desktop: {
        label: "Total Sales Over Time",
        color: "from-red-800 to-purple-800",
    },
} satisfies ChartConfig;

interface SalesChartProps {
    chartData: MonthlySales[],
    salesComparison: SalesComparison
}

export const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const SalesChart = ({ chartData, salesComparison }: SalesChartProps) => {
    // Sort and format data
    const allFormattedData = useMemo(() => {
        const formatted = chartData.map(item => ({
            ...item,
            yearMonth: `${monthNames[item.month - 1]} ${item.year}`
        })).sort((a, b) => (a.year - b.year) || (a.month - b.month));

        // Remove last incomplete month if needed
        if (formatted.length > 0) {
            formatted.shift();
            formatted.pop();
        }

        return formatted;
    }, [chartData]);

    // State for slider range
    const [rangeValues, setRangeValues] = useState<number[]>([0, Math.max(0, allFormattedData.length - 1)]);

    // Filter data based on slider range
    const displayedData = useMemo(() => {
        return allFormattedData.slice(rangeValues[0], rangeValues[1] + 1);
    }, [allFormattedData, rangeValues]);

    const getIcon = () => {
        const message = salesComparison.message.toLowerCase();

        if (message.includes('increased')) {
            return TrendingUp;
        } else if (message.includes('decreased')) {
            return TrendingDown;
        } else if (message.includes('unchanged')) {
            return Minus;
        }

        return null;
    };

    const Icon = getIcon();

    // Get date range for description
    const getDateRange = () => {
        if (displayedData.length === 0) return "";
        const first = displayedData[0];
        const last = displayedData[displayedData.length - 1];
        return `${monthNames[first.month - 1]} ${first.year} - ${monthNames[last.month - 1]} ${last.year}`;
    };

    const handleSliderChange = (values: number[]) => {
        setRangeValues(values);
    };

    return (
        <div>
            <Card>
                <CardHeader>
                    <CardTitle>Total Sales over time</CardTitle>
                    <CardDescription>{getDateRange()}</CardDescription>
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
                        <LineChart data={displayedData} margin={{ left: 24, right: 24 }}>
                            <CartesianGrid vertical={false} />
                            <XAxis
                                dataKey="yearMonth"
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
                                type="natural" // smooth line
                                stroke="red"
                                strokeWidth={2}
                                dot={false}
                            />
                        </LineChart>
                    </ChartContainer>
                </CardContent>
                <CardFooter className="flex-col items-start gap-2 text-sm">
                    <div className="flex gap-2 leading-none font-medium">
                        {salesComparison.message}
                        {Icon && <Icon className="h-4 w-4" />}
                    </div>
                </CardFooter>
            </Card>
        </div>
    )
}

export default SalesChart
// components/CustomerSegmentAnalysis.tsx
"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Users, Wine, TrendingUp, Download } from "lucide-react";
import { Bar, BarChart, CartesianGrid, Legend, ResponsiveContainer, Tooltip, XAxis, YAxis, PieChart, Pie, Cell } from "recharts";
import { exportCustomerSegmentationToExcel } from "@/lib/customerSegmenationExport";
import { useState } from "react";

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

type Props = {
    genderSegments: GenderSegment[];
    ageSegments: AgeSegment[];
    combinedSegments: CombinedSegment[];
};

const COLORS = ['#8b5cf6', '#ec4899', '#f59e0b', '#10b981', '#3b82f6', '#ef4444'];

export default function CustomerSegmentAnalysis({ genderSegments, ageSegments, combinedSegments }: Props) {
    const [isExporting, setIsExporting] = useState(false);
    
    const handleExportToExcel = async () => {
        setIsExporting(true);
        try {
            await exportCustomerSegmentationToExcel(
                genderSegments,
                ageSegments,
                combinedSegments
            );
        } catch (error) {
            console.error('Export failed:', error);
            alert('Failed to export data. Please try again.');
        } finally {
            setIsExporting(false);
        }
    };

    return (
        <div className="p-6 space-y-6 bg-gradient-to-br from-zinc-50 to-zinc-100 dark:from-black dark:to-zinc-900 min-h-screen">
            {/* Header */}
            <div className="space-y-2">
                <div className="flex items-center justify-between">
                    <div>
                        <h1 className="text-4xl font-bold tracking-tight bg-gradient-to-r from-purple-600 to-pink-600 bg-clip-text text-transparent">
                            Customer Segmentation Analysis
                        </h1>
                        <p className="text-lg text-muted-foreground mt-2">
                            Discover wine preferences across different customer demographics
                        </p>
                    </div>
                    <Button 
                        onClick={handleExportToExcel}
                        disabled={isExporting}
                        className="bg-gradient-to-r from-green-600 to-emerald-600 hover:from-green-700 hover:to-emerald-700 text-white shadow-lg hover:shadow-xl transition-all"
                        size="lg"
                    >
                        <Download className="mr-2 h-5 w-5" />
                        {isExporting ? 'Exporting...' : 'Export to Excel'}
                    </Button>
                </div>
            </div>

            <Tabs defaultValue="gender" className="space-y-6">
                <TabsList className="grid w-full max-w-md grid-cols-3 bg-white dark:bg-zinc-800 p-1 shadow-sm">
                    <TabsTrigger 
                        value="gender"
                        className="data-[state=active]:bg-gradient-to-r data-[state=active]:from-purple-600 data-[state=active]:to-pink-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        By Gender
                    </TabsTrigger>
                    <TabsTrigger 
                        value="age"
                        className="data-[state=active]:bg-gradient-to-r data-[state=active]:from-purple-600 data-[state=active]:to-pink-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        By Age
                    </TabsTrigger>
                    <TabsTrigger 
                        value="combined"
                        className="data-[state=active]:bg-gradient-to-r data-[state=active]:from-purple-600 data-[state=active]:to-pink-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        Combined
                    </TabsTrigger>
                </TabsList>

                {/* Gender Segmentation */}
                <TabsContent value="gender" className="space-y-6">
                    <div className="grid gap-6 md:grid-cols-2">
                        {genderSegments.map((segment, idx) => (
                            <Card key={segment.gender} className="border-2 hover:border-purple-400 transition-all hover:shadow-xl bg-white dark:bg-zinc-900">
                                <CardHeader className="bg-gradient-to-r from-purple-50 to-pink-50 dark:from-purple-950/20 dark:to-pink-950/20 border-b">
                                    <div className="flex items-center justify-between">
                                        <CardTitle className="flex items-center gap-2 text-2xl">
                                            <div className={`p-2 rounded-lg ${idx === 0 ? 'bg-blue-100 dark:bg-blue-900' : 'bg-pink-100 dark:bg-pink-900'}`}>
                                                <Users className={`h-6 w-6 ${idx === 0 ? 'text-blue-600' : 'text-pink-600'}`} />
                                            </div>
                                            {segment.gender}
                                        </CardTitle>
                                        <Badge variant="secondary" className="text-lg px-4 py-1">
                                            {segment.customerCount} customers
                                        </Badge>
                                    </div>
                                </CardHeader>
                                <CardContent className="space-y-6 pt-6">
                                    {/* Top Categories */}
                                    <div className="bg-gradient-to-br from-purple-50 to-transparent dark:from-purple-950/10 p-4 rounded-lg">
                                        <h4 className="text-sm font-bold mb-3 flex items-center gap-2 text-purple-700 dark:text-purple-400">
                                            <Wine className="h-4 w-4" />
                                            Preferred Wine Categories
                                        </h4>
                                        <div className="space-y-3">
                                            {segment.topCategories.map((cat, idx) => (
                                                <div key={cat.name} className="flex items-center justify-between bg-white dark:bg-zinc-800 p-2 rounded-lg shadow-sm">
                                                    <div className="flex items-center gap-2">
                                                        <span className="flex items-center justify-center w-6 h-6 rounded-full bg-purple-600 text-white text-xs font-bold">
                                                            {idx + 1}
                                                        </span>
                                                        <span className="font-medium">{cat.name}</span>
                                                    </div>
                                                    <Badge variant="outline" className="bg-purple-50 dark:bg-purple-950/20">
                                                        {cat.count} purchases
                                                    </Badge>
                                                </div>
                                            ))}
                                        </div>
                                    </div>

                                    {/* Top Varieties */}
                                    <div>
                                        <h4 className="text-sm font-bold mb-3 text-pink-700 dark:text-pink-400">Top Grape Varieties</h4>
                                        <div className="flex flex-wrap gap-2">
                                            {segment.topVarieties.slice(0, 5).map((variety, idx) => (
                                                <Badge 
                                                    key={variety.name} 
                                                    className="bg-gradient-to-r from-pink-500 to-purple-500 text-white hover:shadow-lg transition-all"
                                                >
                                                    {variety.name} ({variety.count})
                                                </Badge>
                                            ))}
                                        </div>
                                    </div>

                                    {/* Price Preferences */}
                                    <div className="bg-gradient-to-br from-pink-50 to-transparent dark:from-pink-950/10 p-4 rounded-lg">
                                        <h4 className="text-sm font-bold mb-3 text-pink-700 dark:text-pink-400">Price Range Preferences</h4>
                                        <div className="space-y-2">
                                            {segment.topPriceRanges.map((price, idx) => (
                                                <div key={price.name} className="flex items-center justify-between text-sm bg-white dark:bg-zinc-800 p-2 rounded-lg">
                                                    <span className="font-medium">{price.name}</span>
                                                    <div className="flex items-center gap-2">
                                                        <div className="w-24 h-2 bg-gray-200 dark:bg-gray-700 rounded-full overflow-hidden">
                                                            <div 
                                                                className="h-full bg-gradient-to-r from-pink-500 to-purple-500"
                                                                style={{ width: `${(price.count / segment.topPriceRanges[0].count) * 100}%` }}
                                                            />
                                                        </div>
                                                        <span className="text-muted-foreground w-8 text-right">{price.count}</span>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                </CardContent>
                            </Card>
                        ))}
                    </div>

                    {/* Gender Comparison Chart */}
                    <Card className="border-2 border-purple-200 dark:border-purple-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-purple-50 to-pink-50 dark:from-purple-950/20 dark:to-pink-950/20 border-b">
                            <CardTitle className="text-2xl">Wine Category Preferences by Gender</CardTitle>
                            <CardDescription className="text-base">Compare wine category preferences across genders</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <GenderComparisonChart genderSegments={genderSegments} />
                        </CardContent>
                    </Card>
                </TabsContent>

                {/* Age Segmentation */}
                <TabsContent value="age" className="space-y-6">
                    <div className="grid gap-6 md:grid-cols-2 lg:grid-cols-3">
                        {ageSegments.map((segment, idx) => (
                            <Card key={segment.ageGroup} className="border-2 hover:border-purple-400 transition-all hover:shadow-xl bg-white dark:bg-zinc-900">
                                <CardHeader className="bg-gradient-to-br from-purple-100 to-pink-100 dark:from-purple-950/30 dark:to-pink-950/30 border-b">
                                    <div className="flex items-center justify-between">
                                        <CardTitle className="text-xl">{segment.ageGroup} years</CardTitle>
                                        <Badge className="bg-purple-600 text-white">
                                            {segment.customerCount}
                                        </Badge>
                                    </div>
                                    <CardDescription className="text-base font-semibold text-purple-700 dark:text-purple-400">
                                        Avg age: {segment.avgAge} years
                                    </CardDescription>
                                </CardHeader>
                                <CardContent className="space-y-4 pt-6">
                                    <div className="bg-purple-50 dark:bg-purple-950/10 p-3 rounded-lg">
                                        <h4 className="text-sm font-bold mb-2 text-purple-700 dark:text-purple-400">Top Categories</h4>
                                        <div className="space-y-2">
                                            {segment.topCategories.slice(0, 3).map((cat, idx) => (
                                                <div key={cat.name} className="flex items-center justify-between text-sm bg-white dark:bg-zinc-800 p-2 rounded">
                                                    <span className="font-medium">{idx + 1}. {cat.name}</span>
                                                    <Badge variant="outline" className="text-xs">{cat.count}</Badge>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                    <div>
                                        <h4 className="text-sm font-bold mb-2 text-pink-700 dark:text-pink-400">Top Varieties</h4>
                                        <div className="flex flex-wrap gap-1">
                                            {segment.topVarieties.slice(0, 3).map((variety) => (
                                                <Badge key={variety.name} className="bg-gradient-to-r from-pink-500 to-purple-500 text-white text-xs">
                                                    {variety.name}
                                                </Badge>
                                            ))}
                                        </div>
                                    </div>
                                </CardContent>
                            </Card>
                        ))}
                    </div>

                    <Card className="border-2 border-purple-200 dark:border-purple-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-purple-50 to-pink-50 dark:from-purple-950/20 dark:to-pink-950/20 border-b">
                            <CardTitle className="text-2xl">Age Group Distribution</CardTitle>
                            <CardDescription className="text-base">Customer distribution across age groups</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <AgeDistributionChart ageSegments={ageSegments} />
                        </CardContent>
                    </Card>
                </TabsContent>

                {/* Combined Segmentation */}
                <TabsContent value="combined" className="space-y-6">
                    <Card className="border-2 border-purple-200 dark:border-purple-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-purple-50 to-pink-50 dark:from-purple-950/20 dark:to-pink-950/20 border-b">
                            <CardTitle className="text-2xl">Combined Demographics Analysis</CardTitle>
                            <CardDescription className="text-base">Wine preferences by gender and age group</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <div className="space-y-4">
                                {combinedSegments.map((segment, idx) => (
                                    <Card key={segment.segment} className="border-2 hover:border-purple-300 transition-all bg-gradient-to-br from-white to-purple-50 dark:from-zinc-900 dark:to-purple-950/10">
                                        <CardHeader className="pb-3 bg-white/50 dark:bg-zinc-800/50">
                                            <div className="flex items-center justify-between">
                                                <CardTitle className="text-lg font-bold">{segment.segment}</CardTitle>
                                                <Badge className="bg-gradient-to-r from-purple-600 to-pink-600 text-white">
                                                    {segment.customerCount} customers
                                                </Badge>
                                            </div>
                                        </CardHeader>
                                        <CardContent className="pt-4">
                                            <div className="grid gap-4 md:grid-cols-2">
                                                <div className="bg-purple-50 dark:bg-purple-950/10 p-3 rounded-lg">
                                                    <h5 className="text-sm font-bold mb-2 text-purple-700 dark:text-purple-400">Wine Categories</h5>
                                                    {segment.topCategories.map((cat, idx) => (
                                                        <div key={cat.name} className="flex justify-between text-sm py-1">
                                                            <span className="font-medium">{idx + 1}. {cat.name}</span>
                                                            <span className="text-muted-foreground">{cat.count}</span>
                                                        </div>
                                                    ))}
                                                </div>
                                                <div className="bg-pink-50 dark:bg-pink-950/10 p-3 rounded-lg">
                                                    <h5 className="text-sm font-bold mb-2 text-pink-700 dark:text-pink-400">Grape Varieties</h5>
                                                    <div className="flex flex-wrap gap-1">
                                                        {segment.topVarieties.map((variety) => (
                                                            <Badge key={variety.name} variant="outline" className="text-xs bg-white dark:bg-zinc-800">
                                                                {variety.name}
                                                            </Badge>
                                                        ))}
                                                    </div>
                                                </div>
                                            </div>
                                        </CardContent>
                                    </Card>
                                ))}
                            </div>
                        </CardContent>
                    </Card>
                </TabsContent>
            </Tabs>

            {/* Key Insights */}
            <Card className="border-2 border-green-200 dark:border-green-900 shadow-lg">
                <CardHeader className="bg-gradient-to-r from-green-50 to-emerald-50 dark:from-green-950/20 dark:to-emerald-950/20 border-b">
                    <CardTitle className="flex items-center gap-2 text-2xl">
                        <TrendingUp className="h-6 w-6 text-green-600" />
                        Key Insights
                    </CardTitle>
                </CardHeader>
                <CardContent className="pt-6">
                    <KeyInsights 
                        genderSegments={genderSegments} 
                        ageSegments={ageSegments} 
                    />
                </CardContent>
            </Card>
        </div>
    );
}

function GenderComparisonChart({ genderSegments }: { genderSegments: GenderSegment[] }) {
    const categories = new Set<string>();
    genderSegments.forEach(seg => {
        seg.topCategories.forEach(cat => categories.add(cat.name));
    });

    const chartData = Array.from(categories).map(category => {
        const dataPoint: any = { category };
        genderSegments.forEach(seg => {
            const cat = seg.topCategories.find(c => c.name === category);
            dataPoint[seg.gender] = cat?.count || 0;
        });
        return dataPoint;
    });

    return (
        <ResponsiveContainer width="100%" height={350}>
            <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                <XAxis dataKey="category" />
                <YAxis />
                <Tooltip 
                    contentStyle={{ 
                        backgroundColor: 'white', 
                        border: '2px solid #8b5cf6',
                        borderRadius: '8px',
                        boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
                    }}
                />
                <Legend />
                {genderSegments.map((seg, idx) => (
                    <Bar key={seg.gender} dataKey={seg.gender} fill={COLORS[idx]} radius={[8, 8, 0, 0]} />
                ))}
            </BarChart>
        </ResponsiveContainer>
    );
}

function AgeDistributionChart({ ageSegments }: { ageSegments: AgeSegment[] }) {
    const data = ageSegments.map(seg => ({
        name: seg.ageGroup,
        value: seg.customerCount
    }));

    return (
        <ResponsiveContainer width="100%" height={350}>
            <PieChart>
                <Pie
                    data={data}
                    cx="50%"
                    cy="50%"
                    labelLine={false}
                    label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                    outerRadius={100}
                    fill="#8884d8"
                    dataKey="value"
                >
                    {data.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                    ))}
                </Pie>
                <Tooltip 
                    contentStyle={{ 
                        backgroundColor: 'white', 
                        border: '2px solid #8b5cf6',
                        borderRadius: '8px',
                        boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
                    }}
                />
            </PieChart>
        </ResponsiveContainer>
    );
}

function KeyInsights({ genderSegments, ageSegments }: { genderSegments: GenderSegment[], ageSegments: AgeSegment[] }) {
    const insights: string[] = [];

    // Gender insights
    genderSegments.forEach(seg => {
        if (seg.topCategories.length > 0) {
            const topCat = seg.topCategories[0];
            insights.push(`${seg.gender} customers show strong preference for ${topCat.name} wines`);
        }
        if (seg.topVarieties.length > 0) {
            const topVar = seg.topVarieties[0];
            insights.push(`${topVar.name} is the most popular variety among ${seg.gender} customers`);
        }
    });

    // Age insights
    const olderSegment = ageSegments.find(s => s.ageGroup === "55+");
    if (olderSegment && olderSegment.topCategories.length > 0) {
        insights.push(`Customers over 55 prefer ${olderSegment.topCategories[0].name} wines`);
    }

    const youngerSegment = ageSegments.find(s => s.ageGroup === "18-25");
    if (youngerSegment && youngerSegment.topCategories.length > 0) {
        insights.push(`Younger customers (18-25) favor ${youngerSegment.topCategories[0].name} wines`);
    }

    return (
        <div className="space-y-3">
            {insights.map((insight, idx) => (
                <div key={idx} className="flex items-start gap-3 p-3 bg-green-50 dark:bg-green-950/10 rounded-lg hover:bg-green-100 dark:hover:bg-green-950/20 transition-colors">
                    <div className="h-2 w-2 rounded-full bg-green-500 mt-2 flex-shrink-0" />
                    <p className="text-sm font-medium">{insight}</p>
                </div>
            ))}
        </div>
    );
}
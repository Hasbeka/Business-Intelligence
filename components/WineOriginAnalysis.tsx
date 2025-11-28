// components/WineOriginAnalysis.tsx
"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Globe, MapPin, TrendingUp, DollarSign, Package, Award, Download, Map as MapIcon } from "lucide-react";
import { Bar, BarChart, CartesianGrid, Legend, ResponsiveContainer, Tooltip, XAxis, YAxis, PieChart, Pie, Cell } from "recharts";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { exportToExcel, generateFilename, formatCurrency, type ExcelSheet } from "@/lib/ExcelExport";
import { exportWineOriginWithMap } from "@/lib/excelMapExport";

type CountrySummary = {
    country: string;
    totalSales: number;
    totalRevenue: number;
    totalQuantity: number;
    uniqueWines: number;
    avgRevenuePerSale: number;
    topCategories: { name: string; count: number }[];
    topRegions: { name: string; count: number }[];
};

type RegionSummary = {
    region: string;
    country: string;
    totalSales: number;
    totalRevenue: number;
    totalQuantity: number;
    avgRevenuePerSale: number;
};

type Props = {
    countrySummary: CountrySummary[];
    regionSummary: RegionSummary[];
};

const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#84cc16'];

export default function WineOriginAnalysis({ countrySummary, regionSummary }: Props) {
    const topCountries = countrySummary.slice(0, 10);
    const totalRevenue = countrySummary.reduce((sum, c) => sum + c.totalRevenue, 0);
    const totalSales = countrySummary.reduce((sum, c) => sum + c.totalSales, 0);

    const handleExportData = () => {
        const countryData = countrySummary.map((country, idx) => ({
            'Rank': idx + 1,
            'Country': country.country,
            'Total Sales': country.totalSales,
            'Total Revenue': country.totalRevenue.toFixed(2),
            'Total Quantity': country.totalQuantity,
            'Unique Wines': country.uniqueWines,
            'Avg Revenue per Sale': country.avgRevenuePerSale.toFixed(2),
            'Top Category': country.topCategories[0]?.name || 'N/A',
            'Top Region': country.topRegions[0]?.name || 'N/A'
        }));

        const regionData = regionSummary.map((region, idx) => ({
            'Rank': idx + 1,
            'Region': region.region,
            'Country': region.country,
            'Total Sales': region.totalSales,
            'Total Revenue': region.totalRevenue.toFixed(2),
            'Total Quantity': region.totalQuantity,
            'Avg Revenue per Sale': region.avgRevenuePerSale.toFixed(2)
        }));

        exportToExcel({
            filename: generateFilename("Wine_Origin_Data"),
            sheets: [
                {
                    name: "Countries",
                    columns: Object.keys(countryData[0]).map(k => ({ header: k, key: k, width: 18 })),
                    data: countryData
                },
                {
                    name: "Regions",
                    columns: Object.keys(regionData[0]).map(k => ({ header: k, key: k, width: 18 })),
                    data: regionData
                }
            ]
        });
    };

    const handleExportWithMap = async () => {
        await exportWineOriginWithMap(countrySummary, regionSummary);
    };

    return (
        <div className="p-6 space-y-6 bg-gradient-to-br from-zinc-50 to-zinc-100 dark:from-black dark:to-zinc-900 min-h-screen">
            {/* Header */}
            <div className="space-y-2">
                <div className="flex items-center justify-between">
                    <div>
                        <h1 className="text-4xl font-bold tracking-tight bg-gradient-to-r from-blue-600 to-cyan-600 bg-clip-text text-transparent flex items-center gap-3">
                            <Globe className="h-10 w-10 text-blue-600" />
                            Wine Origin Analysis
                        </h1>
                        <p className="text-lg text-muted-foreground mt-2">
                            Explore the relationship between wine country of origin and sales performance
                        </p>
                    </div>
                    <div className="flex gap-3">
                        <Button 
                            onClick={handleExportData}
                            className="bg-gradient-to-r from-green-600 to-emerald-600 hover:from-green-700 hover:to-emerald-700 text-white shadow-lg"
                            size="lg"
                        >       
                            <Download className="mr-2 h-5 w-5" />
                            Export Data
                        </Button>
                        <Button 
                            onClick={handleExportWithMap}
                            className="bg-gradient-to-r from-blue-600 to-cyan-600 hover:from-blue-700 hover:to-cyan-700 text-white shadow-lg"
                            size="lg"
                        >
                            <MapIcon className="mr-2 h-5 w-5" />
                            Export with Map
                        </Button>
                    </div>
                </div>
            </div>

            {/* Key Metrics */}
            <div className="grid gap-4 md:grid-cols-4">
                <Card className="border-2 border-blue-200 dark:border-blue-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-gradient-to-br from-blue-50 to-transparent dark:from-blue-950/20">
                        <CardTitle className="text-sm font-medium">Total Countries</CardTitle>
                        <div className="p-2 bg-blue-100 dark:bg-blue-900 rounded-lg">
                            <Globe className="h-4 w-4 text-blue-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-3xl font-bold text-blue-600">{countrySummary.length}</div>
                        <p className="text-xs text-muted-foreground mt-1">Wine producing countries</p>
                    </CardContent>
                </Card>

                <Card className="border-2 border-green-200 dark:border-green-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-gradient-to-br from-green-50 to-transparent dark:from-green-950/20">
                        <CardTitle className="text-sm font-medium">Total Revenue</CardTitle>
                        <div className="p-2 bg-green-100 dark:bg-green-900 rounded-lg">
                            <DollarSign className="h-4 w-4 text-green-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-3xl font-bold text-green-600">${totalRevenue.toLocaleString()}</div>
                        <p className="text-xs text-muted-foreground mt-1">Across all origins</p>
                    </CardContent>
                </Card>

                <Card className="border-2 border-orange-200 dark:border-orange-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-gradient-to-br from-orange-50 to-transparent dark:from-orange-950/20">
                        <CardTitle className="text-sm font-medium">Total Sales</CardTitle>
                        <div className="p-2 bg-orange-100 dark:bg-orange-900 rounded-lg">
                            <Package className="h-4 w-4 text-orange-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-3xl font-bold text-orange-600">{totalSales.toLocaleString()}</div>
                        <p className="text-xs text-muted-foreground mt-1">Wine transactions</p>
                    </CardContent>
                </Card>

                <Card className="border-2 border-yellow-200 dark:border-yellow-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-gradient-to-br from-yellow-50 to-transparent dark:from-yellow-950/20">
                        <CardTitle className="text-sm font-medium">Top Country</CardTitle>
                        <div className="p-2 bg-yellow-100 dark:bg-yellow-900 rounded-lg">
                            <Award className="h-4 w-4 text-yellow-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-2xl font-bold text-yellow-600">{countrySummary[0]?.country || 'N/A'}</div>
                        <p className="text-xs text-muted-foreground mt-1">
                            ${countrySummary[0]?.totalRevenue.toLocaleString() || 0} revenue
                        </p>
                    </CardContent>
                </Card>
            </div>

            <Tabs defaultValue="countries" className="space-y-6">
                <TabsList className="grid w-full max-w-md grid-cols-3 bg-white dark:bg-zinc-800 p-1 shadow-sm">
                    <TabsTrigger 
                        value="countries"
                        className="data-[state=active]:bg-gradient-to-r data-[state=active]:from-blue-600 data-[state=active]:to-cyan-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        By Country
                    </TabsTrigger>
                    <TabsTrigger 
                        value="regions"
                        className="data-[state=active]:bg-gradient-to-r data-[state=active]:from-blue-600 data-[state=active]:to-cyan-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        By Region
                    </TabsTrigger>
                    <TabsTrigger 
                        value="comparison"
                        className="data-[state=active]:bg-gradient-to-r data-[state=active]:from-blue-600 data-[state=active]:to-cyan-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        Comparison
                    </TabsTrigger>
                </TabsList>

                {/* Countries Tab */}
                <TabsContent value="countries" className="space-y-6">
                    {/* Top Countries Chart */}
                    <Card className="border-2 border-blue-200 dark:border-blue-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-blue-50 to-cyan-50 dark:from-blue-950/20 dark:to-cyan-950/20 border-b">
                            <CardTitle className="text-2xl">Top 10 Countries by Revenue</CardTitle>
                            <CardDescription className="text-base">Total revenue generated from each wine-producing country</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <ResponsiveContainer width="100%" height={400}>
                                <BarChart data={topCountries}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                                    <XAxis dataKey="country" angle={-45} textAnchor="end" height={100} />
                                    <YAxis />
                                    <Tooltip 
                                        formatter={(value) => `$${Number(value).toLocaleString()}`}
                                        contentStyle={{ 
                                            backgroundColor: 'white', 
                                            border: '2px solid #3b82f6',
                                            borderRadius: '8px',
                                            boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
                                        }}
                                    />
                                    <Bar dataKey="totalRevenue" fill="#3b82f6" name="Revenue" radius={[8, 8, 0, 0]} />
                                </BarChart>
                            </ResponsiveContainer>
                        </CardContent>
                    </Card>

                    {/* Country Details Table */}
                    <Card className="border-2 border-blue-200 dark:border-blue-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-blue-50 to-cyan-50 dark:from-blue-950/20 dark:to-cyan-950/20 border-b">
                            <CardTitle className="text-2xl">Detailed Country Performance</CardTitle>
                            <CardDescription className="text-base">Comprehensive breakdown of sales metrics by country</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <Table>
                                <TableHeader>
                                    <TableRow className="bg-blue-50 dark:bg-blue-950/10 hover:bg-blue-100 dark:hover:bg-blue-950/20">
                                        <TableHead className="font-bold">Rank</TableHead>
                                        <TableHead className="font-bold">Country</TableHead>
                                        <TableHead className="text-right font-bold">Total Sales</TableHead>
                                        <TableHead className="text-right font-bold">Revenue</TableHead>
                                        <TableHead className="text-right font-bold">Avg Sale</TableHead>
                                        <TableHead className="text-right font-bold">Unique Wines</TableHead>
                                        <TableHead className="font-bold">Top Categories</TableHead>
                                    </TableRow>
                                </TableHeader>
                                <TableBody>
                                    {countrySummary.map((country, idx) => (
                                        <TableRow key={country.country} className="hover:bg-blue-50 dark:hover:bg-blue-950/10 transition-colors">
                                            <TableCell>
                                                {idx < 3 ? (
                                                    <Badge className={
                                                        idx === 0 ? "bg-yellow-500" :
                                                        idx === 1 ? "bg-gray-400" :
                                                        "bg-orange-600"
                                                    }>
                                                        {idx + 1}
                                                    </Badge>
                                                ) : (
                                                    <span className="font-medium text-muted-foreground">{idx + 1}</span>
                                                )}
                                            </TableCell>
                                            <TableCell className="font-bold">{country.country}</TableCell>
                                            <TableCell className="text-right">{country.totalSales}</TableCell>
                                            <TableCell className="text-right font-semibold text-green-600">
                                                ${country.totalRevenue.toLocaleString()}
                                            </TableCell>
                                            <TableCell className="text-right">${country.avgRevenuePerSale.toFixed(2)}</TableCell>
                                            <TableCell className="text-right">{country.uniqueWines}</TableCell>
                                            <TableCell>
                                                <div className="flex flex-wrap gap-1">
                                                    {country.topCategories.slice(0, 2).map(cat => (
                                                        <Badge key={cat.name} variant="secondary" className="text-xs">
                                                            {cat.name}
                                                        </Badge>
                                                    ))}
                                                </div>
                                            </TableCell>
                                        </TableRow>
                                    ))}
                                </TableBody>
                            </Table>
                        </CardContent>
                    </Card>

                    {/* Market Share Pie Chart */}
                    <Card className="border-2 border-blue-200 dark:border-blue-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-blue-50 to-cyan-50 dark:from-blue-950/20 dark:to-cyan-950/20 border-b">
                            <CardTitle className="text-2xl">Market Share by Country</CardTitle>
                            <CardDescription className="text-base">Revenue distribution across top wine-producing countries</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <ResponsiveContainer width="100%" height={400}>
                                <PieChart>
                                    <Pie
                                        data={topCountries.slice(0, 8)}
                                        cx="50%"
                                        cy="50%"
                                        labelLine={false}
                                        label={({ country, percent }) => `${country}: ${(percent * 100).toFixed(1)}%`}
                                        outerRadius={120}
                                        fill="#8884d8"
                                        dataKey="totalRevenue"
                                    >
                                        {topCountries.slice(0, 8).map((entry, index) => (
                                            <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                                        ))}
                                    </Pie>
                                    <Tooltip 
                                        formatter={(value) => `$${Number(value).toLocaleString()}`}
                                        contentStyle={{ 
                                            backgroundColor: 'white', 
                                            border: '2px solid #3b82f6',
                                            borderRadius: '8px',
                                            boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
                                        }}
                                    />
                                </PieChart>
                            </ResponsiveContainer>
                        </CardContent>
                    </Card>
                </TabsContent>

                {/* Regions Tab */}
                <TabsContent value="regions" className="space-y-6">
                    <Card className="border-2 border-green-200 dark:border-green-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-green-50 to-emerald-50 dark:from-green-950/20 dark:to-emerald-950/20 border-b">
                            <CardTitle className="text-2xl">Top Wine Regions by Revenue</CardTitle>
                            <CardDescription className="text-base">Discover the highest-performing wine regions globally</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <ResponsiveContainer width="100%" height={500}>
                                <BarChart data={regionSummary} layout="horizontal">
                                    <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                                    <XAxis type="number" />
                                    <YAxis dataKey="region" type="category" width={150} />
                                    <Tooltip 
                                        formatter={(value) => `$${Number(value).toLocaleString()}`}
                                        contentStyle={{ 
                                            backgroundColor: 'white', 
                                            border: '2px solid #10b981',
                                            borderRadius: '8px',
                                            boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
                                        }}
                                    />
                                    <Bar dataKey="totalRevenue" fill="#10b981" name="Revenue" radius={[0, 8, 8, 0]} />
                                </BarChart>
                            </ResponsiveContainer>
                        </CardContent>
                    </Card>

                    <Card className="border-2 border-green-200 dark:border-green-900 shadow-lg">
                        <CardHeader className="bg-gradient-to-r from-green-50 to-emerald-50 dark:from-green-950/20 dark:to-emerald-950/20 border-b">
                            <CardTitle className="text-2xl">Regional Performance Details</CardTitle>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <Table>
                                <TableHeader>
                                    <TableRow className="bg-green-50 dark:bg-green-950/10 hover:bg-green-100 dark:hover:bg-green-950/20">
                                        <TableHead className="font-bold">Rank</TableHead>
                                        <TableHead className="font-bold">Region</TableHead>
                                        <TableHead className="font-bold">Country</TableHead>
                                        <TableHead className="text-right font-bold">Sales</TableHead>
                                        <TableHead className="text-right font-bold">Revenue</TableHead>
                                        <TableHead className="text-right font-bold">Quantity</TableHead>
                                        <TableHead className="text-right font-bold">Avg Sale</TableHead>
                                    </TableRow>
                                </TableHeader>
                                <TableBody>
                                    {regionSummary.map((region, idx) => (
                                        <TableRow key={region.region} className="hover:bg-green-50 dark:hover:bg-green-950/10 transition-colors">
                                            <TableCell>
                                                {idx < 3 ? (
                                                    <Badge className={
                                                        idx === 0 ? "bg-yellow-500" :
                                                        idx === 1 ? "bg-gray-400" :
                                                        "bg-orange-600"
                                                    }>
                                                        {idx + 1}
                                                    </Badge>
                                                ) : (
                                                    <span className="font-medium text-muted-foreground">{idx + 1}</span>
                                                )}
                                            </TableCell>
                                            <TableCell className="font-semibold">
                                                <div className="flex items-center gap-2">
                                                    <MapPin className="h-4 w-4 text-green-600" />
                                                    {region.region}
                                                </div>
                                            </TableCell>
                                            <TableCell>{region.country}</TableCell>
                                            <TableCell className="text-right">{region.totalSales}</TableCell>
                                            <TableCell className="text-right font-semibold text-green-600">
                                                ${region.totalRevenue.toLocaleString()}
                                            </TableCell>
                                            <TableCell className="text-right">{region.totalQuantity}</TableCell>
                                            <TableCell className="text-right">${region.avgRevenuePerSale.toFixed(2)}</TableCell>
                                        </TableRow>
                                    ))}
                                </TableBody>
                            </Table>
                        </CardContent>
                    </Card>
                </TabsContent>

                {/* Comparison Tab */}
                <TabsContent value="comparison" className="space-y-6">
                    <div className="grid gap-6 md:grid-cols-2">
                        {topCountries.slice(0, 6).map((country, idx) => (
                            <Card key={country.country} className="border-2 hover:border-blue-400 transition-all hover:shadow-xl bg-gradient-to-br from-white to-blue-50 dark:from-zinc-900 dark:to-blue-950/10">
                                <CardHeader className="bg-gradient-to-r from-blue-100 to-cyan-100 dark:from-blue-950/30 dark:to-cyan-950/30 border-b">
                                    <CardTitle className="flex items-center gap-2 text-xl">
                                        <Globe className="h-6 w-6 text-blue-600" />
                                        {country.country}
                                    </CardTitle>
                                    <CardDescription className="text-base font-semibold">
                                        Market share: {((country.totalRevenue / totalRevenue) * 100).toFixed(1)}%
                                    </CardDescription>
                                </CardHeader>
                                <CardContent className="space-y-4 pt-6">
                                    <div className="grid grid-cols-2 gap-4">
                                        <div className="bg-blue-50 dark:bg-blue-950/10 p-3 rounded-lg">
                                            <p className="text-sm text-muted-foreground font-semibold">Total Sales</p>
                                            <p className="text-2xl font-bold text-blue-600">{country.totalSales}</p>
                                        </div>
                                        <div className="bg-green-50 dark:bg-green-950/10 p-3 rounded-lg">
                                            <p className="text-sm text-muted-foreground font-semibold">Revenue</p>
                                            <p className="text-2xl font-bold text-green-600">${country.totalRevenue.toLocaleString()}</p>
                                        </div>
                                        <div className="bg-orange-50 dark:bg-orange-950/10 p-3 rounded-lg">
                                            <p className="text-sm text-muted-foreground font-semibold">Avg Sale</p>
                                            <p className="text-xl font-semibold text-orange-600">${country.avgRevenuePerSale.toFixed(2)}</p>
                                        </div>
                                        <div className="bg-purple-50 dark:bg-purple-950/10 p-3 rounded-lg">
                                            <p className="text-sm text-muted-foreground font-semibold">Unique Wines</p>
                                            <p className="text-xl font-semibold text-purple-600">{country.uniqueWines}</p>
                                        </div>
                                    </div>

                                    <div className="bg-gradient-to-br from-blue-50 to-transparent dark:from-blue-950/10 p-4 rounded-lg">
                                        <p className="text-sm font-bold mb-2 text-blue-700 dark:text-blue-400">Top Wine Categories</p>
                                        <div className="space-y-2">
                                            {country.topCategories.map((cat, idx) => (
                                                <div key={cat.name} className="flex justify-between items-center bg-white dark:bg-zinc-800 p-2 rounded-lg">
                                                    <span className="text-sm font-medium">{idx + 1}. {cat.name}</span>
                                                    <Badge variant="outline" className="bg-blue-50 dark:bg-blue-950/20">{cat.count} sales</Badge>
                                                </div>
                                            ))}
                                        </div>
                                    </div>

                                    {country.topRegions.length > 0 && (
                                        <div>
                                            <p className="text-sm font-bold mb-2 text-green-700 dark:text-green-400">Top Regions</p>
                                            <div className="flex flex-wrap gap-2">
                                                {country.topRegions.map(region => (
                                                    <Badge key={region.name} className="bg-gradient-to-r from-green-500 to-emerald-500 text-white">
                                                        {region.name}
                                                    </Badge>
                                                ))}
                                            </div>
                                        </div>
                                    )}
                                </CardContent>
                            </Card>
                        ))}
                    </div>
                </TabsContent>
            </Tabs>

            {/* Key Insights */}
            <Card className="border-2 border-cyan-200 dark:border-cyan-900 shadow-lg">
                <CardHeader className="bg-gradient-to-r from-cyan-50 to-blue-50 dark:from-cyan-950/20 dark:to-blue-950/20 border-b">
                    <CardTitle className="flex items-center gap-2 text-2xl">
                        <TrendingUp className="h-6 w-6 text-cyan-600" />
                        Key Insights
                    </CardTitle>
                </CardHeader>
                <CardContent className="pt-6">
                    <KeyInsights countrySummary={countrySummary} regionSummary={regionSummary} />
                </CardContent>
            </Card>
        </div>
    );
}

function KeyInsights({ countrySummary, regionSummary }: { countrySummary: CountrySummary[], regionSummary: RegionSummary[] }) {
    const insights: string[] = [];

    if (countrySummary.length > 0) {
        const top = countrySummary[0];
        insights.push(`${top.country} dominates the market with $${top.totalRevenue.toLocaleString()} in revenue (${top.totalSales} sales)`);
        
        if (top.topCategories.length > 0) {
            insights.push(`${top.country}'s best-selling category is ${top.topCategories[0].name}`);
        }
    }

    if (countrySummary.length >= 2) {
        const second = countrySummary[1];
        insights.push(`${second.country} ranks second with $${second.totalRevenue.toLocaleString()} in revenue`);
    }

    if (regionSummary.length > 0) {
        const topRegion = regionSummary[0];
        insights.push(`${topRegion.region} is the highest-performing region with $${topRegion.totalRevenue.toLocaleString()} revenue`);
    }

    const highAvgCountries = countrySummary.filter(c => c.avgRevenuePerSale > 50).slice(0, 2);
    if (highAvgCountries.length > 0) {
        insights.push(`Premium markets: ${highAvgCountries.map(c => c.country).join(', ')} have average sales over $50`);
    }

    return (
        <div className="space-y-3">
            {insights.map((insight, idx) => (
                <div key={idx} className="flex items-start gap-3 p-3 bg-cyan-50 dark:bg-cyan-950/10 rounded-lg hover:bg-cyan-100 dark:hover:bg-cyan-950/20 transition-colors">
                    <div className="h-2 w-2 rounded-full bg-cyan-500 mt-2 flex-shrink-0" />
                    <p className="text-sm font-medium">{insight}</p>
                </div>
            ))}
        </div>
    );
}
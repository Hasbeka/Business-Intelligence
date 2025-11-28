// components/ProductPerformance.tsx
"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Wine, TrendingUp, Users, DollarSign, Package, Link2, Award, ShoppingCart, Search, Download } from "lucide-react";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Input } from "@/components/ui/input";
import { useState } from "react";
import { Bar, BarChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { exportProductPerformanceToExcel } from "@/lib/productPerformanceExport";

type WineAssociation = {
    wine: string;
    wineCategory: string;
    wineCountry: string;
    totalCustomers: number;
    associations: Array<{
        associatedWine: string;
        associatedCategory: string;
        associatedCountry: string;
        count: number;
        confidence: number;
    }>;
};

type CategoryAssociation = {
    category: string;
    associations: Array<{ category: string; count: number }>;
};

type WinePerformance = {
    wine: string;
    category: string;
    country: string;
    totalSales: number;
    totalRevenue: number;
    totalQuantity: number;
    uniqueCustomers: number;
    avgSaleAmount: number;
};

type Props = {
    wineAssociations: WineAssociation[];
    categoryAssociations: CategoryAssociation[];
    winePerformance: WinePerformance[];
};

export default function ProductPerformance({ wineAssociations, categoryAssociations, winePerformance }: Props) {
    const [searchTerm, setSearchTerm] = useState("");
    const [isExporting, setIsExporting] = useState(false);

    const filteredAssociations = wineAssociations.filter(w =>
        w.wine.toLowerCase().includes(searchTerm.toLowerCase())
    );

    const totalWines = wineAssociations.length;
    const avgAssociations = wineAssociations.reduce((sum, w) => sum + w.associations.length, 0) / totalWines;
    const strongestAssociation = wineAssociations
        .flatMap(w => w.associations.map(a => ({ ...a, baseWine: w.wine })))
        .sort((a, b) => b.confidence - a.confidence)[0];

    const handleExportExcel = async () => {
        setIsExporting(true);
        try {
            await exportProductPerformanceToExcel(
                wineAssociations,
                categoryAssociations,
                winePerformance
            );
        } catch (error) {
            console.error('Export failed:', error);
            alert('Failed to export data. Please try again.');
        } finally {
            setIsExporting(false);
        }
    };

    return (
        <div className="p-6 space-y-6 bg-linear-to-br from-zinc-50 to-zinc-100 dark:from-black dark:to-zinc-900 min-h-screen">
            {/* Header with Export Button */}
            <div className="flex items-start justify-between gap-4">
                <div className="space-y-2">
                    <h1 className="text-4xl font-bold tracking-tight bg-linear-to-r from-purple-600 to-pink-600 bg-clip-text text-transparent flex items-center gap-3">
                        <Wine className="h-10 w-10 text-purple-600" />
                        Product Performance & Associations
                    </h1>
                    <p className="text-lg text-muted-foreground">
                        Discover wine purchase patterns and product associations
                    </p>
                </div>
                <Button
                    onClick={handleExportExcel}
                    disabled={isExporting}
                    size="lg"
                    className="bg-gradient-to-r from-green-600 to-emerald-600 hover:from-green-700 hover:to-emerald-700 text-white shadow-lg hover:shadow-xl transition-all"
                >
                    <Download className="mr-2 h-5 w-5" />
                    {isExporting ? 'Exporting...' : 'Export to Excel'}
                </Button>
            </div>

            {/* Key Metrics */}
            <div className="grid gap-4 md:grid-cols-4">
                <Card className="border-2 border-purple-200 dark:border-purple-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-linear-to-br from-purple-50 to-transparent dark:from-purple-950/20">
                        <CardTitle className="text-sm font-medium">Wines Analyzed</CardTitle>
                        <div className="p-2 bg-purple-100 dark:bg-purple-900 rounded-lg">
                            <Wine className="h-4 w-4 text-purple-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-3xl font-bold text-purple-600">{totalWines}</div>
                        <p className="text-xs text-muted-foreground mt-1">With purchase associations</p>
                    </CardContent>
                </Card>

                <Card className="border-2 border-pink-200 dark:border-pink-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-linear-to-br from-pink-50 to-transparent dark:from-pink-950/20">
                        <CardTitle className="text-sm font-medium">Avg Associations</CardTitle>
                        <div className="p-2 bg-pink-100 dark:bg-pink-900 rounded-lg">
                            <Link2 className="h-4 w-4 text-pink-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-3xl font-bold text-pink-600">{avgAssociations.toFixed(1)}</div>
                        <p className="text-xs text-muted-foreground mt-1">Per wine product</p>
                    </CardContent>
                </Card>

                <Card className="border-2 border-yellow-200 dark:border-yellow-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-linear-to-br from-yellow-50 to-transparent dark:from-yellow-950/20">
                        <CardTitle className="text-sm font-medium">Top Performance</CardTitle>
                        <div className="p-2 bg-yellow-100 dark:bg-yellow-900 rounded-lg">
                            <Award className="h-4 w-4 text-yellow-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-2xl font-bold text-yellow-600">{winePerformance[0]?.wine.slice(0, 15)}...</div>
                        <p className="text-xs text-muted-foreground mt-1">
                            ${winePerformance[0]?.totalRevenue.toLocaleString()} revenue
                        </p>
                    </CardContent>
                </Card>

                <Card className="border-2 border-green-200 dark:border-green-900 hover:shadow-xl transition-all">
                    <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2 bg-linear-to-br from-green-50 to-transparent dark:from-green-950/20">
                        <CardTitle className="text-sm font-medium">Strongest Link</CardTitle>
                        <div className="p-2 bg-green-100 dark:bg-green-900 rounded-lg">
                            <TrendingUp className="h-4 w-4 text-green-600" />
                        </div>
                    </CardHeader>
                    <CardContent>
                        <div className="text-3xl font-bold text-green-600">{strongestAssociation?.confidence.toFixed(0)}%</div>
                        <p className="text-xs text-muted-foreground mt-1">Confidence score</p>
                    </CardContent>
                </Card>
            </div>

            <Tabs defaultValue="associations" className="space-y-6">
                <TabsList className="grid w-full max-w-md grid-cols-3 bg-white dark:bg-zinc-800 p-1 shadow-sm">
                    <TabsTrigger 
                        value="associations"
                        className="data-[state=active]:bg-linear-to-r data-[state=active]:from-purple-600 data-[state=active]:to-pink-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        Wine Associations
                    </TabsTrigger>
                    <TabsTrigger 
                        value="categories"
                        className="data-[state=active]:bg-linear-to-r data-[state=active]:from-purple-600 data-[state=active]:to-pink-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        Category Patterns
                    </TabsTrigger>
                    <TabsTrigger 
                        value="performance"
                        className="data-[state=active]:bg-linear-to-r data-[state=active]:from-purple-600 data-[state=active]:to-pink-600 data-[state=active]:text-white data-[state=active]:shadow-lg transition-all"
                    >
                        Top Performers
                    </TabsTrigger>
                </TabsList>

                {/* Wine Associations Tab */}
                <TabsContent value="associations" className="space-y-6">
                    {/* Search */}
                    <Card className="border-2 border-purple-200 dark:border-purple-900 shadow-lg">
                        <CardHeader className="bg-linear-to-r from-purple-50 to-pink-50 dark:from-purple-950/20 dark:to-pink-950/20 border-b">
                            <CardTitle className="text-2xl flex items-center gap-2">
                                <Search className="h-6 w-6 text-purple-600" />
                                Search Wine Associations
                            </CardTitle>
                            <CardDescription className="text-base">Find what customers buy together with specific wines</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <div className="relative">
                                <Search className="absolute left-3 top-3 h-4 w-4 text-muted-foreground" />
                                <Input
                                    placeholder="Search for a wine (e.g., Murfatlar, Bucium)..."
                                    value={searchTerm}
                                    onChange={(e) => setSearchTerm(e.target.value)}
                                    className="pl-10 max-w-md border-2 focus:border-purple-500"
                                />
                            </div>
                        </CardContent>
                    </Card>

                    {/* Associations List */}
                    <div className="grid gap-6">
                        {filteredAssociations.slice(0, 20).map((wine) => (
                            <Card key={wine.wine} className="hover:shadow-xl transition-all border-2 hover:border-purple-300 bg-linear-to-br from-white to-purple-50 dark:from-zinc-900 dark:to-purple-950/10">
                                <CardHeader className="bg-linear-to-r from-purple-100 to-pink-100 dark:from-purple-950/30 dark:to-pink-950/30 border-b">
                                    <div className="flex items-start justify-between">
                                        <div className="space-y-2">
                                            <CardTitle className="flex items-center gap-2 text-xl">
                                                <div className="p-2 bg-purple-200 dark:bg-purple-800 rounded-lg">
                                                    <Wine className="h-5 w-5 text-purple-600" />
                                                </div>
                                                {wine.wine}
                                            </CardTitle>
                                            <CardDescription className="flex items-center gap-3 text-base">
                                                <Badge className="bg-linear-to-r from-purple-500 to-pink-500 text-white">
                                                    {wine.wineCategory}
                                                </Badge>
                                                <span className="text-sm font-medium">{wine.wineCountry}</span>
                                                <span className="text-sm flex items-center gap-1 font-semibold">
                                                    <Users className="h-4 w-4" />
                                                    {wine.totalCustomers} customers
                                                </span>
                                            </CardDescription>
                                        </div>
                                    </div>
                                </CardHeader>
                                <CardContent className="pt-6">
                                    <div className="space-y-4">
                                        <div className="flex items-center gap-2 text-sm font-bold text-purple-700 dark:text-purple-400">
                                            <Link2 className="h-5 w-5" />
                                            Customers who bought this also bought:
                                        </div>
                                        <div className="space-y-3">
                                            {wine.associations.map((assoc, idx) => (
                                                <div
                                                    key={assoc.associatedWine}
                                                    className="flex items-center justify-between p-4 rounded-lg bg-white dark:bg-zinc-800 hover:bg-purple-50 dark:hover:bg-purple-950/20 transition-all border-2 border-transparent hover:border-purple-300 shadow-sm"
                                                >
                                                    <div className="flex-1">
                                                        <div className="flex items-center gap-3">
                                                            <span className="flex items-center justify-center w-7 h-7 rounded-full bg-linear-to-r from-purple-600 to-pink-600 text-white text-sm font-bold">
                                                                {idx + 1}
                                                            </span>
                                                            <span className="font-semibold text-base">
                                                                {assoc.associatedWine}
                                                            </span>
                                                        </div>
                                                        <div className="flex items-center gap-2 mt-2 ml-10">
                                                            <Badge variant="outline" className="text-xs bg-purple-50 dark:bg-purple-950/20">
                                                                {assoc.associatedCategory}
                                                            </Badge>
                                                            <span className="text-xs text-muted-foreground font-medium">
                                                                {assoc.associatedCountry}
                                                            </span>
                                                        </div>
                                                    </div>
                                                    <div className="flex flex-col items-end gap-2">
                                                        <Badge className="bg-linear-to-r from-purple-600 to-pink-600 text-white text-sm px-3 py-1">
                                                            {assoc.confidence.toFixed(0)}% confidence
                                                        </Badge>
                                                        <span className="text-xs text-muted-foreground font-semibold">
                                                            {assoc.count} co-purchases
                                                        </span>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                </CardContent>
                            </Card>
                        ))}
                    </div>

                    {filteredAssociations.length === 0 && searchTerm && (
                        <Card className="border-2 border-purple-200 dark:border-purple-900">
                            <CardContent className="p-12 text-center">
                                <Wine className="h-12 w-12 text-muted-foreground mx-auto mb-4" />
                                <p className="text-lg text-muted-foreground">No wines found matching "{searchTerm}"</p>
                            </CardContent>
                        </Card>
                    )}
                </TabsContent>

                {/* Category Associations Tab */}
                <TabsContent value="categories" className="space-y-6">
                    <Card className="border-2 border-pink-200 dark:border-pink-900 shadow-lg">
                        <CardHeader className="bg-linear-to-r from-pink-50 to-purple-50 dark:from-pink-950/20 dark:to-purple-950/20 border-b">
                            <CardTitle className="text-2xl">Wine Category Purchase Patterns</CardTitle>
                            <CardDescription className="text-base">
                                Which wine categories are frequently purchased together
                            </CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <div className="space-y-6">
                                {categoryAssociations.map((cat) => (
                                    <div key={cat.category} className="space-y-3 p-4 bg-linear-to-br from-pink-50 to-transparent dark:from-pink-950/10 rounded-lg">
                                        <div className="flex items-center gap-3">
                                            <Badge className="text-lg px-4 py-2 bg-linear-to-r from-pink-600 to-purple-600 text-white">
                                                {cat.category}
                                            </Badge>
                                            <span className="text-base font-semibold text-muted-foreground">is often bought with:</span>
                                        </div>
                                        <div className="grid gap-3 md:grid-cols-3">
                                            {cat.associations.map((assoc, idx) => (
                                                <Card key={assoc.category} className="border-2 hover:border-pink-400 transition-all hover:shadow-md">
                                                    <CardContent className="p-4">
                                                        <div className="flex items-center justify-between">
                                                            <div className="flex items-center gap-3">
                                                                <span className="flex items-center justify-center w-8 h-8 rounded-full bg-linear-to-r from-pink-600 to-purple-600 text-white font-bold">
                                                                    {idx + 1}
                                                                </span>
                                                                <span className="font-semibold">{assoc.category}</span>
                                                            </div>
                                                            <Badge variant="secondary" className="bg-pink-100 dark:bg-pink-950/20">
                                                                {assoc.count} times
                                                            </Badge>
                                                        </div>
                                                    </CardContent>
                                                </Card>
                                            ))}
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </CardContent>
                    </Card>

                    {/* Category Chart */}
                    <Card className="border-2 border-pink-200 dark:border-pink-900 shadow-lg">
                        <CardHeader className="bg-linear-to-r from-pink-50 to-purple-50 dark:from-pink-950/20 dark:to-purple-950/20 border-b">
                            <CardTitle className="text-2xl">Category Co-Purchase Frequency</CardTitle>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <ResponsiveContainer width="100%" height={400}>
                                <BarChart
                                    data={categoryAssociations.map(c => ({
                                        category: c.category,
                                        total: c.associations.reduce((sum, a) => sum + a.count, 0)
                                    }))}
                                >
                                    <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                                    <XAxis dataKey="category" angle={-45} textAnchor="end" height={100} />
                                    <YAxis />
                                    <Tooltip 
                                        contentStyle={{ 
                                            backgroundColor: 'white', 
                                            border: '2px solid #ec4899',
                                            borderRadius: '8px',
                                            boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
                                        }}
                                    />
                                    <Bar dataKey="total" fill="#ec4899" name="Co-purchases" radius={[8, 8, 0, 0]} />
                                </BarChart>
                            </ResponsiveContainer>
                        </CardContent>
                    </Card>
                </TabsContent>

                {/* Top Performers Tab */}
                <TabsContent value="performance" className="space-y-6">
                    <Card className="border-2 border-yellow-200 dark:border-yellow-900 shadow-lg">
                        <CardHeader className="bg-linear-to-r from-yellow-50 to-orange-50 dark:from-yellow-950/20 dark:to-orange-950/20 border-b">
                            <CardTitle className="text-2xl">Top 20 Best-Selling Wines</CardTitle>
                            <CardDescription className="text-base">Ranked by total revenue generated</CardDescription>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <Table>
                                <TableHeader>
                                    <TableRow className="bg-yellow-50 dark:bg-yellow-950/10 hover:bg-yellow-100 dark:hover:bg-yellow-950/20">
                                        <TableHead className="w-12 font-bold">Rank</TableHead>
                                        <TableHead className="font-bold">Wine</TableHead>
                                        <TableHead className="font-bold">Category</TableHead>
                                        <TableHead className="text-right font-bold">Sales</TableHead>
                                        <TableHead className="text-right font-bold">Revenue</TableHead>
                                        <TableHead className="text-right font-bold">Customers</TableHead>
                                        <TableHead className="text-right font-bold">Avg Sale</TableHead>
                                    </TableRow>
                                </TableHeader>
                                <TableBody>
                                    {winePerformance.map((wine, idx) => (
                                        <TableRow key={wine.wine} className="hover:bg-yellow-50 dark:hover:bg-yellow-950/10 transition-colors">
                                            <TableCell>
                                                {idx + 1 <= 3 ? (
                                                    <Badge className={
                                                        idx === 0 ? "bg-yellow-500 text-lg px-3 py-1" :
                                                        idx === 1 ? "bg-gray-400 text-lg px-3 py-1" :
                                                        "bg-orange-600 text-lg px-3 py-1"
                                                    }>
                                                        {idx + 1}
                                                    </Badge>
                                                ) : (
                                                    <span className="text-muted-foreground font-semibold">{idx + 1}</span>
                                                )}
                                            </TableCell>
                                            <TableCell className="font-semibold">{wine.wine}</TableCell>
                                            <TableCell>
                                                <Badge className="bg-linear-to-r from-purple-500 to-pink-500 text-white">
                                                    {wine.category}
                                                </Badge>
                                            </TableCell>
                                            <TableCell className="text-right font-medium">{wine.totalSales}</TableCell>
                                            <TableCell className="text-right font-bold text-green-600">
                                                ${wine.totalRevenue.toLocaleString()}
                                            </TableCell>
                                            <TableCell className="text-right font-medium">{wine.uniqueCustomers}</TableCell>
                                            <TableCell className="text-right font-medium">
                                                ${wine.avgSaleAmount.toFixed(2)}
                                            </TableCell>
                                        </TableRow>
                                    ))}
                                </TableBody>
                            </Table>
                        </CardContent>
                    </Card>

                    {/* Performance Chart */}
                    <Card className="border-2 border-yellow-200 dark:border-yellow-900 shadow-lg">
                        <CardHeader className="bg-linear-to-r from-yellow-50 to-orange-50 dark:from-yellow-950/20 dark:to-orange-950/20 border-b">
                            <CardTitle className="text-2xl">Revenue Comparison - Top 10 Wines</CardTitle>
                        </CardHeader>
                        <CardContent className="pt-6">
                            <ResponsiveContainer width="100%" height={400}>
                                <BarChart data={winePerformance.slice(0, 10)}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                                    <XAxis dataKey="wine" angle={-45} textAnchor="end" height={150} />
                                    <YAxis />
                                    <Tooltip 
                                        formatter={(value) => `$${Number(value).toLocaleString()}`}
                                        contentStyle={{ 
                                            backgroundColor: 'white', 
                                            border: '2px solid #f59e0b',
                                            borderRadius: '8px',
                                            boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
                                        }}
                                    />
                                    <Bar dataKey="totalRevenue" fill="#f59e0b" name="Revenue" radius={[8, 8, 0, 0]} />
                                </BarChart>
                            </ResponsiveContainer>
                        </CardContent>
                    </Card>
                </TabsContent>
            </Tabs>

            {/* Key Insights */}
            <Card className="border-2 border-purple-200 dark:border-purple-900 shadow-lg">
                <CardHeader className="bg-linear-to-r from-purple-50 to-pink-50 dark:from-purple-950/20 dark:to-pink-950/20 border-b">
                    <CardTitle className="flex items-center gap-2 text-2xl">
                        <ShoppingCart className="h-6 w-6 text-purple-600" />
                        Key Purchase Insights
                    </CardTitle>
                </CardHeader>
                <CardContent className="pt-6">
                    <KeyInsights 
                        wineAssociations={wineAssociations}
                        strongestAssociation={strongestAssociation}
                        winePerformance={winePerformance}
                    />
                </CardContent>
            </Card>
        </div>
    );
}

function KeyInsights({ 
    wineAssociations, 
    strongestAssociation,
    winePerformance 
}: { 
    wineAssociations: WineAssociation[];
    strongestAssociation: any;
    winePerformance: WinePerformance[];
}) {
    const insights: string[] = [];

    if (strongestAssociation) {
        insights.push(
            `Strongest association: ${strongestAssociation.confidence.toFixed(0)}% of customers who bought ${strongestAssociation.baseWine} also purchased ${strongestAssociation.associatedWine}`
        );
    }

    const highConfidence = wineAssociations
        .flatMap(w => w.associations.map(a => ({ ...a, base: w.wine })))
        .filter(a => a.confidence >= 50)
        .length;
    
    if (highConfidence > 0) {
        insights.push(`Found ${highConfidence} wine pairs with 50%+ purchase correlation`);
    }

    if (winePerformance.length >= 3) {
        insights.push(
            `Top 3 wines (${winePerformance.slice(0, 3).map(w => w.wine).join(', ')}) account for $${winePerformance.slice(0, 3).reduce((sum, w) => sum + w.totalRevenue, 0).toLocaleString()} in revenue`
        );
    }

    const crossCategory = wineAssociations.filter(w => 
        w.associations.some(a => a.associatedCategory !== w.wineCategory)
    ).length;
    
    insights.push(`${crossCategory} wines have cross-category purchase patterns`);

    return (
        <div className="space-y-3">
            {insights.map((insight, idx) => (
                <div key={idx} className="flex items-start gap-3 p-3 bg-purple-50 dark:bg-purple-950/10 rounded-lg hover:bg-purple-100 dark:hover:bg-purple-950/20 transition-colors">
                    <div className="h-2 w-2 rounded-full bg-purple-500 mt-2 shrink-0" />
                    <p className="text-sm font-medium">{insight}</p>
                </div>
            ))}
        </div>
    );
}
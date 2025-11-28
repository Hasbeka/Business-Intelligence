// app/customer-insights/page.tsx
"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { UsersRound, Globe, ArrowRight, Users, TrendingUp } from "lucide-react";
import Link from "next/link";

const page = () => {
    return (
        <div className="min-h-screen w-full bg-zinc-50 font-sans dark:bg-black p-6">
            <div className="max-w-7xl mx-auto space-y-6">
                {/* Header */}
                <div className="space-y-2">
                    <h1 className="text-4xl font-bold tracking-tight bg-linear-to-r from-purple-600 to-pink-600 bg-clip-text text-transparent flex items-center gap-3">
                        <Users className="h-10 w-10 text-purple-600" />
                        Customer and Product Insights
                    </h1>
                    <p className="text-lg text-muted-foreground">
                        Explore customer behavior patterns and wine preferences
                    </p>
                </div>

                {/* Main Cards */}
                <div className="grid gap-6 md:grid-cols-2 mt-8">
                    {/* Customer Segmentation Card */}
                    <Card className="relative overflow-hidden group hover:shadow-xl transition-all duration-300 border-2 hover:border-purple-500">
                        <div className="absolute top-0 right-0 w-32 h-32 bg-linear-to-brrom-purple-500/10 to-pink-500/10 rounded-bl-full" />
                        <CardHeader className="relative">
                            <div className="flex items-start justify-between">
                                <div className="space-y-2">
                                    <div className="flex items-center gap-2">
                                        <div className="p-2 bg-purple-100 dark:bg-purple-900/20 rounded-lg">
                                            <UsersRound className="h-6 w-6 text-purple-600" />
                                        </div>
                                        <CardTitle className="text-2xl">Customer Segmentation</CardTitle>
                                    </div>
                                    <CardDescription className="text-base">
                                        Identify customer segments based on demographics and wine preferences
                                    </CardDescription>
                                </div>
                            </div>
                        </CardHeader>
                        <CardContent className="relative space-y-4">
                            <div className="space-y-2">
                                <h4 className="font-semibold text-sm text-muted-foreground">ANALYSIS INCLUDES:</h4>
                                <ul className="space-y-2">
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-purple-500 mt-2" />
                                        <span className="text-sm">Gender-based wine preferences (e.g., males prefer dry varieties)</span>
                                    </li>
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-purple-500 mt-2" />
                                        <span className="text-sm">Age group analysis (e.g., customers over 55 prefer red wines)</span>
                                    </li>
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-purple-500 mt-2" />
                                        <span className="text-sm">Combined demographic patterns and insights</span>
                                    </li>
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-purple-500 mt-2" />
                                        <span className="text-sm">Price range and variety preferences by segment</span>
                                    </li>
                                </ul>
                            </div>
                            <Link href="/customer-segmentation">
                                <Button className="w-full bg-linear-to-r from-purple-600 to-pink-600 hover:from-purple-700 hover:to-pink-700 text-white group-hover:shadow-lg transition-all duration-300">
                                    Explore Segmentation
                                    <ArrowRight className="ml-2 h-4 w-4 group-hover:translate-x-1 transition-transform" />
                                </Button>
                            </Link>
                        </CardContent>
                    </Card>

                    {/* Wine Origin Analysis Card */}
                    <Card className="relative overflow-hidden group hover:shadow-xl transition-all duration-300 border-2 hover:border-blue-500">
                        <div className="absolute top-0 right-0 w-32 h-32 bg-linear-to-br from-blue-500/10 to-cyan-500/10 rounded-bl-full" />
                        <CardHeader className="relative">
                            <div className="flex items-start justify-between">
                                <div className="space-y-2">
                                    <div className="flex items-center gap-2">
                                        <div className="p-2 bg-blue-100 dark:bg-blue-900/20 rounded-lg">
                                            <Globe className="h-6 w-6 text-blue-600" />
                                        </div>
                                        <CardTitle className="text-2xl">Wine Origin Analysis</CardTitle>
                                    </div>
                                    <CardDescription className="text-base">
                                        Discover the relationship between wine country origin and sales data
                                    </CardDescription>
                                </div>
                            </div>
                        </CardHeader>
                        <CardContent className="relative space-y-4">
                            <div className="space-y-2">
                                <h4 className="font-semibold text-sm text-muted-foreground">ANALYSIS INCLUDES:</h4>
                                <ul className="space-y-2">
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-blue-500 mt-2" />
                                        <span className="text-sm">Sales performance by wine country of origin</span>
                                    </li>
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-blue-500 mt-2" />
                                        <span className="text-sm">Top-selling wine regions and countries</span>
                                    </li>
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-blue-500 mt-2" />
                                        <span className="text-sm">Revenue comparison across different origins</span>
                                    </li>
                                    <li className="flex items-start gap-2">
                                        <div className="h-1.5 w-1.5 rounded-full bg-blue-500 mt-2" />
                                        <span className="text-sm">Regional trends and market insights</span>
                                    </li>
                                </ul>
                            </div>
                            <Link href="/wine-origin-analysis">
                                <Button className="w-full bg-linear-to-r from-blue-600 to-cyan-600 hover:from-blue-700 hover:to-cyan-700 text-white group-hover:shadow-lg transition-all duration-300">
                                    Explore Wine Origins
                                    <ArrowRight className="ml-2 h-4 w-4 group-hover:translate-x-1 transition-transform" />
                                </Button>
                            </Link>
                        </CardContent>
                    </Card>
                </div>

                {/* Info Section */}
                <Card className="bg-linear-to-br from-zinc-50 to-zinc-100 dark:from-zinc-900 dark:to-zinc-800 border-zinc-200 dark:border-zinc-700">
                    <CardContent className="p-6">
                        <div className="flex items-start gap-4">
                            <div className="p-3 bg-white dark:bg-zinc-800 rounded-lg shadow-sm">
                                <TrendingUp className="h-6 w-6 text-green-600" />
                            </div>
                            <div className="space-y-1">
                                <h3 className="font-semibold text-lg">Data-Driven Insights</h3>
                                <p className="text-sm text-muted-foreground">
                                    These analyses combine customer demographics, purchase history, and wine characteristics 
                                    to provide actionable insights for marketing strategies and inventory management.
                                </p>
                            </div>
                        </div>
                    </CardContent>
                </Card>
            </div>
        </div>
    );
};

export default page;
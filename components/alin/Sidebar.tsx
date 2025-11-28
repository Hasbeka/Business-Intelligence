'use client';

import React, { useEffect } from 'react'
import { Sidebar, SidebarContent, SidebarGroup, SidebarGroupContent, SidebarGroupLabel, SidebarMenu, SidebarMenuButton, SidebarMenuItem } from '@/components/ui/sidebar'
import { Wine, BarChart3, TrendingUp, Users, UsersRound } from 'lucide-react'
import Link from 'next/link';
import { usePathname } from 'next/navigation';


const items = [
    { label: 'Dashboard', href: '/', icon: BarChart3 },
    { label: 'Sales Analysis', href: '/sales-analysis', icon: TrendingUp },
    { label: 'Customer & Wine Stats', href: '/customer-insights', icon: Users },
    { label: 'Product Performance', href: '/product-performance', icon: Wine },
]

const AppSidebar = () => {

    const [activeItem, setActiveItem] = React.useState('Dashboard')
    const urlElement = usePathname();

    useEffect(() => {
        setActiveItem(urlElement)
    }, [urlElement])

    return (
        <Sidebar className="border-r border-neutral-700 bg-black z-20 ">
            <SidebarContent className="px-4 py-6">
                <SidebarGroup className="space-y-8">
                    {/* Header with wine accent */}
                    <div className="space-y-2">
                        <div className="flex items-center gap-2 mb-4">
                            <div className="w-10 h-10 rounded-lg bg-liniar-to-br from-red-900 to-purple-900 flex items-center justify-center">
                                <Wine className="w-6 h-6 text-white" />
                            </div>
                            <SidebarGroupLabel className="text-xl font-bold text-gray-100 leading-tight flex flex-col">
                                Wine Sales
                                <span className="block text-sm font-normal text-gray-500">Decision Support</span>
                            </SidebarGroupLabel>
                        </div>
                        <div className="h-px bg-liniar-to-r from-red-900/20 via-purple-900/20 to-transparent"></div>
                    </div>

                    {/* Navigation Menu */}
                    <SidebarGroupContent>
                        <SidebarMenu className="space-y-2">
                            {items.map((item) => {
                                const Icon = item.icon
                                const isActive = activeItem === item.href

                                return (
                                    <SidebarMenuItem key={item.label}>
                                        <SidebarMenuButton
                                            asChild
                                            className={`
                                                group relative overflow-hidden rounded-lg transition-all duration-200
                                                ${isActive
                                                    ? 'bg-linear-to-r from-red-900 to-purple-900 text-white shadow-lg shadow-red-900/20'
                                                    : 'text-gray-200 hover:bg-neutral-700 hover:text-neutral-50'
                                                }
                                            `}
                                        >
                                            <Link
                                                href={item.href}
                                                onClick={(e) => {

                                                    setActiveItem(item.label)
                                                }}
                                                className="flex items-center gap-3 px-4 py-3"
                                            >
                                                <Icon className={`w-5 h-5 transition-transform duration-200 ${isActive ? 'scale-110' : 'group-hover:scale-105'}`} />
                                                <span className="font-medium">{item.label}</span>
                                                {isActive && (
                                                    <div className="absolute right-0 top-1/2 -translate-y-1/2 w-1 h-8 bg-white rounded-l-full"></div>
                                                )}
                                            </Link>
                                        </SidebarMenuButton>
                                    </SidebarMenuItem>
                                )
                            })}
                        </SidebarMenu>
                    </SidebarGroupContent>
                </SidebarGroup>

                {/* Footer accent */}
                <div className="mt-auto pt-6">
                    <div className="px-4 py-3 rounded-lg bg-liniar-to-br from-red-50 to-purple-50 border border-red-100">
                        <p className="text-xs text-gray-600 font-medium">Analytics Dashboard</p>
                        <p className="text-xs text-gray-500 mt-1">Powered by data insights</p>
                    </div>
                </div>
            </SidebarContent>
        </Sidebar>
    )
}

export default AppSidebar
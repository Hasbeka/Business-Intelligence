"use client";

import { ColumnDef } from "@tanstack/react-table"
import { SalesBetterFormat } from '@/app/types';
import { DataTable } from "../ui/data-table";
import ExportButton from "@/components/ui/ExportButton";
import { exportToExcel, generateFilename } from "@/lib/ExcelExport";

type GridProps = {
    sales: SalesBetterFormat[]
}

import { ArrowUpDown } from "lucide-react"
import { Button } from "@/components/ui/button"
import { parseDate } from "@/lib/utils";


export const columns: ColumnDef<SalesBetterFormat>[] = [
    {
        accessorKey: "saleID",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    SaleID
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
    },
    {
        accessorKey: "customerName",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Customer
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
    },
    {
        accessorKey: "wineDesignation",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Wine Name
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
    },
    {
        accessorKey: "quantity",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Quantity
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
    },
    {
        accessorKey: "saleAmount",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Amount
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
        cell: ({ row }) => {
            const amount = parseFloat(row.getValue("saleAmount"))
            const formatted = new Intl.NumberFormat("en-US", {
                style: "currency",
                currency: "USD",
            }).format(amount)

            return <div className="text-right font-medium">{formatted}</div>
        },
    },

    {
        accessorKey: "saleDate",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Date
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
        sortingFn: (rowA, rowB, columnId) => {
            const dateA = parseDate(rowA.getValue(columnId));
            const dateB = parseDate(rowB.getValue(columnId));
            if (!dateA || !dateB) return 0;
            return dateA.getTime() - dateB.getTime();
        },

        filterFn: (row, columnId, filterValue) => {

            if (typeof filterValue === 'string') {
                const dateStr = row.getValue(columnId) as string;
                const date = parseDate(dateStr);
                if (!date) return false;


                const formatted = date.toLocaleDateString("en-US", {
                    year: 'numeric',
                    month: 'short',
                    day: 'numeric'
                });

                return formatted.toLowerCase().includes(filterValue.toLowerCase()) ||
                    dateStr.toLowerCase().includes(filterValue.toLowerCase());
            } else {

                const dateStr = row.getValue(columnId) as string;
                const date = parseDate(dateStr);

                if (!date) return false;

                const { from, to } = filterValue as { from: string; to: string };

                if (!from && !to) return true;

                date.setHours(0, 0, 0, 0);

                const fromDate = from ? new Date(from) : null;
                const toDate = to ? new Date(to) : null;

                if (fromDate) {
                    fromDate.setHours(0, 0, 0, 0);
                }
                if (toDate) {
                    toDate.setHours(0, 0, 0, 0);
                }

                if (fromDate && date < fromDate) return false;
                if (toDate && date > toDate) return false;

                return true;
            }
        }
    },
    {
        accessorKey: "wineCategory",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Wine Category
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
    },
    {
        accessorKey: "customerState",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Customer State
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
    },
    {
        accessorKey: "wineCountry",
        header: ({ column }) => {
            return (
                <Button
                    variant="ghost"
                    onClick={() => column.toggleSorting(column.getIsSorted() === "asc")}
                >
                    Wine Origin Country
                    <ArrowUpDown className="ml-2 h-4 w-4" />
                </Button>
            )
        },
    },

]

const GridDataSales = (props: GridProps) => {

    const handleExport = async (
        filteredData: SalesBetterFormat[],
        tableState?: {
            visibleColumns: { id: string; visible: boolean }[];
            filters: any;
            sorting: any;
        }
    ) => {
        // Calculate summaries
        const totalAmount = filteredData.reduce((sum, sale) => sum + sale.saleAmount, 0);
        const totalQuantity = filteredData.reduce((sum, sale) => sum + sale.quantity, 0);
        const avgAmount = filteredData.length > 0 ? totalAmount / filteredData.length : 0;

        // Define all possible columns
        const allColumns = [
            { header: 'Sale ID', key: 'saleID', width: 12 },
            { header: 'Customer Name', key: 'customerName', width: 25 },
            { header: 'Wine Designation', key: 'wineDesignation', width: 30 },
            { header: 'Quantity', key: 'quantity', width: 12 },
            { header: 'Sale Amount ($)', key: 'saleAmount', width: 15 },
            { header: 'Sale Date', key: 'saleDate', width: 15 },
            { header: 'Wine Category', key: 'wineCategory', width: 18 },
            { header: 'Customer State', key: 'customerState', width: 18 },
            { header: 'Wine Origin Country', key: 'wineCountry', width: 20 },
        ];

        // Filter to only visible columns if tableState is provided
        const columnsToExport = tableState?.visibleColumns
            ? allColumns.filter(col =>
                tableState.visibleColumns.some(vc => vc.id === col.key && vc.visible)
            )
            : allColumns;

        // Build summary with active filters info
        const summary: any[] = [
            { label: 'Total Records', value: filteredData.length.toString(), style: 'info' },
            { label: 'Total Revenue', value: `$${totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, style: 'success' },
            { label: 'Total Quantity Sold', value: totalQuantity.toString(), style: 'info' },
            { label: 'Average Sale Amount', value: `$${avgAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, style: 'info' },
        ];

        // Add active filters to summary
        if (tableState?.filters && tableState.filters.length > 0) {
            summary.push({ label: '---', value: '---', style: 'info' });
            summary.push({ label: 'Active Filters', value: tableState.filters.length.toString(), style: 'warning' });

            tableState.filters.forEach((filter: any) => {
                let filterValue = '';
                if (Array.isArray(filter.value)) {
                    filterValue = filter.value.join(', ');
                } else if (typeof filter.value === 'object' && filter.value.from) {
                    filterValue = `${filter.value.from || 'start'} to ${filter.value.to || 'end'}`;
                } else {
                    filterValue = String(filter.value);
                }
                summary.push({
                    label: `  ${filter.id}`,
                    value: filterValue,
                    style: 'info'
                });
            });
        }

        // Add sorting info
        if (tableState?.sorting && tableState.sorting.length > 0) {
            summary.push({ label: '---', value: '---', style: 'info' });
            summary.push({ label: 'Sort Order', value: '', style: 'info' });
            tableState.sorting.forEach((sort: any, index: number) => {
                summary.push({
                    label: `  ${index + 1}. ${sort.id}`,
                    value: sort.desc ? 'Descending' : 'Ascending',
                    style: 'info'
                });
            });
        }

        await exportToExcel({
            filename: generateFilename('sales-data'),
            sheets: [
                {
                    name: 'Sales Data',
                    columns: columnsToExport,
                    data: filteredData,
                    summary: summary
                }
            ]
        });
    };

    return (
        <div>
            <DataTable
                columns={columns}
                data={props.sales}
                onExport={handleExport}
            />
        </div>
    )
}

export default GridDataSales
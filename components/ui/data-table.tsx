"use client"

import {
    ColumnDef,
    SortingState,
    getSortedRowModel,
    ColumnFiltersState,
    flexRender,
    getCoreRowModel,
    getFilteredRowModel,
    useReactTable,
    getPaginationRowModel,
    VisibilityState
} from "@tanstack/react-table";

import {
    Select,
    SelectContent,
    SelectGroup,
    SelectItem,
    SelectLabel,
    SelectTrigger,
    SelectValue,
} from "@/components/ui/select"

import {
    DropdownMenu,
    DropdownMenuCheckboxItem,
    DropdownMenuContent,
    DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu"

import {
    Popover,
    PopoverContent,
    PopoverTrigger,
} from "@/components/ui/popover"

import {
    Table,
    TableBody,
    TableCell,
    TableHead,
    TableHeader,
    TableRow,
} from "@/components/ui/table";
import { Input } from "@/components/ui/input"
import { Button } from "@/components/ui/button";
import { Slider } from "@/components/ui/slider";
import { Checkbox } from "@/components/ui/checkbox";
import ExportButton from "@/components/ui/ExportButton";
import React from "react";
import { ChevronDown, ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight, Filter, X, Calendar } from "lucide-react";

interface DataTableProps<TData, TValue> {
    columns: ColumnDef<TData, TValue>[]
    data: TData[]
    onExport?: (filteredData: TData[], tableState?: {
        visibleColumns: { id: string; visible: boolean }[];
        filters: ColumnFiltersState;
        sorting: SortingState;
    }) => Promise<void>
}

export function DataTable<TData, TValue>({
    columns,
    data,
    onExport,
}: DataTableProps<TData, TValue>) {
    const [sorting, setSorting] = React.useState<SortingState>([]);
    const [columnFilters, setColumnFilters] = React.useState<ColumnFiltersState>([]);
    const [columnVisibility, setColumnVisibility] = React.useState<VisibilityState>({});
    const [selectedColumn, setSelectedColumn] = React.useState<string>("customerName");
    const [filterValue, setFilterValue] = React.useState<string>("");

    // Advanced filters
    const [amountRange, setAmountRange] = React.useState<[number, number]>([0, 10000]);
    const [quantityRange, setQuantityRange] = React.useState<[number, number]>([0, 100]);
    const [selectedCategories, setSelectedCategories] = React.useState<string[]>([]);
    const [selectedStates, setSelectedStates] = React.useState<string[]>([]);
    const [selectedCountries, setSelectedCountries] = React.useState<string[]>([]);
    const [dateRange, setDateRange] = React.useState<{ from: string; to: string }>({ from: "", to: "" });

    const table = useReactTable({
        data,
        columns,
        getCoreRowModel: getCoreRowModel(),
        getPaginationRowModel: getPaginationRowModel(),
        onSortingChange: setSorting,
        getSortedRowModel: getSortedRowModel(),
        onColumnFiltersChange: setColumnFilters,
        getFilteredRowModel: getFilteredRowModel(),
        onColumnVisibilityChange: setColumnVisibility,
        initialState: {
            pagination: {
                pageSize: 10,
            },
        },
        state: {
            sorting,
            columnFilters,
            columnVisibility
        },
    });

    // Extract unique values from data
    const uniqueCategories = React.useMemo(() =>
        Array.from(new Set(data.map((item: any) => item.wineCategory).filter(Boolean))).sort(),
        [data]
    );

    const uniqueStates = React.useMemo(() =>
        Array.from(new Set(data.map((item: any) => item.customerState).filter(Boolean))).sort(),
        [data]
    );

    const uniqueCountries = React.useMemo(() =>
        Array.from(new Set(data.map((item: any) => item.wineCountry).filter(Boolean))).sort(),
        [data]
    );

    // Calculate min/max for ranges
    const amountMinMax = React.useMemo(() => {
        const amounts = data.map((item: any) => parseFloat(item.saleAmount)).filter(n => !isNaN(n));
        return { min: Math.floor(Math.min(...amounts)), max: Math.ceil(Math.max(...amounts)) };
    }, [data]);

    const quantityMinMax = React.useMemo(() => {
        const quantities = data.map((item: any) => parseInt(item.quantity)).filter(n => !isNaN(n));
        return { min: Math.floor(Math.min(...quantities)), max: Math.ceil(Math.max(...quantities)) };
    }, [data]);

    // Initialize ranges
    React.useEffect(() => {
        setAmountRange([amountMinMax.min, amountMinMax.max]);
    }, [amountMinMax]);

    React.useEffect(() => {
        setQuantityRange([quantityMinMax.min, quantityMinMax.max]);
    }, [quantityMinMax]);

    // Apply advanced filters
    React.useEffect(() => {
        const filters: ColumnFiltersState = [];

        // Amount filter
        if (amountRange[0] !== amountMinMax.min || amountRange[1] !== amountMinMax.max) {
            filters.push({
                id: "saleAmount",
                value: amountRange,
            });
        }

        // Quantity filter
        if (quantityRange[0] !== quantityMinMax.min || quantityRange[1] !== quantityMinMax.max) {
            filters.push({
                id: "quantity",
                value: quantityRange,
            });
        }

        // Category filter
        if (selectedCategories.length > 0) {
            filters.push({
                id: "wineCategory",
                value: selectedCategories,
            });
        }

        // State filter
        if (selectedStates.length > 0) {
            filters.push({
                id: "customerState",
                value: selectedStates,
            });
        }

        // Country filter
        if (selectedCountries.length > 0) {
            filters.push({
                id: "wineCountry",
                value: selectedCountries,
            });
        }

        // Date filter
        if (dateRange.from || dateRange.to) {
            filters.push({
                id: "saleDate",
                value: dateRange,
            });
        }

        setColumnFilters(prev => {
            const basicFilter = prev.find(f => f.id === selectedColumn && f.id !== "saleAmount" && f.id !== "quantity" && f.id !== "wineCategory" && f.id !== "customerState" && f.id !== "wineCountry" && f.id !== "saleDate");
            return basicFilter ? [...filters, basicFilter] : filters;
        });
    }, [amountRange, quantityRange, selectedCategories, selectedStates, selectedCountries, dateRange, amountMinMax, quantityMinMax]);

    // Handle basic filter
    const handleFilterChange = (value: string) => {
        setFilterValue(value);
        table.getColumn(selectedColumn)?.setFilterValue(value);
    };

    const handleColumnChange = (value: string) => {
        table.getColumn(selectedColumn)?.setFilterValue("");
        setSelectedColumn(value);
        setFilterValue("");
    };

    const filterableColumns = table.getAllColumns().filter(column => column.getCanFilter());

    const columnLabels: Record<string, string> = {
        saleID: "Sale ID",
        customerName: "Customer",
        wineDesignation: "Wine Name",
        quantity: "Quantity",
        saleAmount: "Amount",
        saleDate: "Date",
        wineCategory: "Wine Category",
        customerState: "Customer State",
        wineCountry: "Wine Origin Country"
    };

    const clearAllFilters = () => {
        setAmountRange([amountMinMax.min, amountMinMax.max]);
        setQuantityRange([quantityMinMax.min, quantityMinMax.max]);
        setSelectedCategories([]);
        setSelectedStates([]);
        setSelectedCountries([]);
        setDateRange({ from: "", to: "" });
        setFilterValue("");
        table.getColumn(selectedColumn)?.setFilterValue("");
    };

    const activeFiltersCount =
        (amountRange[0] !== amountMinMax.min || amountRange[1] !== amountMinMax.max ? 1 : 0) +
        (quantityRange[0] !== quantityMinMax.min || quantityRange[1] !== quantityMinMax.max ? 1 : 0) +
        (selectedCategories.length > 0 ? 1 : 0) +
        (selectedStates.length > 0 ? 1 : 0) +
        (selectedCountries.length > 0 ? 1 : 0) +
        (dateRange.from || dateRange.to ? 1 : 0) +
        (filterValue ? 1 : 0);

    return (
        <div className="w-full bg-black text-white p-6 rounded-lg">
            <div className="flex items-center gap-4 py-4 flex-wrap">
                <div className="flex items-center gap-2 flex-1 min-w-[300px]">
                    <Select value={selectedColumn} onValueChange={handleColumnChange}>
                        <SelectTrigger className="w-[200px] bg-neutral-900 border-neutral-700 text-white hover:bg-neutral-800 transition-colors">
                            <SelectValue placeholder="Select column" />
                        </SelectTrigger>
                        <SelectContent className="bg-neutral-900 border-neutral-700">
                            <SelectGroup>
                                <SelectLabel className="text-neutral-400">Search by column</SelectLabel>
                                {filterableColumns.map((column) => (
                                    <SelectItem
                                        key={column.id}
                                        value={column.id}
                                        className="text-white hover:bg-neutral-800 focus:bg-neutral-800"
                                    >
                                        {columnLabels[column.id] || column.id}
                                    </SelectItem>
                                ))}
                            </SelectGroup>
                        </SelectContent>
                    </Select>

                    <Input
                        placeholder={`Search ${columnLabels[selectedColumn] || selectedColumn}...`}
                        value={filterValue}
                        onChange={(event) => handleFilterChange(event.target.value)}
                        className="flex-1 max-w-md bg-neutral-900 border-neutral-700 text-white placeholder:text-neutral-500 focus:border-purple-500 focus:ring-1 focus:ring-purple-500 transition-all"
                    />
                </div>

                <Popover>
                    <PopoverTrigger asChild>
                        <Button
                            variant="outline"
                            className="bg-liniar-to-r from-red-600 to-purple-600 hover:from-red-700 hover:to-purple-700 border-0 text-white shadow-lg hover:shadow-xl transition-all relative"
                        >
                            <Filter className="mr-2 h-4 w-4" />
                            Advanced Filters
                            {activeFiltersCount > 0 && (
                                <span className="absolute -top-2 -right-2 bg-red-500 text-white text-xs rounded-full h-5 w-5 flex items-center justify-center">
                                    {activeFiltersCount}
                                </span>
                            )}
                        </Button>
                    </PopoverTrigger>
                    <PopoverContent className="w-[400px] bg-neutral-900 border-neutral-700 text-white" align="end">
                        <div className="space-y-6">
                            <div className="flex items-center justify-between">
                                <h4 className="font-semibold text-lg bg-liniar-to-r from-red-400 to-purple-400 bg-clip-text text-transparent">
                                    Advanced Filters
                                </h4>
                                {activeFiltersCount > 0 && (
                                    <Button
                                        variant="ghost"
                                        size="sm"
                                        onClick={clearAllFilters}
                                        className="text-xs text-neutral-400 hover:text-white"
                                    >
                                        <X className="h-3 w-3 mr-1" />
                                        Clear All
                                    </Button>
                                )}
                            </div>

                            {/* Amount Range */}
                            <div className="space-y-3">
                                <label className="text-sm font-medium text-neutral-300">Sale Amount</label>
                                <div className="px-2">
                                    <Slider
                                        min={amountMinMax.min}
                                        max={amountMinMax.max}
                                        step={10}
                                        value={amountRange}
                                        onValueChange={(value) => setAmountRange(value as [number, number])}
                                        className="[[role=slider]]:bg-purple-600 [[role=slider]]:border-purple-600"
                                    />
                                </div>
                                <div className="flex justify-between text-xs text-neutral-400">
                                    <span>${amountRange[0].toFixed(0)}</span>
                                    <span>${amountRange[1].toFixed(0)}</span>
                                </div>
                            </div>

                            {/* Quantity Range */}
                            <div className="space-y-3">
                                <label className="text-sm font-medium text-neutral-300">Quantity</label>
                                <div className="px-2">
                                    <Slider
                                        min={quantityMinMax.min}
                                        max={quantityMinMax.max}
                                        step={1}
                                        value={quantityRange}
                                        onValueChange={(value) => setQuantityRange(value as [number, number])}
                                        className="[[role=slider]]:bg-red-600 [[role=slider]]:border-red-600"
                                    />
                                </div>
                                <div className="flex justify-between text-xs text-neutral-400">
                                    <span>{quantityRange[0]}</span>
                                    <span>{quantityRange[1]}</span>
                                </div>
                            </div>

                            {/* Date Range */}
                            <div className="space-y-3">
                                <label className="text-sm font-medium text-neutral-300 flex items-center gap-2">
                                    <Calendar className="h-4 w-4" />
                                    Date Range
                                </label>
                                <div className="grid grid-cols-2 gap-2">
                                    <Input
                                        type="date"
                                        value={dateRange.from}
                                        onChange={(e) => setDateRange(prev => ({ ...prev, from: e.target.value }))}
                                        className="bg-neutral-800 border-neutral-700 text-white text-xs"
                                    />
                                    <Input
                                        type="date"
                                        value={dateRange.to}
                                        onChange={(e) => setDateRange(prev => ({ ...prev, to: e.target.value }))}
                                        className="bg-neutral-800 border-neutral-700 text-white text-xs"
                                    />
                                </div>
                            </div>

                            {/* Wine Categories */}
                            <div className="space-y-3">
                                <label className="text-sm font-medium text-neutral-300">Wine Categories</label>
                                <div className="max-h-32 overflow-y-auto space-y-2 pr-2">
                                    {uniqueCategories.map((category) => (
                                        <div key={category} className="flex items-center space-x-2">
                                            <Checkbox
                                                id={`cat-${category}`}
                                                checked={selectedCategories.includes(category)}
                                                onCheckedChange={(checked: any) => {
                                                    setSelectedCategories(prev =>
                                                        checked
                                                            ? [...prev, category]
                                                            : prev.filter(c => c !== category)
                                                    );
                                                }}
                                                className="border-neutral-600 data-[state=checked]:bg-purple-600 data-[state=checked]:border-purple-600"
                                            />
                                            <label
                                                htmlFor={`cat-${category}`}
                                                className="text-sm text-neutral-300 cursor-pointer"
                                            >
                                                {category}
                                            </label>
                                        </div>
                                    ))}
                                </div>
                            </div>

                            {/* Customer States */}
                            <div className="space-y-3">
                                <label className="text-sm font-medium text-neutral-300">Customer States</label>
                                <div className="max-h-32 overflow-y-auto space-y-2 pr-2">
                                    {uniqueStates.map((state) => (
                                        <div key={state} className="flex items-center space-x-2">
                                            <Checkbox
                                                id={`state-${state}`}
                                                checked={selectedStates.includes(state)}
                                                onCheckedChange={(checked: any) => {
                                                    setSelectedStates(prev =>
                                                        checked
                                                            ? [...prev, state]
                                                            : prev.filter(s => s !== state)
                                                    );
                                                }}
                                                className="border-neutral-600 data-[state=checked]:bg-red-600 data-[state=checked]:border-red-600"
                                            />
                                            <label
                                                htmlFor={`state-${state}`}
                                                className="text-sm text-neutral-300 cursor-pointer"
                                            >
                                                {state}
                                            </label>
                                        </div>
                                    ))}
                                </div>
                            </div>

                            {/* Wine Countries */}
                            <div className="space-y-3">
                                <label className="text-sm font-medium text-neutral-300">Wine Origin Countries</label>
                                <div className="max-h-32 overflow-y-auto space-y-2 pr-2">
                                    {uniqueCountries.map((country) => (
                                        <div key={country} className="flex items-center space-x-2">
                                            <Checkbox
                                                id={`country-${country}`}
                                                checked={selectedCountries.includes(country)}
                                                onCheckedChange={(checked: any) => {
                                                    setSelectedCountries(prev =>
                                                        checked
                                                            ? [...prev, country]
                                                            : prev.filter(c => c !== country)
                                                    );
                                                }}
                                                className="border-neutral-600 data-[state=checked]:bg-purple-600 data-[state=checked]:border-purple-600"
                                            />
                                            <label
                                                htmlFor={`country-${country}`}
                                                className="text-sm text-neutral-300 cursor-pointer"
                                            >
                                                {country}
                                            </label>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </div>
                    </PopoverContent>
                </Popover>

                <DropdownMenu>
                    <DropdownMenuTrigger asChild>
                        <Button
                            variant="outline"
                            className="bg-neutral-900 border-neutral-700 text-white hover:bg-neutral-800 transition-all"
                        >
                            Columns <ChevronDown className="ml-2 h-4 w-4" />
                        </Button>
                    </DropdownMenuTrigger>
                    <DropdownMenuContent align="end" className="bg-neutral-900 border-neutral-700 w-[200px]">
                        {table
                            .getAllColumns()
                            .filter((column) => column.getCanHide())
                            .map((column) => {
                                return (
                                    <DropdownMenuCheckboxItem
                                        key={column.id}
                                        className="text-white hover:bg-neutral-800 focus:bg-neutral-800 capitalize"
                                        checked={column.getIsVisible()}
                                        onCheckedChange={(value) =>
                                            column.toggleVisibility(!!value)
                                        }
                                    >
                                        {columnLabels[column.id] || column.id}
                                    </DropdownMenuCheckboxItem>
                                )
                            })}
                    </DropdownMenuContent>
                </DropdownMenu>

                {onExport && (
                    <ExportButton
                        onExport={async () => {
                            const filteredRows = table.getFilteredRowModel().rows;
                            const sortedRows = table.getSortedRowModel().rows;
                            const visibleColumns = table.getVisibleLeafColumns();

                            // Get the actual displayed data (filtered + sorted)
                            const displayedData = sortedRows
                                .filter(row => filteredRows.some(fr => fr.id === row.id))
                                .map(row => row.original);

                            // Get column info
                            const columnInfo = visibleColumns.map(col => ({
                                id: col.id,
                                visible: col.getIsVisible()
                            }));

                            // Get active filters
                            const activeFilters = table.getState().columnFilters;

                            // Get sorting state
                            const sortingState = table.getState().sorting;

                            await onExport(displayedData, {
                                visibleColumns: columnInfo,
                                filters: activeFilters,
                                sorting: sortingState
                            });
                        }}
                        showLabel
                    />
                )}
            </div>

            <div className="flex items-center justify-between mb-4">
                <div className="text-sm text-neutral-400">
                    Showing <span className="text-white font-medium">{table.getFilteredRowModel().rows.length}</span> result(s)
                    {table.getFilteredRowModel().rows.length !== data.length && (
                        <span> of <span className="text-white font-medium">{data.length}</span> total</span>
                    )}
                </div>
                <div className="flex items-center gap-2">
                    <span className="text-sm text-neutral-400">Rows per page:</span>
                    <Select
                        value={`${table.getState().pagination.pageSize}`}
                        onValueChange={(value) => {
                            table.setPageSize(Number(value))
                        }}
                    >
                        <SelectTrigger className="h-8 w-[70px] bg-neutral-900 border-neutral-700 text-white">
                            <SelectValue placeholder={table.getState().pagination.pageSize} />
                        </SelectTrigger>
                        <SelectContent side="top" className="bg-neutral-900 border-neutral-700">
                            {[10, 20, 30, 50, 100].map((pageSize) => (
                                <SelectItem
                                    key={pageSize}
                                    value={`${pageSize}`}
                                    className="text-white hover:bg-neutral-800 focus:bg-neutral-800"
                                >
                                    {pageSize}
                                </SelectItem>
                            ))}
                        </SelectContent>
                    </Select>
                </div>
            </div>

            <div className="overflow-hidden rounded-md border border-neutral-800 shadow-2xl">
                <Table>
                    <TableHeader>
                        {table.getHeaderGroups().map((headerGroup) => (
                            <TableRow
                                key={headerGroup.id}
                                className="border-b border-neutral-800 bg-liniar-to-r from-red-900/20 to-purple-900/20 hover:from-red-900/30 hover:to-purple-900/30 transition-all"
                            >
                                {headerGroup.headers.map((header) => {
                                    return (
                                        <TableHead key={header.id} className="text-white font-semibold">
                                            {header.isPlaceholder
                                                ? null
                                                : flexRender(
                                                    header.column.columnDef.header,
                                                    header.getContext()
                                                )}
                                        </TableHead>
                                    )
                                })}
                            </TableRow>
                        ))}
                    </TableHeader>
                    <TableBody>
                        {table.getRowModel().rows?.length ? (
                            table.getRowModel().rows.map((row) => (
                                <TableRow
                                    key={row.id}
                                    data-state={row.getIsSelected() && "selected"}
                                    className="border-b border-neutral-800 hover:bg-neutral-900 transition-colors"
                                >
                                    {row.getVisibleCells().map((cell) => (
                                        <TableCell key={cell.id} className="text-neutral-200">
                                            {flexRender(cell.column.columnDef.cell, cell.getContext())}
                                        </TableCell>
                                    ))}
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell
                                    colSpan={columns.length}
                                    className="h-24 text-center text-neutral-400"
                                >
                                    No results found.
                                </TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>

            <div className="flex items-center justify-between py-4 flex-wrap gap-4">
                <div className="flex items-center gap-2 text-sm text-neutral-400">
                    Page <span className="text-white font-medium">{table.getState().pagination.pageIndex + 1}</span> of{" "}
                    <span className="text-white font-medium">{table.getPageCount()}</span>
                </div>

                <div className="flex items-center space-x-2">
                    <Button
                        variant="outline"
                        size="sm"
                        onClick={() => table.setPageIndex(0)}
                        disabled={!table.getCanPreviousPage()}
                        className="h-8 w-8 p-0 bg-neutral-900 border-neutral-700 text-white hover:bg-neutral-800 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                        <ChevronsLeft className="h-4 w-4" />
                    </Button>
                    <Button
                        variant="outline"
                        size="sm"
                        onClick={() => table.previousPage()}
                        disabled={!table.getCanPreviousPage()}
                        className="h-8 px-3 bg-neutral-900 border-neutral-700 text-white hover:bg-neutral-800 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                        <ChevronLeft className="h-4 w-4 mr-1" />
                        Previous
                    </Button>
                    <Button
                        variant="outline"
                        size="sm"
                        onClick={() => table.nextPage()}
                        disabled={!table.getCanNextPage()}
                        className="h-8 px-3 bg-neutral-900 border-neutral-700 text-white hover:bg-neutral-800 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                        Next
                        <ChevronRight className="h-4 w-4 ml-1" />
                    </Button>
                    <Button
                        variant="outline"
                        size="sm"
                        onClick={() => table.setPageIndex(table.getPageCount() - 1)}
                        disabled={!table.getCanNextPage()}
                        className="h-8 w-8 p-0 bg-neutral-900 border-neutral-700 text-white hover:bg-neutral-800 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                        <ChevronsRight className="h-4 w-4" />
                    </Button>
                </div>
            </div>
        </div>
    )
}
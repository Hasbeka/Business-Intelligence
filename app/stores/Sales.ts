import { create } from "zustand";
import { GeneralSalesInfo, MonthlySales, MonthlyStats, Sale, SalesBetterFormat, SalesComparison, MonthlySalesF } from "../types";
import { readCsv } from "@/lib/CsvReader";
import { toSales } from "@/lib/utils";
import { useCustomersStore } from "./Customers";
import { getSalesByMonthAndCategory, calculateAvgSaleAmount, calculateTotalRevenue, compareWithPreviousMonth, getBestSellingWine, getCustomerWithHighestPurchase, getMonthlySales, getWorstSellingWine, getMonthlySalesStats } from "@/lib/GeneralSalesInfo";
import { useWineStore } from "./Wine";
import { useLocationsStore } from "./Locations";


type SalesState = {
    sales: Sale[];
    salesBetterFormat: SalesBetterFormat[]
    setSales: () => void;
    getSalesStats: () => GeneralSalesInfo;
    amountOnMonths: () => MonthlySales[];
    salesComparison: () => SalesComparison;
    getSalesBetterFormat: () => void;
    montlySalesOnCategory: () => MonthlySalesF[];
    montlyStats: () => MonthlyStats[];
}

export const useSalesStore = create<SalesState>((set, get) => ({
    sales: [],
    salesBetterFormat: [],
    setSales: async () => {
        const salesData = await readCsv('Sales_Data_Dataset.csv');
        const salesList = salesData.map(toSales);
        set({ sales: salesList });
    },
    getSalesStats: () => {
        const salesList = get().sales;
        const customerList = useCustomersStore.getState().customers;
        return {
            totalSales: salesList.length,
            totalRevenue: calculateTotalRevenue(salesList),
            avgSaleAmount: calculateAvgSaleAmount(salesList),
            bestSellingWine: getBestSellingWine(salesList),
            worstSellingWine: getWorstSellingWine(salesList),
            customerWithHighestPurchase: getCustomerWithHighestPurchase(salesList, customerList)
        }
    },
    amountOnMonths: () => {
        return getMonthlySales(get().sales);
    },
    salesComparison: () => {
        const salesList = get().sales;
        return compareWithPreviousMonth(salesList);
    },
    getSalesBetterFormat: () => {
        const salesList = get().sales;
        const customers = useCustomersStore.getState().customers;
        const wines = useWineStore.getState().wines;
        const locations = useLocationsStore.getState().locations;

        if (!customers.length || !wines.length || !locations.length) {
            return [];
        }


        const customerMap = new Map(customers.map(c => [c.customerID, c]));
        const wineMap = new Map(wines.map(w => [w.wineDesignation, w]));
        const locationMap = new Map(locations.map(l => [l.postalCode, l]));

        set({
            salesBetterFormat: salesList.map(sale => {
                const customer = customerMap.get(sale.customerID);
                const wine = wineMap.get(sale.wineDesignation);

                const postalCode = customer?.address.match(/\b\d{5}\b/)?.[0];
                const location = postalCode ? locationMap.get(postalCode) : undefined;

                return {
                    saleID: sale.saleID,
                    customerName: customer?.name || 'Unknown Customer',
                    wineDesignation: sale.wineDesignation,
                    saleDate: new Date(sale.saleDate).toLocaleDateString(),
                    saleAmount: sale.saleAmount,
                    quantity: sale.quantity,
                    wineCategory: wine?.category || 'Unknown Category',
                    customerCountry: location?.country || '-',
                    customerState: location?.state || '-',
                    wineCountry: wine?.country || '-'
                };
            })
        })
    },
    montlySalesOnCategory: () => {
        const sales = get().salesBetterFormat;
        return getSalesByMonthAndCategory(sales);
    },
    montlyStats: (): MonthlyStats[] => {
        return getMonthlySalesStats(useSalesStore.getState().sales)
    },
}))
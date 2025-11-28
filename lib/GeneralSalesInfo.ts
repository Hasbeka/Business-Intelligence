import { MonthlySales, MonthlySalesF, MonthlyStats, Sale, SalesBetterFormat, SalesComparison } from "@/app/types";


export function calculateTotalRevenue(salesList: any[]): number {
    return salesList.reduce((total, sale) => total + sale.saleAmount, 0);
}
export function calculateAvgSaleAmount(salesList: any[]): number {
    if (salesList.length === 0) return 0;
    const totalRevenue = calculateTotalRevenue(salesList);
    return totalRevenue / salesList.length;
}

export function getBestSellingWine(salesList: Sale[]): string {
    const wineSalesMap: { [key: string]: number } = {};
    salesList.forEach(sale => {
        if (!wineSalesMap[sale.wineDesignation]) {
            wineSalesMap[sale.wineDesignation] = 0;
        }
        wineSalesMap[sale.wineDesignation] += sale.saleAmount;
    });
    let bestSellingWine = '';
    let maxSales = 0;
    for (const wine in wineSalesMap) {
        if (wineSalesMap[wine] > maxSales) {
            maxSales = wineSalesMap[wine];
            bestSellingWine = wine;
        }
    }
    return bestSellingWine;
}
export function getWorstSellingWine(salesList: Sale[]): string {
    const wineSalesMap: { [key: string]: number } = {};
    salesList.forEach(sale => {
        if (!wineSalesMap[sale.wineDesignation]) {
            wineSalesMap[sale.wineDesignation] = 0;
        }
        wineSalesMap[sale.wineDesignation] += sale.saleAmount;
    });

    let worstSellingWine = '';
    let minSales = Infinity;

    for (const wine in wineSalesMap) {
        if (wineSalesMap[wine] < minSales) {
            minSales = wineSalesMap[wine];
            worstSellingWine = wine;
        }
    }

    return worstSellingWine;
}
export function getCustomerWithHighestPurchase(salesList: any[], customerList: any[]): Map<string, number> {
    const customerPurchaseMap: { [key: number]: number } = {};
    salesList.forEach(sale => {
        if (!customerPurchaseMap[sale.customerID]) {
            customerPurchaseMap[sale.customerID] = 0;
        }
        customerPurchaseMap[sale.customerID] += sale.saleAmount;
    });
    let topCustomerID = -1;
    let maxPurchase = 0;
    for (const customerID in customerPurchaseMap) {
        if (customerPurchaseMap[customerID] > maxPurchase) {
            maxPurchase = customerPurchaseMap[customerID];
            topCustomerID = parseInt(customerID);
        }
    }
    const topCustomer = customerList.find(customer => customer.customerID === topCustomerID);
    return new Map([[topCustomer ? topCustomer.name : 'Unknown', maxPurchase]]);
}


export function getSalesAmountByMonth(sales: Sale[], year: number, month: number): number {
    return sales.reduce((total, sale) => {
        const saleYear = sale.saleDate.getFullYear();
        const saleMonth = sale.saleDate.getMonth() + 1;

        if (saleYear === year && saleMonth === month) {
            return total + sale.saleAmount;
        }
        return total;
    }, 0);
}

export function getMonthlySales(sales: Sale[]): MonthlySales[] {
    const monthlyMap = new Map<string, number>();

    sales.forEach(sale => {
        const year = sale.saleDate.getFullYear();
        const month = sale.saleDate.getMonth() + 1;
        const key = `${year}-${month}`;

        monthlyMap.set(key, (monthlyMap.get(key) || 0) + sale.saleAmount);
    });

    return Array.from(monthlyMap.entries()).map(([key, totalAmount]) => {
        const [year, month] = key.split('-').map(Number);
        return { year, month, totalAmount };
    });
}


export function compareWithPreviousMonth(sales: Sale[]): SalesComparison {
    const monthlySales = getMonthlySales(sales);

    const sortedSales = monthlySales.sort((a, b) => {
        if (a.year !== b.year) return a.year - b.year;
        return a.month - b.month;
    });

    if (sortedSales.length < 2) {
        return {
            currentMonth: sortedSales[sortedSales.length - 1] || { year: 0, month: 0, totalAmount: 0 },
            previousMonth: null,
            growthPercentage: null,
            isGrowth: null,
            message: "Not enough data to compare months",
        };
    }

    const currentMonth = sortedSales[sortedSales.length - 2];
    const previousMonth = sortedSales[sortedSales.length - 3];

    const growthPercentage = ((currentMonth.totalAmount - previousMonth.totalAmount) / previousMonth.totalAmount) * 100;
    const isGrowth = growthPercentage > 0;

    let message = "";
    if (isGrowth) {
        message = `The sales have increased by ${growthPercentage.toFixed(2)}% compared to last month`;

    } else if (growthPercentage < 0) {
        message = `The sales have decreased by ${Math.abs(growthPercentage).toFixed(2)}% compared to last month`;

    } else {
        message = `The sales are unchanged compared to last month`;

    }

    return {
        currentMonth,
        previousMonth,
        growthPercentage,
        isGrowth,
        message,

    };
}

export function getSalesByMonthAndCategory(sales: SalesBetterFormat[]): MonthlySalesF[] {
    const monthlyMap = new Map<string, MontlyNumbers>();

    sales.forEach(sale => {
        const date = parseDateDDMMYYYY(sale.saleDate);
        if (!date) return; // skip invalid dates

        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const category = sale.wineCategory;

        const key = `${year}-${month}-${category}`;
        monthlyMap.set(key, {
            quantity: (monthlyMap.get(key)?.quantity || 0) + sale.quantity,
            saleAmount: (monthlyMap.get(key)?.saleAmount || 0) + sale.saleAmount,
            avgPrice: ((monthlyMap.get(key)?.saleAmount || 0) + sale.saleAmount) / (monthlyMap.get(key)?.quantity || 0) + sale.quantity
        });
    });

    const result: MonthlySalesF[] = Array.from(monthlyMap.entries()).map(([key, numbers]) => {
        const [year, month, category] = key.split("-");
        const { quantity, avgPrice, saleAmount } = numbers;
        return {
            year: Number(year),
            month: Number(month),
            category,
            quantity,
            avgPrice,
            totalAmount: saleAmount
        };
    });

    return result.sort((a, b) =>
        a.year - b.year || a.month - b.month || a.category.localeCompare(b.category)
    );
}

export function parseDateDDMMYYYY(dateStr: string): Date | null {
    const [day, month, year] = dateStr.split("/").map(Number);
    if (!day || !month || !year) return null;
    return new Date(year, month - 1, day);
}

type MontlyNumbers = {
    quantity: number;
    saleAmount: number;
    avgPrice: number;
}


export function getMonthlySalesStats(sales: Sale[]): MonthlyStats[] {
    const monthlyMap = new Map<string, MontlyNumbers>();

    sales.forEach(sale => {
        const year = sale.saleDate.getFullYear();
        const month = sale.saleDate.getMonth() + 1;
        const key = `${year}-${month}`;


        monthlyMap.set(key, {
            quantity: (monthlyMap.get(key)?.quantity || 0) + sale.quantity,
            saleAmount: (monthlyMap.get(key)?.saleAmount || 0) + sale.saleAmount,
            avgPrice: ((monthlyMap.get(key)?.saleAmount || 0) + sale.saleAmount) / (monthlyMap.get(key)?.quantity || 0) + sale.quantity
        });
    });

    return Array.from(monthlyMap.entries()).map(([key, numbers]) => {
        const [year, month] = key.split('-').map(Number);
        const { quantity, saleAmount, avgPrice } = numbers;
        return { year, month, totalQty: quantity, totalAmount: saleAmount, avgPrice };
    });
}

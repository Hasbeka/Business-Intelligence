

export type GeneralCustomerInfo = {
    total: number;
    avgAge: number;
    avgLoyaltyPoints: number;
    gender: string;
    minAge: number;
    maxAge: number;
    womenPercentage: number;
    menPercentage: number;
}

export type Customer = {
    customerID: number,
    name: string,
    email: string,
    phone: string,
    address: string,
    age: number,
    gender: string,
    purchaseHistory: string[],
    loyaltyPoints: number

}

export type Sale = {
    saleID: number,
    customerID: number,
    wineDesignation: string,
    quantity: number,
    saleAmount: number,
    saleDate: Date
}

export type GeneralSalesInfo = {
    totalSales: number;
    totalRevenue: number;
    avgSaleAmount: number;
    bestSellingWine: string;
    worstSellingWine: string;
    customerWithHighestPurchase: Map<string, number>;
}

export type MonthlySales = {
    year: number,
    month: number,
    totalAmount: number
}
export type Wine = {
    wineId: number;
    wineDesignation: string;
    category: string;
    country: string;
    region: string;
    grapeVariety: string;
    year: number;
    alcoholContent: number;
    bottleSize: string;
    priceRange: string;
}

export type MonthlySalesF = {
    year: number,
    month: number,
    totalAmount: number,
    quantity: number,
    avgPrice: number,
    category: string
}

export type SalesComparison = {
    currentMonth: MonthlySales,
    previousMonth: MonthlySales | null,
    growthPercentage: number | null,
    isGrowth: boolean | null,
    message: string
}



export type SalesBetterFormat = {
    saleID: number;
    customerName: string;
    wineDesignation: string;
    quantity: number;
    saleAmount: number;
    saleDate: string;
    wineCategory: string;
    customerCountry: string;
    wineCountry: string;
    customerState: string;
}

export type CustLocation = {
    postalCode: string;
    country: string;
    state: string;
    city: string;
}

export type MonthlyStats = {
    year: number;
    month: number;
    avgPrice: number;
    totalQty: number;
    totalAmount: number;
};

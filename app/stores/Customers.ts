import { create } from 'zustand'
import { Customer, GeneralCustomerInfo } from '@/app/types'
import { readCsv } from '@/lib/CsvReader';
import { toCustomer } from '@/lib/utils';
import { calculateAvgAge, calculateAvgLoyaltyPoints, calculateGenderPercentages, calculateMaxAge, calculateMinAge, getTheMostCommonGender } from '@/lib/GeneralCustInfo';

type CustomersState = {
    customers: Customer[];
    setCustomers: () => void;
    getCustomerStats: () => GeneralCustomerInfo;
}

export const useCustomersStore = create<CustomersState>((set, get) => ({
    customers: [],
    setCustomers: async () => {
        const customerData = await readCsv('Customers_Dataset.csv');
        const customers = customerData.map(toCustomer);
        set({ customers })
    },
    getCustomerStats: () => {
        const customerList = get().customers;
        return {
            total: customerList.length,
            avgAge: calculateAvgAge(customerList),
            avgLoyaltyPoints: calculateAvgLoyaltyPoints(customerList),
            gender: getTheMostCommonGender(customerList),
            minAge: calculateMinAge(customerList),
            maxAge: calculateMaxAge(customerList),
            womenPercentage: calculateGenderPercentages(customerList).womenPercentage,
            menPercentage: calculateGenderPercentages(customerList).menPercentage
        }
    }
}))
import { Customer } from "@/app/types";

export function calculateMaxAge(customers: Customer[]): number {
    return Math.max(...customers.map(c => {

        return c.age
    }));
}

export function calculateMinAge(customers: Customer[]): number {
    return Math.min(...customers.map(c => c.age));
}

export function calculateAvgAge(customers: Customer[]): number {
    const totalAge = customers.reduce((sum, customer) => sum + customer.age, 0);
    return parseFloat((totalAge / customers.length).toFixed(2));
}

export function calculateAvgLoyaltyPoints(customers: Customer[]): number {
    const totalPoints = customers.reduce((sum, customer) => sum + customer.loyaltyPoints, 0);
    return parseFloat((totalPoints / customers.length).toFixed(2));
}

export function getTheMostCommonGender(customers: Customer[]): string {
    const genderCount: { [key: string]: number } = {};
    customers.forEach(customer => {
        genderCount[customer.gender] = (genderCount[customer.gender] || 0) + 1;
    });
    let mostCommonGender = '';
    let maxCount = 0;
    for (const gender in genderCount) {
        if (genderCount[gender] > maxCount) {
            maxCount = genderCount[gender];
            mostCommonGender = gender;
        }
    }
    return mostCommonGender;
}

export function calculateGenderPercentages(customers: Customer[]): { menPercentage: number, womenPercentage: number } {
    const totalCustomers = customers.length;
    const menCount = customers.filter(c => {
        if (!c.gender) return false;
        return c.gender.toLowerCase() === 'male'
    }).length;
    const womenCount = customers.filter(c => {
        if (!c.gender) return false;

        return c.gender.toLowerCase() === 'female'
    }).length;

    const menPercentage = parseFloat(((menCount / totalCustomers) * 100).toFixed(2));
    const womenPercentage = parseFloat(((womenCount / totalCustomers) * 100).toFixed(2));

    return { menPercentage, womenPercentage };
}
// app/customer-segmentation/page.tsx
import { useCustomersStore } from "../stores/Customers";
import { useSalesStore } from "../stores/Sales";
import { useWineStore } from "../stores/Wine";
import CustomerSegmentAnalysis from "@/components/CustomerSegmentAnalysis";
import { Customer, Sale, Wine } from "../types";

// Helper function to analyze customer segments
function analyzeCustomerSegments(customers: Customer[], sales: Sale[], wines: Wine[]) {
    // Create wine map for quick lookup
    const wineMap = new Map(wines.map(w => [w.wineDesignation, w]));
    
    // Create customer purchase map
    const customerPurchases = new Map<number, { customer: Customer; wines: Wine[] }>();
    
    sales.forEach(sale => {
        const wine = wineMap.get(sale.wineDesignation);
        if (wine) {
            const existing = customerPurchases.get(sale.customerID);
            if (existing) {
                existing.wines.push(wine);
            } else {
                const customer = customers.find(c => c.customerID === sale.customerID);
                if (customer) {
                    customerPurchases.set(sale.customerID, {
                        customer,
                        wines: [wine]
                    });
                }
            }
        }
    });

    // Analyze by Gender
    const genderSegments = analyzeByGender(customerPurchases);
    
    // Analyze by Age Groups
    const ageSegments = analyzeByAgeGroup(customerPurchases);
    
    // Analyze by Gender and Age combined
    const combinedSegments = analyzeByGenderAndAge(customerPurchases);

    return {
        genderSegments,
        ageSegments,
        combinedSegments
    };
}

function analyzeByGender(customerPurchases: Map<number, { customer: Customer; wines: Wine[] }>) {
    const segments = new Map<string, {
        count: number;
        categories: Map<string, number>;
        varieties: Map<string, number>;
        countries: Map<string, number>;
        priceRanges: Map<string, number>;
    }>();

    customerPurchases.forEach(({ customer, wines }) => {
        const gender = customer.gender;
        
        if (!segments.has(gender)) {
            segments.set(gender, {
                count: 0,
                categories: new Map(),
                varieties: new Map(),
                countries: new Map(),
                priceRanges: new Map()
            });
        }

        const segment = segments.get(gender)!;
        segment.count++;

        wines.forEach(wine => {
            // Categories
            segment.categories.set(
                wine.category,
                (segment.categories.get(wine.category) || 0) + 1
            );
            
            // Varieties
            segment.varieties.set(
                wine.grapeVariety,
                (segment.varieties.get(wine.grapeVariety) || 0) + 1
            );
            
            // Countries
            segment.countries.set(
                wine.country,
                (segment.countries.get(wine.country) || 0) + 1
            );
            
            // Price Ranges
            segment.priceRanges.set(
                wine.priceRange,
                (segment.priceRanges.get(wine.priceRange) || 0) + 1
            );
        });
    });

    return Array.from(segments.entries()).map(([gender, data]) => ({
        gender,
        customerCount: data.count,
        topCategories: getTopItems(data.categories, 5),
        topVarieties: getTopItems(data.varieties, 5),
        topCountries: getTopItems(data.countries, 3),
        topPriceRanges: getTopItems(data.priceRanges, 3)
    }));
}

function analyzeByAgeGroup(customerPurchases: Map<number, { customer: Customer; wines: Wine[] }>) {
    const ageGroups = [
        { name: "18-25", min: 18, max: 25 },
        { name: "26-35", min: 26, max: 35 },
        { name: "36-45", min: 36, max: 45 },
        { name: "46-55", min: 46, max: 55 },
        { name: "55+", min: 55, max: 150 }
    ];

    const segments = new Map<string, {
        count: number;
        categories: Map<string, number>;
        varieties: Map<string, number>;
        avgAge: number;
        totalAge: number;
    }>();

    ageGroups.forEach(group => {
        segments.set(group.name, {
            count: 0,
            categories: new Map(),
            varieties: new Map(),
            avgAge: 0,
            totalAge: 0
        });
    });

    customerPurchases.forEach(({ customer, wines }) => {
        const ageGroup = ageGroups.find(g => customer.age >= g.min && customer.age <= g.max);
        if (ageGroup) {
            const segment = segments.get(ageGroup.name)!;
            segment.count++;
            segment.totalAge += customer.age;

            wines.forEach(wine => {
                segment.categories.set(
                    wine.category,
                    (segment.categories.get(wine.category) || 0) + 1
                );
                segment.varieties.set(
                    wine.grapeVariety,
                    (segment.varieties.get(wine.grapeVariety) || 0) + 1
                );
            });
        }
    });

    return Array.from(segments.entries()).map(([ageGroup, data]) => ({
        ageGroup,
        customerCount: data.count,
        avgAge: data.count > 0 ? Math.round(data.totalAge / data.count) : 0,
        topCategories: getTopItems(data.categories, 5),
        topVarieties: getTopItems(data.varieties, 5)
    }));
}

function analyzeByGenderAndAge(customerPurchases: Map<number, { customer: Customer; wines: Wine[] }>) {
    const ageGroups = [
        { name: "18-35", min: 18, max: 35 },
        { name: "36-55", min: 36, max: 55 },
        { name: "55+", min: 55, max: 150 }
    ];

    const segments = new Map<string, {
        count: number;
        categories: Map<string, number>;
        varieties: Map<string, number>;
    }>();

    customerPurchases.forEach(({ customer, wines }) => {
        const ageGroup = ageGroups.find(g => customer.age >= g.min && customer.age <= g.max);
        if (ageGroup) {
            const key = `${customer.gender} (${ageGroup.name})`;
            
            if (!segments.has(key)) {
                segments.set(key, {
                    count: 0,
                    categories: new Map(),
                    varieties: new Map()
                });
            }

            const segment = segments.get(key)!;
            segment.count++;

            wines.forEach(wine => {
                segment.categories.set(
                    wine.category,
                    (segment.categories.get(wine.category) || 0) + 1
                );
                segment.varieties.set(
                    wine.grapeVariety,
                    (segment.varieties.get(wine.grapeVariety) || 0) + 1
                );
            });
        }
    });

    return Array.from(segments.entries()).map(([segment, data]) => ({
        segment,
        customerCount: data.count,
        topCategories: getTopItems(data.categories, 3),
        topVarieties: getTopItems(data.varieties, 3)
    }));
}

function getTopItems(map: Map<string, number>, limit: number) {
    return Array.from(map.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, limit)
        .map(([name, count]) => ({ name, count }));
}

const page = async () => {
    // Load all data
    await useCustomersStore.getState().setCustomers();
    const customers = useCustomersStore.getState().customers;

    await useSalesStore.getState().setSales();
    const sales = useSalesStore.getState().sales;

    await useWineStore.getState().setWines();
    const wines = useWineStore.getState().wines;

    // Analyze segments
    const segmentData = analyzeCustomerSegments(customers, sales, wines);

    return (
        <div className="min-h-screen w-full bg-zinc-50 font-sans dark:bg-black">
            <CustomerSegmentAnalysis 
                genderSegments={segmentData.genderSegments}
                ageSegments={segmentData.ageSegments}
                combinedSegments={segmentData.combinedSegments}
            />
        </div>
    );
};

export default page;
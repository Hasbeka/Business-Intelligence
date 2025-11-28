// app/product-performance/page.tsx
import { useCustomersStore } from "../stores/Customers";
import { useSalesStore } from "../stores/Sales";
import { useWineStore } from "../stores/Wine";
import ProductPerformance from "@/components/ProductPerformance";
import { Customer, Sale, Wine } from "../types";

// Helper function to analyze wine associations
function analyzeWineAssociations(customers: Customer[], sales: Sale[], wines: Wine[]) {
    const wineMap = new Map(wines.map(w => [w.wineDesignation, w]));

    // Group sales by customer
    const customerPurchases = new Map<number, Set<string>>();
    3
    sales.forEach(sale => {
        if (!customerPurchases.has(sale.customerID)) {
            customerPurchases.set(sale.customerID, new Set());
        }
        customerPurchases.get(sale.customerID)!.add(sale.wineDesignation);
    });

    // Calculate wine associations
    const associations = new Map<string, Map<string, number>>();

    // For each customer's purchase history
    customerPurchases.forEach((wineSet, customerID) => {
        const wineArray = Array.from(wineSet);

        // For each pair of wines purchased by the same customer
        for (let i = 0; i < wineArray.length; i++) {
            const wine1 = wineArray[i];

            if (!associations.has(wine1)) {
                associations.set(wine1, new Map());
            }

            for (let j = 0; j < wineArray.length; j++) {
                if (i !== j) {
                    const wine2 = wineArray[j];
                    const currentCount = associations.get(wine1)!.get(wine2) || 0;
                    associations.get(wine1)!.set(wine2, currentCount + 1);
                }
            }
        }
    });

    // Calculate support (how many customers bought each wine)
    const wineSupport = new Map<string, number>();
    customerPurchases.forEach((wineSet) => {
        wineSet.forEach(wine => {
            wineSupport.set(wine, (wineSupport.get(wine) || 0) + 1);
        });
    });

    // Format associations with confidence scores
    const formattedAssociations: Array<{
        wine: string;
        wineCategory: string;
        wineCountry: string;
        totalCustomers: number;
        associations: Array<{
            associatedWine: string;
            associatedCategory: string;
            associatedCountry: string;
            count: number;
            confidence: number;
        }>;
    }> = [];

    associations.forEach((assocMap, wine) => {
        const wineInfo = wineMap.get(wine);
        const wineCustomers = wineSupport.get(wine) || 0;

        if (wineCustomers >= 3) { // Only include wines bought by at least 3 customers
            const sortedAssociations = Array.from(assocMap.entries())
                .map(([assocWine, count]) => {
                    const assocInfo = wineMap.get(assocWine);
                    return {
                        associatedWine: assocWine,
                        associatedCategory: assocInfo?.category || 'Unknown',
                        associatedCountry: assocInfo?.country || 'Unknown',
                        count,
                        confidence: (count / wineCustomers) * 100
                    };
                })
                .sort((a, b) => b.confidence - a.confidence)
                .slice(0, 5); // Top 5 associations

            if (sortedAssociations.length > 0) {
                formattedAssociations.push({
                    wine,
                    wineCategory: wineInfo?.category || 'Unknown',
                    wineCountry: wineInfo?.country || 'Unknown',
                    totalCustomers: wineCustomers,
                    associations: sortedAssociations
                });
            }
        }
    });

    // Sort by total customers (most popular wines first)
    formattedAssociations.sort((a, b) => b.totalCustomers - a.totalCustomers);

    // Calculate category associations
    const categoryAssociations = analyzeCategoryAssociations(customerPurchases, wineMap);

    // Calculate wine performance metrics
    const winePerformance = calculateWinePerformance(sales, wines);

    return {
        wineAssociations: formattedAssociations.slice(0, 50), // Top 50 wines
        categoryAssociations,
        winePerformance: winePerformance.slice(0, 20) // Top 20 performing wines
    };
}

function analyzeCategoryAssociations(
    customerPurchases: Map<number, Set<string>>,
    wineMap: Map<string, Wine>
) {
    const categoryPairs = new Map<string, Map<string, number>>();

    customerPurchases.forEach((wineSet) => {
        const categories = new Set(
            Array.from(wineSet)
                .map(w => wineMap.get(w)?.category)
                .filter(c => c !== undefined)
        );

        const catArray = Array.from(categories) as string[];

        for (let i = 0; i < catArray.length; i++) {
            const cat1 = catArray[i];
            if (!categoryPairs.has(cat1)) {
                categoryPairs.set(cat1, new Map());
            }

            for (let j = 0; j < catArray.length; j++) {
                if (i !== j) {
                    const cat2 = catArray[j];
                    const current = categoryPairs.get(cat1)!.get(cat2) || 0;
                    categoryPairs.get(cat1)!.set(cat2, current + 1);
                }
            }
        }
    });

    const result: Array<{
        category: string;
        associations: Array<{ category: string; count: number }>;
    }> = [];

    categoryPairs.forEach((assocMap, category) => {
        const sorted = Array.from(assocMap.entries())
            .map(([cat, count]) => ({ category: cat, count }))
            .sort((a, b) => b.count - a.count)
            .slice(0, 3);

        result.push({ category, associations: sorted });
    });

    return result.sort((a, b) =>
        b.associations.reduce((sum, a) => sum + a.count, 0) -
        a.associations.reduce((sum, a) => sum + a.count, 0)
    );
}

function calculateWinePerformance(sales: Sale[], wines: Wine[]) {
    const wineStats = new Map<string, {
        totalSales: number;
        totalRevenue: number;
        totalQuantity: number;
        uniqueCustomers: Set<number>;
    }>();

    sales.forEach(sale => {
        if (!wineStats.has(sale.wineDesignation)) {
            wineStats.set(sale.wineDesignation, {
                totalSales: 0,
                totalRevenue: 0,
                totalQuantity: 0,
                uniqueCustomers: new Set()
            });
        }

        const stats = wineStats.get(sale.wineDesignation)!;
        stats.totalSales += 1;
        stats.totalRevenue += sale.saleAmount;
        stats.totalQuantity += sale.quantity;
        stats.uniqueCustomers.add(sale.customerID);
    });

    const wineMap = new Map(wines.map(w => [w.wineDesignation, w]));

    return Array.from(wineStats.entries())
        .map(([wine, stats]) => {
            const wineInfo = wineMap.get(wine);
            return {
                wine,
                category: wineInfo?.category || 'Unknown',
                country: wineInfo?.country || 'Unknown',
                totalSales: stats.totalSales,
                totalRevenue: stats.totalRevenue,
                totalQuantity: stats.totalQuantity,
                uniqueCustomers: stats.uniqueCustomers.size,
                avgSaleAmount: stats.totalRevenue / stats.totalSales
            };
        })
        .sort((a, b) => b.totalRevenue - a.totalRevenue);
}

const page = async () => {
    // Load all data
    await useCustomersStore.getState().setCustomers();
    const customers = useCustomersStore.getState().customers;

    await useSalesStore.getState().setSales();
    const sales = useSalesStore.getState().sales;

    await useWineStore.getState().setWines();
    const wines = useWineStore.getState().wines;

    // Analyze wine associations
    const analysisData = analyzeWineAssociations(customers, sales, wines);

    return (
        <div className="min-h-screen w-full bg-zinc-50 font-sans dark:bg-black">
            <ProductPerformance
                wineAssociations={analysisData.wineAssociations}
                categoryAssociations={analysisData.categoryAssociations}
                winePerformance={analysisData.winePerformance}
            />
        </div>
    );
};

export default page;
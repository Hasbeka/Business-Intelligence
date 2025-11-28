// app/wine-origin-analysis/page.tsx
import { useSalesStore } from "../stores/Sales";
import { useWineStore } from "../stores/Wine";
import WineOriginAnalysis from "@/components/WineOriginAnalysis";
import { Sale, Wine } from "../types";

// Helper function to analyze wine origin sales
function analyzeWineOriginSales(sales: Sale[], wines: Wine[]) {
    const wineMap = new Map(wines.map(w => [w.wineDesignation, w]));
    
    // Sales by country
    const countryData = new Map<string, {
        totalSales: number;
        totalRevenue: number;
        totalQuantity: number;
        wines: Set<string>;
        categories: Map<string, number>;
        regions: Map<string, number>;
    }>();

    sales.forEach(sale => {
        const wine = wineMap.get(sale.wineDesignation);
        if (wine) {
            if (!countryData.has(wine.country)) {
                countryData.set(wine.country, {
                    totalSales: 0,
                    totalRevenue: 0,
                    totalQuantity: 0,
                    wines: new Set(),
                    categories: new Map(),
                    regions: new Map()
                });
            }

            const data = countryData.get(wine.country)!;
            data.totalSales += 1;
            data.totalRevenue += sale.saleAmount;
            data.totalQuantity += sale.quantity;
            data.wines.add(wine.wineDesignation);
            
            // Track categories
            data.categories.set(
                wine.category,
                (data.categories.get(wine.category) || 0) + 1
            );
            
            // Track regions
            data.regions.set(
                wine.region,
                (data.regions.get(wine.region) || 0) + 1
            );
        }
    });

    // Convert to array and sort
    const countrySummary = Array.from(countryData.entries())
        .map(([country, data]) => ({
            country,
            totalSales: data.totalSales,
            totalRevenue: data.totalRevenue,
            totalQuantity: data.totalQuantity,
            uniqueWines: data.wines.size,
            avgRevenuePerSale: data.totalRevenue / data.totalSales,
            topCategories: Array.from(data.categories.entries())
                .sort((a, b) => b[1] - a[1])
                .slice(0, 3)
                .map(([name, count]) => ({ name, count })),
            topRegions: Array.from(data.regions.entries())
                .sort((a, b) => b[1] - a[1])
                .slice(0, 3)
                .map(([name, count]) => ({ name, count }))
        }))
        .sort((a, b) => b.totalRevenue - a.totalRevenue);

    // Regional analysis (group by region)
    const regionData = new Map<string, {
        country: string;
        totalSales: number;
        totalRevenue: number;
        totalQuantity: number;
    }>();

    sales.forEach(sale => {
        const wine = wineMap.get(sale.wineDesignation);
        if (wine && wine.region) {
            const key = `${wine.region}, ${wine.country}`;
            if (!regionData.has(key)) {
                regionData.set(key, {
                    country: wine.country,
                    totalSales: 0,
                    totalRevenue: 0,
                    totalQuantity: 0
                });
            }

            const data = regionData.get(key)!;
            data.totalSales += 1;
            data.totalRevenue += sale.saleAmount;
            data.totalQuantity += sale.quantity;
        }
    });

    const regionSummary = Array.from(regionData.entries())
        .map(([region, data]) => ({
            region,
            country: data.country,
            totalSales: data.totalSales,
            totalRevenue: data.totalRevenue,
            totalQuantity: data.totalQuantity,
            avgRevenuePerSale: data.totalRevenue / data.totalSales
        }))
        .sort((a, b) => b.totalRevenue - a.totalRevenue)
        .slice(0, 15);

    return {
        countrySummary,
        regionSummary
    };
}

const page = async () => {
    // Load sales and wine data
    await useSalesStore.getState().setSales();
    const sales = useSalesStore.getState().sales;

    await useWineStore.getState().setWines();
    const wines = useWineStore.getState().wines;

    // Analyze wine origin sales
    const originAnalysis = analyzeWineOriginSales(sales, wines);

    return (
        <div className="min-h-screen w-full bg-zinc-50 font-sans dark:bg-black">
            <WineOriginAnalysis 
                countrySummary={originAnalysis.countrySummary}
                regionSummary={originAnalysis.regionSummary}
            />
        </div>
    );
};

export default page;
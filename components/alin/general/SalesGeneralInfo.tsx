import { ChartNoAxesCombined, BadgeDollarSign, Coins, Wine, WineOff, User, TrendingUp } from "lucide-react"
import { Card, CardContent, CardTitle } from "@/components/ui/card"
import { GeneralSalesInfo } from "@/app/types"

const CustomersGeneralInfo = (data: GeneralSalesInfo) => {
    const stats = [
        {
            label: 'Total Sales',
            value: data.totalSales.toLocaleString(),
            icon: ChartNoAxesCombined,
            gradient: 'from-red-900 to-purple-900'
        },
        {
            label: 'Avg Sales Amount',
            value: data.avgSaleAmount.toLocaleString(),
            icon: BadgeDollarSign,
            gradient: 'from-purple-900 to-red-900'
        },
        {
            label: 'Average Sales Amount',
            value: Math.round(data.avgSaleAmount).toString(),
            icon: Coins,
            gradient: 'from-red-900 to-purple-900'
        },
        {
            label: 'Best Selling Wine',
            value: data.bestSellingWine,
            icon: Wine,
            gradient: 'from-purple-900 to-red-900'
        },
        {
            label: 'Worst Selling Wine',
            value: data.worstSellingWine,
            icon: WineOff,
            gradient: 'from-red-900 to-purple-900'
        },
        {
            label: 'Customer With Highest Purchase',
            value: Array.from(data.customerWithHighestPurchase.entries()).map(([name, amount]) => `${name} ($${amount.toFixed(2)})`).join(', '),
            icon: User,
            gradient: 'from-purple-900 to-red-900'
        },
    ]

    return (
        <div className="w-full">
            <Card className="bg-black border-2 border-white overflow-hidden">
                <div className="bg-liniar-to-r from-red-900 to-purple-900 px-6 py-4">
                    <CardTitle className="text-white text-2xl font-bold flex items-center gap-3">
                        <TrendingUp className="w-7 h-7" />
                        General Sales Data
                    </CardTitle>
                </div>

                <CardContent className="p-4">
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-2 gap-4">
                        {stats.map((stat, index) => {
                            const Icon = stat.icon
                            return (
                                <div
                                    key={index}
                                    className="group relative bg-black border border-white rounded-lg p-4 hover:border-2 transition-all duration-300 hover:shadow-lg hover:shadow-white/10"
                                >
                                    <div className="absolute inset-0 bg-liniar-to-br opacity-0 group-hover:opacity-5 transition-opacity duration-300 rounded-lg"
                                        style={{ backgroundImage: `linear-gradient(to bottom right, rgb(127, 29, 29), rgb(88, 28, 135))` }}></div>

                                    <div className="relative z-10 flex items-start gap-3">
                                        <div className={`w-10 h-10 rounded-lg bg-liniar-to-br ${stat.gradient} flex items-center justify-center shrink-0 group-hover:scale-110 transition-transform duration-300`}>
                                            <Icon className="w-5 h-5 text-white" />
                                        </div>

                                        <div className="flex-1">
                                            <p className="text-neutral-400 text-sm font-medium mb-1">{stat.label}</p>
                                            <p className="text-white text-l font-bold">{stat.value}</p>
                                        </div>
                                    </div>
                                </div>
                            )
                        })}
                    </div>
                </CardContent>
            </Card>
        </div>
    )
}

export default CustomersGeneralInfo
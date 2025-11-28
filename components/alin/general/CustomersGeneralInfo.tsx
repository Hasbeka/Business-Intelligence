import { Users, Trophy, Calendar, TrendingUp, UserCheck } from "lucide-react"
import { Card, CardContent, CardTitle } from "@/components/ui/card"

interface GeneralCustomerInfo {
    total: number
    avgLoyaltyPoints: number
    avgAge: number
    gender: string
    minAge: number
    maxAge: number
    menPercentage: number
    womenPercentage: number
}

const CustomersGeneralInfo = (data: GeneralCustomerInfo) => {
    const stats = [
        {
            label: 'Total Customers',
            value: data.total.toLocaleString(),
            icon: Users,
            gradient: 'from-red-900 to-purple-900'
        },
        {
            label: 'Avg Loyalty Points',
            value: Math.round(data.avgLoyaltyPoints).toLocaleString(),
            icon: Trophy,
            gradient: 'from-purple-900 to-red-900'
        },
        {
            label: 'Average Age',
            value: Math.round(data.avgAge),
            icon: Calendar,
            gradient: 'from-red-900 to-purple-900'
        },
        {
            label: 'Most Common Gender',
            value: data.gender,
            icon: UserCheck,
            gradient: 'from-purple-900 to-red-900'
        },
        {
            label: 'Age Range',
            value: `${data.minAge} - ${data.maxAge}`,
            icon: TrendingUp,
            gradient: 'from-red-900 to-purple-900'
        },
        {
            label: 'Gender Distribution',
            value: `${data.menPercentage}% M / ${data.womenPercentage}% W`,
            icon: Users,
            gradient: 'from-purple-900 to-red-900'
        },
    ]

    return (
        <div className="w-full">
            <Card className="bg-black border-2 border-white overflow-hidden">
                <div className="bg-liniar-to-r from-red-900 to-purple-900 px-2 py-2">
                    <CardTitle className="text-white text-2xl font-bold flex items-center gap-3">
                        <Users className="w-7 h-7" />
                        General Customer Data
                    </CardTitle>
                </div>

                <CardContent className="px-4 py-8">
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
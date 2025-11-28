
import { GeneralCustomerInfo, GeneralSalesInfo, MonthlySales, SalesComparison } from "@/app/types"
import CustomersGeneralInfo from "./CustomersGeneralInfo"
import SalesGeneralInfo from "./SalesGeneralInfo"
import { Separator } from "../../ui/separator"
import SalesChart from "../sales/SalesChart"

interface MainStatisticsProps {
    customerData: GeneralCustomerInfo
    salesData: GeneralSalesInfo
    amountOnMonths: MonthlySales[]
    salesComparison: SalesComparison
}

const MainStatistics = (props: MainStatisticsProps) => {
    return (
        <div className="w-full px-6 py-8 flex flex-col items-center">
            <div className="w-full max-w-7xl space-y-8">
                {/* Statistics Cards Grid */}
                <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                    <div className="w-full">
                        <CustomersGeneralInfo {...props.customerData} />
                    </div>
                    <div className="w-full">
                        <SalesGeneralInfo {...props.salesData} />
                    </div>
                </div>

                {/* Elegant Separator */}
                <div className="relative py-4">
                    <Separator className="bg-white/20" />
                    <div className="absolute left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 w-2 h-2 bg-liniar-to-r from-red-900 to-purple-900 rounded-full"></div>
                </div>

                {/* Sales Chart */}

                <SalesChart chartData={props.amountOnMonths} salesComparison={props.salesComparison} />
            </div>
        </div>
    )
}

export default MainStatistics
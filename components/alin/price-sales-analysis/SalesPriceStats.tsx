import { MonthlySalesF, MonthlyStats } from '@/app/types'
import GeneralSalesPriceStats from '../general/GeneralSalesPriceStats'
import SalesPriceStatsOnCat from './SalesPriceStatsOnCat'

type Props = {
    salesStats: MonthlyStats[]
    salesCatStats: MonthlySalesF[]
}

const SalesPriceStats = (props: Props) => {
    return (
        <div className='mt-10 min-h-screen w-full'>
            <h2 className='text-2xl mx-4 font-semibold mb-4'>Sales-Price analysis</h2>
            {/* General analysis */}

            <div>
                <h3 className='mx-6 text-xl font-semibold'>General analysis for every wine</h3>
                <GeneralSalesPriceStats chartData={props.salesStats} />
            </div>

            {/* Analysis for every category */}
            <div>
                <h3 className='mx-6 text-xl font-semibold mt-3'> Analysis for every wine category</h3>
                <SalesPriceStatsOnCat chartData={props.salesCatStats} />
            </div>
        </div>
    )
}

export default SalesPriceStats
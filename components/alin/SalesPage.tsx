import React from 'react'
import GridDataSales from './GridDataSales'
import { SalesBetterFormat } from '@/app/types'

interface SalesProps {
    sales: SalesBetterFormat[]
}

const SalesPage = (props: SalesProps) => {
    return (
        <div className="w-full max-w-7xl space-y-8">
            {/* Grid and all functionalities */}
            <GridDataSales sales={props.sales} />
        </div>
    )
}

export default SalesPage 
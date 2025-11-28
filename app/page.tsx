import { GeneralCustomerInfo } from "./types";
import MainStatistics from "@/components/alin/general/MainStatistics";
import { useCustomersStore } from "./stores/Customers";
import { useSalesStore } from "./stores/Sales";


export default async function Home() {

  await useCustomersStore.getState().setCustomers();
  const customerList = useCustomersStore.getState().customers;
  const generalCustomerInfo: GeneralCustomerInfo = useCustomersStore.getState().getCustomerStats();


  // Sales data
  await useSalesStore.getState().setSales();
  const salesList = useSalesStore.getState().sales;
  const generalSalesInfo = useSalesStore.getState().getSalesStats();
  const amountOnMonths = useSalesStore.getState().amountOnMonths();
  const salesComparison = useSalesStore.getState().salesComparison();

  return (
    <div className="min-h-screen w-full  bg-zinc-50 font-sans dark:bg-black" >
      <MainStatistics customerData={generalCustomerInfo} salesData={generalSalesInfo} amountOnMonths={amountOnMonths} salesComparison={salesComparison} />
    </div >
  );
}

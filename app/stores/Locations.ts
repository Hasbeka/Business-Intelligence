import { create } from "zustand";
import { CustLocation } from "../types"
import { readCsv } from "@/lib/CsvReader";
import { toLocation } from "@/lib/utils";

type LocationsStore = {
    locations: CustLocation[];
    setLocations: () => void;
}

export const useLocationsStore = create<LocationsStore>((set, get) => ({
    locations: [],
    setLocations: async () => {
        const locationsData = await readCsv('Customers_Location.csv');
        const locations = locationsData.map(toLocation)
        set({ locations: locations });
    }

}))
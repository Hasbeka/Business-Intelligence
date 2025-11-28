import { create } from "zustand";
import { Wine } from "../types";
import { readCsv } from "@/lib/CsvReader";
import { toWine } from "@/lib/utils";


type WineState = {
    wines: Wine[];
    setWines: () => void;
}

export const useWineStore = create<WineState>((set, get) => ({
    wines: [],
    setWines: async () => {
        const wineData = await readCsv('Wines.csv');
        const wineList = wineData.map(toWine);
        set({ wines: wineList });
    }
}))
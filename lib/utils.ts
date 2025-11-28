import { CustLocation, Customer, Sale, Wine } from "@/app/types";
import { clsx, type ClassValue } from "clsx"
import { twMerge } from "tailwind-merge"

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}

export function fromStringToArray(str: string): string[] {
  if (!str) return [];
  return str.split(',').map(s => s.trim());
}

export function toCustomer(obj: any): Customer {
  return {
    customerID: obj.CustomerID,
    name: obj.Name,
    email: obj.Email,
    phone: obj.Phone,
    address: obj.Address,
    age: obj.Age,
    gender: obj.Gender,
    purchaseHistory: fromStringToArray(obj.PurchaseHistory),
    loyaltyPoints: obj.LoyaltyPoints
  };
}

export function toSales(obj: any): Sale {
  return {
    saleID: obj.SaleID,
    customerID: obj.CustomerID,
    wineDesignation: obj.WineDesignation,
    quantity: obj.Quantity,
    saleAmount: obj.SaleAmount,
    saleDate: new Date(obj.SaleDate)
  }
}

export function toWine(obj: any): Wine {
  return {
    wineId: obj.WineID,
    wineDesignation: obj.WineDesignation,
    country: obj.CountryOfOrigin,
    region: obj.Region,
    year: obj.Year,
    grapeVariety: obj.GrapeVariety,
    category: obj.Category,
    priceRange: obj.PriceRange,
    alcoholContent: obj.AlcoholContent,
    bottleSize: obj.BottleSize
  }
}

export function toLocation(obj: any): CustLocation {
  return {
    postalCode: obj.ZipCode.toString(),
    country: obj.Country,
    state: obj.Region,
    city: obj.City
  }
}

export function collectionWithoutDuplicates<T>(arr: T[], keyFn: (item: T) => any): T[] {
  const seen = new Set();
  return arr.filter(item => {
    const key = keyFn(item);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

export function parseDate(dateStr: string): Date | null {
  const [day, month, year] = dateStr.split("/").map(Number);
  if (!day || !month || !year) return null;
  return new Date(year, month - 1, day);
}

export function formatNumberEnglishStyle(number: number) {

  if (isNaN(number) || number === null || number === undefined) {
    return "Invalid number";
  }

  return Number(number).toLocaleString('en-US');
}
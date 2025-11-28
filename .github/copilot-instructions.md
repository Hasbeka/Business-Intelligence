# AI Agent Instructions for Wine Sales Analysis Project

## Project Overview
This is a Next.js-based business intelligence application for wine sales analysis, focusing on customer segmentation, sales patterns, and promotional campaign optimization.

## Architecture & Data Flow

### Core Components
- **Data Storage**: Uses Zustand for state management (`app/stores/`)
  - `Wine.ts`, `Sales.ts`, `Customers.ts`, `Locations.ts` handle respective domain data
  - Data is loaded from CSV files in `data/` directory using `lib/CsvReader.ts`

### Key Pages & Features
- **Sales Analysis** (`app/sales-analysis/page.tsx`):
  - Combines data from all stores for comprehensive analysis
  - Renders multiple visualization components:
    - `SalesPage`: General sales overview
    - `SalesChartOnCategory`: Category-based analysis
    - `SalesPriceStats`: Price statistics
    - `MarketingAnalytics`: Campaign optimization insights

### Data Model
Core types defined in `app/types.ts`:
- `Customer`: Customer profile with demographics and loyalty data
- `Sale`: Transaction details linking customers and wines
- `Wine`: Product information including category, origin, and characteristics
- `SalesBetterFormat`: Enhanced sale record with denormalized data for analysis

## Development Patterns

### State Management
- Use Zustand stores for global state management
- Each domain has its store (`useWineStore`, `useSalesStore`, etc.)
- Data is loaded asynchronously using `set` actions in stores

### Data Loading Pattern
```typescript
// Example from stores:
setData: async () => {
    const data = await readCsv('filename.csv');
    const processedData = data.map(toEntityType);
    set({ entity: processedData });
}
```

### Data Analysis Components
- Components under `components/` follow a pattern of:
  1. Receiving processed data as props
  2. Providing specific analytics visualization
  3. Using shadcn/ui components for consistent styling

## Common Tasks

### Adding New Analytics
1. Define new types in `app/types.ts` if needed
2. Add data processing methods in relevant store
3. Create visualization component in `components/`
4. Integrate into relevant page

### Data Updates
- CSV files in `data/` directory are the source of truth
- Update through stores to maintain consistency
- Use provided type converters in `lib/utils.ts`

## Integration Points
- shadcn/ui for component library
- CSV file-based data storage
- Zustand for state management
- Next.js 13+ app router for routing and API

## Project-Specific Conventions
- All dates should be handled as Date objects
- Sales amounts should be processed as numbers
- Categories and status fields use predefined string literals
- Always use type-safe data transformations through utility functions

Need any clarification or have sections you'd like me to expand upon?
import { PbiTableRow } from "./visual";
import { formatCell, BLANK_TOKEN, isBlank, displayFilterValue } from "./utils/format";

export type FiltersState = Map<number, Set<string>>;
export type FilterSearchState = Map<number, string>;

export class FilterManager {

    private filters: FiltersState = new Map();
    private filterSearch: FilterSearchState = new Map();

    constructor () {}

    public getFilters() {
        return this.filters;
    }

    public getSearch() {
        return this.filterSearch;
    }

    public clearAll() {
        this.filters.clear();
        this.filterSearch.clear();
    }

    public clearColumn(colIndex: number) {
        this.filters.delete(colIndex);
        this.filterSearch.delete(colIndex);
    }

    public getFilteredRows(rows: PbiTableRow[]): PbiTableRow[] {

        if (this.filters.size === 0) return rows;

        return rows.filter(row => {
            for (const [colIndex, selectedValues] of this.filters.entries()) {
                if (selectedValues.size === 0) return false;
                const raw = row[colIndex];
                if (isBlank(raw)) {
                    if (!selectedValues.has(BLANK_TOKEN)) return false;
                } else {
                    const cellVal = formatCell(raw);
                    if (!selectedValues.has(cellVal)) return false;
                }
            }
            return true;
        });
    }

    public getCacheKey(): string {
        const parts: string[] = [];

        Array.from(this.filters.entries())
            .sort((a, b) => a[0] - b[0])
            .forEach(([colIndex, values]) => {
                const sortedVals = Array.from(values).sort();
                parts.push(`${colIndex}:${sortedVals.join("~~")}`);
            });

        return parts.join("|");
    }

    public getButtonLabel(selected: Set<string>, total: number): string {
        const values = Array.from(selected).map(displayFilterValue);

        if (selected.size === 0) return "(All)";
        if (selected.size === 1) return `= ${values[0]}`;
        if (selected.size === total) return "(All)";
        return `${selected.size} of ${total} selected`;
    }
}
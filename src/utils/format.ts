export function formatCell(v: powerbi.PrimitiveValue): string {
    if (v === null || v === undefined) return "";
    if (v instanceof Date) return v.toISOString();
    return String(v);
}

export const BLANK_TOKEN = "__PBI_BLANK__";

export function isBlank(v: powerbi.PrimitiveValue): boolean {
    if (v === null || v === undefined) return true;
    return String(v).trim() === "";
}

const BLANK_VALUES = new Set(["BLANK_TOKEN", "__PBI_BLANK__"]);

export function displayFilterValue(v: string): string {
    return BLANK_VALUES.has(v) ? "(Blanks)" : v;
}
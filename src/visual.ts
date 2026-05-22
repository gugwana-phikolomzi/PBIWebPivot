"use strict";

import "./styles/index.css";
import { formatCell, BLANK_TOKEN, isBlank, displayFilterValue } from "./utils/format";
import { FilterManager } from "./FilterManager";

import powerbi from "powerbi-visuals-api";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;


type PbiTableColumn = powerbi.DataViewMetadataColumn;
export type PbiTableRow = powerbi.PrimitiveValue[];

type StatKey =
    | "count" | "min" | "max" | "mean" | "median" | "sum" | "stddev" | "unique"
    | "skewness" | "kurtosis" | "excess_kurtosis" | "jarque_bera"
    | "p1" | "p5" | "p10" | "p25" | "p75" | "p90" | "p95" | "p99"
    | "mode"
    | "null_count" | "null_pct" | "zero_count" | "negative_count";

interface ColumnProfile {
    type: "text" | "numeric";
    uniqueValues: string[];
    numericValues: number[];
    totalRowCount: number;
    stats?: {
        count: number;
        min?: number;
        max?: number;
        mean?: number;
        median?: number;
        sum?: number;
        stddev?: number;
        unique?: number;
        skewness?: number;
        kurtosis?: number;
        excess_kurtosis?: number;
        jarque_bera?: number;
        p1?: number;
        p5?: number;
        p10?: number;
        p25?: number;
        p75?: number;
        p90?: number;
        p95?: number;
        p99?: number;
        mode?: number;
        null_count?: number;
        null_pct?: number;
        zero_count?: number;
        negative_count?: number;
    };
}

const NUMERIC_STATS: { key: StatKey; label: string }[] = [
    { key: "count",           label: "Count" },
    { key: "unique",          label: "Unique" },
    { key: "min",             label: "Min" },
    { key: "max",             label: "Max" },
    { key: "mean",            label: "Mean" },
    { key: "median",          label: "Median" },
    { key: "sum",             label: "Sum" },
    { key: "stddev",          label: "Std Dev" },
    { key: "skewness",        label: "Skewness" },
    { key: "kurtosis",        label: "Kurtosis" },
    { key: "excess_kurtosis", label: "Excess Kurt." },
    { key: "jarque_bera",     label: "Jarque-Bera" },
    { key: "p1",              label: "P1" },
    { key: "p5",              label: "P5" },
    { key: "p10",             label: "P10" },
    { key: "p25",             label: "Q1 (P25)" },
    { key: "p75",             label: "Q3 (P75)" },
    { key: "p90",             label: "P90" },
    { key: "p95",             label: "P95" },
    { key: "p99",             label: "P99" },
    { key: "mode",            label: "Mode" },
    { key: "null_count",      label: "Null Count" },
    { key: "null_pct",        label: "Null %" },
    { key: "zero_count",      label: "Zero Count" },
    { key: "negative_count",  label: "Negative Count" },
];

const STAT_TOOLTIP: Record<StatKey, string> = {
    count: "Number of non-empty values",
    unique: "Number of distinct values",
    min: "Minimum value",
    max: "Maximum value",
    mean: "Average (arithmetic mean)",
    median: "Middle value when sorted",
    sum: "Sum of all values",
    stddev: "Population standard deviation",
    skewness: "Measure of asymmetry of the distribution",
    kurtosis: "Measure of tail heaviness (Fisher's definition)",
    excess_kurtosis: "Kurtosis minus 3 (excess over normal distribution)",
    jarque_bera: "Normality test statistic (larger = less normal)",
    p1: "1st percentile",
    p5: "5th percentile",
    p10: "10th percentile",
    p25: "25th percentile (Q1)",
    p75: "75th percentile (Q3)",
    p90: "90th percentile",
    p95: "95th percentile",
    p99: "99th percentile",
    mode: "Most frequently occurring value",
    null_count: "Number of null / missing values",
    null_pct: "Percentage of null / missing values",
    zero_count: "Number of zero values",
    negative_count: "Number of negative values",
};

type SortDir = "asc" | "desc" | null;

export class Visual implements IVisual {
    private target: HTMLElement;
    private host: powerbi.extensibility.visual.IVisualHost;

    private selectedFields: Set<number> = new Set();
    private fieldSearch: string = "";

    private filterManager: FilterManager;

    private filters;
    private filterSearch;

    private sortCol: number | null = null;
    private sortDir: SortDir = null;

    private columns: PbiTableColumn[] = [];
    private rows: PbiTableRow[] = [];
    private allRows: PbiTableRow[] = [];

    private columnProfiles: Map<number, ColumnProfile> = new Map();

    private showSummaryStats: boolean = false;
    private selectedStats: Set<StatKey> = new Set(["count", "min", "max", "mean"]);

    private filteredRowsCache: PbiTableRow[] | null = null;
    private filteredRowsCacheKey: string = "";

    private groupedRowsCache: PbiTableRow[] | null = null;
    private groupedRowsCacheKey: string = "";

    private hasMoreData: boolean = false;
    private isLoadingMore: boolean = false;

    private openDropdown: HTMLElement | null = null;
    private openDropdownCloseFn: (() => void) | null = null;

    // --- COPY / SELECTION STATE ---
    private selectedCells: Set<string> = new Set(); // "row_col"
    private selectionAnchor: { r: number, c: number } | null = null;
    private isDraggingSelection: boolean = false;

    private readonly MAX_RENDER_ROWS = 2000;
    private readonly MAX_STATS_ROWS = 10000;
    private readonly MAX_DISTINCT_FILTER_VALUES = 1000;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;

        this.filterManager = new FilterManager();

        this.filters = this.filterManager.getFilters();
        this.filterSearch = this.filterManager.getSearch();

        // END DRAG when mouse released anywhere
        document.addEventListener("mouseup", () => {
            this.isDraggingSelection = false;
            this.selectionAnchor = null;
        });

        document.addEventListener("mousedown", (e: MouseEvent) => {
            const table = this.target.querySelector(".data-table");
            if (table && !table.contains(e.target as Node)) {
                this.selectedCells.clear();
                this.refreshCellSelectionStyles();
            }
        });

        // CTRL/CMD + C copy shortcut
        this.target.setAttribute("tabindex", "0");
        this.target.addEventListener("keydown", (e: KeyboardEvent) => {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c") {
                this.copySelectionToClipboard();
            }
        });
    }

    public update(options: VisualUpdateOptions) {
        const dataView = options.dataViews?.[0];

        if (!dataView || !dataView.table) {
            Array.from(this.target.children).forEach(c => {
                if (c.tagName !== "STYLE") this.target.removeChild(c);
            });
            const msg = document.createElement("div");
            msg.className = "empty-msg";
            msg.innerText = "No data";
            this.target.appendChild(msg);
            return;
        }

        const table = dataView.table;
        this.columns = table.columns as PbiTableColumn[];

        const incomingRows = (table.rows || []) as PbiTableRow[];

        if (options.operationKind === powerbi.VisualDataChangeOperationKind.Create) {
            this.allRows = incomingRows.slice();
        } else if (incomingRows.length > 0) {
            this.allRows.push(...incomingRows);
        }

        this.rows = this.allRows;
        this.hasMoreData = !!dataView.metadata?.segment;
        this.isLoadingMore = false;
        this.invalidateCaches();
        this.prepareData();

        Array.from(this.target.children).forEach(c => {
            if (c.tagName !== "STYLE") this.target.removeChild(c);
        });

        if (this.selectedFields.size === 0 && this.columns.length > 0) {
            this.selectedFields.add(0);
        }

        const root = document.createElement("div");
        root.className = "pbi-root";

        const fieldPanel = document.createElement("div");
        fieldPanel.className = "field-panel";

        const fieldTitle = document.createElement("div");
        fieldTitle.className = "panel-title";
        fieldTitle.innerText = "Fields";
        fieldPanel.appendChild(fieldTitle);

        const fieldPanelInner = document.createElement("div");
        fieldPanelInner.className = "field-panel-inner";

        const searchInput = document.createElement("input");
        searchInput.type = "text";
        searchInput.className = "field-search";
        searchInput.placeholder = "Search fields…";
        searchInput.value = this.fieldSearch;
        searchInput.oninput = () => {
            this.fieldSearch = searchInput.value;
            this.renderFieldList(fieldListContainer, tablePanel);
        };
        fieldPanelInner.appendChild(searchInput);

        const actionsRow = document.createElement("div");
        actionsRow.className = "action-row";

        const btnAll = document.createElement("button");
        btnAll.className = "btn";
        btnAll.innerText = "All";
        btnAll.onclick = () => {
            this.columns.forEach((_, i) => this.selectedFields.add(i));
            this.invalidateCaches();
            this.renderFieldList(fieldListContainer, tablePanel);
            this.renderMainContent(tablePanel);
        };

        const btnClear = document.createElement("button");
        btnClear.className = "btn";
        btnClear.innerText = "Clear";
        btnClear.onclick = () => {
            this.selectedFields.clear();
            this.filterManager.clearAll();
            this.sortCol = null;
            this.sortDir = null;
            this.invalidateCaches();
            this.renderFieldList(fieldListContainer, tablePanel);
            this.renderMainContent(tablePanel);
        };

        const btnResetFilters = document.createElement("button");
        btnResetFilters.className = "btn";
        btnResetFilters.innerText = "Reset Filters";
        btnResetFilters.onclick = () => {
            this.filterManager.clearAll();
            this.invalidateCaches();
            this.renderMainContent(tablePanel);
        };

        actionsRow.appendChild(btnAll);
        actionsRow.appendChild(btnClear);
        actionsRow.appendChild(btnResetFilters);
        fieldPanelInner.appendChild(actionsRow);

        const fieldListContainer = document.createElement("div");
        fieldListContainer.className = "field-list";
        fieldPanelInner.appendChild(fieldListContainer);

        fieldPanel.appendChild(fieldPanelInner);

        const statsSection = document.createElement("div");

        const statsToggleRow = document.createElement("div");
        statsToggleRow.className = "stats-toggle-row";

        const statsCheckbox = document.createElement("input");
        statsCheckbox.type = "checkbox";
        statsCheckbox.checked = this.showSummaryStats;

        const statsToggleLabel = document.createElement("label");
        statsToggleLabel.innerText = "Summary Stats";

        statsCheckbox.onchange = () => {
            this.showSummaryStats = statsCheckbox.checked;
            statOptionsContainer.style.display = this.showSummaryStats ? "flex" : "none";
            this.renderMainContent(tablePanel);
        };

        statsToggleRow.appendChild(statsCheckbox);
        statsToggleRow.appendChild(statsToggleLabel);
        statsSection.appendChild(statsToggleRow);

        const statOptionsContainer = document.createElement("div");
        statOptionsContainer.className = "stat-options";
        statOptionsContainer.style.display = this.showSummaryStats ? "flex" : "none";

        const statGroups: { label: string; keys: StatKey[] }[] = [
            {
                label: "Basic",
                keys: ["count", "unique", "min", "max", "mean", "median", "sum", "stddev"],
            },
            {
                label: "Distribution",
                keys: ["skewness", "kurtosis", "excess_kurtosis", "jarque_bera", "mode"],
            },
            {
                label: "Percentiles",
                keys: ["p1", "p5", "p10", "p25", "p75", "p90", "p95", "p99"],
            },
            {
                label: "Data Quality",
                keys: ["null_count", "null_pct", "zero_count", "negative_count"],
            },
        ];

        statGroups.forEach(group => {
            const groupLabel = document.createElement("div");
            groupLabel.className = "stat-group-label";
            groupLabel.innerText = group.label;
            statOptionsContainer.appendChild(groupLabel);

            group.keys.forEach(key => {
                const statMeta = NUMERIC_STATS.find(s => s.key === key);
                if (!statMeta) return;

                const row = document.createElement("div");
                row.className = "stat-option-row";

                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.checked = this.selectedStats.has(key);
                cb.onchange = () => {
                    if (cb.checked) this.selectedStats.add(key);
                    else this.selectedStats.delete(key);
                    this.renderMainContent(tablePanel);
                };

                const lbl = document.createElement("label");
                lbl.innerText = statMeta.label;

                const tooltip = document.createElement("span");
                tooltip.className = "stat-tooltip";
                tooltip.innerText = STAT_TOOLTIP[key];

                row.addEventListener("mouseenter", () => { tooltip.style.opacity = "1"; });
                row.addEventListener("mouseleave", () => { tooltip.style.opacity = "0"; });
                row.addEventListener("mousemove", (e: MouseEvent) => {
                    tooltip.style.left = (e.clientX + 16) + "px";
                    tooltip.style.top = (e.clientY - 6) + "px";
                });

                row.appendChild(cb);
                row.appendChild(lbl);
                row.appendChild(tooltip);
                statOptionsContainer.appendChild(row);
            });
        });

        statsSection.appendChild(statOptionsContainer);
        fieldPanel.appendChild(statsSection);

        const tablePanel = document.createElement("div");
        tablePanel.className = "table-panel";

        root.appendChild(fieldPanel);
        root.appendChild(tablePanel);
        this.target.appendChild(root);

        this.renderFieldList(fieldListContainer, tablePanel);
        this.renderMainContent(tablePanel);
    }

    private invalidateCaches(): void {
        this.filteredRowsCache = null;
        this.filteredRowsCacheKey = "";
        this.groupedRowsCache = null;
        this.groupedRowsCacheKey = "";
    }

    private getGroupCacheKey(selected: number[], filterKey: string): string {
        return `${filterKey}__${selected.join(",")}__${this.sortCol ?? "null"}__${this.sortDir ?? "null"}`;
    }

    private prepareData() {
        this.columnProfiles.clear();
        this.columns.forEach((_, index) => {
            const profile = this.buildProfile(this.rows, index);
            this.columnProfiles.set(index, profile);
        });
    }

    private buildProfile(rows: PbiTableRow[], index: number): ColumnProfile {
        const allValues: string[] = [];
        const numericValues: number[] = [];
        let nullCount = 0;
        let zeroCount = 0;
        let negativeCount = 0;

        rows.forEach(row => {
            const raw = row[index];

            if (isBlank(raw)) {
                nullCount++;
                allValues.push(BLANK_TOKEN);
                return;
            } 

            const v = formatCell(raw);
            allValues.push(v);

            const n = parseFloat(v);
            if (!isNaN(n)) {
                numericValues.push(n);
                if (n === 0) zeroCount++;
                if (n < 0) negativeCount++;
            }
        });

        const unique = Array.from(new Set(allValues));
        const nonNullStrings = unique.filter(v => v !== BLANK_TOKEN);

        const totalRows = rows.length;
        const isNumeric = totalRows > 0 && (numericValues.length / totalRows) > 0.8;

        const profile: ColumnProfile = {
            type: isNumeric ? "numeric" : "text",
            uniqueValues: unique,
            numericValues,
            totalRowCount: totalRows,
        };

        if (isNumeric && numericValues.length > 0) {
            const sorted = [...numericValues].sort((a, b) => a - b);
            const n = sorted.length;
            const sum = sorted.reduce((a, b) => a + b, 0);
            const mean = sum / n;
            const median = n % 2 === 0
                ? (sorted[n / 2 - 1] + sorted[n / 2]) / 2
                : sorted[Math.floor(n / 2)];
            const variance = sorted.reduce((acc, v) => acc + (v - mean) ** 2, 0) / n;
            const stddev = Math.sqrt(variance);

            const skewness = n > 2 && stddev > 0
                ? sorted.reduce((acc, v) => acc + ((v - mean) / stddev) ** 3, 0) / n
                : 0;

            const kurtosis = n > 3 && stddev > 0
                ? sorted.reduce((acc, v) => acc + ((v - mean) / stddev) ** 4, 0) / n
                : 0;

            const excess_kurtosis = kurtosis - 3;

            const jarque_bera = n > 0
                ? (n / 6) * (skewness ** 2 + (excess_kurtosis ** 2) / 4)
                : 0;

            const freqMap = new Map<number, number>();
            numericValues.forEach(v => freqMap.set(v, (freqMap.get(v) ?? 0) + 1));
            let mode: number | undefined;
            let maxFreq = 0;
            freqMap.forEach((freq, val) => {
                if (freq > maxFreq) {
                    maxFreq = freq;
                    mode = val;
                }
            });

            profile.stats = {
                count: n,
                min: sorted[0],
                max: sorted[n - 1],
                mean,
                median,
                sum,
                stddev,
                unique: unique.length,
                skewness,
                kurtosis,
                excess_kurtosis,
                jarque_bera,
                p1: this.percentile(sorted, 1),
                p5: this.percentile(sorted, 5),
                p10: this.percentile(sorted, 10),
                p25: this.percentile(sorted, 25),
                p75: this.percentile(sorted, 75),
                p90: this.percentile(sorted, 90),
                p95: this.percentile(sorted, 95),
                p99: this.percentile(sorted, 99),
                mode,
                null_count: nullCount,
                null_pct: totalRows > 0 ? (nullCount / totalRows) * 100 : 0,
                zero_count: zeroCount,
                negative_count: negativeCount,
            };
        } else {
            profile.stats = {
                count: nonNullStrings.length,
                unique: unique.length,
                null_count: nullCount,
                null_pct: totalRows > 0 ? (nullCount / totalRows) * 100 : 0,
                zero_count: 0,
                negative_count: 0,
            };
        }

        return profile;
    }

    private percentile(sorted: number[], p: number): number {
        const n = sorted.length;
        if (n === 0) return 0;
        if (n === 1) return sorted[0];
        const idx = (p / 100) * (n - 1);
        const lower = Math.floor(idx);
        const upper = Math.ceil(idx);
        const frac = idx - lower;
        return sorted[lower] + frac * (sorted[upper] - sorted[lower]);
    }

    private clearElement(element: HTMLElement) {
        while (element.firstChild) {
            element.removeChild(element.firstChild);
        }
    }

    private renderFieldList(container: HTMLElement, tablePanel: HTMLElement) {
        this.clearElement(container);
        const search = this.fieldSearch.toLowerCase();

        this.columns.forEach((col, index) => {
            if (search && !col.displayName.toLowerCase().includes(search)) return;

            const row = document.createElement("div");
            row.className = "field-row";

            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.checked = this.selectedFields.has(index);
            checkbox.onchange = () => {
                if (checkbox.checked) {
                    this.selectedFields.add(index);
                } else {
                    this.selectedFields.delete(index);
                    this.filters.delete(index);
                    this.filterSearch.delete(index);
                    if (this.sortCol === index) {
                        this.sortCol = null;
                        this.sortDir = null;
                    }
                }
                this.invalidateCaches();
                this.renderMainContent(tablePanel);
            };

            const label = document.createElement("label");
            label.innerText = col.displayName;
            label.title = col.displayName;

            row.appendChild(checkbox);
            row.appendChild(label);

            row.onclick = (e) => {
                if ((e.target as HTMLElement).tagName !== "INPUT") {
                    checkbox.checked = !checkbox.checked;
                    checkbox.onchange!(new Event("change"));
                }
            };

            container.appendChild(row);
        });
    }

    private buildMultiSelectFilter(colIndex: number, tableContainer: HTMLElement): HTMLElement {
        const profile = this.columnProfiles.get(colIndex);
        const rawValues: string[] = profile?.uniqueValues ?? [];

        const sortedValues = [...rawValues].sort((a, b) => {
            // Put blanks first
            if (a === BLANK_TOKEN) return -1;
            if (b === BLANK_TOKEN) return 1;

            // Normal alphabetical sort
            return a.localeCompare(b);
        });

        const allValues: string[] = sortedValues.slice(0, this.MAX_DISTINCT_FILTER_VALUES);

        // console.log("Filter values for column", colIndex, allValues);
        console.table({
            rawValues,
            sortedValues,
            allValues
        });

        const activeFilter: Set<string> = this.filters.get(colIndex) ?? new Set();

        const wrap = document.createElement("div");
        wrap.className = "col-filter-wrap";

        const btn = document.createElement("button");
        btn.className = "col-filter-btn" + (activeFilter.size > 0 ? " has-filter" : "");

        const filterLabel = document.createElement("span");
        filterLabel.className = "filter-label";
        filterLabel.innerText = this.filterManager.getButtonLabel(activeFilter, allValues.length);

        const arrow = document.createElement("span");
        arrow.className = "filter-arrow";
        arrow.innerText = "▼";

        btn.appendChild(filterLabel);
        btn.appendChild(arrow);
        wrap.appendChild(btn);

        let dropdown: HTMLElement | null = null;
        let outsideClickHandler: ((e: MouseEvent) => void) | null = null;

        const openDropdown = () => {
            // close any existing open dropdown first
            if (this.openDropdownCloseFn) {
                this.openDropdownCloseFn();
            }

            dropdown = document.createElement("div");
            dropdown.className = "col-filter-dropdown";


            this.openDropdown = dropdown;
            this.openDropdownCloseFn = closeDropdown;

            const searchInput = document.createElement("input");
            searchInput.type = "text";
            searchInput.className = "col-filter-search";
            searchInput.placeholder = "Search values…";
            searchInput.value = this.filterSearch.get(colIndex) ?? "";
            dropdown.appendChild(searchInput);

            if (rawValues.length > this.MAX_DISTINCT_FILTER_VALUES) {
                const note = document.createElement("div");
                note.className = "truncate-note";
                note.innerText = `Showing first ${this.MAX_DISTINCT_FILTER_VALUES.toLocaleString()} values`;
                dropdown.appendChild(note);
            }

            const actionsBar = document.createElement("div");
            actionsBar.className = "col-filter-actions";

            const btnSelectAll = document.createElement("button");
            btnSelectAll.className = "btn";
            btnSelectAll.innerText = "All";
            btnSelectAll.onclick = (e) => {
                e.stopPropagation();
                this.filters.delete(colIndex);
                this.filterSearch.delete(colIndex);
                this.invalidateCaches();
                closeDropdown();
                this.renderMainContent(tableContainer);
            };

            const btnClearSel = document.createElement("button");
            btnClearSel.className = "btn";
            btnClearSel.innerText = "None";
            btnClearSel.onclick = (e) => {
                e.stopPropagation();
                this.filters.set(colIndex, new Set());
                this.invalidateCaches();
                renderList("");
            };

            actionsBar.appendChild(btnSelectAll);
            actionsBar.appendChild(btnClearSel);
            dropdown.appendChild(actionsBar);

            const list = document.createElement("div");
            list.className = "col-filter-list";
            dropdown.appendChild(list);

            const renderList = (searchTerm: string) => {
                this.clearElement(list);
                const term = searchTerm.toLowerCase();
                const visible = term
                    ? allValues.filter(v => v.toLowerCase().includes(term))
                    : allValues;

                if (visible.length === 0) {
                    const empty = document.createElement("div");
                    empty.className = "col-filter-empty";
                    empty.innerText = "No matching values";
                    list.appendChild(empty);
                    return;
                }

                const current = this.filters.get(colIndex);

                visible.forEach(v => {
                    const item = document.createElement("div");
                    item.className = "col-filter-item";

                    const cb = document.createElement("input");
                    cb.type = "checkbox";
                    cb.checked = !current || current.has(v);

                    const lbl = document.createElement("label");
                    lbl.innerText = displayFilterValue(v);
                    lbl.title = displayFilterValue(v);

                    cb.onchange = () => {
                        let sel = this.filters.get(colIndex);
                        if (!sel) {
                            sel = new Set(allValues);
                            this.filters.set(colIndex, sel);
                        }

                        if (cb.checked) sel.add(v);
                        else sel.delete(v);

                        if (sel.size >= allValues.length) {
                            this.filters.delete(colIndex);
                        }

                        this.invalidateCaches();

                        const newActive = this.filters.get(colIndex) ?? new Set();
                        filterLabel.innerText = this.filterManager.getButtonLabel(newActive, allValues.length);
                        btn.className = "col-filter-btn" + (newActive.size > 0 ? " has-filter open" : " open");
                        this.renderMainContent(tableContainer);
                    };

                    item.appendChild(cb);
                    item.appendChild(lbl);
                    item.onclick = (e) => {
                        if ((e.target as HTMLElement).tagName !== "INPUT") {
                            cb.checked = !cb.checked;
                            cb.onchange!(new Event("change"));
                        }
                    };
                    list.appendChild(item);
                });
            };

            searchInput.oninput = () => {
                this.filterSearch.set(colIndex, searchInput.value);
                renderList(searchInput.value);
            };

            renderList(searchInput.value);

            wrap.appendChild(dropdown);
            btn.classList.add("open");

            outsideClickHandler = (e: MouseEvent) => {
                if (!wrap.contains(e.target as Node)) closeDropdown();
            };
            setTimeout(() => document.addEventListener("click", outsideClickHandler!), 0);
        };

        const closeDropdown = () => {
            
            if (this.openDropdown === dropdown) {
                this.openDropdown = null;
                this.openDropdownCloseFn = null;
            }

            if (dropdown) {
                dropdown.remove();
                dropdown = null;
            }
            if (outsideClickHandler) {
                document.removeEventListener("click", outsideClickHandler);
                outsideClickHandler = null;
            }
            btn.classList.remove("open");
            const cur = this.filters.get(colIndex) ?? new Set();
            filterLabel.innerText = this.filterManager.getButtonLabel(cur, allValues.length);
            btn.className = "col-filter-btn" + (cur.size > 0 ? " has-filter" : "");
        };

        btn.onclick = (e) => {
            e.stopPropagation();
            if (dropdown) closeDropdown();
            else openDropdown();
        };

        return wrap;
    }

    private getGroupedRows(filteredRows: PbiTableRow[], selected: number[]): PbiTableRow[] {
        const filterKey = this.filterManager.getCacheKey();
        const groupKey = this.getGroupCacheKey(selected, filterKey);

        if (this.groupedRowsCache && this.groupedRowsCacheKey === groupKey) {
            return this.groupedRowsCache;
        }

        const groupedMap = new Map<string, PbiTableRow>();

        for (const row of filteredRows) {
            const key = selected.map(i => formatCell(row[i])).join("||");
            if (!groupedMap.has(key)) groupedMap.set(key, row);
        }

        let result = Array.from(groupedMap.values());

        if (this.sortCol !== null && this.sortDir !== null) {
            const sortColIndex = this.sortCol;
            const dir = this.sortDir;
            const profile = this.columnProfiles.get(sortColIndex);
            const isNumeric = profile?.type === "numeric";

            result = result.sort((a, b) => {
                const aRaw = formatCell(a[sortColIndex]);
                const bRaw = formatCell(b[sortColIndex]);

                let cmp: number;
                if (isNumeric) {
                    const aNum = parseFloat(aRaw);
                    const bNum = parseFloat(bRaw);
                    const aIsNaN = isNaN(aNum);
                    const bIsNaN = isNaN(bNum);

                    if (aIsNaN && bIsNaN) cmp = 0;
                    else if (aIsNaN) cmp = 1;
                    else if (bIsNaN) cmp = -1;
                    else cmp = aNum - bNum;
                } else {
                    cmp = aRaw.localeCompare(bRaw);
                }

                return dir === "asc" ? cmp : -cmp;
            });
        }

        this.groupedRowsCache = result;
        this.groupedRowsCacheKey = groupKey;
        return result;
    }

    private renderMainContent(container: HTMLElement) {
        this.clearElement(container);

        const filteredRows = this.filterManager.getFilteredRows(this.rows);
        const isFiltered = this.filters.size > 0;

        if (this.showSummaryStats && this.selectedStats.size > 0) {
            this.renderStatsCards(container, filteredRows, isFiltered);
        }

        this.renderTable(container, filteredRows);

        if (this.hasMoreData || this.isLoadingMore) {
            container.appendChild(this.buildLoadMoreBar(container));
        }
    }

    private buildLoadMoreBar(container: HTMLElement): HTMLElement {
        const bar = document.createElement("div");
        bar.className = "row-count-bar";

        const btn = document.createElement("button");
        btn.className = "btn";
        btn.innerText = this.isLoadingMore ? "Loading..." : "Load more rows";
        btn.disabled = this.isLoadingMore;

        btn.onclick = () => {
            if (this.isLoadingMore) return;

            this.isLoadingMore = true;
            this.renderMainContent(container);

            const accepted = this.host.fetchMoreData(false);
            if (!accepted) {
                this.isLoadingMore = false;
                this.hasMoreData = false;
                this.renderMainContent(container);
            }
        };

        bar.appendChild(btn);

        const status = document.createElement("span");
        status.className = "load-more-status";
        status.innerText = `${this.rows.length.toLocaleString()} rows loaded`;
        bar.appendChild(status);

        if (this.hasMoreData) {
            const note = document.createElement("span");
            note.innerText = "More rows available";
            bar.appendChild(note);
        }

        return bar;
    }

    private renderStatsCards(
        container: HTMLElement,
        filteredRows: PbiTableRow[],
        isFiltered: boolean,
    ) {
        const selected = Array.from(this.selectedFields);
        if (selected.length === 0) return;

        const allAvailableKeys = NUMERIC_STATS.map(s => s.key);
        const statsToShow = allAvailableKeys.filter(k => this.selectedStats.has(k));
        if (statsToShow.length === 0) return;

        const statsRows =
            filteredRows.length > this.MAX_STATS_ROWS
                ? filteredRows.slice(0, this.MAX_STATS_ROWS)
                : filteredRows;

        const liveProfiles = new Map<number, ColumnProfile>();
        selected.forEach(i => {
            liveProfiles.set(i, this.buildProfile(statsRows, i));
        });

        const section = document.createElement("div");
        section.className = "stats-section";

        const sectionTitle = document.createElement("div");
        sectionTitle.className = "stats-section-title";
        sectionTitle.innerText =
            filteredRows.length > this.MAX_STATS_ROWS
                ? `Summary Statistics (sample of ${this.MAX_STATS_ROWS.toLocaleString()} rows)`
                : "Summary Statistics";

        if (isFiltered) {
            const badge = document.createElement("span");
            badge.className = "stats-filtered-badge";
            badge.innerText = "Filtered";
            sectionTitle.appendChild(badge);
        }

        section.appendChild(sectionTitle);

        const table = document.createElement("table");
        table.className = "stats-table";

        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");

        const cornerTh = document.createElement("th");
        cornerTh.innerText = "Field";
        headerRow.appendChild(cornerTh);

        statsToShow.forEach(key => {
            const statMeta = NUMERIC_STATS.find(s => s.key === key);
            const th = document.createElement("th");
            th.className = "stat-col";
            th.innerText = statMeta?.label ?? key;
            th.title = STAT_TOOLTIP[key];
            headerRow.appendChild(th);
        });

        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");

        selected.forEach(i => {
            const profile = liveProfiles.get(i);
            const col = this.columns[i];
            const tr = document.createElement("tr");

            const nameTd = document.createElement("td");
            nameTd.className = "field-name-cell";
            nameTd.innerText = col.displayName;
            nameTd.title = col.displayName;
            tr.appendChild(nameTd);

            const numericOnlyStats: StatKey[] = [
                "min", "max", "mean", "median", "sum", "stddev",
                "skewness", "kurtosis", "excess_kurtosis", "jarque_bera",
                "p1", "p5", "p10", "p25", "p75", "p90", "p95", "p99",
                "mode", "zero_count", "negative_count",
            ];

            statsToShow.forEach(key => {
                const td = document.createElement("td");
                td.className = "stat-val-cell";

                if (!profile || !profile.stats) {
                    td.innerText = "—";
                    td.style.color = "#bbb";
                } else if (profile.type === "text" && numericOnlyStats.includes(key)) {
                    td.innerText = "—";
                    td.style.color = "#bbb";
                } else {
                    const val = (profile.stats as Record<string, number | undefined>)[key];
                    td.innerText = val !== undefined ? this.formatStat(key, val) : "—";
                }

                tr.appendChild(td);
            });

            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        section.appendChild(table);
        container.appendChild(section);
    }

    private formatStat(key: StatKey, val: number): string {
        if (key === "count" || key === "unique" || key === "null_count" || key === "zero_count" || key === "negative_count") {
            return String(Math.round(val));
        }
        if (key === "null_pct") {
            return val.toLocaleString(undefined, { maximumFractionDigits: 1 }) + "%";
        }
        if (key === "sum") {
            return val.toLocaleString(undefined, { maximumFractionDigits: 2 });
        }
        if (key === "jarque_bera") {
            return val.toLocaleString(undefined, { maximumFractionDigits: 2 });
        }
        return val.toLocaleString(undefined, { maximumFractionDigits: 4 });
    }

    private showCopyToast() {
        let toast = this.target.querySelector(".copy-toast") as HTMLElement;
        if (!toast) {
            toast = document.createElement("div");
            toast.className = "copy-toast";
            toast.innerText = "Copied to clipboard";
            this.target.appendChild(toast);
        }
        toast.classList.add("show");
        setTimeout(() => toast.classList.remove("show"), 1800);
    }

    private copySelectionToClipboard() {
        if (this.selectedCells.size === 0) return;

        // Convert selection to grid
        const cells = Array.from(this.selectedCells).map(k => {
            const [r, c] = k.split("_").map(Number);
            return { r, c };
        });

        const minR = Math.min(...cells.map(c => c.r));
        const maxR = Math.max(...cells.map(c => c.r));
        const minC = Math.min(...cells.map(c => c.c));
        const maxC = Math.max(...cells.map(c => c.c));

        // Build header row from visible column names in the selected range
        const selected = Array.from(this.selectedFields);
        const headerCols: string[] = [];
        for (let c = minC; c <= maxC; c++) {
            const colIndex = selected[c];
            headerCols.push(colIndex !== undefined ? this.columns[colIndex].displayName : "");
        }

        // Build data rows
        let text = headerCols.join("\t") + "\n";

        for (let r = minR; r <= maxR; r++) {
            const rowVals: string[] = [];
            for (let c = minC; c <= maxC; c++) {
                const td = this.target.querySelector(`[data-cell="${r}_${c}"]`) as HTMLElement;
                rowVals.push(td ? td.innerText : "");
            }
            text += rowVals.join("\t") + "\n";
        }

        navigator.clipboard.writeText(text).catch(() => {
            // Fallback for non-secure contexts
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.opacity = "0";
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
        });

        this.showCopyToast();
    }

    private refreshCellSelectionStyles() {
        this.target.querySelectorAll("[data-cell]").forEach(el => {
            el.classList.remove("copied");
        });

        this.selectedCells.forEach(key => {
            const el = this.target.querySelector(`[data-cell="${key}"]`);
            if (el) el.classList.add("copied");
        });
    }

    private renderTable(container: HTMLElement, filteredRows: PbiTableRow[]) {
        const selected = Array.from(this.selectedFields);

        if (selected.length === 0) {
            const msg = document.createElement("div");
            msg.className = "empty-msg";
            msg.innerText = "Select at least one field to display.";
            container.appendChild(msg);
            return;
        }

        const displayRows = this.getGroupedRows(filteredRows, selected);

        const wrap = document.createElement("div");
        wrap.className = "data-table-wrap";

        const table = document.createElement("table");
        table.className = "data-table";

        const thead = document.createElement("thead");
        const header = document.createElement("tr");

        selected.forEach(colIndex => {
            const th = document.createElement("th");

            const titleRow = document.createElement("div");
            titleRow.className = "col-header-title-row";
            titleRow.title = "Click to sort";

            const nameSpan = document.createElement("span");
            nameSpan.className = "col-header-name";
            nameSpan.innerText = this.columns[colIndex].displayName;

            const sortBtn = document.createElement("span");
            const isSorted = this.sortCol === colIndex;
            sortBtn.className = "sort-btn" + (isSorted ? " active" : "");

            const upArrow = document.createElement("span");
            upArrow.className = "sort-arrow" + (isSorted && this.sortDir === "asc" ? " active-arrow" : "");
            upArrow.innerText = "▲";

            const downArrow = document.createElement("span");
            downArrow.className = "sort-arrow" + (isSorted && this.sortDir === "desc" ? " active-arrow" : "");
            downArrow.innerText = "▼";

            sortBtn.appendChild(upArrow);
            sortBtn.appendChild(downArrow);

            titleRow.appendChild(nameSpan);
            titleRow.appendChild(sortBtn);
            th.appendChild(titleRow);

            titleRow.onclick = () => {
                if (this.sortCol !== colIndex) {
                    this.sortCol = colIndex;
                    this.sortDir = "asc";
                } else if (this.sortDir === "asc") {
                    this.sortDir = "desc";
                } else if (this.sortDir === "desc") {
                    this.sortCol = null;
                    this.sortDir = null;
                }
                this.invalidateCaches();
                this.renderMainContent(container);
            };

            const filterWidget = this.buildMultiSelectFilter(colIndex, container);
            th.appendChild(filterWidget);

            header.appendChild(th);
        });

        thead.appendChild(header);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        const MAX_PREVIEW_ROWS = this.MAX_RENDER_ROWS;

        displayRows.slice(0, MAX_PREVIEW_ROWS).forEach((row, rowIndex) => {

            const tr = document.createElement("tr");

            selected.forEach((colIndex, visibleColIndex) => {

                const td = document.createElement("td");
                const cellKey = `${rowIndex}_${visibleColIndex}`;
                td.dataset.cell = cellKey;

                const value = formatCell(row[colIndex]);
                td.innerText = value;

                // 🟦 START selection
                td.onmousedown = (e: MouseEvent) => {
                    e.preventDefault();
                    this.target.focus({ preventScroll: true });
                    this.isDraggingSelection = true;
                    this.selectionAnchor = { r: rowIndex, c: visibleColIndex };
                    this.selectedCells.clear();
                    this.selectedCells.add(cellKey);
                    this.refreshCellSelectionStyles();
                };

                // 🟦 DRAG selection
                td.onmouseover = () => {
                    if (!this.isDraggingSelection || !this.selectionAnchor) return;

                    this.selectedCells.clear();

                    const r1 = this.selectionAnchor.r;
                    const c1 = this.selectionAnchor.c;
                    const r2 = rowIndex;
                    const c2 = visibleColIndex;

                    const minR = Math.min(r1, r2);
                    const maxR = Math.max(r1, r2);
                    const minC = Math.min(c1, c2);
                    const maxC = Math.max(c1, c2);

                    for (let r = minR; r <= maxR; r++) {
                        for (let c = minC; c <= maxC; c++) {
                            this.selectedCells.add(`${r}_${c}`);
                        }
                    }

                    this.refreshCellSelectionStyles();
                };

                tr.appendChild(td);
            });

            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        wrap.appendChild(table);
        container.appendChild(wrap);

        const bar = document.createElement("div");
        bar.className = "row-count-bar";
        const shown = Math.min(displayRows.length, MAX_PREVIEW_ROWS);
        const sortInfo = this.sortCol !== null && this.sortDir !== null
            ? `  |  sorted by ${this.columns[this.sortCol].displayName} ${this.sortDir.toUpperCase()}`
            : "";

        bar.innerText = displayRows.length > MAX_PREVIEW_ROWS
            ? `Showing ${shown.toLocaleString()} of ${displayRows.length.toLocaleString()} rows (preview limited for performance)  |  ${this.rows.length.toLocaleString()} loaded${sortInfo}`
            : `${displayRows.length.toLocaleString()} rows  |  ${this.rows.length.toLocaleString()} loaded${sortInfo}`;

        container.appendChild(bar);
    }

}
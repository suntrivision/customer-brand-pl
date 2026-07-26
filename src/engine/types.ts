export const MONTHS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
] as const;

export type Month = (typeof MONTHS)[number];

export type Filters = {
  month: Month;
  channel: string;
  brand: string;
};

/** key → 12 monthly values (Jan..Dec), missing = 0 */
export type MonthlyMap = Map<string, number[]>;

export type WorkbookData = {
  /** Working(Local) flat lookup: account+channel+brand+month+year → amount */
  working: Map<string, number>;
  /** P&L channel (AD) → Budget channel (AC) */
  channelMap: Map<string, string>;
  budget: MonthlyMap;
  sellOutQty2025: MonthlyMap;
  sellOutQty2026: MonthlyMap;
  sellOutAmt2025: MonthlyMap;
  sellOutAmt2026: MonthlyMap;
  sellInQty2025: MonthlyMap;
  sellInAmt2025: MonthlyMap;
  sellInQty2026: MonthlyMap;
  sellInAmt2026: MonthlyMap;
  channels: string[];
  brands: string[];
};

export type MetricTriple = {
  ly: number | null;
  budget: number | null;
  actual: number | null;
};

export type PLRowValues = {
  mtd: MetricTriple;
  ytd: MetricTriple;
  mtdIndexBudget: number | null;
  mtdIndexLy: number | null;
  ytdIndexBudget: number | null;
  ytdIndexLy: number | null;
  mtdRatio: MetricTriple;
  ytdRatio: MetricTriple;
};

export type PLRow = {
  id: string;
  label: string;
  indent: number;
  bold?: boolean;
  isAup?: boolean;
  isRatioOnly?: boolean;
  values: PLRowValues;
  /** MTD Actual for Jan..Dec under the same Channel/Brand */
  monthlyActuals: number[];
};

export type PLReport = {
  filters: Filters;
  rows: PLRow[];
};

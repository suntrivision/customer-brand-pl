import type { Filters, PLReport, WorkbookData } from "./types";
import { computePL } from "./computePL";

/**
 * Line order / labels aligned to ComparitiveAnalysis.xlsx
 * (sheet "Comparitive Analysis"), mapped to Local P&L rows where possible.
 */
export type CompareLineDef = {
  id: string;
  label: string;
  /** Matching P&L row id in computePL, if any */
  plRowId: string | null;
  /** Spacer / blank row in the Excel layout */
  spacer?: boolean;
};

export const COMPARE_LINES: CompareLineDef[] = [
  { id: "gsv", label: "Sell In (Gross Sales Value - GSV) - Total", plRowId: "gsv" },
  { id: "oid", label: "On Invoice Discount", plRowId: "oid" },
  { id: "dist-margin", label: "Distributor Margin", plRowId: "dist-margin" },
  { id: "net-invoice", label: "Net Invoice sales", plRowId: null },
  { id: "spacer-1", label: "", plRowId: null, spacer: true },
  { id: "trade-spend", label: "Trade Spend", plRowId: "trade-spend" },
  { id: "sales-exp", label: "Sales Expenses", plRowId: "sales-exp" },
  { id: "ppd", label: "PPD", plRowId: "ppd" },
  { id: "kby-reimb", label: "Kobayashi Reimbursement", plRowId: "kby-reimb" },
  { id: "dc-charges", label: "DC Charges", plRowId: "dc-charges" },
  { id: "listing-fee", label: "Listing Fee", plRowId: "listing-fee" },
  { id: "list-price-disc", label: "List Price Discount", plRowId: "list-price-disc" },
  { id: "spacer-2", label: "", plRowId: null, spacer: true },
  { id: "net-revenue", label: "Net Revenue", plRowId: "net-sales" },
  { id: "cogs", label: "Cost of Goods Sold", plRowId: null },
  { id: "cogs-direct", label: "Cost of Sales - Direct", plRowId: "cogs-direct" },
  { id: "cogs-indirect", label: "Cost of Sales - Indirect", plRowId: "cogs-indirect" },
  { id: "indir-mk", label: "Indirect COGS-MK", plRowId: "indir-mk" },
  { id: "indir-sc", label: "Indirect COGS-SC", plRowId: "indir-sc" },
  { id: "twinpack", label: "Kobayashi Reimbursement (Twinpack rebate)", plRowId: "twinpack" },
  { id: "logistic", label: "Logistic Cost", plRowId: "logistic" },
  { id: "cogs-logistic", label: "COGS & Logistic", plRowId: "cogs-logistic" },
  {
    id: "gm-after-promoter",
    label: "Gross Margin After Promoter & Merchandiser",
    plRowId: "gross-margin",
  },
  { id: "promoter", label: "Promoter & Merchandiser", plRowId: "promoter" },
  { id: "npd", label: "NPD", plRowId: "npd" },
  { id: "mkt-others", label: "Marketing Cost - Others", plRowId: "mkt-others" },
  { id: "promotion", label: "Promotion", plRowId: "promotion" },
  { id: "advertising", label: "Advertising", plRowId: "advertising" },
  { id: "mkt-cost", label: "Marketing Cost", plRowId: "mkt-cost" },
  { id: "profit-contrib", label: "Profit contribution", plRowId: "gp-after-mkt" },
  { id: "gp-after-mkt", label: "Gross Profit after Marketing", plRowId: "gp-after-mkt" },
  { id: "ga", label: "G&A Expenses", plRowId: "ga" },
  { id: "staff", label: "Staff Remuneration", plRowId: "staff" },
  { id: "other-ga", label: "Other G&A Cost", plRowId: "other-ga" },
  { id: "finance", label: "Finance Expenses", plRowId: "finance" },
  { id: "admin", label: "Administration Expenses", plRowId: "admin" },
  { id: "outdoor", label: "Outdoor Expenses", plRowId: "outdoor" },
  { id: "nop", label: "Net Operating Profit", plRowId: "nop" },
  { id: "other-op", label: "Other Operating (Income)/Expense", plRowId: "other-op" },
  { id: "other-exp", label: "Other Expenses", plRowId: "other-exp" },
  { id: "other-inc", label: "Other Income", plRowId: "other-inc" },
  { id: "spacer-3", label: "", plRowId: null, spacer: true },
  { id: "npbt", label: "Net Profit Before Tax", plRowId: "op-income" },
  { id: "tax", label: "Tax", plRowId: "tax" },
  { id: "npat", label: "Net Profit After Tax", plRowId: "nop" },
];

/** Input Actuals keyed by compare line id */
export type ActualsInput = Record<string, number | null>;

export type CompareRow = {
  id: string;
  label: string;
  spacer?: boolean;
  inputActual: number | null;
  plActual: number | null;
  /** Difference = P&L Actual − Input (same as ComparitiveAnalysis.xlsx) */
  diff: number | null;
  diffPct: number | null;
  hasInput: boolean;
};

export type CompareReport = {
  filters: Filters;
  sourceLabel: string;
  rows: CompareRow[];
  inputCount: number;
};

function normLabel(s: string): string {
  return s
    .toLowerCase()
    .replace(/['’]/g, "")
    .replace(/\s+/g, " ")
    .replace(/[()–—-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

const LABEL_TO_ID = new Map<string, string>();
for (const line of COMPARE_LINES) {
  if (line.spacer || !line.label) continue;
  LABEL_TO_ID.set(normLabel(line.label), line.id);
}
// Local P&L / mapper aliases
LABEL_TO_ID.set(normLabel("Sell in (Gross Sales Value - GSV) - Total"), "gsv");
LABEL_TO_ID.set(normLabel("Net Sales"), "net-revenue");
LABEL_TO_ID.set(normLabel("Gross Margin"), "gm-after-promoter");
LABEL_TO_ID.set(normLabel("Operating Income before Tax"), "npbt");
LABEL_TO_ID.set(normLabel("Net Operating Profit"), "nop");

export function emptyActualsInput(): ActualsInput {
  const out: ActualsInput = {};
  for (const line of COMPARE_LINES) {
    if (!line.spacer) out[line.id] = null;
  }
  return out;
}

function plAmt(
  byId: Map<string, { values: { mtd: { actual: number | null } } }>,
  id: string,
): number | null {
  const v = byId.get(id)?.values.mtd.actual;
  return v == null || !Number.isFinite(v) ? null : v;
}

function resolvePlActual(
  byId: Map<string, { values: { mtd: { actual: number | null } } }>,
  line: CompareLineDef,
): number | null {
  if (line.spacer) return null;
  if (line.plRowId) {
    return plAmt(byId, line.plRowId);
  }
  // Derived lines (Local P&L equivalents)
  if (line.id === "net-invoice") {
    const gsv = plAmt(byId, "gsv") ?? 0;
    const oid = plAmt(byId, "oid") ?? 0;
    const dist = plAmt(byId, "dist-margin") ?? 0;
    return gsv - oid - dist;
  }
  if (line.id === "cogs") {
    const d = plAmt(byId, "cogs-direct") ?? 0;
    const i = plAmt(byId, "cogs-indirect") ?? 0;
    return d + i;
  }
  return null;
}

export function parseActualsText(text: string): {
  input: ActualsInput;
  matched: number;
  unmatched: string[];
} {
  const input = emptyActualsInput();
  const unmatched: string[] = [];
  let matched = 0;

  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  for (const raw of lines) {
    // Formats: "Label,1234" | "Label\t1234" | "Label: 1234" | "Label 1234"
    const m =
      raw.match(/^(.+?)[,:\t]\s*(-?[\d,]+(?:\.\d+)?)\s*$/) ||
      raw.match(/^(.+?)\s+(-?[\d,]+(?:\.\d+)?)\s*$/);
    if (!m) {
      unmatched.push(raw);
      continue;
    }
    const label = m[1]!.trim();
    const num = Number(m[2]!.replace(/,/g, ""));
    if (!Number.isFinite(num)) {
      unmatched.push(raw);
      continue;
    }
    const id = LABEL_TO_ID.get(normLabel(label));
    if (!id) {
      unmatched.push(raw);
      continue;
    }
    input[id] = num;
    matched++;
  }

  return { input, matched, unmatched };
}

export function parseActualsRows(
  rows: { label: string; amount: number }[],
): { input: ActualsInput; matched: number; unmatched: string[] } {
  const input = emptyActualsInput();
  const unmatched: string[] = [];
  let matched = 0;
  for (const row of rows) {
    const id = LABEL_TO_ID.get(normLabel(row.label));
    if (!id || !Number.isFinite(row.amount)) {
      unmatched.push(row.label);
      continue;
    }
    input[id] = row.amount;
    matched++;
  }
  return { input, matched, unmatched };
}

/**
 * Parse ComparitiveAnalysis.xlsx layout:
 * Col A = P&L line, Col B = Trivision/input Actual, Col C = TTK customer Actual.
 * Optionally detect customer from header "Trivision - DIST-CHC" / "DIST-CHC".
 */
export function parseComparativeSheet(
  rows: (string | number | null | undefined)[][],
  customerHint?: string,
): {
  input: ActualsInput;
  matched: number;
  unmatched: string[];
  detectedCustomer: string | null;
  sourceLabel: string;
} {
  const input = emptyActualsInput();
  const unmatched: string[] = [];
  let matched = 0;
  let detectedCustomer: string | null = null;
  let sourceLabel = "Comparative Analysis Actuals";

  // Find the block for the requested customer (or first data block)
  let start = -1;
  for (let i = 0; i < rows.length; i++) {
    const a = String(rows[i]?.[0] ?? "").trim();
    const b = String(rows[i]?.[1] ?? "").trim();
    const c = String(rows[i]?.[2] ?? "").trim();
    if (/^trivision/i.test(b) || /^trivision/i.test(String(rows[i]?.[0] ?? ""))) {
      const cust = c || b.replace(/^trivision\s*-\s*/i, "").trim();
      if (customerHint) {
        if (normLabel(cust).includes(normLabel(customerHint)) || normLabel(b).includes(normLabel(customerHint))) {
          start = i + 1;
          detectedCustomer = cust || customerHint;
          sourceLabel = b || sourceLabel;
          break;
        }
      } else if (start < 0 && /DIST-|MT_|All Chain/i.test(cust + b)) {
        start = i + 1;
        detectedCustomer = cust || null;
        sourceLabel = b || sourceLabel;
      }
    }
    if (/^p&l line/i.test(a) && start < 0) {
      start = i + 1;
      if (c) detectedCustomer = c;
      if (b) sourceLabel = b;
    }
  }
  if (start < 0) start = 0;

  for (let i = start; i < rows.length; i++) {
    const row = rows[i] ?? [];
    const label = String(row[0] ?? "").trim();
    if (!label) continue;
    // Next customer block
    if (/^trivision$/i.test(label) || /^trivision\s*-/i.test(String(row[1] ?? ""))) break;
    if (/^p&l line/i.test(label)) continue;

    const amount = Number(row[1]);
    if (!Number.isFinite(amount)) continue;

    const id = LABEL_TO_ID.get(normLabel(label));
    if (!id) {
      unmatched.push(label);
      continue;
    }
    input[id] = amount;
    matched++;
  }

  return { input, matched, unmatched, detectedCustomer, sourceLabel };
}

export function buildCustomerCompare(
  data: WorkbookData,
  filters: Filters,
  input: ActualsInput,
  sourceLabel = "Input Actuals",
): CompareReport {
  const report: PLReport = computePL(data, filters);
  const byId = new Map(report.rows.map((r) => [r.id, r]));

  let inputCount = 0;
  const rows: CompareRow[] = COMPARE_LINES.map((line) => {
    if (line.spacer) {
      return {
        id: line.id,
        label: "",
        spacer: true,
        inputActual: null,
        plActual: null,
        diff: null,
        diffPct: null,
        hasInput: false,
      };
    }

    const plActual = resolvePlActual(byId, line);
    const inputActual = input[line.id] ?? null;
    const hasInput = inputActual != null && Number.isFinite(inputActual);
    if (hasInput) inputCount++;

    // Excel: Difference = TTK (P&L) − Trivision (Input)
    let diff: number | null = null;
    let diffPct: number | null = null;
    if (hasInput && plActual != null) {
      diff = plActual - inputActual!;
      if (inputActual !== 0) diffPct = (diff / Math.abs(inputActual!)) * 100;
    } else if (hasInput && plActual == null) {
      diff = 0 - inputActual!;
    }

    return {
      id: line.id,
      label: line.label,
      inputActual: hasInput ? inputActual : null,
      plActual,
      diff,
      diffPct,
      hasInput,
    };
  });

  return { filters, sourceLabel, rows, inputCount };
}

export function seedInputFromPl(data: WorkbookData, filters: Filters): ActualsInput {
  const report = computePL(data, filters);
  const byId = new Map(report.rows.map((r) => [r.id, r]));
  const input = emptyActualsInput();
  for (const line of COMPARE_LINES) {
    if (line.spacer) continue;
    input[line.id] = resolvePlActual(byId, line);
  }
  return input;
}

export function inputtableLineCount(): number {
  return COMPARE_LINES.filter((l) => !l.spacer).length;
}

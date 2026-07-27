import * as XLSX from "xlsx";
import type { Month } from "./types";

export type AllocationStep = {
  step: number;
  label: string;
  detail?: string;
  value?: number | null;
};

export type AuditGlRow = {
  grouping: string;
  brand: string;
  costCenter: string;
  description: string;
  pre: number;
  post: number;
  delta: number;
};

export type AuditPrePost = {
  grouping: string;
  pre: number;
  post: number;
  delta: number;
  glRows: number;
};

export type AuditRule = {
  grouping: string;
  ruleId: string;
  ruleName: string;
  driver: string;
  exclusions: string;
  pool: string;
};

export type AuditCalc = {
  grouping: string;
  pool: number | null;
  formula: string;
  reconstructed: number | null;
  engineActual: number | null;
  method: string;
};

export type AllocationAuditData = {
  sourceFile: string;
  customerCode: string;
  period: string; // YYYY-MM
  runId: string | null;
  methodNotes: string[];
  salesWeightNote: string | null;
  prePost: AuditPrePost[];
  rules: AuditRule[];
  calc: AuditCalc[];
  glRows: AuditGlRow[];
};

export type PreAllocationDetail = {
  available: boolean;
  reason?: string;
  sourceFile?: string;
  customerCode?: string;
  period?: string;
  grouping?: string;
  steps: AllocationStep[];
  nativeGlRows: AuditGlRow[];
  rule?: AuditRule | null;
  calc?: AuditCalc | null;
};

const MONTH_TO_NUM: Record<Month, string> = {
  Jan: "01",
  Feb: "02",
  Mar: "03",
  Apr: "04",
  May: "05",
  Jun: "06",
  Jul: "07",
  Aug: "08",
  Sep: "09",
  Oct: "10",
  Nov: "11",
  Dec: "12",
};

/** Working(Local) account / P&L label → audit accounting grouping */
const ACCOUNT_TO_GROUPING: Record<string, string> = {
  Revenue: "Revenue",
  "Distributor Margin": "Distributor Margin",
  "On Invoice Discount": "On Invoice Discount",
  "Sales Expenses": "Sales Expenses",
  "Cost of Sales - Direct": "DIRECT COGS",
  "Indirect COGS-SC": "Indirect COGS-SC",
  "Indirect COGS-MK": "Indirect COGS-MK",
  "Kobayashi Reimbursement": "Kobayashi Reimbursement",
  "Kobayashi Reimbursement (Twinpack rebate)": "Kobayashi Reimbursement",
  "Promoter & Merchandiser": "Promoter Expenses",
  Promotion: "Trade Marketing Expenses",
  Advertising: "Marketing Expenses",
  NPD: "NPD Expenses",
  "Staff Remuneration": "Staff Remuneration",
  "Finance Expenses": "Finance Expenses",
  "Administration Expenses": "Administration Expenses",
  "Outdoor Expenses": "Outdoor Expenses",
  "Other Expenses": "Other Expenses",
  "Other Income": "Other Income",
  Tax: "Tax",
  "Logistic Cost": "Transport Expenses",
  "DC Charges": "DC Charges",
  PPD: "PPD",
  "Listing Fee": "Listing Fee",
  "List Price Discount": "List Price Discounts",
};

/** Local P&L channel → audit customer code (and reverse helpers) */
const CHANNEL_TO_CODE: Record<string, string> = {
  "DIST-CHC": "DTCH",
  "DIST-CK DISTRIBUTORS": "DTCK",
  "DIST-VILLY": "DTVL",
  "DIST-LS": "DTLS",
};

function sheetRows(wb: XLSX.WorkBook, name: string): (string | number | null)[][] {
  const sheet = wb.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
    header: 1,
    defval: null,
    raw: true,
  });
}

function findSheet(wb: XLSX.WorkBook, re: RegExp): string | null {
  return wb.SheetNames.find((n) => re.test(n)) ?? null;
}

function num(v: unknown): number {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function str(v: unknown): string {
  return v == null ? "" : String(v).trim();
}

export function isAllocationAuditWorkbook(wb: XLSX.WorkBook): boolean {
  const names = wb.SheetNames.join(" | ");
  return /pre\s*vs\s*post/i.test(names) || /allocation audit/i.test(names) || /source gl/i.test(names);
}

export function parseAllocationAudit(buffer: ArrayBuffer, sourceFile: string): AllocationAuditData {
  const wb = XLSX.read(buffer, { type: "array" });
  if (!isAllocationAuditWorkbook(wb)) {
    throw new Error(
      "Not an Allocation Audit workbook (expected sheets like “3. Engine Pre vs Post”, “6. Source GL rows”).",
    );
  }

  const coverName = findSheet(wb, /cover|method/i) ?? wb.SheetNames[0]!;
  const cover = sheetRows(wb, coverName);
  let customerCode = "";
  let period = "";
  let runId: string | null = null;
  const methodNotes: string[] = [];
  for (const row of cover) {
    const k = str(row[0]).toLowerCase();
    const v = str(row[1]);
    if (k === "customer") customerCode = v.toUpperCase();
    if (k === "period") period = v.slice(0, 7);
    if (k.includes("allocation run")) runId = v || null;
    if (/^\d+\./.test(str(row[0])) && v) methodNotes.push(`${str(row[0])} ${v}`);
  }

  const prePostName = findSheet(wb, /pre\s*vs\s*post/i);
  const prePost: AuditPrePost[] = [];
  if (prePostName) {
    for (const row of sheetRows(wb, prePostName).slice(1)) {
      const grouping = str(row[0]);
      if (!grouping || /^total$/i.test(grouping)) continue;
      prePost.push({
        grouping,
        pre: num(row[1]),
        post: num(row[2]),
        delta: num(row[3]),
        glRows: num(row[4]),
      });
    }
  }

  const rulesName = findSheet(wb, /rules applied/i);
  const rules: AuditRule[] = [];
  if (rulesName) {
    for (const row of sheetRows(wb, rulesName).slice(1)) {
      const grouping = str(row[0]);
      if (!grouping || /^note/i.test(grouping)) continue;
      rules.push({
        grouping,
        ruleId: str(row[1]),
        ruleName: str(row[2]),
        driver: str(row[3]),
        exclusions: str(row[5]),
        pool: str(row[6]),
      });
    }
  }

  const calcName = findSheet(wb, /calculation detail/i);
  const calc: AuditCalc[] = [];
  let salesWeightNote: string | null = null;
  if (calcName) {
    const rows = sheetRows(wb, calcName);
    for (const row of rows) {
      const a = str(row[0]);
      if (/sales-%|sales %|weight/i.test(a) && !salesWeightNote) {
        salesWeightNote = a + (str(row[1]) ? ` ${str(row[1])}` : "");
      }
    }
    let headerSeen = false;
    for (const row of rows) {
      const a = str(row[0]);
      if (/accounting grouping/i.test(a)) {
        headerSeen = true;
        continue;
      }
      if (!headerSeen || !a || /per-brand|→ total/i.test(a)) continue;
      // stop at brand-aware subtables that reuse col layout differently
      if (/^(Indirect COGS-MK|Marketing Expenses|Trade Marketing)/i.test(a) && str(row[1]) && !str(row[2])) {
        // brand subtable rows have brand in col B without pool number pattern — still ok if pool is number
      }
      const poolRaw = row[1];
      if (typeof poolRaw === "string" && !/^-?[\d.]+$/.test(poolRaw) && poolRaw !== "") {
        // likely brand subtable
        if (!/direct GL|×|%|Σ|sales/i.test(str(row[2]))) continue;
      }
      calc.push({
        grouping: a,
        pool: row[1] == null || row[1] === "" ? null : num(row[1]),
        formula: str(row[2]),
        reconstructed: row[3] == null || row[3] === "" ? null : num(row[3]),
        engineActual: row[4] == null || row[4] === "" ? null : num(row[4]),
        method: str(row[6] ?? row[5]),
      });
    }
  }

  const glName = findSheet(wb, /source gl/i);
  const glRows: AuditGlRow[] = [];
  if (glName) {
    for (const row of sheetRows(wb, glName).slice(1)) {
      const grouping = str(row[0]);
      if (!grouping) continue;
      glRows.push({
        grouping,
        brand: str(row[1]),
        costCenter: str(row[2]),
        description: str(row[3]),
        pre: num(row[4]),
        post: num(row[5]),
        delta: num(row[6]),
      });
    }
  }

  if (!customerCode) throw new Error("Allocation Audit missing Customer on Cover sheet");
  if (!period) throw new Error("Allocation Audit missing Period on Cover sheet");

  return {
    sourceFile,
    customerCode,
    period,
    runId,
    methodNotes: methodNotes.slice(0, 8),
    salesWeightNote,
    prePost,
    rules,
    calc,
    glRows,
  };
}

export function auditPeriodForMonth(month: Month, year = 2026): string {
  return `${year}-${MONTH_TO_NUM[month]}`;
}

export function channelToAuditCode(channel: string): string | null {
  if (CHANNEL_TO_CODE[channel]) return CHANNEL_TO_CODE[channel];
  // DIST-CHC → try trailing token / DTCH style
  const m = channel.match(/^DIST-([A-Z0-9]+)/i);
  if (m) {
    const code = m[1]!.toUpperCase();
    if (code === "CK") return "DTCK";
    if (code.length <= 4) return `DT${code}`;
  }
  // already a code
  if (/^[A-Z]{3,6}$/i.test(channel)) return channel.toUpperCase();
  return null;
}

export function resolveAuditGrouping(accountOrLabel: string): string | null {
  if (ACCOUNT_TO_GROUPING[accountOrLabel]) return ACCOUNT_TO_GROUPING[accountOrLabel];
  const norm = accountOrLabel.trim().toLowerCase();
  for (const [k, v] of Object.entries(ACCOUNT_TO_GROUPING)) {
    if (k.toLowerCase() === norm) return v;
  }
  // direct match to grouping names
  return accountOrLabel;
}

/**
 * Build pre-allocation (native → allocated) steps from an uploaded audit workbook.
 */
export function buildPreAllocationDetail(
  audit: AllocationAuditData | null | undefined,
  opts: {
    channel: string;
    month: Month;
    /** Working account or P&L label */
    account: string | null;
    year?: number;
  },
): PreAllocationDetail {
  if (!audit) {
    return {
      available: false,
      reason:
        "Upload an Allocation Audit workbook (e.g. DTCH_Allocation_Audit_Workbook_2026-04.xlsx) to see Pre → Post allocation steps for this customer/month.",
      steps: [],
      nativeGlRows: [],
    };
  }

  const code = channelToAuditCode(opts.channel);
  const period = auditPeriodForMonth(opts.month, opts.year ?? 2026);
  if (!code || code !== audit.customerCode.toUpperCase()) {
    return {
      available: false,
      reason: `Loaded audit is for customer ${audit.customerCode} · ${audit.period}. Current filter is ${opts.channel}${code ? ` (${code})` : ""} · ${opts.month}. Upload a matching audit or switch Channel/Month.`,
      sourceFile: audit.sourceFile,
      customerCode: audit.customerCode,
      period: audit.period,
      steps: [],
      nativeGlRows: [],
    };
  }
  if (period !== audit.period) {
    return {
      available: false,
      reason: `Loaded audit period is ${audit.period}; selected month is ${opts.month} (${period}).`,
      sourceFile: audit.sourceFile,
      customerCode: audit.customerCode,
      period: audit.period,
      steps: [],
      nativeGlRows: [],
    };
  }

  if (!opts.account) {
    return {
      available: false,
      reason: "This P&L line is derived from other rows — open a base Working account (e.g. Indirect COGS-SC) for pre-allocation detail.",
      sourceFile: audit.sourceFile,
      steps: [],
      nativeGlRows: [],
    };
  }

  const grouping = resolveAuditGrouping(opts.account);
  if (!grouping) {
    return {
      available: false,
      reason: `No audit grouping mapped for “${opts.account}”.`,
      sourceFile: audit.sourceFile,
      steps: [],
      nativeGlRows: [],
    };
  }

  const pp = audit.prePost.find((r) => r.grouping.toLowerCase() === grouping.toLowerCase());
  const rule = audit.rules.find((r) => r.grouping.toLowerCase() === grouping.toLowerCase()) ?? null;
  const calc =
    audit.calc.find((r) => r.grouping.toLowerCase() === grouping.toLowerCase() && r.formula) ?? null;

  const glForGroup = audit.glRows.filter((r) => r.grouping.toLowerCase() === grouping.toLowerCase());
  const nativeBrandRows = glForGroup.filter(
    (r) => r.brand && r.brand !== "(all)" && Math.abs(r.pre) > 1e-9,
  );
  const postAll = glForGroup.find((r) => r.brand === "(all)");
  const nativeSum = nativeBrandRows.reduce((s, r) => s + r.pre, 0);
  const post = pp?.post ?? postAll?.post ?? null;
  const pre = pp?.pre ?? nativeSum;
  const allocatedIn = post != null ? post - pre : pp?.delta ?? null;

  if (pp == null && nativeBrandRows.length === 0 && calc == null) {
    return {
      available: false,
      reason: `Audit has no Pre/Post or GL rows for grouping “${grouping}”.`,
      sourceFile: audit.sourceFile,
      customerCode: audit.customerCode,
      period: audit.period,
      grouping,
      steps: [],
      nativeGlRows: [],
    };
  }

  const steps: AllocationStep[] = [
    {
      step: 1,
      label: "Prior to allocation — freeze GL snapshot",
      detail: `Allocation Audit ${audit.sourceFile} · customer ${audit.customerCode} · period ${audit.period}${audit.runId ? ` · run ${audit.runId}` : ""}`,
    },
  ];

  if (audit.salesWeightNote) {
    steps.push({
      step: steps.length + 1,
      label: "Sales-mix weight for this customer",
      detail: audit.salesWeightNote,
    });
  }

  if (rule) {
    steps.push({
      step: steps.length + 1,
      label: "Rule applied",
      detail: `${rule.ruleId} — ${rule.ruleName} · driver ${rule.driver || "—"} · pool ${rule.pool || "—"}${rule.exclusions ? ` · exclusions: ${rule.exclusions}` : ""}`,
    });
  }

  steps.push({
    step: steps.length + 1,
    label: "Step A — Native (already on this customer before pool fan-out)",
    detail:
      nativeBrandRows.length > 0
        ? `Sum of branded Pre GL rows = ${fmt(nativeSum)} (${nativeBrandRows.length} rows). These stay attributed to ${opts.channel} before overhead pools are spread.`
        : pre === 0
          ? "No native Pre amount on this customer (Pre = 0). Entire Post comes from pool allocation."
          : `Pre (GL) for grouping = ${fmt(pre)}`,
    value: pre,
  });

  if (calc?.formula) {
    steps.push({
      step: steps.length + 1,
      label: "Step B — Pool allocation math",
      detail: [
        calc.pool != null ? `Pool to allocate = ${fmt(calc.pool)}` : null,
        `Formula: ${calc.formula}`,
        calc.method ? `Method: ${calc.method}` : null,
        calc.reconstructed != null ? `Reconstructed = ${fmt(calc.reconstructed)}` : null,
      ]
        .filter(Boolean)
        .join(" · "),
      value: calc.reconstructed ?? calc.engineActual,
    });
  } else if (allocatedIn != null && Math.abs(allocatedIn) > 1e-6) {
    steps.push({
      step: steps.length + 1,
      label: "Step B — Allocated-in from blank / overhead pools",
      detail: `Post − Pre = ${fmt(post)} − ${fmt(pre)} = ${fmt(allocatedIn)} (share of pools assigned by sales % / brand mix)`,
      value: allocatedIn,
    });
  }

  steps.push({
    step: steps.length + 1,
    label: "Step C — Post (allocated) amount for this customer",
    detail: `Native Pre ${fmt(pre)} + allocated-in ${fmt(allocatedIn ?? 0)} → Post ${fmt(post)}`,
    value: post,
  });

  steps.push({
    step: steps.length + 1,
    label: "Then lands in Customer Brand PL Working(Local)",
    detail:
      "Working(Local) stores this Post amount (or a later run’s equivalent) under Account&Channel&Brand&Month&Year. Local P&L Actual is a VLOOKUP of that key — see steps below.",
    value: post,
  });

  return {
    available: true,
    sourceFile: audit.sourceFile,
    customerCode: audit.customerCode,
    period: audit.period,
    grouping,
    steps,
    nativeGlRows: nativeBrandRows.sort((a, b) => Math.abs(b.pre) - Math.abs(a.pre)),
    rule,
    calc,
  };
}

function fmt(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return "－";
  return v.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

/** Working account for a P&L row id (null if derived-only). */
export function workingAccountForRowId(rowId: string): string | null {
  const map: Record<string, string> = {
    gsv: "Revenue",
    "dist-margin": "Distributor Margin",
    oid: "On Invoice Discount",
    "sales-exp": "Sales Expenses",
    ppd: "PPD",
    "kby-reimb": "Kobayashi Reimbursement",
    "dc-charges": "DC Charges",
    "listing-fee": "Listing Fee",
    "list-price-disc": "List Price Discount",
    "cogs-direct": "Cost of Sales - Direct",
    "indir-mk": "Indirect COGS-MK",
    "indir-sc": "Indirect COGS-SC",
    twinpack: "Kobayashi Reimbursement (Twinpack rebate)",
    logistic: "Logistic Cost",
    promoter: "Promoter & Merchandiser",
    npd: "NPD",
    promotion: "Promotion",
    advertising: "Advertising",
    staff: "Staff Remuneration",
    finance: "Finance Expenses",
    admin: "Administration Expenses",
    outdoor: "Outdoor Expenses",
    "other-exp": "Other Expenses",
    "other-inc": "Other Income",
    temporary: "Temporary (Income)/Expense",
    tax: "Tax",
  };
  return map[rowId] ?? null;
}

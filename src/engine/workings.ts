import {
  MONTHS,
  type Filters,
  type Month,
  type MonthlyMap,
  type PLReport,
  type PLRow,
  type WorkbookData,
} from "./types";
import {
  budgetKey,
  getMonthly,
  monthIndex,
  mtdFrom,
  sellKey,
  workingKey,
  ytdFrom,
} from "./lookup";
import { computePL } from "./computePL";
import {
  buildPreAllocationDetail,
  workingAccountForRowId,
  type AllocationAuditData,
  type AllocationStep,
  type PreAllocationDetail,
} from "./allocationAudit";

export type { AllocationStep, PreAllocationDetail } from "./allocationAudit";

export type WorkingStep = {
  label: string;
  detail?: string;
  value?: number | null;
};

export type MonthlyBreak = {
  month: Month;
  value: number;
};

export type MonthlyTriple = {
  month: Month;
  ly: number;
  budget: number;
  actual: number;
  isSelected: boolean;
  inYtd: boolean;
};

export type KeyPart = {
  cell: string;
  part: string;
  value: string;
};

export type BrandBreak = {
  brand: string;
  key: string;
  value: number;
  /** rollup = TTK/Kobayashi/All Brand; leaf = named brand */
  kind?: "rollup" | "leaf";
};

export type SpreadsheetDetail = {
  /** Excel formula as written on Local P&L */
  excelFormula: string;
  /** Same formula with filter values substituted */
  excelFormulaResolved: string;
  keyParts?: KeyPart[];
  notes?: string[];
  /** When filter brand is All Brand, show brand-level Working keys that roll up */
  brandBreakdown?: BrandBreak[];
  /** Numbered allocation / lookup steps for Actuals */
  allocationSteps?: AllocationStep[];
};

export type MetricWorkings = {
  title: string;
  result: number | null;
  source?: string;
  key?: string;
  steps: WorkingStep[];
  monthly?: MonthlyBreak[];
  spreadsheet?: SpreadsheetDetail;
  /** True for Actual MTD/YTD cards — UI expands allocation trail */
  isActual?: boolean;
  /** Always-visible Actuals build-up steps */
  allocationSteps?: AllocationStep[];
};

export type RowWorkings = {
  rowId: string;
  label: string;
  filters: Filters;
  budgetChannel: string;
  budgetBrand: string;
  channelMapped: boolean;
  method: string;
  /** Month-by-month MTD values (full year) for review */
  monthlyActuals: MonthlyTriple[];
  metrics: MetricWorkings[];
  indexes: WorkingStep[];
  ratios: WorkingStep[];
  /** Pre-allocation (native → pool → post) from Allocation Audit workbook */
  preAllocation: PreAllocationDetail;
};

const LY = 2025;
const CY = 2026;

function mapBudgetChannel(data: WorkbookData, channel: string): string {
  return data.channelMap.get(channel) ?? channel;
}

function fmt(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return "－";
  return v.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

function ytdMonths(month: Month): Month[] {
  return MONTHS.slice(0, monthIndex(month) + 1) as Month[];
}

function monthlyBreaks(map: MonthlyMap, key: string, month: Month): MonthlyBreak[] {
  const arr = getMonthly(map, key);
  return ytdMonths(month).map((m) => ({
    month: m,
    value: arr[monthIndex(m)] ?? 0,
  }));
}

function workingMonthly(
  data: WorkbookData,
  account: string,
  channel: string,
  brand: string,
  year: number,
  month: Month,
  sign = 1,
): MonthlyBreak[] {
  return ytdMonths(month).map((m) => {
    const key = workingKey(account, channel, brand, m, year);
    return { month: m, value: sign * (data.working.get(key) ?? 0) };
  });
}

function rowVal(report: PLReport, id: string): PLRow | undefined {
  return report.rows.find((r) => r.id === id);
}

function pick(
  report: PLReport,
  id: string,
  period: "mtd" | "ytd",
  field: "ly" | "budget" | "actual",
): number {
  const r = rowVal(report, id);
  if (!r) return 0;
  return r.values[period][field] ?? 0;
}

type SellKind = "outQty" | "outAmt" | "inQty" | "inAmt";

function sellMaps(data: WorkbookData, kind: SellKind): { ly: MonthlyMap; cy: MonthlyMap; lySheet: string; cySheet: string; budgetLine: string } {
  if (kind === "outQty") {
    return {
      ly: data.sellOutQty2025,
      cy: data.sellOutQty2026,
      lySheet: "2025 Sell Out Qty",
      cySheet: "2026 Sell Out Qty",
      budgetLine: "Sell Out - TTK/KBY ( Vol in pcs )",
    };
  }
  if (kind === "outAmt") {
    return {
      ly: data.sellOutAmt2025,
      cy: data.sellOutAmt2026,
      lySheet: "2025 Sell Out Amount",
      cySheet: "2026 Sell Out Amount",
      budgetLine: "Sell Out - TTK/KBY ( Val in RM'000 )",
    };
  }
  if (kind === "inQty") {
    return {
      ly: data.sellInQty2025,
      cy: data.sellInQty2026,
      lySheet: "2025 Sell In Qty & Amt",
      cySheet: "2026 Sell In Qty",
      budgetLine: "Sell In - TTK/KBY ( Vol in pcs )",
    };
  }
  return {
    ly: data.sellInAmt2025,
    cy: data.sellInAmt2026,
    lySheet: "2025 Sell In Qty & Amt (Amount)",
    cySheet: "2026 Sell In Amount",
    budgetLine: "Sell In - TTK/KBY ( Val in RM'000 )",
  };
}

function workingBrandBreakdown(
  data: WorkbookData,
  account: string,
  channel: string,
  month: Month,
  year: number,
  sign: number,
  includeAllBrand = false,
): BrandBreak[] {
  const prefix = `${account}${channel}`;
  const suffix = `${month}${year}`;
  const out: BrandBreak[] = [];
  for (const [key, amount] of data.working) {
    if (!key.startsWith(prefix) || !key.endsWith(suffix)) continue;
    const brand = key.slice(prefix.length, key.length - suffix.length);
    if (!brand) continue;
    if (!includeAllBrand && brand === "All Brand") continue;
    const kind: "rollup" | "leaf" =
      brand === "All Brand" || brand === "TTK" || brand === "Kobayashi" ? "rollup" : "leaf";
    out.push({ brand, key, value: sign * amount, kind });
  }
  out.sort((a, b) => {
    const rank = (k?: string) => (k === "rollup" ? 0 : 1);
    const dr = rank(a.kind) - rank(b.kind);
    if (dr !== 0) return dr;
    return Math.abs(b.value) - Math.abs(a.value);
  });
  return out;
}

/** Build numbered Actuals trail from Working(Local) (post-allocation stored amounts). */
function buildWorkingActualTrail(
  data: WorkbookData,
  account: string,
  channel: string,
  brand: string,
  month: Month,
  year: number,
  key: string,
  stored: number,
  result: number,
  flipSign: boolean,
  isYtd: boolean,
): { allocationSteps: AllocationStep[]; brandBreakdown: BrandBreak[] } {
  const brandBreak = workingBrandBreakdown(data, account, channel, month, year, flipSign ? -1 : 1, true);
  const allBrand = brandBreak.find((b) => b.brand === "All Brand");
  const ttk = brandBreak.find((b) => b.brand === "TTK");
  const kby = brandBreak.find((b) => b.brand === "Kobayashi");
  const leaves = brandBreak.filter((b) => b.kind === "leaf");
  const leafSum = leaves.reduce((s, b) => s + b.value, 0);
  const rollupSum = (ttk?.value ?? 0) + (kby?.value ?? 0);

  const allocationSteps: AllocationStep[] = [
    {
      step: 1,
      label: "Source spreadsheet",
      detail:
        "Working(Local) Raw table — amounts are already post-allocation (native customer booking + blank-customer / pool fan-out done upstream before load).",
    },
    {
      step: 2,
      label: "Build VLOOKUP key",
      detail: `Account & Channel & Brand & Month & Year → ${key}`,
    },
    {
      step: 3,
      label: isYtd ? `Sum monthly Working keys Jan → ${month}` : "Look up Amount (column E)",
      detail: isYtd
        ? `For each month m in Jan…${month}: VLOOKUP(Account&Channel&Brand&m&${year}, Working(Local)!$A:$E, 5, 0)`
        : `=IFERROR(VLOOKUP("${key}", 'Working(Local)'!$A:$E, 5, FALSE), 0)`,
      value: stored,
    },
  ];

  if (flipSign) {
    allocationSteps.push({
      step: allocationSteps.length + 1,
      label: "Apply Local P&L sign convention",
      detail: "Revenue is stored negative in Working(Local); GSV Actual = −Revenue",
      value: result,
    });
  } else {
    allocationSteps.push({
      step: allocationSteps.length + 1,
      label: "Amount used on Local P&L Actual",
      detail: `Column K (CY ${year}) for selected Month / Channel / Brand`,
      value: result,
    });
  }

  if (brand === "All Brand" || brandBreak.length > 1) {
    allocationSteps.push({
      step: allocationSteps.length + 1,
      label: "Brand allocation view (same account / channel / month)",
      detail: [
        allBrand != null ? `All Brand stored = ${fmt(allBrand.value)}` : null,
        ttk != null || kby != null
          ? `TTK + Kobayashi = ${fmt(ttk?.value ?? 0)} + ${fmt(kby?.value ?? 0)} = ${fmt(rollupSum)}`
          : null,
        leaves.length
          ? `Sum of ${leaves.length} leaf brands = ${fmt(leafSum)} (do not add leaves on top of TTK — TTK already rolls them up)`
          : null,
        allBrand != null && Math.abs(allBrand.value - rollupSum) > 0.05
          ? `Note: All Brand (${fmt(allBrand.value)}) ≠ TTK+Kobayashi (${fmt(rollupSum)}) by ${fmt(allBrand.value - rollupSum)} — use the All Brand key for “All Brand” filter, not a re-sum of leaves.`
          : null,
      ]
        .filter(Boolean)
        .join(" · "),
      value: allBrand?.value ?? result,
    });
  }

  // Related Working accounts that often explain the same Actual (allocation components)
  const related: { account: string; label: string }[] = [];
  if (account === "Indirect COGS-SC" || account === "Indirect COGS-MK") {
    related.push(
      { account: "Indirect COGS-MK", label: "Indirect COGS-MK" },
      { account: "Indirect COGS-SC", label: "Indirect COGS-SC" },
    );
  } else if (account === "Cost of Sales - Direct") {
    related.push(
      { account: "Cost of Sales - Direct", label: "Direct" },
      { account: "Indirect COGS-MK", label: "Indirect COGS-MK" },
      { account: "Indirect COGS-SC", label: "Indirect COGS-SC" },
    );
  }

  if (related.length) {
    const parts = related.map((r) => {
      const k = workingKey(r.account, channel, brand, month, year);
      const amt = data.working.get(k) ?? 0;
      return { label: r.label, key: k, amount: amt };
    });
    const sum = parts.reduce((s, p) => s + p.amount, 0);
    allocationSteps.push({
      step: allocationSteps.length + 1,
      label: "Related Working accounts (same Channel / Brand / Month)",
      detail: parts.map((p) => `${p.label} = ${fmt(p.amount)}`).join(" · ") + ` · combined ${fmt(sum)}`,
      value: sum,
    });
  }

  allocationSteps.push({
    step: allocationSteps.length + 1,
    label: "What this file does / does not show",
    detail:
      "Customer Brand PL Working(Local) stores the final allocated Actual. Upstream allocation math (native booking + blank-customer SC/MK pools × sales %) lives in the Trivision / CombinedPL engine — compare that export in Customer comparison when you need pre-allocation steps.",
  });

  return { allocationSteps, brandBreakdown: brandBreak };
}

function sellSpreadsheet(
  sheet: string,
  key: string,
  channel: string,
  brand: string,
  ttkKby: "TTK" | "KBY",
  month: Month,
  yearColNote: string,
  isYtd: boolean,
): SpreadsheetDetail {
  const formula = isYtd
    ? `=IFERROR(SUM(INDEX('${sheet}'!$A:$P,MATCH($B$3&$B$4&"${ttkKby}",'${sheet}'!$A:$A,0), firstMonthCol):INDEX(..., monthCol)),)`
    : `=IFERROR(VLOOKUP($B$3&$B$4&"${ttkKby}",'${sheet}'!$A:$P,MATCH($B$2,{"Jan",..."Dec"},0)+offset,0),)`;
  return {
    excelFormula: formula,
    excelFormulaResolved: isYtd
      ? `=SUM of ${sheet} row key "${key}" columns Jan → ${month}`
      : `=VLOOKUP("${key}", '${sheet}', column ${month}, FALSE)`,
    keyParts: [
      { cell: "$B$3", part: "Channel", value: channel },
      { cell: "$B$4", part: "Brand", value: brand },
      { cell: "literal", part: "TTK/KBY", value: ttkKby },
      { cell: "$B$2", part: "Month", value: month },
    ],
    notes: [
      `Lookup key = Channel & Brand & "${ttkKby}" → ${key}`,
      yearColNote,
    ],
  };
}

function budgetSpreadsheet(
  budgetChannel: string,
  budgetBrand: string,
  lineLabel: string,
  month: Month,
  isYtd: boolean,
  displayChannel: string,
): SpreadsheetDetail {
  const key = budgetKey(budgetChannel, budgetBrand, lineLabel);
  return {
    excelFormula: isYtd
      ? `=IFERROR(SUM(INDEX('2026 Budget P&L'!$A:$N,MATCH($S$7&$S$8&line,'2026 Budget P&L'!$A:$A,0),2):INDEX(...,MATCH(month)+2)),0)`
      : `=IFERROR(VLOOKUP($J$7&$J$8&lineLabel,'2026 Budget P&L'!$A:$N,MATCH($B$2,{"Jan",..."Dec"},0)+2,0),0)`,
    excelFormulaResolved: isYtd
      ? `=SUM Budget row "${key}" Jan → ${month}`
      : `=VLOOKUP("${key}", '2026 Budget P&L', ${month} column, FALSE)`,
    keyParts: [
      { cell: "$J$7 / $S$7", part: "Budget Channel", value: budgetChannel },
      { cell: "$J$8 / $S$8", part: "Budget Brand", value: budgetBrand },
      { cell: "line label", part: "P&L line", value: lineLabel },
      { cell: "$B$2", part: "Month", value: month },
    ],
    notes: [
      budgetChannel !== displayChannel
        ? `Channel mapped for budget: ${displayChannel} → ${budgetChannel} (XLOOKUP Working AD→AC)`
        : "Budget channel equals display channel",
      displayChannel !== "All Chain"
        ? "Budget brand forced to All Brand when Channel ≠ All Chain"
        : "Budget brand follows selected brand",
    ],
  };
}

function workingSpreadsheet(
  account: string,
  channel: string,
  brand: string,
  month: Month,
  year: number,
  key: string,
  isYtd: boolean,
  flipSign: boolean,
  brandBreakdown?: BrandBreak[],
  allocationSteps?: AllocationStep[],
): SpreadsheetDetail {
  const yearCell = year === LY ? "I$10" : "K$10";
  const col = year === LY ? "I" : "K";
  const rollups = brandBreakdown?.filter((b) => b.kind === "rollup") ?? [];
  const leaves = brandBreakdown?.filter((b) => b.kind === "leaf") ?? [];
  return {
    excelFormula: isYtd
      ? `=IF(T$9="YTD ${month}", SUM(VLOOKUP(account&channel&brand&eachMonth&year,'Working(Local)'!$A:$G,5,0) for Jan→${month}), …)`
      : `=IFERROR(VLOOKUP($Cxx&$B$3&$B$4&$B$2&${yearCell},'Working(Local)'!$A:$E,5,0),0)`,
    excelFormulaResolved: isYtd
      ? `=SUM of Working keys ${account}${channel}${brand}Jan${year} → …${month}${year}${flipSign ? " (then × −1 for GSV)" : ""}`
      : `=IFERROR(VLOOKUP("${key}", 'Working(Local)'!$A:$E, 5, FALSE), 0)${flipSign ? " → then × −1" : ""}`,
    keyParts: [
      { cell: `$Cxx / account`, part: "Account", value: account },
      { cell: "$B$3", part: "Channel", value: channel },
      { cell: "$B$4", part: "Brand", value: brand },
      { cell: "$B$2", part: "Month", value: month },
      { cell: yearCell, part: "Year", value: String(year) },
    ],
    notes: [
      `Composite key = Account & Channel & Brand & Month & Year`,
      `VLOOKUP returns column E (Amount) from Working(Local) Raw table`,
      `Local P&L ${col} column uses year ${year}`,
      ...(flipSign ? ["Revenue is stored negative in Working; Local P&L flips sign for GSV"] : []),
      ...(rollups.length
        ? [`Rollup brands in Working: ${rollups.map((b) => `${b.brand}=${fmt(b.value)}`).join(", ")}`]
        : []),
      ...(leaves.length
        ? [`${leaves.length} leaf brand keys available below (TTK already includes its leaves)`]
        : []),
    ],
    brandBreakdown,
    allocationSteps,
  };
}

function explainSell(
  data: WorkbookData,
  filters: Filters,
  budgetChannel: string,
  budgetBrand: string,
  ttkKby: "TTK" | "KBY",
  kind: SellKind,
): MetricWorkings[] {
  const maps = sellMaps(data, kind);
  const key = sellKey(filters.channel, filters.brand, ttkKby);
  const budgetLine = maps.budgetLine.replace("TTK/KBY", ttkKby);
  const bKey = budgetKey(budgetChannel, budgetBrand, budgetLine);
  const month = filters.month;

  return [
    {
      title: "MTD · LY",
      result: mtdFrom(maps.ly, key, month),
      source: maps.lySheet,
      key,
      steps: [
        { label: "Lookup key", detail: `channel & brand & "${ttkKby}"` },
        { label: "Column", detail: month },
        { label: "Result", value: mtdFrom(maps.ly, key, month) },
      ],
      spreadsheet: sellSpreadsheet(
        maps.lySheet,
        key,
        filters.channel,
        filters.brand,
        ttkKby,
        month,
        "LY sell sheets use 2025 monthly columns",
        false,
      ),
    },
    {
      title: "MTD · Budget",
      result: mtdFrom(data.budget, bKey, month),
      source: "2026 Budget P&L",
      key: bKey,
      steps: [
        {
          label: "Budget channel",
          detail:
            budgetChannel === filters.channel
              ? budgetChannel
              : `${filters.channel} → ${budgetChannel} (Working channel map)`,
        },
        {
          label: "Budget brand",
          detail:
            filters.channel === "All Chain"
              ? filters.brand
              : `All Brand (forced when Channel ≠ All Chain)`,
        },
        { label: "Line label", detail: budgetLine },
        { label: "Column", detail: month },
        { label: "Result", value: mtdFrom(data.budget, bKey, month) },
      ],
      spreadsheet: budgetSpreadsheet(
        budgetChannel,
        budgetBrand,
        budgetLine,
        month,
        false,
        filters.channel,
      ),
    },
    {
      title: "MTD · Actual",
      result: mtdFrom(maps.cy, key, month),
      source: maps.cySheet,
      key,
      isActual: true,
      allocationSteps: [
        {
          step: 1,
          label: "Source spreadsheet",
          detail: `${maps.cySheet} — Sell Actuals (2026), not Working(Local)`,
        },
        {
          step: 2,
          label: "Build lookup key",
          detail: `Channel & Brand & "${ttkKby}" → ${key}`,
        },
        {
          step: 3,
          label: `Take ${month} column amount`,
          detail: `=VLOOKUP("${key}", '${maps.cySheet}', ${month} column, FALSE)`,
          value: mtdFrom(maps.cy, key, month),
        },
        {
          step: 4,
          label: "Local P&L Actual",
          detail: "Posted to MTD Actual for this Sell line",
          value: mtdFrom(maps.cy, key, month),
        },
      ],
      steps: [
        { label: "Lookup key", detail: key },
        { label: "Column", detail: month },
        { label: "Result", value: mtdFrom(maps.cy, key, month) },
      ],
      spreadsheet: {
        ...sellSpreadsheet(
          maps.cySheet,
          key,
          filters.channel,
          filters.brand,
          ttkKby,
          month,
          "Actual sell sheets use 2026 monthly columns",
          false,
        ),
        allocationSteps: [
          {
            step: 1,
            label: "Source spreadsheet",
            detail: maps.cySheet,
          },
          {
            step: 2,
            label: "Key",
            detail: key,
            value: mtdFrom(maps.cy, key, month),
          },
        ],
      },
    },
    {
      title: "YTD · LY",
      result: ytdFrom(maps.ly, key, month),
      source: maps.lySheet,
      key,
      steps: [
        { label: "Sum", detail: `Jan → ${month}` },
        { label: "Total", value: ytdFrom(maps.ly, key, month) },
      ],
      monthly: monthlyBreaks(maps.ly, key, month),
      spreadsheet: sellSpreadsheet(
        maps.lySheet,
        key,
        filters.channel,
        filters.brand,
        ttkKby,
        month,
        "YTD = SUM Jan→selected month on the same sell row",
        true,
      ),
    },
    {
      title: "YTD · Budget",
      result: ytdFrom(data.budget, bKey, month),
      source: "2026 Budget P&L",
      key: bKey,
      steps: [
        { label: "Sum", detail: `Jan → ${month}` },
        { label: "Total", value: ytdFrom(data.budget, bKey, month) },
      ],
      monthly: monthlyBreaks(data.budget, bKey, month),
      spreadsheet: budgetSpreadsheet(
        budgetChannel,
        budgetBrand,
        budgetLine,
        month,
        true,
        filters.channel,
      ),
    },
    {
      title: "YTD · Actual",
      result: ytdFrom(maps.cy, key, month),
      source: maps.cySheet,
      key,
      isActual: true,
      allocationSteps: [
        {
          step: 1,
          label: "Source spreadsheet",
          detail: `${maps.cySheet} — Sell Actuals (2026)`,
        },
        {
          step: 2,
          label: "Sum Jan → selected month",
          detail: `Key ${key} on columns Jan…${month}`,
          value: ytdFrom(maps.cy, key, month),
        },
      ],
      steps: [
        { label: "Sum", detail: `Jan → ${month}` },
        { label: "Total", value: ytdFrom(maps.cy, key, month) },
      ],
      monthly: monthlyBreaks(maps.cy, key, month),
      spreadsheet: sellSpreadsheet(
        maps.cySheet,
        key,
        filters.channel,
        filters.brand,
        ttkKby,
        month,
        "YTD = SUM Jan→selected month on the same sell row",
        true,
      ),
    },
  ];
}

function explainWorking(
  data: WorkbookData,
  filters: Filters,
  budgetChannel: string,
  budgetBrand: string,
  account: string,
  opts?: { budgetLabel?: string; flipSign?: boolean },
): MetricWorkings[] {
  const budgetLabel = opts?.budgetLabel ?? account;
  const sign = opts?.flipSign ? -1 : 1;
  const flipSign = !!opts?.flipSign;
  const month = filters.month;
  const { channel, brand } = filters;
  const bKey = budgetKey(budgetChannel, budgetBrand, budgetLabel);

  const lyKey = workingKey(account, channel, brand, month, LY);
  const cyKey = workingKey(account, channel, brand, month, CY);
  const lyRaw = data.working.get(lyKey) ?? 0;
  const cyRaw = data.working.get(cyKey) ?? 0;
  const brandBreakLy =
    brand === "All Brand" ? workingBrandBreakdown(data, account, channel, month, LY, sign, true) : undefined;
  const mtdActualTrail = buildWorkingActualTrail(
    data,
    account,
    channel,
    brand,
    month,
    CY,
    cyKey,
    cyRaw,
    sign * cyRaw,
    flipSign,
    false,
  );
  const ytdActualMonths = workingMonthly(data, account, channel, brand, CY, month, sign);
  const ytdActualResult = ytdActualMonths.reduce((s, x) => s + x.value, 0);
  const ytdActualTrail = buildWorkingActualTrail(
    data,
    account,
    channel,
    brand,
    month,
    CY,
    cyKey,
    ytdActualMonths.reduce((s, x) => s + (data.working.get(workingKey(account, channel, brand, x.month, CY)) ?? 0), 0),
    ytdActualResult,
    flipSign,
    true,
  );

  return [
    {
      title: "MTD · LY",
      result: sign * lyRaw,
      source: "Working(Local)",
      key: lyKey,
      steps: [
        { label: "Key pattern", detail: "account + channel + brand + month + year" },
        { label: "Stored amount", value: lyRaw },
        ...(flipSign
          ? [{ label: "Sign flip", detail: "GSV = −Revenue (Excel matches Working sign convention)", value: sign * lyRaw }]
          : []),
      ],
      spreadsheet: workingSpreadsheet(
        account,
        channel,
        brand,
        month,
        LY,
        lyKey,
        false,
        flipSign,
        brandBreakLy,
      ),
    },
    {
      title: "MTD · Budget",
      result: mtdFrom(data.budget, bKey, month),
      source: "2026 Budget P&L",
      key: bKey,
      steps: [
        { label: "Budget key", detail: bKey },
        { label: "Column", detail: month },
        { label: "Result", value: mtdFrom(data.budget, bKey, month) },
      ],
      spreadsheet: budgetSpreadsheet(budgetChannel, budgetBrand, budgetLabel, month, false, channel),
    },
    {
      title: "MTD · Actual",
      result: sign * cyRaw,
      source: "Working(Local)",
      key: cyKey,
      isActual: true,
      allocationSteps: mtdActualTrail.allocationSteps,
      steps: mtdActualTrail.allocationSteps.map((s) => ({
        label: `${s.step}. ${s.label}`,
        detail: s.detail,
        value: s.value,
      })),
      spreadsheet: workingSpreadsheet(
        account,
        channel,
        brand,
        month,
        CY,
        cyKey,
        false,
        flipSign,
        mtdActualTrail.brandBreakdown,
        mtdActualTrail.allocationSteps,
      ),
    },
    {
      title: "YTD · LY",
      result: workingMonthly(data, account, channel, brand, LY, month, sign).reduce((s, x) => s + x.value, 0),
      source: "Working(Local)",
      steps: [
        { label: "Sum of monthly keys", detail: `Jan${LY} → ${month}${LY}` },
      ],
      monthly: workingMonthly(data, account, channel, brand, LY, month, sign),
      spreadsheet: workingSpreadsheet(account, channel, brand, month, LY, lyKey, true, flipSign),
    },
    {
      title: "YTD · Budget",
      result: ytdFrom(data.budget, bKey, month),
      source: "2026 Budget P&L",
      key: bKey,
      steps: [{ label: "Sum", detail: `Jan → ${month}` }],
      monthly: monthlyBreaks(data.budget, bKey, month),
      spreadsheet: budgetSpreadsheet(budgetChannel, budgetBrand, budgetLabel, month, true, channel),
    },
    {
      title: "YTD · Actual",
      result: ytdActualResult,
      source: "Working(Local)",
      isActual: true,
      allocationSteps: ytdActualTrail.allocationSteps,
      steps: [
        ...ytdActualTrail.allocationSteps.map((s) => ({
          label: `${s.step}. ${s.label}`,
          detail: s.detail,
          value: s.value,
        })),
        { label: "Monthly Actuals in YTD", detail: `Jan${CY} → ${month}${CY}` },
      ],
      monthly: ytdActualMonths,
      spreadsheet: workingSpreadsheet(
        account,
        channel,
        brand,
        month,
        CY,
        cyKey,
        true,
        flipSign,
        ytdActualTrail.brandBreakdown,
        ytdActualTrail.allocationSteps,
      ),
    },
  ];
}

function formulaMetrics(
  report: PLReport,
  expression: string,
  parts: { id: string; label: string; op: "+" | "-" }[],
  extras?: WorkingStep[],
  excelHint?: string,
): MetricWorkings[] {
  const build = (period: "mtd" | "ytd", field: "ly" | "budget" | "actual"): MetricWorkings => {
    const steps: WorkingStep[] = parts.map((p) => ({
      label: `${p.op === "-" ? "−" : "+"} ${p.label}`,
      detail: p.id,
      value: pick(report, p.id, period, field),
    }));
    if (extras) steps.push(...extras);
    let result = 0;
    for (const p of parts) {
      const v = pick(report, p.id, period, field);
      result += p.op === "-" ? -v : v;
    }
    const resolved = parts
      .map((p, i) => {
        const v = pick(report, p.id, period, field);
        const sign = p.op === "-" ? "−" : i === 0 ? "" : "+";
        return `${sign}${fmt(v)}`;
      })
      .join(" ");

    const isActual = field === "actual";
    const allocationSteps: AllocationStep[] | undefined = isActual
      ? [
          {
            step: 1,
            label: "Derived Actual (not a single Working key)",
            detail: `Local P&L computes this Actual from component rows: ${expression}`,
          },
          ...parts.map((p, i) => ({
            step: i + 2,
            label: `${p.op === "-" ? "Subtract" : "Add"} ${p.label}`,
            detail: `Use ${period.toUpperCase()} Actual of row “${p.label}” (see that line’s Workings for Working(Local) allocation)`,
            value: pick(report, p.id, period, field),
          })),
          {
            step: parts.length + 2,
            label: "Resulting Actual",
            detail: `=${resolved}`,
            value: result,
          },
        ]
      : undefined;

    return {
      title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
      result,
      isActual,
      allocationSteps,
      steps: [
        { label: "Formula", detail: expression },
        ...steps,
        { label: "Result", value: result },
      ],
      spreadsheet: {
        excelFormula: excelHint ?? `=${expression}`,
        excelFormulaResolved: `=${resolved} = ${fmt(result)}`,
        notes: [
          "Derived on Local P&L from other rows (not a direct Working/Budget lookup)",
          `Period: ${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
          ...(isActual
            ? ["Open each component line’s Workings to see Working(Local) key / brand allocation for that Actual."]
            : []),
        ],
        keyParts: parts.map((p) => ({
          cell: p.id,
          part: p.label,
          value: fmt(pick(report, p.id, period, field)),
        })),
        allocationSteps,
      },
    };
  };

  return [
    build("mtd", "ly"),
    build("mtd", "budget"),
    build("mtd", "actual"),
    build("ytd", "ly"),
    build("ytd", "budget"),
    build("ytd", "actual"),
  ];
}

/** Recompute MTD LY/Budget/Actual for every calendar month under the same Channel/Brand. */
function buildMonthlyActuals(
  data: WorkbookData,
  filters: Filters,
  rowId: string,
): MonthlyTriple[] {
  const selectedIdx = monthIndex(filters.month);
  return MONTHS.map((month, idx) => {
    const report = computePL(data, { ...filters, month });
    const row = report.rows.find((r) => r.id === rowId);
    return {
      month,
      ly: row?.values.mtd.ly ?? 0,
      budget: row?.values.mtd.budget ?? 0,
      actual: row?.values.mtd.actual ?? 0,
      isSelected: month === filters.month,
      inYtd: idx <= selectedIdx,
    };
  });
}

function indexAndRatioSteps(row: PLRow): { indexes: WorkingStep[]; ratios: WorkingStep[] } {
  const { values: v } = row;
  return {
    indexes: [
      {
        label: "MTD Index vs Budget",
        detail: `Actual ÷ Budget × 100 = ${fmt(v.mtd.actual)} ÷ ${fmt(v.mtd.budget)} × 100`,
        value: v.mtdIndexBudget,
      },
      {
        label: "MTD Index vs LY",
        detail: `Actual ÷ LY × 100 = ${fmt(v.mtd.actual)} ÷ ${fmt(v.mtd.ly)} × 100`,
        value: v.mtdIndexLy,
      },
      {
        label: "YTD Index vs Budget",
        detail: `Actual ÷ Budget × 100 = ${fmt(v.ytd.actual)} ÷ ${fmt(v.ytd.budget)} × 100`,
        value: v.ytdIndexBudget,
      },
      {
        label: "YTD Index vs LY",
        detail: `Actual ÷ LY × 100 = ${fmt(v.ytd.actual)} ÷ ${fmt(v.ytd.ly)} × 100`,
        value: v.ytdIndexLy,
      },
    ],
    ratios: [
      {
        label: "MTD % to base (LY / Budget / Actual)",
        detail: `${fmt(v.mtdRatio.ly)} / ${fmt(v.mtdRatio.budget)} / ${fmt(v.mtdRatio.actual)} (shown as % of GSV or group total)`,
      },
      {
        label: "YTD % to base (LY / Budget / Actual)",
        detail: `${fmt(v.ytdRatio.ly)} / ${fmt(v.ytdRatio.budget)} / ${fmt(v.ytdRatio.actual)}`,
      },
    ],
  };
}

export function buildRowWorkings(
  data: WorkbookData,
  filters: Filters,
  report: PLReport,
  rowId: string,
  audit?: AllocationAuditData | null,
): RowWorkings | null {
  const row = rowVal(report, rowId);
  if (!row) return null;

  const budgetChannel = mapBudgetChannel(data, filters.channel);
  const budgetBrand = filters.channel === "All Chain" ? filters.brand : "All Brand";
  const channelMapped = budgetChannel !== filters.channel;
  const workingAccount = workingAccountForRowId(rowId);
  const preAllocation = buildPreAllocationDetail(audit, {
    channel: filters.channel,
    month: filters.month,
    account: workingAccount,
  });

  let method = "";
  let metrics: MetricWorkings[] = [];

  const sellSpec: Record<string, { ttk: "TTK" | "KBY"; kind: SellKind }> = {
    "so-ttk-qty": { ttk: "TTK", kind: "outQty" },
    "so-kby-qty": { ttk: "KBY", kind: "outQty" },
    "so-ttk-amt": { ttk: "TTK", kind: "outAmt" },
    "so-kby-amt": { ttk: "KBY", kind: "outAmt" },
    "si-ttk-qty": { ttk: "TTK", kind: "inQty" },
    "si-kby-qty": { ttk: "KBY", kind: "inQty" },
    "si-ttk-amt": { ttk: "TTK", kind: "inAmt" },
    "si-kby-amt": { ttk: "KBY", kind: "inAmt" },
  };

  const workingSpec: Record<string, { account: string; budgetLabel?: string; flipSign?: boolean }> = {
    gsv: {
      account: "Revenue",
      budgetLabel: "Sell in (Gross Sales Value - GSV) - Total ",
      flipSign: true,
    },
    "dist-margin": { account: "Distributor Margin" },
    oid: { account: "On Invoice Discount" },
    "sales-exp": { account: "Sales Expenses" },
    ppd: { account: "PPD" },
    "kby-reimb": { account: "Kobayashi Reimbursement" },
    "dc-charges": { account: "DC Charges" },
    "listing-fee": { account: "Listing Fee" },
    "list-price-disc": { account: "List Price Discount" },
    "cogs-direct": { account: "Cost of Sales - Direct" },
    "indir-mk": { account: "Indirect COGS-MK" },
    "indir-sc": { account: "Indirect COGS-SC" },
    twinpack: { account: "Kobayashi Reimbursement (Twinpack rebate)" },
    logistic: { account: "Logistic Cost" },
    promoter: { account: "Promoter & Merchandiser" },
    npd: { account: "NPD" },
    promotion: { account: "Promotion" },
    advertising: { account: "Advertising" },
    staff: { account: "Staff Remuneration", budgetLabel: "G&A Expenses" },
    finance: { account: "Finance Expenses" },
    admin: { account: "Administration Expenses" },
    outdoor: { account: "Outdoor Expenses" },
    "other-exp": { account: "Other Expenses" },
    "other-inc": { account: "Other Income" },
    temporary: { account: "Temporary (Income)/Expense" },
    tax: { account: "Tax" },
  };

  if (sellSpec[rowId]) {
    const s = sellSpec[rowId]!;
    method = `VLOOKUP sell sheet by key channel+brand+${s.ttk}; Budget by mapped channel + brand + line label.`;
    metrics = explainSell(data, filters, budgetChannel, budgetBrand, s.ttk, s.kind);
  } else if (workingSpec[rowId]) {
    const w = workingSpec[rowId]!;
    method = w.flipSign
      ? "Working(Local) Revenue lookup with sign flip for GSV; Budget from Budget P&L GSV line."
      : "Working(Local) composite-key lookup for LY/Actual; Budget from 2026 Budget P&L.";
    metrics = explainWorking(data, filters, budgetChannel, budgetBrand, w.account, w);
  } else if (rowId === "so-tot-qty" || rowId === "so-tot-amt" || rowId === "si-tot-qty" || rowId === "si-tot-amt") {
    const a = rowId.replace("tot", "ttk");
    const b = rowId.replace("tot", "kby");
    method = "Sum of TTK + KBY components.";
    metrics = formulaMetrics(report, "TTK + KBY", [
      { id: a, label: rowVal(report, a)?.label ?? a, op: "+" },
      { id: b, label: rowVal(report, b)?.label ?? b, op: "+" },
    ]);
  } else if (rowId === "aup-ttk" || rowId === "aup-kby") {
    const prefix = rowId === "aup-ttk" ? "so-ttk" : "so-kby";
    method = "Average Unit Price = Sell Out Value ÷ Sell Out Volume.";
    metrics = formulaMetrics(report, "Sell Out Val ÷ Sell Out Qty", [
      { id: `${prefix}-amt`, label: "Sell Out Value", op: "+" },
    ]);
    // richer custom for AUP
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        const amt = pick(report, `${prefix}-amt`, period, field);
        const qty = pick(report, `${prefix}-qty`, period, field);
        const result = qty === 0 ? null : amt / qty;
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
          result,
          steps: [
            { label: "Value", value: amt },
            { label: "Volume", value: qty },
            { label: "AUP = Value ÷ Volume", value: result },
          ],
        };
      }),
    );
  } else if (rowId === "trade-spend") {
    method =
      "LY/Actual = sum of Trade Spend children. Budget = direct Budget P&L lookup for “Trade Spend”.";
    const children = [
      "sales-exp",
      "ppd",
      "kby-reimb",
      "dc-charges",
      "listing-fee",
      "list-price-disc",
    ];
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        if (field === "budget") {
          const bKey = budgetKey(budgetChannel, budgetBrand, "Trade Spend");
          const result =
            period === "mtd"
              ? mtdFrom(data.budget, bKey, filters.month)
              : ytdFrom(data.budget, bKey, filters.month);
          return {
            title: `${period.toUpperCase()} · Budget`,
            result,
            source: "2026 Budget P&L",
            key: bKey,
            steps: [
              { label: "Direct budget lookup", detail: "Trade Spend" },
              { label: "Result", value: result },
            ],
            monthly: period === "ytd" ? monthlyBreaks(data.budget, bKey, filters.month) : undefined,
          };
        }
        const steps = children.map((id) => ({
          label: `+ ${rowVal(report, id)?.label ?? id}`,
          value: pick(report, id, period, field),
        }));
        const result = children.reduce((s, id) => s + pick(report, id, period, field), 0);
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : "Actual"}`,
          result,
          steps: [...steps, { label: "Sum", value: result }],
        };
      }),
    );
  } else if (rowId === "net-sales") {
    method =
      "LY/Actual: GSV − Dist Margin − OID − Trade Spend. Budget also subtracts List Price Discount (Excel J42).";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        const gsv = pick(report, "gsv", period, field);
        const dist = pick(report, "dist-margin", period, field);
        const oid = pick(report, "oid", period, field);
        const trade = pick(report, "trade-spend", period, field);
        const lpd = pick(report, "list-price-disc", period, field);
        const steps: WorkingStep[] = [
          { label: "GSV", value: gsv },
          { label: "− Distributor Margin", value: dist },
          { label: "− On Invoice Discount", value: oid },
          { label: "− Trade Spend", value: trade },
        ];
        let result = gsv - dist - oid - trade;
        if (field === "budget") {
          steps.push({ label: "− List Price Discount (budget only)", value: lpd });
          result -= lpd;
        }
        steps.push({ label: "Net Sales", value: result });
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
          result,
          steps,
        };
      }),
    );
  } else if (rowId === "cogs-indirect") {
    method = "LY/Actual = Indirect MK + SC + Twinpack. Budget = direct “Cost of Sales - Indirect”.";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        if (field === "budget") {
          const v = pick(report, "cogs-indirect", period, "budget");
          return {
            title: `${period.toUpperCase()} · Budget`,
            result: v,
            steps: [
              { label: "Direct budget lookup", detail: "Cost of Sales - Indirect" },
              { label: "Result", value: v },
            ],
          };
        }
        const mk = pick(report, "indir-mk", period, field);
        const sc = pick(report, "indir-sc", period, field);
        const tp = pick(report, "twinpack", period, field);
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : "Actual"}`,
          result: mk + sc + tp,
          steps: [
            { label: "+ Indirect COGS-MK", value: mk },
            { label: "+ Indirect COGS-SC", value: sc },
            { label: "+ Twinpack rebate", value: tp },
            { label: "Sum", value: mk + sc + tp },
          ],
        };
      }),
    );
  } else if (rowId === "cogs-logistic") {
    method =
      "LY/Actual: Direct + Indirect + Logistic. Budget: Direct + Indirect + Logistic + Twinpack (Excel adds Twinpack again on budget).";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        const d = pick(report, "cogs-direct", period, field);
        const i = pick(report, "cogs-indirect", period, field);
        const l = pick(report, "logistic", period, field);
        const t = pick(report, "twinpack", period, field);
        const steps: WorkingStep[] = [
          { label: "+ Cost of Sales - Direct", value: d },
          { label: "+ Cost of Sales - Indirect", value: i },
          { label: "+ Logistic Cost", value: l },
        ];
        let result = d + i + l;
        if (field === "budget") {
          steps.push({ label: "+ Twinpack (budget only, Excel J49)", value: t });
          result += t;
        }
        steps.push({ label: "COGS & Logistic", value: result });
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
          result,
          steps,
        };
      }),
    );
  } else if (rowId === "gross-margin") {
    method = "Net Sales − COGS & Logistic.";
    metrics = formulaMetrics(report, "Net Sales − COGS & Logistic", [
      { id: "net-sales", label: "Net Sales", op: "+" },
      { id: "cogs-logistic", label: "COGS & Logistic", op: "-" },
    ]);
  } else if (rowId === "mkt-others") {
    method = "LY/Actual = Promoter + NPD. Budget = direct “Marketing Cost - Others”.";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        if (field === "budget") {
          const v = pick(report, "mkt-others", period, "budget");
          return {
            title: `${period.toUpperCase()} · Budget`,
            result: v,
            steps: [{ label: "Direct budget lookup", detail: "Marketing Cost - Others", value: v }],
          };
        }
        const a = pick(report, "promoter", period, field);
        const b = pick(report, "npd", period, field);
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : "Actual"}`,
          result: a + b,
          steps: [
            { label: "+ Promoter & Merchandiser", value: a },
            { label: "+ NPD", value: b },
            { label: "Sum", value: a + b },
          ],
        };
      }),
    );
  } else if (rowId === "mkt-cost") {
    method = "LY/Actual = Promotion + Advertising. Budget = direct “Marketing Cost”.";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        if (field === "budget") {
          const v = pick(report, "mkt-cost", period, "budget");
          return {
            title: `${period.toUpperCase()} · Budget`,
            result: v,
            steps: [{ label: "Direct budget lookup", detail: "Marketing Cost", value: v }],
          };
        }
        const a = pick(report, "promotion", period, field);
        const b = pick(report, "advertising", period, field);
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : "Actual"}`,
          result: a + b,
          steps: [
            { label: "+ Promotion", value: a },
            { label: "+ Advertising", value: b },
            { label: "Sum", value: a + b },
          ],
        };
      }),
    );
  } else if (rowId === "gp-after-mkt") {
    method = "Gross Margin − Marketing Cost - Others − Marketing Cost.";
    metrics = formulaMetrics(report, "GM − Mkt Others − Mkt Cost", [
      { id: "gross-margin", label: "Gross Margin", op: "+" },
      { id: "mkt-others", label: "Marketing Cost - Others", op: "-" },
      { id: "mkt-cost", label: "Marketing Cost", op: "-" },
    ]);
  } else if (rowId === "ga") {
    method = "LY/Actual = Staff Remuneration. Budget = direct “G&A Expenses”.";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        const v = pick(report, "ga", period, field);
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
          result: v,
          steps:
            field === "budget"
              ? [{ label: "Budget lookup", detail: "G&A Expenses", value: v }]
              : [{ label: "= Staff Remuneration", value: pick(report, "staff", period, field) }],
        };
      }),
    );
  } else if (rowId === "other-ga") {
    method = "LY/Actual = Finance + Admin + Outdoor. Budget = direct “Other G&A Cost”.";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        if (field === "budget") {
          const v = pick(report, "other-ga", period, "budget");
          return {
            title: `${period.toUpperCase()} · Budget`,
            result: v,
            steps: [{ label: "Direct budget lookup", detail: "Other G&A Cost", value: v }],
          };
        }
        const a = pick(report, "finance", period, field);
        const b = pick(report, "admin", period, field);
        const c = pick(report, "outdoor", period, field);
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : "Actual"}`,
          result: a + b + c,
          steps: [
            { label: "+ Finance Expenses", value: a },
            { label: "+ Administration Expenses", value: b },
            { label: "+ Outdoor Expenses", value: c },
            { label: "Sum", value: a + b + c },
          ],
        };
      }),
    );
  } else if (rowId === "other-op") {
    method =
      "LY/Actual = Other Exp + Other Inc + Temporary. Budget = Other Operating lookup + Temporary budget (Excel J64).";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        if (field === "budget") {
          const opKey = budgetKey(budgetChannel, budgetBrand, "Other Operating (Income)/Expense");
          const tmpKey = budgetKey(budgetChannel, budgetBrand, "Temporary (Income)/Expense");
          const op =
            period === "mtd"
              ? mtdFrom(data.budget, opKey, filters.month)
              : ytdFrom(data.budget, opKey, filters.month);
          const tmp =
            period === "mtd"
              ? mtdFrom(data.budget, tmpKey, filters.month)
              : ytdFrom(data.budget, tmpKey, filters.month);
          return {
            title: `${period.toUpperCase()} · Budget`,
            result: op + tmp,
            steps: [
              { label: "Other Operating (budget)", detail: opKey, value: op },
              { label: "+ Temporary (budget)", detail: tmpKey, value: tmp },
              { label: "Sum (Excel J64)", value: op + tmp },
            ],
          };
        }
        const a = pick(report, "other-exp", period, field);
        const b = pick(report, "other-inc", period, field);
        const c = pick(report, "temporary", period, field);
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : "Actual"}`,
          result: a + b + c,
          steps: [
            { label: "+ Other Expenses", value: a },
            { label: "+ Other Income", value: b },
            { label: "+ Temporary", value: c },
            { label: "Sum", value: a + b + c },
          ],
        };
      }),
    );
  } else if (rowId === "op-income") {
    method = "GP after Marketing − G&A − Other G&A − Other Operating.";
    metrics = formulaMetrics(report, "GP after Mkt − G&A − Other G&A − Other Op", [
      { id: "gp-after-mkt", label: "Gross Profit after Marketing", op: "+" },
      { id: "ga", label: "G&A Expenses", op: "-" },
      { id: "other-ga", label: "Other G&A Cost", op: "-" },
      { id: "other-op", label: "Other Operating", op: "-" },
    ]);
  } else if (rowId === "nop") {
    method = "Operating Income before Tax − Tax.";
    metrics = formulaMetrics(report, "Op Income − Tax", [
      { id: "op-income", label: "Operating Income before Tax", op: "+" },
      { id: "tax", label: "Tax", op: "-" },
    ]);
  } else if (rowId === "dist-vs-so") {
    method = "Distributor Margin ÷ Sell Out Total Value.";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        const num_ = pick(report, "dist-margin", period, field);
        const den = pick(report, "so-tot-amt", period, field);
        const result = den === 0 ? null : num_ / den;
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
          result,
          steps: [
            { label: "Distributor Margin", value: num_ },
            { label: "Sell Out Total Value", value: den },
            { label: "Ratio", value: result },
          ],
        };
      }),
    );
  } else if (rowId === "oid-trade-vs-so") {
    method = "(Trade Spend + OID) ÷ Sell Out Total Value.";
    metrics = (["mtd", "ytd"] as const).flatMap((period) =>
      (["ly", "budget", "actual"] as const).map((field) => {
        const trade = pick(report, "trade-spend", period, field);
        const oid = pick(report, "oid", period, field);
        const den = pick(report, "so-tot-amt", period, field);
        const result = den === 0 ? null : (trade + oid) / den;
        return {
          title: `${period.toUpperCase()} · ${field === "ly" ? "LY" : field === "budget" ? "Budget" : "Actual"}`,
          result,
          steps: [
            { label: "Trade Spend", value: trade },
            { label: "+ OID", value: oid },
            { label: "Sell Out Total Value", value: den },
            { label: "Ratio", value: result },
          ],
        };
      }),
    );
  } else {
    method = "See calculated row values; detailed source mapping not classified.";
    metrics = [
      {
        title: "MTD values",
        result: row.values.mtd.actual,
        steps: [
          { label: "LY", value: row.values.mtd.ly },
          { label: "Budget", value: row.values.mtd.budget },
          { label: "Actual", value: row.values.mtd.actual },
        ],
      },
      {
        title: "YTD values",
        result: row.values.ytd.actual,
        steps: [
          { label: "LY", value: row.values.ytd.ly },
          { label: "Budget", value: row.values.ytd.budget },
          { label: "Actual", value: row.values.ytd.actual },
        ],
      },
    ];
  }

  const { indexes, ratios } = indexAndRatioSteps(row);
  const monthlyActuals = buildMonthlyActuals(data, filters, rowId);

  // Ensure every Actual card exposes an allocation / build-up trail for the UI
  const decoratedMetrics = metrics.map((m) => {
    if (!/actual/i.test(m.title)) return m;

    let postSteps: AllocationStep[] = m.allocationSteps?.length
      ? m.allocationSteps
      : [
          {
            step: 1,
            label: "Actual build-up",
            detail: m.source
              ? `Source sheet: ${m.source}${m.key ? ` · key ${m.key}` : ""}`
              : "Derived from component P&L Actuals / Working(Local) lookups",
            value: m.result,
          },
          ...m.steps.map((s, i) => ({
            step: i + 2,
            label: s.label,
            detail: s.detail,
            value: s.value,
          })),
        ];

    // Prepend engine Pre → Post steps when Allocation Audit matches this customer/month/account
    let allocationSteps = postSteps;
    if (preAllocation.available && preAllocation.steps.length && /MTD/i.test(m.title)) {
      const pre = preAllocation.steps;
      const offset = pre.length;
      allocationSteps = [
        ...pre,
        {
          step: offset + 1,
          label: "── After allocation: Customer Brand PL lookup ──",
          detail: "Working(Local) / Sell sheets hold the Post amount used on Local P&L Actual",
        },
        ...postSteps.map((s) => ({ ...s, step: offset + 1 + s.step })),
      ];
    }

    return {
      ...m,
      isActual: true,
      allocationSteps,
      spreadsheet: m.spreadsheet
        ? { ...m.spreadsheet, allocationSteps }
        : {
            excelFormula: "(see steps)",
            excelFormulaResolved: `Actual = ${fmt(m.result)}`,
            allocationSteps,
          },
    };
  });

  return {
    rowId,
    label: row.label,
    filters,
    budgetChannel,
    budgetBrand,
    channelMapped,
    method,
    monthlyActuals,
    metrics: decoratedMetrics,
    indexes,
    ratios,
    preAllocation,
  };
}

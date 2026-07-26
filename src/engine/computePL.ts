import {
  type Filters,
  type MetricTriple,
  type PLReport,
  type PLRow,
  type PLRowValues,
  type WorkbookData,
  MONTHS,
  type Month,
} from "./types";
import {
  budgetKey,
  indexPct,
  mtdFrom,
  safeDiv,
  sellKey,
  workingKey,
  ytdFrom,
} from "./lookup";

const LY_YEAR = 2025;
const CY_YEAR = 2026;

function emptyTriple(): MetricTriple {
  return { ly: 0, budget: 0, actual: 0 };
}

function withIndexes(v: PLRowValues): PLRowValues {
  return {
    ...v,
    mtdIndexBudget: indexPct(v.mtd.actual, v.mtd.budget),
    mtdIndexLy: indexPct(v.mtd.actual, v.mtd.ly),
    ytdIndexBudget: indexPct(v.ytd.actual, v.ytd.budget),
    ytdIndexLy: indexPct(v.ytd.actual, v.ytd.ly),
  };
}

function addTriples(a: MetricTriple, b: MetricTriple): MetricTriple {
  return {
    ly: (a.ly ?? 0) + (b.ly ?? 0),
    budget: (a.budget ?? 0) + (b.budget ?? 0),
    actual: (a.actual ?? 0) + (b.actual ?? 0),
  };
}

function subTriples(a: MetricTriple, b: MetricTriple): MetricTriple {
  return {
    ly: (a.ly ?? 0) - (b.ly ?? 0),
    budget: (a.budget ?? 0) - (b.budget ?? 0),
    actual: (a.actual ?? 0) - (b.actual ?? 0),
  };
}

function makeValues(mtd: MetricTriple, ytd: MetricTriple): PLRowValues {
  return withIndexes({
    mtd,
    ytd,
    mtdIndexBudget: null,
    mtdIndexLy: null,
    ytdIndexBudget: null,
    ytdIndexLy: null,
    mtdRatio: { ly: null, budget: null, actual: null },
    ytdRatio: { ly: null, budget: null, actual: null },
  });
}

function ratioTo(gsv: PLRowValues, row: PLRowValues): PLRowValues {
  return {
    ...row,
    mtdRatio: {
      ly: safeDiv(row.mtd.ly, gsv.mtd.ly),
      budget: safeDiv(row.mtd.budget, gsv.mtd.budget),
      actual: safeDiv(row.mtd.actual, gsv.mtd.actual),
    },
    ytdRatio: {
      ly: safeDiv(row.ytd.ly, gsv.ytd.ly),
      budget: safeDiv(row.ytd.budget, gsv.ytd.budget),
      actual: safeDiv(row.ytd.actual, gsv.ytd.actual),
    },
  };
}

function shareOf(part: PLRowValues, total: PLRowValues): PLRowValues {
  return {
    ...part,
    mtdRatio: {
      ly: safeDiv(part.mtd.ly, total.mtd.ly),
      budget: safeDiv(part.mtd.budget, total.mtd.budget),
      actual: safeDiv(part.mtd.actual, total.mtd.actual),
    },
    ytdRatio: {
      ly: safeDiv(part.ytd.ly, total.ytd.ly),
      budget: safeDiv(part.ytd.budget, total.ytd.budget),
      actual: safeDiv(part.ytd.actual, total.ytd.actual),
    },
  };
}

type Ctx = {
  data: WorkbookData;
  filters: Filters;
  budgetChannel: string;
  budgetBrand: string;
};

function mapBudgetChannel(data: WorkbookData, channel: string): string {
  return data.channelMap.get(channel) ?? channel;
}

function workingMtd(ctx: Ctx, account: string, year: number): number {
  const { channel, brand, month } = ctx.filters;
  const key = workingKey(account, channel, brand, month, year);
  return ctx.data.working.get(key) ?? 0;
}

function workingYtd(ctx: Ctx, account: string, year: number): number {
  const { channel, brand, month } = ctx.filters;
  const end = MONTHS.indexOf(month);
  let sum = 0;
  for (let i = 0; i <= end; i++) {
    const m = MONTHS[i] as Month;
    const key = workingKey(account, channel, brand, m, year);
    sum += ctx.data.working.get(key) ?? 0;
  }
  return sum;
}

/** Working P&L amount (as stored). Revenue is negated by caller to get GSV. */
function workingTriple(ctx: Ctx, account: string, flipSign = false): { mtd: MetricTriple; ytd: MetricTriple } {
  const sign = flipSign ? -1 : 1;
  return {
    mtd: {
      ly: sign * workingMtd(ctx, account, LY_YEAR),
      budget: 0, // filled separately
      actual: sign * workingMtd(ctx, account, CY_YEAR),
    },
    ytd: {
      ly: sign * workingYtd(ctx, account, LY_YEAR),
      budget: 0,
      actual: sign * workingYtd(ctx, account, CY_YEAR),
    },
  };
}

function budgetMtd(ctx: Ctx, lineLabel: string): number {
  const key = budgetKey(ctx.budgetChannel, ctx.budgetBrand, lineLabel);
  return mtdFrom(ctx.data.budget, key, ctx.filters.month);
}

function budgetYtd(ctx: Ctx, lineLabel: string): number {
  const key = budgetKey(ctx.budgetChannel, ctx.budgetBrand, lineLabel);
  return ytdFrom(ctx.data.budget, key, ctx.filters.month);
}

function sellTriple(
  ctx: Ctx,
  ttkKby: "TTK" | "KBY",
  kind: "outQty" | "outAmt" | "inQty" | "inAmt",
): { mtd: MetricTriple; ytd: MetricTriple } {
  const { channel, brand, month } = ctx.filters;
  const key = sellKey(channel, brand, ttkKby);
  const budgetLine =
    kind === "outQty"
      ? `Sell Out - ${ttkKby} ( Vol in pcs )`
      : kind === "outAmt"
        ? `Sell Out - ${ttkKby} ( Val in RM'000 )`
        : kind === "inQty"
          ? `Sell In - ${ttkKby} ( Vol in pcs )`
          : `Sell In - ${ttkKby} ( Val in RM'000 )`;

  const lyMap =
    kind === "outQty"
      ? ctx.data.sellOutQty2025
      : kind === "outAmt"
        ? ctx.data.sellOutAmt2025
        : kind === "inQty"
          ? ctx.data.sellInQty2025
          : ctx.data.sellInAmt2025;
  const cyMap =
    kind === "outQty"
      ? ctx.data.sellOutQty2026
      : kind === "outAmt"
        ? ctx.data.sellOutAmt2026
        : kind === "inQty"
          ? ctx.data.sellInQty2026
          : ctx.data.sellInAmt2026;

  return {
    mtd: {
      ly: mtdFrom(lyMap, key, month),
      budget: budgetMtd(ctx, budgetLine),
      actual: mtdFrom(cyMap, key, month),
    },
    ytd: {
      ly: ytdFrom(lyMap, key, month),
      budget: budgetYtd(ctx, budgetLine),
      actual: ytdFrom(cyMap, key, month),
    },
  };
}

function workingLine(ctx: Ctx, account: string, budgetLabel?: string): PLRowValues {
  const w = workingTriple(ctx, account, false);
  const label = budgetLabel ?? account;
  w.mtd.budget = budgetMtd(ctx, label);
  w.ytd.budget = budgetYtd(ctx, label);
  return makeValues(w.mtd, w.ytd);
}

function row(
  id: string,
  label: string,
  indent: number,
  values: PLRowValues,
  opts?: { bold?: boolean; isAup?: boolean },
): PLRow {
  return { id, label, indent, values, monthlyActuals: [], ...opts };
}

/** Attach Jan–Dec MTD Actual columns for the same Channel/Brand. */
export function withMonthlyActuals(data: WorkbookData, filters: Filters): PLReport {
  const base = computePL(data, filters);
  const byMonth = MONTHS.map((month) => computePL(data, { ...filters, month }));
  return {
    ...base,
    rows: base.rows.map((r) => ({
      ...r,
      monthlyActuals: byMonth.map((rep) => {
        const match = rep.rows.find((x) => x.id === r.id);
        return match?.values.mtd.actual ?? 0;
      }),
    })),
  };
}

export function computePL(data: WorkbookData, filters: Filters): PLReport {
  const budgetChannel = mapBudgetChannel(data, filters.channel);
  const budgetBrand = filters.channel === "All Chain" ? filters.brand : "All Brand";
  const ctx: Ctx = { data, filters, budgetChannel, budgetBrand };

  // --- Sell Out ---
  const soTtkQty = makeValues(...tuple(sellTriple(ctx, "TTK", "outQty")));
  const soKbyQty = makeValues(...tuple(sellTriple(ctx, "KBY", "outQty")));
  const soTotQty = makeValues(addTriples(soTtkQty.mtd, soKbyQty.mtd), addTriples(soTtkQty.ytd, soKbyQty.ytd));

  const soTtkAmt = makeValues(...tuple(sellTriple(ctx, "TTK", "outAmt")));
  const soKbyAmt = makeValues(...tuple(sellTriple(ctx, "KBY", "outAmt")));
  const soTotAmt = makeValues(addTriples(soTtkAmt.mtd, soKbyAmt.mtd), addTriples(soTtkAmt.ytd, soKbyAmt.ytd));

  // AUP
  const aupTtk = makeValues(
    {
      ly: safeDiv(soTtkAmt.mtd.ly, soTtkQty.mtd.ly) ?? 0,
      budget: safeDiv(soTtkAmt.mtd.budget, soTtkQty.mtd.budget) ?? 0,
      actual: safeDiv(soTtkAmt.mtd.actual, soTtkQty.mtd.actual) ?? 0,
    },
    {
      ly: safeDiv(soTtkAmt.ytd.ly, soTtkQty.ytd.ly) ?? 0,
      budget: safeDiv(soTtkAmt.ytd.budget, soTtkQty.ytd.budget) ?? 0,
      actual: safeDiv(soTtkAmt.ytd.actual, soTtkQty.ytd.actual) ?? 0,
    },
  );
  const aupKby = makeValues(
    {
      ly: safeDiv(soKbyAmt.mtd.ly, soKbyQty.mtd.ly) ?? 0,
      budget: safeDiv(soKbyAmt.mtd.budget, soKbyQty.mtd.budget) ?? 0,
      actual: safeDiv(soKbyAmt.mtd.actual, soKbyQty.mtd.actual) ?? 0,
    },
    {
      ly: safeDiv(soKbyAmt.ytd.ly, soKbyQty.ytd.ly) ?? 0,
      budget: safeDiv(soKbyAmt.ytd.budget, soKbyQty.ytd.budget) ?? 0,
      actual: safeDiv(soKbyAmt.ytd.actual, soKbyQty.ytd.actual) ?? 0,
    },
  );

  // --- Sell In ---
  const siTtkQty = makeValues(...tuple(sellTriple(ctx, "TTK", "inQty")));
  const siKbyQty = makeValues(...tuple(sellTriple(ctx, "KBY", "inQty")));
  const siTotQty = makeValues(addTriples(siTtkQty.mtd, siKbyQty.mtd), addTriples(siTtkQty.ytd, siKbyQty.ytd));

  const siTtkAmt = makeValues(...tuple(sellTriple(ctx, "TTK", "inAmt")));
  const siKbyAmt = makeValues(...tuple(sellTriple(ctx, "KBY", "inAmt")));
  const siTotAmt = makeValues(addTriples(siTtkAmt.mtd, siKbyAmt.mtd), addTriples(siTtkAmt.ytd, siKbyAmt.ytd));

  // GSV from Working Revenue (sign flip) + budget line
  const rev = workingTriple(ctx, "Revenue", true);
  const gsvLabel = "Sell in (Gross Sales Value - GSV) - Total ";
  rev.mtd.budget = budgetMtd(ctx, gsvLabel);
  rev.ytd.budget = budgetYtd(ctx, gsvLabel);
  const gsv = makeValues(rev.mtd, rev.ytd);

  const distMargin = workingLine(ctx, "Distributor Margin");
  const oid = workingLine(ctx, "On Invoice Discount");

  const salesExp = workingLine(ctx, "Sales Expenses");
  const ppd = workingLine(ctx, "PPD");
  const kbyReimb = workingLine(ctx, "Kobayashi Reimbursement");
  const dcCharges = workingLine(ctx, "DC Charges");
  const listingFee = workingLine(ctx, "Listing Fee");
  const listPriceDisc = workingLine(ctx, "List Price Discount");

  // Trade Spend actual = sum children; budget = direct lookup
  const tradeSpendChildren = [salesExp, ppd, kbyReimb, dcCharges, listingFee, listPriceDisc];
  const tradeSpendActualMtd = tradeSpendChildren.reduce((a, r) => addTriples(a, r.mtd), emptyTriple());
  const tradeSpendActualYtd = tradeSpendChildren.reduce((a, r) => addTriples(a, r.ytd), emptyTriple());
  const tradeSpend = makeValues(
    {
      ly: tradeSpendActualMtd.ly,
      budget: budgetMtd(ctx, "Trade Spend"),
      actual: tradeSpendActualMtd.actual,
    },
    {
      ly: tradeSpendActualYtd.ly,
      budget: budgetYtd(ctx, "Trade Spend"),
      actual: tradeSpendActualYtd.actual,
    },
  );

  // Net Sales: actual GSV - dist - oid - trade; budget also subtracts List Price Discount
  const netSales = makeValues(
    {
      ly: (gsv.mtd.ly ?? 0) - (distMargin.mtd.ly ?? 0) - (oid.mtd.ly ?? 0) - (tradeSpend.mtd.ly ?? 0),
      budget:
        (gsv.mtd.budget ?? 0) -
        (distMargin.mtd.budget ?? 0) -
        (oid.mtd.budget ?? 0) -
        (tradeSpend.mtd.budget ?? 0) -
        (listPriceDisc.mtd.budget ?? 0),
      actual:
        (gsv.mtd.actual ?? 0) - (distMargin.mtd.actual ?? 0) - (oid.mtd.actual ?? 0) - (tradeSpend.mtd.actual ?? 0),
    },
    {
      ly: (gsv.ytd.ly ?? 0) - (distMargin.ytd.ly ?? 0) - (oid.ytd.ly ?? 0) - (tradeSpend.ytd.ly ?? 0),
      budget:
        (gsv.ytd.budget ?? 0) -
        (distMargin.ytd.budget ?? 0) -
        (oid.ytd.budget ?? 0) -
        (tradeSpend.ytd.budget ?? 0) -
        (listPriceDisc.ytd.budget ?? 0),
      actual:
        (gsv.ytd.actual ?? 0) - (distMargin.ytd.actual ?? 0) - (oid.ytd.actual ?? 0) - (tradeSpend.ytd.actual ?? 0),
    },
  );

  const cogsDirect = workingLine(ctx, "Cost of Sales - Direct");
  const indirMk = workingLine(ctx, "Indirect COGS-MK");
  const indirSc = workingLine(ctx, "Indirect COGS-SC");
  const twinpack = workingLine(ctx, "Kobayashi Reimbursement (Twinpack rebate)");
  const cogsIndirect = makeValues(
    {
      ly: (indirMk.mtd.ly ?? 0) + (indirSc.mtd.ly ?? 0) + (twinpack.mtd.ly ?? 0),
      budget: budgetMtd(ctx, "Cost of Sales - Indirect"),
      actual: (indirMk.mtd.actual ?? 0) + (indirSc.mtd.actual ?? 0) + (twinpack.mtd.actual ?? 0),
    },
    {
      ly: (indirMk.ytd.ly ?? 0) + (indirSc.ytd.ly ?? 0) + (twinpack.ytd.ly ?? 0),
      budget: budgetYtd(ctx, "Cost of Sales - Indirect"),
      actual: (indirMk.ytd.actual ?? 0) + (indirSc.ytd.actual ?? 0) + (twinpack.ytd.actual ?? 0),
    },
  );
  const logistic = workingLine(ctx, "Logistic Cost");

  // COGS & Logistic: actual Direct+Indirect+Logistic; budget + twinpack separately (Excel J49)
  const cogsLogistic = makeValues(
    {
      ly: (cogsDirect.mtd.ly ?? 0) + (cogsIndirect.mtd.ly ?? 0) + (logistic.mtd.ly ?? 0),
      budget:
        (cogsDirect.mtd.budget ?? 0) +
        (cogsIndirect.mtd.budget ?? 0) +
        (logistic.mtd.budget ?? 0) +
        (twinpack.mtd.budget ?? 0),
      actual: (cogsDirect.mtd.actual ?? 0) + (cogsIndirect.mtd.actual ?? 0) + (logistic.mtd.actual ?? 0),
    },
    {
      ly: (cogsDirect.ytd.ly ?? 0) + (cogsIndirect.ytd.ly ?? 0) + (logistic.ytd.ly ?? 0),
      budget:
        (cogsDirect.ytd.budget ?? 0) +
        (cogsIndirect.ytd.budget ?? 0) +
        (logistic.ytd.budget ?? 0) +
        (twinpack.ytd.budget ?? 0),
      actual: (cogsDirect.ytd.actual ?? 0) + (cogsIndirect.ytd.actual ?? 0) + (logistic.ytd.actual ?? 0),
    },
  );

  const grossMargin = makeValues(subTriples(netSales.mtd, cogsLogistic.mtd), subTriples(netSales.ytd, cogsLogistic.ytd));

  const promoter = workingLine(ctx, "Promoter & Merchandiser");
  const npd = workingLine(ctx, "NPD");
  const mktOthers = makeValues(addTriples(promoter.mtd, npd.mtd), addTriples(promoter.ytd, npd.ytd));
  // budget for Marketing Cost - Others
  mktOthers.mtd.budget = budgetMtd(ctx, "Marketing Cost - Others");
  mktOthers.ytd.budget = budgetYtd(ctx, "Marketing Cost - Others");
  const mktOthersFinal = withIndexes(mktOthers);

  const promotion = workingLine(ctx, "Promotion");
  const advertising = workingLine(ctx, "Advertising");
  const mktCost = makeValues(addTriples(promotion.mtd, advertising.mtd), addTriples(promotion.ytd, advertising.ytd));
  mktCost.mtd.budget = budgetMtd(ctx, "Marketing Cost");
  mktCost.ytd.budget = budgetYtd(ctx, "Marketing Cost");
  const mktCostFinal = withIndexes(mktCost);

  const gpAfterMkt = makeValues(
    {
      ly: (grossMargin.mtd.ly ?? 0) - (mktOthersFinal.mtd.ly ?? 0) - (mktCostFinal.mtd.ly ?? 0),
      budget: (grossMargin.mtd.budget ?? 0) - (mktOthersFinal.mtd.budget ?? 0) - (mktCostFinal.mtd.budget ?? 0),
      actual: (grossMargin.mtd.actual ?? 0) - (mktOthersFinal.mtd.actual ?? 0) - (mktCostFinal.mtd.actual ?? 0),
    },
    {
      ly: (grossMargin.ytd.ly ?? 0) - (mktOthersFinal.ytd.ly ?? 0) - (mktCostFinal.ytd.ly ?? 0),
      budget: (grossMargin.ytd.budget ?? 0) - (mktOthersFinal.ytd.budget ?? 0) - (mktCostFinal.ytd.budget ?? 0),
      actual: (grossMargin.ytd.actual ?? 0) - (mktOthersFinal.ytd.actual ?? 0) - (mktCostFinal.ytd.actual ?? 0),
    },
  );

  const staff = workingLine(ctx, "Staff Remuneration", "G&A Expenses");
  // G&A actual = staff; budget = G&A Expenses lookup (staff budget mirrors G&A in Excel)
  const ga = makeValues(
    {
      ly: staff.mtd.ly,
      budget: budgetMtd(ctx, "G&A Expenses"),
      actual: staff.mtd.actual,
    },
    {
      ly: staff.ytd.ly,
      budget: budgetYtd(ctx, "G&A Expenses"),
      actual: staff.ytd.actual,
    },
  );
  // Staff remuneration budget = G&A budget
  const staffFinal = makeValues(
    { ly: staff.mtd.ly, budget: ga.mtd.budget, actual: staff.mtd.actual },
    { ly: staff.ytd.ly, budget: ga.ytd.budget, actual: staff.ytd.actual },
  );

  const finance = workingLine(ctx, "Finance Expenses");
  const admin = workingLine(ctx, "Administration Expenses");
  const outdoor = workingLine(ctx, "Outdoor Expenses");
  const otherGa = makeValues(
    {
      ly: (finance.mtd.ly ?? 0) + (admin.mtd.ly ?? 0) + (outdoor.mtd.ly ?? 0),
      budget: budgetMtd(ctx, "Other G&A Cost"),
      actual: (finance.mtd.actual ?? 0) + (admin.mtd.actual ?? 0) + (outdoor.mtd.actual ?? 0),
    },
    {
      ly: (finance.ytd.ly ?? 0) + (admin.ytd.ly ?? 0) + (outdoor.ytd.ly ?? 0),
      budget: budgetYtd(ctx, "Other G&A Cost"),
      actual: (finance.ytd.actual ?? 0) + (admin.ytd.actual ?? 0) + (outdoor.ytd.actual ?? 0),
    },
  );

  const otherExp = workingLine(ctx, "Other Expenses");
  const otherInc = workingLine(ctx, "Other Income");
  const temporary = workingLine(ctx, "Temporary (Income)/Expense");
  // Budget: Excel J64 = VLOOKUP(Other Operating) + Temporary budget (not sum of all children)
  const otherOp = makeValues(
    {
      ly: (otherExp.mtd.ly ?? 0) + (otherInc.mtd.ly ?? 0) + (temporary.mtd.ly ?? 0),
      budget:
        budgetMtd(ctx, "Other Operating (Income)/Expense") +
        budgetMtd(ctx, "Temporary (Income)/Expense"),
      actual: (otherExp.mtd.actual ?? 0) + (otherInc.mtd.actual ?? 0) + (temporary.mtd.actual ?? 0),
    },
    {
      ly: (otherExp.ytd.ly ?? 0) + (otherInc.ytd.ly ?? 0) + (temporary.ytd.ly ?? 0),
      budget:
        budgetYtd(ctx, "Other Operating (Income)/Expense") +
        budgetYtd(ctx, "Temporary (Income)/Expense"),
      actual: (otherExp.ytd.actual ?? 0) + (otherInc.ytd.actual ?? 0) + (temporary.ytd.actual ?? 0),
    },
  );

  const opIncome = makeValues(
    {
      ly: (gpAfterMkt.mtd.ly ?? 0) - (ga.mtd.ly ?? 0) - (otherGa.mtd.ly ?? 0) - (otherOp.mtd.ly ?? 0),
      budget: (gpAfterMkt.mtd.budget ?? 0) - (ga.mtd.budget ?? 0) - (otherGa.mtd.budget ?? 0) - (otherOp.mtd.budget ?? 0),
      actual:
        (gpAfterMkt.mtd.actual ?? 0) - (ga.mtd.actual ?? 0) - (otherGa.mtd.actual ?? 0) - (otherOp.mtd.actual ?? 0),
    },
    {
      ly: (gpAfterMkt.ytd.ly ?? 0) - (ga.ytd.ly ?? 0) - (otherGa.ytd.ly ?? 0) - (otherOp.ytd.ly ?? 0),
      budget: (gpAfterMkt.ytd.budget ?? 0) - (ga.ytd.budget ?? 0) - (otherGa.ytd.budget ?? 0) - (otherOp.ytd.budget ?? 0),
      actual:
        (gpAfterMkt.ytd.actual ?? 0) - (ga.ytd.actual ?? 0) - (otherGa.ytd.actual ?? 0) - (otherOp.ytd.actual ?? 0),
    },
  );

  const tax = workingLine(ctx, "Tax");
  const nop = makeValues(subTriples(opIncome.mtd, tax.mtd), subTriples(opIncome.ytd, tax.ytd));

  // Disclosure ratios
  const distVsSellOut = makeValues(
    {
      ly: safeDiv(distMargin.mtd.ly, soTotAmt.mtd.ly),
      budget: safeDiv(distMargin.mtd.budget, soTotAmt.mtd.budget),
      actual: safeDiv(distMargin.mtd.actual, soTotAmt.mtd.actual),
    },
    {
      ly: safeDiv(distMargin.ytd.ly, soTotAmt.ytd.ly),
      budget: safeDiv(distMargin.ytd.budget, soTotAmt.ytd.budget),
      actual: safeDiv(distMargin.ytd.actual, soTotAmt.ytd.actual),
    },
  );

  const oidTradeVsSellOut = makeValues(
    {
      ly: safeDiv((tradeSpend.mtd.ly ?? 0) + (oid.mtd.ly ?? 0), soTotAmt.mtd.ly),
      budget: safeDiv((tradeSpend.mtd.budget ?? 0) + (oid.mtd.budget ?? 0), soTotAmt.mtd.budget),
      actual: safeDiv((tradeSpend.mtd.actual ?? 0) + (oid.mtd.actual ?? 0), soTotAmt.mtd.actual),
    },
    {
      ly: safeDiv((tradeSpend.ytd.ly ?? 0) + (oid.ytd.ly ?? 0), soTotAmt.ytd.ly),
      budget: safeDiv((tradeSpend.ytd.budget ?? 0) + (oid.ytd.budget ?? 0), soTotAmt.ytd.budget),
      actual: safeDiv((tradeSpend.ytd.actual ?? 0) + (oid.ytd.actual ?? 0), soTotAmt.ytd.actual),
    },
  );

  const applyGsv = (v: PLRowValues) => ratioTo(gsv, v);

  const rows: PLRow[] = [
    row("aup-ttk", "Average Unit Price (AUP) - TTK", 1, aupTtk, { isAup: true }),
    row("aup-kby", "Average Unit Price (AUP) - KBY", 1, aupKby, { isAup: true }),
    row("so-ttk-qty", "Sell Out - TTK ( Vol in pcs )", 1, shareOf(soTtkQty, soTotQty)),
    row("so-kby-qty", "Sell Out - KBY ( Vol in pcs )", 1, shareOf(soKbyQty, soTotQty)),
    row("so-tot-qty", "Sell Out - Total  ( Vol in pcs )", 0, withIndexes(soTotQty), { bold: true }),
    row("so-ttk-amt", "Sell Out - TTK ( Val in RM'000 )", 1, shareOf(soTtkAmt, soTotAmt)),
    row("so-kby-amt", "Sell Out - KBY ( Val in RM'000 )", 1, shareOf(soKbyAmt, soTotAmt)),
    row("so-tot-amt", "Sell Out - Total ( Val in RM'000 )", 0, withIndexes(soTotAmt), { bold: true }),
    row("si-ttk-qty", "Sell In - TTK ( Vol in pcs )", 1, shareOf(siTtkQty, siTotQty)),
    row("si-kby-qty", "Sell In - KBY ( Vol in pcs )", 1, shareOf(siKbyQty, siTotQty)),
    row("si-tot-qty", "Sell in - Total ( Vol in pcs )", 0, withIndexes(siTotQty), { bold: true }),
    row("si-ttk-amt", "Sell In - TTK ( Val in RM'000 )", 1, shareOf(siTtkAmt, siTotAmt)),
    row("si-kby-amt", "Sell In - KBY ( Val in RM'000 )", 1, shareOf(siKbyAmt, siTotAmt)),
    row("si-tot-amt", "Sell in - Total ( Val in RM'000 )", 0, withIndexes(siTotAmt), { bold: true }),
    row("gsv", "Sell in (Gross Sales Value - GSV) - Total ", 0, applyGsv(gsv), { bold: true }),
    row("dist-margin", "Distributor Margin", 0, applyGsv(distMargin), { bold: true }),
    row("dist-vs-so", "Distributor Margin vs Sell Out (Disclosure)", 1, distVsSellOut, {
      isAup: true,
    }),
    row("oid", "On Invoice Discount", 0, applyGsv(oid), { bold: true }),
    row("trade-spend", "Trade Spend", 0, applyGsv(tradeSpend), { bold: true }),
    row("sales-exp", "Sales Expenses", 1, applyGsv(salesExp)),
    row("ppd", "PPD", 1, applyGsv(ppd)),
    row("kby-reimb", "Kobayashi Reimbursement", 1, applyGsv(kbyReimb)),
    row("dc-charges", "DC Charges", 1, applyGsv(dcCharges)),
    row("listing-fee", "Listing Fee", 1, applyGsv(listingFee)),
    row("list-price-disc", "List Price Discount", 1, applyGsv(listPriceDisc)),
    row("oid-trade-vs-so", "OID & Trade Spend  vs Sell Out (Disclosure)", 1, oidTradeVsSellOut, {
      isAup: true,
    }),
    row("net-sales", "Net Sales", 0, applyGsv(netSales), { bold: true }),
    row("cogs-direct", "Cost of Sales - Direct", 1, applyGsv(cogsDirect)),
    row("cogs-indirect", "Cost of Sales - Indirect", 1, applyGsv(cogsIndirect)),
    row("indir-mk", "Indirect COGS-MK", 1, applyGsv(indirMk)),
    row("indir-sc", "Indirect COGS-SC", 1, applyGsv(indirSc)),
    row("twinpack", "Kobayashi Reimbursement (Twinpack rebate)", 1, applyGsv(twinpack)),
    row("logistic", "Logistic Cost", 1, applyGsv(logistic)),
    row("cogs-logistic", "COGS & Logistic", 0, applyGsv(cogsLogistic), { bold: true }),
    row("gross-margin", "Gross Margin", 0, applyGsv(grossMargin), { bold: true }),
    row("promoter", "Promoter & Merchandiser", 1, applyGsv(promoter)),
    row("npd", "NPD", 1, applyGsv(npd)),
    row("mkt-others", "Marketing Cost - Others", 0, applyGsv(mktOthersFinal), { bold: true }),
    row("promotion", "Promotion", 1, applyGsv(promotion)),
    row("advertising", "Advertising", 1, applyGsv(advertising)),
    row("mkt-cost", "Marketing Cost", 0, applyGsv(mktCostFinal), { bold: true }),
    row("gp-after-mkt", "Gross Profit after Marketing", 0, applyGsv(gpAfterMkt), { bold: true }),
    row("ga", "G&A Expenses", 0, applyGsv(ga), { bold: true }),
    row("staff", "Staff Remuneration", 1, applyGsv(staffFinal)),
    row("other-ga", "Other G&A Cost", 0, applyGsv(otherGa), { bold: true }),
    row("finance", "Finance Expenses", 1, applyGsv(finance)),
    row("admin", "Administration Expenses", 1, applyGsv(admin)),
    row("outdoor", "Outdoor Expenses", 1, applyGsv(outdoor)),
    row("other-op", "Other Operating (Income)/Expense", 0, applyGsv(otherOp), { bold: true }),
    row("other-exp", "Other Expenses", 1, applyGsv(otherExp)),
    row("other-inc", "Other Income", 1, applyGsv(otherInc)),
    row("temporary", "Temporary (Income)/Expense", 1, applyGsv(temporary)),
    row("op-income", "Operating Income before Tax", 0, applyGsv(opIncome), { bold: true }),
    row("tax", "Tax", 0, applyGsv(tax), { bold: true }),
    row("nop", "Net Operating Profit", 0, applyGsv(nop), { bold: true }),
  ];

  return { filters, rows };
}

function tuple<T extends { mtd: MetricTriple; ytd: MetricTriple }>(
  v: T,
): [MetricTriple, MetricTriple] {
  return [v.mtd, v.ytd];
}

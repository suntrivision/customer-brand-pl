import * as XLSX from "xlsx";
import { MONTHS, type Month, type MonthlyMap, type WorkbookData } from "./types";
import { num, zeros } from "./lookup";

const REQUIRED_SHEETS = [
  "Working(Local)",
  "2026 Budget P&L",
  "2025 Sell Out Qty",
  "2026 Sell Out Qty",
  "2025 Sell Out Amount",
  "2026 Sell Out Amount",
  "2025 Sell In Qty & Amt",
  "2026 Sell In Qty",
  "2026 Sell In Amount",
] as const;

function sheetToRows(wb: XLSX.WorkBook, name: string): unknown[][] {
  const sheet = wb.Sheets[name];
  if (!sheet) throw new Error(`Missing required sheet: ${name}`);
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: null,
    raw: true,
  });
}

function ensureMonthly(map: MonthlyMap, key: string): number[] {
  let arr = map.get(key);
  if (!arr) {
    arr = zeros();
    map.set(key, arr);
  }
  return arr;
}

function parseMonthlyBlock(
  rows: unknown[][],
  opts: {
    startRow: number; // 0-based
    keyCol: number;
    janCol: number;
    channelCol?: number;
    brandCol?: number;
  },
  channels: Set<string>,
  brands: Set<string>,
): MonthlyMap {
  const map: MonthlyMap = new Map();
  const { startRow, keyCol, janCol, channelCol, brandCol } = opts;

  for (let r = startRow; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const key = row[keyCol];
    if (key == null || key === "") continue;
    const keyStr = String(key);
    // skip header-like keys
    if (keyStr === "Automated" || keyStr === "Automation" || keyStr === "A" || keyStr === "Qty") {
      continue;
    }
    const arr = ensureMonthly(map, keyStr);
    for (let m = 0; m < 12; m++) {
      arr[m] = num(row[janCol + m]);
    }
    if (channelCol != null && row[channelCol] != null) {
      channels.add(String(row[channelCol]));
    }
    if (brandCol != null && row[brandCol] != null) {
      brands.add(String(row[brandCol]));
    }
  }
  return map;
}

function parseWorking(rows: unknown[][]): {
  working: Map<string, number>;
  channelMap: Map<string, string>;
  channels: Set<string>;
  brands: Set<string>;
} {
  const working = new Map<string, number>();
  const channelMap = new Map<string, string>();
  const channels = new Set<string>();
  const brands = new Set<string>();

  // Channel map: AC = budget name (col 28), AD = P&L name (col 29), 0-based
  for (let r = 0; r < Math.min(rows.length, 5000); r++) {
    const row = rows[r];
    if (!row) continue;
    const budgetName = row[28];
    const plName = row[29];
    if (
      plName != null &&
      budgetName != null &&
      String(plName) !== "Name from P&L" &&
      String(budgetName) !== "Name from Budget" &&
      String(budgetName) !== "For 2026 Budget"
    ) {
      channelMap.set(String(plName), String(budgetName));
    }
  }

  // Flat lookup table: find header "Raw" then data rows
  let start = -1;
  for (let r = 0; r < rows.length; r++) {
    const a = rows[r]?.[0];
    if (a === "Raw") {
      start = r + 1;
      break;
    }
  }
  if (start < 0) {
    // fallback: first row where col A looks like a composite key with a year
    for (let r = 0; r < rows.length; r++) {
      const a = rows[r]?.[0];
      if (typeof a === "string" && /20\d{2}$/.test(a) && a.length > 15) {
        start = r;
        break;
      }
    }
  }
  if (start < 0) throw new Error("Could not find Working(Local) lookup table");

  for (let r = start; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const key = row[0];
    if (key == null || key === "") continue;
    const keyStr = String(key);
    if (keyStr === "Raw") continue;
    working.set(keyStr, num(row[4]));
    const ch = row[2];
    const br = row[3];
    if (ch != null && ch !== "") channels.add(String(ch));
    if (br != null && br !== "") brands.add(String(br));
  }

  return { working, channelMap, channels, brands };
}

function parseBudget(rows: unknown[][]): MonthlyMap {
  const map: MonthlyMap = new Map();
  // Row 0 headers, row 1 dates, data from row 2. Key col 0, Jan col 2
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const key = row[0];
    if (key == null || key === "") continue;
    const arr = ensureMonthly(map, String(key));
    for (let m = 0; m < 12; m++) {
      arr[m] = num(row[2 + m]);
    }
  }
  return map;
}

export function parseWorkbook(buffer: ArrayBuffer): WorkbookData {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });

  const missing = REQUIRED_SHEETS.filter((s) => !wb.SheetNames.includes(s));
  if (missing.length) {
    throw new Error(`Missing sheets: ${missing.join(", ")}`);
  }

  const channels = new Set<string>(["All Chain"]);
  const brands = new Set<string>(["All Brand"]);

  const workingRows = sheetToRows(wb, "Working(Local)");
  const { working, channelMap, channels: wCh, brands: wBr } = parseWorking(workingRows);
  wCh.forEach((c) => channels.add(c));
  wBr.forEach((b) => brands.add(b));

  const budget = parseBudget(sheetToRows(wb, "2026 Budget P&L"));

  // 2025 Sell Out Qty: key A, channel D(3), brand E(4), Jan col G(6)
  const sellOutQty2025 = parseMonthlyBlock(
    sheetToRows(wb, "2025 Sell Out Qty"),
    { startRow: 1, keyCol: 0, janCol: 6, channelCol: 3, brandCol: 4 },
    channels,
    brands,
  );

  // 2026 Sell Out Qty: key A, channel B(1), brand C(2), Jan col E(4)
  const sellOutQty2026 = parseMonthlyBlock(
    sheetToRows(wb, "2026 Sell Out Qty"),
    { startRow: 1, keyCol: 0, janCol: 4, channelCol: 1, brandCol: 2 },
    channels,
    brands,
  );

  // 2025 Sell Out Amount: key A, channel B, brand D(3), Jan col F(5)
  const sellOutAmt2025 = parseMonthlyBlock(
    sheetToRows(wb, "2025 Sell Out Amount"),
    { startRow: 1, keyCol: 0, janCol: 5, channelCol: 1, brandCol: 3 },
    channels,
    brands,
  );

  // 2026 Sell Out Amount: key A, channel B, brand C, Jan col E(4)
  const sellOutAmt2026 = parseMonthlyBlock(
    sheetToRows(wb, "2026 Sell Out Amount"),
    { startRow: 1, keyCol: 0, janCol: 4, channelCol: 1, brandCol: 2 },
    channels,
    brands,
  );

  // 2025 Sell In: header on row 2 (index 1), data from row 3 (index 2)
  // Qty: key A, Jan F(5); Amount: key U(20), Jan Z(25)
  const sellInRows = sheetToRows(wb, "2025 Sell In Qty & Amt");
  const sellInQty2025 = parseMonthlyBlock(
    sellInRows,
    { startRow: 2, keyCol: 0, janCol: 5, channelCol: 2, brandCol: 4 },
    channels,
    brands,
  );
  const sellInAmt2025 = parseMonthlyBlock(
    sellInRows,
    { startRow: 2, keyCol: 20, janCol: 25, channelCol: 22, brandCol: 24 },
    channels,
    brands,
  );

  const sellInQty2026 = parseMonthlyBlock(
    sheetToRows(wb, "2026 Sell In Qty"),
    { startRow: 1, keyCol: 0, janCol: 4, channelCol: 1, brandCol: 2 },
    channels,
    brands,
  );

  const sellInAmt2026 = parseMonthlyBlock(
    sheetToRows(wb, "2026 Sell In Amount"),
    { startRow: 1, keyCol: 0, janCol: 4, channelCol: 1, brandCol: 2 },
    channels,
    brands,
  );

  const prefer = (list: string[], first: string[]) => {
    const rest = list.filter((x) => !first.includes(x)).sort((a, b) => a.localeCompare(b));
    return [...first.filter((x) => list.includes(x) || x === "All Chain" || x === "All Brand"), ...rest];
  };

  return {
    working,
    channelMap,
    budget,
    sellOutQty2025,
    sellOutQty2026,
    sellOutAmt2025,
    sellOutAmt2026,
    sellInQty2025,
    sellInAmt2025,
    sellInQty2026,
    sellInAmt2026,
    channels: prefer([...channels], ["All Chain"]),
    brands: prefer([...brands], ["All Brand"]),
  };
}

export function isMonth(v: string): v is Month {
  return (MONTHS as readonly string[]).includes(v);
}

export { REQUIRED_SHEETS };

import { MONTHS, type Month, type MonthlyMap } from "./types";

export function monthIndex(month: Month): number {
  return MONTHS.indexOf(month);
}

export function num(v: unknown): number {
  if (v == null || v === "") return 0;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

export function getMonthly(map: MonthlyMap, key: string): number[] {
  return map.get(key) ?? zeros();
}

export function zeros(): number[] {
  return [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
}

export function mtdFrom(map: MonthlyMap, key: string, month: Month): number {
  return getMonthly(map, key)[monthIndex(month)] ?? 0;
}

export function ytdFrom(map: MonthlyMap, key: string, month: Month): number {
  const arr = getMonthly(map, key);
  const end = monthIndex(month);
  let sum = 0;
  for (let i = 0; i <= end; i++) sum += arr[i] ?? 0;
  return sum;
}

export function workingKey(
  account: string,
  channel: string,
  brand: string,
  month: Month,
  year: number,
): string {
  return `${account}${channel}${brand}${month}${year}`;
}

export function sellKey(channel: string, brand: string, ttkKby: "TTK" | "KBY"): string {
  return `${channel}${brand}${ttkKby}`;
}

export function budgetKey(channel: string, brand: string, lineLabel: string): string {
  return `${channel}${brand}${lineLabel}`;
}

/** Index vs Budget / LY — Excel returns "－" when invalid; we use null */
export function indexPct(actual: number | null, base: number | null): number | null {
  if (actual == null || base == null) return null;
  if (base <= 0 || actual <= 0) return null;
  if (!Number.isFinite(actual / base)) return null;
  return (actual / base) * 100;
}

export function safeDiv(num_: number | null, den: number | null): number | null {
  if (num_ == null || den == null || den === 0) return null;
  const r = num_ / den;
  return Number.isFinite(r) ? r : null;
}

export function nOrNull(v: number): number | null {
  return Number.isFinite(v) ? v : null;
}

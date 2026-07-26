export function fmtNumber(v: number | null | undefined, decimals = 0): string {
  if (v == null || !Number.isFinite(v)) return "－";
  return v.toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function fmtIndex(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return "－";
  return v.toLocaleString("en-US", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  });
}

export function fmtRatio(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return " -";
  return (v * 100).toLocaleString("en-US", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  }) + "%";
}

export function fmtAup(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v) || v === 0) return "";
  return v.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

import { MONTHS, type Month, type PLReport, type PLRow } from "../engine/types";
import { fmtAup, fmtIndex, fmtNumber, fmtRatio } from "./format";

type Props = {
  report: PLReport;
  selectedId: string | null;
  onSelect: (id: string) => void;
};

function cellClass(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return "num";
  if (v < 0) return "num neg";
  return "num";
}

function MetricCells({
  row,
  period,
}: {
  row: PLRow;
  period: "mtd" | "ytd";
}) {
  const t = period === "mtd" ? row.values.mtd : row.values.ytd;
  const idxB = period === "mtd" ? row.values.mtdIndexBudget : row.values.ytdIndexBudget;
  const idxL = period === "mtd" ? row.values.mtdIndexLy : row.values.ytdIndexLy;
  const ratio = period === "mtd" ? row.values.mtdRatio : row.values.ytdRatio;
  const fmt = row.isAup ? fmtAup : (v: number | null | undefined) => fmtNumber(v, 0);

  if (row.isAup && row.id.includes("vs")) {
    return (
      <>
        <td className={cellClass(t.ly)}>{fmtRatio(t.ly)}</td>
        <td className={cellClass(t.budget)}>{fmtRatio(t.budget)}</td>
        <td className={cellClass(t.actual)}>{fmtRatio(t.actual)}</td>
        <td className="num">{fmtIndex(idxB)}</td>
        <td className="num">{fmtIndex(idxL)}</td>
        <td className="num muted">-</td>
        <td className="num muted">-</td>
        <td className="num muted">-</td>
      </>
    );
  }

  return (
    <>
      <td className={cellClass(t.ly)}>{fmt(t.ly)}</td>
      <td className={cellClass(t.budget)}>{fmt(t.budget)}</td>
      <td className={cellClass(t.actual)}>{fmt(t.actual)}</td>
      <td className="num">{fmtIndex(idxB)}</td>
      <td className="num">{fmtIndex(idxL)}</td>
      <td className="num">{fmtRatio(ratio.ly)}</td>
      <td className="num">{fmtRatio(ratio.budget)}</td>
      <td className="num">{fmtRatio(ratio.actual)}</td>
    </>
  );
}

function MonthlyActualCells({ row, selectedMonth }: { row: PLRow; selectedMonth: Month }) {
  const selectedIdx = MONTHS.indexOf(selectedMonth);
  const fmt = (v: number) => {
    if (row.isAup && row.id.includes("vs")) return fmtRatio(v);
    if (row.isAup) return fmtAup(v);
    return fmtNumber(v, 0);
  };

  return (
    <>
      {MONTHS.map((m, i) => {
        const v = row.monthlyActuals[i] ?? 0;
        const classes = [
          cellClass(v),
          "mth-actual",
          m === selectedMonth ? "is-selected-mth" : "",
          i > selectedIdx ? "out-ytd-mth" : "",
        ]
          .filter(Boolean)
          .join(" ");
        return (
          <td key={m} className={classes}>
            {fmt(v)}
          </td>
        );
      })}
    </>
  );
}

export function PLTable({ report, selectedId, onSelect }: Props) {
  const { filters, rows } = report;

  return (
    <div className="pl-wrap">
      <div className="pl-meta">
        <div>
          <span className="meta-label">Month</span> {filters.month}
        </div>
        <div>
          <span className="meta-label">Channel</span> {filters.channel}
        </div>
        <div>
          <span className="meta-label">Brand</span> {filters.brand}
        </div>
        <div className="meta-note">
          Click a line for workings · Monthly Actuals = 2026 MTD by month
        </div>
      </div>

      <div className="pl-scroll">
        <table className="pl-table">
          <thead>
            <tr>
              <th className="sticky-col label-col" rowSpan={3}>
                Line item
              </th>
              <th colSpan={8} className="group-mtd">
                MTD
              </th>
              <th colSpan={8} className="group-ytd">
                YTD
              </th>
              <th colSpan={12} className="group-actuals">
                Monthly Actuals (2026)
              </th>
            </tr>
            <tr>
              <th colSpan={3} className="group-mtd">
                Value
              </th>
              <th colSpan={2} className="group-mtd">
                Index
              </th>
              <th colSpan={3} className="group-mtd">
                Financial Ratio
              </th>
              <th colSpan={3} className="group-ytd">
                Value
              </th>
              <th colSpan={2} className="group-ytd">
                Index
              </th>
              <th colSpan={3} className="group-ytd">
                Financial Ratio
              </th>
              <th colSpan={12} className="group-actuals">
                Actual
              </th>
            </tr>
            <tr>
              <th className="group-mtd">LY</th>
              <th className="group-mtd">Budget</th>
              <th className="group-mtd">Actual</th>
              <th className="group-mtd">vs Bud</th>
              <th className="group-mtd">vs LY</th>
              <th className="group-mtd">LY</th>
              <th className="group-mtd">Budget</th>
              <th className="group-mtd">Actual</th>
              <th className="group-ytd">LY</th>
              <th className="group-ytd">Budget</th>
              <th className="group-ytd">Actual</th>
              <th className="group-ytd">vs Bud</th>
              <th className="group-ytd">vs LY</th>
              <th className="group-ytd">LY</th>
              <th className="group-ytd">Budget</th>
              <th className="group-ytd">Actual</th>
              {MONTHS.map((m) => (
                <th
                  key={m}
                  className={`group-actuals ${m === filters.month ? "is-selected-mth-h" : ""}`}
                >
                  {m}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr
                key={r.id}
                className={[r.bold ? "row-bold" : "", selectedId === r.id ? "row-selected" : ""]
                  .filter(Boolean)
                  .join(" ")}
                onClick={() => onSelect(r.id)}
              >
                <td
                  className="sticky-col label-col"
                  style={{ paddingLeft: `${12 + r.indent * 16}px` }}
                >
                  {r.label}
                </td>
                <MetricCells row={r} period="mtd" />
                <MetricCells row={r} period="ytd" />
                <MonthlyActualCells row={r} selectedMonth={filters.month} />
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

import { useState } from "react";
import type { MetricWorkings, RowWorkings } from "../engine/workings";
import { fmtNumber } from "./format";

type Props = {
  workings: RowWorkings | null;
  onClose: () => void;
};

function val(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return "－";
  return v.toLocaleString("en-US", {
    maximumFractionDigits: 4,
  });
}

function SpreadsheetDetailBlock({ metric }: { metric: MetricWorkings }) {
  const [open, setOpen] = useState(false);
  const detail = metric.spreadsheet;
  if (!detail) return null;

  return (
    <div className="ss-detail">
      <button type="button" className="ss-toggle" onClick={() => setOpen((v) => !v)}>
        {open ? "Hide spreadsheet detail" : "Further detail (as per spreadsheet)"}
      </button>
      {open ? (
        <div className="ss-body">
          <div className="ss-block">
            <div className="ss-label">Excel formula</div>
            <pre className="ss-formula">{detail.excelFormula}</pre>
          </div>
          <div className="ss-block">
            <div className="ss-label">Resolved with filters</div>
            <pre className="ss-formula resolved">{detail.excelFormulaResolved}</pre>
          </div>
          {detail.keyParts && detail.keyParts.length > 0 ? (
            <div className="ss-block">
              <div className="ss-label">Key / formula parts</div>
              <table className="monthly-table">
                <thead>
                  <tr>
                    <th>Cell / part</th>
                    <th>Meaning</th>
                    <th>Value</th>
                  </tr>
                </thead>
                <tbody>
                  {detail.keyParts.map((p) => (
                    <tr key={`${p.cell}-${p.part}`}>
                      <td>
                        <code>{p.cell}</code>
                      </td>
                      <td>{p.part}</td>
                      <td className="num">{p.value}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}
          {detail.notes && detail.notes.length > 0 ? (
            <ul className="ss-notes">
              {detail.notes.map((n) => (
                <li key={n}>{n}</li>
              ))}
            </ul>
          ) : null}
          {detail.brandBreakdown && detail.brandBreakdown.length > 0 ? (
            <div className="ss-block">
              <div className="ss-label">
                Brand breakdown in Working(Local) (same account / channel / month / year)
              </div>
              <table className="monthly-table">
                <thead>
                  <tr>
                    <th>Brand</th>
                    <th>Key</th>
                    <th>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {detail.brandBreakdown.map((b) => (
                    <tr key={b.key}>
                      <td>{b.brand}</td>
                      <td className="key-cell">
                        <code>{b.key}</code>
                      </td>
                      <td className={b.value < 0 ? "neg" : undefined}>{fmtNumber(b.value, 2)}</td>
                    </tr>
                  ))}
                  <tr className="monthly-total">
                    <td colSpan={2}>Sum of brands</td>
                    <td>
                      {fmtNumber(
                        detail.brandBreakdown.reduce((s, b) => s + b.value, 0),
                        2,
                      )}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}

function MetricCard({ metric }: { metric: MetricWorkings }) {
  return (
    <section className="workings-card">
      <header>
        <strong>{metric.title}</strong>
        <span className={metric.result != null && metric.result < 0 ? "neg" : undefined}>
          {val(metric.result)}
        </span>
      </header>
      {metric.source ? (
        <div className="workings-source">
          Sheet: <code>{metric.source}</code>
        </div>
      ) : null}
      {metric.key ? (
        <div className="workings-key">
          Key: <code>{metric.key}</code>
        </div>
      ) : null}
      <ul>
        {metric.steps.map((s, i) => (
          <li key={`${metric.title}-${i}`}>
            <span className="step-label">{s.label}</span>
            {s.detail ? <span className="step-detail">{s.detail}</span> : null}
            {s.value != null ? (
              <span className={`step-value ${s.value < 0 ? "neg" : ""}`}>{val(s.value)}</span>
            ) : null}
          </li>
        ))}
      </ul>
      {metric.monthly && metric.monthly.length > 0 ? (
        <table className="monthly-table">
          <thead>
            <tr>
              <th>Month</th>
              <th>Amount</th>
            </tr>
          </thead>
          <tbody>
            {metric.monthly.map((row) => (
              <tr key={row.month}>
                <td>{row.month}</td>
                <td className={row.value < 0 ? "neg" : undefined}>{fmtNumber(row.value, 2)}</td>
              </tr>
            ))}
            <tr className="monthly-total">
              <td>YTD</td>
              <td>{fmtNumber(metric.monthly.reduce((s, x) => s + x.value, 0), 2)}</td>
            </tr>
          </tbody>
        </table>
      ) : null}
      <SpreadsheetDetailBlock metric={metric} />
    </section>
  );
}

export function WorkingsPanel({ workings, onClose }: Props) {
  if (!workings) {
    return (
      <aside className="workings empty">
        <h2>Workings</h2>
        <p>Click any P&amp;L line to see how LY / Budget / Actual were calculated.</p>
      </aside>
    );
  }

  const { filters, monthlyActuals } = workings;
  const ytdActual = monthlyActuals.filter((m) => m.inYtd).reduce((s, m) => s + m.actual, 0);
  const ytdLy = monthlyActuals.filter((m) => m.inYtd).reduce((s, m) => s + m.ly, 0);
  const ytdBudget = monthlyActuals.filter((m) => m.inYtd).reduce((s, m) => s + m.budget, 0);

  return (
    <aside className="workings">
      <div className="workings-head">
        <div>
          <h2>Workings</h2>
          <p className="workings-title">{workings.label}</p>
        </div>
        <button type="button" className="btn-secondary" onClick={onClose}>
          Close
        </button>
      </div>

      <div className="workings-filters">
        <div>
          <span>Month</span> {filters.month}
        </div>
        <div>
          <span>Channel</span> {filters.channel}
          {workings.channelMapped ? (
            <em className="map-note"> → budget {workings.budgetChannel}</em>
          ) : null}
        </div>
        <div>
          <span>Brand</span> {filters.brand}
          {filters.channel !== "All Chain" ? (
            <em className="map-note"> → budget {workings.budgetBrand}</em>
          ) : null}
        </div>
      </div>

      <p className="workings-method">{workings.method}</p>

      <section className="workings-card monthly-actuals-card">
        <header>
          <strong>Monthly actuals</strong>
          <span className="muted-tiny">
            MTD by month · {filters.channel} / {filters.brand}
          </span>
        </header>
        <table className="monthly-table monthly-actuals-table">
          <thead>
            <tr>
              <th>Mth</th>
              <th>LY</th>
              <th>Budget</th>
              <th>Actual</th>
            </tr>
          </thead>
          <tbody>
            {monthlyActuals.map((row) => (
              <tr
                key={row.month}
                className={[
                  row.isSelected ? "is-selected-month" : "",
                  row.inYtd ? "in-ytd" : "out-ytd",
                ]
                  .filter(Boolean)
                  .join(" ")}
              >
                <td>{row.month}</td>
                <td className={row.ly < 0 ? "neg" : undefined}>{fmtNumber(row.ly, 0)}</td>
                <td className={row.budget < 0 ? "neg" : undefined}>{fmtNumber(row.budget, 0)}</td>
                <td className={row.actual < 0 ? "neg" : undefined}>{fmtNumber(row.actual, 0)}</td>
              </tr>
            ))}
            <tr className="monthly-total">
              <td>YTD→{filters.month}</td>
              <td className={ytdLy < 0 ? "neg" : undefined}>{fmtNumber(ytdLy, 0)}</td>
              <td className={ytdBudget < 0 ? "neg" : undefined}>{fmtNumber(ytdBudget, 0)}</td>
              <td className={ytdActual < 0 ? "neg" : undefined}>{fmtNumber(ytdActual, 0)}</td>
            </tr>
          </tbody>
        </table>
      </section>

      <div className="workings-metrics">
        {workings.metrics.map((m) => (
          <MetricCard key={m.title} metric={m} />
        ))}
      </div>

      <section className="workings-card">
        <header>
          <strong>Indexes</strong>
        </header>
        <ul>
          {workings.indexes.map((s) => (
            <li key={s.label}>
              <span className="step-label">{s.label}</span>
              {s.detail ? <span className="step-detail">{s.detail}</span> : null}
              {s.value != null ? <span className="step-value">{val(s.value)}</span> : null}
            </li>
          ))}
        </ul>
      </section>

      <section className="workings-card">
        <header>
          <strong>Financial ratios</strong>
        </header>
        <ul>
          {workings.ratios.map((s) => (
            <li key={s.label}>
              <span className="step-label">{s.label}</span>
              {s.detail ? <span className="step-detail">{s.detail}</span> : null}
            </li>
          ))}
        </ul>
      </section>
    </aside>
  );
}

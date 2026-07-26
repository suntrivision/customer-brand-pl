import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  buildCustomerCompare,
  emptyActualsInput,
  fillDerivedInput,
  inputtableLineCount,
  isTrivisionMapperSheet,
  parseActualsRows,
  parseActualsText,
  parseComparativeSheet,
  parseTrivisionMapperSheet,
  resolveCustomerChannel,
  seedInputFromPl,
  type ActualsInput,
} from "../engine/compare";
import { MONTHS, type Filters, type Month, type WorkbookData } from "../engine/types";
import { fmtNumber } from "./format";

type Props = {
  data: WorkbookData;
  channels: string[];
  brands: string[];
  initialFilters: Filters;
};

function numClass(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return "num";
  if (v < 0) return "num neg";
  return "num";
}

export function CustomerCompare({ data, channels, brands, initialFilters }: Props) {
  const [filters, setFilters] = useState<Filters>({
    ...initialFilters,
    channel: channels.includes("DIST-CHC")
      ? "DIST-CHC"
      : channels.find((c) => c !== "All Chain") ?? initialFilters.channel,
    brand: brands.includes("All Brand") ? "All Brand" : initialFilters.brand,
  });
  const [input, setInput] = useState<ActualsInput>(() => emptyActualsInput());
  const [sourceLabel, setSourceLabel] = useState("Trivision / Input Actuals");
  const [pasteText, setPasteText] = useState("");
  const [status, setStatus] = useState<string | null>(null);
  const [showOnlyWithInput, setShowOnlyWithInput] = useState(false);

  const report = useMemo(
    () => buildCustomerCompare(data, filters, input, sourceLabel),
    [data, filters, input, sourceLabel],
  );

  const visibleRows = showOnlyWithInput
    ? report.rows.filter((r) => r.hasInput || r.spacer)
    : report.rows;

  function setAmount(id: string, raw: string) {
    const trimmed = raw.trim();
    if (trimmed === "") {
      setInput((prev) => ({ ...prev, [id]: null }));
      return;
    }
    const n = Number(trimmed.replace(/,/g, ""));
    if (!Number.isFinite(n)) return;
    setInput((prev) => ({ ...prev, [id]: n }));
  }

  function applyPaste() {
    const { input: next, matched, unmatched } = parseActualsText(pasteText);
    setInput((prev) => ({ ...prev, ...next }));
    setSourceLabel("Pasted Actuals");
    setStatus(
      unmatched.length
        ? `Matched ${matched} lines · ${unmatched.length} unmatched`
        : `Matched ${matched} lines`,
    );
  }

  function applyTrivisionInput(
    parsed: {
      input: ActualsInput;
      matched: number;
      unmatched: string[];
      detectedCustomer: string | null;
      sourceLabel: string;
    },
    fileName: string,
    formatNote: string,
  ) {
    setInput(fillDerivedInput({ ...emptyActualsInput(), ...parsed.input }));
    setSourceLabel(parsed.sourceLabel || fileName);
    const channel = resolveCustomerChannel(parsed.detectedCustomer, channels);
    if (channel) {
      setFilters((f) => ({ ...f, channel }));
    }
    const custNote = channel
      ? ` · customer ${channel}`
      : parsed.detectedCustomer
        ? ` · customer “${parsed.detectedCustomer}” (not matched — pick Channel)`
        : "";
    setStatus(
      parsed.unmatched.length
        ? `Loaded ${parsed.matched} from ${fileName}${custNote} · ${parsed.unmatched.length} unmatched`
        : `Loaded ${parsed.matched} lines from ${fileName}${custNote} · ${formatNote}`,
    );
  }

  async function onUploadActuals(file: File) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheetName =
        wb.SheetNames.find((n) => /compar/i.test(n)) ?? wb.SheetNames[0]!;
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
        header: 1,
        defval: null,
        raw: true,
      });

      // Trivision / P&L Mapper (dtck.xlsx): Label | base/calculated | brand | customer | amount
      if (isTrivisionMapperSheet(rows)) {
        applyTrivisionInput(parseTrivisionMapperSheet(rows), file.name, "Trivision P&L");
        return;
      }

      // ComparitiveAnalysis.xlsx layout
      const looksComparative = rows.some((r) => {
        const a = String(r?.[0] ?? "");
        const b = String(r?.[1] ?? "");
        return /p&l line/i.test(a) || /trivision/i.test(b) || /trivision/i.test(a);
      });

      if (looksComparative) {
        const parsed = parseComparativeSheet(rows, filters.channel);
        applyTrivisionInput(parsed, file.name, "Comparitive Analysis");
        return;
      }

      // Simple Label | Amount (first two columns)
      const parsedRows: { label: string; amount: number }[] = [];
      for (const row of rows) {
        if (!row || row.length < 2) continue;
        const label = String(row[0] ?? "").trim();
        const amount = Number(row[1]);
        if (!label || /p&l line/i.test(label) || /^line$/i.test(label)) continue;
        if (!Number.isFinite(amount)) continue;
        parsedRows.push({ label, amount });
      }
      const { input: next, matched, unmatched } = parseActualsRows(parsedRows);
      setInput(fillDerivedInput({ ...emptyActualsInput(), ...next }));
      setSourceLabel(file.name);
      setStatus(
        unmatched.length
          ? `Loaded ${matched} from ${file.name} · ${unmatched.length} unmatched`
          : `Loaded ${matched} lines from ${file.name}`,
      );
    } catch (e) {
      setStatus(e instanceof Error ? e.message : "Failed to read actuals file");
    }
  }

  const lineTotal = inputtableLineCount();

  return (
    <div className="compare-wrap">
      <div className="compare-toolbar">
        <div className="compare-title">
          <h2>Customer comparison</h2>
          <p>
            Upload a Trivision P&amp;L export (e.g. dtck.xlsx) or ComparitiveAnalysis.xlsx, then
            compare line-by-line to Customer Brand Local P&amp;L Actuals for that customer.
          </p>
        </div>

        <div className="toolbar-filters">
          <label>
            Month
            <select
              value={filters.month}
              onChange={(e) => setFilters({ ...filters, month: e.target.value as Month })}
            >
              {MONTHS.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          </label>
          <label>
            Customer (Channel)
            <select
              value={filters.channel}
              onChange={(e) => setFilters({ ...filters, channel: e.target.value })}
            >
              {channels.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
          </label>
          <label>
            Brand
            <select
              value={filters.brand}
              onChange={(e) => setFilters({ ...filters, brand: e.target.value })}
            >
              {brands.map((b) => (
                <option key={b} value={b}>
                  {b}
                </option>
              ))}
            </select>
          </label>
        </div>
      </div>

      <div className="compare-input-panel">
        <div className="compare-input-actions">
          <label className="btn-secondary file-btn">
            Upload Trivision P&amp;L (e.g. dtck.xlsx)
            <input
              type="file"
              accept=".xlsx,.xlsm,.csv"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) void onUploadActuals(f);
              }}
            />
          </label>
          <button
            type="button"
            className="btn-secondary"
            onClick={() => {
              setInput(seedInputFromPl(data, filters));
              setSourceLabel("Seeded from P&L Actuals");
              setStatus("Filled input with current P&L Actuals (edit to compare)");
            }}
          >
            Seed from P&amp;L
          </button>
          <button
            type="button"
            className="btn-secondary"
            onClick={() => {
              setInput(emptyActualsInput());
              setPasteText("");
              setSourceLabel("Trivision / Input Actuals");
              setStatus("Cleared input Actuals");
            }}
          >
            Clear input
          </button>
          <label className="compare-check">
            <input
              type="checkbox"
              checked={showOnlyWithInput}
              onChange={(e) => setShowOnlyWithInput(e.target.checked)}
            />
            Only lines with input
          </label>
        </div>

        <label className="paste-label">
          Paste Actuals (Label, Amount — Comparative Analysis Col A / Col B)
          <textarea
            value={pasteText}
            onChange={(e) => setPasteText(e.target.value)}
            placeholder={`Sell In (Gross Sales Value - GSV) - Total, 526374\nIndirect COGS-SC, 22345.91\nGross Margin After Promoter & Merchandiser, 75675.53`}
            rows={4}
          />
        </label>
        <button type="button" className="btn-primary" onClick={applyPaste}>
          Apply pasted Actuals
        </button>
        {status ? <p className="compare-status">{status}</p> : null}
      </div>

      <div className="compare-meta">
        <span>
          <strong>{sourceLabel}</strong> vs Customer Brand P&amp;L · {filters.channel} ·{" "}
          {filters.brand} · {filters.month} 2026
        </span>
        <span>
          {report.inputCount} / {lineTotal} lines entered · Diff = P&amp;L − Input
        </span>
      </div>

      <div className="compare-scroll">
        <table className="compare-table">
          <thead>
            <tr>
              <th className="sticky-col">P&amp;L line items</th>
              <th className="group-input">{sourceLabel}</th>
              <th className="group-pl">{filters.channel}</th>
              <th className="group-diff">Difference</th>
              <th className="group-diff">Diff %</th>
            </tr>
          </thead>
          <tbody>
            {visibleRows.map((r) =>
              r.spacer ? (
                <tr key={r.id} className="spacer-row">
                  <td className="sticky-col" colSpan={5} />
                </tr>
              ) : (
                <tr key={r.id} className={r.hasInput ? undefined : "no-input"}>
                  <td className="sticky-col">{r.label}</td>
                  <td className="input-cell">
                    <input
                      type="text"
                      inputMode="decimal"
                      value={r.inputActual == null ? "" : String(r.inputActual)}
                      onChange={(e) => setAmount(r.id, e.target.value)}
                      placeholder="—"
                    />
                  </td>
                  <td className={numClass(r.plActual)}>{fmtNumber(r.plActual, 2)}</td>
                  <td className={numClass(r.diff)}>
                    {r.diff == null ? "—" : fmtNumber(r.diff, 2)}
                  </td>
                  <td className={numClass(r.diffPct)}>
                    {r.diffPct == null ? "—" : `${fmtNumber(r.diffPct, 1)}%`}
                  </td>
                </tr>
              ),
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

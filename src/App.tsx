import { useMemo, useState } from "react";
import { CustomerCompare } from "./components/CustomerCompare";
import { FileUpload } from "./components/FileUpload";
import { FilterBar } from "./components/FilterBar";
import { PLTable } from "./components/PLTable";
import { WorkingsPanel } from "./components/WorkingsPanel";
import { computePL, withMonthlyActuals } from "./engine/computePL";
import { parseWorkbook } from "./engine/parseWorkbook";
import { buildRowWorkings } from "./engine/workings";
import type { Filters, WorkbookData } from "./engine/types";
import "./App.css";

const DEFAULT_FILTERS: Filters = {
  month: "May",
  channel: "All Chain",
  brand: "All Brand",
};

type AppView = "pl" | "compare";

export default function App() {
  const [data, setData] = useState<WorkbookData | null>(null);
  const [fileName, setFileName] = useState("");
  const [filters, setFilters] = useState<Filters>(DEFAULT_FILTERS);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>("gsv");
  const [view, setView] = useState<AppView>("pl");

  const report = useMemo(() => {
    if (!data) return null;
    return withMonthlyActuals(data, filters);
  }, [data, filters]);

  const workings = useMemo(() => {
    if (!data || !report || !selectedId || view !== "pl") return null;
    const base = computePL(data, filters);
    return buildRowWorkings(data, filters, base, selectedId);
  }, [data, filters, report, selectedId, view]);

  async function handleFile(file: File) {
    setBusy(true);
    setError(null);
    try {
      const buffer = await file.arrayBuffer();
      await new Promise((r) => setTimeout(r, 30));
      const parsed = parseWorkbook(buffer);
      setData(parsed);
      setFileName(file.name);
      setFilters({
        month: "May",
        channel: parsed.channels.includes("All Chain") ? "All Chain" : parsed.channels[0]!,
        brand: parsed.brands.includes("All Brand") ? "All Brand" : parsed.brands[0]!,
      });
      setSelectedId("gsv");
      setView("pl");
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "Failed to parse workbook");
    } finally {
      setBusy(false);
    }
  }

  function reset() {
    setData(null);
    setFileName("");
    setError(null);
    setFilters(DEFAULT_FILTERS);
    setSelectedId(null);
    setView("pl");
  }

  if (!data || !report) {
    return <FileUpload onFile={handleFile} busy={busy} error={error} />;
  }

  return (
    <div className="app">
      <FilterBar
        filters={filters}
        channels={data.channels}
        brands={data.brands}
        fileName={fileName}
        hideFilters={view === "compare"}
        onChange={(next) => {
          setFilters(next);
        }}
        onReset={reset}
      />
      <nav className="view-tabs" aria-label="App views">
        <button
          type="button"
          className={view === "pl" ? "tab active" : "tab"}
          onClick={() => setView("pl")}
        >
          Local P&amp;L
        </button>
        <button
          type="button"
          className={view === "compare" ? "tab active" : "tab"}
          onClick={() => setView("compare")}
        >
          Customer comparison
        </button>
      </nav>
      {view === "pl" ? (
        <div className="app-main">
          <PLTable report={report} selectedId={selectedId} onSelect={setSelectedId} />
          <WorkingsPanel workings={workings} onClose={() => setSelectedId(null)} />
        </div>
      ) : (
        <CustomerCompare
          data={data}
          channels={data.channels}
          brands={data.brands}
          initialFilters={filters}
        />
      )}
    </div>
  );
}

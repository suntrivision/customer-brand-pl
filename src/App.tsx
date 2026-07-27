import { useMemo, useState } from "react";
import { CustomerCompare } from "./components/CustomerCompare";
import { FileUpload } from "./components/FileUpload";
import { FilterBar } from "./components/FilterBar";
import { PLTable } from "./components/PLTable";
import { WorkingsPanel } from "./components/WorkingsPanel";
import {
  parseAllocationAudit,
  type AllocationAuditData,
} from "./engine/allocationAudit";
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
  const [audit, setAudit] = useState<AllocationAuditData | null>(null);
  const [auditFileName, setAuditFileName] = useState<string | null>(null);
  const [auditStatus, setAuditStatus] = useState<string | null>(null);

  const report = useMemo(() => {
    if (!data) return null;
    return withMonthlyActuals(data, filters);
  }, [data, filters]);

  const workings = useMemo(() => {
    if (!data || !report || !selectedId || view !== "pl") return null;
    const base = computePL(data, filters);
    return buildRowWorkings(data, filters, base, selectedId, audit);
  }, [data, filters, report, selectedId, view, audit]);

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
    setAudit(null);
    setAuditFileName(null);
    setAuditStatus(null);
  }

  async function handleAuditFile(file: File) {
    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseAllocationAudit(buffer, file.name);
      setAudit(parsed);
      setAuditFileName(file.name);
      setAuditStatus(
        `Loaded audit ${parsed.customerCode} · ${parsed.period}` +
          (parsed.runId ? ` · run ${parsed.runId.slice(0, 8)}…` : ""),
      );
      // Align filters to audit customer/month when possible
      const monthMap: Record<string, Filters["month"]> = {
        "01": "Jan",
        "02": "Feb",
        "03": "Mar",
        "04": "Apr",
        "05": "May",
        "06": "Jun",
        "07": "Jul",
        "08": "Aug",
        "09": "Sep",
        "10": "Oct",
        "11": "Nov",
        "12": "Dec",
      };
      const mm = parsed.period.slice(5, 7);
      const month = monthMap[mm];
      const codeToChannel: Record<string, string> = {
        DTCH: "DIST-CHC",
        DTCK: "DIST-CK DISTRIBUTORS",
      };
      const channel = codeToChannel[parsed.customerCode.toUpperCase()];
      setFilters((f) => ({
        ...f,
        ...(month ? { month } : {}),
        ...(channel && data?.channels.includes(channel) ? { channel } : {}),
      }));
    } catch (e) {
      setAuditStatus(e instanceof Error ? e.message : "Failed to read allocation audit");
    }
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
          <WorkingsPanel
            workings={workings}
            onClose={() => setSelectedId(null)}
            auditFileName={auditFileName}
            auditStatus={auditStatus}
            onUploadAudit={(f) => void handleAuditFile(f)}
            onClearAudit={() => {
              setAudit(null);
              setAuditFileName(null);
              setAuditStatus("Cleared allocation audit");
            }}
          />
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

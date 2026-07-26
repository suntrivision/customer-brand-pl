import { MONTHS, type Filters, type Month } from "../engine/types";

type Props = {
  filters: Filters;
  channels: string[];
  brands: string[];
  fileName: string;
  hideFilters?: boolean;
  onChange: (next: Filters) => void;
  onReset: () => void;
};

export function FilterBar({
  filters,
  channels,
  brands,
  fileName,
  hideFilters = false,
  onChange,
  onReset,
}: Props) {
  return (
    <header className="toolbar">
      <div className="toolbar-brand">
        <strong>Customer Brand P&amp;L</strong>
        <span className="toolbar-file" title={fileName}>
          {fileName}
        </span>
      </div>
      <div className="toolbar-filters">
        {!hideFilters ? (
          <>
            <label>
              Month
              <select
                value={filters.month}
                onChange={(e) => onChange({ ...filters, month: e.target.value as Month })}
              >
                {MONTHS.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Channel
              <select
                value={filters.channel}
                onChange={(e) => onChange({ ...filters, channel: e.target.value })}
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
                onChange={(e) => onChange({ ...filters, brand: e.target.value })}
              >
                {brands.map((b) => (
                  <option key={b} value={b}>
                    {b}
                  </option>
                ))}
              </select>
            </label>
          </>
        ) : null}
        <button type="button" className="btn-secondary" onClick={onReset}>
          Load another file
        </button>
      </div>
    </header>
  );
}

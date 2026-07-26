type Props = {
  onFile: (file: File) => void;
  busy: boolean;
  error: string | null;
};

export function FileUpload({ onFile, busy, error }: Props) {
  return (
    <div className="upload">
      <div className="upload-card">
        <h1>Customer Brand P&amp;L</h1>
        <p className="upload-sub">
          Upload the Customer Brand PL workbook to view Local P&amp;L with Month, Channel, and Brand
          filters.
        </p>
        <label className={`upload-drop ${busy ? "is-busy" : ""}`}>
          <input
            type="file"
            accept=".xlsx,.xlsm"
            disabled={busy}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) onFile(f);
            }}
          />
          <span className="upload-cta">{busy ? "Parsing workbook…" : "Choose Excel file"}</span>
          <span className="upload-hint">.xlsx · Local P&amp;L sheets are read in the browser</span>
        </label>
        {error ? <p className="upload-error">{error}</p> : null}
      </div>
    </div>
  );
}

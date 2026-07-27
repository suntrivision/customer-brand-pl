import { readFileSync } from "node:fs";
import { parseAllocationAudit, buildPreAllocationDetail } from "../src/engine/allocationAudit.ts";

const path =
  "c:/Totenku/Premchand/checkexpenses/DTCH_Allocation_Audit_Workbook_2026-04.xlsx";
const buf = readFileSync(path);
const audit = parseAllocationAudit(
  buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  "DTCH_Allocation_Audit_Workbook_2026-04.xlsx",
);
console.log(audit.customerCode, audit.period, "prePost", audit.prePost.length, "gl", audit.glRows.length);
const sc = buildPreAllocationDetail(audit, {
  channel: "DIST-CHC",
  month: "Apr",
  account: "Indirect COGS-SC",
});
console.log("available", sc.available, sc.grouping);
for (const s of sc.steps) {
  console.log(s.step, s.label, s.value ?? "", (s.detail ?? "").slice(0, 100));
}
console.log("native brands", sc.nativeGlRows.length, "sum", sc.nativeGlRows.reduce((a, r) => a + r.pre, 0));

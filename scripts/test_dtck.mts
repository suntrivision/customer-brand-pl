import { readFileSync } from "node:fs";
import * as XLSX from "xlsx";
import {
  isTrivisionMapperSheet,
  parseTrivisionMapperSheet,
  resolveCustomerChannel,
  buildCustomerCompare,
} from "../src/engine/compare.ts";
import { parseWorkbook } from "../src/engine/parseWorkbook.ts";

const dtckPath = "c:/Totenku/Premchand/checkexpenses/26July/dtck.xlsx";
const plPath =
  "c:/Totenku/Premchand/checkexpenses/25July2026/06 Customer Brand PL Jun'26 .07.2026.xlsx";

const dtckBuf = readFileSync(dtckPath);
const wb = XLSX.read(dtckBuf, { type: "buffer" });
const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(wb.Sheets[wb.SheetNames[0]!], {
  header: 1,
  defval: null,
  raw: true,
});
console.log("isTrivisionMapperSheet", isTrivisionMapperSheet(rows));
const parsed = parseTrivisionMapperSheet(rows);
console.log("matched", parsed.matched, "unmatched", parsed.unmatched);
console.log("customer raw", parsed.detectedCustomer);
console.log("gsv", parsed.input.gsv, "indir-sc", parsed.input["indir-sc"], "gm", parsed.input["gm-after-promoter"]);

const plBuf = readFileSync(plPath);
const data = parseWorkbook(plBuf.buffer.slice(plBuf.byteOffset, plBuf.byteOffset + plBuf.byteLength));
const channel = resolveCustomerChannel(parsed.detectedCustomer, data.channels);
console.log("resolved channel", channel);

const report = buildCustomerCompare(
  data,
  { month: "Apr", channel: channel!, brand: "All Brand" },
  parsed.input,
  parsed.sourceLabel,
);
const focus = ["gsv", "indir-sc", "gm-after-promoter", "net-revenue", "nop"];
for (const id of focus) {
  const r = report.rows.find((x) => x.id === id)!;
  console.log(
    id,
    "input=",
    r.inputActual?.toFixed?.(2) ?? r.inputActual,
    "pl=",
    r.plActual?.toFixed?.(2) ?? r.plActual,
    "diff=",
    r.diff?.toFixed?.(2) ?? r.diff,
  );
}

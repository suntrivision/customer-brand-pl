import fs from "node:fs";
import { parseWorkbook } from "../src/engine/parseWorkbook.ts";
import { computePL } from "../src/engine/computePL.ts";

const path =
  process.argv[2] ??
  String.raw`c:\Totenku\Premchand\checkexpenses\25July2026\06 Customer Brand PL Jun'26 .07.2026.xlsx`;

const buf = fs.readFileSync(path);
const data = parseWorkbook(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));

const filters = { month: "May" as const, channel: "All Chain", brand: "All Brand" };
const report = computePL(data, filters);

// Excel cached (from inspect) for May / All Chain / All Brand
const expected: Record<string, { mtd: [number, number, number]; ytd?: [number, number, number] }> = {
  "so-tot-qty": { mtd: [1947822, 1941038.6160192448, 1514467], ytd: [7901860, 9403651.253572589, 8672141] },
  "so-tot-amt": {
    mtd: [22080204.03147001, 22944193.331990696, 16765417.752960017],
    ytd: [90124751.24578008, 109810632.35649005, 98611265.32240498],
  },
  gsv: {
    mtd: [25944219.40000002, 22862998.29433462, 18715519.520000007],
    ytd: [93102875.51000002, 112824762.79817432, 109405076.11999999],
  },
  "net-sales": {
    mtd: [18739723.660000004, 17149615.949316185, 13428468.000000004],
    ytd: [68333849.58, 85081800.98734671, 82317429.21999998],
  },
  "gross-margin": {
    mtd: [4166330.3099999987, 3849984.5850084033, 2051806.5500000007],
    ytd: [12824804.869999982, 20361638.50805617, 18547167.14999999],
  },
  nop: {
    mtd: [320395.1477741222, -499546.50391739653, -2183812.644438208],
    ytd: [-4839021.363890576, -1341531.0338547723, -3894195.325416434],
  },
};

function close(a: number | null, b: number, tol = 1) {
  if (a == null) return false;
  return Math.abs(a - b) <= tol || Math.abs(a - b) / Math.max(Math.abs(b), 1) < 1e-6;
}

let fails = 0;
for (const [id, exp] of Object.entries(expected)) {
  const row = report.rows.find((r) => r.id === id);
  if (!row) {
    console.log("MISSING ROW", id);
    fails++;
    continue;
  }
  const checks: [string, number | null, number][] = [
    ["mtd.ly", row.values.mtd.ly, exp.mtd[0]],
    ["mtd.budget", row.values.mtd.budget, exp.mtd[1]],
    ["mtd.actual", row.values.mtd.actual, exp.mtd[2]],
  ];
  if (exp.ytd) {
    checks.push(
      ["ytd.ly", row.values.ytd.ly, exp.ytd[0]],
      ["ytd.budget", row.values.ytd.budget, exp.ytd[1]],
      ["ytd.actual", row.values.ytd.actual, exp.ytd[2]],
    );
  }
  for (const [name, got, want] of checks) {
    const ok = close(got, want, Math.max(2, Math.abs(want) * 1e-6));
    if (!ok) {
      fails++;
      console.log(`FAIL ${id} ${name}: got=${got} want=${want} diff=${(got ?? 0) - want}`);
    } else {
      console.log(`OK   ${id} ${name}`);
    }
  }
}

console.log("\nchannels", data.channels.length, "brands", data.brands.length, "working keys", data.working.size);
console.log(fails ? `\n${fails} mismatches` : "\nAll spot-checks passed");
process.exit(fails ? 1 : 0);

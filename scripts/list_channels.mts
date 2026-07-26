import { readFileSync } from "node:fs";
import { parseWorkbook } from "../src/engine/parseWorkbook.ts";

const path =
  "c:/Totenku/Premchand/checkexpenses/25July2026/06 Customer Brand PL Jun'26 .07.2026.xlsx";
const buf = readFileSync(path);
const data = parseWorkbook(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
console.log(
  "CK/DTCK channels:",
  data.channels.filter((c) => /ck|dtck/i.test(c)),
);
console.log(
  "DIST channels:",
  data.channels.filter((c) => /^DIST/i.test(c)),
);

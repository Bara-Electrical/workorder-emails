#!/usr/bin/env node
// Regression tests for payerLinePm() — run with: npm test
//
// Bricks + Agent's newer template names the PM only on the PDF's payer line. Same harness as
// the other tools/*.test.mjs files: the function is lifted out of index.js and run alone.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src  = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");
const i = src.indexOf("function payerLinePm(text, agencyName)");
if (i === -1) throw new Error("index.js no longer contains payerLinePm");
const fn = src.slice(i, src.indexOf("\n}\n", i) + 3);

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "payer-pm-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, fn + "\nexport { payerLinePm };\n");
const { payerLinePm } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, got, expected) => results.push({ name, ok: got === expected, detail: `got ${JSON.stringify(got)}, expected ${JSON.stringify(expected)}` });

// PDF text as the AI receives it (whitespace collapsed), verbatim from the work orders.
const GAMAGE = "Invoice Details Initial Payment Balance Payment $ 0.00 $ 0.00 Payer Information Franziska Scharl- Paid via rental funds on behalf of owner Owner: Stewart Gordon Ogden C/O Professionals: The Wright Team 0893494655";
const PROSPECTOR = "Invoice Details Initial Payment Balance Payment $ 0.00 $ 0.00 Payer Information Owner: Fay Stockdale C/O Professionals: The Wright Team 0893494655 Alyssa Radmore-Collard - Paid via rental funds on behalf of owner";
const PEAK = "Invoice Details Initial Payment Balance Payment $ 0.00 $ 0.00 Payer Information Peak Central - Paid via rental funds on behalf of owner Terms & conditions";

check("27 Gamage Way (108740): name hard against the dash", payerLinePm(GAMAGE, "Professionals TWT Realty - The Wright Team"), "Franziska Scharl");
check("30 Prospector Loop (108486): hyphenated surname after the owner block", payerLinePm(PROSPECTOR, "Professionals TWT Realty - The Wright Team"), "Alyssa Radmore-Collard");
check("Peak Central puts its own name there: not a PM", payerLinePm(PEAK, "Peak Central"), null);
check("…matched case- and punctuation-insensitively", payerLinePm(PEAK, "PEAK CENTRAL."), null);
check("…and when the extracted agency is longer", payerLinePm(PEAK, "Peak Central Property Management"), null);
check("no payer line, no PM", payerLinePm("Job Details Job Title Rangehood", "Professionals"), null);
check("a single word is not taken as a person", payerLinePm("Payer Information Admin - Paid via rental funds", "X"), null);
check("empty text is safe", payerLinePm("", "X"), null);
check("null text is safe", payerLinePm(null, null), null);
check("no agency name still finds the person", payerLinePm(GAMAGE, null), "Franziska Scharl");

const failures = results.filter(r => !r.ok);
console.log(`payer pm: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}\n    ${f.detail}`).join("\n"));
  process.exit(1);
}

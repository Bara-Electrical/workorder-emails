#!/usr/bin/env node
// Regression tests for the "After hours" email tag — run with: npm test
//
// Tagged "After hours" in Outlook, a work order becomes an After Hours Callout due today
// (Perth). Same harness as the other tools/*.test.mjs files: the functions are lifted out of
// index.js so the tests run what ships.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src  = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");
function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  return src.slice(i, src.indexOf(endMarker, i) + endMarker.length);
}
const module_ = [
  grab("const AFTER_HOURS_CATEGORY", ";\n"),
  grab("function isAfterHoursEmail(categories)", "\n}\n"),
  grab("function pickAfterHoursTaskType(taskTypes)", "\n}\n"),
  grab("function perthDateString(", "\n}\n"),
  "export { isAfterHoursEmail, pickAfterHoursTaskType, perthDateString };",
].join("\n");
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "after-hours-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { isAfterHoursEmail, pickAfterHoursTaskType, perthDateString } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

check("the tag is found whatever its case", isAfterHoursEmail(["Bara AI", "After hours"]) && isAfterHoursEmail([" AFTER HOURS "]));
check("no tag, not after hours", !isAfterHoursEmail(["Bara AI", "Urgent"]) && !isAfterHoursEmail([]) && !isAfterHoursEmail(null));
check("'After hours callout requested' in another tag doesn't count", !isAfterHoursEmail(["After hours callout requested"]));

const TYPES = [
  { tasktypeid: "OLD", tasktype: "After Hours Callout", archived: "true" },
  { tasktypeid: "ALT", tasktype: "After Hours Call Out", archived: "false" },
  { tasktypeid: "AHC", tasktype: "After Hours Callout", archived: "false" },
  { tasktypeid: "EC1", tasktype: "$120 Standard Electrical Compliance", archived: "false" },
];
check("picks the live After Hours Callout", pickAfterHoursTaskType(TYPES) === "AHC", String(pickAfterHoursTaskType(TYPES)));
check("falls back to the other spelling", pickAfterHoursTaskType(TYPES.filter(t => t.tasktypeid !== "AHC")) === "ALT");
check("never an archived one, and null when there's none", pickAfterHoursTaskType(TYPES.slice(0, 1)) === null && pickAfterHoursTaskType([]) === null);

// 2am Perth on 10 Oct is 18:00 UTC on 9 Oct — still the 10th in Perth.
check("today is Perth's today", perthDateString(new Date("2026-10-09T18:00:00Z")) === "2026/10/10", perthDateString(new Date("2026-10-09T18:00:00Z")));
check("and in the evening too", perthDateString(new Date("2026-10-10T13:30:00Z")) === "2026/10/10");

const failures = results.filter(r => !r.ok);
console.log(`after hours: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

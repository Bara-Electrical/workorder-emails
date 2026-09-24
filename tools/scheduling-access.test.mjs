#!/usr/bin/env node
// Regression tests for schedulingAccessLines() — run with: npm test
//
// index.js starts an HTTP server and the poll loop the moment it is imported, so the
// functions under test are lifted out of the source text and evaluated in isolation, the same
// way as the other tools/*.test.mjs files.
//
// What these guard: a lockbox code must reach the scheduling note whoever manages the job
// sees it, not only when the AI happened to mark the property "Vacant". Job 108212
// (3 Marungi Way) had a lockbox code and no tenant, was never flagged vacant, and the code
// only reached the scheduling note when the office added it by hand.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src  = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");

function grab(startMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf("\n}\n", i);
  if (j === -1) throw new Error(`no end of function after: ${startMarker}`);
  return src.slice(i, j + 3);
}

const module_ = [
  grab("function extractLockboxDetails(accessDetails)"),
  grab("function extractKeyCollectionLine(taskDescription)"),
  grab("function schedulingAccessLines(result)"),
  "export { schedulingAccessLines };",
].join("\n");

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "scheduling-access-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { schedulingAccessLines } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const CASES = [
  // The live case, exactly as the AI logged it: no tenant, not flagged vacant, a lockbox.
  ["108212: lockbox with no tenant and no 'vacant'",
    { "tenant-name": null, "access-details": "Lockbox code: 214",
      "task-description": "Repair front yard light fitting.\nReplace missing globes." },
    ["Lockbox code: 214"]],

  // A named tenant doesn't make the code useless — they may not be home.
  ["lockbox with a named tenant",
    { "tenant-name": "Emma", "access-details": "Lockbox code: 5" },
    ["Lockbox code: 5"]],

  // Unchanged from before: vacant properties keep the "Vacant - " prefix.
  ["vacant with a lockbox",
    { "tenant-name": "Vacant", "access-details": "Lockbox code: 214" },
    ["Vacant - Lockbox code: 214"]],
  ["vacant with a key-collection line",
    { "tenant-name": "Vacant", "access-details": null,
      "task-description": "Replace globes.\nCollect keys from office after 20/09." },
    ["Vacant - Collect keys from office after 20/09."]],
  ["vacant with both",
    { "tenant-name": "  vacant ", "access-details": "Lockbox code: 9",
      "task-description": "Collect keys after 1pm." },
    ["Vacant - Lockbox code: 9", "Vacant - Collect keys after 1pm."]],

  // A key-collection line is still vacant-only: someone is there to let the tech in.
  ["key collection without vacant stays out",
    { "tenant-name": null, "access-details": null,
      "task-description": "Collect keys after 1pm." },
    []],

  // Only lockbox info surfaces; keys we hold and gate/swipe access are arranged separately.
  ["key number and gate code alone",
    { "tenant-name": null, "access-details": "Key: 12, Gate code: 789" },
    []],
  ["lockbox picked out of mixed access details",
    { "tenant-name": null, "access-details": "Key: 12, Lockbox code: 7, Gate code: 789" },
    ["Lockbox code: 7"]],

  ["nothing at all", { "tenant-name": null, "access-details": null }, []],
  ["empty object", {}, []],
];

const failures = [];
for (const [name, input, expected] of CASES) {
  const got = schedulingAccessLines(input);
  if (JSON.stringify(got) !== JSON.stringify(expected)) {
    failures.push(`  ${name}\n    expected: ${JSON.stringify(expected)}\n    got:      ${JSON.stringify(got)}`);
  }
}

console.log(`scheduling access: ${CASES.length - failures.length}/${CASES.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n${failures.join("\n")}`);
  process.exit(1);
}

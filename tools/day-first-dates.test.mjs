#!/usr/bin/env node
// dayFirstDates(): ISO dates the model copies into a task description become DD/MM/YYYY
// (job 110885: "certificate expiry date 2026-11-11"). Lifted from index.js like the other
// tools/*.test.mjs files, because importing index.js starts the server.
import fs from "node:fs";
import path from "node:path";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");
const i = src.indexOf("function dayFirstDates(");
const j = src.indexOf("\n}\n", i);
const dayFirstDates = new Function(`${src.slice(i, j + 3)}; return dayFirstDates;`)();

assert.equal(
  dayFirstDates("Annual electrical compliance check — schedule after current certificate expiry date 2026-11-11."),
  "Annual electrical compliance check — schedule after current certificate expiry date 11/11/2026.",
);
assert.equal(dayFirstDates("Between 2026-01-05 and 2026-12-31"), "Between 05/01/2026 and 31/12/2026");
assert.equal(dayFirstDates("Already 11/11/2026"), "Already 11/11/2026");
assert.equal(dayFirstDates("Order 2026-11-1123 and part 2026-13-01"), "Order 2026-11-1123 and part 2026-13-01");
console.log("day-first-dates: ok");

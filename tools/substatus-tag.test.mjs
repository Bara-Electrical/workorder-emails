#!/usr/bin/env node
// Regression tests for Outlook-category → Aroflo substatus resolution — run with npm test.
//
// Lifts SUBSTATUS_TAG_MAP and the two resolution lines out of index.js's source rather than
// importing the module; see client-match.test.mjs's header for why importing boots the app.
//
// The point of these tests: the map is keyed on the LOWERCASED category text, so a key with
// a capital letter or stray space is silently dead — the tag is ignored and the job falls
// back to the task-type default. That is exactly how "Weather Dependant" ended up as PPM1
// on job 107481, and nothing would have caught it.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src = fs.readFileSync(path.join(here, "..", "index.js"), "utf8");

function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf(endMarker, i);
  if (j === -1) throw new Error(`no end marker "${endMarker}" after: ${startMarker}`);
  return src.slice(i, j + endMarker.length);
}

const mapSrc = grab("const SUBSTATUS_TAG_MAP = {", "\n};\n");
const resolveSrc = grab("const categoriesLower = new Set(", "?.[1] || null;");

// Runs the real resolution lines against a message, exactly as processMessage does.
const resolve = new Function("message", `${mapSrc}\n${resolveSrc}\nreturn substatusTagId;`);

// Every id the map can return, for readable failure output.
const ID_NAMES = Object.fromEntries(
  [...mapSrc.matchAll(/"([^"]+)":\s*"([^"]+)",\s*\/\/\s*(.+)/g)].map(m => [m[2], m[3].trim()])
);
const name = id => (id ? `${id} (${ID_NAMES[id] ?? "?"})` : "null");

const WEATHER = "IydKLyIK";
const AWAITING = "Iyc6LycK";
const URGENT = "Iyc6UyMK";

const CASES = [
  // Job 107481: the tag that was being ignored.
  [["Weather Dependant"],              WEATHER],
  // The exact category string in use, per the mailbox: "Weather Dependant", 36 uses.
  [["Bara AI", "Weather Dependant"],   WEATHER],
  // Casing and padding must not matter — the lookup lowercases and trims.
  [["weather dependant"],              WEATHER],
  [["WEATHER DEPENDANT"],              WEATHER],
  [["  Weather Dependant  "],          WEATHER],
  // Aroflo's own spelling, in case a category is ever created to match it.
  [["Weather Dependent"],              WEATHER],
  // First match wins: urgent work stays urgent even when also weather dependent.
  [["Weather Dependant", "Urgent"],    URGENT],
  // "Send To Awaiting Confirmation" — the category actually in use on workorders@.
  [["Send To Awaiting Confirmation"],  AWAITING],
  [["send to awaiting confirmation"],  AWAITING],
  [["Bara AI", "Send To Awaiting Confirmation"], AWAITING],
  // Urgent still wins over it, same as weather.
  [["Send To Awaiting Confirmation", "Urgent"], URGENT],
  // The retired "Waiting confirmation" category is deliberately unmapped.
  [["Waiting confirmation"],           null],
  // Untagged and unknown tags fall through to the task-type default.
  [[],                                 null],
  [["Bara AI"],                        null],
  [["Weather"],                        null],
];

let pass = 0;
const failures = [];
for (const [categories, expected] of CASES) {
  const got = resolve({ categories });
  if (got === expected) pass++;
  else failures.push(`  ${JSON.stringify(categories)}\n    expected: ${name(expected)}\n    got:      ${name(got)}`);
}

// Every key must already be lowercase and trimmed, or it can never match.
for (const [, key] of mapSrc.matchAll(/^\s*"([^"]+)":/gm)) {
  if (key === key.trim().toLowerCase()) pass++;
  else failures.push(`  SUBSTATUS_TAG_MAP key ${JSON.stringify(key)} is not lowercase/trimmed — it can never match`);
}

const total = CASES.length + [...mapSrc.matchAll(/^\s*"([^"]+)":/gm)].length;
console.log(`substatus tags: ${pass}/${total} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n${failures.join("\n")}`);
  process.exit(1);
}

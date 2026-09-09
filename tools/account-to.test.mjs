#!/usr/bin/env node
// Regression tests for the account-to placeholder strip — run with: npm test
//
// Lifts the normalisation out of index.js's source rather than importing the module, for
// the same reason as client-match.test.mjs: importing index.js boots the server and the
// email poll loop. See that file's header for the full rationale.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src = fs.readFileSync(path.join(here, "..", "index.js"), "utf8");

const marker = 'if (parsed["account-to"]) {';
const start = src.indexOf(marker);
if (start === -1) throw new Error(`index.js no longer contains: ${marker}`);
const end = src.indexOf("\n  }\n", start);
if (end === -1) throw new Error("could not find the end of the account-to block");
const block = src.slice(start, end + 4);

// Run the real block against a parsed object, exactly as processMessage does.
function normalise(accountTo) {
  const parsed = { "account-to": accountTo };
  const console = { log() {} }; // silence the block's logging
  new Function("parsed", "console", block)(parsed, console);
  return parsed["account-to"];
}

const CASES = [
  // The reported bug: no owner named, so the agency stands alone.
  ["Owners c/o Smith Realty",              "Smith Realty"],
  ["owners c/o Smith Realty",              "Smith Realty"],
  ["Owner c/o Smith Realty",               "Smith Realty"],
  ["The Owners c/o Smith Realty",          "Smith Realty"],
  ["The Owner c/o Rightly.Realestate",     "Rightly.Realestate"],
  ["Owners c/- Smith Realty",              "Smith Realty"],
  ["Owner / Owners c/o Smith Realty",      "Smith Realty"],
  ["OWNERS C/O SMITH REALTY",              "SMITH REALTY"],
  // A real owner name must survive untouched — this is the whole point of the field.
  ["Tony Antoniou c/o Realmark North Coastal", "Tony Antoniou c/o Realmark North Coastal"],
  ["Helen Jackson, Ian Jackson c/o Canquan Pty Ltd", "Helen Jackson, Ian Jackson c/o Canquan Pty Ltd"],
  ["Delstrat Pty Ltd c/o BOSS Real Estate",     "Delstrat Pty Ltd c/o BOSS Real Estate"],
  // Someone whose name merely starts with "Owen" is not the word "owner".
  ["Owen Smith c/o Smith Realty",          "Owen Smith c/o Smith Realty"],
  // Already correct, or no c/o at all — left alone.
  ["Smith Realty",                         "Smith Realty"],
  ["Owners Corporation Strata Services",   "Owners Corporation Strata Services"],
];

let pass = 0;
const failures = [];
for (const [input, expected] of CASES) {
  const got = normalise(input);
  if (got === expected) pass++;
  else failures.push(`  ${JSON.stringify(input)}\n    expected: ${expected}\n    got:      ${got}`);
}

console.log(`account-to: ${pass}/${CASES.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n${failures.join("\n")}`);
  process.exit(1);
}

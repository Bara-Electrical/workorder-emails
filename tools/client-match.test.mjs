#!/usr/bin/env node
// Regression tests for findClient() — run with: npm test
//
// index.js is a single module that starts an HTTP server and an email poll loop the
// moment it is imported, so importing it here would boot the app. Instead the three
// declarations under test are lifted out of the source text and evaluated in isolation.
// That keeps production code untouched while still testing the shipped source rather
// than a hand-copied paraphrase that could drift away from it.
//
// If index.js is refactored and a marker below no longer matches, this fails loudly with
// the marker it could not find — that is a prompt to update the marker, never to weaken
// the assertions.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const indexPath = path.join(here, "..", "index.js");
const src = fs.readFileSync(indexPath, "utf8");

function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf(endMarker, i);
  if (j === -1) throw new Error(`no end marker "${endMarker}" after: ${startMarker}`);
  return src.slice(i, j);
}

const extracted = [
  grab("function normaliseClientName(name)", "\n}\n") + "\n}\n",
  grab("function clientNameForms(realEstateName)", "\n}\n") + "\n}\n",
  grab("function clientNameCandidates(realEstateName)", "\n}\n") + "\n}\n",
  grab("async function findClient(realEstateName)", "\n// Full state names"),
].join("\n");

// The cache maps and the API fallback are module-level in index.js; stub them so the
// cache path under test runs exactly as it does in production. arofloGet returning no
// clients means a test that reaches the live-API fallback fails rather than silently
// passing through it.
const module_ = `
const clientCache = new Map();
const clientCacheNormalised = new Map();
const toArray = x => (Array.isArray(x) ? x : [x]);
const arofloGet = async () => ({ clients: [] });
${extracted}
export { findClient, clientCache, clientCacheNormalised, normaliseClientName };
`;
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "client-match-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const m = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

// Real Aroflo client names, chosen to include the duplicate and junk records that make
// name normalisation risky — not just the happy path.
const CLIENTS = [
  "Rightly.Realestate",
  "L.J. Hooker Leederville/Kwinana",
  "LJ Hooker Victoria Park - Belmont",
  "A.D. Engineering International",
  "Michaelkeil.com",
  "R & P.M Buswell",
  "Rentwest Solutions",
  "Driven Property Group",
  "Professionals Armadale Real Estate WA",
  "Robin  Wright", "Robin Wright",              // duplicate owner records, double space
  "S Class Property Group", "S-Class Property Group",
  "cam .", "Cameron Burchell", "Cameron  Best", // junk single-name records
  "emma .", "Emma Smith",
  "Steven  .", "SDRE Steven Davis Real Estate", // work order 3378: junk card vs real agency
  "Scott .", "Regina .",
  "Pro Property Group Real Estate", "ProProperty Group",
];

for (const name of CLIENTS) m.clientCache.set(name.toLowerCase(), { clientname: name, clientid: name });
m.clientCacheNormalised.clear();
for (const c of m.clientCache.values()) {
  const key = m.normaliseClientName(c.clientname);
  if (!key) continue;
  const bucket = m.clientCacheNormalised.get(key);
  if (bucket) bucket.push(c); else m.clientCacheNormalised.set(key, [c]);
}

const CASES = [
  // The reported bug: Aroflo has "Rightly.Realestate", work orders say "Rightly Realestate".
  ["Rightly Realestate",            "Rightly.Realestate"],
  ["Rightly Real Estate",           "Rightly.Realestate"],
  ["Rightly.Realestate",            "Rightly.Realestate"],
  ["RIGHTLY REALESTATE",            "Rightly.Realestate"],
  ["Rightly Realestate Pty Ltd",    "Rightly.Realestate"],
  // Same shape, other live clients.
  ["AD Engineering International",  "A.D. Engineering International"],
  ["R & PM Buswell",                "R & P.M Buswell"],
  // Must not regress: exact and starts-with matching that already worked.
  ["Rentwest Solutions",            "Rentwest Solutions"],
  ["Driven Property Group Pty Ltd", "Driven Property Group"],
  ["LJ Hooker Victoria Park",       "LJ Hooker Victoria Park - Belmont"],
  // Must not regress: names colliding with duplicate/junk records under normalisation.
  ["Cameron Burchell",              "Cameron Burchell"],
  ["Emma Smith",                    "Emma Smith"],
  ["Robin Wright",                  "Robin Wright"],
  ["S Class Property Group",        "S Class Property Group"],
  // Genuinely ambiguous — must stay unmatched rather than pick one of the two.
  ["SClass Property Group",         null],
  // Work order 3378. The agency's real card is "SDRE Steven Davis Real Estate", which no
  // tier can reach from "Steven Davis Real Estate" — so the correct result is NO match,
  // surfaced as "Client not found" for an admin. It must never fall back to the junk
  // single-name card "Steven  ." via the derived first-word candidate "Steven".
  ["Steven Davis Real Estate",      null],
  ["Scott Palmer Realty",           null],
  ["Regina Property Group",         null],
  // A genuinely single-word agency name is still a whole name, so it may normalise-match.
  ["Steven",                        "Steven  ."],
];

let pass = 0;
const failures = [];
for (const [input, expected] of CASES) {
  const got = (await m.findClient(input))?.clientname ?? null;
  if (got === expected) pass++;
  else failures.push(`  ${JSON.stringify(input)}\n    expected: ${expected}\n    got:      ${got}`);
}

console.log(`findClient: ${pass}/${CASES.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n${failures.join("\n")}`);
  process.exit(1);
}

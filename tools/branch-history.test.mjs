#!/usr/bin/env node
// Regression tests for pickBranchFromHistory() — run with: npm test
//
// When a work order names more than one branch, the agency's own records break the tie: the
// branch card that already holds the property, else the one whose contacts include the PM.
// Same harness as the other tools/*.test.mjs files: the declarations under test are lifted
// out of index.js and evaluated in isolation, since importing it starts the service.
//
// Booking against the wrong office is worse than not booking, so these are as much about
// what must still decline as what resolves.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src  = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");

function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf(endMarker, i);
  if (j === -1) throw new Error(`no end marker "${endMarker}" after: ${startMarker}`);
  return src.slice(i, j + endMarker.length);
}

const module_ = [
  grab("const STATE_NAMES = {", "\n};\n"),
  grab("const STATE_PATTERN", ";\n"),
  grab("function parseAustralianAddress(address)", "\n}\n"),
  grab("function matchContact(contacts, pmName)", "\n}\n"),
  grab("function matchLocation(locations, address)", "\n}\n"),
  grab("function pickBranchFromHistory(candidates, address, pmName)", "\n}\n"),
  "export { pickBranchFromHistory, matchLocation };",
].join("\n");

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "branch-history-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { pickBranchFromHistory, matchLocation } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

const loc = (locationname, suburb, archived = "false") => ({ locationname, suburb, archived });
const pm  = (givennames, surname, archived = "false") => ({ givennames, surname, archived });
const card = (name, locations = [], contacts = []) => ({ name, locations, contacts });

// The address as the AI extracts it from RMA's work order.
const SPEARWOOD = "284 Hamilton Road, Spearwood WA 6163";

// ---- 1 Oct: 284 Hamilton Road, Spearwood. RMA listed both office addresses; the property's
// earlier job sits on the Osborne Park card. ----
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("284 Hamilton Road", "Spearwood"), loc("47 Fishbone Turn", "Banksia Grove")]),
    card("RMA - Port Kennedy", [loc("16 Hoskin Way", "Baldivis")]),
  ], SPEARWOOD, "Tasha Sharman");
  check("Spearwood resolves to the card that holds the property", r?.name === "RMA - Osborne Park", JSON.stringify(r));
  check("…and says why", r?.how === "property already on this card", JSON.stringify(r));
}

// Order of the candidates must not matter.
{
  const r = pickBranchFromHistory([
    card("RMA - Port Kennedy", [loc("16 Hoskin Way", "Baldivis")]),
    card("RMA - Osborne Park", [loc("284 Hamilton Road", "Spearwood")]),
  ], SPEARWOOD, null);
  check("candidate order does not change the answer", r?.name === "RMA - Osborne Park", JSON.stringify(r));
}

// ---- Must still decline ----
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("284 Hamilton Road", "Spearwood")]),
    card("RMA - Port Kennedy", [loc("284 Hamilton Road", "Spearwood")]),
  ], SPEARWOOD, null);
  check("property on both cards declines", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("284 Hamilton Road", "Spearwood")], [pm("Tasha", "Sharman")]),
    card("RMA - Port Kennedy", [loc("284 Hamilton Road", "Spearwood")]),
  ], SPEARWOOD, "Tasha Sharman");
  check("property on both cards declines even when the PM is on one", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("12 Other Street", "Spearwood")]),
    card("RMA - Port Kennedy", [loc("16 Hoskin Way", "Baldivis")]),
  ], SPEARWOOD, null);
  check("a new property with no PM declines", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("284 Hamilton Road", "Spearwood", "true")]),
    card("RMA - Port Kennedy"),
  ], SPEARWOOD, null);
  check("an archived location is not evidence", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("2/284 Hamilton Road", "Spearwood")]),
    card("RMA - Port Kennedy"),
  ], SPEARWOOD, null);
  check("a different unit at the same number is not the property", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("1284 Hamilton Road", "Spearwood")]),
    card("RMA - Port Kennedy"),
  ], SPEARWOOD, null);
  check("a different street number is not the property", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [loc("284 Hamilton Road", "Coogee")]),
    card("RMA - Port Kennedy"),
  ], SPEARWOOD, null);
  check("the same street in another suburb is not the property", r === null, JSON.stringify(r));
}

// ---- Falls back to the PM when no card holds the property ----
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [], [pm("Tasha", "Sharman")]),
    card("RMA - Port Kennedy", [], [pm("Rhianna", "Mckenzie")]),
  ], SPEARWOOD, "Tasha Sharman");
  check("new property resolves by the PM's card", r?.name === "RMA - Osborne Park", JSON.stringify(r));
  check("…and says why", r?.how === "PM is a contact on this card", JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [], [pm("Tasha", "Sharman")]),
    card("RMA - Port Kennedy", [], [pm("Tasha", "Sharman")]),
  ], SPEARWOOD, "Tasha Sharman");
  check("PM on both cards declines", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [], [pm("Tasha", "Sharman", "true")]),
    card("RMA - Port Kennedy"),
  ], SPEARWOOD, "Tasha Sharman");
  check("an archived PM contact is not evidence", r === null, JSON.stringify(r));
}
{
  const r = pickBranchFromHistory([
    card("RMA - Osborne Park", [], [pm("Tasha", "Sharman")]),
    card("RMA - Port Kennedy"),
  ], SPEARWOOD, null);
  check("no PM extracted means no PM tie-break", r === null, JSON.stringify(r));
}

// ---- matchLocation is the same rule findOrUpdateLocation applies ----
{
  const locs = [loc("284 Hamilton Road", "Spearwood")];
  check("matchLocation finds the property", matchLocation(locs, SPEARWOOD)?.locationname === "284 Hamilton Road");
  check("matchLocation with no address is null", matchLocation(locs, "") === null);
  check("matchLocation with no locations is null", matchLocation([], SPEARWOOD) === null);
}

const failures = results.filter(r => !r.ok);
console.log(`branch history: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

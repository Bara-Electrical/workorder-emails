#!/usr/bin/env node
// Regression tests for resolveBranch() — run with: npm test
//
// index.js starts an HTTP server and the poll loop the moment it is imported, so the
// declarations under test are lifted out of the source text and evaluated in isolation.
// Same approach as the other tools/*.test.mjs files: it tests the shipped source rather than
// a hand-copied paraphrase that could drift away from it.
//
// Some agencies keep one Aroflo card per branch while the work order names only the group,
// so the extracted name reaches every branch or none. The branch is in the email's own text
// and is read back out of it. Booking against the wrong branch is worse than not booking, so
// these cases are as much about what it must REFUSE to resolve as what it resolves.

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
  grab("const BRANCH_MAPS = [", "\n];\n"),
  grab("const BRANCH_PROXIMITY", "= 40;"),
  grab("function resolveBranch(realEstate, rawEmail)", "\n}\n"),
  "export { resolveBranch, BRANCH_MAPS, BRANCH_PROXIMITY };",
].join("\n");

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "branch-resolve-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const m = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });
const resolved = (name, extracted, email, expected) => {
  const r = m.resolveBranch(extracted, email);
  check(name, r?.name === expected, `got ${JSON.stringify(r)}`);
};
const declined = (name, extracted, email) => {
  const r = m.resolveBranch(extracted, email);
  check(name, r !== null && !r.name, `got ${JSON.stringify(r)}`);
};

// ---- Real emails from the 23 Sep client-not-found alerts. ----

// Austpro signs off with the branch. Extracted name is the bare group.
resolved(
  "Austpro resolves from its signature",
  "Austpro Properties",
  "Ada sent a message Hi Bara, The owner would like to carry out the AC service before the upcoming summer. Can you please assist arrange service with the tenants? View issue report Regards, Austpro Properties Booragoon Do Not Share This Email",
  "Austpro Properties - Booragoon"
);
resolved(
  "Austpro resolves on its other branch too",
  "Austpro Properties",
  "Regards, Austpro Properties South Perth",
  "Austpro Properties - South Perth"
);

// Bellcourt names the branch in the subject line.
resolved(
  "Bellcourt resolves from the subject",
  "Bellcourt Property",
  "Work order: Hi Work, Bellcourt Mount Lawley 1 has a job request. Anish Patel from Bellcourt Mount Lawley 1 has sent you a work order to complete a job at 10 Cantle Street, Perth, WA, 6000.",
  "Bellcourt Property Group Mount Lawley"
);

// The case proximity exists for: a Mount Lawley work order whose PROPERTY is in South Perth.
// A bare suburb search matches both and picks whichever key comes first — here that is
// "mount lawley", right by luck, so the test uses the order that would get it wrong.
resolved(
  "the branch beside the agency name beats a property suburb elsewhere",
  "Bellcourt Property",
  "Bellcourt Mount Lawley 1 has sent you a work order for 12 Angelo Street, South Perth WA 6151.",
  "Bellcourt Property Group Mount Lawley"
);
resolved(
  "and the same the other way round",
  "Bellcourt Property",
  "Bellcourt South Perth has sent you a work order for 5 Second Avenue, Mount Lawley WA 6050.",
  "Bellcourt Property Group South Perth"
);

// RMA must keep working: its branch comes from the office address on the work order, which is
// not always within reach of the agency name, so the distant fallback still has to fire.
resolved(
  "RMA still resolves from the work order's branch address",
  "Rental Management Australia (WA)",
  "Congratulations on being selected for the job. Account to: The Owners C/O Rental Management Australia (WA), 17 Drake St, Osborne Park WA 6017",
  "RMA - Osborne Park"
);
resolved(
  "RMA resolves a branch far from its name",
  "Rental Management Australia",
  "Rental Management Australia\n" + "filler ".repeat(40) + "\n7 Sunlight Dr, Port Kennedy WA 6172",
  "RMA - Port Kennedy"
);

// ---- What it must refuse. ----

declined("no branch named anywhere declines", "Austpro Properties", "Please attend the property. Regards, Ada");
declined("two branches equally far declines", "Austpro Properties", "Property at 3 Booragoon Ave; owner lives in South Perth.");
declined("two branches both beside the name declines", "Bellcourt Property", "Bellcourt Mount Lawley and Bellcourt Shenton Park both manage this.");
declined("an empty email declines", "Austpro Properties", "");
declined("a null email declines", "Austpro Properties", null);

{
  const r = m.resolveBranch("Austpro Properties", "3 Booragoon Ave; owner in South Perth.");
  check("an ambiguous result names the candidates", Array.isArray(r.ambiguous) && r.ambiguous.length === 2, JSON.stringify(r));
}

// An agency with a single Aroflo card must not be touched by any of this.
{
  check("a non-branch agency returns null", m.resolveBranch("Bourkes", "Bourkes, South Perth WA") === null);
  check("an empty name returns null", m.resolveBranch("", "anything") === null);
  check("a null name returns null", m.resolveBranch(null, "anything") === null);
}

// Proximity must stay wide enough for the real signatures and far short of an address
// elsewhere in a work order.
{
  check("proximity spans a real signature", m.BRANCH_PROXIMITY >= "  Properties  ".length);
  const far = "Bellcourt" + " ".repeat(m.BRANCH_PROXIMITY + 10) + "south perth";
  const r = m.resolveBranch("Bellcourt Property", far + " and mount lawley");
  check("proximity does not reach past its limit", !r?.how || r.how === "elsewhere in the email", JSON.stringify(r));
}

// The table itself: a branch whose card name is wrong is invisible until a job lands on the
// wrong client, so assert the shape rather than trusting it by eye.
{
  check("every entry has branches", m.BRANCH_MAPS.every(b => Object.keys(b.branches || {}).length >= 2));
  check("every suburb key is lowercase and trimmed", m.BRANCH_MAPS.every(b => Object.keys(b.branches).every(k => k === k.toLowerCase().trim())));
  check("every entry has agency and near patterns", m.BRANCH_MAPS.every(b => b.agency instanceof RegExp && b.near instanceof RegExp));
  check("every agency pattern matches its own card names", m.BRANCH_MAPS.every(b => Object.values(b.branches).some(v => b.agency.test(v)) || b.near.test(Object.values(b.branches)[0])));
}

const failures = results.filter(r => !r.ok);
console.log(`branch resolve: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

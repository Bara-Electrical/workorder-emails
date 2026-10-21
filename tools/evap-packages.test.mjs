#!/usr/bin/env node
// Regression tests for the evap packages — run with: npm test
//
// From 1 Oct 2026 an evap service is EV1 and an evap service with an electrical compliance
// check is EVEC1; AC1/ACEC1 cover reverse-cycle units only. Same harness as the other
// tools/*.test.mjs files: the functions are lifted out of index.js so the tests run what ships.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, "..");
const src  = fs.readFileSync(path.join(root, "index.js"), "utf8").replace(/\r\n/g, "\n");
function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  return src.slice(i, src.indexOf(endMarker, i) + endMarker.length);
}
const module_ = [
  grab("function pickEvapTaskType(taskTypes, pkg)", "\n}\n"),
  grab("const EVAP_TASK_TYPES", ";\n"),
  grab("const EVAP_PACKAGE_FOR", ";\n"),
  grab("function applyEvapTag(result, airconUnitType)", "\n}\n"),
  grab("const SUBSTATUS_MAP", "\n};\n"),
  grab("function arofloRateLimit(data)", "\n}\n"),
  // arofloTaskTypes with Aroflo stubbed out, so a failed refresh can be played.
  "const toArray = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]);",
  "let getImpl = null; const arofloGet = (q) => getImpl(q);",
  grab("const TASK_TYPES_TTL_MS", ";\n"),
  grab("let taskTypesCache", ";\n"),
  grab("async function arofloTaskTypes()", "\n}\n"),
  "const stubAroflo = (f) => { getImpl = f; }; const expireTaskTypes = () => { taskTypesCache.at = 0; };",
  "export { pickEvapTaskType, applyEvapTag, SUBSTATUS_MAP, arofloRateLimit, arofloTaskTypes, stubAroflo, expireTaskTypes };",
].join("\n");
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "evap-packages-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { pickEvapTaskType, applyEvapTag, SUBSTATUS_MAP, arofloRateLimit, arofloTaskTypes, stubAroflo, expireTaskTypes } = await import(pathToFileURL(tmp).href);
const { PACKAGE_TEMPLATES } = await import(pathToFileURL(path.join(root, "templates.js")).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

// ---- Finding the Aroflo task type by name ----
const TYPES = [
  { tasktypeid: "AC1", tasktype: "$200+ Standard Air Con Service", archived: "false" },
  { tasktypeid: "ACEC1", tasktype: "$290+ Air Con & Electrical Compliance", archived: "false" },
  { tasktypeid: "OLDEV", tasktype: "$280 Evap Service", archived: "true" },
  { tasktypeid: "PRIVEV", tasktype: "Private - $300+ Evap Service", archived: "false" },
  { tasktypeid: "STOREY", tasktype: "$500 2 Story Evap Service", archived: "false" },
  { tasktypeid: "EV1", tasktype: "$300+ Evap Service", archived: "false" },
  { tasktypeid: "EVEC1", tasktype: "$390+ Evap & Electrical Compliance.", archived: "false" },
];
check("EV1 is the live real-estate evap service", pickEvapTaskType(TYPES, "EV1") === "EV1", String(pickEvapTaskType(TYPES, "EV1")));
check("EVEC1 is the evap & compliance type", pickEvapTaskType(TYPES, "EVEC1") === "EVEC1", String(pickEvapTaskType(TYPES, "EVEC1")));
check("an EVEC1-only list gives no EV1", pickEvapTaskType(TYPES.filter(t => t.tasktypeid !== "EV1"), "EV1") === null);
check("never archived, private or 2 storey", pickEvapTaskType(TYPES.filter(t => ["OLDEV", "PRIVEV", "STOREY"].includes(t.tasktypeid)), "EV1") === null);
check("'Evaporative' in the name counts", pickEvapTaskType([{ tasktypeid: "X", tasktype: "Evaporative Aircon Service", archived: "false" }], "EV1") === "X");
check("nothing for other packages", pickEvapTaskType(TYPES, "AC1") === null && pickEvapTaskType(null, "EV1") === null);

// ---- The office's Evap tag ----
{
  const r = { "task-type": "AC1", "package": "AC1" };
  check("Evap-only tag turns AC1 into EV1", applyEvapTag(r, "Evap") && r["task-type"] === "EV1" && r.package === "EV1", JSON.stringify(r));
}
{
  const r = { "task-type": "ACEC1", "package": "ACEC1" };
  applyEvapTag(r, "Evap");
  check("Evap-only tag turns ACEC1 into EVEC1", r["task-type"] === "EVEC1" && r.package === "EVEC1", JSON.stringify(r));
}
{
  const r = { "task-type": "Real Estate General Maintenance", "package": "ACEC1" };
  applyEvapTag(r, "Evap");
  check("package changes, a maintenance task type stays", r["task-type"] === "Real Estate General Maintenance" && r.package === "EVEC1", JSON.stringify(r));
}
{
  const r = { "task-type": "AC1", "package": "AC1" };
  check("a split tagged as well leaves AC1 alone", applyEvapTag(r, "Split System, Evap") === null && r.package === "AC1");
  check("no tag leaves AC1 alone", applyEvapTag(r, null) === null && r.package === "AC1");
}
{
  const r = { "task-type": "AC2", "package": "AC2" };
  check("AC2 is never switched", applyEvapTag(r, "Evap") === null && r.package === "AC2");
}

// ---- Templates and substatus ----
for (const pkg of ["EV1", "EVEC1"]) {
  check(`${pkg} has a template naming it`, PACKAGE_TEMPLATES[pkg]?.includes(`${pkg}.`));
  check(`${pkg} is Ready to schedule PPM1`, SUBSTATUS_MAP[pkg] === SUBSTATUS_MAP.AC1);
}
check("EVEC1 carries the compliance check", PACKAGE_TEMPLATES.EVEC1.includes("Smoke Alarm Testing") && PACKAGE_TEMPLATES.EVEC1.includes("Evap Service."));
for (const pkg of ["AC1", "ACEC1"]) {
  check(`${pkg} no longer covers an evap`, !/evaporative/i.test(PACKAGE_TEMPLATES[pkg]));
}

// ---- Aroflo's rate-limit replies are retried, other failures aren't ----
check("per-minute limit (status 6) is retried", arofloRateLimit({ status: "6", statusmessage: "Exceeded Rate Limit - Requests per Minute" }) === "minute");
check("per-second limit is retried", arofloRateLimit({ status: "11", statusmessage: "Exceeded Rate Limit - Requests per Second" }) === "second");
check("daily limit isn't retried", arofloRateLimit({ status: "7", statusmessage: "Exceeded Rate Limit - Daily" }) === null);
check("size limit isn't retried", arofloRateLimit({ status: "8", statusmessage: "Exceeded Size Rate Limit - Please try filtering your request down" }) === null);
check("a good reply isn't a rate limit", arofloRateLimit({ status: "0", statusmessage: "Login OK" }) === null);

// ---- A failed task type refresh keeps the last list ----
{
  let threw = null;
  stubAroflo(async () => { throw new Error("Aroflo GET failed: Exceeded Rate Limit - Requests per Minute"); });
  try { await arofloTaskTypes(); } catch (err) { threw = err; }
  check("with no list yet, a failed read still fails", threw && /Rate Limit/.test(threw.message));

  stubAroflo(async () => ({ tasktypes: TYPES }));
  check("a good read gives the list", (await arofloTaskTypes()).length === TYPES.length);

  expireTaskTypes();
  stubAroflo(async () => { throw new Error("Aroflo GET failed: Exceeded Rate Limit - Requests per Minute"); });
  const warn = console.warn; console.warn = () => {};
  let stale; try { stale = await arofloTaskTypes(); } finally { console.warn = warn; }
  check("a failed refresh keeps the last list, so EVEC1 is still found", pickEvapTaskType(stale, "EVEC1") === "EVEC1");
}

const failures = results.filter(r => !r.ok);
console.log(`evap packages: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

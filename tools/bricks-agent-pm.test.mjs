#!/usr/bin/env node
// Regression tests for bricksAgentIds() / bricksAgentAssignedPm() — run with: npm test
//
// Bricks + Agent work orders link to the job, and the job's record names the assigned PM. Same
// harness as the other tools/*.test.mjs files; fetch is stubbed, so nothing leaves the machine.

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
  grab("const BRICKS_AGENT_LINK", ";\n"),
  grab("function bricksAgentIds(html)", "\n}\n"),
  grab("async function bricksAgentAssignedPm(html)", "\n}\n"),
  "export { bricksAgentIds, bricksAgentAssignedPm };",
].join("\n");
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "bricks-agent-pm-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { bricksAgentIds, bricksAgentAssignedPm } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

// Verbatim href from the 30 Prospector Loop email (job 108486), Safe Links-wrapped.
const PROSPECTOR = '<a href="https://aus01.safelinks.protection.outlook.com/?url=https%3A%2F%2Ftrade.bricksandagent.com%2Fexternal%2Foverview%2F47255f4a-174b-45bf-a861-3d0e7744639d%2F425fc1ac-c147-4180-da70-08df1aac8363&amp;data=05%7C02%7Cadmin%40baraelectrical.com.au" style="display:inline-block">Review Work Order </a>';
// Verbatim from the Peak Central email (16 Kumarina Drive), with a query string after the job id.
const PEAK_SCHEDULE = 'https%3A%2F%2Ftrade.bricksandagent.com%2Fexternal%2Foverview%2Fb386e634-5afe-4952-a6d3-09da98096cec%2Fb223e1b3-c75f-4cdd-22f2-08df1f805407%3FisSchedule%3D1';

{
  const ids = bricksAgentIds(PROSPECTOR);
  check("Safe Links-encoded link: user id", ids?.userId === "47255f4a-174b-45bf-a861-3d0e7744639d", JSON.stringify(ids));
  check("Safe Links-encoded link: job id", ids?.jobId === "425fc1ac-c147-4180-da70-08df1aac8363", JSON.stringify(ids));
  check("a trailing query string is not part of the id", bricksAgentIds(PEAK_SCHEDULE)?.jobId === "b223e1b3-c75f-4cdd-22f2-08df1f805407");
  const plain = bricksAgentIds("see https://trade.bricksandagent.com/external/overview/47255f4a-174b-45bf-a861-3d0e7744639d/425fc1ac-c147-4180-da70-08df1aac8363");
  check("an unwrapped link reads the same", plain?.jobId === "425fc1ac-c147-4180-da70-08df1aac8363");
  check("no link, no ids", bricksAgentIds("<p>PropertyMe work order</p>") === null);
  check("null is safe", bricksAgentIds(null) === null);
}

// ---- The lookup, against a stubbed API ----
const calls = [];
const respond = (status, body) => async (url) => { calls.push(url); return { ok: status >= 200 && status < 300, status, json: async () => body }; };
const origWarn = console.warn; console.warn = () => {};
try {
  globalThis.fetch = respond(200, { job: { assignedPmName: "Alyssa Radmore-Collard ", assignedPmAddress: "pm9@professionalstwt.com.au" } });
  check("assigned PM is returned, trimmed", await bricksAgentAssignedPm(PROSPECTOR) === "Alyssa Radmore-Collard");
  check("…from the jobdetails call with both ids",
    calls.at(-1) === "https://services.bricksandagent.com/external/jobdetails?jobId=425fc1ac-c147-4180-da70-08df1aac8363&userId=47255f4a-174b-45bf-a861-3d0e7744639d", calls.at(-1));

  calls.length = 0;
  check("no link means no call", await bricksAgentAssignedPm("<p>no link</p>") === null && calls.length === 0);

  globalThis.fetch = respond(404, { job: { assignedPmName: "From An Error Page" } });
  check("an error status leaves the AI's answer (null)", await bricksAgentAssignedPm(PROSPECTOR) === null);

  globalThis.fetch = respond(200, { job: { assignedPmName: "   " } });
  check("a blank name is null", await bricksAgentAssignedPm(PROSPECTOR) === null);

  globalThis.fetch = respond(200, { job: {} });
  check("a missing name is null", await bricksAgentAssignedPm(PROSPECTOR) === null);

  globalThis.fetch = async () => { throw new Error("network down"); };
  check("a network failure is null, not a throw", await bricksAgentAssignedPm(PROSPECTOR) === null);
} finally {
  console.warn = origWarn;
}

const failures = results.filter(r => !r.ok);
console.log(`bricks agent pm: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

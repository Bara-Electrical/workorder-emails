#!/usr/bin/env node
// Regression tests for thread deduplication — run with npm test.
//
// Guards the bug that created job 107496 as a duplicate of 106941: findJobTagInThread
// returned null both when a thread genuinely had no job AND when the Graph lookup itself
// failed, and the caller reads null as "safe to create a job". Graph intermittently
// rejects the conversationId $filter, so an infrastructure blip silently became a
// duplicate job.
//
// Lifts the real function out of index.js's source rather than importing the module; see
// client-match.test.mjs's header for why importing boots the app.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n"); // a Windows checkout has CRLF; the markers below are LF

function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf(endMarker, i);
  if (j === -1) throw new Error(`no end marker "${endMarker}" after: ${startMarker}`);
  return src.slice(i, j + endMarker.length);
}

const sentinelSrc = grab("const THREAD_LOOKUP_FAILED = Symbol(", ";");
const fnSrc = grab("async function findJobTagInThread(", "\n}\n");
const tagFnSrc = grab("async function tagWholeConversation(", "\n}\n");
// STATUS_CATEGORIES is built from the individual category constants, so take the whole
// block through to the end of the array rather than the array alone.
const statusCatsSrc = grab("const TRIGGER_CATEGORY", "];");

// graphFetch is stubbed per-case so each Graph outcome can be exercised exactly as the
// real function sees it.
function build(graphFetch) {
  return new Function("graphFetch", "console", `
    ${sentinelSrc}
    ${fnSrc}
    return { findJobTagInThread, THREAD_LOOKUP_FAILED };
  `)(graphFetch, { warn() {}, log() {} });
}

const ok = (messages) => async () => ({ ok: true, status: 200, json: async () => ({ value: messages }) });
const httpError = (status) => async () => ({ ok: false, status, json: async () => ({ error: { code: "InefficientFilter" } }) });
const throws = () => async () => { throw new Error("socket hang up"); };

// Runs the real tagWholeConversation over a stubbed mailbox, returning the categories it
// PATCHed onto each message so the tagging rules can be asserted directly.
async function tagAndCapture(messages, tag) {
  const patched = {};
  const graphFetch = async (url, opts) => {
    if (!opts || opts.method !== "PATCH") {
      return { ok: true, status: 200, json: async () => ({ value: messages }) };
    }
    patched[url.split("/messages/")[1]] = JSON.parse(opts.body).categories;
    return { ok: true, status: 200, json: async () => ({}) };
  };
  const fn = new Function("graphFetch", "console", `
    ${statusCatsSrc}
    ${tagFnSrc}
    return tagWholeConversation;
  `)(graphFetch, { warn() {}, log() {} });
  await fn("mb", "conv", null, tag);
  return patched;
}

const CASES = [
  {
    name: "sibling already has a job — returns its tag so no second job is created",
    graphFetch: ok([{ id: "other", categories: ["Job created - 106941"] }]),
    expect: (r, S) => r === "Job created - 106941",
  },
  {
    name: "sibling tagged Existing job — also counts as already handled",
    graphFetch: ok([{ id: "other", categories: ["Existing job - 106941"] }]),
    expect: (r, S) => r === "Existing job - 106941",
  },
  {
    name: "genuinely untagged thread — returns null so a job IS created",
    graphFetch: ok([{ id: "other", categories: ["Bara AI"] }]),
    expect: (r, S) => r === null,
  },
  {
    name: "only the message itself is in the thread — returns null",
    graphFetch: ok([{ id: "self", categories: ["Job created - 106941"] }]),
    expect: (r, S) => r === null,
  },
  {
    name: "Graph rejects the filter (the 107496 bug) — must NOT look like an untagged thread",
    graphFetch: httpError(400),
    expect: (r, S) => r === S,
  },
  {
    name: "Graph 429 throttling — must NOT look like an untagged thread",
    graphFetch: httpError(429),
    expect: (r, S) => r === S,
  },
  {
    name: "network error — must NOT look like an untagged thread, and must not throw",
    graphFetch: throws(),
    expect: (r, S) => r === S,
  },
];

let pass = 0;
const failures = [];
for (const c of CASES) {
  const { findJobTagInThread, THREAD_LOOKUP_FAILED } = build(c.graphFetch);
  let result;
  try {
    result = await findJobTagInThread("mb", "conv", "self");
  } catch (err) {
    failures.push(`  ${c.name}\n    threw: ${err.message}`);
    continue;
  }
  if (c.expect(result, THREAD_LOOKUP_FAILED)) pass++;
  else failures.push(`  ${c.name}\n    got: ${String(result)}`);
}

// --- tagWholeConversation: what actually lands on each message ---
const TAG_CASES = [
  {
    name: "the email that created the job keeps Job created, not downgraded to Existing job",
    messages: [{ id: "orig", categories: ["Job created - 106941"] }],
    tag: "Existing job - 106941",
    expect: p => p.orig === undefined, // untouched
  },
  {
    name: "a reply gets the Existing job tag",
    messages: [{ id: "reply", categories: ["Bara AI"] }],
    tag: "Existing job - 106941",
    expect: p => JSON.stringify(p.reply) === JSON.stringify(["Bara AI", "Existing job - 106941"]),
  },
  {
    name: "a stale Existing job tag is replaced, not duplicated",
    messages: [{ id: "reply", categories: ["Existing job - 999"] }],
    tag: "Existing job - 106941",
    expect: p => JSON.stringify(p.reply) === JSON.stringify(["Existing job - 106941"]),
  },
  {
    name: "transient status categories are cleared",
    messages: [{ id: "reply", categories: ["Reading Email", "Bara AI"] }],
    tag: "Existing job - 106941",
    expect: p => JSON.stringify(p.reply) === JSON.stringify(["Bara AI", "Existing job - 106941"]),
  },
  {
    name: "a Job created tag naming a DIFFERENT job survives (the 107496 evidence loss)",
    messages: [{ id: "orig", categories: ["Job created - 106941"] }],
    tag: "Job created - 107496",
    expect: p => p.orig.includes("Job created - 106941") && p.orig.includes("Job created - 107496"),
  },
  {
    name: "the incoming tag is never added twice",
    messages: [{ id: "m", categories: ["Job created - 106941"] }],
    tag: "Job created - 106941",
    expect: p => p.m.filter(c => c === "Job created - 106941").length === 1,
  },
];

for (const c of TAG_CASES) {
  let patched;
  try {
    patched = await tagAndCapture(c.messages, c.tag);
  } catch (err) {
    failures.push(`  ${c.name}\n    threw: ${err.message}`);
    continue;
  }
  if (c.expect(patched)) pass++;
  else failures.push(`  ${c.name}\n    patched: ${JSON.stringify(patched)}`);
}

// The failure sentinel must never be a falsy value: the caller guards with `if (siblingTag)`
// after the sentinel check, so a falsy sentinel would fall straight through to job creation.
const { THREAD_LOOKUP_FAILED } = build(ok([]));
if (THREAD_LOOKUP_FAILED) pass++;
else failures.push("  THREAD_LOOKUP_FAILED is falsy — it would fall through to job creation");

const total = CASES.length + TAG_CASES.length + 1;
console.log(`thread dedup: ${pass}/${total} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n${failures.join("\n")}`);
  process.exit(1);
}

#!/usr/bin/env node
// Regression tests for the poll's candidate window — run with: npm test
//
// index.js starts an HTTP server and the poll loop the moment it is imported, so the
// functions under test are lifted out of the source text and evaluated in isolation with
// graphFetch stubbed. Same approach as tools/thread-dedup.test.mjs — it tests the shipped
// source rather than a hand-copied paraphrase that could drift away from it.
//
// What these guard: an already-created work order still matches the Graph query (the job tag
// carries the job number, so there is nothing constant to filter on) and is dropped in code.
// On 2026-09-21 the window was one page of ten, a batch of created Bourkes work orders filled
// it, and oldest-first meant four newer work orders sat behind that wall — every poll from
// 02:19 reported "0 conversation(s) found" while work piled up. The window must span the
// whole inbox, and the heavy per-message content must NOT be paid for across it.

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

const extracted = [
  grab("const POLL_WINDOW", "POLL_BATCH_SIZE = 10;"),
  grab("async function fetchPollCandidates(mailbox, filter)", "\n}\n"),
  grab("async function fetchFullMessage(mailbox, messageId)", "\n}\n"),
].join("\n");

const module_ = `
export let calls, responder;
export function setup(c, r) { calls = c; responder = r; }
const graphFetch = async (path) => { calls.push(path); return responder(path); };
${extracted}
export { fetchPollCandidates, fetchFullMessage, POLL_WINDOW, POLL_BATCH_SIZE };
`;
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "poll-window-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const m = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const done  = (n) => ({ id: `done${n}`, subject: `done ${n}`,  categories: ["Bara AI", `Job created - ${100000 + n}`], conversationId: `cdone${n}`, receivedDateTime: `2026-09-21T01:${String(n % 60).padStart(2, "0")}:00Z` });
const fresh = (n) => ({ id: `new${n}`,  subject: `new ${n}`,   categories: ["Bara AI"],                                conversationId: `cnew${n}`,  receivedDateTime: `2026-09-21T02:${String(n % 60).padStart(2, "0")}:00Z` });

let calls = [];
function serveList(value) {
  return async () => ({ ok: true, status: 200, json: async () => ({ value }) });
}
async function candidates(value) {
  calls = [];
  m.setup(calls, serveList(value));
  const out = await m.fetchPollCandidates("mb", "FILTER");
  return { ...out, calls: [...calls] };
}

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

// The live failure of 2026-09-21: created work orders ahead of newer mail in oldest-first
// order. Under the old ten-message window this returned nothing at all.
{
  const inbox = [...Array.from({ length: 14 }, (_, i) => done(i + 1)), fresh(17), fresh(35)];
  const r = await candidates(inbox);
  check("created mail ahead of new mail does not hide it", r.messages.map(x => x.id).join(",") === "new17,new35", JSON.stringify(r.messages.map(x => x.id)));
  check("created mail is counted, not silently dropped", r.skipped === 14, `skipped=${r.skipped}`);
  check("it takes a single request", r.calls.length === 1, `calls=${r.calls.length}`);
}

// An inbox of nothing but created work orders is genuinely no work — but the count must say
// so, because "0 conversation(s) found" with no explanation is what hid this for hours.
{
  const r = await candidates(Array.from({ length: 30 }, (_, i) => done(i + 1)));
  check("an all-created inbox yields no candidates", r.messages.length === 0);
  check("an all-created inbox reports what it skipped", r.skipped === 30, `skipped=${r.skipped}`);
}

// More work than one tick processes is still all visible here; the batch cap lives in the
// poll, not the window, so the window must not truncate to it.
{
  const inbox = Array.from({ length: 25 }, (_, i) => fresh(i + 1));
  const r = await candidates(inbox);
  check("the window is not capped to the batch size", r.messages.length === 25, `messages=${r.messages.length} batch=${m.POLL_BATCH_SIZE}`);
  check("the window is wider than the batch", m.POLL_WINDOW > m.POLL_BATCH_SIZE, `window=${m.POLL_WINDOW} batch=${m.POLL_BATCH_SIZE}`);
  check("oldest stays first", r.messages[0].id === "new1" && r.messages[24].id === "new25");
}

// The query itself: the whole inbox, oldest first, and NOT the heavy fields — paying for
// bodies and attachments across the whole window is what made a wide window unaffordable.
{
  const r = await candidates([fresh(1)]);
  const q = r.calls[0];
  check("the query asks for the whole inbox", q.includes(`$top=${m.POLL_WINDOW}`), q);
  check("the query is oldest-first", /\$orderby=receivedDateTime(%20| )asc/.test(q), q);
  check("the query keeps the filter", q.includes("$filter=FILTER"), q);
  check("the query does not fetch bodies", !q.includes("body"), q);
  check("the query does not expand attachments", !q.includes("$expand"), q);
  check("the query still carries what phase 1 needs", ["id", "subject", "categories", "conversationId", "receivedDateTime"].every(f => q.includes(f)), q);
}

// A failed candidate query is the poll's own error to report, not something to swallow.
{
  calls = [];
  m.setup(calls, async () => ({ ok: false, status: 503, json: async () => ({ error: "nope" }) }));
  let threw = null;
  try { await m.fetchPollCandidates("mb", "FILTER"); } catch (err) { threw = err; }
  check("a failed candidate query fails the poll", threw !== null && /503/.test(threw.message), String(threw && threw.message));
}

// fetchFullMessage is where the heavy fields are paid for — once, per message processed.
{
  calls = [];
  m.setup(calls, async () => ({ ok: true, status: 200, json: async () => ({ id: "x", body: { content: "hi" }, attachments: [] }) }));
  const full = await m.fetchFullMessage("mb", "x");
  const q = calls[0];
  check("the full fetch targets one message", q.includes("/users/mb/messages/x"), q);
  check("the full fetch asks for the body", q.includes("body"), q);
  check("the full fetch expands attachments", q.includes("$expand=attachments"), q);
  check("the full fetch returns the message", full.body.content === "hi");
  check("the full fetch is one request per message", calls.length === 1, `calls=${calls.length}`);
}

{
  calls = [];
  m.setup(calls, async () => ({ ok: false, status: 404, json: async () => ({ error: "gone" }) }));
  let threw = null;
  try { await m.fetchFullMessage("mb", "x"); } catch (err) { threw = err; }
  check("a failed full fetch throws so the caller can retry", threw !== null && /404/.test(threw.message), String(threw && threw.message));
}

const failures = results.filter(r => !r.ok);
console.log(`poll window: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    got: ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

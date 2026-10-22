// The work-order gate: before a job is created, ask the dashboard what it knows about the
// property (jobs there in the last 30 days, the aircon unit tally) and, when the gate is
// armed, hold the email for a human decision instead of making a likely duplicate. The
// dashboard also keeps a record per email so the Outlook plugin can show the checks.
//
// Every dashboard call here is best effort — the callers catch. Nothing in this module may
// lose a work order: an unreachable dashboard means "create the job as today".
//
// Everything takes `fetchImpl` and `env` so the tests can drive it without a network.

export const NEEDS_DECISION_CATEGORY = "Needs Decision";
export const STOPPED_CATEGORY        = "Stopped";
export const GATE_CONTINUE_CATEGORY  = "Gate: Continue";
export const GATE_TEST_CATEGORY      = "Gate Test";
// An email whose agency name matched no client card waits under this until someone picks
// the card in the plugin. Taking the tag off by hand also retries it, as "Client not found"
// always did.
export const PICK_CLIENT_CATEGORY    = "Pick Client";
// What a pick puts on the email in place of "Pick Client", so it comes back through the
// candidate query and the job is made on the picked card. Same role as "Gate: Continue".
export const CLIENT_PICKED_CATEGORY  = "Client: Picked";

// Thrown by createArofloJob in place of creating the task. Carries the checks so the poll
// loop can log what the hold was decided on.
export class GateHold extends Error {
  constructor({ recentJobs = [], aircon = {} } = {}) {
    super(`Held for decision: ${recentJobs.length} recent job(s) at this location`);
    this.name = "GateHold";
    this.recentJobs = recentJobs;
    this.aircon = aircon;
  }
}

// Thrown by createArofloJob when the agency name matched no client card and the email is held
// for the office to pick one (the record is already NEEDS_CLIENT).
// `alreadyAsked`: the email was held for a pick before (its tag was taken off to retry), so
// the office has had its alert already.
export class ClientHold extends Error {
  constructor({ extractedName = null, aiPick = null, alreadyAsked = false } = {}) {
    super(`Held for a client pick: "${extractedName ?? ""}" matched no client card${aiPick ? ` (AI suggests ${aiPick})` : ""}`);
    this.name = "ClientHold";
    this.extractedName = extractedName;
    this.aiPick = aiPick;
    this.alreadyAsked = alreadyAsked;
  }
}

// What the plugin shows on an email held for a client pick. Lives on the record
// (checks.clientPick) like the duplicate prompt, so the wording is this service's to change.
//
//   { question, extractedName, from, ai: { aroFloId, name, reason } | null,
//     suggestions: [{ aroFloId, name }], askedBefore: [{ aroFloId, name }] }
//
// `extractedName` is also the key the dashboard remembers the pick under, with the sender.
export function clientPickPrompt({ extractedName, from, ai = null, suggestions = [], askedBefore = [] }) {
  const named = extractedName ? `"${extractedName}"` : "this agency";
  const question = askedBefore.length > 1
    ? `${named} has been put on ${askedBefore.length} different cards before. Which one is this work order for?`
    : ai
    ? `No client card is called ${named}. Is it ${ai.name}?`
    : `No client card is called ${named}. Which client is this work order for?`;
  return {
    question,
    extractedName: extractedName || null,
    from: from || null,
    ai,
    suggestions: suggestions.filter(c => c.aroFloId !== ai?.aroFloId).slice(0, 3),
    askedBefore,
  };
}

const GATE_MODES = ["off", "tagged", "on"];
let warnedUnknownMode = false;

// WORKORDER_GATE: off (default) runs the checks but never holds; tagged holds only emails
// carrying "Gate Test" so Brandon can trial it on real emails; on holds everything.
export function gateMode(env = process.env) {
  const raw = (env.WORKORDER_GATE || "off").trim().toLowerCase();
  if (GATE_MODES.includes(raw)) return raw;
  if (!warnedUnknownMode) {
    warnedUnknownMode = true;
    console.warn(`[gate] WORKORDER_GATE="${env.WORKORDER_GATE}" is not one of ${GATE_MODES.join("/")} — gate off`);
  }
  return "off";
}

// "Gate: Continue" is how a Continue decision reaches the poller: the email comes back
// through the normal candidate query wearing it, and the gate must not hold it a second time.
export function gateApplies(mode, categories = []) {
  if (categories.includes(GATE_CONTINUE_CATEGORY)) return false;
  if (mode === "on") return true;
  if (mode === "tagged") return categories.includes(GATE_TEST_CATEGORY);
  return false;
}

// The poller strips every status category, "Gate: Continue" included, the moment it picks
// an email up ("Reading email"), and only then reads the full message. So the Continue that
// put the email back in the queue is gone by the time the gate looks for it, and the gate
// held the email again — every "New job — create it" click looped straight back to Needs
// Decision. The poller remembers the Continue from the candidate it picked up and puts it
// back on the full message's categories with this.
export function carryGateContinue(pickedUp = [], fetched = []) {
  return carry(GATE_CONTINUE_CATEGORY, pickedUp, fetched);
}

// Same for "Client: Picked": createArofloJob reads it to put the job on the picked card
// even if the name matchers would now land somewhere else.
export function carryClientPicked(pickedUp = [], fetched = []) {
  return carry(CLIENT_PICKED_CATEGORY, pickedUp, fetched);
}

function carry(category, pickedUp, fetched) {
  if (!pickedUp.includes(category) || fetched.includes(category)) return fetched;
  return [...fetched, category];
}

// What the plugin asks the office, and the buttons it offers, for an email held because
// the property had a job lately. Lives on the record (checks.prompt) so the wording and
// the answers are this service's to change — the plugin draws whatever is here.
//
//   { question, answers: [{ id, label, effect: "CONTINUE" | "STOP", primary, then }] }
//
// `id` comes back as the record's `decision`; `effect` is what the dashboard turns it into
// (CONTINUED or STOPPED, which the poller acts on); `then` is what the office sees while
// the answer is in flight. Add answers freely — a new id with effect CONTINUE can carry a
// meaning of its own once applyGateDecisions reads `decision`.
export function duplicatePrompt(checks, now = new Date()) {
  const recent = Array.isArray(checks?.recentJobs) ? checks.recentJobs : [];
  const first = recent[0];
  const site = checks?.site ? ` at ${checks.site}` : " at this property";
  const when = first?.requestedAt ? ` on ${perthDate(first.requestedAt)}` : "";
  const sameOrder = Array.isArray(checks?.sameOrder) ? checks.sameOrder.filter(j => j?.jobNumber) : [];
  const question = sameOrder.length && checks?.orderNumber
    ? `Job ${sameOrder[0].jobNumber} already has order number ${checks.orderNumber}${sameOrder.length > 1 ? ` (and ${sameOrder.length - 1} more)` : ""}. Is this work order for a new job, or the same one?`
    : first
    ? `I found job ${first.jobNumber}${first.taskType ? ` (${first.taskType})` : ""}${site}, raised${when}${recent.length > 1 ? `, and ${recent.length - 1} more` : ""}. Is this work order for a new job, or the same one?`
    : `Something needs checking before I create this job${site}. Create it, or leave it?`;
  return {
    question,
    answers: [
      { id: "CONTINUE", label: "New job — create it", effect: "CONTINUE", primary: true, then: "Creating the job now…" },
      { id: "STOP", label: "Same job — don't create", effect: "STOP", then: (sameOrder[0] ?? first) ? `Not created — treated as the same job as ${(sameOrder[0] ?? first).jobNumber}.` : "Not created." },
    ],
  };
}

// dd/mm/yyyy in Perth, the way the office reads dates; an unparseable date is left out.
function perthDate(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Perth", day: "2-digit", month: "2-digit", year: "numeric" }).format(d);
}

// "Site: 2× Split, 1× Ducted" — the tally the dashboard derives from compliance forms, in
// the order the office reads them. Accepts either the bare tally or the { units } wrapper
// /api/property-check returns. Empty string when there is nothing to say, so the caller
// can append nothing.
const UNIT_ORDER = ["Split", "Ducted", "Evaporative"];
export function siteLine(aircon) {
  const units = aircon?.units ?? aircon ?? {};
  const parts = UNIT_ORDER.filter(t => units[t] > 0).map(t => `${units[t]}× ${t}`);
  return parts.length ? `Site: ${parts.join(", ")}` : "";
}

function dashboard(env) {
  const base = env.DASHBOARD_URL, secret = env.DASHBOARD_API_SECRET;
  if (!base || !secret) throw new Error("DASHBOARD_URL and DASHBOARD_API_SECRET must be set");
  return { base: base.replace(/\/$/, ""), headers: { Authorization: `Bearer ${secret}`, "Content-Type": "application/json" } };
}

async function dashboardFetch(path, init, label, { fetchImpl = fetch, env = process.env } = {}) {
  const { base, headers } = dashboard(env);
  const res = await fetchImpl(`${base}${path}`, { ...init, headers });
  if (!res.ok) throw new Error(`${label} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

// → { locationId, recentJobs: [{ jobNumber, taskType, status, requestedAt }], aircon: { Split: 2 } }.
// The dashboard finds the site by client + street (its Location rows carry no AroFlo
// location id — the reports never include one), and uses the v1 location id only to match
// its own earlier gate records. An unknown site is a 200 with empty lists, so a throw here
// means the dashboard itself is unreachable or refusing us.
//
// `orderNumber` (the agency's job number off the email) also brings back the client's jobs
// with that number in AroFlo's Cust ON, from the last year, wherever they are — a follow-up
// sent as a new thread is the same job.
export function propertyCheck({ aroFloLocationId, clientAroFloId, street, suburb, orderNumber }, opts) {
  const q = new URLSearchParams();
  if (orderNumber) q.set("orderNumber", orderNumber);
  if (aroFloLocationId) q.set("aroFloLocationId", aroFloLocationId);
  if (clientAroFloId) q.set("clientAroFloId", clientAroFloId);
  if (street) q.set("street", street);
  if (suburb) q.set("suburb", suburb);
  return dashboardFetch(`/api/property-check?${q}`, {}, "Property check", opts);
}

// Upsert by messageId; any subset of the record's fields. Called at each stage of an email
// (CHECKED / NEEDS_DECISION, then CREATED or FAILED) so the plugin always sees the latest.
export function upsertGateRecord(fields, opts) {
  return dashboardFetch("/api/workorder-gate", { method: "PUT", body: JSON.stringify(fields) }, "Gate record upsert", opts);
}

// Records whose decision the poller has not yet acted on: status CONTINUED or STOPPED.
export async function pendingDecisions(status, opts) {
  const body = await dashboardFetch(`/api/workorder-gate/pending?status=${encodeURIComponent(status)}`, {}, `Pending ${status}`, opts);
  return body?.items ?? [];
}

// Every active client card the dashboard holds: [{ aroFloId, name, category }]. Filled from
// AroFlo's client report, so it costs no AroFlo API calls and a card made this morning is in
// it within the hour.
export async function dashboardClients(opts) {
  const body = await dashboardFetch("/api/workorder-gate/clients", {}, "Client list", opts);
  return Array.isArray(body?.clients) ? body.clients : [];
}

// → { picks: [{ aroFloId, name, timesPicked }], propertyCards: [{ aroFloId, name }] }.
// Cards people picked before for this sender and agency name, and the card on the latest job
// at the property (a hint for the AI, never an answer on its own).
export function clientContext({ from, name, street, suburb }, opts) {
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (name) q.set("name", name);
  if (street) q.set("street", street);
  if (suburb) q.set("suburb", suburb);
  return dashboardFetch(`/api/workorder-gate/client-context?${q}`, {}, "Client context", opts);
}

// The gate record for one email, or null when there is none.
export async function gateRecord(messageId, { fetchImpl = fetch, env = process.env } = {}) {
  const { base, headers } = dashboard(env);
  const res = await fetchImpl(`${base}/api/workorder-gate?messageId=${encodeURIComponent(messageId)}`, { headers });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Gate record read failed: ${res.status} ${await res.text()}`);
  return res.json();
}

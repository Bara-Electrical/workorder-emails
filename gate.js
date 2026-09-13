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
  const question = first
    ? `I found job ${first.jobNumber}${first.taskType ? ` (${first.taskType})` : ""}${site}, raised${when}${recent.length > 1 ? `, and ${recent.length - 1} more` : ""}. Is this work order for a new job, or the same one?`
    : `Something needs checking before I create this job${site}. Create it, or leave it?`;
  return {
    question,
    answers: [
      { id: "CONTINUE", label: "New job — create it", effect: "CONTINUE", primary: true, then: "Creating the job now…" },
      { id: "STOP", label: "Same job — don't create", effect: "STOP", then: first ? `Not created — treated as the same job as ${first.jobNumber}.` : "Not created." },
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
export function propertyCheck({ aroFloLocationId, clientAroFloId, street, suburb }, opts) {
  const q = new URLSearchParams();
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

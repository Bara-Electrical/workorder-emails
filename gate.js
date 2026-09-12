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

// → { locationId, recentJobs: [{ jobNumber, taskType, status, requestedAt }], aircon: { units } }.
// An unknown location is a 200 with empty lists, so a throw here means the dashboard itself
// is unreachable or refusing us.
export function propertyCheck(aroFloLocationId, opts) {
  return dashboardFetch(`/api/property-check?aroFloLocationId=${encodeURIComponent(aroFloLocationId)}`, {}, "Property check", opts);
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

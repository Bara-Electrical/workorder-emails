// Reads the new client out of an AroFlo client webhook. AroFlo's webhooks arrive as
// { eventname, timestamputc, payload: { api_related_ids: { ... }, ...fields } } (seen on the
// dashboard's task, schedule and compliance-form webhooks). The client one has not been seen
// yet, so the id and name are looked for in each place they could plausibly be, and the
// first delivery is logged in full by the caller.

const ID_KEYS = ["clientid", "clientId", "client_id", "orgid", "id"];
const NAME_KEYS = ["clientname", "clientName", "client_name", "orgname", "name"];

function pick(obj, keys) {
  if (!obj || typeof obj !== "object") return null;
  for (const k of keys) {
    const v = obj[k];
    if (typeof v === "string" && v.trim()) return v.trim();
    if (typeof v === "number") return String(v);
  }
  return null;
}

// Returns { clientId, clientName }; either may be null.
export function clientFromWebhook(body) {
  if (!body || typeof body !== "object") return { clientId: null, clientName: null };
  const inner = body.payload && typeof body.payload === "object" ? body.payload : {};
  const ids = inner.api_related_ids || body.api_related_ids || {};
  const clientId =
    pick(ids, ["clientid", "clientId", "client_id", "orgid"]) ||
    pick(inner, ID_KEYS) || pick(body, ["clientid", "clientId", "client_id"]);
  const clientName = pick(inner, NAME_KEYS) || pick(body, ["clientname", "clientName", "client_name"]);
  return { clientId, clientName };
}

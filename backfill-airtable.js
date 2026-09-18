// ONE-OFF: copies every row out of Airtable into the Bara dashboard's Postgres, then
// verifies the copy field by field.
//
// WHY IT LIVES HERE, in a work-order email poller. Nothing else has both halves. The
// dashboard owns the database but has no Airtable key; a laptop has neither. This service
// already holds AIRTABLE_API_KEY, DASHBOARD_URL and LOG_API_SECRET, so it is the only
// place the two ends meet.
//
// Runs at startup when AIRTABLE_BACKFILL is set, and does nothing at all otherwise — the
// flag is set on Railway for one deploy and then removed. Delete this file once Airtable
// is gone.
//
// Safe to run more than once: every row upserts on its Airtable record id.
//
// Talks to Airtable's REST API with fetch rather than the airtable package, which this
// service dropped when its logging moved to the dashboard. Re-adding a dependency for a
// file that gets deleted next week is not worth it.

const BARA_AI = "app3KX3c4jIGMS0Zf";
const COMPLIANCE = "appaUXvN4fKrIffWx";
const BATCH = 200;

const text = (v) => (typeof v === "string" && v.trim() ? v.trim() : null);
const sel = (v) => (typeof v === "string" ? v : v && typeof v === "object" && typeof v.name === "string" ? v.name : null);
const list = (v) => (Array.isArray(v) ? v.map(sel).filter(Boolean) : []);
const numOrNull = (v) => (typeof v === "number" && Number.isFinite(v) ? v : null);
const day = (v) => (text(v) ? text(v).slice(0, 10) : null);

const TABLES = [
  {
    key: "activity",
    label: "Activity Log",
    base: BARA_AI,
    name: "Activity Log",
    map: (r) => ({
      airtableId: r.id,
      action: r.get("Action") || "Activity",
      department: sel(r.get("Department")),
      jobNumber: text(r.get("Job number")),
      source: "airtable-backfill",
      createdAt: r.get("Timestamp") || r._rawJson?.createdTime,
    }),
  },
  {
    key: "work-order",
    label: "Work Order AI Log",
    base: BARA_AI,
    name: "Work Order AI Log",
    map: (r) => ({
      airtableId: r.id,
      emailSubject: text(r.get("Email Subject")),
      taskType: text(r.get("Task Type")),
      package: text(r.get("Package")),
      address: text(r.get("Address")),
      realEstate: text(r.get("Real Estate")),
      propertyManager: text(r.get("Property Manager")),
      tenantName: text(r.get("Tenant Name")),
      tenantContact: text(r.get("Tenant Contact")),
      orderNumber: text(r.get("Order Number")),
      accountTo: text(r.get("Account To")),
      accessDetails: text(r.get("Access Details")),
      expenditureLimit: text(r.get("Expenditure Limit")),
      taskDescription: text(r.get("Task Description")),
      confidence: numOrNull(r.get("Confidence")),
      aiNotes: text(r.get("AI Notes")),
      createdAt: r.get("Created Time") || r._rawJson?.createdTime,
    }),
  },
  {
    // The one that has to be exact — the follow-up service decides who gets chased from
    // these dates. The verify pass below compares every field of every row, not a sample.
    key: "quote-follow-up",
    label: "Quote Follow-up Tracking",
    base: BARA_AI,
    name: "Quote Follow-up Tracking",
    map: (r) => ({
      airtableId: r.id,
      jobNumber: text(r.get("Job Number")),
      clientName: text(r.get("Client Name")),
      address: text(r.get("Address")),
      estimator: text(r.get("Estimator")),
      value: numOrNull(r.get("Value")),
      lastContactDate: day(r.get("Last Contact Date")),
      stage: sel(r.get("Stage")),
      lastVerified: day(r.get("Last Verified")),
    }),
  },
  {
    key: "compliance-property",
    label: "Properties",
    base: COMPLIANCE,
    name: "Properties",
    map: (r) => ({
      airtableId: r.id,
      reference: text(r.get("ID")),
      realEstate: sel(r.get("Real Estate")),
      propertyManager: text(r.get("Property Manager")),
      streetAddress: text(r.get("Street Address")),
      suburb: text(r.get("Suburb")),
      postcode: text(r.get("Postcode")),
      tenant: text(r.get("Tenant")),
      tenantContact: text(r.get("Tenant Contact")),
      managementStatus: sel(r.get("Management Status")),
      subscribedTo: list(r.get("Subscribed To")),
      dueDate: day(r.get("Due Date")),
    }),
  },
];

function dash(path, init) {
  const base = process.env.DASHBOARD_URL.replace(/\/$/, "");
  return fetch(`${base}${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${process.env.LOG_API_SECRET}`, "Content-Type": "application/json" },
  });
}

// A stand-in for the airtable package's record, so the maps above can keep reading fields
// by their Airtable column names.
function asRecord(record) {
  return {
    id: record.id,
    get: (field) => record.fields?.[field],
    _rawJson: record,
  };
}

async function allRows(baseId, tableName) {
  const out = [];
  let offset;
  do {
    const params = new URLSearchParams({ pageSize: "100" });
    if (offset) params.set("offset", offset);
    const res = await fetch(`https://api.airtable.com/v0/${baseId}/${encodeURIComponent(tableName)}?${params}`, {
      headers: { Authorization: `Bearer ${process.env.AIRTABLE_API_KEY}` },
    });
    if (!res.ok) throw new Error(`Airtable ${tableName} returned ${res.status}: ${(await res.text()).slice(0, 200)}`);
    const body = await res.json();
    for (const record of body.records ?? []) out.push(asRecord(record));
    offset = body.offset;
  } while (offset);
  return out;
}

export async function runAirtableBackfill() {
  console.log("[backfill] starting");
  const summary = [];

  for (const spec of TABLES) {
    let records;
    try {
      records = await allRows(spec.base, spec.name);
    } catch (err) {
      console.error(`[backfill] ${spec.label}: could not read Airtable — ${err.message}`);
      summary.push(`${spec.label}: READ FAILED`);
      continue;
    }

    let written = 0;
    let skipped = 0;
    for (let i = 0; i < records.length; i += BATCH) {
      const rows = records.slice(i, i + BATCH).map(spec.map);
      const res = await dash("/api/log/import", { method: "POST", body: JSON.stringify({ table: spec.key, rows }) });
      if (!res.ok) {
        console.error(`[backfill] ${spec.label}: import batch failed ${res.status} ${(await res.text()).slice(0, 200)}`);
        summary.push(`${spec.label}: IMPORT FAILED`);
        written = -1;
        break;
      }
      const body = await res.json();
      written += body.written;
      skipped += body.skipped;
      console.log(`[backfill] ${spec.label}: ${written}/${records.length}`);
    }
    if (written < 0) continue;

    // Verified by sending the same mapped rows back for comparison, so the check uses the
    // same normalisation the write used. A row count would pass a copy that silently
    // dropped every date in the follow-up table.
    let missing = 0;
    let differing = 0;
    let total = 0;
    const examples = [];
    for (let i = 0; i < records.length; i += BATCH) {
      const rows = records.slice(i, i + BATCH).map(spec.map);
      const res = await dash("/api/log/verify", { method: "POST", body: JSON.stringify({ table: spec.key, rows }) });
      if (!res.ok) {
        console.error(`[backfill] ${spec.label}: verify batch failed ${res.status} ${(await res.text()).slice(0, 200)}`);
        missing = -1;
        break;
      }
      const body = await res.json();
      missing += body.missing;
      differing += body.differing;
      total = body.total;
      for (const line of body.examples ?? []) if (examples.length < 10) examples.push(line);
    }
    if (missing < 0) {
      summary.push(`${spec.label}: VERIFY FAILED`);
      continue;
    }

    const ok = missing === 0 && differing === 0;
    for (const line of examples) console.error(`[backfill]   ${line}`);
    summary.push(
      `${ok ? "PASS" : "FAIL"} ${spec.label}: ${records.length} in Airtable, ${written} written, ${total} in Postgres` +
        `${skipped ? `, ${skipped} skipped` : ""}${missing ? `, ${missing} MISSING` : ""}${differing ? `, ${differing} FIELDS DIFFER` : ""}`
    );
  }

  console.log("[backfill] done\n  " + summary.join("\n  "));
  if (summary.some((line) => !line.startsWith("PASS"))) console.error("[backfill] NOT every table passed — do not delete anything in Airtable");
  return summary;
}

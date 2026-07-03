import "dotenv/config";
import express from "express";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { createHmac } from "crypto";
import { PACKAGE_TEMPLATES } from "./templates.js";
import Airtable from "airtable";

const REQUIRED_ENV = [
  "OPENAI_API_KEY",
  "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET",
  "UENCODED", "PENCODED", "ORGENCODED", "SECRET_KEY",
];
for (const key of REQUIRED_ENV) {
  if (!process.env[key]) {
    console.error(`[startup] ${key} is not set`);
    process.exit(1);
  }
}

let airtableBase = null;
if (process.env.AIRTABLE_API_KEY && process.env.AIRTABLE_BASE_ID) {
  airtableBase = new Airtable({ apiKey: process.env.AIRTABLE_API_KEY }).base(process.env.AIRTABLE_BASE_ID);
} else {
  console.warn("[startup] AIRTABLE_API_KEY or AIRTABLE_BASE_ID not set — activity and AI logging disabled");
}

if (!process.env.ADMIN_API_KEY) {
  console.warn("[startup] ADMIN_API_KEY not set — /clients, /find-client, and /aroflo-webhook are open with no auth");
}

// Known sender name → Aroflo client name mappings
const CLIENT_NAME_MAP = {
  "warilla pty ltd": "Peter Kuhne Real Estate",
  "warrilla pty ltd": "Peter Kuhne Real Estate",
  "professionals cannington": "Professionals BW Backhouse & Associates",
  "bw backhouse & associates": "Professionals BW Backhouse & Associates",
  "acton belle property south perth & victoria park": "Acton | Belle Property South Perth and Victoria Park",
  "leifield": "Leifield Real Estate",
  "leifield - wa": "Leifield Real Estate",
};

// Sender email domain → Aroflo client name (fallback when AI can't extract name from compound domains)
const EMAIL_DOMAIN_MAP = {
  "platinumelectricians.com.au": "Platinum Electricians",
};

const TRIGGER_CATEGORY          = "Bara AI";
const PROCESSING_CATEGORY       = "Processing";
const FAILED_CATEGORY           = "Failed";
const CLIENT_NOT_FOUND_CATEGORY = "Client not found";
const NO_ADDRESS_CATEGORY       = "No address";
const READING_EMAIL_CATEGORY    = "Reading Email";
const SENDING_TO_AI_CATEGORY    = "Sending to AI";
const CREATING_JOB_CATEGORY     = "Creating Job";
const WORKORDERS_EMAIL          = "workorders@baraelectrical.com.au";
const BRANDON_EMAIL             = "brandon.roberts@baraelectrical.com.au";
const POLL_INTERVAL_MS          = 5 * 60 * 1000;

// All transient status categories — stripped when transitioning to next stage
const STATUS_CATEGORIES = [
  PROCESSING_CATEGORY, FAILED_CATEGORY, READING_EMAIL_CATEGORY,
  SENDING_TO_AI_CATEGORY, CREATING_JOB_CATEGORY,
  CLIENT_NOT_FOUND_CATEGORY, NO_ADDRESS_CATEGORY,
];

let pollRunning = false;

const app = express();
app.use(express.json({ limit: "25mb" }));

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// Gate the debug/introspection endpoints behind a shared secret, checked as either an
// x-api-key header or a ?key= query param (so it can be embedded directly in a webhook
// URL, e.g. Aroflo's). Left open with a startup warning if ADMIN_API_KEY isn't set yet,
// so adding this check can't accidentally take the poll loop down before the env var is set.
function requireApiKey(req, res, next) {
  const configured = process.env.ADMIN_API_KEY;
  if (!configured) return next();
  const provided = req.headers["x-api-key"] || req.query.key;
  if (provided !== configured) return res.status(401).json({ error: "Unauthorized" });
  next();
}

// Aroflo's XML API returns a bare object for a single result and an array for multiple —
// normalise both shapes to an array so callers never need to branch on it themselves.
const toArray = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]);

// Wrap a value in a CDATA section for Aroflo's XML API. A literal "]]>" inside the value
// would otherwise prematurely close the section and corrupt the rest of the POST body —
// split it the standard way so the value round-trips safely either way.
function cdata(value) {
  return `<![CDATA[${String(value ?? "").replace(/]]>/g, "]]]]><![CDATA[>")}]]>`;
}

// Escape a plain-text value for safe interpolation into HTML (notes, alert emails).
// Not for use on values that are themselves meant to contain markup (e.g. the cleaned
// email body, or the hand-authored HTML templates).
function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// ================================================================
// RETRY HELPER — transient network errors and 429/5xx get retried with
// backoff; anything else (4xx, business-logic failures) is not retried.
// ================================================================
async function fetchWithRetry(url, options = {}, { attempts = 3, baseDelayMs = 500, label = "fetch" } = {}) {
  let lastErr, lastRes;
  for (let i = 0; i < attempts; i++) {
    try {
      const res = await fetch(url, options);
      if (res.ok || (res.status < 500 && res.status !== 429)) return res;
      lastRes = res;
      lastErr = new Error(`${label}: HTTP ${res.status}`);
    } catch (err) {
      lastErr = err;
    }
    if (i < attempts - 1) {
      const delay = baseDelayMs * 2 ** i;
      console.warn(`${label} — attempt ${i + 1}/${attempts} failed, retrying in ${delay}ms: ${lastErr.message}`);
      await new Promise(r => setTimeout(r, delay));
    }
  }
  if (lastRes) return lastRes; // exhausted retries but got a (failing) response — let the caller inspect/report it
  throw lastErr; // every attempt was a raw network failure
}

// ================================================================
// GRAPH API
// ================================================================
let tokenCache = { token: null, expiry: 0 };

async function getAccessToken() {
  if (tokenCache.token && Date.now() < tokenCache.expiry - 60000) return tokenCache.token;
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     process.env.GRAPH_CLIENT_ID,
        client_secret: process.env.GRAPH_CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error(`Graph auth failed: ${JSON.stringify(data)}`);
  tokenCache = { token: data.access_token, expiry: Date.now() + data.expires_in * 1000 };
  return tokenCache.token;
}

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  return fetchWithRetry(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization:  `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  }, { label: `graphFetch ${path}` });
}

// ================================================================
// AROFLO API
// ================================================================
const AROFLO_BASE   = "https://api.aroflo.com/";
const AROFLO_ACCEPT = "text/json";

function arofloAuth() {
  return (
    "uencoded="   + encodeURIComponent(process.env.UENCODED) +
    "&pencoded="  + encodeURIComponent(process.env.PENCODED) +
    "&orgEncoded=" + encodeURIComponent(process.env.ORGENCODED)
  );
}

function arofloSign(method, query, ts) {
  return createHmac("sha512", process.env.SECRET_KEY)
    .update([method, "", AROFLO_ACCEPT, arofloAuth(), ts, query].join("+"))
    .digest("hex");
}

// Aroflo's HMAC signature is tied to a specific timestamp, so a retry can't reuse the
// same signed request — each attempt re-signs with a fresh timestamp. Only network
// errors and 429/5xx are retried; a clean HTTP response with a non-"0" Aroflo status is
// a completed request (possibly already applied server-side for POSTs), so it's surfaced
// immediately rather than retried to avoid duplicating side effects like job/note creation.
async function arofloRequest(method, urlSuffix, body, label, attempt = 1) {
  const query = method === "GET" ? urlSuffix : body;
  const ts    = new Date().toISOString();
  const auth  = arofloAuth();
  const headers = {
    Accept:          AROFLO_ACCEPT,
    Authorization:   auth,
    Authentication:  "HMAC " + arofloSign(method, query, ts),
    afdatetimeutc:   ts,
  };
  if (method === "POST") headers["Content-Type"] = "application/x-www-form-urlencoded";

  let res;
  try {
    res = await fetch(
      method === "GET" ? AROFLO_BASE + "?" + urlSuffix : AROFLO_BASE + "?",
      method === "GET" ? { headers } : { method: "POST", headers, body }
    );
  } catch (err) {
    if (attempt < 3) {
      const delay = 500 * 2 ** (attempt - 1);
      console.warn(`${label} network error (attempt ${attempt}/3) — retrying in ${delay}ms: ${err.message}`);
      await new Promise(r => setTimeout(r, delay));
      return arofloRequest(method, urlSuffix, body, label, attempt + 1);
    }
    throw err;
  }

  if (!res.ok) {
    if ((res.status === 429 || res.status >= 500) && attempt < 3) {
      const delay = 500 * 2 ** (attempt - 1);
      console.warn(`${label} HTTP ${res.status} (attempt ${attempt}/3) — retrying in ${delay}ms`);
      await new Promise(r => setTimeout(r, delay));
      return arofloRequest(method, urlSuffix, body, label, attempt + 1);
    }
    throw new Error(`${label} request failed: HTTP ${res.status}`);
  }

  let data;
  try {
    data = await res.json();
  } catch {
    throw new Error(`${label} returned non-JSON response (HTTP ${res.status})`);
  }
  if (data.status !== "0") throw new Error(`${label} failed: ${data.statusmessage}`);
  return data.zoneresponse;
}

async function arofloGet(params) {
  return arofloRequest("GET", params, null, "Aroflo GET");
}

async function arofloPost(body) {
  return arofloRequest("POST", null, body, "Aroflo POST");
}

// Task type IDs sourced from live Aroflo system on 2026-06-24
const TASK_TYPE_MAP = {
  "EC1":                            "JCYqLyBSQCAgCg==", // $120 Standard Electrical Compliance
  "AC1":                            "JCYqWyVQUCAgCg==", // $180 Standard Air Con Service
  "AC2":                            "JCYqLyNRMCAgCg==", // $270 Deluxe Air Con Service
  "ACEC1":                          "JCYqLyBRMCAgCg==", // $240 Air Con & Electrical Compliance
  "Real Estate Aircon Maintenance":  "JCZaSyBSQCAgCg==", // Real-Estate Air Con Maintenance
  "Real Estate General Maintenance": "JCYqWyVQICAgCg==", // Real-Estate General Maintenance
};

const SUBSTATUS_MAP = {
  "EC1":                            "IyQ6SyYK", // Ready to schedule PPM1
  "AC1":                            "IyQ6SyYK", // Ready to schedule PPM1
  "AC2":                            "Iyc6LyUK", // Ready to schedule (Specialised)
  "ACEC1":                          "IyQ6SyYK", // Ready to schedule PPM1
  "Real Estate Aircon Maintenance":  "Iyc6LyYK", // Ready to schedule
  "Real Estate General Maintenance": "Iyc6LyYK", // Ready to schedule
};

// In-memory client cache: lowercase clientname → client object.
// Populated at startup and updated via the /aroflo-webhook endpoint.
const clientCache = new Map();
let clientCacheLastLoaded = null;

async function loadClientCache() {
  // where=clientid!=0 forces Aroflo to return the full client list.
  // Without a WHERE clause the API silently limits to recently active clients only.
  const WHERE = encodeURIComponent("and|clientid|!=|0");
  let page = 1, loaded = 0;
  try {
    while (true) {
      const zone = await arofloGet(`zone=clients&where=${WHERE}&page=${page}`);
      const raw  = zone?.clients;
      if (!raw) break;
      const arr  = toArray(raw);
      for (const c of arr) clientCache.set(c.clientname.toLowerCase(), c);
      loaded += arr.length;
      const current = parseInt(zone.currentpageresults ?? 0);
      const max     = parseInt(zone.maxpageresults ?? 500);
      if (current < max) break;
      page++;
    }
    clientCacheLastLoaded = new Date().toISOString();
    console.log(`[client] Cache loaded: ${clientCache.size} unique clients across ${page} page(s)`);
  } catch (err) {
    console.error("[client] Failed to load cache — fuzzy matching unavailable:", err.message);
  }
}

// Fetch a client's locations and contacts in a single call (join=locations,contacts)
// instead of two separate lookups.
async function findLocationsAndContacts(clientId) {
  try {
    const zone = await arofloGet(
      "zone=clients" +
      "&join=" + encodeURIComponent("locations,contacts") +
      "&where=" + encodeURIComponent(`and|clientid|=|${clientId}`) +
      "&where=" + encodeURIComponent("and|archived|=|false") +
      "&page=1"
    );
    const raw    = zone.clients;
    const client = toArray(raw)[0];
    return {
      locations: toArray(client?.locations),
      contacts:  toArray(client?.contacts),
    };
  } catch (err) {
    console.warn("[client] Location/contact search failed:", err.message);
    return { locations: [], contacts: [] };
  }
}

// Match a PM contact by name from an already-fetched contacts array.
function matchContact(contacts, pmName) {
  if (!pmName) return null;
  const nameLower = pmName.toLowerCase();
  return contacts.find(c => `${c.givennames} ${c.surname}`.toLowerCase().includes(nameLower)) || null;
}

// Candidate names to try when matching a client, from most to least specific:
// the raw name, the part before a "|"/"," separator, that with Pty/Ltd/etc stripped,
// and just its first word — deduped in order.
function clientNameCandidates(realEstateName) {
  const baseName     = realEstateName.split(/[|,]/)[0].trim();
  const strippedName = baseName.replace(/\s+(?:Pty\.?\s*)?(?:Ltd\.?|Limited|Inc\.?|LLC)\.?$/i, "").trim();
  return [realEstateName, baseName, strippedName, baseName.split(" ")[0]]
    .filter((v, i, a) => v && a.indexOf(v) === i);
}

// Search for a client by name.
// 1. Exact match against local cache (no API call needed).
// 2. Starts-with fuzzy match against local cache — only if exactly one result.
// 3. Falls back to Aroflo API exact match if cache is empty (not yet loaded).
async function findClient(realEstateName) {
  if (!realEstateName) return null;

  const candidates = clientNameCandidates(realEstateName);
  const baseName   = realEstateName.split(/[|,]/)[0].trim();

  // Cache lookup — check exact match on each candidate
  if (clientCache.size > 0) {
    for (const name of candidates) {
      const hit = clientCache.get(name.toLowerCase());
      if (hit) return hit;
    }

    // Bidirectional starts-with fuzzy match — handles "Driven Property Group Pty Ltd" → "Driven Property Group"
    const query   = baseName.toLowerCase();
    const matches = [...clientCache.values()].filter(c => {
      const name = c.clientname.toLowerCase();
      return name.startsWith(query) || query.startsWith(name);
    });
    if (matches.length === 1) {
      console.log(`[client] Fuzzy match: "${realEstateName}" → "${matches[0].clientname}"`);
      return matches[0];
    }
    if (matches.length > 1) {
      console.warn(`[client] Ambiguous name "${realEstateName}" — ${matches.length} cache matches: ${matches.map(c => c.clientname).join(", ")}`);
    }
    return null;
  }

  // Cache not loaded yet — fall back to API
  for (const name of candidates) {
    const zone = await arofloGet(
      "zone=clients&where=" + encodeURIComponent(`and|clientname|=|${name}`) + "&page=1"
    );
    const arr = toArray(zone.clients);
    if (arr.length > 0) return arr[0];
  }

  return null;
}

function parseAustralianAddress(address) {
  const parts = address.split(",").map(p => p.trim());
  const street = parts[0] || "";
  const rest   = parts.slice(1).join(", ").trim();
  // Optional comma before state handles "Perth, WA 6000" and "Perth WA 6000".
  // Explicit state list handles title-cased abbreviations like "Wa" or "Nsw".
  const match  = rest.match(/^(.*?),?\s+(NSW|VIC|QLD|SA|WA|TAS|ACT|NT)(?:[,\s]+(\d{4}))?$/i);
  return match
    ? { street, suburb: match[1].trim(), state: match[2].toUpperCase(), postcode: match[3] || "" }
    : { street, suburb: rest, state: "", postcode: "" };
}

async function geocodeAddress(address) {
  const key = process.env.HERE_API_KEY;
  if (!key) { console.warn("[location] HERE_API_KEY not set — skipping geocode"); return null; }
  try {
    const url  = "https://geocode.search.hereapi.com/v1/geocode?q=" +
                 encodeURIComponent(address) + "&in=countryCode:AUS&apiKey=" + key;
    const res  = await fetch(url);
    const data = await res.json();
    const item = data?.items?.[0];
    if (item?.position) return { lat: item.position.lat, lon: item.position.lng };
  } catch (err) {
    console.warn("[location] Geocode failed:", err.message);
  }
  return null;
}

async function createLocation(clientId, address, tenantName, tenantContact, tenantEmail) {
  const { street, suburb, state, postcode } = parseAustralianAddress(address);
  const coords = await geocodeAddress(address);
  const xml =
`<clients><client>
  <clientid>${clientId}</clientid>
  <locations><location>
    <locationname>${cdata(street)}</locationname>
    <suburb>${cdata(suburb)}</suburb>
    <state>${cdata(state)}</state>
    <postcode>${cdata(postcode)}</postcode>
    <country><![CDATA[Australia]]></country>
    ${coords ? `<gpslat>${coords.lat}</gpslat><gpslong>${coords.lon}</gpslong>` : ""}
    ${tenantName    ? `<sitecontact>${cdata(tenantName)}</sitecontact>`  : ""}
    ${tenantContact ? `<sitephone>${cdata(tenantContact)}</sitephone>`   : ""}
    ${tenantEmail   ? `<siteemail>${cdata(tenantEmail)}</siteemail>`     : ""}
  </location></locations>
</client></clients>`;
  const createZone = await arofloPost("zone=clients&postxml=" + encodeURIComponent(xml));
  console.log("[location] Create response:", JSON.stringify(createZone?.postresults));

  // Aroflo returns the new locationid under updates.clients[0].locations[0]
  const pr          = createZone?.postresults;
  const updClients  = pr?.updates?.clients;
  const updClient   = toArray(updClients)[0];
  const updLocs     = updClient?.locations;
  const newId       = toArray(updLocs)[0]?.locationid;
  if (newId) {
    console.log("[location] Created:", newId, address);
    return { locationid: newId, locationname: address };
  }

  // Fall back: re-query to find the location we just created
  console.log("[location] locationid not in response — fetching newly created location");
  const zone = await arofloGet(
    "zone=clients" +
    "&where=" + encodeURIComponent(`and|clientid|=|${clientId}`) +
    "&join=locations"
  );
  const client = toArray(zone.clients)[0];
  const all = toArray(client?.locations);
  const streetPart = address.replace(/^\d+\//, "").split(",")[0].trim().toLowerCase();
  const found = all.find(l => l.locationname?.toLowerCase().includes(streetPart));
  if (found) {
    console.log("[location] Found after creation:", found.locationid, found.locationname);
    return found;
  }

  console.warn("[location] Created but could not retrieve locationid");
  return null;
}

// Aroflo silently truncates SiteContact/SitePhone to 50 characters each (confirmed
// empirically — no validation error, it just clips). For multi-tenant jobs this can
// cut off a name or number mid-word. Keep the same number of tenants in both fields
// (so name[i] still lines up with phone[i]) and report anything dropped separately.
const SITE_FIELD_LIMIT = 50;
function fitTenantFields(tenantName, tenantContact, maxLen = SITE_FIELD_LIMIT) {
  const names  = tenantName    ? tenantName.split(",").map(s => s.trim())    : [];
  const phones = tenantContact ? tenantContact.split(",").map(s => s.trim()) : [];
  const total  = Math.max(names.length, phones.length);
  const fitsWithin = (items, count) => items.slice(0, count).join(", ").length <= maxLen;

  let keptCount = total;
  while (keptCount > 0 &&
         ((names.length  && !fitsWithin(names, keptCount)) ||
          (phones.length && !fitsWithin(phones, keptCount)))) {
    keptCount--;
  }

  // Single tenant longer than the limit on its own — hard-truncate rather than drop it entirely.
  if (keptCount === 0 && total > 0) {
    return {
      keptName:     names[0]  ? names[0].slice(0, maxLen)  : null,
      keptPhone:    phones[0] ? phones[0].slice(0, maxLen) : null,
      overflowName:  names.slice(1),
      overflowPhone: phones.slice(1),
      truncated: true,
    };
  }

  return {
    keptName:      names.slice(0, keptCount).join(", ")  || null,
    keptPhone:      phones.slice(0, keptCount).join(", ") || null,
    overflowName:   names.slice(keptCount),
    overflowPhone:  phones.slice(keptCount),
    truncated: keptCount < total,
  };
}

// Find a location by street address from an already-fetched locations array, then
// update SiteContact/SitePhone if stale, or create the location if it doesn't exist yet.
// Matches by street address locally — the locationname|like WHERE clause is not
// supported by Aroflo.
async function findOrUpdateLocation(clientId, locations, address, tenantName, tenantContact, tenantEmail) {
  if (!address) return null;

  // Strip unit prefix: "1412/380 Murray Street, Perth WA" → "380 Murray Street"
  const streetPart = address.replace(/^\d+\//, "").split(",")[0].trim().toLowerCase();
  const incomingUnit = address.match(/^(\d+)\//)?.[1] || null;
  const { suburb: incomingSuburb } = parseAustralianAddress(address);

  const forClient = locations;
  const active = forClient.filter(l => l.archived?.toUpperCase() !== "TRUE");
  console.log(`[location] Search — client ${clientId}: ${forClient.length} location(s) (${active.length} active), searching for "${streetPart}"${incomingSuburb ? ` in "${incomingSuburb}"` : ""}`);
  // A street-only match isn't enough when a building has multiple numbered units on
  // file — "10/27 X" contains "27 x" as a substring, so it would wrongly match an
  // incoming "9/27 X" and silently attach the job to a different unit's tenant.
  // If both the incoming address and the candidate location have a unit number,
  // they must match exactly. Same idea for suburb: the same street name can exist in
  // two different suburbs for a large client, so require the suburb to agree too
  // whenever both are known — otherwise a job (and the tenant's details) could get
  // silently attached to the wrong property.
  const location = active.find(l => {
    if (!l.locationname?.toLowerCase().includes(streetPart)) return false;
    const storedUnit = l.locationname?.match(/^(\d+)\//)?.[1] || null;
    if (incomingUnit && storedUnit && incomingUnit !== storedUnit) return false;
    if (incomingSuburb && l.suburb && incomingSuburb.toLowerCase() !== l.suburb.toLowerCase()) return false;
    return true;
  });

  if (!location) {
    console.log("[location] No match for:", streetPart, "— creating new location:", address);
    try {
      return await createLocation(clientId, address, tenantName, tenantContact, tenantEmail);
    } catch (err) {
      console.warn("[location] Creation failed:", err.message);
      return null;
    }
  }

  console.log("[location] Found:", location.locationid, location.locationname);

  if (tenantName || tenantContact || tenantEmail) {
    // Must be wrapped in <clients><client> — a bare <locations><location> POST to
    // zone=locations returns status "0" with no error but silently does not apply.
    const xml =
`<clients><client>
  <clientid>${clientId}</clientid>
  <locations><location>
    <locationid>${location.locationid}</locationid>
    ${tenantName    ? `<sitecontact>${cdata(tenantName)}</sitecontact>`  : ""}
    ${tenantContact ? `<sitephone>${cdata(tenantContact)}</sitephone>`   : ""}
    ${tenantEmail   ? `<siteemail>${cdata(tenantEmail)}</siteemail>`     : ""}
  </location></locations>
</client></clients>`;
    try {
      const updateRes = await arofloPost("zone=clients&postxml=" + encodeURIComponent(xml));
      console.log("[location] Tenant details updated:", JSON.stringify(updateRes?.postresults));
    } catch (err) {
      console.warn("[location] Tenant update failed:", err.message);
    }
  }

  return location;
}

async function logActivity(action, jobNumber) {
  if (!airtableBase) return;
  try {
    await airtableBase("Activity Log").create([{
      fields: { "Action": action, "Job number": jobNumber || null, "Department": "Admin" },
    }]);
  } catch (err) {
    console.warn("[airtable] Activity log failed:", err.message);
  }
}

async function logAiOutput(result, emailSubject) {
  if (!airtableBase) return;
  try {
    await airtableBase("Work Order AI Log").create([{
      fields: {
        "Email Subject":     emailSubject || null,
        "Task Type":         result["task-type"] || null,
        "Package":           result["package"] || null,
        "Address":           result["address"] || null,
        "Real Estate":       result["real-estate"] || null,
        "Property Manager":  result["property-manager"] || null,
        "Tenant Name":       result["tenant-name"] || null,
        "Tenant Contact":    result["tenant-contact"] || null,
        "Order Number":      result["order-number"] || null,
        "Account To":        result["account-to"] || null,
        "Access Details":    result["access-details"] || null,
        "Expenditure Limit": result["expenditure-limit"] || null,
        "Task Description":  result["task-description"] || null,
        "Confidence":        result["confidence"] != null ? Number(result["confidence"]) : null,
        "AI Notes":          result["notes"] || null,
      },
    }]);
  } catch (err) {
    console.warn("[airtable] AI log failed:", err.message);
  }
}

function buildDescription(result) {
  const parts = [];
  const spacer = `<p>&nbsp;</p>`;

  if (result["task-description"] || result["task-type"]) {
    parts.push(`<p>${escapeHtml(result["task-description"] || result["task-type"])}</p>`);
  }

  const hasHighlights = result["expenditure-limit"] || result["access-details"];
  if (hasHighlights) parts.push(spacer);

  if (result["expenditure-limit"]) {
    parts.push(`<p><span style="background:#cce5ff;font-weight:bold">Expenditure Limit: ${escapeHtml(result["expenditure-limit"])}</span></p>`);
  }

  if (result["access-details"]) {
    parts.push(`<p><span style="background:#ccffcc;font-weight:bold">Access Details: ${escapeHtml(result["access-details"])}</span></p>`);
  }

  // Fall back to task-type if AI forgot to set package (e.g. task-type is EC1 but package is null)
  const pkg = (result["package"] && result["package"] !== "null" ? result["package"] : null)
    ?? (PACKAGE_TEMPLATES[result["task-type"]] ? result["task-type"] : null);
  if (pkg && PACKAGE_TEMPLATES[pkg]) {
    parts.push(spacer);
    parts.push(PACKAGE_TEMPLATES[pkg]);
  }

  return parts.join("\n");
}

async function createArofloJob(result, rawEmail, pdfAttachment = null, emailMeta = null, imageAttachments = []) {
  console.log("[job] Creating Aroflo job...");
  const warnings = [];

  if (!result.address) throw new Error("No address found in work order");

  const taskTypeId  = TASK_TYPE_MAP[result["task-type"]];
  const substatusId = SUBSTATUS_MAP[result["task-type"]] || "Iyc6LyYK"; // default: Ready to schedule
  if (!taskTypeId) {
    const detail = `Unknown task type: "${result["task-type"]}" — task created without a task type`;
    console.warn("[job]", detail);
    warnings.push({ tag: "Unknown task type", detail });
  }

  const realEstate = CLIENT_NAME_MAP[result["real-estate"]?.toLowerCase()] || result["real-estate"];
  console.log(`[job] Client lookup — AI extracted real-estate: "${result["real-estate"]}", resolved to: "${realEstate}", from: "${emailMeta?.from}"`);
  let client = await findClient(realEstate);
  let clientFoundVia = "name";
  if (!client && emailMeta?.from) {
    const domain = emailMeta.from.split("@")[1];
    const domainName = domain && EMAIL_DOMAIN_MAP[domain.toLowerCase()];
    console.log(`[job] Client not found by name — domain: "${domain}", domain map hit: "${domainName || "none"}"`);
    if (domainName) {
      client = await findClient(domainName);
      if (!client) {
        // Client not in cache (may be a supplier/subcontractor type) — try live Aroflo API
        console.log(`[job] Cache miss for "${domainName}" — trying live Aroflo API lookup`);
        const zone = await arofloGet(`zone=clients&where=${encodeURIComponent(`and|clientname|=|${domainName}`)}&page=1`);
        client = toArray(zone?.clients)[0] || null;
      }
      clientFoundVia = `email domain (${domainName})`;
    }
  }
  if (!client) throw new Error(`Client not found in Aroflo: name="${realEstate}", from="${emailMeta?.from}"`);
  console.log(`[job] Client (via ${clientFoundVia}):`, client.clientid, client.clientname);

  const { locations, contacts } = await findLocationsAndContacts(client.clientid);

  const tenantFit = fitTenantFields(result["tenant-name"], result["tenant-contact"]);
  let additionalTenantNote = null;
  if (tenantFit.truncated) {
    const rows = Math.max(tenantFit.overflowName.length, tenantFit.overflowPhone.length);
    const lines = [];
    for (let i = 0; i < rows; i++) {
      const line = [tenantFit.overflowName[i], tenantFit.overflowPhone[i]].filter(Boolean).join(", ");
      if (line) lines.push(`Additional tenant - ${escapeHtml(line)}`);
    }
    additionalTenantNote = lines.join("<br/>");
    const detail = `SiteContact/SitePhone are capped at ${SITE_FIELD_LIMIT} characters by Aroflo — full list: "${result["tenant-name"] || ""}" / "${result["tenant-contact"] || ""}"`;
    console.warn("[job]", detail);
    warnings.push({ tag: "Tenant details truncated", detail });
  }

  const location = await findOrUpdateLocation(
    client.clientid,
    locations,
    result.address,
    tenantFit.keptName,
    tenantFit.keptPhone,
    result["tenant-email"]
  );
  if (!location && result.address) {
    const detail = `Location not linked for "${result.address}" — address used as site name fallback`;
    console.warn("[job]", detail);
    warnings.push({ tag: "Location not linked", detail });
  }

  const pmContact = matchContact(contacts, result["property-manager"]);
  if (pmContact) {
    console.log("[job] PM contact:", pmContact.contactid, pmContact.contactname);
  } else if (result["property-manager"]) {
    const detail = `PM contact not found in Aroflo: "${result["property-manager"]}"`;
    console.warn("[job]", detail);
    warnings.push({ tag: "PM not in Aroflo", detail });
  }

  const dueDate = (() => {
    const d = new Date();
    d.setDate(d.getDate() + 7);
    return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
  })();

  const { street: taskStreet, suburb: taskSuburb } = parseAustralianAddress(result.address || "");
  const taskName = [taskStreet, taskSuburb]
    .filter(Boolean)
    .join(" ")
    .replace(/,?\s*(?:NSW|VIC|QLD|SA|WA|TAS|ACT|NT)\b\s*\d{0,4}\s*$/i, "")
    .trim();

  const xml =
`<tasks>
  <task>
    <org><orgid>JiYqTydSXDcmCg==</orgid></org>
    <contactname>Bara AI</contactname>
    ${taskTypeId ? `<tasktype><tasktypeid>${taskTypeId}</tasktypeid></tasktype>`                  : ""}
    <client><clientid>${client.clientid}</clientid></client>
    ${pmContact ? `<contact><userid>${pmContact.userid}</userid></contact>` : ""}
    ${location ? `<location><locationid>${location.locationid}</locationid></location>` : ""}
    ${result.address && !location  ? `<sitename>${cdata(result.address)}</sitename>`          : ""}
    <taskname>${cdata(taskName)}</taskname>
    <description>${cdata(buildDescription(result))}</description>
    <duedate>${dueDate}</duedate>
    ${result["order-number"] ? `<custon>${cdata(result["order-number"])}</custon>` : ""}
    ${(result["account-to"] || realEstate) ? `<customfields><customfield><name><![CDATA[ Account To: ]]></name><type><![CDATA[ text ]]></type><value>${cdata(result["account-to"] || realEstate)}</value></customfield></customfields>` : ""}
  </task>
</tasks>`;

  const zone     = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(xml));
  const pr          = zone.postresults;
  const insertTotal = Number(pr?.inserttotal ?? 0);

  if (insertTotal < 1) {
    const errArr = toArray(pr?.errors);
    const msgs   = errArr.length
      ? errArr.map(e => e.detail || e.message || JSON.stringify(e)).join("; ")
      : "No job inserted";
    throw new Error(`Aroflo task creation failed: ${msgs}`);
  }

  const task     = toArray(pr?.inserts?.tasks)[0];
  const taskId   = task?.taskid;

  // jobnumber isn't in the insert response — fetch it. Also use the fetched taskId
  // for note posting since it's confirmed to work with this format.
  let jobNumber = "(see Aroflo)";
  let confirmedTaskId = taskId;
  if (taskId) {
    try {
      const fetched = await arofloGet("zone=tasks&where=" + encodeURIComponent(`and|taskid|=|${taskId}`) + "&page=1");
      const arr = toArray(fetched.tasks);
      jobNumber = arr[0]?.jobnumber || taskId;
      confirmedTaskId = arr[0]?.taskid || taskId;
    } catch (err) {
      console.warn("[job] Could not fetch job number:", err.message);
      jobNumber = taskId;
    }
  }
  console.log("[job] Aroflo job created — job number:", jobNumber, "taskId:", confirmedTaskId);

  // Upload PDF and any photo attachments to SharePoint
  let oneDriveUrl = null;
  const photos = [];
  if (jobNumber !== "(see Aroflo)") {
    if (pdfAttachment) {
      try {
        oneDriveUrl = await uploadWorkOrderToOneDrive(jobNumber, pdfAttachment.name, pdfAttachment.data);
        console.log("[sharepoint] PDF uploaded:", oneDriveUrl);
      } catch (err) {
        console.warn("[sharepoint] PDF upload failed:", err.message);
        warnings.push({ tag: "PDF upload failed", detail: err.message });
      }
    }
    if (imageAttachments.length > 0) {
      const driveId = await getSharepointDriveId();
      const photoResults = await Promise.all(
        imageAttachments.map(async img => {
          try {
            const item = await uploadPhotoToOneDrive(jobNumber, img.name, img.data, img.contentType);
            const thumbnailUrl = await getThumbnailUrl(driveId, item.id).catch(() => null);
            console.log("[sharepoint] Photo uploaded:", img.name);
            return { name: img.name, webUrl: item.webUrl, thumbnailUrl };
          } catch (err) {
            console.warn("[sharepoint] Photo upload failed:", img.name, err.message);
            warnings.push({ tag: "Photo upload failed", detail: `${img.name}: ${err.message}` });
            return null;
          }
        })
      );
      photos.push(...photoResults.filter(Boolean));
    }
  }

  // Post the original email as a note and set the substatus in one combined task
  // update (Aroflo doesn't apply substatus on create, so a follow-up write is
  // always needed — piggyback the note content on the same call).
  if (confirmedTaskId) {
    let noteHtml = null;
    if (rawEmail) {
      try {
        noteHtml = await emailHtmlForNote(rawEmail, oneDriveUrl, emailMeta);
      } catch (err) {
        const detail = `Email note not posted to job: ${err.message}`;
        console.warn("[job]", detail);
        warnings.push({ tag: "Note not posted", detail });
      }
    } else {
      warnings.push({ tag: "Note not posted", detail: "No email content available — note not posted to job" });
    }

    const photoGalleryHtml = photos.length > 0 ? buildPhotoGalleryNote(photos) : null;

    const notesXml = [
      noteHtml             ? `<note><content>${cdata(noteHtml)}</content></note>`             : "",
      photoGalleryHtml     ? `<note><content>${cdata(photoGalleryHtml)}</content></note>`      : "",
      additionalTenantNote ? `<note><content>${cdata(additionalTenantNote)}</content></note>` : "",
    ].join("");

    if (notesXml || substatusId) {
      const updateXml =
`<tasks>
  <task>
    <taskid>${confirmedTaskId}</taskid>
    ${substatusId ? `<status>not started</status><substatus><substatusid>${substatusId}</substatusid></substatus>` : ""}
    ${notesXml ? `<notes>${notesXml}</notes>` : ""}
  </task>
</tasks>`;
      try {
        const upZone = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(updateXml));
        const upPr   = upZone.postresults;
        if (Number(upPr?.updatetotal ?? 0) > 0) {
          console.log("[job] Task update applied — note:", !!noteHtml, "additional tenant note:", !!additionalTenantNote, "substatus:", substatusId || "n/a");
        } else {
          if (noteHtml) warnings.push({ tag: "Note not posted", detail: "Combined task update did not apply" });
          if (substatusId) warnings.push({ tag: "Substatus failed", detail: "Combined task update did not apply — job may need manual scheduling status update" });
        }
      } catch (err) {
        if (noteHtml) warnings.push({ tag: "Note not posted", detail: `Email note not posted to job: ${err.message}` });
        if (substatusId) warnings.push({ tag: "Substatus failed", detail: `Substatus not applied: ${err.message}` });
      }
    }
  }

  return { jobNumber, warnings };
}

// ================================================================
// PDF + HTML helpers
// ================================================================
async function extractPDF(data) {
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf         = await loadingTask.promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page    = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(" ") + "\n";
  }
  return text;
}

function cleanHtml(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Decode SafeLinks/Inky wrapped URLs. Inky requires following the redirect.
async function decodeWrappedLinks(html) {
  const matches = [...html.matchAll(/href="([^"]+)"/gi)];
  const unique  = [...new Set(matches.map(m => m[1]))];
  const map     = {};

  await Promise.all(unique.map(async href => {
    const decoded = href.replace(/&amp;/g, "&");
    try {
      if (/safelinks\.protection\.outlook\.com/i.test(decoded)) {
        const url = new URL(decoded).searchParams.get("url");
        if (url) map[href] = decodeURIComponent(url);
      } else if (/shared\.outlook\.inky\.com/i.test(decoded)) {
        const res = await fetch(decoded.includes("confirm=True") ? decoded : decoded + "&confirm=True", {
          redirect: "follow",
          headers: { "User-Agent": "Mozilla/5.0" },
        });
        map[href] = res.url;
      }
    } catch { /* leave as-is */ }
  }));

  return html.replace(/href="([^"]+)"/gi, (match, href) =>
    map[href] ? `href="${map[href]}"` : match
  );
}

// Strip scripts/styles/tracking pixels but keep HTML structure for display in Aroflo notes
async function emailHtmlForNote(html, oneDriveUrl = null, emailMeta = null) {
  let cleaned = html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<img[^>]*>/gi, "")
    .trim();

  // Remove Inky security banner — strip from start of HTML up to and including the
  // ipw-end anchor (raw email uses "ipw-end-*", Outlook renderer prefixes it "x_ipw-end-*")
  cleaned = cleaned.replace(/[\s\S]*?<a\b[^>]*name="(?:x_)?ipw-end-\d+"[^>]*><\/a>(?:\s*<\/\w+>)*/i, "");

  cleaned = await decodeWrappedLinks(cleaned);

  const cell = (label, value) =>
    `<tr><td style="color:#888888;font-size:12px;font-weight:bold;padding:1px 12px 1px 0;white-space:nowrap;vertical-align:top">${label}</td><td style="color:#444444;font-size:12px;padding:1px 0">${value}</td></tr>`;

  const metaRows = [
    emailMeta?.from    ? cell("From:",       escapeHtml(emailMeta.from))    : "",
    emailMeta?.to      ? cell("To:",         escapeHtml(emailMeta.to))      : "",
    emailMeta?.subject ? cell("Subject:",    escapeHtml(emailMeta.subject)) : "",
    oneDriveUrl        ? cell("Attachment:", `<a href="${oneDriveUrl}" style="color:#1a6bbf" target="_blank">View Work Order PDF</a>`) : "",
  ].filter(Boolean).join("");

  const titleRow = `<tr><td colspan="2" style="font-size:16px;font-weight:bold;color:#444444;padding:0 0 5px 0">Work Order</td></tr>`;
  const metaHtml = `<table style="border-collapse:collapse;margin:0 0 12px 0">${titleRow}${metaRows}</table>`;

  return `${metaHtml}<hr style="border:none;border-top:1px solid #dddddd;margin:0 0 14px 0"><div>${cleaned}</div>`;
}

// A plain driveItem webUrl opens the file in isolation with no folder context — no gallery
// navigation. Browsing to the parent folder and clicking a photo from there does give
// SharePoint's native arrow navigation between sibling files (confirmed manually), so every
// thumbnail links to the shared folder rather than a single-file deep link.
function buildFolderUrl(webUrl) {
  const u = new URL(webUrl);
  return u.origin + u.pathname.slice(0, u.pathname.lastIndexOf("/"));
}

// Small clickable thumbnail grid, posted as its own note. Every thumbnail links to the same
// job photo folder — opening it gives SharePoint's native gallery view where the tech can
// click through all of the job's photos with arrow navigation.
function buildPhotoGalleryNote(photos) {
  const folderUrl = buildFolderUrl(photos[0].webUrl);
  const thumbs = photos.map(p =>
    `<a href="${folderUrl}" target="_blank"><img src="${p.thumbnailUrl || p.webUrl}" alt="${escapeHtml(p.name)}" style="width:110px;height:110px;object-fit:cover;margin:0 8px 8px 0;border-radius:4px;border:1px solid #dddddd" /></a>`
  ).join("");

  const titleRow = `Photos (${photos.length})`;
  return `<div style="font-size:14px;font-weight:bold;color:#444444;margin:0 0 8px 0">${titleRow}</div><div>${thumbs}</div>`;
}

// Cached drive ID for the Bara Electrical Services SharePoint document library
let sharepointDriveId = null;

async function getSharepointDriveId() {
  if (sharepointDriveId) return sharepointDriveId;
  const siteRes  = await graphFetch(`/sites/baraelectrical.sharepoint.com:/sites/BaraElectricalServices`);
  const site     = await siteRes.json();
  const drivesRes = await graphFetch(`/sites/${site.id}/drives`);
  const drives   = await drivesRes.json();
  const drive    = (drives.value || []).find(d => d.name === "Documents" || d.name === "Shared Documents");
  if (!drive) throw new Error("SharePoint Documents drive not found");
  sharepointDriveId = drive.id;
  return sharepointDriveId;
}

async function uploadWorkOrderToOneDrive(jobNumber, filename, contentBytes, contentType = "application/pdf") {
  const safeName = filename.replace(/#/g, "").trim();
  const itemPath = ["General", "Other", "AI Workorders [dont touch]", `${jobNumber} - ${safeName}`]
    .map(s => encodeURIComponent(s)).join("/");
  const uploadData = await putSharepointFile(itemPath, contentBytes, contentType);
  return uploadData.webUrl || null;
}

// Photos go in a per-job subfolder (rather than the flat filename-prefixed layout used for
// the PDF) so that opening any one of them from the note still gives SharePoint's native
// gallery view — arrow through the rest of the job's photos — instead of a dead-end preview.
async function uploadPhotoToOneDrive(jobNumber, filename, contentBytes, contentType) {
  const safeName = filename.replace(/#/g, "").trim();
  const itemPath = ["General", "Other", "AI Workorders [dont touch]", jobNumber, safeName]
    .map(s => encodeURIComponent(s)).join("/");
  const uploadData = await putSharepointFile(itemPath, contentBytes, contentType);
  return { id: uploadData.id, webUrl: uploadData.webUrl };
}

// Thumbnail image URL for a driveItem, to embed as a small clickable preview in the note.
// These Graph-issued URLs are short-lived (not the permanent webUrl) — fine for a tech
// checking the job soon after creation, but the thumbnail image itself can go stale later
// even though the click-through link keeps working.
async function getThumbnailUrl(driveId, itemId) {
  const res = await graphFetch(`/drives/${driveId}/items/${itemId}/thumbnails`);
  if (!res.ok) return null;
  const data = await res.json();
  return data.value?.[0]?.medium?.url || data.value?.[0]?.large?.url || null;
}

async function putSharepointFile(itemPath, contentBytes, contentType) {
  const driveId = await getSharepointDriveId();

  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), 20000);

  try {
    const uploadRes = await graphFetch(`/drives/${driveId}/root:/${itemPath}:/content`, {
      method: "PUT",
      headers: { "Content-Type": contentType },
      body: contentBytes,
      signal: ac.signal,
    });
    if (!uploadRes.ok) {
      const err = await uploadRes.json().catch(() => ({}));
      throw new Error(`SharePoint upload failed ${uploadRes.status}: ${err?.error?.message || JSON.stringify(err)}`);
    }
    return await uploadRes.json();
  } finally {
    clearTimeout(timer);
  }
}

// Known work order portal domains
const WORKORDER_DOMAINS = /tapihq\.com|propertytree\.com|propertyme\.com\.au|console\.net\.au|inspection\.express/i;

// Search raw HTML anchor tags — match on link text containing "work order" or "workorder".
function findWorkOrderLink(rawHtml) {
  const unescaped = rawHtml.replace(/&amp;/g, "&");

  // Extract <a href="...">text</a> pairs
  const anchors = [...unescaped.matchAll(/<a\s[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi)];

  for (const [, href, rawText] of anchors) {
    if (!href.startsWith("https://")) continue;
    const text = rawText.replace(/<[^>]*>/g, "").trim();
    if (/work\s*order/i.test(text)) return href;
  }

  // Fallback: known portal domains or "workorder" in the URL itself
  for (const [, href] of anchors) {
    if (!href.startsWith("https://")) continue;
    const dest = /safelinks\.protection\.outlook\.com/i.test(href)
      ? (() => { try { return decodeURIComponent(new URL(href).searchParams.get("url") || href); } catch { return href; } })()
      : href;
    if (/workorder/i.test(dest) || WORKORDER_DOMAINS.test(dest)) return href;
  }

  // Last resort: plain-text TAPI URL
  const plainText = unescaped.replace(/<[^>]*>/g, " ");
  const m = plainText.match(/https:\/\/url\d+\.tapihq\.com\/ls\/click\S+/i);
  if (m) return m[0].split(/[>")\s]/)[0].trim();

  return null;
}

// ================================================================
// EMAIL PROCESSING
// ================================================================
async function processMessage(message, mailbox = WORKORDERS_EMAIL, onStatus = null) {

  const attachments = message.attachments || [];
  const rawBody     = message.body?.content || "";
  const emailBodyText = (message.body?.contentType === "html" ? cleanHtml(rawBody) : rawBody).replace(/\s+/g, " ").trim();
  let   textForAI     = emailBodyText;
  let   pdfAttachment = null;

  function withEmailBody(primary) {
    return `--- WORK ORDER CONTENT (prefer this) ---\n${primary}\n\n--- EMAIL BODY (use for anything not found above) ---\n${emailBodyText}`;
  }

  // Work order link detection
  const workOrderLink = findWorkOrderLink(rawBody);

  if (workOrderLink) {
    try {
      let response = await fetch(workOrderLink, {
        redirect: "follow",
        headers: { "User-Agent": "Mozilla/5.0" },
      });

      // Inky link protection stops the redirect — bypass with confirm=True
      if (response.url.includes("shared.outlook.inky.com") && !response.url.includes("confirm=True")) {
        response = await fetch(response.url + "&confirm=True", {
          redirect: "follow",
          headers: { "User-Agent": "Mozilla/5.0" },
        });
      }

      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("pdf")) {
        const buffer  = await response.arrayBuffer();
        const pdfText = await extractPDF(new Uint8Array(buffer));
        textForAI = withEmailBody(pdfText.replace(/\s+/g, " ").trim());
      } else {
        const html     = await response.text();
        const linkText = cleanHtml(html).slice(0, 50000);
        if (linkText.length > 200) textForAI = withEmailBody(linkText);
      }
    } catch (err) {
      console.error("[email] Link fetch error:", err.message);
    }
  }

  // Work order PDF attachment — takes priority over link
  const workorderAttachment = attachments.find(a =>
    (a.name || "").toLowerCase().endsWith(".pdf")
  );

  if (workorderAttachment) {
    const attachRes  = await graphFetch(
      `/users/${mailbox}/messages/${message.id}/attachments/${workorderAttachment.id}`
    );
    const attachData = await attachRes.json();
    const data       = Uint8Array.from(atob(attachData.contentBytes), c => c.charCodeAt(0));
    pdfAttachment = { name: workorderAttachment.name, data: new Uint8Array(data) }; // copy before pdfjs detaches the buffer
    const pdfText    = await extractPDF(data);
    textForAI = withEmailBody(pdfText.replace(/\s+/g, " ").trim());
  }
  if (onStatus) await onStatus(SENDING_TO_AI_CATEGORY);

  const responseAI = await openai.responses.create({
    model: "gpt-5-mini",
    text: { format: { type: "json_object" } },
    instructions: `You are a work order extraction system for an electrical company in Australia. Today's date is ${new Date().toLocaleDateString("en-AU", { day: "2-digit", month: "2-digit", year: "numeric" })}.

CRITICAL RULES:
- tenant-name and tenant-contact come from the Tenant Details section, OR any section labelled "Contact for job access" or similar, OR an inline mention anywhere in the description/notes such as "contact tenant emma for access 0414152246" (extract "Emma" as tenant-name and "0414152246" as tenant-contact in that case). Always scan the full description/instructions text for a phrase like "contact tenant <name>" or "tenant <name> on <number>" even if there is no dedicated tenant section. Never use the Owner Details / Owner section as tenant-name, tenant-contact, or tenant-email — the owner is not the tenant, even if the PM states tenant details were not provided. If no tenant is listed anywhere (including inline) and the property is explicitly stated as vacant, set tenant-name to "Vacant" and ensure access-details captures any lockbox or key collection info. If no tenant is listed and the property is NOT stated as vacant, leave tenant-name null. Never use access details, lockbox info, or key numbers as the tenant name. If multiple tenants are listed, include ALL of them separated by commas — do not drop any.
- If the email states the tenant is vacating or moving out and the move-out date is within 7 days of today, treat the property as vacant: set tenant-name to "Vacant", set tenant-contact to null, and include in task-description that keys should be collected after the move-out date (e.g. "Collect keys after 3/7").
- access-details is ONLY physical access codes/numbers — key numbers, lockbox codes, gate codes, swipe card numbers. e.g. "Key: 1234", "Lockbox code: 56", "Gate code: 789". Do NOT include contact instructions, tenant names, safety instructions, or anything that is not a physical code or number.
- expenditure-limit is the dollar amount only — e.g. "$330". Strip any conditions, notes, or extra text after the amount. If the expenditure limit is $0 or zero, return null.
- confidence is a float 0.0–1.0 rating how confident you are in the overall extraction. 1.0 = all fields clearly present, 0.0 = guessing most fields.
- notes is any concerns, ambiguities, or flags worth mentioning — e.g. missing fields, conflicting info, unusual job details. Leave null if nothing to flag.
- tenant-contact must contain phone numbers ONLY — no names, no labels, just the numbers. Only use a number if it is explicitly and unambiguously tied to the tenant (e.g. appears in a Tenant section, is labelled "Tenant Phone"/"Tenant Mobile"/"Contact Number", or immediately follows an inline "contact tenant <name>" style phrase). If you are unsure whether a number belongs to the tenant, leave tenant-contact null. If there are multiple confirmed tenant numbers, separate with commas. Prefer mobile over home numbers. Australian numbers always start with 0 (e.g. 0412 345 678) — always include the leading 0.
- tenant-email is the tenant's email address. Only include if explicitly labelled as the tenant's email. Leave null if not present or uncertain.
- property-manager comes from the Property Manager section, OR from an Agency Details section where the manager is listed (e.g. "Manager: Jane Smith"). Use the person's name only, not the agency name.
- account-to must include ALL owners exactly as written, always in the format: owners c/o real estate.
- real-estate must always be a company or agency name — never a URL or domain. If the source contains something like "aussieproperty.com.au", convert it to a readable name (e.g. "Aussie Property") by stripping the domain extension and formatting as a proper name. If you cannot find it directly, look for it in account-to after the c/o. The sender's email address is provided at the top of the input — use the domain as an additional hint to identify real-estate if the company name is not clearly stated in the content (e.g. "noreply@raywhite.com.au" → "Ray White").
- order-number is the job/work order number.
- address is required — if it isn't clearly stated in the body/PDF content, check the email subject line (provided at the top of the input) since it often contains the property address.
- task-description must be a concise electrician job summary. If anything is listed as conditional or requires approval (e.g. "deluxe clean if approved", "AC2 if required"), include that in the description too.
- Do NOT include instructions to contact the tenant or PM for access in task-description. Contacting the tenant for access is the default assumption for every job and must not be stated.
- Key numbers in access-details are reference numbers for our existing key management system — we already hold these keys. Do NOT include any instruction to collect, pick up, or obtain keys in task-description based solely on a key number being listed in access-details.
- Do NOT guess missing fields — if missing return null.
- The text may be a structured form (with clear sections) OR plain prose in an email. Extract the same fields either way — don't return null just because sections aren't labelled.
- If the input contains both a PDF CONTENT section and an EMAIL BODY section, prefer the PDF for all fields but check the email body for anything not found in the PDF.
- For task-type: if the text explicitly mentions EC1, AC1, AC2, or ACEC1 anywhere, use that — it takes priority over everything else, even if other work is also mentioned. If a compliance check or aircon service is recommended or suggested (e.g. "recommend compliance check", "suggest an EC1"), treat it as that task type — recommendations from a PM or owner should be acted on. EXCEPTION: if the work order includes a package (EC1/AC1/AC2/ACEC1) AND additional work beyond the package, set task-type to "Real Estate General Maintenance" (or "Real Estate Aircon Maintenance" if the extra work is aircon-related) — but still set the package field to the package code. NOTE: smoke alarm testing, replacement, or supply is already included within EC1 — do not treat it as additional work beyond the package.
- package: set to "EC1", "AC1", "AC2", or "ACEC1" if the job involves one of these standard packages — even if task-type has been set to general/aircon maintenance due to extra work being present. Only return null if there is genuinely no package involved at all (e.g. a pure repair or general maintenance job with no compliance check or aircon service).

TASK TYPES:
EC1 = Electrical Compliance Check only
AC1 = Aircon Servicing only
AC2 = Deluxe Aircon Clean
ACEC1 = Aircon service AND electrical compliance check combined
Real Estate Aircon Maintenance = aircon related jobs
Real Estate General Maintenance = everything else

Return ONLY valid JSON with these exact keys:
{
  "task-type": "",
  "tenant-name": "",
  "tenant-contact": "",
  "tenant-email": "",
  "address": "",
  "task-description": "",
  "real-estate": "",
  "property-manager": "",
  "account-to": "",
  "order-number": "",
  "access-details": "",
  "expenditure-limit": "",
  "package": "EC1 | AC1 | AC2 | ACEC1 | null",
  "confidence": 0.0,
  "notes": ""
}`,
    input: `From: ${message.from?.emailAddress?.address || ""}\nSubject: ${message.subject || ""}\n\nExtract the following work order and return JSON:\n\n${textForAI}`,
  });

  const parsed = JSON.parse(responseAI.output_text);

  // Normalise address to title case in case the source had ALL CAPS suburb
  if (parsed.address) {
    parsed.address = parsed.address.replace(/,?\s*Australia\s*$/i, "").trim();
    parsed.address = parsed.address.replace(/\b\w+/g, w =>
      w.length <= 2 ? w.toUpperCase() : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()
    );
    // Expand common street type abbreviations
    const streetAbbr = {
      "\\bSt\\b": "Street", "\\bRd\\b": "Road",    "\\bAve\\b": "Avenue",
      "\\bDr\\b": "Drive",  "\\bLn\\b": "Lane",     "\\bCt\\b": "Court",
      "\\bPl\\b": "Place",  "\\bCl\\b": "Close",    "\\bCres\\b": "Crescent",
      "\\bBlvd\\b": "Boulevard", "\\bHwy\\b": "Highway", "\\bFwy\\b": "Freeway",
      "\\bTce\\b": "Terrace",    "\\bPde\\b": "Parade",  "\\bGr\\b": "Grove",
      "\\bBvd\\b": "Boulevard",  "\\bCct\\b": "Circuit", "\\bEsp\\b": "Esplanade",
    };
    for (const [abbr, full] of Object.entries(streetAbbr)) {
      parsed.address = parsed.address.replace(new RegExp(abbr, "gi"), full);
    }
  }

  // If AI returned a URL/domain for real-estate, strip the TLD and convert to a readable name
  if (parsed["real-estate"] && /\.(com\.au|net\.au|org\.au|com|net|org|io|au)$/i.test(parsed["real-estate"])) {
    parsed["real-estate"] = parsed["real-estate"]
      .replace(/\.(com\.au|net\.au|org\.au|com|net|org|io|au)$/i, "")
      .replace(/[-_.]/g, " ")
      .replace(/\b\w/g, c => c.toUpperCase())
      .trim();
  }

  // Ignore $0 expenditure limits
  if (parsed["expenditure-limit"] && /^\$?\s*0+(\.0+)?\s*$/.test(parsed["expenditure-limit"].trim())) {
    parsed["expenditure-limit"] = null;
  }

  // Australian mobile numbers are 10 digits starting with 0 — if AI drops the leading 0, restore it
  if (parsed["tenant-contact"]) {
    parsed["tenant-contact"] = parsed["tenant-contact"]
      .split(",")
      .map(n => {
        const digits = n.replace(/\D/g, "");
        return digits.length === 9 && digits.startsWith("4") ? "0" + n.trim() : n.trim();
      })
      .join(", ");
  }

  const emailMeta = {
    from:    message.from?.emailAddress?.address || null,
    to:      (message.toRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(", ") || null,
    subject: message.subject || null,
  };

  // Download image attachments in parallel
  const imageAttachments = (await Promise.all(
    attachments
      .filter(a => !/inky/i.test(a.name || ""))
      .filter(a => /\.(jpe?g|png|gif|bmp|webp)$/i.test(a.name || "") || (a.contentType || "").startsWith("image/"))
      .map(async a => {
        try {
          const attRes  = await graphFetch(`/users/${mailbox}/messages/${message.id}/attachments/${a.id}`);
          const attData = await attRes.json();
          const imgData = Uint8Array.from(atob(attData.contentBytes), c => c.charCodeAt(0));
          console.log("[email] Image attachment downloaded:", a.name);
          return { name: a.name, data: imgData, contentType: a.contentType || "image/jpeg" };
        } catch (err) {
          console.warn("[email] Failed to download image attachment:", a.name, err.message);
          return null;
        }
      })
  )).filter(Boolean);

  return { result: parsed, rawEmail: rawBody, pdfAttachment, imageAttachments, emailMeta };
}

// Best-effort alert email to Brandon — used for both per-job warnings and
// poll-loop-level failures. Never throws; a failure to send is only logged.
async function sendAlertEmail(subject, htmlBody, context) {
  try {
    await graphFetch(`/users/${WORKORDERS_EMAIL}/sendMail`, {
      method: "POST",
      body: JSON.stringify({
        message: {
          subject,
          toRecipients: [{ emailAddress: { address: BRANDON_EMAIL } }],
          body: { contentType: "HTML", content: htmlBody },
        },
        saveToSentItems: false,
      }),
    });
    console.log(`[alert] Email sent — ${context}`);
  } catch (err) {
    console.error(`[alert] Failed to send email — ${context}:`, err.message);
  }
}

// POLL LOOP
// ================================================================
async function setJobStatus(mailbox, messageId, currentCategories, newStatus) {
  const updated = [
    ...currentCategories.filter(c => !STATUS_CATEGORIES.includes(c)),
    newStatus,
  ];
  await graphFetch(`/users/${mailbox}/messages/${messageId}`, {
    method: "PATCH",
    body: JSON.stringify({ categories: updated }),
  });
  return updated;
}

// Look for a sibling message already tagged "Job created - X" or "Existing job - X"
// elsewhere in the same reply thread — replies to an already-processed work
// order get their own "Bara AI" tag but shouldn't trigger their own job.
// Mailbox-wide, not Inbox-scoped — existing filing rules (e.g. "move mail from
// this client into its own subfolder") can relocate a message out of Inbox
// well before a later reply in the same thread gets processed.
async function findJobTagInThread(mailbox, conversationId, excludeMessageId) {
  const filter = encodeURIComponent(`conversationId eq '${conversationId}'`);
  const res = await graphFetch(
    `/users/${mailbox}/messages?$filter=${filter}&$select=id,categories&$top=25`
  );
  const data = await res.json();
  if (!res.ok) return null;
  for (const m of (data.value || [])) {
    if (m.id === excludeMessageId) continue;
    const tag = (m.categories || []).find(c => c.startsWith("Job created") || c.startsWith("Existing job"));
    if (tag) return tag;
  }
  return null;
}

// Stamp the same job tag onto every other message already in the thread, so
// anyone looking at any message in the conversation sees the job status
// immediately instead of just the "Bara AI" trigger tag.
async function tagWholeConversation(mailbox, conversationId, excludeMessageId, tag) {
  if (!conversationId) return;
  try {
    const filter = encodeURIComponent(`conversationId eq '${conversationId}'`);
    const res = await graphFetch(
      `/users/${mailbox}/messages?$filter=${filter}&$select=id,categories&$top=25`
    );
    const data = await res.json();
    if (!res.ok) return;
    for (const m of (data.value || [])) {
      if (m.id === excludeMessageId) continue;
      const categories = [
        ...(m.categories || []).filter(c => !STATUS_CATEGORIES.includes(c) && !c.startsWith("Job created") && !c.startsWith("Existing job")),
        tag,
      ];
      try {
        await graphFetch(`/users/${mailbox}/messages/${m.id}`, {
          method: "PATCH",
          body: JSON.stringify({ categories }),
        });
      } catch (err) {
        console.warn("[poll] Failed to tag conversation message:", m.id, err.message);
      }
    }
  } catch (err) {
    console.warn("[poll] tagWholeConversation failed:", err.message);
  }
}

async function pollInbox(mailbox) {
  const filter = encodeURIComponent(
    `categories/any(c:c eq '${TRIGGER_CATEGORY}')` +
    ` and not categories/any(c:c eq '${CLIENT_NOT_FOUND_CATEGORY}')` +
    ` and not categories/any(c:c eq '${NO_ADDRESS_CATEGORY}')` +
    ` and not categories/any(c:c eq '${PROCESSING_CATEGORY}')` +
    ` and not categories/any(c:c eq '${FAILED_CATEGORY}')` +
    ` and not categories/any(c:c eq '${READING_EMAIL_CATEGORY}')` +
    ` and not categories/any(c:c eq '${SENDING_TO_AI_CATEGORY}')` +
    ` and not categories/any(c:c eq '${CREATING_JOB_CATEGORY}')`
  );

  const res  = await graphFetch(
    `/users/${mailbox}/mailFolders/inbox/messages` +
    `?$filter=${filter}` +
    `&$select=id,subject,body,categories,from,toRecipients,conversationId,receivedDateTime` +
    `&$expand=attachments($select=id,name,contentType,size)` +
    `&$top=10`
  );
  const data = await res.json();
  if (!res.ok) throw new Error(`Graph API error ${res.status}: ${JSON.stringify(data?.error || data)}`);

  // Newest first — when a thread has multiple untagged messages, this ensures the
  // most recent one (which quotes everything underneath it) is the one processed
  // and posted to Aroflo; the earlier ones get caught as duplicates by the thread
  // check below since the newest is tagged first.
  const messages = (data.value || [])
    .filter(m => !m.categories.some(c => c.startsWith("Job created")))
    .sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
  console.log(`[poll] (${mailbox}): ${messages.length} email(s) found`);

  for (const message of messages) {
    const siblingTag = message.conversationId
      ? await findJobTagInThread(mailbox, message.conversationId, message.id)
      : null;
    if (siblingTag) {
      const jobNumber = siblingTag.split(" - ")[1] || siblingTag;
      console.log(`[poll] Reply to already-processed thread — tagging as duplicate of ${siblingTag}:`, message.subject);
      await graphFetch(`/users/${mailbox}/messages/${message.id}`, {
        method: "PATCH",
        body: JSON.stringify({ categories: [...message.categories.filter(c => !STATUS_CATEGORIES.includes(c)), `Existing job - ${jobNumber}`] }),
      });
      continue;
    }

    let currentCategories = await setJobStatus(mailbox, message.id, message.categories, READING_EMAIL_CATEGORY);
    console.log("[poll] Reading:", message.subject);

    try {
      const onStatus = async (status) => {
        currentCategories = await setJobStatus(mailbox, message.id, currentCategories, status);
      };

      const { result, rawEmail, pdfAttachment, imageAttachments, emailMeta } = await processMessage(message, mailbox, onStatus);

      // Pre-job validation warnings — checked before Aroflo job creation
      const preWarnings = [];
      if (!result["property-manager"]) {
        preWarnings.push({ tag: "No PM in email", detail: "Property manager not listed in the work order" });
      }
      if (!result["tenant-name"] && !result["tenant-contact"]) {
        preWarnings.push({ tag: "No tenant info", detail: "No tenant name or contact in the work order, and property not stated as vacant" });
      }

      await onStatus(CREATING_JOB_CATEGORY);
      const { jobNumber, warnings: jobWarnings } = await createArofloJob(result, rawEmail, pdfAttachment, emailMeta, imageAttachments);
      logAiOutput(result, message.subject).catch(err => console.warn("[airtable] logAiOutput:", err.message));
      logActivity("Job created", jobNumber).catch(err => console.warn("[airtable] logActivity:", err.message));
      const allWarnings = [...preWarnings, ...jobWarnings];

      // Always apply job tag to prevent re-processing; add a specific tag for each distinct
      // failure type (multiple failures of the same kind — e.g. several failed photo
      // uploads — share one tag rather than repeating it).
      const jobTag = `Job created - ${jobNumber}`;
      const warningTags = [...new Set(allWarnings.map(w => w.tag))];
      const finalCategories = [
        ...currentCategories.filter(c => !STATUS_CATEGORIES.includes(c)),
        jobTag,
        ...warningTags,
      ];
      await graphFetch(`/users/${mailbox}/messages/${message.id}`, {
        method: "PATCH",
        body: JSON.stringify({ categories: finalCategories }),
      });
      currentCategories = finalCategories;
      tagWholeConversation(mailbox, message.conversationId, message.id, jobTag)
        .catch(err => console.warn("[poll] tagWholeConversation:", err.message));

      if (allWarnings.length > 0) {
        console.warn("[job] Created with issues:", allWarnings.map(w => w.tag));
        const warningLines = allWarnings.map(w => `<li style="margin:4px 0;font-family:sans-serif;font-size:14px"><strong>${escapeHtml(w.tag)}:</strong> ${escapeHtml(w.detail)}</li>`).join("");
        await sendAlertEmail(
          `Action required — Job ${jobNumber} created with issues`,
          `<p style="font-family:sans-serif;font-size:14px">Job <strong>${escapeHtml(jobNumber)}</strong> was created in Aroflo but the following need attention:</p><ul>${warningLines}</ul><p style="font-family:sans-serif;font-size:12px;color:#888">Original email: ${escapeHtml(message.subject)}</p>`,
          `job ${jobNumber}`
        );
      } else {
        console.log("[poll] Tagged as done:", message.subject);
      }
    } catch (err) {
      console.error("[poll] Error processing message:", message.subject, err.message);
      if (err.message.startsWith("Client not found")) {
        await setJobStatus(mailbox, message.id, currentCategories, CLIENT_NOT_FOUND_CATEGORY);
        console.log("[poll] Tagged as client not found:", message.subject);
      } else if (err.message.startsWith("No address found")) {
        await setJobStatus(mailbox, message.id, currentCategories, NO_ADDRESS_CATEGORY);
        console.log("[poll] Tagged as no address:", message.subject);
      } else {
        // Tag as "Failed" — remove it in Outlook to retry
        await setJobStatus(mailbox, message.id, currentCategories, FAILED_CATEGORY);
        console.log("[poll] Tagged as failed:", message.subject);
      }
    }
  }
}

async function pollEmails() {
  if (pollRunning) {
    console.log("[poll] Skipped — previous run still in progress");
    return;
  }
  pollRunning = true;
  try {
    await pollInbox(WORKORDERS_EMAIL);
  } catch (err) {
    console.error("[poll] Poll error:", err);
    await sendAlertEmail(
      "Action required — work order poll failed",
      `<p style="font-family:sans-serif;font-size:14px">The email poll loop failed with an error and did not finish checking for new work orders:</p><pre style="font-family:monospace;font-size:12px;background:#f5f5f5;padding:8px;border-radius:4px">${escapeHtml(err.message || String(err))}</pre><p style="font-family:sans-serif;font-size:12px;color:#888">This will retry automatically on the next poll cycle.</p>`,
      "poll loop failure"
    );
  } finally {
    pollRunning = false;
  }
}

// ================================================================
// TEMP: Dump client cache as JSON — GET /clients
// ================================================================
app.get("/clients", requireApiKey, (req, res) => {
  res.json({
    count: clientCache.size,
    lastLoaded: clientCacheLastLoaded,
    clients: [...clientCache.values()].map(c => ({ clientid: c.clientid, clientname: c.clientname })),
  });
});

app.get("/find-client", requireApiKey, async (req, res) => {
  const name = req.query.name;
  if (!name) return res.status(400).json({ error: "Pass ?name=..." });
  const result = await findClient(name);
  const key = name.toLowerCase();
  const exactCacheHit = clientCache.get(key);
  const partialMatches = [...clientCache.values()]
    .filter(c => c.clientname.toLowerCase().includes(key) || key.includes(c.clientname.toLowerCase()))
    .map(c => c.clientname);
  res.json({ result, exactCacheHit: exactCacheHit?.clientname || null, partialMatches });
});

// ================================================================
// AROFLO WEBHOOK — new client created
// ================================================================
app.post("/aroflo-webhook", requireApiKey, express.json(), async (req, res) => {
  console.log("[webhook] Aroflo webhook received:", JSON.stringify(req.body));
  res.sendStatus(200);
  await loadClientCache();
});

// ================================================================
// START
// ================================================================
app.listen(process.env.PORT || 3000, async () => {
  console.log(`[startup] Server running — deployed ${new Date().toISOString()}`);
  // Wait for the client cache before polling — otherwise a work order sitting
  // in the inbox at deploy time gets processed against an empty cache and
  // falls back to a less reliable single-candidate live API lookup.
  await loadClientCache();
  pollEmails();
  setInterval(pollEmails, POLL_INTERVAL_MS);
});

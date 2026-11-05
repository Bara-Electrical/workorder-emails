// Evap and two-storey check for aircon service work orders (Brandon, 5 Nov 2026).
//
// Work orders often call an evap service "aircon service", and the property manager rarely
// says the house is two storey. Both change the job: an evap is EV1/EVEC1, not AC1/ACEC1,
// and a two-storey evap is priced separately (2 Story Evap Service) and needs the PM's OK.
// So for each aircon service job, one AI call looks at two things:
//
//   1. The property's sale or rental listing (realestate.com.au, Domain, agency sites),
//      found with OpenAI's web search. realestate.com.au blocks automated page loads, so
//      the listing is read through search, the same way the quoting agent does it.
//   2. An aerial photo of the roof: Landgate's free public WA imagery first (about 7 cm a
//      pixel over Perth, at least 400 days old), else HERE satellite with the HERE key the
//      service already uses for geocoding. The photo is sent to the model and never kept.
//
// The answer only adds warning tags and a line in the description. It never changes the
// task type: a house can have an evap and splits, and the office decides. "Unknown" adds
// nothing, so a vague listing or a blurry roof does not flag the job.
//
// Its own key, OPENAI_12_WORKORDER_LISTING, so OpenAI's usage page shows what it costs.
// Without the key the check is skipped.

export const PROPERTY_LOOKUP_MODEL = process.env.PROPERTY_LOOKUP_MODEL || "gpt-5-mini";

const WEB_SEARCH_TOOL = { type: "web_search", user_location: { type: "approximate", country: "AU", region: "Western Australia", city: "Perth" } };

// The roof photo's footprint: a 45 m square holds a Perth block and a little of each
// neighbour, so the house in the middle is clear.
const BOX_METRES = 45;
const IMAGE_PX = 1024;
const FETCH_TIMEOUT_MS = 20_000;

const LANDGATE_EXPORT = "https://services.slip.wa.gov.au/public/rest/services/SLIP_Public_Services/Locate/MapServer/export";

export const INSTRUCTIONS = `You check one Perth, Western Australia residential property for an electrical and air conditioning business, before a tech is sent to service its air conditioning.

Answer two questions about the property at the given address:
A. Does it have an EVAPORATIVE cooler (an "evap": Breezair, Coolair, Braemar, Bonaire, Coolbreeze, Climate Wizard; listings say "evaporative air conditioning", "evaporative cooling", "evap cooling", "ducted evaporative")? Reverse cycle, split system, ducted reverse cycle or refrigerated air are NOT evaporative.
B. Is the house single storey or two storey (double storey, two-storey, split level with an upper floor)?

Use two sources:
1. Web search. Search for the exact street address and read the property's own sale or rental listing (realestate.com.au, domain.com.au, allhomes, reiwa.com.au, an agency's site). Only use a listing whose street number, street and suburb all match the address. A different unit or a neighbour's house does not count. If you find no matching listing, say so.
2. The attached aerial photo, taken straight down, with the property at the CENTRE. Ignore neighbouring houses. A Perth evaporative cooler sits ON the roof: a squat square or rectangular box about 1.0 to 1.5 m across, usually beige, grey or cream, often near the ridge, with a short shadow. Do not confuse it with solar hot water (flat panel with a long cylinder), solar PV (dark flat rectangles), skylights (flush with the roof), satellite dishes (round), whirlybirds (small round cowls under 0.5 m) or a refrigerated condenser (bigger, with a round fan grille). Storeys cannot be read reliably from straight above: use the photo for storeys only when it is obvious, otherwise "unknown".

Be strict. "unknown" is the right answer whenever the evidence is not clear. A listing that does not mention cooling at all is "unknown" for the evap, not "no".

Return JSON only, with no other text:
{
  "listing": {"found": true|false, "url": "the listing's link or null", "evap": "yes"|"no"|"unknown", "storeys": 1|2|"unknown", "evidence": "the listing's own words, short"},
  "roof": {"usable": true|false, "evap": "yes"|"no"|"unknown", "confidence": 0.0-1.0, "evidence": "what you see and where in the photo"}
}`;

function bbox(lat, lon, metres) {
  const dLat = metres / 2 / 111320;
  const dLon = metres / 2 / (111320 * Math.cos((lat * Math.PI) / 180));
  return [lon - dLon, lat - dLat, lon + dLon, lat + dLat];
}

export function landgateUrl(lat, lon) {
  const [xmin, ymin, xmax, ymax] = bbox(lat, lon, BOX_METRES);
  const p = new URLSearchParams({
    bbox: `${xmin},${ymin},${xmax},${ymax}`, bboxSR: "4326", imageSR: "3857",
    size: `${IMAGE_PX},${IMAGE_PX}`, format: "jpg", f: "image",
  });
  return `${LANDGATE_EXPORT}?${p}`;
}

// Zoom 20 is about 15 cm a pixel at Perth's latitude: enough to see a 1 m box.
export function hereSatelliteUrl(lat, lon, apiKey) {
  return `https://image.maps.hereapi.com/mia/v3/base/mc/center:${lat},${lon};zoom=20/${IMAGE_PX}x${IMAGE_PX}/jpeg?style=satellite.day&apiKey=${encodeURIComponent(apiKey)}`;
}

// An ArcGIS or HERE error can come back as HTTP 200 with JSON or XML, so the content type
// and a plausible size are both checked before a reply counts as a photo.
async function fetchImage(url, fetchImpl) {
  const res = await fetchImpl(url, { signal: AbortSignal.timeout(FETCH_TIMEOUT_MS) });
  const type = res.headers.get("content-type") || "";
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  if (!type.startsWith("image/")) throw new Error(`not an image (${type || "no content type"})`);
  const data = Buffer.from(await res.arrayBuffer());
  if (data.length < 5000) throw new Error(`${data.length} bytes, probably blank`);
  return { contentType: type.split(";")[0], data };
}

export async function fetchRoofPhoto({ lat, lon, hereApiKey, fetchImpl = fetch }) {
  const tried = [];
  try {
    return { source: "Landgate", ...(await fetchImage(landgateUrl(lat, lon), fetchImpl)) };
  } catch (err) { tried.push(`Landgate: ${err.message}`); }
  if (hereApiKey) {
    try {
      return { source: "HERE", ...(await fetchImage(hereSatelliteUrl(lat, lon, hereApiKey), fetchImpl)) };
    } catch (err) { tried.push(`HERE: ${err.message}`); }
  }
  console.warn(`[property] No roof photo: ${tried.join("; ")}`);
  return null;
}

const asEvap = (v) => (v === "yes" || v === true ? "yes" : v === "no" || v === false ? "no" : "unknown");
const asStoreys = (v) => (v === 1 || v === "1" ? 1 : v === 2 || v === "2" ? 2 : "unknown");

// The model's JSON, made safe: anything it got wrong or left out reads as "unknown".
export function parseLookup(text) {
  let body;
  try {
    const raw = String(text || "");
    body = JSON.parse(raw.slice(raw.indexOf("{"), raw.lastIndexOf("}") + 1));
  } catch {
    return null;
  }
  const l = body?.listing || {};
  const r = body?.roof || {};
  const url = typeof l.url === "string" && /^https?:\/\//i.test(l.url) ? l.url.trim() : null;
  const found = l.found === true;
  return {
    listing: {
      found,
      url: found ? url : null,
      evap: found ? asEvap(l.evap) : "unknown",
      storeys: found ? asStoreys(l.storeys) : "unknown",
      evidence: typeof l.evidence === "string" ? l.evidence.trim().slice(0, 300) : "",
    },
    roof: {
      usable: r.usable === true,
      evap: r.usable === true ? asEvap(r.evap) : "unknown",
      confidence: Number.isFinite(Number(r.confidence)) ? Number(r.confidence) : 0,
      evidence: typeof r.evidence === "string" ? r.evidence.trim().slice(0, 300) : "",
    },
  };
}

// The roof photo only counts when the model is fairly sure: a beige box on a beige tile
// roof is easy to imagine.
export const ROOF_CONFIDENCE = 0.7;

// What the lookup means for the job: warning tags and one description line.
// pkg is the job's package (AC1, AC2, ACEC1, EV1, EVEC1).
export function lookupFindings(lookup, pkg) {
  if (!lookup) return { warnings: [], line: null };
  const { listing, roof } = lookup;
  const roofEvap = roof.evap === "yes" && roof.confidence >= ROOF_CONFIDENCE;
  const roofNoEvap = roof.evap === "no" && roof.confidence >= ROOF_CONFIDENCE;
  const evapPackage = pkg === "EV1" || pkg === "EVEC1";
  const evapFound = listing.evap === "yes" || roofEvap;
  const twoStorey = listing.storeys === 2;
  const warnings = [];

  if (evapFound && !evapPackage) {
    const where = [
      listing.evap === "yes" ? `the listing says evaporative${listing.evidence ? ` ("${listing.evidence}")` : ""}` : null,
      roofEvap ? `the roof photo shows an evap (${roof.evidence || "box on the roof"})` : null,
    ].filter(Boolean).join(" and ");
    warnings.push({ tag: "Evap found", detail: `Booked as ${pkg || "an aircon service"}, but ${where}${listing.url ? ` ${listing.url}` : ""}. Check whether it should be EV1/EVEC1.` });
  }
  // Two storey only matters for an evap: it sits on the roof (2 Story Evap Service, call the PM).
  if (twoStorey && (evapPackage || evapFound)) {
    warnings.push({ tag: "Two storey", detail: `The listing says two storey${listing.evidence ? ` ("${listing.evidence}")` : ""}${listing.url ? ` ${listing.url}` : ""}. Evap on a two-storey roof: 2 Story Evap Service, call the PM before going up.` });
  }

  const bits = [];
  if (evapFound) bits.push("evap");
  else if (evapPackage && (listing.evap === "no" || roofNoEvap)) bits.push("no evap seen");
  if (twoStorey) bits.push("2 storey");
  else if (listing.storeys === 1) bits.push("single storey");
  if (!bits.length) return { warnings, line: null };
  const sources = [listing.found ? "listing" : null, roof.usable ? "roof photo" : null].filter(Boolean).join(" + ");
  return { warnings, line: { text: `Property check (${sources}): ${bits.join(", ")}`, url: listing.url } };
}

// → the parsed lookup or null. Never throws: a failed check leaves the job as it was.
export async function lookupProperty({ openai, address, coords, hereApiKey, model = PROPERTY_LOOKUP_MODEL, fetchImpl = fetch }) {
  if (!openai || !address) return null;
  try {
    const photo = coords ? await fetchRoofPhoto({ lat: coords.lat, lon: coords.lon, hereApiKey, fetchImpl }) : null;
    const content = [{ type: "input_text", text: `Address: ${address}\n${photo ? `Aerial photo attached (${photo.source}).` : "No aerial photo: answer roof.usable false."}` }];
    if (photo) content.push({ type: "input_image", image_url: `data:${photo.contentType};base64,${photo.data.toString("base64")}`, detail: "high" });
    const response = await openai.responses.create({
      model,
      instructions: INSTRUCTIONS,
      tools: [WEB_SEARCH_TOOL],
      input: [{ role: "user", content }],
    });
    const lookup = parseLookup(response.output_text);
    if (lookup) lookup.photoSource = photo?.source ?? null;
    return lookup;
  } catch (err) {
    console.warn("[property] Lookup failed:", err.message);
    return null;
  }
}

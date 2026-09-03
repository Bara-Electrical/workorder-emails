#!/usr/bin/env node
// Standalone experiment — NOT wired into the email pipeline.
//
// Question it answers: given only an address, can free Landgate aerial imagery plus a
// vision model tell us (a) whether the property has a roof-mounted evaporative cooler
// and (b) whether it is single or double storey — accurately enough to put in front of
// a tech?
//
// Landgate's imagery hosts are blocked from the Claude sandbox by egress policy, so this
// has to be run from somewhere with open outbound HTTPS (a laptop, or the Railway box).
//
//   node tools/property-lookup-test.js "3 Cable Cove, Mosman Park WA 6012"
//   node tools/property-lookup-test.js --lat -32.0123 --lon 115.7654
//   node tools/property-lookup-test.js --batch tools/property-test-set.csv
//   node tools/property-lookup-test.js --capabilities     # list real Landgate layer names
//
// Needs HERE_API_KEY (address mode only) and OPENAI_API_KEY — the same two the app
// already uses. Saves every tile it fetches to ./property-test-out/ so the imagery can
// be eyeballed next to the model's answer; a wrong answer on a tile that clearly shows
// the evap is a prompt problem, a wrong answer on a blurry or misplaced tile is not.

import fs from "node:fs/promises";
import path from "node:path";
import OpenAI from "openai";

const OUT_DIR = "property-test-out";

// Candidate imagery endpoints, best first. Which of these actually serves usable
// resolution over Perth metro is exactly what this script is meant to find out, so it
// tries each in turn and reports which one answered rather than assuming.
const SOURCES = [
  {
    name: "slip-locate",
    kind: "arcgis",
    url: "https://services.slip.wa.gov.au/public/rest/services/SLIP_Public_Services/Locate/MapServer/export",
    note: "Free whole-of-state base layer. Imagery is deliberately >=400 days old.",
  },
  {
    name: "landgate-wms-public",
    kind: "wms",
    url: "https://www2.landgate.wa.gov.au/ows/wmspublicimagery",
    layer: process.env.LANDGATE_WMS_LAYER || "public:aerial_photography_metro",
    note: "Public WMS. Layer name is a guess — run --capabilities to get the real list.",
  },
];

function parseArgs(argv) {
  const args = { boxMetres: 45, size: 1024 };
  const rest = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--lat") args.lat = Number(argv[++i]);
    else if (a === "--lon") args.lon = Number(argv[++i]);
    else if (a === "--batch") args.batch = argv[++i];
    else if (a === "--box") args.boxMetres = Number(argv[++i]);
    else if (a === "--size") args.size = Number(argv[++i]);
    else if (a === "--capabilities") args.capabilities = true;
    else rest.push(a);
  }
  args.address = rest.join(" ").trim() || null;
  return args;
}

// Same HERE geocoder the app already uses in geocodeAddress(), so this test exercises
// the real coordinate quality rather than a better one we wouldn't have in production.
// resultType/houseNumberType come back too: a "PA" (point address) is a rooftop-grade
// hit, anything else means HERE interpolated along the street and the tile may well be
// centred on next door.
async function geocode(address) {
  const key = process.env.HERE_API_KEY;
  if (!key) throw new Error("HERE_API_KEY not set (or pass --lat/--lon)");
  const url = "https://geocode.search.hereapi.com/v1/geocode?q=" +
    encodeURIComponent(address) + "&in=countryCode:AUS&apiKey=" + key;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HERE geocode HTTP ${res.status}`);
  const item = (await res.json())?.items?.[0];
  if (!item?.position) throw new Error("HERE returned no position");
  return {
    lat: item.position.lat,
    lon: item.position.lng,
    label: item.title,
    resultType: item.resultType,
    houseNumberType: item.houseNumberType,
    rooftop: item.resultType === "houseNumber" && item.houseNumberType === "PA",
  };
}

function bbox(lat, lon, metres) {
  const dLat = metres / 2 / 111320;
  const dLon = metres / 2 / (111320 * Math.cos((lat * Math.PI) / 180));
  return [lon - dLon, lat - dLat, lon + dLon, lat + dLat];
}

function tileUrl(source, lat, lon, boxMetres, size) {
  const [xmin, ymin, xmax, ymax] = bbox(lat, lon, boxMetres);
  if (source.kind === "arcgis") {
    const p = new URLSearchParams({
      bbox: `${xmin},${ymin},${xmax},${ymax}`,
      bboxSR: "4326", imageSR: "3857",
      size: `${size},${size}`,
      format: "png32", transparent: "false", f: "image",
    });
    return `${source.url}?${p}`;
  }
  const p = new URLSearchParams({
    SERVICE: "WMS", VERSION: "1.1.1", REQUEST: "GetMap",
    LAYERS: source.layer, STYLES: "", SRS: "EPSG:4326",
    BBOX: `${xmin},${ymin},${xmax},${ymax}`,
    WIDTH: String(size), HEIGHT: String(size), FORMAT: "image/png",
  });
  return `${source.url}?${p}`;
}

// A WMS/ArcGIS error comes back as HTTP 200 with an XML or JSON body, so checking the
// status code alone would happily hand a page of XML to the vision model. Require an
// image content-type and a plausible size before trusting the response.
async function fetchTile(lat, lon, boxMetres, size) {
  const failures = [];
  for (const source of SOURCES) {
    const url = tileUrl(source, lat, lon, boxMetres, size);
    try {
      const res = await fetch(url);
      const type = res.headers.get("content-type") || "";
      if (!res.ok) { failures.push(`${source.name}: HTTP ${res.status}`); continue; }
      if (!type.startsWith("image/")) {
        const body = (await res.text()).slice(0, 300).replace(/\s+/g, " ");
        failures.push(`${source.name}: got ${type || "no content-type"} — ${body}`);
        continue;
      }
      const data = Buffer.from(await res.arrayBuffer());
      if (data.length < 5000) { failures.push(`${source.name}: ${data.length}B — probably blank`); continue; }
      return { source: source.name, contentType: type, data, url };
    } catch (err) {
      failures.push(`${source.name}: ${err.message}`);
    }
  }
  throw new Error("No imagery source answered:\n  " + failures.join("\n  "));
}

async function capabilities() {
  const src = SOURCES.find(s => s.kind === "wms");
  const res = await fetch(`${src.url}?SERVICE=WMS&VERSION=1.1.1&REQUEST=GetCapabilities`);
  const xml = await res.text();
  const names = [...xml.matchAll(/<Name>([^<]+)<\/Name>/g)].map(m => m[1]);
  console.log(`WMS ${src.url} — HTTP ${res.status}`);
  console.log(names.length ? names.map(n => "  " + n).join("\n") : "  (no <Name> elements found)");
  const arc = SOURCES.find(s => s.kind === "arcgis");
  const meta = await fetch(arc.url.replace(/\/export$/, "") + "?f=json");
  console.log(`\nArcGIS ${arc.url} — HTTP ${meta.status}`);
  console.log("  " + (await meta.text()).slice(0, 600).replace(/\s+/g, " "));
}

// Deliberately asks for evidence and a confidence per field, not a bare yes/no. The whole
// question here is whether the answer is trustworthy enough to show a tech, and a model
// that says "beige box on the ridge, 1.2m across, clear shadow" is checkable against the
// saved tile in a way that a lone "true" is not.
const PROMPT = `This is a nadir (straight-down) aerial photo of a single residential property in Perth, Western Australia, at roughly 10cm per pixel. The property of interest is the building at the CENTRE of the image; ignore neighbouring properties.

Answer two questions.

1. EVAPORATIVE COOLER. A Perth evaporative cooler sits ON the roof: a squat rectangular or square box roughly 1.0-1.5m across (10-15 pixels here), usually beige, grey or cream, often near the ridge, normally casting a short distinct shadow. Do NOT confuse it with:
   - solar hot water (a large flat panel with a long horizontal cylinder above it)
   - solar PV (large flat very dark rectangles, usually several together)
   - skylights (small, flush with the roof plane, often bright or translucent)
   - satellite dishes (round, on an angled mount)
   - whirlybirds/roof vents (small round cowls, typically under 0.5m, often several in a row)
   - a ducted refrigerative condenser (larger, with a visible circular fan grille on top)

2. STOREYS. From nadir imagery this is an inference, not an observation. Use shadow length cast by the building relative to its footprint, the apparent height of the eaves, and roof complexity. Say "unknown" rather than guessing when the shadow is unclear or the sun angle is high.

Return JSON only:
{
  "evap": {"present": true|false|"unknown", "confidence": 0.0-1.0, "evidence": "what you actually see and where in the frame"},
  "storeys": {"estimate": 1|2|"unknown", "confidence": 0.0-1.0, "evidence": "the shadow/height cue you used"},
  "roof_material": "tile"|"metal"|"unknown",
  "roof_pitch": "shallow"|"steep"|"unknown",
  "other_roof_plant": ["solar_pv", "solar_hot_water", "whirlybird", "ducted_condenser", "satellite_dish"],
  "image_usable": true|false,
  "notes": "anything that would make this answer unreliable — blur, tree cover, the centre building being ambiguous"
}`;

async function analyse(openai, tile) {
  const res = await openai.responses.create({
    model: "gpt-5-mini",
    text: { format: { type: "json_object" } },
    input: [{
      role: "user",
      content: [
        { type: "input_text", text: PROMPT },
        { type: "input_image", image_url: `data:${tile.contentType};base64,${tile.data.toString("base64")}`, detail: "high" },
      ],
    }],
  });
  return JSON.parse(res.output_text);
}

function slug(s) { return s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60); }

async function runOne(openai, { address, lat, lon, boxMetres, size }) {
  let geo = null;
  if (lat === undefined || lon === undefined) {
    geo = await geocode(address);
    ({ lat, lon } = geo);
  }
  const tile = await fetchTile(lat, lon, boxMetres, size);
  await fs.mkdir(OUT_DIR, { recursive: true });
  const file = path.join(OUT_DIR, `${slug(address || `${lat}_${lon}`)}.png`);
  await fs.writeFile(file, tile.data);
  const verdict = await analyse(openai, tile);
  return { address, lat, lon, geo, tile: { source: tile.source, file, bytes: tile.data.length }, verdict };
}

// Batch mode takes ground truth alongside each address so the run ends in a hit rate
// rather than a pile of plausible-looking JSON. CSV: address,expected_evap,expected_storeys
// e.g.  "3 Cable Cove, Mosman Park WA 6012",yes,2
async function runBatch(openai, file, opts) {
  const lines = (await fs.readFile(file, "utf8")).split("\n").map(l => l.trim()).filter(l => l && !l.startsWith("#"));
  const rows = [];
  for (const line of lines) {
    const m = line.match(/^"([^"]+)"\s*,\s*([^,]*)\s*,\s*(.*)$/) || line.match(/^([^,]+),\s*([^,]*)\s*,\s*(.*)$/);
    if (!m) { console.warn(`skipping unparseable line: ${line}`); continue; }
    const [, address, expEvap, expStoreys] = m;
    try {
      const r = await runOne(openai, { address, ...opts });
      const gotEvap = r.verdict?.evap?.present;
      const gotStoreys = r.verdict?.storeys?.estimate;
      rows.push({
        address,
        expected_evap: expEvap.trim(), got_evap: String(gotEvap),
        evap_conf: r.verdict?.evap?.confidence,
        expected_storeys: expStoreys.trim(), got_storeys: String(gotStoreys),
        storeys_conf: r.verdict?.storeys?.confidence,
        rooftop_geocode: r.geo?.rooftop ?? "n/a",
      });
      console.log(`  ✓ ${address}`);
    } catch (err) {
      rows.push({ address, error: err.message });
      console.log(`  ✗ ${address} — ${err.message}`);
    }
  }
  console.table(rows);

  const scored = rows.filter(r => !r.error);
  const norm = v => ({ yes: "true", y: "true", true: "true", no: "false", n: "false", false: "false" }[String(v).toLowerCase()] ?? String(v).toLowerCase());
  const evapHits = scored.filter(r => norm(r.expected_evap) === norm(r.got_evap)).length;
  const storeyHits = scored.filter(r => norm(r.expected_storeys) === norm(r.got_storeys)).length;
  console.log(`\nevap:    ${evapHits}/${scored.length}`);
  console.log(`storeys: ${storeyHits}/${scored.length}`);
  console.log(`tiles saved to ./${OUT_DIR}/ — check the misses against the imagery before blaming the prompt`);
}

const args = parseArgs(process.argv.slice(2));

if (args.capabilities) {
  await capabilities();
} else if (args.batch) {
  await runBatch(new OpenAI({ apiKey: process.env.OPENAI_API_KEY }), args.batch, { boxMetres: args.boxMetres, size: args.size });
} else if (!args.address && args.lat === undefined) {
  console.error("usage: node tools/property-lookup-test.js \"<address>\" | --lat <n> --lon <n> | --batch <csv> | --capabilities");
  process.exit(1);
} else {
  const r = await runOne(new OpenAI({ apiKey: process.env.OPENAI_API_KEY }), args);
  console.log(JSON.stringify(r, null, 2));
}

// Photo check: what the photos on a new work order show that changes how the job is booked
// (Brandon, 27 Nov 2026).
//
// A work order's photos often show what its text leaves out: a 90cm cooktop, a double oven,
// the aircon units on site, a two-storey house under the evap, a ceiling too high for the
// small ladder. Each of those changes the booking (two people, an AC tech, a tall ladder),
// and today the office only finds out on site.
//
// One AI call looks at the job's real photos (already filtered by filterRealPhotos) next to
// the work order text and lists what it can SEE, each with a confidence and the evidence.
// Only clear findings (confidence >= WRITE_CONFIDENCE, evidence given, not already in the
// work order text) count. What happens with them is PHOTO_CHECK:
//
//   off     nothing runs
//   email   (default) Brandon gets an email when something is found; AroFlo is not touched
//   aroflo  as email, and the findings are also APPENDED to the job: one dated line per need
//           in the pinned note's SCHEDULING NOTE section (where the booking checker reads
//           needs), and an aircon unit line at the bottom of the task description on aircon
//           jobs that have no unit type from the office. Nothing already there is changed.
//
// Brandon asked for the AroFlo writes ("it should change the task description and the
// scheduling notes when needed") on the condition that it is accurate, so it starts on
// email until the accuracy run (backtest below) has been checked.
//
// It runs on the work order key (OPENAI_10_WORKORDER_EMAILS): Brandon, 7 Nov 2026, work order
// features share that key.

export const PHOTO_CHECK_MODEL = process.env.PHOTO_CHECK_MODEL || "gpt-5-mini";

// A finding is acted on only at or above this. Measured against the backtest before the
// AroFlo writes are turned on; raise it rather than loosen the prompt if anything is wrong.
export const WRITE_CONFIDENCE = 0.9;

// Photos sent per job. Work orders rarely carry more than a handful of real photos; past
// this the extra ones are almost always the same thing from another angle.
export const MAX_PHOTOS = 8;

export const KINDS = ["appliance_90cm", "double_oven", "aircon", "two_storey", "high_ceiling"];

export function photoCheckMode(env = process.env) {
  const v = String(env.PHOTO_CHECK || "").trim().toLowerCase();
  return v === "off" || v === "aroflo" ? v : "email";
}

export const INSTRUCTIONS = `You look at the photos attached to a property maintenance work order for an electrical and air conditioning business in Perth, Western Australia, before the job is booked.

Report ONLY things you can clearly SEE in the photos that change how the job is booked. Do not report anything you only read in the work order text. Do not guess. If a photo is blurry, too close, or does not show the thing, report nothing for it. Reporting nothing is the normal answer.

What to look for:
1. appliance_90cm: a 90cm wide cooktop, oven, upright cooker or rangehood. Only when the width is evident: a rating plate, model number or label that says 90cm/900mm (model numbers with 90 or 900 in the width position, e.g. OBO9011, IO9060, WVE9515), a measuring tape, or a cooktop with 5 burners/zones laid out across clearly wider than a standard 60cm one. A normal 4-burner 60cm cooktop is NOT 90cm. If unsure, do not report it.
2. double_oven: two separate oven cavities with two doors (stacked wall ovens, or an upright cooker with two oven doors). A single oven with a grill compartment or warming drawer is NOT a double oven.
3. aircon: each air conditioner you can see, with its type: "split" (wall mounted indoor head or its outdoor condenser), "ducted" (ceiling return air grille with filter, ducted outdoor unit, ducted controller), "evap" (evaporative cooler on the roof, ceiling vents of an evap, evap wall controller like Breezair/Coolair/Braemar), "cassette", "window". Give the count you can actually see. A ceiling exhaust fan, a range hood or a ceiling fan is NOT an aircon.
4. two_storey: the building itself clearly has an upper floor (a second row of windows above, an internal staircase going up, a balcony on an upper level). Only for a house or townhouse; an apartment block with many floors is not this.
5. high_ceiling: the work is plainly on a ceiling or wall higher than about 3 metres (a raked or cathedral ceiling, a void, a stairwell, a double height entry) so a tall ladder is needed. A normal 2.4 to 2.7 metre ceiling is NOT this.

For each finding give:
- kind: one of appliance_90cm, double_oven, aircon, two_storey, high_ceiling
- detail: a few words a scheduler understands, e.g. "90cm Westinghouse cooktop", "stacked double wall oven", "2x split system", "1x ducted reverse cycle", "1x evap", "raked ceiling about 4.5m at the light"
- aircon_type and aircon_count: only for aircon (type one of split, ducted, evap, cassette, window)
- photos: the photo numbers (1-based) that show it
- confidence: 0.0 to 1.0, how sure you are from the photos alone
- evidence: what in the photo shows it, e.g. "rating plate reads OBO9011AM", "five burners across, wider than the 60cm wall oven below"

Return ONLY JSON:
{"findings":[{"kind":"","detail":"","aircon_type":null,"aircon_count":null,"photos":[1],"confidence":0.0,"evidence":""}]}`;

// The model's answer, held to the schema: unknown kinds, empty evidence and out-of-range
// numbers are dropped rather than trusted. Null when the answer is not JSON at all.
export function parseFindings(text, photoCount = MAX_PHOTOS) {
  let body;
  try {
    const raw = String(text || "");
    const start = raw.indexOf("{");
    const end = raw.lastIndexOf("}");
    if (start < 0 || end < start) return null;
    body = JSON.parse(raw.slice(start, end + 1));
  } catch {
    return null;
  }
  const list = Array.isArray(body?.findings) ? body.findings : [];
  const findings = [];
  for (const f of list) {
    if (!f || !KINDS.includes(f.kind)) continue;
    const confidence = Number(f.confidence);
    if (!Number.isFinite(confidence) || confidence < 0 || confidence > 1) continue;
    const evidence = String(f.evidence || "").trim();
    const detail = String(f.detail || "").trim();
    if (!evidence || !detail) continue;
    const photos = (Array.isArray(f.photos) ? f.photos : [])
      .map(Number).filter(n => Number.isInteger(n) && n >= 1 && n <= photoCount);
    const out = { kind: f.kind, detail: detail.slice(0, 80), photos, confidence, evidence: evidence.slice(0, 200) };
    if (f.kind === "aircon") {
      const type = String(f.aircon_type || "").toLowerCase();
      const count = Number(f.aircon_count);
      if (!["split", "ducted", "evap", "cassette", "window"].includes(type)) continue;
      out.airconType = type;
      out.airconCount = Number.isInteger(count) && count >= 1 && count <= 10 ? count : null;
    }
    findings.push(out);
  }
  return { findings };
}

// What the work order already says, per kind, so a finding the office can already read is
// not written again. Same wording the booking checker keys on (bara-dashboard
// src/lib/schedulingChecks/needs.ts TWO_PERSON_RE / ROOF_TWO_PERSON_RE).
const ALREADY_SAID = {
  appliance_90cm: /\b90\s?cm\b|\b900\s?mm\b/i,
  double_oven: /\bdouble\s+ovens?\b|\btwin\s+ovens?\b/i,
  two_storey: /\b(?:two|double|2)[\s-]?stor(?:e?y|eys|ies)\b|\bupstairs\b/i,
  high_ceiling: /\b(?:tall|high|extension|\d+\s*(?:ft|foot|feet|m))\s+ladder\b|\b(?:raked|cathedral|high)\s+ceilings?\b|\bscaffold/i,
};

// Roof work, where a two-storey house matters (booking checker check.ts roofWork: an evap, or
// work said to be on the roof; a two-storey split service is not roof work).
const ROOF_WORK_RE = /\bevap(?:orative)?\b|\bon\s+the\s+roof\b|\broof\s?top\b|\bEV(?:EC)?1\b|\bsolar\b|\bantenna\b|\bwhirlybird\b/i;

const AIRCON_PACKAGES = ["AC1", "AC2", "ACEC1", "EV1", "EVEC1"];

const TYPE_LABEL = { split: "Split", ducted: "Ducted", evap: "Evap", cassette: "Cassette", window: "Window" };

/**
 * What the job gets from the findings. Pure, for the tests.
 *   text        the work order text the office already reads (task description + email/PDF)
 *   pkg         the job's package / task type, for "is this an aircon job"
 *   airconKnown the office already tagged a unit type, or the site record has units
 * Returns { schedulingLines, descriptionLine, used, skipped } where skipped carries the reason.
 */
export function photoActions(findings, { text = "", pkg = null, taskType = null, airconKnown = false } = {}) {
  const used = [];
  const skipped = [];
  const schedulingLines = [];
  let descriptionLine = null;
  const roofWork = ROOF_WORK_RE.test(text) || ["EV1", "EVEC1"].includes(pkg) || ["EV1", "EVEC1"].includes(taskType);
  const airconJob = AIRCON_PACKAGES.includes(pkg) || AIRCON_PACKAGES.includes(taskType) || /air\s?con/i.test(taskType || "");

  for (const f of findings || []) {
    if (f.confidence < WRITE_CONFIDENCE) { skipped.push({ ...f, reason: "not sure enough" }); continue; }
    if (ALREADY_SAID[f.kind]?.test(text)) { skipped.push({ ...f, reason: "work order already says so" }); continue; }
    if (f.kind === "appliance_90cm") {
      schedulingLines.push(`Photos: ${f.detail} - 90cm, 2 people`);
    } else if (f.kind === "double_oven") {
      schedulingLines.push(`Photos: ${f.detail} - double oven, 2 people`);
    } else if (f.kind === "two_storey") {
      if (!roofWork) { skipped.push({ ...f, reason: "not roof work" }); continue; }
      schedulingLines.push(`Photos: two storey - 2 people for the roof`);
    } else if (f.kind === "high_ceiling") {
      schedulingLines.push(`Photos: ${f.detail} - tall ladder`);
    } else if (f.kind === "aircon") {
      if (!airconJob) { skipped.push({ ...f, reason: "not an aircon job" }); continue; }
      if (airconKnown) { skipped.push({ ...f, reason: "unit type already known" }); continue; }
      if (descriptionLine) { skipped.push({ ...f, reason: "one aircon line only" }); continue; }
      descriptionLine = `Photos: ${f.airconCount ? `${f.airconCount}× ` : ""}${TYPE_LABEL[f.airconType]}`;
    }
    used.push(f);
  }
  // One line per need, never the same line twice.
  return { schedulingLines: [...new Set(schedulingLines)], descriptionLine, used, skipped };
}

function photoInput(img, n) {
  return [
    { type: "input_text", text: `Photo ${n}${img.name ? ` (${img.name})` : ""}:` },
    { type: "input_image", image_url: `data:${img.contentType || "image/jpeg"};base64,${Buffer.from(img.data).toString("base64")}`, detail: "high" },
  ];
}

/**
 * Runs the check on a job's real photos. Null when there is nothing to look at or the call
 * failed: a failure never holds up or changes the job.
 */
export async function checkPhotos({ openai, images, workOrderText = "", model = PHOTO_CHECK_MODEL }) {
  const photos = (images || []).slice(0, MAX_PHOTOS);
  if (!openai || photos.length === 0) return null;
  try {
    const response = await openai.responses.create({
      model,
      instructions: INSTRUCTIONS,
      text: { format: { type: "json_object" } },
      input: [{
        role: "user",
        content: [
          { type: "input_text", text: `Work order text, for context only (do not report what it says, only what the photos show):\n${String(workOrderText).slice(0, 4000)}` },
          ...photos.flatMap((img, i) => photoInput(img, i + 1)),
        ],
      }],
    });
    const parsed = parseFindings(response.output_text, photos.length);
    if (!parsed) { console.warn("[photos] Photo check answer was not JSON"); return null; }
    return { ...parsed, photoCount: photos.length, model };
  } catch (err) {
    console.warn("[photos] Photo check failed:", err.message);
    return null;
  }
}

const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * The email to Brandon for one job. Sent only when something was found (used or not), so a
 * clean job sends nothing. `written` says whether the lines went on the job (PHOTO_CHECK=aroflo
 * and the writes worked) or would have.
 */
export function photoEmail({ jobNumber, jobUrl, address, check, actions, written, writeErrors = [] }) {
  const found = check?.findings || [];
  const subject = `Photo check: JN ${jobNumber} - ${actions.used.length ? actions.used.map(f => f.detail).join(", ") : "nothing clear enough"}`;
  const p = (s) => `<p style="font-family:sans-serif;font-size:14px;margin:6px 0">${s}</p>`;
  const job = jobUrl ? `<a href="${esc(jobUrl)}">JN ${esc(jobNumber)}</a>` : `JN ${esc(jobNumber)}`;
  const lines = [...actions.schedulingLines.map(l => `Scheduling note: ${l}`), ...(actions.descriptionLine ? [`Description: ${actions.descriptionLine}`] : [])];
  const row = (f, extra = "") => `<li style="font-family:sans-serif;font-size:14px;margin:4px 0"><strong>${esc(f.detail)}</strong> (${esc(f.kind)}, ${Math.round(f.confidence * 100)}%, photo ${esc(f.photos.join(", ") || "?")}): ${esc(f.evidence)}${extra}</li>`;
  const html = [
    p(`${job} - ${esc(address || "")}. ${check?.photoCount ?? 0} photo(s) checked.`),
    lines.length
      ? p(`<strong>${written ? "Added to the job" : "Would add to the job (AroFlo writes are off)"}:</strong>`) + `<ul>${lines.map(l => `<li style="font-family:sans-serif;font-size:14px">${esc(l)}</li>`).join("")}</ul>`
      : p("Nothing to add to the job."),
    writeErrors.length ? p(`<strong>Not written:</strong> ${esc(writeErrors.join("; "))}`) : "",
    actions.used.length ? p("<strong>Used:</strong>") + `<ul>${actions.used.map(f => row(f)).join("")}</ul>` : "",
    actions.skipped.length ? p("<strong>Seen but not used:</strong>") + `<ul>${actions.skipped.map(f => row(f, ` <em>(${esc(f.reason)})</em>`)).join("")}</ul>` : "",
  ].join("");
  return { subject, html, worthSending: found.length > 0 };
}

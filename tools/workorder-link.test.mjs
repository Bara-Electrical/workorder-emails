#!/usr/bin/env node
// Regression tests for findWorkOrderLink() / findLinkedPhotoLinks() — run with: npm test
//
// The link the app follows is where the work order's details come from, and the page
// snapshot opens it in a real browser. Same harness as the other tools/*.test.mjs files.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src  = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");
function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  return src.slice(i, src.indexOf(endMarker, i) + endMarker.length);
}
const module_ = [
  grab("const WORKORDER_DOMAINS", ";\n"),
  grab("function isActionLink(text, dest)", "\n}\n"),
  grab("function linkDestination(href)", "\n}\n"),
  grab("function findWorkOrderLink(rawHtml)", "\n}\n"),
  grab("function findLinkedPhotoLinks(rawHtml)", "\n}\n"),
  "export { findWorkOrderLink, findLinkedPhotoLinks };",
].join("\n");
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "workorder-link-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { findWorkOrderLink, findLinkedPhotoLinks } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

// ---- 6 Oct: Ray White Canning River, 111 Dulwich Street (job 108887). Verbatim anchor. ----
const AILO_PDF = 'https://aus01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fshared.outlook.inky.com%2Flink%3Fdomain%3Dfile.ailousercontent.com%26t%3Dh.eJxVjkuOwyAQRK8SsR4-zc84q1ylgUaxQuIIY0XWaO6eMJvRbOuV6tU321tl5xO79v7czlKWpZLApa77Ri2tj06PLtJ6_wWbVGWetfeWezLAraKJB3KFQykBTZ6nOWRpnUuQguHZYeB2UpljDJGTBx9sihEgsa8Tuw1xxIacKqXeloRVvtZ2W1umtl0G-iPjhcBd6hDRknYBAhX12dI4QUHvNEWTdZQwzaCdtqCEU0ZbY4eMhuzZaKt0iHuq2PpxaXi8rkunsT1KeZT-hT9v_jlZcA.MEYCIQDixpCSzGJwkXttnoLSya3PtG3_3_GSlQoRqktyPQrPzQIhAPERvyInUMrOvnpTuAUhbLHbv64J3bzZyEmcbiSETH2W&amp;data=05%7C02%7Cworkorders%40baraelectrical.com.au&amp;reserved=0';
const AILO_PHOTO = AILO_PDF.replace("h.eJxVjk", "h.eJxPHOTO");
const DULWICH = `<p><b>Files</b></p><p><a href="${AILO_PHOTO}">IMG_1234.jpg</a></p><p><a href="${AILO_PDF}">WO-2230-111-Dulwich-Street-Beckenham.pdf</a></p>`;
{
  const link = findWorkOrderLink(DULWICH);
  check("Ailo: the work order PDF is followed", link === AILO_PDF.replace(/&amp;/g, "&"), String(link).slice(0, 120));
  check("Ailo: the PDF is picked over a photo listed before it", link && !link.includes("PHOTO"));
  const photos = findLinkedPhotoLinks(DULWICH);
  check("Ailo: the photo is still picked up as a photo", photos.length === 1 && photos[0].name === "IMG_1234.jpg", JSON.stringify(photos.map(p => p.name)));
}

// ---- Bricks + Agent action buttons must never be followed. Verbatim from Peak Central. ----
const BNA = "https://aus01.safelinks.protection.outlook.com/?url=https%3A%2F%2Ftrade.bricksandagent.com%2Fexternal%2Foverview%2Fb386e634-5afe-4952-a6d3-09da98096cec%2Fb223e1b3-c75f-4cdd-22f2-08df1f805407";
const PEAK = [
  `<a href="${BNA}%3FisWorkOrderReject%3D1&amp;data=05">Reject Work Order</a>`,
  `<a href="${BNA}%3FisSchedule%3D1&amp;data=05">Schedule Job </a>`,
  `<a href="${BNA}%3FisStart%3D1&amp;data=05">Start </a>`,
  `<a href="${BNA}%3FunableToContactTenant%3D1&amp;data=05">Click here</a>`,
].join(" ");
check("Bricks + Agent: no action button is followed", findWorkOrderLink(PEAK) === null, String(findWorkOrderLink(PEAK)));
check("…even with the action only in the URL", findWorkOrderLink(`<a href="${BNA}%3FisWorkOrderReject%3D1">Open</a>`) === null);
check("…or only a camel-cased flag in the URL, behind innocent text",
  findWorkOrderLink(`<a href="${BNA}%3FisSchedule%3D1">View work order</a>`) === null);
check("…including unableToContactTenant",
  findWorkOrderLink(`<a href="${BNA}%3FunableToContactTenant%3D1">View work order</a>`) === null);
check("…or only in the text", findWorkOrderLink('<a href="https://example.com/x">Decline this work order</a>') === null);

// Professionals' "Review Work Order" opens the job and is fine to follow.
const REVIEW = `<a href="${BNA}&amp;data=05">Review Work Order </a>`;
check("Bricks + Agent: 'Review Work Order' is still followed", findWorkOrderLink(PEAK + REVIEW) === `${BNA}&data=05`);

// ---- Unchanged behaviour ----
check("plain 'work order' link text still wins", findWorkOrderLink('<a href="https://agency.example/wo/1">View work order</a>') === "https://agency.example/wo/1");
check("a known portal by domain still works", findWorkOrderLink('<a href="https://my.propertyme.com.au/doc/1">Open</a>') === "https://my.propertyme.com.au/doc/1");
check("an ordinary ?issue= parameter is not an action", findWorkOrderLink('<a href="https://tapihq.com/x?issue=42">View</a>') === "https://tapihq.com/x?issue=42");
check("http links are ignored", findWorkOrderLink('<a href="http://agency.example/wo">Work order</a>') === null);
check("no links, no result", findWorkOrderLink("<p>nothing</p>") === null);

const failures = results.filter(r => !r.ok);
console.log(`workorder link: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

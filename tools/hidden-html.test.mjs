#!/usr/bin/env node
// Regression tests for stripHiddenElements() / cleanHtml() — run with: npm test
//
// The AI reads the email as cleanHtml leaves it, so text hidden with display:none must go
// before the tags are stripped: nobody in Outlook can see it, and Bricks & Agent hides a
// placeholder PM in exactly that way. Same harness as the other tools/*.test.mjs files.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src  = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");

function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf(endMarker, i);
  if (j === -1) throw new Error(`no end marker "${endMarker}" after: ${startMarker}`);
  return src.slice(i, j + endMarker.length);
}

const module_ = [
  grab("const VOID_TAGS", ";\n"),
  grab("function stripHiddenElements(html)", "\n}\n"),
  grab("function cleanHtml(html)", "\n}\n"),
  "export { stripHiddenElements, cleanHtml };",
].join("\n");

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "hidden-html-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { stripHiddenElements, cleanHtml } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

// ---- 2 Oct: Peak Central, 16 Kumarina Drive (job 108733). Verbatim from the email. ----
const PEAK_CENTRAL =
  '<tr><td style="padding:10px 35px 10px 35px"><table class="content-row"><tbody><tr><td style="text-align:center;font-size:12px;line-height:1.58;color:#000"><p style="font-weight:bold;display:none">Property Manager Details</p></td></tr></tbody></table></td></tr>' +
  '<tr><td style="padding:0px 35px 0px 35px"><table><tbody><tr><td id="item" style="display:none"><table class="content-row"><tbody><tr><td style="font-weight:bold;width:30%"><p>Name</p></td><td style="width:70%"><p>: Jodie Mordacz</p></td></tr><tr><td style="font-weight:bold"><p>Mobile</p></td><td><p>: +61894149055</p></td></tr><tr><td style="font-weight:bold"><p>Email</p></td><td><p>: <a href="mailto:admin@peakcentral.com.au">admin@peakcentral.com.au</a></p></td></tr></tbody></table></td></tr></tbody></table></td></tr>' +
  '<tr><td style="padding:0px 35px 0px 35px;font-size:14px;text-align:center;line-height:1.58;color:#000"><table><tbody><tr><td style="text-align:center;font-size:12px;line-height:1.58;color:#000;padding-top:10px"><p style="font-weight:bold">Tenant Details</p></td></tr></tbody></table></td></tr>' +
  '<tr style="display:contents"><td id="item"><table class="content-row"><tbody><tr><td style="font-weight:bold;width:30%"><p>Contact</p></td><td style="width:70%"><p>: Julia Rickards</p></td></tr><tr><td style="font-weight:bold"><p>Email</p></td><td><p>: <a href="mailto:pm1@peakcentral.com.au">pm1@peakcentral.com.au</a></p></td></tr></tbody></table></td></tr>';
{
  const text = cleanHtml(PEAK_CENTRAL);
  check("the hidden placeholder PM is gone", !text.includes("Jodie Mordacz"), text);
  check("…with its admin@ email", !text.includes("admin@peakcentral.com.au"), text);
  check("…and the hidden heading", !text.includes("Property Manager Details"), text);
  check("the visible poster is kept", text.includes("Julia Rickards") && text.includes("pm1@peakcentral.com.au"), text);
  check("the visible Tenant Details heading is kept", text.includes("Tenant Details"), text);
  check("display:contents is not hidden", text.includes("Contact"), text);
}

// RMA's hidden-then-visible address paragraphs, verbatim: only the visible one survives.
{
  const html = '<p style="display:none;font-size:12px;text-align:center">284 Hamilton Road, Spearwood, WA 6163 <br />C/O Rental Management Australia (WA) </p><p style="display:block;font-size:12px;text-align:center">Melinda Burmas <br />C/O Rental Management Australia (WA) </p>';
  const text = cleanHtml(html);
  check("a hidden <p> with a <br> inside goes", !text.includes("284 Hamilton Road"), text);
  check("the visible <p> after it stays", text.includes("Melinda Burmas"), text);
}

// ---- The rule itself ----
{
  check("nested same-name tags are removed as one element",
    stripHiddenElements('<div style="display:none">a<div>b<div>c</div></div>d</div>after') === "after");
  check("text after the hidden element is untouched",
    stripHiddenElements('before<span style="display: none">x</span>after') === "beforeafter");
  check("the last display declaration wins (none then block stays)",
    stripHiddenElements('<p style="display:none;display:block">kept</p>').includes("kept"));
  check("the last display declaration wins (block then none goes)",
    !stripHiddenElements('<p style="display:block;display:none">gone</p>').includes("gone"));
  check("Inky's !important visible banner stays",
    stripHiddenElements('<div style="display:block !important;display:block">banner</div>').includes("banner"));
  check("single-quoted style is read", !stripHiddenElements("<p style='display:none'>gone</p>").includes("gone"));
  check("upper-case markup is read", !stripHiddenElements('<P STYLE="DISPLAY:NONE">gone</P>').includes("gone"));
  check("a hidden void element goes alone",
    stripHiddenElements('a<img style="display:none" src="x">b') === "ab");
  check("<p> does not pair with <pre>",
    stripHiddenElements('<p style="display:none">x<pre>y</pre></p>z') === "z");
  check("an unclosed hidden element hides to the end",
    stripHiddenElements('kept<div style="display:none">gone') === "kept");
  check("visibility:hidden is not display:none",
    stripHiddenElements('<p style="visibility:hidden">kept</p>').includes("kept"));
  check("no style, no change", stripHiddenElements("<p>plain</p>") === "<p>plain</p>");
  check("two hidden siblings both go",
    stripHiddenElements('<p style="display:none">a</p>mid<p style="display:none">b</p>') === "mid");
}

const failures = results.filter(r => !r.ok);
console.log(`hidden html: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    ${f.detail}` : ""}`).join("\n"));
  process.exit(1);
}

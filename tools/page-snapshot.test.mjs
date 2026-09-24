#!/usr/bin/env node
// Tests for page-snapshot.js — run with: npm test
//
// This one really launches a browser and really prints a PDF, because the things that break a
// snapshot are not logic: a missing executable, a missing font, a page that never goes quiet.
// None of that shows up in a mock.
//
// It skips cleanly when there is no browser to drive (playwright-core not installed, or no
// Chromium on PATH), so a checkout without the image's system packages still runs npm test.
// In the deployed image both are present and these run for real.

import fs from "node:fs";
import { execFileSync } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, "..");

// Chromium lives in the image at CHROMIUM_PATH; a dev machine may have a Playwright download.
function findBrowser() {
  if (process.env.CHROMIUM_PATH && fs.existsSync(process.env.CHROMIUM_PATH)) return process.env.CHROMIUM_PATH;
  for (const base of ["/opt/pw-browsers", process.env.PLAYWRIGHT_BROWSERS_PATH].filter(Boolean)) {
    if (!fs.existsSync(base)) continue;
    for (const dir of fs.readdirSync(base)) {
      const exe = path.join(base, dir, "chrome-linux", "chrome");
      if (fs.existsSync(exe)) return exe;
    }
  }
  for (const exe of ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"]) {
    if (fs.existsSync(exe)) return exe;
  }
  return null;
}

async function resolvable(spec) {
  try { await import(spec); return true; } catch { return false; }
}

const browser = findBrowser();
if (!browser || !(await resolvable("playwright-core"))) {
  console.log(`page snapshot: skipped (${!browser ? "no chromium found" : "playwright-core not installed"})`);
  process.exit(0);
}
process.env.CHROMIUM_PATH = browser;

const m = await import(pathToFileURL(path.join(root, "page-snapshot.js")).href);

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "page-snapshot-"));
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

// A stand-in for a work-order page, carrying the things that must survive into the PDF: the
// scope, the spending limit, the branch in the signature, and the page's own colour coding.
const pageFile = path.join(tmpDir, "page.html");
fs.writeFileSync(pageFile, `<!doctype html><html><head><style>
body{font-family:Helvetica,Arial,sans-serif;background:#f6f6f6;padding:40px}
.card{background:#fff;padding:24px;border-left:8px solid #ffae00}
h1{color:#ffae00} td{border:1px solid #ddd;padding:6px}
</style></head><body><div class="card"><h1>New work order</h1>
<table><tr><td>Address</td><td>116 Amherst Road, Canning Vale WA</td></tr>
<tr><td>Expenditure Limit</td><td>$330</td></tr></table>
<p>Regards, Austpro Properties Booragoon</p></div></body></html>`);
const pageUrl = pathToFileURL(pageFile).href;

const results = [];
const check = (name, cond, detail = "") => results.push({ name, ok: !!cond, detail });

{
  const snap = await m.snapshotWorkOrderPage(pageUrl);
  check("it returns a snapshot", snap !== null, String(snap));
  if (snap) {
    const magic = Buffer.from(snap.bytes.slice(0, 5)).toString("latin1");
    check("the bytes are a PDF", magic === "%PDF-", magic);
    // A blank A4 from the same Chromium is under 1KB. Anything near that rendered nothing —
    // the failure mode where fonts are missing and the page comes out empty.
    check("the PDF has real content, not a blank page", snap.bytes.length > 5000, `${snap.bytes.length} bytes`);
    check("the filename is dated", /^Work order page \d{4}-\d{2}-\d{2}\.pdf$/.test(snap.filename), snap.filename);
  }
}

// The label is what distinguishes one snapshot from another on the job card.
{
  const snap = await m.snapshotWorkOrderPage(pageUrl, { label: "Issue report" });
  check("the label is used", snap?.filename.startsWith("Issue report "), snap?.filename);
}

// Phase 2 snapshots up to POLL_BATCH_SIZE pages at once, through one shared browser.
{
  const many = await Promise.all(Array.from({ length: 10 }, () => m.snapshotWorkOrderPage(pageUrl)));
  check("ten at once all succeed", many.every(s => s?.bytes?.length > 5000), `${many.filter(Boolean).length}/10`);
}

// Every failure path must warn and return null. A snapshot is a nice-to-have on a job that
// already exists, so throwing here would cost a job to save a PDF.
{
  const warn = console.warn;
  console.warn = () => {};
  const dead    = await m.snapshotWorkOrderPage("https://127.0.0.1:9/nope");
  const nothing = await m.snapshotWorkOrderPage("");
  const nullish = await m.snapshotWorkOrderPage(null);
  const garbage = await m.snapshotWorkOrderPage("not-a-url");
  console.warn = warn;
  check("an unreachable page returns null", dead === null, String(dead));
  check("an empty url returns null", nothing === null, String(nothing));
  check("a null url returns null", nullish === null, String(nullish));
  check("a malformed url returns null", garbage === null, String(garbage));
}

// A browser that died must be relaunched, not handed out disconnected forever.
{
  await m.closeSnapshotBrowser();
  const after = await m.snapshotWorkOrderPage(pageUrl);
  check("it relaunches after the browser closes", after?.bytes?.length > 5000, String(after && after.bytes.length));
  await m.closeSnapshotBrowser();
}

// No Chromium may outlive closeSnapshotBrowser. This is the check that matters most for a
// service that runs for weeks: a browser launched per snapshot instead of shared passes every
// check above — it is correct, and fast enough — but it orphans a ~350MB Chromium each time,
// and the service climbs until it runs out of memory. After everything above, a leak shows
// up here as browsers still parented to this process once the one we know about is closed.
{
  const chromiumChildren = () => {
    try {
      return execFileSync("ps", ["-o", "pid=,args=", "--ppid", String(process.pid)], { encoding: "utf8" })
        .split("\n").filter(line => /chrom/i.test(line));
    } catch { return null; } // ps without --ppid (macOS): cannot tell, so do not guess
  };
  let left = chromiumChildren();
  // A closed browser takes a moment to reap; give it a few seconds before calling it a leak.
  for (let i = 0; left?.length && i < 30; i++) {
    await new Promise(r => setTimeout(r, 100));
    left = chromiumChildren();
  }
  if (left === null) console.log("page snapshot: leak check skipped (ps has no --ppid here)");
  else check("no browser outlives closeSnapshotBrowser", left.length === 0, `${left.length} still running`);
}

const failures = results.filter(r => !r.ok);
console.log(`page snapshot: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}${f.detail ? `\n    got: ${f.detail}` : ""}`).join("\n"));
}
// Explicit, because a leaked browser holds the event loop open: without this a leak turns
// into a run that never finishes instead of the failure recorded above.
process.exit(failures.length ? 1 : 0);

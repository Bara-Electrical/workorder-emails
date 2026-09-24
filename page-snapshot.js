// ================================================================
// WORK-ORDER PAGE SNAPSHOT
// ================================================================
// Agencies edit a work-order page after sending it — the scope changes, the spending limit
// changes — and until now there was no record of what it said when we accepted it. This
// prints the page to PDF at the moment the email is processed and the result is uploaded to
// the job, so the original is on the job card whatever the page says later.
//
// Chromium comes from the image (see Dockerfile), not from npm: playwright-core never
// downloads a browser, so the dependency is a few hundred kilobytes of client code and the
// browser is a system package. CHROMIUM_PATH points at it.
//
// This is best effort throughout. A snapshot is a nice-to-have on a job that already exists,
// so every failure here returns null with a warning and never interrupts job creation.

import { chromium } from "playwright-core";

const NAV_TIMEOUT_MS    = 30000;
const SETTLE_TIMEOUT_MS = 8000;
const PDF_TIMEOUT_MS    = 30000;

// One browser for the process, reused across polls. Launching Chromium costs about a second
// and ~350MB, and phase 2 snapshots up to POLL_BATCH_SIZE pages at once — a browser each
// would be ten launches and several gigabytes for no benefit, where ten pages in one browser
// is one launch. Pages are what leak, and each is closed in its own finally below.
let browser = null;
let launching = null;

async function getBrowser() {
  if (browser?.isConnected()) return browser;
  // A crashed or killed browser leaves a disconnected handle behind; drop it and relaunch
  // rather than handing out something every call will fail on.
  browser = null;
  // Concurrent callers in phase 2 must not each start their own launch.
  launching ??= chromium.launch({
    executablePath: process.env.CHROMIUM_PATH || undefined,
    // --no-sandbox: required in a container — Chromium's sandbox needs privileges the runtime
    // doesn't grant, and this process only ever visits work-order URLs the agencies sent us.
    // --disable-domain-reliability / --no-pings: a headless renderer on a server has no reason
    // to contact Google. Playwright already turns off background networking and component
    // updates, but Chromium still reached www.google.com and redirector.gvt1.com on launch
    // without these (counted at the egress proxy while testing).
    args: ["--no-sandbox", "--disable-dev-shm-usage", "--disable-domain-reliability", "--no-pings"],
  }).then(b => { browser = b; launching = null; return b; },
          err => { launching = null; throw err; });
  return launching;
}

// Print one work-order page to PDF. Returns { filename, bytes } or null.
export async function snapshotWorkOrderPage(url, { label = "Work order page" } = {}) {
  if (!url) return null;
  let page = null;
  try {
    const b = await getBrowser();
    page = await b.newPage({ viewport: { width: 1280, height: 1600 } });
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: NAV_TIMEOUT_MS });
    // Photos and webfonts land after domcontentloaded, and a page that never goes quiet
    // (polling, analytics) would hang on networkidle — so wait for quiet, but don't insist.
    await page.waitForLoadState("networkidle", { timeout: SETTLE_TIMEOUT_MS }).catch(() => {});
    const bytes = await page.pdf({
      format: "A4",
      printBackground: true,          // the page's own colour coding is part of the record
      margin: { top: "10mm", bottom: "10mm", left: "10mm", right: "10mm" },
      timeout: PDF_TIMEOUT_MS,
    });
    // Dated, because the point of the file is which version of the page this was.
    const stamp = new Date().toISOString().slice(0, 10);
    return { filename: `${label} ${stamp}.pdf`, bytes: new Uint8Array(bytes) };
  } catch (err) {
    console.warn("[snapshot] Could not capture the work-order page:", err.message);
    return null;
  } finally {
    if (page) await page.close().catch(() => {});
  }
}

// For a clean shutdown; the poll loop does not need it.
export async function closeSnapshotBrowser() {
  const b = browser;
  browser = null;
  if (b) await b.close().catch(() => {});
}

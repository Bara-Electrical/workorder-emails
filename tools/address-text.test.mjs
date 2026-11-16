import assert from "node:assert/strict";
import { normaliseAddress } from "../addressText.js";

const cases = [
  // The bug: St James suburb came out as "Street James" (16 Nov 2026).
  ["12 Smith St, St James WA 6102", "12 Smith Street, St James WA 6102"],
  ["12 SMITH ST, ST JAMES WA 6102", "12 Smith Street, St James WA 6102"],
  ["12 Smith St, Saint James, WA, 6102, Australia", "12 Smith Street, St James, WA, 6102"],
  ["12 Smith St St James WA 6102", "12 Smith Street St James WA 6102"],
  ["12 Smith Rd St James WA 6102", "12 Smith Road St James WA 6102"],
  ["4/20 Jarrah Rd, St James, WA 6102", "4/20 Jarrah Road, St James, WA 6102"],
  // "St" at the start of a street name stays St.
  ["100 St Georges Tce, Perth WA 6000", "100 St Georges Terrace, Perth WA 6000"],
  ["5 St Kilda Rd, Rockingham WA 6168", "5 St Kilda Road, Rockingham WA 6168"],
  // Ordinary street-type "St" still expands.
  ["72 Hamilton St, Bentley WA 6102", "72 Hamilton Street, Bentley WA 6102"],
  ["72 Hamilton St Bentley WA 6102", "72 Hamilton Street Bentley WA 6102"],
  ["9 James St, Bassendean WA 6054", "9 James Street, Bassendean WA 6054"],
  ["9 Smith St. Bentley WA", "9 Smith Street Bentley WA"],
  // Other street types unchanged from before.
  ["3 Oak Ave, Como WA", "3 Oak Avenue, Como WA"],
  ["1 Lake Dr, Joondalup WA 6027", "1 Lake Drive, Joondalup WA 6027"],
  ["8 Bay Cres, Mandurah WA", "8 Bay Crescent, Mandurah WA"],
];

let failed = 0;
for (const [input, want] of cases) {
  const got = normaliseAddress(input);
  try { assert.equal(got, want); }
  catch { failed++; console.error(`FAIL ${JSON.stringify(input)}\n  got  ${JSON.stringify(got)}\n  want ${JSON.stringify(want)}`); }
}
if (failed) { console.error(`${failed} of ${cases.length} address cases failed`); process.exit(1); }
console.log(`address-text: ${cases.length} cases passed`);

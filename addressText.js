// Tidies the address the AI pulls out of a work order: title case, street types spelt out
// ("Rd" → "Road"), and "St" read the right way round. "St" is either a street type
// ("12 Smith St") or the start of a place name ("St James", "100 St Georges Tce"). Expanding
// every "St" to "Street" made St James jobs "Street James" (16 Nov 2026), so a place-name "St"
// is kept as "St", the way Landgate and AroFlo spell the suburb. A written-out "Saint" becomes
// "St" too, so "Saint James" and "St James" are one suburb when sites are matched.

const STREET_TYPES = {
  st: "Street", rd: "Road", ave: "Avenue", dr: "Drive", ln: "Lane", ct: "Court",
  pl: "Place", cl: "Close", cres: "Crescent", blvd: "Boulevard", hwy: "Highway",
  fwy: "Freeway", tce: "Terrace", pde: "Parade", gr: "Grove", bvd: "Boulevard",
  cct: "Circuit", esp: "Esplanade",
};

// Spelt-out street types, so "Smith Road St James" knows the street has already ended.
const STREET_TYPE_WORDS = new Set([
  ...Object.values(STREET_TYPES).map(w => w.toLowerCase()),
  "way", "loop", "rise", "mews", "vista", "view", "gardens", "square", "walk", "retreat",
  "chase", "link", "promenade", "avenue", "row", "pass", "gate", "heights", "ramble", "bend",
]);

// Names that follow "St" in Australian suburbs and streets, for "12 Smith St St James" style
// addresses with no comma between street and suburb.
const SAINT_NAMES = new Set([
  "james", "georges", "george", "kilda", "albans", "ives", "marys", "mary", "leonards",
  "peters", "clair", "helens", "lucia", "morris", "agnes", "andrews", "arnaud", "johns",
  "kitts", "huberts", "lawrence", "pauls", "patricks", "michaels", "quentin", "albans",
  "helena", "anne", "annes", "brigids", "davids", "lukes", "marks", "matthews", "stephens",
  "thomas", "vincents", "aubyn", "heliers", "germain", "fillans", "kilians",
]);

const STATES = new Set(["wa", "nsw", "vic", "qld", "sa", "tas", "act", "nt"]);

function titleCase(address) {
  return address.replace(/\b\w+/g, w => {
    if (/^(st|saint)$/i.test(w)) return w.toLowerCase(); // decided below
    return w.length <= 2 ? w.toUpperCase() : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  });
}

// One comma-separated part of the address, e.g. "12 Smith St St James WA 6102".
function fixSegment(segment) {
  const tokens = segment.split(/(\s+)/); // keeps the spaces
  const words = []; // indexes of non-space tokens
  tokens.forEach((t, i) => { if (t && !/^\s+$/.test(t)) words.push(i); });

  words.forEach((ti, wi) => {
    const raw = tokens[ti];
    const bare = raw.replace(/\.$/, "").toLowerCase();
    const prev = wi > 0 ? tokens[words[wi - 1]] : null;
    const next = wi < words.length - 1 ? tokens[words[wi + 1]] : null;
    const nextBare = next ? next.replace(/[^a-z]/gi, "").toLowerCase() : "";

    if (bare === "saint") { tokens[ti] = next ? "St" : "Saint"; return; }
    if (bare !== "st") {
      const full = STREET_TYPES[bare];
      if (full) tokens[ti] = full;
      return;
    }

    const nextIsName = !!next && /^[a-z]/i.test(next) && !STREET_TYPES[nextBare] && !STATES.has(nextBare);
    const placeName = nextIsName && (
      prev === null                                   // starts the part: ", St James WA"
      || /^\d/.test(prev)                             // after a house number: "100 St Georges"
      || STREET_TYPE_WORDS.has(prev.toLowerCase())    // street already ended: "Smith Road St James"
      || SAINT_NAMES.has(nextBare)                    // "Smith St St James"
    );
    tokens[ti] = placeName ? "St" : "Street";
  });

  return tokens.join("");
}

export function normaliseAddress(address) {
  let a = String(address || "").replace(/,?\s*Australia\s*$/i, "").trim();
  a = titleCase(a);
  return a.split(",").map(fixSegment).join(",");
}

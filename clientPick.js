// Picking the client card for a work order whose agency name matched no card.
//
// The name matchers in index.js (exact, normalised, starts-with, the alias maps) stay first.
// When they all miss, this asks the AI to choose from the active client cards, and the office
// confirms or corrects the choice in the Outlook plugin before any job is made — so a wrong
// pick costs a click, never a job on the wrong card. A pick the office confirms is saved on
// the dashboard against the sender and the agency name, and the next email from them never
// gets here.
//
// Everything takes its OpenAI client as an argument so the tests can drive it offline.

// The model is told to answer with a card name exactly as listed. Anything it returns that
// is not on the list is treated as no answer rather than fuzzily matched back, because a
// near-miss name is precisely how a job ends up on the wrong card.
export const CLIENT_PICK_MODEL = "gpt-5-mini";

const INSTRUCTIONS = `You match a work order email to the AroFlo client card the job must be booked under, for an electrical company in Perth, Western Australia.

The client is the real estate agency, property manager or business that SENT the work order and will pay for it. It is never:
- the platform that delivered the email (Console, PropertyMe, Tapi, Bricks + Agent, Managed App, "emaildelivery@console.com.au", "hi@tapihq.com", "reply@bricksandagent.com", "no-reply@managedapp.com.au");
- the property owner, landlord, tenant, strata company or a conveyancer named on the work order;
- Bara Electrical itself (an email forwarded from a Bara Electrical address still belongs to the agency that wrote the original).

How to decide:
- A legal or trading name ("Grand Alliance Property Group Pty Ltd T/As Century 21 Grand Alliance", "c/o Canquan Pty Ltd") can point to the card of the agency that trades under it.
- The sender's own email domain (not a platform's) identifies the agency, e.g. @coronis.com.au is "Coronis Now WA".
- Many agencies have one card per branch ("Ray White ELG", "Ray White South Perth", "Austpro Properties - Booragoon", "Austpro Properties - South Perth"). Choose a branch card when the email names that branch, its suburb next to the agency name, or its office address. Never choose a branch because the PROPERTY is in that suburb.
- "Cards that already hold this property" is a hint only. Properties change agencies often, so never pick a card on the hint alone. When the email shows the agency and exactly one of that agency's branch cards already holds the property, that branch may be picked.
- If the email names a group with several branch cards and neither the email nor the hint settles the branch, answer null.
- An agency with no card of its own yet (a new client, or one that took over another agency's rentals) has no answer: return null rather than the closest-sounding card.

Answer with JSON: {"card": "<a card name copied exactly from the list>" or null, "reason": "<one short sentence>"}.`;

// The email as the model sees it: what was extracted, who sent it, and the raw text (the
// signature and the work order's letterhead carry the agency more often than the body).
// Capped so a long thread cannot crowd out the card list.
export function buildPickInput(evidence, cards) {
  const lines = [
    `Agency name read from the email: ${evidence.extractedName || "(none)"}`,
    `Sender: ${[evidence.fromName, evidence.from && `<${evidence.from}>`].filter(Boolean).join(" ") || "(unknown)"}`,
    `Subject: ${evidence.subject || ""}`,
    `Property: ${evidence.address || "(unknown)"}`,
  ];
  if (evidence.propertyCards?.length) {
    lines.push(`Cards that already hold this property (hint only): ${evidence.propertyCards.join("; ")}`);
  }
  const text = String(evidence.emailText || "").replace(/\s+/g, " ").trim().slice(0, 6000);
  if (text) lines.push("", "Email text:", text);
  lines.push("", `Client cards (${cards.length}), one per line:`, ...cards.map(c => c.clientname), "", "Answer with the JSON object described.");
  return lines.join("\n");
}

// The model's answer, held to the list: the exact name (case and spacing forgiven, nothing
// else), or null. When two cards share a name the pick is ambiguous and also null — the
// office's dropdown tells them apart, the model cannot.
export function parsePick(text, cards) {
  let body;
  try {
    body = JSON.parse(text);
  } catch {
    return null;
  }
  const card = typeof body?.card === "string" ? body.card : null;
  if (!card) return null;
  const key = squash(card);
  const hits = cards.filter(c => squash(c.clientname) === key);
  if (hits.length !== 1) return null;
  return { client: hits[0], reason: typeof body.reason === "string" ? body.reason.trim().slice(0, 300) : "" };
}

function squash(name) {
  return String(name || "").toLowerCase().replace(/\s+/g, " ").trim();
}

// → { client, reason } or null. Never throws: a failed call is "no suggestion", and the
// office still gets the dropdown.
export async function aiPickClient({ openai, evidence, cards, model = CLIENT_PICK_MODEL }) {
  if (!openai || !cards?.length) return null;
  try {
    const response = await openai.responses.create({
      model,
      text: { format: { type: "json_object" } },
      instructions: INSTRUCTIONS,
      input: buildPickInput(evidence, cards),
    });
    return parsePick(response.output_text, cards);
  } catch (err) {
    console.warn("[client-pick] AI pick failed:", err.message);
    return null;
  }
}

// Words that say what kind of business a card is rather than which one it is. Left out of
// the overlap score so "Smith Real Estate" is not a strong match for "Jones Real Estate".
const FILLER = new Set([
  "real", "estate", "property", "properties", "group", "realty", "management", "pty", "ltd",
  "the", "and", "of", "co", "wa", "perth", "agents", "residential", "t", "as", "c", "o",
]);

function words(name) {
  return String(name || "").toLowerCase().replace(/&/g, " and ").split(/[^a-z0-9@]+/).filter(w => w && !FILLER.has(w));
}

// The few cards whose names share the most distinctive words with the extracted name, for
// the plugin to offer beside the AI's pick. A plain word overlap: it only orders the
// shortcuts, and the office still searches the full list for anything else.
export function closestClients(name, cards, limit = 3) {
  const want = new Set(words(name));
  if (!want.size) return [];
  return cards
    .map(c => {
      const have = words(c.clientname);
      const shared = have.filter(w => want.has(w)).length;
      return { c, score: shared / Math.max(have.length, 1) + shared };
    })
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score || a.c.clientname.localeCompare(b.c.clientname))
    .slice(0, limit)
    .map(x => x.c);
}

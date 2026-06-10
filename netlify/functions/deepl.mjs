/* DeepL-bro: nettleseren kan ikke kalle DeepL direkte (CORS), så vi videresender her.
   Nøkkelen kommer fra brukeren (lagres aldri her) eller fra DEEPL_API_KEY på siten. */
import { json, requireAuth } from "../lib/util.mjs";

export const config = { path: "/api/deepl" };

const MAX_TEXTS = 60;
const MAX_CHARS = 60000;

export default async (req) => {
  try {
    requireAuth(req);                         // kun innloggede brukere
    if (req.method !== "POST") return json(405, { error: "Bruk POST." });
    let body;
    try { body = await req.json(); } catch (e) { return json(400, { error: "Ugyldig forespørsel." }); }

    const key = String(body.deeplKey || "").trim() || process.env.DEEPL_API_KEY || "";
    if (!key) return json(400, { error: "Ingen DeepL-nøkkel. Legg inn din egen under ⚙︎ (gratis hos deepl.com → 'DeepL API')." });

    let texts = Array.isArray(body.texts) ? body.texts.filter(t => typeof t === "string" && t.trim()) : [];
    if (!texts.length) return json(400, { error: "Ingen tekst å oversette." });
    if (texts.length > MAX_TEXTS) return json(400, { error: `For mange avsnitt på én gang (maks ${MAX_TEXTS}).` });
    const total = texts.reduce((a, t) => a + t.length, 0);
    if (total > MAX_CHARS) return json(400, { error: "For mye tekst på én gang – del kapittelet i to." });

    const host = key.endsWith(":fx") ? "api-free.deepl.com" : "api.deepl.com";
    const res = await fetch(`https://${host}/v2/translate`, {
      method: "POST",
      headers: { "content-type": "application/json", "authorization": "DeepL-Auth-Key " + key },
      body: JSON.stringify({ text: texts, source_lang: "EN", target_lang: "NB" }),
    }).catch(() => null);
    if (!res) return json(502, { error: "Fikk ikke kontakt med DeepL. Prøv igjen." });
    if (res.status === 403) return json(403, { error: "DeepL avviste nøkkelen – sjekk at den er riktig (⚙︎)." });
    if (res.status === 456) return json(456, { error: "DeepL-kvoten er brukt opp for denne perioden." });
    if (res.status === 429) return json(429, { error: "DeepL er opptatt – vent litt og prøv igjen." });
    if (!res.ok) return json(502, { error: "DeepL svarte med feil (" + res.status + "). Prøv igjen." });
    const d = await res.json();
    const out = (d.translations || []).map(t => t.text || "");
    return json(200, { ok: true, translations: out });
  } catch (err) {
    if (err instanceof Response) return err;
    return json(500, { error: "Noe gikk galt på tjeneren. Prøv igjen." });
  }
};

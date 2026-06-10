/* Pålogging: registrer, logg inn, logg ut, hvem er jeg. */
import { getStore } from "@netlify/blobs";
import { json, normEmail, validEmail, hashPassword, verifyPassword, sessionCookie, clearCookie, readSession, userKey } from "../lib/util.mjs";

export const config = { path: "/api/auth" };

export default async (req) => {
  try {
    if (req.method !== "POST") return json(405, { error: "Bruk POST." });
    if (!process.env.SESSION_SECRET) return json(500, { error: "Tjeneren mangler oppsett (SESSION_SECRET)." });
    let body;
    try { body = await req.json(); } catch (e) { return json(400, { error: "Ugyldig forespørsel." }); }
    const users = getStore({ name: "users", consistency: "strong" });
    const action = body.action;

    if (action === "register") {
      const email = normEmail(body.email);
      const name = String(body.name || "").trim().slice(0, 80);
      const password = String(body.password || "");
      const invite = String(body.invite || "").trim();
      if (!validEmail(email)) return json(400, { error: "Skriv en gyldig e-postadresse." });
      if (!name) return json(400, { error: "Skriv navnet ditt." });
      if (password.length < 8) return json(400, { error: "Passordet må ha minst 8 tegn." });
      const wanted = process.env.INVITE_CODE || "";
      if (wanted && invite !== wanted) return json(403, { error: "Feil invitasjonskode. Spør den som ga deg lenken." });
      const key = userKey(email);
      if (await users.get(key)) return json(409, { error: "Det finnes alt en bruker med denne e-posten. Logg inn i stedet." });
      const { salt, hash } = hashPassword(password);
      await users.setJSON(key, { email, name, salt, hash, createdAt: new Date().toISOString(), projects: [] });
      return json(200, { ok: true, name, email }, { "set-cookie": sessionCookie(email) });
    }

    if (action === "login") {
      const email = normEmail(body.email);
      const password = String(body.password || "");
      const key = userKey(email);
      const user = await users.get(key, { type: "json" });
      if (!user || !verifyPassword(password, user.salt, user.hash)) {
        return json(401, { error: "Feil e-post eller passord." });
      }
      return json(200, { ok: true, name: user.name, email: user.email }, { "set-cookie": sessionCookie(email) });
    }

    if (action === "logout") {
      return json(200, { ok: true }, { "set-cookie": clearCookie() });
    }

    if (action === "me") {
      const s = readSession(req);
      if (!s) return json(200, { ok: false });
      const user = await getStore({ name: "users", consistency: "strong" }).get(userKey(s.email), { type: "json" });
      if (!user) return json(200, { ok: false });
      return json(200, { ok: true, name: user.name, email: user.email });
    }

    return json(400, { error: "Ukjent handling." });
  } catch (err) {
    if (err instanceof Response) return err;
    return json(500, { error: "Noe gikk galt på tjeneren. Prøv igjen." });
  }
};

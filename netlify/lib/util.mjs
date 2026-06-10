/* Felles hjelpere for server-funksjonene: sesjoner, passord, svar. */
import crypto from "node:crypto";

const SECRET = process.env.SESSION_SECRET || "";
const COOKIE = "ob_session";
const NINETY_DAYS = 90 * 24 * 60 * 60;

export function json(status, body, extraHeaders) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", ...(extraHeaders || {}) },
  });
}

export function normEmail(e) {
  return String(e || "").trim().toLowerCase();
}

export function validEmail(e) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e) && e.length <= 200;
}

// ---------- Passord (scrypt med salt) ----------
export function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(password, salt, 64).toString("hex");
  return { salt, hash };
}
export function verifyPassword(password, salt, hash) {
  try {
    const got = crypto.scryptSync(password, salt, 64);
    const want = Buffer.from(hash, "hex");
    return got.length === want.length && crypto.timingSafeEqual(got, want);
  } catch (e) { return false; }
}

// ---------- Sesjons-token (HMAC-signert) ----------
function b64url(buf) { return Buffer.from(buf).toString("base64url"); }
function sign(payload) {
  return crypto.createHmac("sha256", SECRET).update(payload).digest("base64url");
}
export function makeSession(email) {
  const payload = b64url(JSON.stringify({ e: email, x: Date.now() + NINETY_DAYS * 1000 }));
  return payload + "." + sign(payload);
}
export function readSession(req) {
  if (!SECRET) return null;
  const cookie = req.headers.get("cookie") || "";
  const m = cookie.match(new RegExp("(?:^|;\\s*)" + COOKIE + "=([^;]+)"));
  if (!m) return null;
  const tok = m[1];
  const dot = tok.lastIndexOf(".");
  if (dot < 1) return null;
  const payload = tok.slice(0, dot), sig = tok.slice(dot + 1);
  const want = sign(payload);
  const a = Buffer.from(sig), b = Buffer.from(want);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try {
    const data = JSON.parse(Buffer.from(payload, "base64url").toString("utf-8"));
    if (!data.e || !data.x || Date.now() > data.x) return null;
    return { email: data.e };
  } catch (e) { return null; }
}
export function sessionCookie(email) {
  return `${COOKIE}=${makeSession(email)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${NINETY_DAYS}`;
}
export function clearCookie() {
  return `${COOKIE}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`;
}

// Bruker-id trygg som blob-nøkkel (e-post kan inneholde rare tegn)
export function userKey(email) {
  return crypto.createHash("sha256").update(email).digest("hex").slice(0, 32);
}

export function requireAuth(req) {
  const s = readSession(req);
  if (!s) throw json(401, { error: "Du må være innlogget." });
  return s;
}

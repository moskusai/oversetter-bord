/* Prosjekter i skyen: liste, lagre, åpne, gi nytt navn, slette. Kun egne prosjekter. */
import { getStore } from "@netlify/blobs";
import crypto from "node:crypto";
import { json, requireAuth, userKey } from "../lib/util.mjs";

export const config = { path: "/api/projects" };

const MAX_PROJECTS = 25;
const MAX_BYTES = 10 * 1024 * 1024;   // 10 MB per prosjekt

export default async (req) => {
  try {
    const s = requireAuth(req);
    const uid = userKey(s.email);
    const users = getStore("users");
    const projects = getStore("projects");
    const user = await users.get(uid, { type: "json" });
    if (!user) return json(401, { error: "Du må være innlogget." });
    user.projects = user.projects || [];

    if (req.method === "GET") {
      const list = [...user.projects].sort((a, b) => (b.updatedAt || "").localeCompare(a.updatedAt || ""));
      return json(200, { ok: true, projects: list, name: user.name });
    }

    if (req.method !== "POST") return json(405, { error: "Bruk POST." });
    let body;
    try { body = await req.json(); } catch (e) { return json(400, { error: "Ugyldig forespørsel." }); }
    const action = body.action;
    const own = (id) => user.projects.find(p => p.id === id);

    if (action === "save") {
      const name = String(body.name || "Uten navn").trim().slice(0, 100) || "Uten navn";
      const data = body.data;
      if (!data || typeof data !== "object") return json(400, { error: "Mangler prosjektdata." });
      const raw = JSON.stringify(data);
      if (raw.length > MAX_BYTES) return json(413, { error: "Prosjektet er for stort til skylagring (over 10 MB)." });
      let id = String(body.id || "");
      let meta = id ? own(id) : null;
      if (id && !meta) return json(404, { error: "Fant ikke prosjektet." });
      // Konfliktvern: hvis prosjektet er endret et annet sted siden klienten sist så det, si fra
      if (meta && body.expectedUpdatedAt && meta.updatedAt && body.expectedUpdatedAt !== meta.updatedAt && !body.force) {
        return json(409, { error: "Prosjektet er endret et annet sted siden sist.", conflict: true, updatedAt: meta.updatedAt });
      }
      if (!meta) {
        if (user.projects.length >= MAX_PROJECTS) return json(400, { error: `Du har alt ${MAX_PROJECTS} prosjekter – slett et gammelt først.` });
        id = crypto.randomUUID();
        meta = { id };
        user.projects.push(meta);
      }
      meta.name = name;
      meta.updatedAt = new Date().toISOString();
      meta.progress = typeof body.progress === "string" ? body.progress.slice(0, 40) : "";
      await projects.set(uid + "/" + id, raw);
      await users.setJSON(uid, user);
      return json(200, { ok: true, id, updatedAt: meta.updatedAt });
    }

    if (action === "load") {
      const meta = own(String(body.id || ""));
      if (!meta) return json(404, { error: "Fant ikke prosjektet." });
      const raw = await projects.get(uid + "/" + meta.id);
      if (!raw) return json(404, { error: "Prosjektdataene mangler." });
      return new Response(`{"ok":true,"name":${JSON.stringify(meta.name)},"updatedAt":${JSON.stringify(meta.updatedAt)},"data":${raw}}`,
        { status: 200, headers: { "content-type": "application/json; charset=utf-8" } });
    }

    if (action === "rename") {
      const meta = own(String(body.id || ""));
      if (!meta) return json(404, { error: "Fant ikke prosjektet." });
      meta.name = String(body.name || "").trim().slice(0, 100) || meta.name;
      await users.setJSON(uid, user);
      return json(200, { ok: true });
    }

    if (action === "delete") {
      const meta = own(String(body.id || ""));
      if (!meta) return json(404, { error: "Fant ikke prosjektet." });
      user.projects = user.projects.filter(p => p.id !== meta.id);
      await projects.delete(uid + "/" + meta.id);
      await users.setJSON(uid, user);
      return json(200, { ok: true });
    }

    return json(400, { error: "Ukjent handling." });
  } catch (err) {
    if (err instanceof Response) return err;
    return json(500, { error: "Noe gikk galt på tjeneren. Prøv igjen." });
  }
};

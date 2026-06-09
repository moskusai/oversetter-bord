/* Oversetter-bord – last opp original + oversettelse, jobb side om side, sjekk ord med AI.
   Ren JS, ingen avhengigheter. */
(function () {
  "use strict";

  const KEY = "oversetter-bord-v2";
  const OLD_KEY = "barnebibel-oversetter-v1";
  const PALETTE = ["#e8a06f", "#7bb0a0", "#c8a2d8", "#d9b24a", "#8fb4e3", "#d98f9e", "#9cc79a", "#caa37a"];
  const PROSE = ["title", "body", "quote", "note", "section"];
  const PROV_NAME = { anthropic: "Claude", openai: "ChatGPT", gemini: "Gemini" };

  // ---------- Lagring ----------
  function freshStore() {
    return {
      source: null,                 // { name, chapters:[{title,segments:[{id,type,en}]}] }
      translations: {}, links: {}, uncertain: {},
      settings: {
        provider: "anthropic",
        keys: { anthropic: "", openai: "", gemini: "" },
        models: { anthropic: "claude-opus-4-8", openai: "gpt-4o", gemini: "gemini-2.0-flash" },
      },
      lastChapter: 0, seenHelp: false,
    };
  }
  function load() {
    try {
      const v = JSON.parse(localStorage.getItem(KEY));
      if (v) return v;
    } catch (e) {}
    // Migrer nøkkel fra tidligere versjon hvis den finnes
    const s = freshStore();
    try {
      const old = JSON.parse(localStorage.getItem(OLD_KEY));
      if (old && old.settings) {
        if (old.settings.apiKey) s.settings.keys.anthropic = old.settings.apiKey;
        if (old.settings.model) s.settings.models.anthropic = old.settings.model;
      }
    } catch (e) {}
    return s;
  }
  let store = load();
  // sørg for at strukturen er komplett
  (function ensure() {
    const f = freshStore();
    store.settings = Object.assign({}, f.settings, store.settings || {});
    store.settings.keys = Object.assign({}, f.settings.keys, store.settings.keys || {});
    store.settings.models = Object.assign({}, f.settings.models, store.settings.models || {});
    store.translations = store.translations || {};
    store.links = store.links || {};
    store.uncertain = store.uncertain || {};
    if (store.lastChapter == null) store.lastChapter = 0;
  })();

  let saveTimer = null, saveWarned = false;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      try { localStorage.setItem(KEY, JSON.stringify(store)); saveWarned = false; }
      catch (e) {
        if (!saveWarned) { saveWarned = true; alert("Obs: nettleseren har lite lagringsplass igjen. Ta en sikkerhetskopi under «≡ Fil» for å være trygg."); }
      }
    }, 200);
  }

  // ---------- Tilstand ----------
  let ci = store.lastChapter || 0;
  let active = null;        // {segId, side, wi, word}
  let editingSeg = null;
  let lookupCtx = null;

  const hasSource = () => !!(store.source && store.source.chapters && store.source.chapters.length);
  const chapter = () => store.source.chapters[ci];
  function getNo(segId) { return (store.translations[ci] && store.translations[ci][segId]) || ""; }
  function setNo(segId, text) {
    if (!store.translations[ci]) store.translations[ci] = {};
    if (text) store.translations[ci][segId] = text; else delete store.translations[ci][segId];
    save();
  }
  function segLinks(segId) { return (store.links[ci] && store.links[ci][segId]) || []; }
  function setSegLinks(segId, arr) {
    if (!store.links[ci]) store.links[ci] = {};
    if (arr && arr.length) store.links[ci][segId] = arr; else delete store.links[ci][segId];
    save();
  }

  // ---------- Lese Word/tekst-filer ----------
  async function readDocxXml(buf) {
    const dv = new DataView(buf), u8 = new Uint8Array(buf);
    if (buf.byteLength < 22) throw new Error("Dette ser ikke ut som en gyldig Word-fil (.docx).");
    let eocd = -1;
    const min = Math.max(0, buf.byteLength - 22 - 65535);
    for (let i = buf.byteLength - 22; i >= min; i--) {
      // krev gyldig EOCD-signatur OG at kommentarlengden stemmer (unngå falsk treff i arkivkommentar)
      if (dv.getUint32(i, true) === 0x06054b50 && i + 22 + dv.getUint16(i + 20, true) === buf.byteLength) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("Dette ser ikke ut som en gyldig Word-fil (.docx).");
    const cdCount = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true), target = null;
    for (let i = 0; i < cdCount; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const compSize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const localOff = dv.getUint32(p + 42, true);
      const fname = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nameLen));
      if (fname === "word/document.xml") { target = { method, compSize, localOff }; break; }
      p += 46 + nameLen + extraLen + commentLen;
    }
    if (!target) throw new Error("Fant ikke teksten inne i Word-fila.");
    if (target.localOff + 30 > buf.byteLength || dv.getUint32(target.localOff, true) !== 0x04034b50)
      throw new Error("Word-fila ser ut til å være skadet.");
    const lhNameLen = dv.getUint16(target.localOff + 26, true);
    const lhExtraLen = dv.getUint16(target.localOff + 28, true);
    const dataStart = target.localOff + 30 + lhNameLen + lhExtraLen;
    if (dataStart + target.compSize > buf.byteLength) throw new Error("Word-fila ser ut til å være skadet.");
    const data = u8.subarray(dataStart, dataStart + target.compSize);
    if (target.method === 0) return new TextDecoder("utf-8").decode(data);
    if (target.method === 8) {
      if (typeof DecompressionStream === "undefined")
        throw new Error("Nettleseren din kan ikke pakke ut Word-filer. Prøv Chrome eller Safari, eller lim inn teksten manuelt.");
      const ds = new DecompressionStream("deflate-raw");
      const ab = await new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer();
      return new TextDecoder("utf-8").decode(ab);
    }
    throw new Error("Ukjent komprimering i Word-fila.");
  }
  function parseDocxXml(xml) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("Klarte ikke å lese innholdet i Word-fila.");
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const ps = doc.getElementsByTagNameNS(W, "p");
    const paras = [];
    for (let i = 0; i < ps.length; i++) {
      const p = ps[i];
      const styleEls = p.getElementsByTagNameNS(W, "pStyle");
      const style = styleEls.length ? styleEls[0].getAttributeNS(W, "val") : null;
      const ts = p.getElementsByTagNameNS(W, "t");
      let txt = ""; for (let j = 0; j < ts.length; j++) txt += ts[j].textContent;
      paras.push({ style: style, text: txt.trim() });
    }
    return paras;
  }
  async function docxToParas(buf) { return parseDocxXml(await readDocxXml(buf)); }

  function isHeading1(s) { return s && (/(heading|overskrift)\s*1(\b|[^0-9])/i.test(s) || /^(title|tittel)$/i.test(s)); }
  function isHeading(s) { return s && /(heading|overskrift|title|tittel)/i.test(s); }
  const SENT_END = /[.!?:"”’'…)»]$/;
  function endsSentence(s) { return SENT_END.test(s.trim()); }
  function startsContinuation(s) {           // ser ut som fortsettelse av forrige setning?
    const m = s.match(/\p{L}/u);
    if (!m) return true;                      // starter med tegn/tall
    return m[0].toLowerCase() === m[0] && m[0].toUpperCase() !== m[0]; // liten forbokstav
  }
  // Bygger kapitler. Slår sammen brødtekst-linjer som er delt midt i en setning
  // (vanlig i bøker med spaltebrudd), og hopper over rene sidetall.
  function parasToChapters(paras) {
    const chapters = []; let cur = null, bodyBuf = "";
    function flushBody() { if (cur && bodyBuf.trim()) cur.segments.push({ type: "body", en: bodyBuf.trim() }); bodyBuf = ""; }
    function open(t) { flushBody(); cur = { title: (t && t.trim()) || "(uten tittel)", segments: [] }; chapters.push(cur); }
    function pushSeg(seg) { flushBody(); if (!cur) open("Dokument"); cur.segments.push(seg); }
    for (const { style, text } of paras) {
      if (!text) continue;
      if (/^\d{1,4}$/.test(text) && !isHeading(style)) continue;      // sidetall
      if (isHeading1(style)) { open(text); cur.segments.push({ type: "title", en: text }); continue; }
      if (isHeading(style)) { pushSeg({ type: "section", en: text }); continue; }
      if (!cur) open("Dokument");
      if (!bodyBuf) bodyBuf = text;
      else if (!endsSentence(bodyBuf) && startsContinuation(text)) bodyBuf += " " + text;
      else { flushBody(); bodyBuf = text; }
    }
    flushBody();
    if (!chapters.length) open("Dokument");
    chapters.forEach(c => c.segments.forEach((s, i) => (s.id = i)));
    return chapters;
  }
  function textToChapters(text) {
    const blocks = text.replace(/\r/g, "").split(/\n\s*\n/).map(b => b.trim()).filter(Boolean);
    const paras = blocks.map(b => {
      const m = b.match(/^(#{1,6})\s+([\s\S]*)/);
      if (m) return { style: m[1].length <= 1 ? "Heading1" : "Heading2", text: m[2].trim() };
      return { style: null, text: b };
    });
    return parasToChapters(paras);
  }
  function assertSupported(file) {
    if (/\.doc$/i.test(file.name))
      throw new Error("Gammelt Word-format (.doc) støttes ikke. Åpne fila i Word og lagre som .docx, eller lim inn teksten direkte.");
  }
  async function fileToChapters(file) {
    assertSupported(file);
    if (/\.docx$/i.test(file.name)) return parasToChapters(await docxToParas(await file.arrayBuffer()));
    return textToChapters(await file.text());
  }
  async function fileToParagraphs(file) {
    assertSupported(file);
    if (/\.docx$/i.test(file.name)) return (await docxToParas(await file.arrayBuffer())).map(p => p.text).filter(Boolean);
    return (await file.text()).replace(/\r/g, "").split(/\n\s*\n/).map(b => b.trim()).filter(Boolean);
  }

  // ---------- Opplasting ----------
  async function handleEnglishFile(file) {
    if (!file) return;
    if (hasSource() && Object.keys(store.translations).length) {
      if (!confirm("Dette erstatter originalen og fjerner den pågående oversettelsen. Ta gjerne sikkerhetskopi først (≡ Fil). Vil du fortsette?")) return;
    }
    try {
      const chapters = await fileToChapters(file);
      const nSeg = chapters.reduce((a, c) => a + c.segments.length, 0);
      if (!nSeg) throw new Error("Fant ingen tekst i fila.");
      store.source = { name: file.name, chapters };
      store.translations = {}; store.links = {}; store.uncertain = {};
      ci = 0; active = null; editingSeg = null;
      save(); closeOverlays(); renderAll(); updateUncCount();
    } catch (err) {
      alert("Klarte ikke å lese fila: " + err.message);
    }
  }
  async function handleNorwegianFile(file) {
    if (!file) return;
    if (!hasSource()) { alert("Last opp den engelske originalen først, så vet appen hvor den norske teksten skal ligge."); return; }
    try {
      const paras = await fileToParagraphs(file);
      if (!paras.length) throw new Error("Fant ingen tekst i fila.");
      const targets = [];
      store.source.chapters.forEach((c, k) => c.segments.forEach(s => { if (PROSE.includes(s.type)) targets.push({ k, seg: s }); }));
      const had = Object.keys(store.translations).length > 0;
      const msg = `Dette legger ${paras.length} avsnitt fra den norske fila på rad over de ${targets.length} feltene i originalen.` +
        (paras.length !== targets.length ? "\n\nMERK: antallet er ulikt, så noen avsnitt kan havne forskjøvet – du må kanskje rette enkelte felt etterpå." : "") +
        (had ? "\n\nDen ERSTATTER den norske teksten som alt ligger der (ta gjerne sikkerhetskopi først)." : "") +
        "\n\nFortsette?";
      if (!confirm(msg)) return;
      // Ren erstatning – ingen gammel tekst blir liggende igjen i halen
      store.translations = {}; store.links = {}; store.uncertain = {};
      const n = Math.min(paras.length, targets.length);
      for (let i = 0; i < n; i++) {
        const { k, seg } = targets[i];
        if (!store.translations[k]) store.translations[k] = {};
        store.translations[k][seg.id] = paras[i];
      }
      save(); closeOverlays(); renderAll(); updateUncCount();
      alert(`Ferdig: la inn ${n} avsnitt.`);
    } catch (err) {
      alert("Klarte ikke å lese fila: " + err.message);
    }
  }

  // ---------- Tokenisering ----------
  function tokenize(text) {
    const re = /[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu;
    const out = []; let last = 0, wi = 0, m;
    while ((m = re.exec(text))) {
      if (m.index > last) out.push({ w: false, t: text.slice(last, m.index) });
      out.push({ w: true, t: m[0], wi: wi++ });
      last = m.index + m[0].length;
    }
    if (last < text.length) out.push({ w: false, t: text.slice(last) });
    return out;
  }
  function renderTokens(container, text, side, segId) {
    const toks = tokenize(text);
    const links = segLinks(segId);
    const colorByWi = {};
    links.forEach((pair, li) => { const wi = side === "en" ? pair[0] : pair[1]; if (colorByWi[wi] == null) colorByWi[wi] = li; });
    const unc = store.uncertain[ci] || [];
    for (const tk of toks) {
      if (!tk.w) { container.appendChild(document.createTextNode(tk.t)); continue; }
      const s = document.createElement("span");
      s.className = "w"; s.dataset.wi = tk.wi; s.textContent = tk.t;
      if (colorByWi[tk.wi] != null) { s.classList.add("linked"); s.style.setProperty("--lc", PALETTE[colorByWi[tk.wi] % PALETTE.length]); }
      if (unc.some(u => u.segId === segId && u.side === side && u.wi === tk.wi)) s.classList.add("w-uncertain");
      container.appendChild(s);
    }
  }

  // ---------- Render ----------
  const grid = document.getElementById("grid");
  const sel = document.getElementById("chapterSelect");

  function renderAll() {
    const show = hasSource();
    document.getElementById("emptyState").style.display = show ? "none" : "flex";
    document.getElementById("mainArea").style.display = show ? "" : "none";
    document.getElementById("chapnav").style.visibility = show ? "visible" : "hidden";
    document.getElementById("progress").style.visibility = show ? "visible" : "hidden";
    if (!show) { document.getElementById("appTitle").innerHTML = 'Oversetter-bord <small>engelsk → norsk</small>'; return; }
    if (ci >= store.source.chapters.length) ci = 0;
    // kapittelvelger
    sel.innerHTML = "";
    store.source.chapters.forEach((c, i) => {
      const o = document.createElement("option");
      o.value = i; o.textContent = (i + 1) + ". " + c.title;
      sel.appendChild(o);
    });
    const docName = store.source.name.replace(/\.(docx|txt|md)$/i, "");
    document.getElementById("appTitle").innerHTML = escapeHtml(docName) + ' <small>engelsk → norsk</small>';
    renderChapter();
  }

  function makeCard(side, seg) {
    const card = document.createElement("div");
    card.className = `card ${side} t-${seg.type}`;
    card.dataset.segId = seg.id; card.dataset.side = side;
    const txt = document.createElement("div"); txt.className = "txt";
    if (side === "en") { renderTokens(txt, seg.en, "en", seg.id); card.appendChild(txt); return card; }

    const noText = getNo(seg.id);
    if (editingSeg === seg.id) {
      const ta = document.createElement("textarea");
      ta.className = "editbox"; ta.value = noText; ta.placeholder = "Skriv eller lim inn norsk her …";
      const row = document.createElement("div"); row.className = "editrow";
      const ok = document.createElement("button"); ok.className = "btn primary"; ok.textContent = "Lagre";
      const cancel = document.createElement("button"); cancel.className = "btn"; cancel.textContent = "Angre endring";
      let done = false;
      const finish = (saveIt) => { if (done) return; done = true; if (saveIt) commitEdit(seg.id, ta.value); else { editingSeg = null; renderChapter(); } };
      ok.onclick = () => finish(true);
      cancel.onclick = () => finish(false);
      ta.addEventListener("keydown", (e) => {
        if (e.key === "Escape") finish(false);
        if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) finish(true);
      });
      // Klikker du bort uten å trykke noe, lagres teksten automatisk (ingen tap)
      ta.addEventListener("blur", () => finish(true));
      row.appendChild(ok); row.appendChild(cancel);
      card.appendChild(ta); card.appendChild(row);
      setTimeout(() => { ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length); }, 0);
      return card;
    }
    if (noText) { renderTokens(txt, noText, "no", seg.id); card.appendChild(txt); }
    else {
      const ph = document.createElement("span");
      ph.className = "no-empty"; ph.textContent = "✎ Lim inn / skriv norsk";
      ph.onclick = () => { editingSeg = seg.id; renderChapter(); };
      card.appendChild(ph);
    }
    const edit = document.createElement("button");
    edit.className = "editbtn"; edit.textContent = "✎ rediger";
    edit.onclick = (e) => { e.stopPropagation(); editingSeg = seg.id; renderChapter(); };
    card.appendChild(edit);
    return card;
  }
  function commitEdit(segId, value) {
    const old = getNo(segId), v = value.trim();
    if (v !== old) {
      setNo(segId, v);
      setSegLinks(segId, []);
      if (store.uncertain[ci]) { store.uncertain[ci] = store.uncertain[ci].filter(u => !(u.segId === segId && u.side === "no")); save(); }
    }
    editingSeg = null; active = null; renderChapter();
  }
  function renderChapter() {
    const c = chapter();
    document.getElementById("enTitle").textContent = c.title;
    const titleSeg = c.segments.find(s => s.type === "title");
    document.getElementById("noTitle").textContent = (titleSeg && getNo(titleSeg.id)) || "—";
    sel.value = ci;
    grid.innerHTML = "";
    for (const seg of c.segments) { grid.appendChild(makeCard("en", seg)); grid.appendChild(makeCard("no", seg)); }
    updateHighlights(); updateProgress();
    store.lastChapter = ci; save();
  }

  // ---------- Markering / kobling ----------
  function findCard(side, segId) { return grid.querySelector(`.card.${side}[data-seg-id="${segId}"]`); }
  function wordSpan(card, wi) { return card ? card.querySelector(`.w[data-wi="${wi}"]`) : null; }
  function isUncertain(a) { return (store.uncertain[ci] || []).some(u => u.segId === a.segId && u.side === a.side && u.wi === a.wi); }

  function updateHighlights() {
    grid.querySelectorAll(".w-active,.w-linkhot").forEach(e => e.classList.remove("w-active", "w-linkhot"));
    grid.querySelectorAll(".seg-highlight,.seg-current").forEach(e => e.classList.remove("seg-highlight", "seg-current"));
    const bar = document.getElementById("actionbar");
    if (!active) { bar.classList.remove("show"); return; }
    const ownCard = findCard(active.side, active.segId);
    const otherSide = active.side === "en" ? "no" : "en";
    const otherCard = findCard(otherSide, active.segId);
    if (ownCard) ownCard.classList.add("seg-current");
    if (otherCard) otherCard.classList.add("seg-highlight");
    const aw = wordSpan(ownCard, active.wi); if (aw) aw.classList.add("w-active");
    let hasLink = false;
    for (const pair of segLinks(active.segId)) {
      const mineWi = active.side === "en" ? pair[0] : pair[1];
      const theirWi = active.side === "en" ? pair[1] : pair[0];
      if (mineWi === active.wi) { const ps = wordSpan(otherCard, theirWi); if (ps) ps.classList.add("w-linkhot"); hasLink = true; }
    }
    bar.classList.add("show");
    document.getElementById("abSel").innerHTML = `Valgt: <b>«${escapeHtml(active.word)}»</b> (${active.side === "en" ? "engelsk" : "norsk"})`;
    document.getElementById("abHint").textContent = active.side === "en"
      ? "Klikk et norsk ord i samme avsnitt for å koble dem."
      : "Klikk et engelsk ord i samme avsnitt for å koble dem.";
    document.getElementById("abUncertain").textContent = isUncertain(active) ? "✓ Usikker (fjern)" : "❓ Merk usikker";
    document.getElementById("abUnlink").style.display = hasLink ? "" : "none";
  }

  grid.addEventListener("click", (e) => {
    const span = e.target.closest(".w"); if (!span) return;
    const card = e.target.closest(".card");
    const segId = +card.dataset.segId, side = card.dataset.side, wi = +span.dataset.wi, word = span.textContent;
    if (active && active.side !== side && active.segId === segId) {
      const enWi = side === "en" ? wi : active.wi;
      const noWi = side === "no" ? wi : active.wi;
      const links = segLinks(segId).slice();
      const idx = links.findIndex(p => p[0] === enWi && p[1] === noWi);
      if (idx >= 0) links.splice(idx, 1);
      else { for (let i = links.length - 1; i >= 0; i--) if (links[i][0] === enWi || links[i][1] === noWi) links.splice(i, 1); links.push([enWi, noWi]); }
      setSegLinks(segId, links); active = { segId, side, wi, word }; renderChapter(); return;
    }
    active = { segId, side, wi, word }; updateHighlights();
  });
  grid.addEventListener("dblclick", (e) => {
    const card = e.target.closest(".card.no"); if (!card) return;
    editingSeg = +card.dataset.segId; renderChapter();
  });

  // ---------- Handlingslinje ----------
  document.getElementById("abClose").onclick = () => { active = null; updateHighlights(); };
  document.getElementById("abUnlink").onclick = () => {
    if (!active) return;
    setSegLinks(active.segId, segLinks(active.segId).filter(p => (active.side === "en" ? p[0] : p[1]) !== active.wi));
    renderChapter();
  };
  document.getElementById("abUncertain").onclick = () => {
    if (!active) return;
    const list = store.uncertain[ci] || (store.uncertain[ci] = []);
    const i = list.findIndex(u => u.segId === active.segId && u.side === active.side && u.wi === active.wi);
    if (i >= 0) list.splice(i, 1);
    else { const seg = chapter().segments[active.segId]; list.push({ segId: active.segId, side: active.side, wi: active.wi, word: active.word, en: seg.en, no: getNo(active.segId) }); }
    save(); renderChapter(); updateUncCount();
  };
  document.getElementById("abLookup").onclick = () => openLookup();

  // ---------- Oppslag via AI ----------
  function buildPrompts(ctx) {
    const sys = "Du er en erfaren oversetter som hjelper med å oversette en bok fra engelsk til norsk bokmål. " +
      "Svar bare med selve svaret på norsk, kort og oversiktlig. Ikke vis tankegang.";
    let user = `Jeg oversetter denne engelske setningen:\n"${ctx.en}"\n\nJeg er usikker på ordet «${ctx.word}».\n\n`;
    if (ctx.no) user += `Min norske oversettelse så langt:\n"${ctx.no}"\n\n`;
    user += "Gi meg:\n1) Hva «" + ctx.word + "» betyr her.\n2) 1–3 gode forslag til norsk oversettelse av ordet i denne sammenhengen.\n3) En kort merknad om nyanse eller valg.";
    if (ctx.no && ctx.side === "no") user += "\n4) Passer det norske ordet jeg har valgt? Hvis ikke, foreslå bedre.";
    return { sys, user };
  }
  async function errMsg(res) {
    let m = "Feil " + res.status;
    try { const j = await res.json(); m = (j.error && (j.error.message || j.error)) || m; } catch (e) {}
    return new Error(typeof m === "string" ? m : JSON.stringify(m));
  }
  async function callProvider(prov, sys, user) {
    const s = store.settings;
    if (prov === "anthropic") {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "content-type": "application/json", "x-api-key": s.keys.anthropic, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({ model: s.models.anthropic || "claude-opus-4-8", max_tokens: 800, system: sys, messages: [{ role: "user", content: user }] }),
      });
      if (!res.ok) throw await errMsg(res);
      const d = await res.json();
      return (d.content || []).filter(b => b.type === "text").map(b => b.text).join("\n").trim();
    }
    if (prov === "openai") {
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: { "content-type": "application/json", "authorization": "Bearer " + s.keys.openai },
        body: JSON.stringify({ model: s.models.openai || "gpt-4o", messages: [{ role: "system", content: sys }, { role: "user", content: user }] }),
      });
      if (!res.ok) throw await errMsg(res);
      const d = await res.json();
      return ((d.choices && d.choices[0] && d.choices[0].message && d.choices[0].message.content) || "").trim();
    }
    if (prov === "gemini") {
      const model = s.models.gemini || "gemini-2.0-flash";
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(s.keys.gemini)}`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ system_instruction: { parts: [{ text: sys }] }, contents: [{ parts: [{ text: user }] }] }),
      });
      if (!res.ok) throw await errMsg(res);
      const d = await res.json();
      const c = d.candidates && d.candidates[0];
      return ((c && c.content && c.content.parts || []).map(p => p.text || "").join("\n")).trim();
    }
    throw new Error("Ukjent AI.");
  }
  function refreshLookupButtons() {
    ["anthropic", "openai", "gemini"].forEach(prov => {
      const btn = document.querySelector(`.lookup-providers .btn[data-prov="${prov}"]`);
      btn.disabled = !store.settings.keys[prov];
      btn.title = store.settings.keys[prov] ? "Spør " + PROV_NAME[prov] : "Legg inn nøkkel under ⚙︎ først";
    });
  }
  function appendLookupResult(prov, text, isError) {
    const box = document.createElement("div"); box.className = "lookup-result";
    box.innerHTML = `<div class="head"${isError ? ' style="color:#b5482a"' : ''}>${isError ? "Noe gikk galt" : "Svar"} <span class="tagprov">${PROV_NAME[prov]}</span></div>${escapeHtml(text)}`;
    document.getElementById("lookupBody").appendChild(box);
  }
  async function runLookup(prov) {
    if (!store.settings.keys[prov]) { openOverlay("settingsOverlay"); return; }
    const { sys, user } = buildPrompts(lookupCtx);
    const loading = document.createElement("div"); loading.className = "lookup-result"; loading.textContent = "Henter svar fra " + PROV_NAME[prov] + " …";
    document.getElementById("lookupBody").appendChild(loading);
    try {
      const text = await callProvider(prov, sys, user); loading.remove();
      if (text) appendLookupResult(prov, text, false);
      else appendLookupResult(prov, "AI-en ga ikke noe svar denne gangen (kan skyldes innholdsfilter, eller at svaret ble for langt). Prøv igjen, en annen AI, eller bruk «❓ Usikre → Kopier».", true);
    }
    catch (err) { loading.remove(); appendLookupResult(prov, err.message + "\n\n(Sjekk nøkkel/modell under ⚙︎, eller bruk «❓ Usikre → Kopier» for å spørre i en vanlig chat.)", true); }
  }
  function openLookup() {
    if (!active) return;
    const seg = chapter().segments[active.segId];
    lookupCtx = { word: active.word, en: seg.en, no: getNo(active.segId), side: active.side };
    document.getElementById("lookupTitle").textContent = `Oppslag: «${active.word}»`;
    document.getElementById("lookupSub").textContent = truncate(seg.en, 110);
    document.getElementById("lookupBody").innerHTML = "";
    refreshLookupButtons();
    openOverlay("lookupOverlay");
    const anyKey = Object.values(store.settings.keys).some(Boolean);
    if (!anyKey) { appendLookupResult(store.settings.provider, "Du har ikke lagt inn noen AI-nøkkel ennå. Åpne ⚙︎ (tannhjulet) og legg inn minst én – Claude, ChatGPT eller Gemini.", true); return; }
    const def = store.settings.provider;
    if (store.settings.keys[def]) runLookup(def);
    else appendLookupResult(def, `Standard-AI-en (${PROV_NAME[def]}) mangler nøkkel. Trykk en av de aktive knappene øverst, eller legg inn nøkkel under ⚙︎.`, true);
  }
  document.querySelectorAll(".lookup-providers .btn").forEach(b => b.onclick = () => runLookup(b.dataset.prov));
  document.getElementById("lookupClose").onclick = closeOverlays;

  // ---------- Navigasjon ----------
  function goto(i) {
    ci = Math.max(0, Math.min(store.source.chapters.length - 1, i));
    active = null; editingSeg = null; renderChapter(); window.scrollTo(0, 0);
  }
  document.getElementById("prevCh").onclick = () => { if (hasSource()) goto(ci - 1); };
  document.getElementById("nextCh").onclick = () => { if (hasSource()) goto(ci + 1); };
  sel.onchange = () => goto(+sel.value);

  // ---------- Progresjon ----------
  function translatableSegs(c) { return c.segments.filter(s => PROSE.includes(s.type)); }
  function updateProgress() {
    const segs = translatableSegs(chapter());
    const done = segs.filter(s => getNo(s.id)).length;
    document.getElementById("progress").textContent = `${done} / ${segs.length} avsnitt`;
  }
  function updateUncCount() {
    let n = 0; for (const k in store.uncertain) n += (store.uncertain[k] || []).length;
    document.getElementById("uncCount").textContent = n ? `(${n})` : "";
  }

  // ---------- Lim inn hele delen ----------
  const pasteText = document.getElementById("pasteText");
  document.getElementById("pasteChapter").onclick = () => { pasteText.value = ""; updatePasteCount(); openOverlay("pasteOverlay"); setTimeout(() => pasteText.focus(), 0); };
  function pasteTargets() { return chapter().segments.filter(s => s.type === "body"); }
  function updatePasteCount() {
    const paras = pasteText.value.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
    document.getElementById("pasteCount").textContent =
      `${paras.length} avsnitt limt inn · denne delen har ${pasteTargets().length} brødtekst-avsnitt. De legges nedover i rekkefølge – tittel og sitater fyller du i egne felt.`;
  }
  pasteText.addEventListener("input", updatePasteCount);
  document.getElementById("pasteApply").onclick = () => {
    const paras = pasteText.value.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
    pasteTargets().forEach((seg, i) => { if (i < paras.length) { setNo(seg.id, paras[i]); setSegLinks(seg.id, []); } });
    closeOverlays(); renderChapter();
  };
  document.getElementById("pasteCancel").onclick = closeOverlays;

  // ---------- Innstillinger ----------
  document.getElementById("settingsBtn").onclick = () => {
    const s = store.settings;
    document.getElementById("providerSelect").value = s.provider;
    document.getElementById("anthropicKey").value = s.keys.anthropic || "";
    document.getElementById("openaiKey").value = s.keys.openai || "";
    document.getElementById("geminiKey").value = s.keys.gemini || "";
    document.getElementById("anthropicModel").value = s.models.anthropic || "claude-opus-4-8";
    document.getElementById("openaiModel").value = s.models.openai || "";
    document.getElementById("geminiModel").value = s.models.gemini || "";
    openOverlay("settingsOverlay");
  };
  document.getElementById("settingsSave").onclick = () => {
    const s = store.settings;
    s.provider = document.getElementById("providerSelect").value;
    s.keys.anthropic = document.getElementById("anthropicKey").value.trim();
    s.keys.openai = document.getElementById("openaiKey").value.trim();
    s.keys.gemini = document.getElementById("geminiKey").value.trim();
    s.models.anthropic = document.getElementById("anthropicModel").value;
    s.models.openai = document.getElementById("openaiModel").value.trim() || "gpt-4o";
    s.models.gemini = document.getElementById("geminiModel").value.trim() || "gemini-2.0-flash";
    save(); closeOverlays();
  };
  document.getElementById("settingsCancel").onclick = closeOverlays;

  // ---------- Usikre ord ----------
  document.getElementById("uncertainBtn").onclick = () => { renderUncertain(); openOverlay("uncertainOverlay"); };
  function allUncertain() {
    const items = [];
    if (!hasSource()) return items;
    store.source.chapters.forEach((c, k) => (store.uncertain[k] || []).forEach((u, idx) => items.push({ c, k, u, idx })));
    return items;
  }
  function renderUncertain() {
    const wrap = document.getElementById("uncertainList"); wrap.innerHTML = "";
    const items = allUncertain();
    if (!items.length) { wrap.innerHTML = '<p class="empty-hint">Ingen ord merket ennå. Klikk et ord i teksten og velg «Merk usikker».</p>'; return; }
    items.forEach(({ c, k, u, idx }) => {
      const d = document.createElement("div"); d.className = "uncertain-item";
      d.innerHTML = `<button class="rm" title="Fjern">✕</button><span class="word">${escapeHtml(u.word)}</span> ` +
        `<span style="color:var(--muted);font-size:13px">– ${escapeHtml(c.title)} (del ${k + 1})</span>` +
        `<div class="ctx">EN: ${escapeHtml(truncate(u.en, 120))}</div>` + (u.no ? `<div class="ctx">NO: ${escapeHtml(truncate(u.no, 120))}</div>` : "");
      d.querySelector(".rm").onclick = () => { store.uncertain[k].splice(idx, 1); save(); renderUncertain(); updateUncCount(); if (k === ci) renderChapter(); };
      wrap.appendChild(d);
    });
  }
  document.getElementById("uncClose").onclick = closeOverlays;
  document.getElementById("uncCopy").onclick = () => {
    const items = allUncertain();
    if (!items.length) { alert("Ingen ord å kopiere ennå."); return; }
    let txt = "Jeg oversetter en bok fra engelsk til norsk bokmål. Kan du sjekke disse ordene jeg er usikker på? For hvert ord: gi norsk oversettelse og en kort begrunnelse.\n\n";
    items.forEach(({ c, k, u }, i) => {
      txt += `${i + 1}) Ord: «${u.word}»  (${c.title}, del ${k + 1})\n   Engelsk: "${u.en}"\n`;
      if (u.no) txt += `   Min norske tekst: "${u.no}"\n`;
      txt += "\n";
    });
    copyText(txt).then(ok => {
      const b = document.getElementById("uncCopy");
      b.textContent = ok ? "✓ Kopiert!" : "Kunne ikke kopiere";
      setTimeout(() => { b.textContent = "📋 Kopier (til ChatGPT/Claude/Gemini)"; }, 1800);
    });
  };

  // ---------- Hjelp ----------
  document.getElementById("helpBtn").onclick = () => openOverlay("helpOverlay");
  document.getElementById("helpClose").onclick = closeOverlays;

  // ---------- Fil-meny / opplasting ----------
  document.getElementById("menuBtn").onclick = () => {
    let nDone = 0, nTot = 0;
    if (hasSource()) store.source.chapters.forEach((c, k) => c.segments.forEach(s => { if (PROSE.includes(s.type)) { nTot++; if (store.translations[k] && store.translations[k][s.id]) nDone++; } }));
    document.getElementById("menuStats").textContent = hasSource()
      ? `Dokument: «${store.source.name}». Oversatt: ${nDone} av ${nTot} avsnitt.`
      : "Ingen original lastet opp ennå.";
    openOverlay("menuOverlay");
  };
  document.getElementById("menuClose").onclick = closeOverlays;

  const enFile = document.getElementById("enFile"), noFile = document.getElementById("noFile"), importFile = document.getElementById("importFile");
  document.getElementById("esUploadEn").onclick = () => enFile.click();
  document.getElementById("esUploadNo").onclick = () => noFile.click();
  document.getElementById("uploadEn").onclick = () => enFile.click();
  document.getElementById("uploadNo").onclick = () => noFile.click();
  enFile.onchange = (e) => { const f = e.target.files[0]; e.target.value = ""; handleEnglishFile(f); };
  noFile.onchange = (e) => { const f = e.target.files[0]; e.target.value = ""; handleNorwegianFile(f); };

  document.getElementById("exportJson").onclick = () => {
    const blob = JSON.stringify({ source: store.source, translations: store.translations, links: store.links, uncertain: store.uncertain }, null, 2);
    download("oversettelse-sikkerhetskopi.json", blob, "application/json");
  };
  function validSource(src) {
    if (src == null) return true;
    if (typeof src !== "object" || !Array.isArray(src.chapters) || !src.chapters.length) return false;
    return src.chapters.every(c => c && Array.isArray(c.segments));
  }
  document.getElementById("importJson").onclick = () => importFile.click();
  importFile.onchange = (e) => {
    const f = e.target.files[0]; if (!f) return;
    const r = new FileReader();
    r.onload = () => {
      let obj;
      try { obj = JSON.parse(r.result); } catch (err) { alert("Klarte ikke å lese fila: ugyldig format."); return; }
      if (!obj || typeof obj !== "object" || !validSource(obj.source)) {
        alert("Sikkerhetskopien ser skadet ut (mangler gyldig dokumentstruktur). Ingen endring gjort."); return;
      }
      const snapshot = JSON.stringify(store);
      try {
        store.source = obj.source || null;
        if (store.source) store.source.chapters.forEach(c => c.segments.forEach((s, i) => (s.id = i))); // re-indekser defensivt
        store.translations = obj.translations || {};
        store.links = obj.links || {};
        store.uncertain = obj.uncertain || {};
        ci = 0; active = null; editingSeg = null;
        renderAll(); updateUncCount(); save(); closeOverlays();
        alert("Sikkerhetskopi hentet inn.");
      } catch (err) {
        store = JSON.parse(snapshot); ci = 0; renderAll(); updateUncCount();
        alert("Klarte ikke å bruke sikkerhetskopien: " + err.message);
      }
    };
    r.readAsText(f); e.target.value = "";
  };
  document.getElementById("exportText").onclick = () => {
    if (!hasSource()) { alert("Last opp en original først."); return; }
    let out = (store.source.name || "Oversettelse") + "\n\n";
    store.source.chapters.forEach((c, k) => {
      out += "\n========== " + (k + 1) + ". " + c.title + " ==========\n\n";
      c.segments.forEach(s => { if (!PROSE.includes(s.type)) return; const no = (store.translations[k] && store.translations[k][s.id]) || ""; out += (no || "[mangler oversettelse]") + "\n\n"; });
    });
    download("oversettelse-norsk.txt", out, "text/plain");
  };
  document.getElementById("clearAll").onclick = () => {
    if (!confirm("Dette tømmer originalen og hele oversettelsen fra appen (nøklene dine beholdes). Ta gjerne sikkerhetskopi først. Fortsette?")) return;
    const keep = store.settings;
    store = freshStore(); store.settings = keep; store.seenHelp = true;
    ci = 0; active = null; editingSeg = null;
    save(); closeOverlays(); renderAll(); updateUncCount();
  };

  // ---------- Hjelpere ----------
  function openOverlay(id) { closeOverlays(); document.getElementById(id).classList.add("show"); }
  function closeOverlays() { document.querySelectorAll(".overlay.show").forEach(o => o.classList.remove("show")); }
  // Klikk på mørk bakgrunn lukker – men ikke oppslagsvinduet (så et AI-svar ikke mistes ved feilklikk)
  document.querySelectorAll(".overlay").forEach(o => { if (o.id === "lookupOverlay") return; o.addEventListener("click", (e) => { if (e.target === o) closeOverlays(); }); });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") { closeOverlays(); if (active) { active = null; updateHighlights(); } }
    if (editingSeg !== null || !hasSource()) return;
    if (e.key === "ArrowRight" && e.altKey) goto(ci + 1);
    if (e.key === "ArrowLeft" && e.altKey) goto(ci - 1);
  });
  function escapeHtml(s) { return String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  function truncate(s, n) { s = s || ""; return s.length > n ? s.slice(0, n) + "…" : s; }
  function download(name, content, mime) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([content], { type: mime })); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text).then(() => true, () => fallbackCopy(text));
    return Promise.resolve(fallbackCopy(text));
  }
  function fallbackCopy(text) {
    try { const ta = document.createElement("textarea"); ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0"; document.body.appendChild(ta); ta.select(); const ok = document.execCommand("copy"); ta.remove(); return ok; }
    catch (e) { return false; }
  }

  // Lagre umiddelbart hvis siden lukkes (debounce-vinduet kan ellers svelge siste endring)
  window.addEventListener("beforeunload", () => {
    clearTimeout(saveTimer);
    try { localStorage.setItem(KEY, JSON.stringify(store)); } catch (e) {}
  });

  // ---------- Start ----------
  try { renderAll(); }
  catch (err) {                              // skadet lagret tilstand skal ikke låse appen
    store.source = null; save();
    try { renderAll(); } catch (e2) {}
  }
  updateUncCount();
  if (!store.seenHelp) { openOverlay("helpOverlay"); store.seenHelp = true; save(); }
})();

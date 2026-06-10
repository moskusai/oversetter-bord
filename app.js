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
      translations: {}, links: {}, uncertain: {}, align: {},
      project: null,                // { id, name } når arbeidet er knyttet til et skyprosjekt
      settings: {
        provider: "anthropic",
        keys: { anthropic: "", openai: "", gemini: "", deepl: "" },
        models: { anthropic: "claude-opus-4-8", openai: "gpt-4o", gemini: "gemini-2.0-flash" },
        autoAlign: true,
        context: "Dette er en barnebok. Bruk et enkelt, varmt og naturlig norsk som er fint å lese høyt for barn, og behold meningen i originalen.",
      },
      lastChapter: 0, seenHelp: false,
    };
  }
  function load() {
    try {
      const v = JSON.parse(localStorage.getItem(KEY));
      if (v && typeof v === "object" && !Array.isArray(v)) return v;   // avvis korrupt/primitiv lagring
    } catch (e) {}
    // Migrer nøkkel fra tidligere versjon hvis den finnes, og rydd den gamle etterpå
    const s = freshStore();
    try {
      const old = JSON.parse(localStorage.getItem(OLD_KEY));
      if (old && old.settings) {
        if (old.settings.apiKey) s.settings.keys.anthropic = old.settings.apiKey;
        if (old.settings.model) s.settings.models.anthropic = old.settings.model;
      }
      localStorage.removeItem(OLD_KEY);
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
    store.align = store.align || {};
    if (store.settings.autoAlign == null) store.settings.autoAlign = true;
    if (store.project === undefined) store.project = null;
    if (store.lastChapter == null) store.lastChapter = 0;
  })();

  let saveTimer = null, saveWarned = false;
  function save(markCloudDirty) {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      try { localStorage.setItem(KEY, JSON.stringify(store)); saveWarned = false; }
      catch (e) {
        if (!saveWarned) { saveWarned = true; alert("Obs: nettleseren har lite lagringsplass igjen. Ta en arbeidsfil under «≡ Fil» for å være trygg."); }
      }
    }, 200);
    scheduleCloudSave(markCloudDirty !== false);   // skylagring (ren navigasjon merker ikke endring)
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
  // Kapittel-eksplisitte varianter – brukes når et AI-svar kan komme etter at brukeren har byttet kapittel
  function setNoAt(k, segId, text) {
    if (!store.translations[k]) store.translations[k] = {};
    if (text) store.translations[k][segId] = text; else delete store.translations[k][segId];
    save();
  }
  function clearMetaAt(k, segId) {
    if (store.links[k]) delete store.links[k][segId];
    if (store.align[k]) delete store.align[k][segId];
    if (store.uncertain[k]) store.uncertain[k] = store.uncertain[k].filter(u => !(u.segId === segId && u.side === "no"));
    triedAlign.delete(k + ":" + segId);
  }
  // Auto-kobling fra AI (egen lagring så den ikke roter til de manuelle koblingene)
  function getAlign(segId) { return store.align[ci] ? store.align[ci][segId] : undefined; }
  function setAlign(segId, arr) { if (!store.align[ci]) store.align[ci] = {}; store.align[ci][segId] = arr || []; save(); }
  function clearAlign(segId) { if (store.align[ci]) delete store.align[ci][segId]; triedAlign.delete(ci + ":" + segId); }
  function resetAlignCaches() { aligning.clear(); triedAlign.clear(); }

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
  function isHeading2(s) { return s && /(heading|overskrift)\s*2(\b|[^0-9])/i.test(s); }
  function isHeading(s) { return s && /(heading|overskrift|title|tittel)/i.test(s); }
  const SENT_END = /[.!?:"”’'…)»]$/;
  function endsSentence(s) { return SENT_END.test(s.trim()); }
  function startsContinuation(s) {           // ser ut som fortsettelse av forrige setning?
    const m = s.match(/\p{L}/u);
    if (!m) return true;                      // starter med tegn/tall
    return m[0].toLowerCase() === m[0] && m[0].toUpperCase() !== m[0]; // liten forbokstav
  }
  function looksLetterspaced(t) {             // f.eks. "T H E  B O O K"
    const k = t.split(/\s+/).filter(Boolean);
    if (k.length < 3) return false;
    return k.filter(x => x.length === 1).length / k.length > 0.55;
  }
  function despace(t) { return /\S {2,}\S/.test(t) ? t.split(/ {2,}/).map(w => w.replace(/ /g, "")).join(" ") : t; }
  function cleanLabel(s) { s = despace(s).trim(); return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

  // Bygger kapitler. Dropper forord/kolofon foran første kapittel, slår sammen
  // brødtekst-linjer delt midt i en setning, hopper over sidetall, rydder bort
  // spredte fotnoter, og merker testament-skille (Overskrift2) for gruppering.
  function parasToChapters(paras) {
    const firstH1 = paras.findIndex(p => p.text && isHeading1(p.style));
    const used = firstH1 >= 0 ? paras.slice(firstH1) : paras;
    const chapters = []; let cur = null, bodyBuf = "", section = "";
    function flushBody() { if (cur && bodyBuf.trim()) cur.segments.push({ type: "body", en: bodyBuf.trim() }); bodyBuf = ""; }
    function open(t) { flushBody(); cur = { title: (t && t.trim()) || "(uten tittel)", section: section, segments: [] }; chapters.push(cur); }
    function pushSeg(seg) { flushBody(); if (!cur) open("Dokument"); cur.segments.push(seg); }
    for (const { style, text } of used) {
      if (!text) continue;
      if (/^\d{1,4}$/.test(text) && !isHeading(style)) continue;                  // sidetall
      if (isHeading1(style)) { open(text); cur.segments.push({ type: "title", en: text }); continue; }
      if (isHeading2(style)) { section = text; pushSeg({ type: "section", en: text }); continue; } // f.eks. testament-skille
      if (isHeading(style)) { pushSeg({ type: "section", en: text }); continue; }
      if (looksLetterspaced(text)) { pushSeg({ type: "source", en: despace(text) }); continue; }   // spredte fotnoter/referanser
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
  function hasAnyWork() {
    return Object.values(store.translations).some(t => t && Object.keys(t).length) ||
      Object.values(store.uncertain).some(u => u && u.length) ||
      Object.values(store.links).some(l => l && Object.keys(l).length);
  }
  async function handleEnglishFile(file) {
    if (!file) return;
    if (hasSource() && hasAnyWork()) {
      if (!confirm("Dette erstatter originalen og fjerner den pågående oversettelsen (også merkede ord og koblinger). Ta gjerne en arbeidsfil først (≡ Fil). Vil du fortsette?")) return;
    }
    try {
      const chapters = await fileToChapters(file);
      const nSeg = chapters.reduce((a, c) => a + c.segments.length, 0);
      if (!nSeg) throw new Error("Fant ingen tekst i fila.");
      store.source = { name: file.name, chapters };
      store.translations = {}; store.links = {}; store.uncertain = {}; store.align = {};
      store.project = null; setCloudStatus(me ? "ikke lagret i skyen ennå" : "");   // ny bok = ikke samme skyprosjekt
      resetAlignCaches();
      ci = 0; active = null; editingSeg = null;
      save(); closeOverlays(); renderAll(); updateUncCount();
      let msg = `Lest inn «${file.name}»: ${chapters.length} kapitler, ${nSeg} avsnitt.`;
      if (chapters.length === 1 && chapters[0].title === "Dokument")
        msg += "\n\nFant ingen kapitteloverskrifter i fila, så alt ligger som ett kapittel. Hvis boka har kapitler, sjekk at overskriftene har overskrift-stil i Word.";
      alert(msg);
    } catch (err) {
      alert("Klarte ikke å lese fila: " + err.message);
    }
  }
  async function handleNorwegianFile(file) {
    if (!file) return;
    if (!hasSource()) { alert("Last opp den engelske originalen først, så vet appen hvor den norske teksten skal ligge."); return; }
    try {
      const noChapters = await fileToChapters(file);
      if (!noChapters.length) throw new Error("Fant ingen tekst i fila.");
      const enChapters = store.source.chapters;
      const had = Object.keys(store.translations).length > 0;
      const same = noChapters.length === enChapters.length;
      const msg = `Stiller den norske fila opp mot originalen – kapittel for kapittel.\n\n` +
        `Original: ${enChapters.length} kapitler. Norsk fil: ${noChapters.length} kapitler.` +
        (same ? " (Samme antall – stiller presist opp.)" : "\n\nMERK: ulikt antall kapitler, så noe kan bli forskjøvet. Sjekk gjerne etterpå.") +
        (had ? "\n\nDette ERSTATTER den norske teksten som alt ligger der (ta gjerne en arbeidsfil først)." : "") +
        "\n\nFortsette?";
      if (!confirm(msg)) return;
      store.translations = {}; store.links = {}; store.uncertain = {}; store.align = {};
      resetAlignCaches(); active = null; editingSeg = null;
      enChapters.forEach(c => delete c.sectionNo);   // ikke la gamle norske gruppe-etiketter henge igjen
      const n = Math.min(noChapters.length, enChapters.length);
      let filled = 0;
      for (let k = 0; k < n; k++) {
        const enSegs = enChapters[k].segments.filter(s => PROSE.includes(s.type));
        const noSegs = noChapters[k].segments.filter(s => PROSE.includes(s.type));
        if (noChapters[k].section) enChapters[k].sectionNo = noChapters[k].section; // norsk gruppe-etikett
        const m = Math.min(enSegs.length, noSegs.length);
        for (let j = 0; j < m; j++) {
          if (!store.translations[k]) store.translations[k] = {};
          store.translations[k][enSegs[j].id] = noSegs[j].en;
          filled++;
        }
      }
      save(); closeOverlays(); renderAll(); updateUncCount();
      alert(`Ferdig: stilte opp ${n} kapitler og ${filled} avsnitt.`);
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
    // kapittelvelger – gruppert etter testament/seksjon
    sel.innerHTML = "";
    let curGroup = "__init__", groupEl = null;
    store.source.chapters.forEach((c, i) => {
      const g = c.sectionNo || c.section || "";
      if (g !== curGroup) {
        curGroup = g;
        if (g) { groupEl = document.createElement("optgroup"); groupEl.label = cleanLabel(g); sel.appendChild(groupEl); }
        else groupEl = null;
      }
      const o = document.createElement("option");
      o.value = i; o.textContent = (i + 1) + ". " + c.title;
      (groupEl || sel).appendChild(o);
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
      // mousedown stjeler ellers fokus fra boksen FØR klikket – da ville blur-lagringen
      // slå til først og «Angre endring» aldri få angret (Chrome/Edge)
      ok.addEventListener("mousedown", (e) => { e.preventDefault(); finish(true); });
      cancel.addEventListener("mousedown", (e) => { e.preventDefault(); finish(false); });
      ok.onclick = () => finish(true);          // tastatur-fallback
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
    const tools = document.createElement("div"); tools.className = "no-tools";
    const mk = (label, title, fn) => { const b = document.createElement("button"); b.className = "tbtn"; b.textContent = label; b.title = title; b.onclick = (e) => { e.stopPropagation(); fn(); }; return b; };
    tools.appendChild(mk("✎", "Rediger / lim inn norsk", () => { editingSeg = seg.id; renderChapter(); }));
    if (PROSE.includes(seg.type)) {
      tools.appendChild(mk("↧", "Sett inn tom linje her – skyver den norske teksten nedover (når norsk ligger ett hakk for høyt)", () => insertNoGap(seg.id)));
      tools.appendChild(mk("🗑", "Fjern denne norske linja – skyver resten oppover (når det er ett avsnitt for mye)", () => deleteNoCell(seg.id)));
    }
    card.appendChild(tools);
    return card;
  }
  function commitEdit(segId, value) {
    const old = getNo(segId), v = value.trim();
    if (v !== old) {
      setNo(segId, v);
      setSegLinks(segId, []); clearAlign(segId);
      if (store.uncertain[ci]) { store.uncertain[ci] = store.uncertain[ci].filter(u => !(u.segId === segId && u.side === "no")); save(); }
    }
    editingSeg = null; active = null; renderChapter();
  }
  // ---------- Manuell justering (skyv norsk kolonne opp/ned) ----------
  function proseSegs() { return chapter().segments.filter(s => PROSE.includes(s.type)); }
  function clearMetaFrom(segs, idx) {           // nullstill koblinger/merker kun fra skiftepunktet og nedover
    for (let j = idx; j < segs.length; j++) clearMetaAt(ci, segs[j].id);
    save();
  }
  function insertNoGap(segId) {                 // sett inn tom linje her -> skyv norsk nedover
    const segs = proseSegs(), idx = segs.findIndex(s => s.id === segId);
    if (idx < 0) return;
    if (idx === segs.length - 1) { alert("Dette er siste avsnitt i kapittelet – det er ingenting å skyve nedover herfra."); return; }
    const texts = segs.map(s => getNo(s.id));
    texts.splice(idx, 0, "");
    const overflow = texts.pop();               // siste faller ut – ikke mist tekst
    if (overflow) texts[texts.length - 1] = (texts[texts.length - 1] ? texts[texts.length - 1] + " " : "") + overflow;
    segs.forEach((s, j) => setNo(s.id, texts[j]));
    clearMetaFrom(segs, idx); active = null; editingSeg = null; save(); renderChapter(); updateUncCount();
    if (overflow) alert("De to nederste avsnittene i kapittelet ble slått sammen for ikke å miste tekst – sjekk slutten av kapittelet.");
  }
  function deleteNoCell(segId) {                 // fjern denne linja -> skyv norsk oppover
    const segs = proseSegs(), idx = segs.findIndex(s => s.id === segId);
    if (idx < 0) return;
    const cur = getNo(segId);
    if (cur && !confirm(`Dette fjerner avsnittet «${truncate(cur, 60)}» og skyver resten oppover. Fortsette?`)) return;
    const texts = segs.map(s => getNo(s.id));
    texts.splice(idx, 1); texts.push("");
    segs.forEach((s, j) => setNo(s.id, texts[j]));
    clearMetaFrom(segs, idx); active = null; editingSeg = null; save(); renderChapter(); updateUncCount();
  }

  function renderChapter() {
    const c = chapter();
    document.getElementById("enTitle").textContent = c.title;
    const titleSeg = c.segments.find(s => s.type === "title");
    document.getElementById("noTitle").textContent = (titleSeg && getNo(titleSeg.id)) || "—";
    sel.value = ci;
    document.getElementById("prevCh").disabled = ci === 0;
    document.getElementById("nextCh").disabled = ci === store.source.chapters.length - 1;
    grid.innerHTML = "";
    for (const seg of c.segments) { grid.appendChild(makeCard("en", seg)); grid.appendChild(makeCard("no", seg)); }
    updateHighlights(); updateProgress();
    store.lastChapter = ci; save(false);        // navigasjon er ikke en innholdsendring
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
    const partners = segLinks(active.segId).concat(getAlign(active.segId) || []);
    for (const pair of partners) {
      const mineWi = active.side === "en" ? pair[0] : pair[1];
      const theirWi = active.side === "en" ? pair[1] : pair[0];
      if (mineWi === active.wi) { const ps = wordSpan(otherCard, theirWi); if (ps) ps.classList.add("w-linkhot"); }
    }
    hasLink = partners.some(p => (active.side === "en" ? p[0] : p[1]) === active.wi);
    bar.classList.add("show");
    document.getElementById("abSel").innerHTML = `Valgt: <b>«${escapeHtml(active.word)}»</b> (${active.side === "en" ? "engelsk" : "norsk"})`;
    document.getElementById("abHint").textContent =
      (active.side === "en" && !getNo(active.segId)) ? "Skriv inn den norske oversettelsen først – da kan ordene kobles."
      : active.side === "en" ? "Klikk et norsk ord i samme avsnitt for å koble dem."
      : "Klikk et engelsk ord i samme avsnitt for å koble dem.";
    document.getElementById("abUncertain").textContent = isUncertain(active) ? "✓ Usikker (fjern)" : "❓ Merk usikker";
    document.getElementById("abUnlink").style.display = hasLink ? "" : "none";
  }

  grid.addEventListener("click", (e) => {
    const span = e.target.closest(".w"); if (!span) return;
    if (e.detail > 1) return;                 // andre klikk i et dobbeltklikk skal ikke lage kobling
    const card = e.target.closest(".card");
    const segId = +card.dataset.segId, side = card.dataset.side, wi = +span.dataset.wi, word = span.textContent;
    if (active && active.side !== side && active.segId === segId) {
      const enWi = side === "en" ? wi : active.wi;
      const noWi = side === "no" ? wi : active.wi;
      const links = segLinks(segId).slice();
      const idx = links.findIndex(p => p[0] === enWi && p[1] === noWi);
      if (idx >= 0) links.splice(idx, 1);
      else { for (let i = links.length - 1; i >= 0; i--) if (links[i][0] === enWi || links[i][1] === noWi) links.splice(i, 1); links.push([enWi, noWi]); }
      // manuell kobling overstyrer en eventuell AI-kobling for de samme ordene
      const al = getAlign(segId);
      if (al && al.length) setAlign(segId, al.filter(p => p[0] !== enWi && p[1] !== noWi));
      setSegLinks(segId, links); active = { segId, side, wi, word }; renderChapter(); return;
    }
    active = { segId, side, wi, word }; updateHighlights();
    autoAlignIfNeeded(segId);
  });
  grid.addEventListener("dblclick", (e) => {
    const card = e.target.closest(".card.no"); if (!card) return;
    editingSeg = +card.dataset.segId; renderChapter();
  });

  // ---------- Handlingslinje ----------
  document.getElementById("abClose").onclick = () => { active = null; updateHighlights(); };
  document.getElementById("abUnlink").onclick = () => {
    if (!active) return;
    const mine = (p) => (active.side === "en" ? p[0] : p[1]) !== active.wi;
    setSegLinks(active.segId, segLinks(active.segId).filter(mine));
    const al = getAlign(active.segId);
    if (al && al.length) setAlign(active.segId, al.filter(mine));   // fjern også AI-koblingen for ordet
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
  document.getElementById("abCopy").onclick = () => {
    if (!active) return;
    const seg = chapter().segments[active.segId];
    const no = getNo(active.segId);
    let txt = aiContext() + "Jeg oversetter en bok fra engelsk til norsk bokmål. Her er et avsnitt med original og forslag til oversettelse:\n\n" +
      `ENGELSK:\n"${seg.en}"\n\nNORSK (forslag):\n"${no || "(ikke oversatt ennå)"}"\n`;
    if (active.word) txt += `\nJeg er spesielt usikker på ordet «${active.word}».`;
    copyText(txt).then(ok => {
      const b = document.getElementById("abCopy"), o = b.textContent;
      b.textContent = ok ? "✓ Kopiert!" : "Kunne ikke kopiere";
      setTimeout(() => { b.textContent = o; }, 1500);
    });
  };

  // ---------- Oppslag via AI ----------
  function aiContext() { const c = (store.settings.context || "").trim(); return c ? "Kontekst om boka: " + c + "\n\n" : ""; }
  function buildPrompts(ctx) {
    const sys = "Du er en erfaren oversetter som hjelper med å oversette en bok fra engelsk til norsk bokmål. " +
      "Svar bare med selve svaret på norsk, kort og oversiktlig. Ikke vis tankegang.";
    let user = aiContext() + `Jeg oversetter denne engelske teksten:\n"${ctx.en}"\n\n`;
    if (ctx.no) user += `Min norske oversettelse så langt:\n"${ctx.no}"\n\n`;
    if (ctx.side === "no") {
      user += `I den norske teksten har jeg brukt ordet «${ctx.word}», og jeg er usikker på det.\n\n` +
        "Gi meg:\n1) Hvilket/hvilke engelske ord det svarer til her.\n2) Om det er et godt valg i denne sammenhengen.\n3) 1–3 alternative norske ord hvis det finnes bedre.";
    } else {
      user += `Jeg er usikker på det engelske ordet «${ctx.word}».\n\n` +
        "Gi meg:\n1) Hva «" + ctx.word + "» betyr her.\n2) 1–3 gode forslag til norsk oversettelse av ordet i denne sammenhengen.\n3) En kort merknad om nyanse eller valg.";
      if (ctx.no) user += "\n4) Har jeg alt truffet godt i min norske tekst? Hvis ikke, foreslå bedre.";
    }
    return { sys, user };
  }
  async function errMsg(res) {
    let raw = "";
    try { const j = await res.json(); raw = (j.error && (j.error.message || j.error)) || ""; if (typeof raw !== "string") raw = JSON.stringify(raw); } catch (e) {}
    const FRIENDLY = {
      401: "Nøkkelen ser ut til å være feil eller utløpt – sjekk den under ⚙︎.",
      403: "Nøkkelen har ikke tilgang – sjekk den under ⚙︎.",
      404: "Fant ikke AI-modellen – sjekk modellnavnet under ⚙︎.",
      429: "AI-tjenesten er opptatt akkurat nå – vent litt og prøv igjen.",
    };
    const base = FRIENDLY[res.status] || (res.status >= 500 ? "AI-tjenesten har trøbbel akkurat nå – prøv igjen om litt." : "Feil " + res.status + " fra AI-tjenesten.");
    return new Error(base + (raw ? " (" + truncate(raw, 120) + ")" : ""));
  }
  function netFetch(url, opts) {                 // vennlig melding ved nettverksbrudd
    return fetch(url, opts).catch(() => { throw new Error("Fikk ikke kontakt med AI-tjenesten – sjekk internettforbindelsen og prøv igjen."); });
  }
  let lastCallTruncated = false;   // satt av callProvider når svaret ble kuttet på token-taket
  async function callProvider(prov, sys, user, maxTokens) {
    const s = store.settings;
    lastCallTruncated = false;
    if (prov === "anthropic") {
      const res = await netFetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "content-type": "application/json", "x-api-key": s.keys.anthropic, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({ model: s.models.anthropic || "claude-opus-4-8", max_tokens: maxTokens || 800, system: sys, messages: [{ role: "user", content: user }] }),
      });
      if (!res.ok) throw await errMsg(res);
      const d = await res.json();
      lastCallTruncated = d.stop_reason === "max_tokens";
      return (d.content || []).filter(b => b.type === "text").map(b => b.text).join("\n").trim();
    }
    if (prov === "openai") {
      const model = s.models.openai || "gpt-4o";
      const body = { model, messages: [{ role: "system", content: sys }, { role: "user", content: user }] };
      // resonneringsmodeller bruker skjulte tenke-tokens av samme budsjett – gi romsligere tak
      if (maxTokens) body.max_completion_tokens = /^(o\d|gpt-5)/i.test(model) ? Math.max(maxTokens, 4000) : maxTokens;
      const res = await netFetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: { "content-type": "application/json", "authorization": "Bearer " + s.keys.openai },
        body: JSON.stringify(body),
      });
      if (!res.ok) throw await errMsg(res);
      const d = await res.json();
      const ch = d.choices && d.choices[0];
      lastCallTruncated = !!ch && ch.finish_reason === "length";
      return ((ch && ch.message && ch.message.content) || "").trim();
    }
    if (prov === "gemini") {
      const model = s.models.gemini || "gemini-2.0-flash";
      const body = { system_instruction: { parts: [{ text: sys }] }, contents: [{ parts: [{ text: user }] }] };
      if (maxTokens) {
        // eldre modeller har 8192-tak; nyere (2.5/3.x) tåler mye mer og bruker tenke-tokens av samme budsjett
        const cap = /gemini-(1\.|2\.0)/i.test(model) ? 8192 : 65536;
        const want = /gemini-(1\.|2\.0)/i.test(model) ? maxTokens : Math.max(maxTokens, 4000);
        body.generationConfig = { maxOutputTokens: Math.min(want, cap) };
      }
      const res = await netFetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
        method: "POST",
        headers: { "content-type": "application/json", "x-goog-api-key": s.keys.gemini },   // nøkkel i header, ikke i URL
        body: JSON.stringify(body),
      });
      if (!res.ok) throw await errMsg(res);
      const d = await res.json();
      const c = d.candidates && d.candidates[0];
      lastCallTruncated = !!c && c.finishReason === "MAX_TOKENS";
      return ((c && c.content && c.content.parts || []).map(p => p.text || "").join("\n")).trim();
    }
    throw new Error("Ukjent AI.");
  }

  // ---------- Auto-kobling av ord (AI) ----------
  const aligning = new Set(), triedAlign = new Set();
  async function alignWords(prov, enText, noText) {
    const enW = tokenize(enText).filter(t => t.w), noW = tokenize(noText).filter(t => t.w);
    if (!enW.length || !noW.length) return [];
    const enList = enW.map((t, i) => i + ":" + t.t).join("  ");
    const noList = noW.map((t, i) => i + ":" + t.t).join("  ");
    const sys = "Du kobler ord mellom en engelsk setning og dens norske oversettelse. Svar KUN med JSON, ingen forklaring.";
    const user = aiContext() + `Engelske ord (indeks:ord):\n${enList}\n\nNorske ord (indeks:ord):\n${noList}\n\n` +
      "For hvert engelske ord som har en tydelig motpart i den norske teksten, gi paret [engelskIndeks, norskIndeks]. " +
      "Flere engelske ord kan peke på samme norske ord (sammensatte ord), og samme engelske ord kan peke på flere norske – ta med alle slike par. " +
      "Hopp over ord uten tydelig motpart. Svar KUN med en JSON-liste, f.eks. [[0,1],[2,0]].";
    const resp = await callProvider(prov, sys, user, Math.min(4000, 200 + enW.length * 12));   // skaler med avsnittslengde
    const m = resp && resp.match(/\[[\s\S]*\]/);
    if (!m) throw new Error("ugyldig svar");      // forbigående – ikke lås segmentet (catch cacher ikke)
    let arr;
    try { arr = JSON.parse(m[0]); }
    catch (e) {
      // avkuttet JSON? behold de hele parene fram til siste komplette
      const cut = m[0].lastIndexOf("],");
      if (cut > 0) { try { arr = JSON.parse(m[0].slice(0, cut + 1) + "]"); } catch (e2) { throw new Error("ugyldig svar"); } }
      else throw new Error("ugyldig svar");
    }
    if (!Array.isArray(arr)) throw new Error("ugyldig svar");
    const out = [];
    for (const p of arr) if (Array.isArray(p) && p.length === 2 && Number.isInteger(p[0]) && Number.isInteger(p[1]) &&
      p[0] >= 0 && p[0] < enW.length && p[1] >= 0 && p[1] < noW.length) out.push([p[0], p[1]]);
    return out;
  }
  function autoAlignIfNeeded(segId) {
    if (!store.settings.autoAlign) return;
    const def = store.settings.provider;
    if (!store.settings.keys[def]) return;
    const seg = chapter().segments[segId];
    if (!seg || !PROSE.includes(seg.type)) return;
    const noText = getNo(segId);
    if (!noText) return;
    if (getAlign(segId) !== undefined || segLinks(segId).length) return;   // alt koblet/forsøkt
    const tag = ci + ":" + segId;
    if (aligning.has(tag) || triedAlign.has(tag)) return;
    aligning.add(tag);
    const myChapter = ci, myText = noText;
    if (active && active.segId === segId) document.getElementById("abHint").textContent = "🔗 Kobler ord med AI …";
    alignWords(def, seg.en, noText).then(pairs => {
      aligning.delete(tag);
      // forkast svaret hvis kapittelet er byttet, teksten endret, eller en manuell kobling er laget i mellomtiden
      if (myChapter !== ci || getNo(segId) !== myText || segLinks(segId).length) return;
      if (pairs.length) setAlign(segId, pairs);
      else triedAlign.add(tag);
      // IKKE renderChapter her – det ville ødelagt en åpen rediger-boks. Markering holder.
      if (active && active.segId === segId) updateHighlights();
    }).catch(() => { aligning.delete(tag); if (active && active.segId === segId) updateHighlights(); });
    // merk: feil (nett/ratelimit) caches ikke – nytt klikk prøver igjen
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
  const lookupBusy = new Set();
  async function runLookup(prov) {
    if (!store.settings.keys[prov]) { openOverlay("settingsOverlay"); return; }
    if (lookupBusy.has(prov)) return;            // ikke send dobbelt ved utålmodige klikk
    lookupBusy.add(prov);
    const { sys, user } = buildPrompts(lookupCtx);
    const loading = document.createElement("div"); loading.className = "lookup-result"; loading.textContent = "Henter svar fra " + PROV_NAME[prov] + " …";
    document.getElementById("lookupBody").appendChild(loading);
    try {
      const text = await callProvider(prov, sys, user, 1500); loading.remove();
      if (text) appendLookupResult(prov, text + (lastCallTruncated ? "\n\n(OBS: svaret ble kuttet – slutten kan mangle.)" : ""), false);
      else appendLookupResult(prov, "AI-en ga ikke noe svar denne gangen (kan skyldes innholdsfilter, eller at svaret ble for langt). Prøv igjen, en annen AI, eller bruk «❓ Usikre → Kopier».", true);
    }
    catch (err) { loading.remove(); appendLookupResult(prov, err.message + "\n\n(Sjekk nøkkel/modell under ⚙︎, eller bruk «❓ Usikre → Kopier» for å spørre i en vanlig chat.)", true); }
    finally { lookupBusy.delete(prov); }
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

  // ---------- Lim inn hele kapittelet ----------
  const pasteText = document.getElementById("pasteText");
  document.getElementById("pasteChapter").onclick = () => {
    if (!pasteTargets().length) { alert("Dette kapittelet har ingen brødtekst-avsnitt å lime inn i – bla til et annet kapittel."); return; }
    // ikke kast tekst som alt står der (f.eks. etter feilklikk på bakgrunnen)
    updatePasteCount(); openOverlay("pasteOverlay"); setTimeout(() => pasteText.focus(), 0);
  };
  function pasteTargets() { return chapter().segments.filter(s => s.type === "body"); }
  function updatePasteCount() {
    const paras = pasteText.value.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
    document.getElementById("pasteCount").textContent =
      `${paras.length} avsnitt limt inn · dette kapittelet har ${pasteTargets().length} brødtekst-avsnitt. De legges nedover i rekkefølge – tittel og sitater fyller du i egne felt.`;
  }
  pasteText.addEventListener("input", updatePasteCount);
  document.getElementById("pasteApply").onclick = () => {
    const paras = pasteText.value.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
    if (!paras.length) { alert("Du har ikke limt inn noen tekst ennå."); return; }
    const targets = pasteTargets();
    targets.forEach((seg, i) => { if (i < paras.length) { setNo(seg.id, paras[i]); setSegLinks(seg.id, []); clearMetaAt(ci, seg.id); } });
    pasteText.value = "";                       // nullstill først etter vellykket innsetting
    active = null; closeOverlays(); renderChapter(); updateUncCount();
    if (paras.length !== targets.length)
      alert(`Satt inn ${Math.min(paras.length, targets.length)} avsnitt. Du limte inn ${paras.length}, kapittelet har ${targets.length} brødtekst-avsnitt – sjekk gjerne at de står på rett plass (bruk ↧/🗑 ved behov).`);
  };
  document.getElementById("pasteCancel").onclick = closeOverlays;

  // ---------- Forbedre kapittelet med AI ----------
  function chapterProse() { return chapter().segments.filter(s => PROSE.includes(s.type)); }
  function chapterBodies() { return chapter().segments.filter(s => s.type === "body"); }
  function btnFlash(btn, ok) { const o = btn.textContent; btn.textContent = ok ? "✓ Kopiert!" : "Kunne ikke kopiere"; setTimeout(() => { btn.textContent = o; }, 1500); }
  function buildOptimizePrompt() {
    const bodies = chapterBodies();
    const sys = "Du er en dyktig oversetter som forbedrer en oversettelse til norsk bokmål. Bruk konsekvent bokmål og norske anførselstegn («…»). Svar bare med selve teksten, ingen forklaring.";
    let user = aiContext() + "Nedenfor er et kapittel fra boka. For hvert avsnitt står den engelske originalen (EN) og en norsk oversettelse (NO).\n\n" +
      "Forbedre den NORSKE oversettelsen så den blir naturlig, korrekt og god – behold betydningen og NØYAKTIG samme antall avsnitt. " +
      "Der NO står som (mangler), oversett den engelske originalen til norsk bokmål i stedet.\n\n" +
      "Svar med KUN den norske teksten: ett avsnitt om gangen, med én tom linje mellom hvert avsnitt, i samme rekkefølge – alltid like mange avsnitt som det er nummererte blokker. Ikke ta med engelsk, ingen avsnittsnummer, ikke noe annet.\n\n" +
      `=== ${chapter().title} ===\n\n`;
    bodies.forEach((s, i) => { user += `[${i + 1}]\nEN: ${s.en}\nNO: ${getNo(s.id) || "(mangler)"}\n\n`; });
    return { sys, user, bodies };
  }
  function applyOptimized(text) {
    const paras = text.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
    const bodies = chapterBodies();
    const n = Math.min(paras.length, bodies.length);
    for (let i = 0; i < n; i++) { setNo(bodies[i].id, paras[i]); setSegLinks(bodies[i].id, []); clearMetaAt(ci, bodies[i].id); }
    active = null; closeOverlays(); renderChapter(); updateUncCount();
    if (paras.length !== bodies.length)
      alert(`Satt inn ${n} avsnitt. AI-svaret hadde ${paras.length}, kapittelet har ${bodies.length} brødtekst-avsnitt – sjekk gjerne at de står på rett plass (bruk ↧/🗑 ved behov).`);
  }
  let distributing = false;
  async function distributeWithAI(pasted) {
    if (!pasted.trim()) return;
    if (distributing) return;                   // ikke send dobbelt
    const count = document.getElementById("optCount");
    const def = store.settings.provider;
    if (!store.settings.keys[def]) {
      count.textContent = "Du må legge inn en AI-nøkkel under ⚙︎ (øverst til høyre) først. Teksten din står trygt her imens – eller bruk «Sett inn på rad».";
      return;
    }
    // Bare felt som faktisk telles og eksporteres (PROSE) – ellers kan tekst "forsvinne" fra den ferdige boka
    const myChapter = ci, mySource = store.source;
    const segs = chapter().segments.filter(s => PROSE.includes(s.type) && s.en && s.en.trim());
    const typeLabel = { title: "tittel", body: "brødtekst", quote: "sitat", section: "seksjon", note: "merknad" };
    const list = segs.map(s => `[${s.id}] (${typeLabel[s.type] || s.type}) ${s.en}`).join("\n");
    const sys = "Du fordeler en ferdig norsk oversettelse på de riktige feltene i et kapittel. Svar KUN med gyldig JSON, ingen forklaring.";
    const user = aiContext() + `Feltene i kapittelet (indeks, type, engelsk original):\n${list}\n\n` +
      `Den norske oversettelsen av hele kapittelet:\n"""\n${pasted}\n"""\n\n` +
      `Fordel den norske teksten på riktig felt ut fra mening og rekkefølge. Bruk teksten ORDRETT – ikke endre, forkort eller omskriv noe. ` +
      `Returner KUN et JSON-objekt der nøklene er tallene i klammene, f.eks. {"0": "Norsk tittel", "2": "Første avsnitt …"}. Ikke ta med engelsk, ingen forklaring.`;
    distributing = true;
    const aiBtn = document.getElementById("optApplyAI"), rowBtn = document.getElementById("optApply");
    aiBtn.disabled = true; rowBtn.disabled = true;
    count.textContent = "🤖 Fordeler teksten med AI …";
    try {
      const resp = await callProvider(def, sys, user, 16000);
      if (lastCallTruncated) throw new Error("Svaret ble for langt og ble kuttet. Prøv igjen, eller del kapittelet i to og lim inn halvparten om gangen.");
      const m = resp && resp.match(/\{[\s\S]*\}/);
      if (!m) throw new Error("Fikk ikke et gyldig svar fra AI-en. Prøv igjen, eller bruk «Sett inn på rad».");
      let obj;
      try { obj = JSON.parse(m[0]); }
      catch (e) { throw new Error("Fikk ikke et gyldig svar fra AI-en. Prøv igjen, eller bruk «Sett inn på rad»."); }
      if (store.source !== mySource) throw new Error("Dokumentet ble byttet mens AI-en jobbet – svaret er forkastet.");   // ny original/arbeidsfil/tømt underveis
      let n = 0;
      Object.keys(obj).forEach(k => {
        const id = parseInt(k, 10), seg = segs.find(s => s.id === id);
        if (seg && typeof obj[k] === "string" && obj[k].trim()) {
          // skriv alltid til kapittelet svaret gjelder – selv om brukeren har blatt videre
          setNoAt(myChapter, seg.id, obj[k].trim());
          clearMetaAt(myChapter, seg.id);
          n++;
        }
      });
      if (!n) { count.textContent = "AI-en klarte ikke å fordele teksten. Prøv igjen, eller bruk «Sett inn på rad». Teksten din står fortsatt her."; return; }
      active = null; closeOverlays();
      if (ci === myChapter) renderChapter(); else renderAll();
      updateUncCount();
      alert(`AI fordelte den norske teksten på ${n} felt i kapittelet «${store.source.chapters[myChapter].title}». Sjekk gjerne gjennom.`);
    } catch (err) { count.textContent = "Feil: " + err.message; }
    finally { distributing = false; aiBtn.disabled = false; rowBtn.disabled = false; }
  }

  const optPaste = document.getElementById("optPaste");
  function updateOptCount() {
    const paras = optPaste.value.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
    document.getElementById("optCount").textContent = optPaste.value.trim()
      ? `${paras.length} avsnitt i svaret · kapittelet har ${chapterBodies().length} brødtekst-avsnitt.` : "";
  }
  document.getElementById("optimizeBtn").onclick = () => {
    if (!chapterBodies().length) { alert("Dette kapittelet har ingen brødtekst å forbedre – bla til et annet kapittel."); return; }
    // ikke kast tekst brukeren alt har limt inn (f.eks. hvis modalen ble lukket ved uhell)
    updateOptCount();
    const def = store.settings.provider;
    const hasKey = store.settings.keys[def];
    const row = document.getElementById("optApiRow");
    row.style.display = hasKey ? "" : "none";
    document.getElementById("optProvName").textContent = PROV_NAME[def] || "";
    document.getElementById("optFetchStatus").textContent = "";
    document.getElementById("optDeeplRow").style.display = cloudAvailable ? "" : "none";
    document.getElementById("optDeeplStatus").textContent = "";
    openOverlay("optimizeOverlay");
  };
  document.getElementById("optCopyAI").onclick = (e) => { const { user } = buildOptimizePrompt(); copyText(user).then(ok => btnFlash(e.target, ok)); };
  // Brødtekst, samme som «Sett inn på rad» – ellers forskyver tittel/sitat alt ved retur-liming
  document.getElementById("optCopyEn").onclick = (e) => { copyText(chapterBodies().map(s => s.en).join("\n\n")).then(ok => btnFlash(e.target, ok)); };
  document.getElementById("optCopyNo").onclick = (e) => { copyText(chapterBodies().map(s => getNo(s.id) || "").join("\n\n").trim()).then(ok => btnFlash(e.target, ok)); };
  optPaste.addEventListener("input", updateOptCount);
  document.getElementById("optApply").onclick = () => { if (optPaste.value.trim()) applyOptimized(optPaste.value); else closeOverlays(); };
  document.getElementById("optApplyAI").onclick = () => distributeWithAI(optPaste.value);
  document.getElementById("optClose").onclick = closeOverlays;
  let fetching = false;
  document.getElementById("optFetch").onclick = async () => {
    const def = store.settings.provider;
    if (!store.settings.keys[def]) { openOverlay("settingsOverlay"); return; }
    if (fetching) return;                        // ikke send dobbelt
    fetching = true;
    const btn = document.getElementById("optFetch"); btn.disabled = true;
    const status = document.getElementById("optFetchStatus");
    status.textContent = "Henter fra " + PROV_NAME[def] + " …";
    const myChapter = ci;
    const { sys, user } = buildOptimizePrompt();
    try {
      const text = await callProvider(def, sys, user, 8000);
      if (ci !== myChapter) { status.textContent = "Kapittelet ble byttet mens AI-en jobbet – gå tilbake og hent på nytt."; return; }
      optPaste.value = text || ""; updateOptCount();
      status.textContent = !text ? "AI-en ga tomt svar. Prøv igjen."
        : lastCallTruncated ? "OBS: Svaret ble kuttet (for langt kapittel). Se gjennom – slutten kan mangle."
        : "Hentet – se gjennom og trykk «Sett inn på rad» (eller «Sett inn med AI» hvis avsnittene ikke stemmer).";
    } catch (err) { status.textContent = "Feil: " + err.message; }
    finally { fetching = false; btn.disabled = false; }
  };

  // ---------- Innstillinger ----------
  document.getElementById("settingsBtn").onclick = () => {
    const s = store.settings;
    document.getElementById("providerSelect").value = s.provider;
    document.getElementById("anthropicKey").value = s.keys.anthropic || "";
    document.getElementById("openaiKey").value = s.keys.openai || "";
    document.getElementById("geminiKey").value = s.keys.gemini || "";
    document.getElementById("deeplKey").value = s.keys.deepl || "";
    // ukjent lagret modell (fra eldre versjon) legges til som eget valg så den ikke mistes
    const setSel = (id, val) => {
      const el = document.getElementById(id);
      if (val && ![...el.options].some(o => o.value === val)) {
        const o = document.createElement("option"); o.value = val; o.textContent = val; el.appendChild(o);
      }
      el.value = val || el.options[0].value;
    };
    setSel("anthropicModel", s.models.anthropic || "claude-opus-4-8");
    setSel("openaiModel", s.models.openai || "gpt-4o");
    setSel("geminiModel", s.models.gemini || "gemini-2.0-flash");
    document.getElementById("autoAlign").checked = s.autoAlign !== false;
    document.getElementById("aiContext").value = s.context || "";
    openOverlay("settingsOverlay");
  };
  document.getElementById("settingsSave").onclick = () => {
    const s = store.settings;
    s.provider = document.getElementById("providerSelect").value;
    s.keys.anthropic = document.getElementById("anthropicKey").value.trim();
    s.keys.openai = document.getElementById("openaiKey").value.trim();
    s.keys.gemini = document.getElementById("geminiKey").value.trim();
    s.keys.deepl = document.getElementById("deeplKey").value.trim();
    s.models.anthropic = document.getElementById("anthropicModel").value;
    s.models.openai = document.getElementById("openaiModel").value.trim() || "gpt-4o";
    s.models.gemini = document.getElementById("geminiModel").value.trim() || "gemini-2.0-flash";
    s.autoAlign = document.getElementById("autoAlign").checked;
    s.context = document.getElementById("aiContext").value;
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
      d.innerHTML = `<button class="rm" title="Fjern">✕</button>` +
        `<button class="btn goto" style="float:right;margin-right:6px;font-size:12px;padding:3px 8px">Gå til →</button>` +
        `<span class="word">${escapeHtml(u.word)}</span> ` +
        `<span style="color:var(--muted);font-size:13px">– ${escapeHtml(c.title)} (kap. ${k + 1})</span>` +
        `<div class="ctx">EN: ${escapeHtml(truncate(u.en, 120))}</div>` + (u.no ? `<div class="ctx">NO: ${escapeHtml(truncate(u.no, 120))}</div>` : "");
      d.querySelector(".rm").onclick = () => { store.uncertain[k].splice(idx, 1); save(); renderUncertain(); updateUncCount(); if (k === ci) renderChapter(); };
      d.querySelector(".goto").onclick = () => {     // hopp rett til ordet i teksten
        closeOverlays(); goto(k);
        active = { segId: u.segId, side: u.side, wi: u.wi, word: u.word };
        updateHighlights();
        const card = findCard(u.side, u.segId);
        if (card) card.scrollIntoView({ behavior: "smooth", block: "center" });
      };
      wrap.appendChild(d);
    });
  }
  document.getElementById("uncClose").onclick = closeOverlays;
  document.getElementById("uncCopy").onclick = () => {
    const items = allUncertain();
    if (!items.length) { alert("Ingen ord å kopiere ennå."); return; }
    let txt = aiContext() + "Jeg oversetter en bok fra engelsk til norsk bokmål. Kan du sjekke disse ordene jeg er usikker på? For hvert ord: gi norsk oversettelse og en kort begrunnelse.\n\n";
    items.forEach(({ c, k, u }, i) => {
      txt += `${i + 1}) Ord: «${u.word}»  (${c.title}, kap. ${k + 1})\n   Engelsk: "${u.en}"\n`;
      if (u.no) txt += `   Min norske tekst: "${u.no}"\n`;
      txt += "\n";
    });
    copyText(txt).then(ok => {
      const b = document.getElementById("uncCopy");
      b.textContent = ok ? "✓ Kopiert!" : "Kunne ikke kopiere";
      setTimeout(() => { b.textContent = "📋 Kopier (til ChatGPT/Claude/Gemini)"; }, 1800);
    });
  };

  // ---------- Lagre nå ----------
  const saveBtn = document.getElementById("saveBtn");
  function commitOpenEditor() {                 // ta med tekst som står i en åpen rediger-boks
    if (editingSeg == null) return;
    const ta = grid.querySelector(".editbox");
    if (!ta) return;
    const v = ta.value.trim();
    if (!store.translations[ci]) store.translations[ci] = {};
    if (v) store.translations[ci][editingSeg] = v; else delete store.translations[ci][editingSeg];
  }
  function saveNow() {
    clearTimeout(saveTimer);
    commitOpenEditor();
    if (me && store.project && hasSource()) cloudSaveNow();   // lagre i skyen samtidig
    try {
      localStorage.setItem(KEY, JSON.stringify(store)); saveWarned = false;
      saveBtn.textContent = "✓ Lagret"; saveBtn.classList.add("ok");
      setTimeout(() => { saveBtn.textContent = "💾 Lagre"; saveBtn.classList.remove("ok"); }, 1500);
    } catch (e) {
      saveBtn.textContent = "⚠ Lite plass";
      setTimeout(() => { saveBtn.textContent = "💾 Lagre"; }, 2500);
      alert("Obs: nettleseren har lite lagringsplass igjen. Ta en arbeidsfil under «≡ Fil → Lagre / del arbeidsfil» for å være helt trygg.");
    }
  }
  saveBtn.onclick = saveNow;

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

  function workFileData() {                     // alt arbeidet – aldri nøkler
    return {
      source: store.source, translations: store.translations, links: store.links,
      uncertain: store.uncertain, align: store.align,
      settings: { context: store.settings.context, autoAlign: store.settings.autoAlign },
    };
  }
  function applyWorkData(obj) {                 // forutsetter at obj.source har passert validSource
    store.source = obj.source || null;
    if (store.source) store.source.chapters.forEach(c => {
      c.title = typeof c.title === "string" ? c.title : "(uten tittel)";
      c.segments = c.segments.filter(s => s && typeof s === "object");
      c.segments.forEach((s, i) => {                                  // re-indekser + tving form defensivt
        s.id = i;
        s.en = typeof s.en === "string" ? s.en : "";
        s.type = typeof s.type === "string" ? s.type : "body";
      });
    });
    const norm = store.source ? normTables(obj, store.source.chapters.length, store.source.chapters) : { t: {}, l: {}, u: {}, a: {} };
    store.translations = norm.t; store.links = norm.l; store.uncertain = norm.u; store.align = norm.a;
    if (obj.settings && typeof obj.settings === "object") {           // kontekst følger med – aldri nøkler
      if (typeof obj.settings.context === "string") store.settings.context = obj.settings.context;
      if (typeof obj.settings.autoAlign === "boolean") store.settings.autoAlign = obj.settings.autoAlign;
    }
    resetAlignCaches();
    ci = 0; active = null; editingSeg = null;
  }
  document.getElementById("exportJson").onclick = () => {
    download("arbeidsfil-oversettelse.json", JSON.stringify(workFileData(), null, 2), "application/json");
  };
  function validSource(src) {
    if (src == null) return true;
    if (typeof src !== "object" || !Array.isArray(src.chapters) || !src.chapters.length) return false;
    return src.chapters.every(c => c && Array.isArray(c.segments));
  }
  // Rens importerte tabeller: riktig form per kapittel, og dropp kapitler utenfor dokumentet
  function isPairArr(v) { return Array.isArray(v) && v.every(p => Array.isArray(p) && p.length === 2 && Number.isInteger(p[0]) && Number.isInteger(p[1]) && p[0] >= 0 && p[1] >= 0); }
  function normTables(obj, nChapters, chapters) {
    const t = {}, l = {}, u = {}, a = {};
    const tbl = (v) => (v && typeof v === "object" && !Array.isArray(v)) ? v : {};
    const tIn = tbl(obj.translations), lIn = tbl(obj.links), uIn = tbl(obj.uncertain), aIn = tbl(obj.align);
    for (let k = 0; k < nChapters; k++) {
      const nSegs = chapters[k].segments.length;
      const tb = tIn[k]; if (tb && typeof tb === "object" && !Array.isArray(tb)) {
        const out = {}; for (const id in tb) if (typeof tb[id] === "string") out[id] = tb[id];
        if (Object.keys(out).length) t[k] = out;
      }
      for (const [src, dst] of [[lIn, l], [aIn, a]]) {
        const b = src[k]; if (b && typeof b === "object" && !Array.isArray(b)) {
          const out = {}; for (const id in b) if (isPairArr(b[id])) out[id] = b[id];
          if (Object.keys(out).length) dst[k] = out;
        }
      }
      const ub = uIn[k]; if (Array.isArray(ub)) {
        const out = ub.filter(x => x && typeof x === "object" && typeof x.word === "string" &&
          Number.isInteger(x.segId) && x.segId >= 0 && x.segId < nSegs && (x.side === "en" || x.side === "no") && Number.isInteger(x.wi));
        if (out.length) u[k] = out;
      }
    }
    return { t, l, u, a };
  }
  document.getElementById("importJson").onclick = () => importFile.click();
  importFile.onchange = (e) => {
    const f = e.target.files[0]; if (!f) return;
    const r = new FileReader();
    r.onload = () => {
      let obj;
      try { obj = JSON.parse(r.result); } catch (err) { alert("Klarte ikke å lese fila: ugyldig format."); return; }
      // Må faktisk være en arbeidsfil fra denne appen – ellers kunne en hvilken som helst
      // JSON-fil «importeres» og nullstille alt arbeid
      if (!obj || typeof obj !== "object" || !("source" in obj) || !obj.source || !validSource(obj.source)) {
        alert("Dette ser ikke ut som en arbeidsfil fra denne appen. Ingen endring gjort."); return;
      }
      const hasWork = hasSource() && Object.values(store.translations).some(t => t && Object.keys(t).length);
      if (hasWork && !confirm("Dette erstatter dokumentet og oversettelsen som ligger her med innholdet i arbeidsfila. Fortsette?")) return;
      const snapshot = JSON.stringify(store);
      try { localStorage.setItem(KEY + "-forrige", snapshot); } catch (e) {}   // angre-kopi
      try {
        applyWorkData(obj);
        store.project = null; setCloudStatus(me ? "ikke lagret i skyen ennå" : "");   // en arbeidsfil er ikke knyttet til et skyprosjekt
        renderAll(); updateUncCount(); save(); closeOverlays();
        alert("Arbeidsfil hentet inn.");
      } catch (err) {
        store = JSON.parse(snapshot); ci = 0; renderAll(); updateUncCount();
        alert("Klarte ikke å bruke arbeidsfila: " + err.message);
      }
    };
    r.readAsText(f); e.target.value = "";
  };
  document.getElementById("exportText").onclick = () => {
    if (!hasSource()) { alert("Last opp en original først."); return; }
    let out = (store.source.name || "Oversettelse") + "\n\n", missing = 0;
    store.source.chapters.forEach((c, k) => {
      out += "\n========== " + (k + 1) + ". " + c.title + " ==========\n\n";
      c.segments.forEach(s => {
        const no = (store.translations[k] && store.translations[k][s.id]) || "";
        if (!PROSE.includes(s.type) && !no) return;   // ikke-PROSE tas med kun når de faktisk er oversatt
        if (!no) missing++;
        out += (no || "[mangler oversettelse]") + "\n\n";
      });
    });
    const docName = (store.source.name || "oversettelse").replace(/\.(docx|txt|md)$/i, "");
    download(docName + " – norsk.txt", out, "text/plain");
    alert(missing ? `Eksportert. Obs: ${missing} avsnitt mangler oversettelse og er merket [mangler oversettelse] i fila.` : "Eksportert – alle avsnitt er oversatt.");
  };
  document.getElementById("clearAll").onclick = () => {
    if (!confirm("Dette tømmer originalen og hele oversettelsen fra appen (nøklene dine beholdes). Ta gjerne en arbeidsfil først. Fortsette?")) return;
    const keep = store.settings;
    store = freshStore(); store.settings = keep; store.seenHelp = true;
    setCloudStatus("");
    resetAlignCaches();
    ci = 0; active = null; editingSeg = null;
    save(); closeOverlays(); renderAll(); updateUncCount();
  };

  // ---------- Sky: innlogging, prosjekter, DeepL ----------
  let me = null, cloudAvailable = false, cloudTimer = null, cloudBusy = false, cloudDirty = false;
  const $ = (id) => document.getElementById(id);

  async function api(path, body) {
    const res = await fetch(path, body === undefined
      ? { method: "GET" }
      : { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) });
    let d = null; try { d = await res.json(); } catch (e) {}
    if (!res.ok || !d) {
      const err = new Error((d && d.error) || "Fikk ikke kontakt med tjeneren – prøv igjen.");
      err.data = d; err.status = res.status;
      throw err;
    }
    return d;
  }
  function setMe(m) {
    me = m;
    $("accountBtn").textContent = m ? "👤 " + (m.name || "").split(" ")[0] : "👤 Logg inn";
    setCloudStatus("");
  }
  function setCloudStatus(t) { $("cloudStatus").textContent = t; }
  function progressSummary() {
    if (!hasSource()) return "";
    let done = 0, tot = 0;
    store.source.chapters.forEach((c, k) => c.segments.forEach(s => {
      if (PROSE.includes(s.type)) { tot++; if (store.translations[k] && store.translations[k][s.id]) done++; }
    }));
    return done + " / " + tot + " avsnitt";
  }
  function scheduleCloudSave(markDirty) {
    if (!me || !store.project || !hasSource()) return;
    if (markDirty !== false) cloudDirty = true;
    if (!cloudDirty) return;                    // ren navigasjon skal ikke utløse skylagring
    clearTimeout(cloudTimer);
    cloudTimer = setTimeout(() => cloudSaveNow(), 4000);
  }
  async function cloudSaveNow(force) {
    if (!me || !store.project || !hasSource() || !cloudDirty) return;
    if (cloudBusy) { clearTimeout(cloudTimer); cloudTimer = setTimeout(() => cloudSaveNow(), 2000); return; }   // prøv igjen straks
    const proj = store.project;                 // fang referansen – brukeren kan bytte prosjekt mens vi venter
    cloudBusy = true; setCloudStatus("☁️ lagrer …");
    try {
      const d = await api("/api/projects", {
        action: "save", id: proj.id || "", name: proj.name, data: workFileData(),
        progress: progressSummary(), expectedUpdatedAt: proj.updatedAt || "", force: !!force,
      });
      if (store.project !== proj) return;       // prosjektet ble byttet underveis – forkast svaret
      proj.id = d.id; proj.updatedAt = d.updatedAt; cloudDirty = false;
      const kl = new Date().toLocaleTimeString("nb-NO", { hour: "2-digit", minute: "2-digit" });
      setCloudStatus("☁️ lagret " + kl);
    } catch (e) {
      if (store.project !== proj) return;
      if (e.data && e.data.conflict) {          // endret et annet sted (annen maskin/fane)
        setCloudStatus("☁️ konflikt");
        const hentNy = confirm("Dette prosjektet er endret et annet sted (kanskje en annen maskin eller fane).\n\n" +
          "Trykk OK for å hente den nyeste versjonen fra skyen (det du har gjort her etter siste lagring går tapt).\n" +
          "Trykk Avbryt for å overskrive skyen med det du har her.");
        if (hentNy) { cloudBusy = false; await reloadProjectFromCloud(proj); return; }
        cloudBusy = false; return cloudSaveNow(true);   // brukeren valgte å overskrive
      }
      setCloudStatus("☁️ får ikke lagret – prøver igjen");
      clearTimeout(cloudTimer); cloudTimer = setTimeout(() => cloudSaveNow(), 30000);
    }
    finally { cloudBusy = false; if (cloudDirty && store.project === proj) scheduleCloudSave(false); }
  }
  async function reloadProjectFromCloud(proj) {
    try {
      const d = await api("/api/projects", { action: "load", id: proj.id });
      if (!d.data || !d.data.source || !validSource(d.data.source)) throw new Error("Prosjektdataene ser skadet ut.");
      applyWorkData(d.data);
      store.project = { id: proj.id, name: d.name || proj.name, updatedAt: d.updatedAt };
      cloudDirty = false;
      renderAll(); updateUncCount(); save(false);
      setCloudStatus("☁️ hentet nyeste versjon");
    } catch (e) { setCloudStatus("☁️ " + e.message); }
  }

  // Konto-vindu
  let accMode = "login";
  function showAccount() {
    $("accLoggedOut").style.display = me ? "none" : "";
    $("accLoggedIn").style.display = me ? "" : "none";
    $("accError").textContent = ""; $("accError2").textContent = "";
    if (me) { $("accHello").textContent = "Innlogget som " + me.name; loadProjects(); }
    openOverlay("accountOverlay");
  }
  function setAccMode(m) {
    accMode = m;
    $("loginForm").style.display = m === "login" ? "" : "none";
    $("registerForm").style.display = m === "register" ? "" : "none";
    $("accSubmit").textContent = m === "login" ? "Logg inn" : "Opprett bruker";
    $("tabLogin").classList.toggle("primary", m === "login");
    $("tabRegister").classList.toggle("primary", m === "register");
  }
  $("accountBtn").onclick = showAccount;
  $("tabLogin").onclick = () => setAccMode("login");
  $("tabRegister").onclick = () => setAccMode("register");
  $("accCancel").onclick = closeOverlays;
  $("accClose").onclick = closeOverlays;
  $("accSubmit").onclick = async () => {
    const err = $("accError"); err.textContent = "";
    const btn = $("accSubmit"); if (btn.disabled) return; btn.disabled = true;
    try {
      const d = accMode === "login"
        ? await api("/api/auth", { action: "login", email: $("loginEmail").value, password: $("loginPassword").value })
        : await api("/api/auth", { action: "register", name: $("regName").value, email: $("regEmail").value, password: $("regPassword").value, invite: $("regInvite").value });
      setMe({ name: d.name, email: d.email });
      $("loginPassword").value = ""; $("regPassword").value = "";
      showAccount();
    } catch (e) { err.textContent = e.message; }
    finally { btn.disabled = false; }
  };
  $("accLogout").onclick = async () => {
    try { await api("/api/auth", { action: "logout" }); } catch (e) {}
    setMe(null); showAccount();
  };

  // Prosjektliste
  async function loadProjects() {
    const wrap = $("projList"); wrap.innerHTML = '<p class="empty-hint">Henter …</p>';
    try {
      const d = await api("/api/projects");
      wrap.innerHTML = "";
      if (!d.projects.length) { wrap.innerHTML = '<p class="empty-hint">Ingen prosjekter ennå. Trykk «☁️ Lagre dette arbeidet som prosjekt» for å komme i gang.</p>'; return; }
      d.projects.forEach(p => {
        const row = document.createElement("div"); row.className = "uncertain-item";
        const isOpen = store.project && store.project.id === p.id;
        const when = p.updatedAt ? new Date(p.updatedAt).toLocaleString("nb-NO", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) : "";
        row.innerHTML = `<button class="rm" title="Slett prosjektet">✕</button>` +
          `<button class="btn popen" style="float:right;margin-right:6px;font-size:12px;padding:3px 8px">${isOpen ? "Åpent nå" : "Åpne →"}</button>` +
          `<span class="word">${escapeHtml(p.name)}</span>` +
          `<div class="ctx">${escapeHtml(when)}${p.progress ? " · " + escapeHtml(p.progress) : ""}${isOpen ? " · dette er åpent nå" : ""}</div>`;
        row.querySelector(".popen").disabled = !!isOpen;
        row.querySelector(".popen").onclick = () => openProject(p);
        row.querySelector(".rm").onclick = async () => {
          if (!confirm(`Slette prosjektet «${p.name}» fra skyen? (Det som er åpent i appen beholdes.)`)) return;
          try {
            await api("/api/projects", { action: "delete", id: p.id });
            if (store.project && store.project.id === p.id) { store.project = null; save(); }
            loadProjects();
          } catch (e) { $("accError2").textContent = e.message; }
        };
        wrap.appendChild(row);
      });
    } catch (e) { wrap.innerHTML = ""; $("accError2").textContent = e.message; }
  }
  async function openProject(p) {
    if (hasSource() && (!store.project || store.project.id !== p.id)) {
      if (!confirm(`Åpne «${p.name}»? Det som står i appen nå erstattes (ta gjerne en arbeidsfil først hvis det ikke er lagret som prosjekt).`)) return;
    }
    try {
      // send eventuelle usendte endringer i gjeldende prosjekt FØR vi bytter
      clearTimeout(cloudTimer);
      if (me && store.project && cloudDirty && hasSource() && store.project.id !== p.id) {
        setCloudStatus("☁️ lagrer det gamle først …");
        await cloudSaveNow();
      }
      const d = await api("/api/projects", { action: "load", id: p.id });
      if (!d.data || !d.data.source || !validSource(d.data.source)) throw new Error("Prosjektdataene ser skadet ut.");
      try { localStorage.setItem(KEY + "-forrige", JSON.stringify(store)); } catch (e) {}
      applyWorkData(d.data);
      store.project = { id: p.id, name: d.name || p.name, updatedAt: d.updatedAt };
      cloudDirty = false;
      renderAll(); updateUncCount(); save(false); closeOverlays();
      setCloudStatus("☁️ åpnet «" + truncate(store.project.name, 24) + "»");
    } catch (e) { $("accError2").textContent = e.message; }
  }
  $("projSaveAs").onclick = async () => {
    if (!hasSource()) { $("accError2").textContent = "Last opp en engelsk original først – så kan arbeidet lagres som prosjekt."; return; }
    const def = (store.project && store.project.name) || (store.source.name || "Min bok").replace(/\.(docx|txt|md)$/i, "");
    const raw = prompt("Hva skal prosjektet hete?", def);
    if (raw === null) return;
    const name = raw.trim() || def;
    if (store.project && store.project.id) {
      if (name !== store.project.name) {
        // samme prosjekt med nytt navn, eller en ny kopi?
        const giNyttNavn = confirm(`Trykk OK for å gi prosjektet nytt navn («${name}»).\nTrykk Avbryt for å lagre som et NYTT prosjekt i tillegg.`);
        if (giNyttNavn) {
          try { await api("/api/projects", { action: "rename", id: store.project.id, name }); store.project.name = name; }
          catch (e) { $("accError2").textContent = e.message; return; }
        } else {
          store.project = { id: "", name };     // bevisst ny kopi
        }
      }
    } else {
      store.project = { id: "", name };
    }
    cloudDirty = true;
    await cloudSaveNow();
    save(false); loadProjects();
  };

  // DeepL-oversettelse av kapittelet (via tjeneren – DeepL blokkerer direktekall fra nettleser)
  let deeplBusy = false;
  $("optDeepl").onclick = async () => {
    const status = $("optDeeplStatus");
    if (!me) { status.textContent = "Logg inn først (👤 øverst til høyre) for å bruke DeepL."; return; }
    if (deeplBusy) return;
    deeplBusy = true; $("optDeepl").disabled = true;
    status.textContent = "🌐 Oversetter med DeepL …";
    const myChapter = ci;
    try {
      const texts = chapterBodies().map(s => s.en);
      const d = await api("/api/deepl", { texts, deeplKey: store.settings.keys.deepl || "" });
      if (ci !== myChapter) { status.textContent = "Kapittelet ble byttet mens DeepL jobbet – gå tilbake og prøv igjen."; return; }
      optPaste.value = (d.translations || []).join("\n\n"); updateOptCount();
      status.textContent = "Oversatt! Se gjennom og trykk «Sett inn på rad».";
    } catch (e) { status.textContent = e.message; }
    finally { deeplBusy = false; $("optDeepl").disabled = false; }
  };

  async function initCloud(attempt) {
    if (!/^https?:$/.test(location.protocol)) return;
    try {
      const d = await api("/api/auth", { action: "me" });
      cloudAvailable = true;
      $("accountBtn").style.display = "";
      if (d.ok) setMe({ name: d.name, email: d.email });
      if (me && store.project && hasSource()) {
        setCloudStatus("☁️ «" + truncate(store.project.name, 24) + "»");
        checkCloudFreshness();                 // er det en nyere versjon i skyen (annen maskin)?
      }
    } catch (e) {
      if (e.status === 404 || e.status === undefined && !(e instanceof TypeError)) {
        // ser ut som ingen tjener (åpnet lokalt) – skjul sky-funksjonene stille
        if (e.status === 404) return;
      }
      const n = (attempt || 0) + 1;
      if (n <= 2) { setTimeout(() => initCloud(n), n * 3000); return; }   // nettglipp – prøv igjen
      if (store.project && hasSource()) setCloudStatus("☁️ får ikke kontakt – lagres bare lokalt");
    }
  }
  async function checkCloudFreshness() {
    try {
      const d = await api("/api/projects");
      const p = d.projects.find(x => x.id === store.project.id);
      if (p && store.project.updatedAt && p.updatedAt && p.updatedAt > store.project.updatedAt) {
        if (confirm(`Prosjektet «${p.name}» er endret et annet sted siden sist (kanskje en annen maskin).\n\nHente den nyeste versjonen fra skyen?`)) {
          await reloadProjectFromCloud(store.project);
        }
      }
    } catch (e) { /* ikke kritisk */ }
  }

  // ---------- Hjelpere ----------
  function openOverlay(id) { closeOverlays(); document.getElementById(id).classList.add("show"); }
  function closeOverlays() { document.querySelectorAll(".overlay.show").forEach(o => o.classList.remove("show")); }
  // Klikk på mørk bakgrunn lukker – men ikke oppslag/innstillinger/konto (så svar/nøkler/passord ikke mistes ved feilklikk)
  document.querySelectorAll(".overlay").forEach(o => { if (o.id === "lookupOverlay" || o.id === "settingsOverlay" || o.id === "accountOverlay") return; o.addEventListener("click", (e) => { if (e.target === o) closeOverlays(); }); });
  // 👁-knapper: vis/skjul API-nøkler så man kan sjekke at limingen ble riktig
  document.querySelectorAll(".prov-block input[type=password]").forEach(inp => {
    const row = document.createElement("div"); row.className = "keyrow";
    inp.parentNode.insertBefore(row, inp); row.appendChild(inp);
    const eye = document.createElement("button");
    eye.type = "button"; eye.textContent = "👁"; eye.className = "btn eye"; eye.title = "Vis/skjul nøkkelen";
    eye.onclick = () => { inp.type = inp.type === "password" ? "text" : "password"; };
    row.appendChild(eye);
  });
  document.addEventListener("keydown", (e) => {
    if ((e.metaKey || e.ctrlKey) && (e.key === "s" || e.key === "S")) { e.preventDefault(); saveNow(); return; }
    if (e.key === "Escape") { closeOverlays(); if (active) { active = null; updateHighlights(); } }
    if (editingSeg !== null || !hasSource()) return;
    if (document.querySelector(".overlay.show")) return;   // ikke bla kapitler mens et vindu er åpent
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
    try { commitOpenEditor(); localStorage.setItem(KEY, JSON.stringify(store)); } catch (e) {}
    // siste skylagrings-forsøk – virker kun for små prosjekter (nettlesere avviser store keepalive-kropper),
    // hovedvernet er visibilitychange-lagringen under
    if (me && store.project && cloudDirty && hasSource()) {
      try {
        fetch("/api/projects", {
          method: "POST", keepalive: true, headers: { "content-type": "application/json" },
          body: JSON.stringify({ action: "save", id: store.project.id || "", name: store.project.name, data: workFileData(), progress: progressSummary(), expectedUpdatedAt: store.project.updatedAt || "" }),
        });
      } catch (e) {}
    }
  });
  // Når fanen skjules (bytter fane, minimerer, lukker snart): lagre til skyen med vanlig fetch mens siden ennå lever
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden" && me && store.project && cloudDirty && hasSource()) {
      commitOpenEditor(); clearTimeout(cloudTimer); cloudSaveNow();
    }
  });

  // Enkel skjermleser-merking på modalene
  document.querySelectorAll(".overlay .modal").forEach(m => { m.setAttribute("role", "dialog"); m.setAttribute("aria-modal", "true"); });

  // ---------- Start ----------
  try { renderAll(); }
  catch (err) {                              // skadet lagret tilstand skal ikke låse appen
    try { localStorage.setItem(KEY + "-skadet", localStorage.getItem(KEY) || ""); } catch (e) {}  // ta vare på en kopi
    store.source = null; save();
    try { renderAll(); } catch (e2) {}
  }
  updateUncCount();
  setAccMode("login");
  initCloud();
  if (!store.seenHelp) { openOverlay("helpOverlay"); store.seenHelp = true; save(); }
})();

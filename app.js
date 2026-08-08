/* =========================================================================
   CoreWeave Cutsheet Lookup — application logic
   Pure client-side. Loads an Excel/CSV cutsheet with SheetJS, lets the user
   map columns, then looks up all connections that reference a given
   sector:rack:RU location and totals the optic transceivers needed.
   ========================================================================= */
(function () {
  "use strict";

  // ---- App state --------------------------------------------------------
  const state = {
    fileName: "",
    workbook: null,
    sheetNames: [],
    activeSheet: null,
    headers: [],      // array of column header strings
    rows: [],         // array of objects keyed by header
    mapping: {
      endpoints: [],  // [{ mode:'combined'|'split', combined:idx, sector:idx, rack:idx, ru:idx, name }]
      optics: [],     // [{ pn:idx, qty:idx|null, desc:idx|null, endpoint:idx|null, name }]
    },
    opticScope: "endpoint", // 'endpoint' | 'all'
    lastResult: null,
  };

  const $ = (id) => document.getElementById(id);
  const el = (tag, cls, txt) => {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (txt != null) n.textContent = txt;
    return n;
  };

  // =======================================================================
  // Location parsing / normalisation
  // A location is (sector, rack, RU). We build a canonical signature so that
  // "05" == "5", "S1:05:30" == "s1-05-30", etc.  Query supports partial
  // matches: "s1:05" matches every RU in that rack.
  // =======================================================================
  function parseLocation(raw) {
    if (raw == null) return null;
    const str = String(raw).trim().toLowerCase();
    if (!str) return null;

    const tokens = str.split(/[^a-z0-9]+/).filter(Boolean);
    let sector = null, rack = null, ru = null;
    const loose = [];

    for (const t of tokens) {
      let m;
      if (sector == null && (m = /^s(\d+)$/.exec(t))) { sector = parseInt(m[1], 10); }
      else if (rack == null && (m = /^r(\d+)$/.exec(t))) { rack = parseInt(m[1], 10); }
      else if (ru == null && (m = /^(?:ru|u)(\d+)$/.exec(t))) { ru = parseInt(m[1], 10); }
      else if (/^\d+$/.test(t)) { loose.push(parseInt(t, 10)); }
      // non-numeric tokens (device names etc.) are ignored
    }
    // Positionally fill remaining components from leftover bare numbers.
    for (const n of loose) {
      if (sector == null) sector = n;
      else if (rack == null) rack = n;
      else if (ru == null) ru = n;
    }
    if (sector == null && rack == null && ru == null) return null;
    return { sector, rack, ru };
  }

  // Build a location directly from separate sector/rack/RU cell values.
  function locFromParts(s, r, u) {
    const num = (v) => {
      if (v == null || String(v).trim() === "") return null;
      const m = /(\d+)/.exec(String(v));
      return m ? parseInt(m[1], 10) : null;
    };
    const sector = num(s), rack = num(r), ru = num(u);
    if (sector == null && rack == null && ru == null) return null;
    return { sector, rack, ru };
  }

  function locLabel(loc) {
    if (!loc) return "—";
    const p = (v) => (v == null ? "?" : v);
    return `S${p(loc.sector)}:${String(p(loc.rack)).padStart(2, "0")}:${p(loc.ru)}`;
  }

  // Does an endpoint location satisfy the (possibly partial) query location?
  // Only the components present in the query must match.
  function locMatches(query, endpoint) {
    if (!query || !endpoint) return false;
    if (query.sector != null && query.sector !== endpoint.sector) return false;
    if (query.rack != null && query.rack !== endpoint.rack) return false;
    if (query.ru != null && query.ru !== endpoint.ru) return false;
    // Require at least the components the user gave to have been comparable.
    return true;
  }

  // =======================================================================
  // File loading
  // =======================================================================
  function loadArrayBuffer(name, buf) {
    let wb;
    try {
      wb = XLSX.read(buf, { type: "array", cellDates: true });
    } catch (err) {
      showResultBanner("danger", "Could not read that file: " + err.message);
      return;
    }
    state.fileName = name;
    state.workbook = wb;
    state.sheetNames = wb.SheetNames.slice();
    state.activeSheet = wb.SheetNames[0];

    // Populate UI
    $("fiName").textContent = name;
    $("fileInfo").classList.add("show");
    $("dropzone").style.display = "none";

    const sel = $("sheetSelect");
    sel.innerHTML = "";
    state.sheetNames.forEach((n) => {
      const o = el("option", null, n);
      o.value = n;
      sel.appendChild(o);
    });
    $("sheetPickWrap").style.display = state.sheetNames.length > 1 ? "" : "none";

    selectSheet(state.activeSheet);
    setStatus(true);
  }

  function selectSheet(name) {
    state.activeSheet = name;
    const ws = state.workbook.Sheets[name];
    // header:1 gives array-of-arrays; we find the header row heuristically.
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
    const { headerRow, headers } = detectHeaderRow(aoa);
    state.headers = headers;
    state.rows = aoa.slice(headerRow + 1).map((arr) => {
      const obj = {};
      headers.forEach((h, i) => { obj[i] = arr[i] != null ? arr[i] : ""; });
      return obj;
    }).filter((obj) => Object.values(obj).some((v) => String(v).trim() !== ""));

    $("fiMeta").textContent =
      `${state.rows.length} rows · ${headers.length} columns` +
      (state.sheetNames.length > 1 ? ` · sheet “${name}”` : "");

    autoDetectMapping();
    $("mapCard").classList.remove("hidden");
    $("searchCard").classList.remove("hidden");
    renderMapping();
  }

  // Pick the row that looks most like a header: the first row (within the
  // first few) whose cells are mostly non-empty short text strings.
  function detectHeaderRow(aoa) {
    const limit = Math.min(aoa.length, 15);
    let best = 0, bestScore = -1;
    for (let r = 0; r < limit; r++) {
      const row = aoa[r] || [];
      const nonEmpty = row.filter((c) => String(c).trim() !== "");
      if (nonEmpty.length < 2) continue;
      let textish = 0;
      for (const c of nonEmpty) {
        const s = String(c).trim();
        if (s.length <= 40 && /[a-z]/i.test(s)) textish++;
      }
      const score = textish + nonEmpty.length * 0.25 - r * 0.5;
      if (score > bestScore) { bestScore = score; best = r; }
    }
    const rawHeaders = (aoa[best] || []).map((c, i) => {
      const s = String(c).trim();
      return s || `Column ${i + 1}`;
    });
    return { headerRow: best, headers: rawHeaders };
  }

  // =======================================================================
  // Auto-detect column mapping from header names
  // =======================================================================
  function autoDetectMapping() {
    const H = state.headers.map((h) => h.toLowerCase());
    const endpoints = [];
    const optics = [];

    const has = (h, ...words) => words.some((w) => h.includes(w));
    const isLoc = (h) => has(h, "location", "loc", "endpoint", "position");
    const isSector = (h) => /\bsector\b/.test(h) || /\bsec\b/.test(h) || /(^|[^a-z])sect/.test(h);
    const isRack = (h) => /\brack\b/.test(h) || /\bcab(inet)?\b/.test(h);
    const isRU = (h) => /\bru\b/.test(h) || /\bu\b/.test(h) || has(h, "elevation", "rack unit", "elev");

    // Column "senses" used to classify optic columns.
    const opticish = (h) => has(h, "optic", "transceiver", "xcvr", "sfp", "qsfp", "osfp");
    const partish = (h) => has(h, "part number", "part no", "part #", "p/n", "partnum") || /\bpn\b/.test(h) || /\bpart\b/.test(h) || h.endsWith(" pn") || h.endsWith("pn");
    const descish = (h) => has(h, "description", "desc", "notes", "comment");
    const typeish = (h) => has(h, "type", "model", "media");
    // A real optic-PN column mentions optics/parts but is NOT the description.
    const isOptic = (h) => (opticish(h) || partish(h)) && !descish(h);
    const isPartCol = (h) => partish(h) && !descish(h);
    const isDesc = (h) => descish(h);
    const isQty = (h) => has(h, "qty", "quantity", "count");

    // ---- Endpoints (combined location columns) ----
    let combinedLocs = [];
    H.forEach((h, i) => { if (isLoc(h)) combinedLocs.push(i); });
    if (combinedLocs.length) {
      combinedLocs.forEach((i) =>
        endpoints.push({ mode: "combined", combined: i, sector: null, rack: null, ru: null, name: sideName(state.headers[i]), side: sideToken(state.headers[i]) })
      );
    } else {
      // ---- Endpoints (split sector/rack/RU columns), grouped by A/Z side ----
      const sectorCols = [], rackCols = [], ruCols = [];
      H.forEach((h, i) => {
        if (isSector(h)) sectorCols.push(i);
        else if (isRack(h)) rackCols.push(i);
        else if (isRU(h)) ruCols.push(i);
      });
      const n = Math.max(rackCols.length, ruCols.length, sectorCols.length);
      for (let k = 0; k < n; k++) {
        const rack = rackCols[k] != null ? rackCols[k] : null;
        const ru = ruCols[k] != null ? ruCols[k] : null;
        const sector = sectorCols[k] != null ? sectorCols[k] : (sectorCols[0] != null ? sectorCols[0] : null);
        if (rack == null && ru == null) continue;
        const refIdx = rack != null ? rack : ru;
        endpoints.push({ mode: "split", combined: null, sector, rack, ru, name: sideName(state.headers[refIdx]), side: sideToken(state.headers[refIdx]) });
      }
    }

    // ---- Optic PN columns (tiered: optic-labelled first, then generic PN) ----
    let opticCols = [];
    H.forEach((h, i) => { if (opticish(h) && !descish(h)) opticCols.push(i); });
    if (!opticCols.length) H.forEach((h, i) => { if (isPartCol(h)) opticCols.push(i); });

    const qtyCol = H.findIndex(isQty);
    opticCols.forEach((i) => {
      const nm = state.headers[i];
      const desc = findSiblingDesc(i, H, isDesc);
      // Optics are tied to a *side* (a/z/1/2). "At searched location" then counts
      // an optic when the query matched any endpoint on that same side — robust
      // even when a side has several location columns (locode, patch-panel, …).
      optics.push({ pn: i, qty: qtyCol >= 0 ? qtyCol : null, desc, endpoint: null, side: sideToken(nm), name: nm });
    });

    state.mapping.endpoints = endpoints;
    state.mapping.optics = optics;
  }

  function sideName(header) {
    return header || "Endpoint";
  }

  // Find a description column that belongs to the same side as an optic col.
  function findSiblingDesc(opticIdx, H, isDesc) {
    const side = sideToken(state.headers[opticIdx]);
    let generic = -1;
    for (let i = 0; i < H.length; i++) {
      if (!isDesc(H[i])) continue;
      if (side && sideToken(state.headers[i]) === side) return i;
      if (generic < 0) generic = i;
    }
    return generic >= 0 ? generic : null;
  }

  function sideToken(header) {
    const h = (header || "").toLowerCase();
    if (/\b(a[- ]?end|a[- ]?side|\ba\b|src|source|near|local)\b/.test(h)) return "a";
    if (/\b(z[- ]?end|z[- ]?side|b[- ]?end|\bz\b|\bb\b|dst|dest|destination|far|remote)\b/.test(h)) return "z";
    if (/\b1\b/.test(h)) return "1";
    if (/\b2\b/.test(h)) return "2";
    return null;
  }

  // Which endpoint indices does this optic belong to?
  //  - explicit override (op.endpoint) wins;
  //  - otherwise every endpoint sharing the optic's side (a/z/1/2);
  //  - null means "can't tell" (caller counts it for the whole connection).
  function opticEndpointIdxs(op, endpoints) {
    if (op.endpoint != null) return [op.endpoint];
    if (op.side) {
      const idxs = [];
      endpoints.forEach((ep, i) => { if (ep.side && ep.side === op.side) idxs.push(i); });
      if (idxs.length) return idxs;
    }
    return null;
  }

  // =======================================================================
  // Mapping UI
  // =======================================================================
  function colOptions(selected, allowNone, noneLabel) {
    const frag = document.createDocumentFragment();
    if (allowNone) {
      const o = el("option", null, noneLabel || "— none —");
      o.value = "";
      if (selected == null || selected === "") o.selected = true;
      frag.appendChild(o);
    }
    state.headers.forEach((h, i) => {
      const o = el("option", null, h);
      o.value = String(i);
      if (String(selected) === String(i)) o.selected = true;
      frag.appendChild(o);
    });
    return frag;
  }

  function renderMapping() {
    renderEndpoints();
    renderOptics();
    updateMapBanner();
  }

  function renderEndpoints() {
    const c = $("endpointsContainer");
    c.innerHTML = "";
    state.mapping.endpoints.forEach((ep, idx) => {
      const block = el("div", "endpoint-block");

      const head = el("div", "eb-head");
      head.appendChild(el("span", null, "◍ Endpoint " + (idx + 1)));
      const toggle = el("div", "mode-toggle");
      ["combined", "split"].forEach((m) => {
        const b = el("button", ep.mode === m ? "active" : null, m === "combined" ? "One column" : "Sector/Rack/RU");
        b.onclick = () => { ep.mode = m; renderMapping(); };
        toggle.appendChild(b);
      });
      head.appendChild(toggle);
      const rm = el("button", "btn ghost sm", "✕");
      rm.style.marginLeft = "8px";
      rm.title = "Remove endpoint";
      rm.onclick = () => { state.mapping.endpoints.splice(idx, 1); renderMapping(); };
      head.appendChild(rm);
      block.appendChild(head);

      if (ep.mode === "combined") {
        const l = el("label", "fld");
        l.style.margin = "0";
        l.appendChild(mkLabel("Location column"));
        const s = document.createElement("select");
        s.appendChild(colOptions(ep.combined, false));
        s.onchange = () => { ep.combined = parseInt(s.value, 10); ep.name = state.headers[ep.combined]; ep.side = sideToken(ep.name); };
        l.appendChild(s);
        block.appendChild(l);
      } else {
        const grid = el("div", "grid-3");
        [["sector", "Sector column"], ["rack", "Rack column"], ["ru", "RU / elevation column"]].forEach(([key, label]) => {
          const l = el("label", "fld");
          l.style.margin = "0";
          l.appendChild(mkLabel(label));
          const s = document.createElement("select");
          s.appendChild(colOptions(ep[key], true, "— n/a —"));
          s.onchange = () => { ep[key] = s.value === "" ? null : parseInt(s.value, 10); };
          l.appendChild(s);
          grid.appendChild(l);
        });
        block.appendChild(grid);
      }
      c.appendChild(block);
    });
    if (!state.mapping.endpoints.length) {
      c.appendChild(el("div", "mini-hint", "No endpoint columns set — add at least one so the tool knows where to look."));
    }
  }

  function renderOptics() {
    const c = $("opticsContainer");
    c.innerHTML = "";
    const epNames = state.mapping.endpoints.map((e, i) => `Endpoint ${i + 1} (${e.name || "?"})`);
    state.mapping.optics.forEach((op, idx) => {
      const row = el("div", "slot-row");

      // PN column
      const pnL = el("label", "fld"); pnL.style.margin = "0";
      pnL.appendChild(mkLabel("Optic PN column"));
      const pnS = document.createElement("select");
      pnS.appendChild(colOptions(op.pn, false));
      pnS.onchange = () => { op.pn = parseInt(pnS.value, 10); op.name = state.headers[op.pn]; op.side = sideToken(op.name); };
      pnL.appendChild(pnS); row.appendChild(pnL);

      // Qty column
      const qL = el("label", "fld"); qL.style.margin = "0";
      qL.appendChild(mkLabel("Quantity col (opt.)"));
      const qS = document.createElement("select");
      qS.appendChild(colOptions(op.qty, true, "— 1 per row —"));
      qS.onchange = () => { op.qty = qS.value === "" ? null : parseInt(qS.value, 10); };
      qL.appendChild(qS); row.appendChild(qL);

      // Description column
      const dL = el("label", "fld"); dL.style.margin = "0";
      dL.appendChild(mkLabel("Description col (opt.)"));
      const dS = document.createElement("select");
      dS.appendChild(colOptions(op.desc, true));
      dS.onchange = () => { op.desc = dS.value === "" ? null : parseInt(dS.value, 10); };
      dL.appendChild(dS); row.appendChild(dL);

      // remove
      const rm = el("button", "btn ghost sm", "✕");
      rm.title = "Remove optic column";
      rm.onclick = () => { state.mapping.optics.splice(idx, 1); renderMapping(); };
      row.appendChild(rm);

      c.appendChild(row);

      // endpoint association (for "at searched location" scope)
      if (epNames.length > 1) {
        const assoc = el("label", "fld");
        assoc.style.margin = "0 0 14px";
        assoc.appendChild(mkLabel("This optic sits at"));
        const aS = document.createElement("select");
        const none = el("option", null, "— auto (by side) —"); none.value = "";
        if (op.endpoint == null) none.selected = true;
        aS.appendChild(none);
        epNames.forEach((n, i) => {
          const o = el("option", null, n); o.value = String(i);
          if (op.endpoint === i) o.selected = true;
          aS.appendChild(o);
        });
        aS.onchange = () => { op.endpoint = aS.value === "" ? null : parseInt(aS.value, 10); };
        assoc.appendChild(aS);
        c.appendChild(assoc);
      }
    });
    if (!state.mapping.optics.length) {
      c.appendChild(el("div", "mini-hint", "No optic columns set — add the column(s) that hold transceiver part numbers."));
    }
  }

  function mkLabel(text) { return el("span", "lbl", text); }

  function updateMapBanner() {
    const b = $("mapBanner");
    const problems = [];
    if (!state.mapping.endpoints.length) problems.push("no endpoint (location) column is mapped");
    if (!state.mapping.optics.length) problems.push("no optic part-number column is mapped");
    const hasAssoc = state.mapping.optics.some((o) => o.endpoint != null);
    if (state.mapping.endpoints.length > 1 && !hasAssoc) {
      // informational, not an error
    }
    if (problems.length) {
      b.className = "banner warn show";
      b.textContent = "Heads up: " + problems.join("; ") + ". Search will still run but results may be incomplete.";
      $("mapSub").textContent = "needs attention";
    } else {
      b.className = "banner";
      const eNames = state.mapping.endpoints.map((e) => e.name).join(", ");
      const oNames = state.mapping.optics.map((o) => o.name).join(", ");
      $("mapSub").textContent = `endpoints: ${eNames}  ·  optics: ${oNames}`;
    }
  }

  // =======================================================================
  // Endpoint location extraction per row
  // =======================================================================
  function endpointLoc(row, ep) {
    if (ep.mode === "combined") {
      return ep.combined != null ? parseLocation(row[ep.combined]) : null;
    }
    return locFromParts(
      ep.sector != null ? row[ep.sector] : null,
      ep.rack != null ? row[ep.rack] : null,
      ep.ru != null ? row[ep.ru] : null
    );
  }

  // =======================================================================
  // Search
  // =======================================================================
  function runSearch() {
    const raw = $("queryInput").value.trim();
    if (!raw) return;
    if (!state.rows.length) { showResultBanner("warn", "Load a cutsheet first."); return; }

    const queries = raw.split(/[,;\n]+/).map((s) => s.trim()).filter(Boolean)
      .map((q) => ({ text: q, loc: parseLocation(q) }))
      .filter((q) => q.loc);

    if (!queries.length) {
      showResultBanner("warn", `Couldn't read a location from “${raw}”. Use the format sector:rack:RU, e.g. s1:05:30.`);
      return;
    }

    const matchedRows = matchRowsPure(state.rows, state.mapping.endpoints, queries);
    state.lastResult = { queries, matchedRows };
    renderResults();
  }

  // Pure: find rows whose mapped endpoints satisfy any query. No DOM/state.
  function matchRowsPure(rows, endpoints, queries) {
    const matchedRows = [];
    rows.forEach((row, ri) => {
      const eps = endpoints.map((ep) => endpointLoc(row, ep));
      const matchedEndpoints = [];
      let matchedQuery = null;
      eps.forEach((loc, ei) => {
        for (const q of queries) {
          if (locMatches(q.loc, loc)) {
            matchedEndpoints.push(ei);
            if (!matchedQuery) matchedQuery = q.text;
            break;
          }
        }
      });
      if (matchedEndpoints.length) {
        matchedRows.push({ ri, row, eps, matchedEndpoints, matchedQuery });
      }
    });
    return matchedRows;
  }

  function aggregateOptics(matchedRows, scope) {
    return aggregatePure(matchedRows, state.mapping.optics, state.mapping.endpoints, scope);
  }

  // Pure: total optics across matched rows. In "endpoint" scope, an optic is
  // counted only when it sits on a side that matched. No DOM/state.
  function aggregatePure(matchedRows, optics, endpoints, scope) {
    const map = new Map(); // pn -> { pn, desc, qty }
    let unassignedNote = false;

    for (const mr of matchedRows) {
      for (const op of optics) {
        if (op.pn == null) continue;
        // Scope filter: only count optics that sit at a matched endpoint.
        if (scope === "endpoint") {
          const idxs = opticEndpointIdxs(op, endpoints);
          if (idxs == null) {
            // Can't tell which side this optic is on; count it but flag.
            if (endpoints.length > 1) unassignedNote = true;
          } else if (!idxs.some((i) => mr.matchedEndpoints.includes(i))) {
            continue; // optic is on a side that didn't match — skip
          }
        }
        const pnVal = String(mr.row[op.pn] || "").trim();
        if (!pnVal) continue;
        let qty = 1;
        if (op.qty != null) {
          const q = parseFloat(String(mr.row[op.qty]).replace(/[^0-9.\-]/g, ""));
          qty = isNaN(q) ? 1 : q;
        }
        const desc = op.desc != null ? String(mr.row[op.desc] || "").trim() : "";
        const key = pnVal.toUpperCase();
        if (!map.has(key)) map.set(key, { pn: pnVal, desc, qty: 0 });
        const rec = map.get(key);
        rec.qty += qty;
        if (!rec.desc && desc) rec.desc = desc;
      }
    }
    const list = Array.from(map.values()).sort((a, b) => b.qty - a.qty || a.pn.localeCompare(b.pn));
    return { list, unassignedNote };
  }

  function renderResults() {
    const { matchedRows } = state.lastResult;
    $("results").classList.add("show");
    $("resultBanner").className = "banner";

    if (!matchedRows.length) {
      $("statConns").textContent = "0";
      $("statOptics").textContent = "0";
      $("statPNs").textContent = "0";
      $("opticsBody").innerHTML = "";
      $("connBody").innerHTML = "";
      $("connHead").innerHTML = "";
      $("opticsCount").textContent = "0";
      $("connCount").textContent = "0";
      showResultBanner("warn", "No connections in the cutsheet reference that location. Double-check the value, or the column mapping in step 2.");
      return;
    }

    const { list, unassignedNote } = aggregateOptics(matchedRows, state.opticScope);
    const totalOptics = list.reduce((s, r) => s + r.qty, 0);

    $("statConns").textContent = matchedRows.length;
    $("statOptics").textContent = totalOptics % 1 === 0 ? totalOptics : totalOptics.toFixed(2);
    $("statPNs").textContent = list.length;

    // Optics table
    const anyDesc = list.some((r) => r.desc);
    $("descHead").style.display = anyDesc ? "" : "none";
    const ob = $("opticsBody");
    ob.innerHTML = "";
    list.forEach((r) => {
      const tr = el("tr");
      const td1 = el("td", "pn");
      td1.appendChild(el("span", "pn-badge", r.pn));
      tr.appendChild(td1);
      if (anyDesc) tr.appendChild(el("td", null, r.desc || "—"));
      else { const t = el("td"); t.style.display = "none"; tr.appendChild(t); }
      tr.appendChild(el("td", "qty", r.qty % 1 === 0 ? String(r.qty) : r.qty.toFixed(2)));
      ob.appendChild(tr);
    });
    $("opticsCount").textContent = list.length;

    if (unassignedNote) {
      showResultBanner("warn", "Some optic columns aren’t linked to a specific endpoint, so they’re counted for the whole connection. Set “This optic sits at” in step 2 for precise per-location counts.");
    }

    // Connections table — show mapped-relevant columns plus a match badge.
    renderConnTable(matchedRows);
    $("connCount").textContent = matchedRows.length;

    document.getElementById("results").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function relevantColumns() {
    const cols = new Set();
    state.mapping.endpoints.forEach((ep) => {
      if (ep.mode === "combined" && ep.combined != null) cols.add(ep.combined);
      else { ["sector", "rack", "ru"].forEach((k) => { if (ep[k] != null) cols.add(ep[k]); }); }
    });
    state.mapping.optics.forEach((op) => {
      if (op.pn != null) cols.add(op.pn);
      if (op.qty != null) cols.add(op.qty);
      if (op.desc != null) cols.add(op.desc);
    });
    // Include a leading identifier column if present (e.g. Cable ID) for context.
    if (state.headers.length) cols.add(0);
    return Array.from(cols).sort((a, b) => a - b);
  }

  function renderConnTable(matchedRows) {
    const cols = relevantColumns();
    const head = $("connHead");
    head.innerHTML = "";
    head.appendChild(el("th", null, "Match"));
    cols.forEach((c) => head.appendChild(el("th", null, state.headers[c])));

    const body = $("connBody");
    body.innerHTML = "";
    matchedRows.forEach((mr) => {
      const tr = el("tr");
      const badgeTd = el("td");
      mr.matchedEndpoints.forEach((ei) => {
        const loc = mr.eps[ei];
        const b = el("span", "match-badge", locLabel(loc));
        b.style.marginRight = "4px";
        badgeTd.appendChild(b);
      });
      tr.appendChild(badgeTd);
      cols.forEach((c) => {
        const isPN = state.mapping.optics.some((o) => o.pn === c);
        const td = el("td", isPN ? "pn" : null, String(mr.row[c] != null ? mr.row[c] : ""));
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
  }

  // =======================================================================
  // Export helpers
  // =======================================================================
  function toCSV(rows) {
    return rows.map((r) => r.map((c) => {
      const s = String(c == null ? "" : c);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(",")).join("\n");
  }
  function download(filename, text) {
    const blob = new Blob([text], { type: "text/csv;charset=utf-8;" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  function queryTag() {
    return (state.lastResult && state.lastResult.queries.map((q) => q.text).join("_").replace(/[^a-z0-9_]+/gi, "-")) || "results";
  }

  function exportOpticsCSV() {
    if (!state.lastResult) return;
    const { list } = aggregateOptics(state.lastResult.matchedRows, state.opticScope);
    const rows = [["Optic Part Number", "Description", "Quantity"]];
    list.forEach((r) => rows.push([r.pn, r.desc, r.qty]));
    download(`optics_${queryTag()}.csv`, toCSV(rows));
  }
  function copyOptics() {
    if (!state.lastResult) return;
    const { list } = aggregateOptics(state.lastResult.matchedRows, state.opticScope);
    const text = list.map((r) => `${r.pn}\t${r.qty}${r.desc ? "\t" + r.desc : ""}`).join("\n");
    navigator.clipboard.writeText(text).then(
      () => flash($("copyOpticsBtn"), "Copied!"),
      () => flash($("copyOpticsBtn"), "Copy failed")
    );
  }
  function exportConnsCSV() {
    if (!state.lastResult) return;
    const cols = relevantColumns();
    const rows = [["Matched location"].concat(cols.map((c) => state.headers[c]))];
    state.lastResult.matchedRows.forEach((mr) => {
      const badge = mr.matchedEndpoints.map((ei) => locLabel(mr.eps[ei])).join(" ");
      rows.push([badge].concat(cols.map((c) => mr.row[c])));
    });
    download(`connections_${queryTag()}.csv`, toCSV(rows));
  }
  function flash(btn, msg) {
    const old = btn.textContent; btn.textContent = msg;
    setTimeout(() => { btn.textContent = old; }, 1400);
  }

  // =======================================================================
  // UI plumbing
  // =======================================================================
  function setStatus(loaded) {
    const chip = $("statusChip");
    if (loaded) {
      chip.classList.add("loaded");
      $("statusText").textContent = state.fileName;
      $("loadSub").textContent = state.fileName;
    } else {
      chip.classList.remove("loaded");
      $("statusText").textContent = "No cutsheet loaded";
      $("loadSub").textContent = "";
    }
  }
  function showResultBanner(kind, msg) {
    const b = $("resultBanner");
    b.className = "banner " + kind + " show";
    b.textContent = msg;
    $("results").classList.add("show");
  }

  function clearFile() {
    state.workbook = null; state.rows = []; state.headers = []; state.fileName = "";
    state.mapping = { endpoints: [], optics: [] };
    state.lastResult = null;
    $("fileInfo").classList.remove("show");
    $("dropzone").style.display = "";
    $("mapCard").classList.add("hidden");
    $("searchCard").classList.add("hidden");
    $("results").classList.remove("show");
    $("fileInput").value = "";
    setStatus(false);
  }

  function initTheme() {
    const saved = localStorage.getItem("cw_theme");
    if (saved) document.documentElement.setAttribute("data-theme", saved);
    $("themeBtn").onclick = () => {
      const cur = document.documentElement.getAttribute("data-theme");
      const isDark = cur === "dark" || (!cur && window.matchMedia("(prefers-color-scheme: dark)").matches);
      const next = isDark ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", next);
      localStorage.setItem("cw_theme", next);
    };
  }

  function wireEvents() {
    // File input / dropzone
    const dz = $("dropzone");
    dz.onclick = () => $("fileInput").click();
    $("fileInput").onchange = (e) => { const f = e.target.files[0]; if (f) readFile(f); };
    ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
    dz.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) readFile(f); });

    $("clearFileBtn").onclick = clearFile;
    $("sheetSelect").onchange = (e) => selectSheet(e.target.value);

    // Card collapse
    document.querySelectorAll("[data-toggle]").forEach((h) => {
      h.addEventListener("click", () => h.closest(".card").classList.toggle("collapsed"));
    });

    // Mapping buttons
    $("addEndpointBtn").onclick = () => {
      state.mapping.endpoints.push({ mode: "combined", combined: 0, sector: null, rack: null, ru: null, name: state.headers[0], side: sideToken(state.headers[0]) });
      renderMapping();
    };
    $("addOpticBtn").onclick = () => {
      state.mapping.optics.push({ pn: 0, qty: null, desc: null, endpoint: null, side: sideToken(state.headers[0]), name: state.headers[0] });
      renderMapping();
    };
    $("remapBtn").onclick = () => { autoDetectMapping(); renderMapping(); };

    // Search
    $("searchBtn").onclick = runSearch;
    $("queryInput").addEventListener("keydown", (e) => { if (e.key === "Enter") runSearch(); });
    document.querySelectorAll(".chip[data-ex]").forEach((c) => {
      c.onclick = () => { $("queryInput").value = c.getAttribute("data-ex"); runSearch(); };
    });

    // Optic scope toggle
    $("opticScopeToggle").querySelectorAll("button").forEach((b) => {
      b.onclick = () => {
        state.opticScope = b.getAttribute("data-scope");
        $("opticScopeToggle").querySelectorAll("button").forEach((x) => x.classList.toggle("active", x === b));
        if (state.lastResult) renderResults();
      };
    });

    // Exports
    $("copyOpticsBtn").onclick = copyOptics;
    $("csvOpticsBtn").onclick = exportOpticsCSV;
    $("csvConnsBtn").onclick = exportConnsCSV;
  }

  function readFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => loadArrayBuffer(file.name, new Uint8Array(e.target.result));
    reader.onerror = () => showResultBanner("danger", "Could not read the file.");
    reader.readAsArrayBuffer(file);
  }

  // ---- boot ----
  document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    wireEvents();
    setStatus(false);
  });

  // Expose a tiny surface for the headless test harness (Node/JSDOM).
  if (typeof module !== "undefined" && module.exports) {
    module.exports = {
      parseLocation, locFromParts, locMatches, locLabel,
      sideToken, endpointLoc, opticEndpointIdxs, matchRowsPure, aggregatePure,
    };
  }
})();

/* Tap Dashboard: live view of the Draft Survey SharePoint lists, after Microsoft 365 sign-in. */
(function () {
  "use strict";
  const CFG = {
    clientId: "77d8f819-d373-40ee-a424-da8a8b15552d",
    tenantId: "0c07cc35-4442-46a2-a2b1-a250d653b92f",
    site: "febdistributing365.sharepoint.com:/sites/DraftSurvey",
    lists: { visits: "cd3d31a7-ae92-4d2f-bb94-ca0693882ebc", stops: "4696936e-1e85-49ae-b5c7-50fc01f1237d", brands: "fd9848ac-b296-440b-ae80-0e28f10ab755", cleaning: "3dc80bc5-6a2e-45c3-b4e7-6ae6ed159aa0" },
    scopes: ["Sites.Read.All"]
  };
  const NICE = { "CONSTELLATION BEER": "Constellation", "ANHEUSER-BUSCH": "Anheuser-Busch", "HEINEKEN USA": "Heineken USA", "MOLSON COORS": "Molson Coors", "PARISH BREWING": "Parish Brewing", "FLY LLAMA BREWING COMPANY": "Fly Llama Brewing", "CHANDELEUR BREWING": "Chandeleur Brewing", "ABITA BREWING CO": "Abita", "BOSTON BEER CO.": "Boston Beer", "THE YUENGLING CO LLC": "Yuengling", "SOUTHERN PROHIBITION BREWING": "Southern Prohibition", "PABST BREWING CO": "Pabst", "GAMBRINUS CO.": "Gambrinus (Shiner)", "DIAGEO-GUINNESS": "Diageo-Guinness", "SAPPORO USA": "Sapporo USA", "NEW BELGIUM": "New Belgium", "LAGUNITAS": "Lagunitas", "SWEETWATER BREWING": "SweetWater", "LAZY MAGNOLIA": "Lazy Magnolia", "URBAN SOUTH BREWERY": "Urban South", "FOUNDERS": "Founders", "SIERRA NEVADA": "Sierra Nevada", "JOCK BEER LLC": "Jock Beer (Garage Beer)", "FERTILE GROUND": "Fertile Ground", "KEY CITY BREWING": "Key City Brewing", "WISEACRE BREWING": "Wiseacre", "KONA BREWING": "Kona", "ACE CIDER": "Ace Cider", "GNARLY BARLEY BREWING": "Gnarly Barley", "PERFECT PLAIN BREWING": "Perfect Plain", "YELLOWHAMMER BREWING": "Yellowhammer Brewing", "KIRIN": "Kirin" };
  const FIX = { "ANHEUSER-BUSCH INBEV": "Anheuser-Busch", "MITCHELL DISTRIBUTING": "Distributor (brand not recorded)", "SOUTHERN BEVERAGES": "Distributor (brand not recorded)", "TRULY HARD SELTZER": "Boston Beer", "BELL'S SEASONAL": "Bell's", "BELL'S TWO-HEARTED ALE": "Bell's", "BELL'S SPECIALTY": "Bell's", "NOLA SEASONAL": "NOLA Brewing" };
  const $ = s => document.querySelector(s);
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const fmt = n => n.toLocaleString(), pct = (a, b) => b ? Math.round(100 * a / b) + "%" : "–";
  const iso = d => d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  const titleCase = s => String(s || "").toLowerCase().replace(/\b\w/g, c => c.toUpperCase());

  /* ---------- sign-in ---------- */
  let msalApp = null, account = null;
  async function token() {
    const req = { scopes: CFG.scopes, account };
    try { return (await msalApp.acquireTokenSilent(req)).accessToken; }
    catch (e) { await msalApp.acquireTokenRedirect(req); throw new Error("redirecting"); }
  }
  async function graph(path) {
    const t = await token();
    let url = path.startsWith("https://") ? path : "https://graph.microsoft.com/v1.0" + path, out = [];
    while (url) {
      const r = await fetch(url, { headers: { Authorization: "Bearer " + t, Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } });
      if (r.status === 401 || r.status === 403) { const e = new Error("noaccess"); e.status = r.status; throw e; }
      if (!r.ok) throw new Error("Graph " + r.status);
      const j = await r.json();
      if (!j.value) return j;
      out = out.concat(j.value); url = j["@odata.nextLink"] || null;
    }
    return out;
  }

  /* ---------- data ---------- */
  let VISITS = [], STOPS = {}, BRANDSUP = {}, CLEAN = [];
  async function loadAll() {
    status("Loading survey data…");
    const site = await graph("/sites/" + CFG.site + "?$select=id");
    const L = id => "/sites/" + site.id + "/lists/" + id + "/items?$top=999&$expand=fields($select=";
    const [v, s, b, cl] = await Promise.all([
      graph(L(CFG.lists.visits) + "Title,AccountNumber,AccountName,VisitDate,Location,Market,LineCleanerName,TapList,DataSource)"),
      graph(L(CFG.lists.stops) + "Title,CustomerName,RouteNum,LineCleaner,Location,Phase,DayOfWeek,Sequence)"),
      graph(L(CFG.lists.brands) + "Title,Supplier)"),
      graph(L(CFG.lists.cleaning) + "Title,AccountNumber,AccountName,VisitDate,Market,Location,LineCleanerName,CleanQ1,CleanQ2,SignatureStatus,DataSource)")
    ]);
    VISITS = v.map(x => x.fields); CLEAN = cl.map(x => x.fields);
    STOPS = {}; s.forEach(x => { STOPS[x.fields.Title] = x.fields; });
    BRANDSUP = {}; b.forEach(x => { BRANDSUP[String(x.fields.Title || "").toUpperCase()] = x.fields.Supplier || ""; });
  }
  function supOf(brand, given) {
    const B = String(brand || "").toUpperCase();
    if (FIX[B]) return FIX[B];
    const raw = BRANDSUP[B] || given || "";
    if (!raw) return "Supplier not in brand list";
    return NICE[raw.toUpperCase()] || titleCase(raw);
  }

  /* ---------- filters + state ---------- */
  const today = new Date();
  const st = { loc: "All", route: "", from: "", to: iso(today), range: "30", src: "all", brand: null, sup: null, excl: false, q: "", sortK: "s", sortD: 1 };
  function setRange(r) {
    st.range = r;
    if (r === "all") { st.from = ""; st.to = iso(today); }
    else if (r !== "custom") { const d = new Date(); d.setDate(d.getDate() - (+r - 1)); st.from = iso(d); st.to = iso(today); }
  }
  setRange("30");
  const OUR = () => st.loc === "Gulfport" ? "FEB" : st.loc === "Jackson" ? "CCB" : "FEB/CCB";

  // One row per account: its most recent visit in the chosen range. Rows: [acct, name, date, loc, brand, ours, comp, route, cleaner, supplier]
  function rows() {
    const latest = {};
    VISITS.forEach(v => {
      const d = String(v.VisitDate || "").slice(0, 10); if (!d) return;
      if (st.from && d < st.from) return; if (st.to && d > st.to) return;
      if (st.src === "app" && v.DataSource === "Encompass import") return;
      if (st.src === "import" && v.DataSource !== "Encompass import") return;
      const a = v.AccountNumber; if (!latest[a] || d > latest[a].d) latest[a] = { d, v };
    });
    const R = [];
    Object.values(latest).forEach(({ d, v }) => {
      const stop = STOPS[v.AccountNumber] || {};
      const loc = v.Market || v.Location || stop.Location || "";
      if (st.loc !== "All" && loc !== st.loc) return;
      if (st.route && String(stop.RouteNum || "") !== st.route) return;
      let taps = []; try { taps = JSON.parse(v.TapList || "[]"); } catch (e) {}
      const by = {};
      taps.forEach(t => { if (t.status === "Pending") return; const k = String(t.brand || "").toUpperCase(); const x = by[k] || (by[k] = { o: 0, c: 0, s: t.supplier }); t.ours ? x.o++ : x.c++; });
      Object.entries(by).forEach(([b, x]) => R.push([v.AccountNumber, titleCase(v.AccountName || stop.CustomerName), d, loc, b, x.o, x.c, stop.RouteNum || "", v.LineCleanerName || stop.LineCleaner || "", supOf(b, x.s)]));
    });
    return R;
  }

  /* ---------- render ---------- */
  function accounts(rs) { const m = {}; rs.forEach(r => { const a = m[r[0]] || (m[r[0]] = { id: r[0], n: r[1], l: r[3], d: r[2], route: r[7], cl: r[8], o: 0, c: 0, taps: [] }); a.o += r[5]; a.c += r[6]; a.taps.push(r); }); return Object.values(m).map(a => (a.t = a.o + a.c, a.s = a.t ? a.o / a.t : 0, a)); }
  function bars(el, items, max, onClick, selKey) {
    el.innerHTML = items.map(([k, o, c]) => `<div class="bar ${selKey === k ? "sel" : ""}" data-k="${esc(k)}"><div class="lab" title="${esc(k)}">${esc(k)}</div><div class="track"><div class="o" style="width:${100 * o / max}%"></div><div class="c" style="width:${100 * c / max}%"></div></div><div class="n">${fmt(o + c)}</div></div>`).join("");
    if (onClick) el.querySelectorAll(".bar").forEach(b => b.onclick = () => onClick(b.dataset.k));
  }
  function render() {
    const rs = rows(); const all = accounts(rs);
    let acc = all;
    if (st.brand) acc = acc.filter(a => a.taps.some(r => r[4] === st.brand));
    if (st.sup) acc = acc.filter(a => a.taps.some(r => r[9] === st.sup));
    if (st.q) { const Q = st.q.toUpperCase(); acc = acc.filter(a => a.n.toUpperCase().includes(Q) || String(a.id).includes(Q) || a.taps.some(r => r[4].includes(Q))); }
    const preEx = acc; if (st.excl) acc = acc.filter(a => a.c === 0);
    const ids = new Set(acc.map(a => a.id)); const fr = rs.filter(r => ids.has(r[0]));
    const O = fr.reduce((s, r) => s + r[5], 0), C = fr.reduce((s, r) => s + r[6], 0);
    const withComp = acc.filter(a => a.c > 0).length, exAll = preEx.filter(a => a.c === 0).length;
    document.querySelectorAll(".our").forEach(e => e.textContent = OUR());
    const chips = [];
    if (st.excl) chips.push(["Exclusive accounts", () => { st.excl = false; }]);
    if (st.sup) chips.push([st.sup, () => { st.sup = null; }]);
    if (st.brand) chips.push([st.brand, () => { st.brand = null; }]);
    $("#active").innerHTML = chips.map(([t], i) => `<span class="chip">${esc(t)}<button data-i="${i}" aria-label="Clear ${esc(t)} filter">×</button></span>`).join("");
    $("#active").querySelectorAll("button").forEach(b => b.onclick = () => { chips[+b.dataset.i][1](); render(); });
    $("#kpis").innerHTML = [[fmt(acc.length), "Accounts surveyed"], [fmt(O + C), "Taps counted"], [fmt(O), OUR() + " taps"], [fmt(C), "Competitor taps"], [pct(O, O + C), OUR() + " tap share"], [pct(withComp, acc.length), "Accounts with a competitor tap"]].map(([b, s]) => `<div class="kpi"><b>${b}</b><span>${s}</span></div>`).join("")
      + `<button class="kpi ex ${st.excl ? "on" : ""}" id="exbtn" aria-pressed="${st.excl}"><b>${fmt(exAll)} <small>${pct(exAll, preEx.length)}</small></b><span>${OUR()} exclusive accounts · ${st.excl ? "showing only these" : "click to show"}</span></button>`;
    $("#exbtn").onclick = () => { st.excl = !st.excl; render(); };
    const L = {}; fr.forEach(r => { const x = L[r[3]] || (L[r[3]] = [r[3] + (r[3] === "Gulfport" ? " (FEB)" : r[3] === "Jackson" ? " (CCB)" : ""), 0, 0]); x[1] += r[5]; x[2] += r[6]; });

    const base = st.excl ? fr : rs, cb = {}, ob = {};
    base.forEach(r => { if (r[6]) cb[r[4]] = (cb[r[4]] || 0) + r[6]; if (r[5]) ob[r[4]] = (ob[r[4]] || 0) + r[5]; });
    const pick = k => { st.brand = st.brand === k ? null : k; render(); };
    const ci = Object.entries(cb).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([k, v]) => [k, 0, v]); bars($("#comp"), ci, Math.max(1, ...ci.map(x => x[2])), pick, st.brand);
    if (!ci.length) $("#comp").innerHTML = '<p class="muted">No competitor taps in these accounts.</p>';
    const oi = Object.entries(ob).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([k, v]) => [k, v, 0]); bars($("#ours"), oi, Math.max(1, ...oi.map(x => x[1])), pick, st.brand);
    const SP = {}; (st.excl || st.brand || st.q ? fr : rs).forEach(r => { const x = SP[r[9]] || (SP[r[9]] = [r[9], 0, 0]); x[1] += r[5]; x[2] += r[6]; });
    const si = Object.values(SP).sort((a, b) => (b[1] + b[2]) - (a[1] + a[2])); bars($("#sup"), si, Math.max(1, ...si.map(x => x[1] + x[2])), k => { st.sup = st.sup === k ? null : k; render(); }, st.sup);
    acc.sort((a, b) => { const x = a[st.sortK], y = b[st.sortK]; return (typeof x === "string" ? String(x).localeCompare(String(y)) : x - y) * st.sortD; });
    $("#rows").innerHTML = acc.map(a => `<tr data-id="${esc(a.id)}"><td>${esc(a.n)} <span class="muted">#${esc(a.id)}</span></td><td>${esc(a.l)}</td><td>${esc(a.route)}</td><td class="num">${a.t}</td><td class="num">${a.o}</td><td class="num">${a.c}</td><td class="num">${pct(a.o, a.t)}</td></tr>`).join("") || `<tr><td colspan="7" class="muted">No surveyed accounts match these filters.</td></tr>`;
    $("#detail").innerHTML = "";
    document.querySelectorAll("#rows tr[data-id]").forEach(tr => tr.onclick = () => {
      const a = acc.find(x => String(x.id) === tr.dataset.id);
      $("#detail").innerHTML = `<h2 style="margin-top:14px">${esc(a.n)} · ${esc(a.l)}${a.route ? " · Route " + esc(a.route) : ""} · surveyed ${esc(a.d)}${a.cl ? " by " + esc(a.cl) : ""}</h2>` + a.taps.map(r => `${r[5] ? `<span class="pill o">${esc(r[4])}${r[5] > 1 ? " ×" + r[5] : ""}</span>` : ""}${r[6] ? `<span class="pill c">${esc(r[4])}${r[6] > 1 ? " ×" + r[6] : ""}</span>` : ""}`).join("");
    });
    const srcCount = new Set(VISITS.map(v => v.DataSource === "Encompass import" ? "import" : "app"));
    renderCleaning();
    renderChanges();
    renderSchedule();
    $("#subtitle").textContent = (st.from ? st.from + " to " + st.to : "All dates through " + st.to) + " · latest visit per account" + (srcCount.has("import") && st.src !== "app" ? " · includes Encompass import" : "");
  }
  /* ---------- line cleaning ---------- */
  const EVERY = ph => { const p = String(ph || ""); return /all/i.test(p) ? 7 : /&/.test(p) ? 14 : 28; };
  const daysBetween = (a, b) => Math.round((Date.parse(b + "T12:00:00Z") - Date.parse(a + "T12:00:00Z")) / 864e5);
  function renderCleaning() {
    const okStop = s => (st.loc === "All" || s.Location === st.loc) && (!st.route || String(s.RouteNum || "") === st.route);
    const inScope = c => { const s = STOPS[c.AccountNumber] || {}; const loc = c.Market || c.Location || s.Location || "";
      if (st.loc !== "All" && loc !== st.loc) return false; if (st.route && String(s.RouteNum || "") !== st.route) return false;
      if (st.src === "app" && c.DataSource === "Encompass import") return false; if (st.src === "import" && c.DataSource !== "Encompass import") return false; return true; };
    const all = CLEAN.filter(inScope);
    const inRange = all.filter(c => { const d = String(c.VisitDate || "").slice(0, 10); return d && (!st.from || d >= st.from) && (!st.to || d <= st.to); });
    const n = inRange.length, acid = inRange.filter(c => /acid/i.test(c.CleanQ1 || "")).length, pump = inRange.filter(c => /yes/i.test(c.CleanQ2 || "")).length;
    const answered = q => inRange.filter(c => String(c[q] || "").trim()).length;
    const appRows = inRange.filter(c => c.DataSource !== "Encompass import"), signed = appRows.filter(c => c.SignatureStatus === "Signed").length;
    const accts = new Set(inRange.map(c => c.AccountNumber)).size;
    $("#ckpis").innerHTML = [[fmt(n), "Cleanings"], [fmt(accts), "Accounts cleaned"], [pct(acid, answered("CleanQ1")), "Acid cleanings"], [pct(pump, answered("CleanQ2")), "Pump used"], [appRows.length ? pct(signed, appRows.length) : "–", "Manager signed (app visits)"]].map(([b, s]) => `<div class="kpi"><b>${b}</b><span>${s}</span></div>`).join("");
    const by = {}; inRange.forEach(c => { const k = c.LineCleanerName || "Unknown"; const x = by[k] || (by[k] = [k, 0, 0]); /acid/i.test(c.CleanQ1 || "") ? x[2]++ : x[1]++; });
    const bi = Object.values(by).sort((a, b) => (b[1] + b[2]) - (a[1] + a[2])); bars($("#bycleaner"), bi, Math.max(1, ...bi.map(x => x[1] + x[2])));
    const last = {}; all.forEach(c => { const d = String(c.VisitDate || "").slice(0, 10); if (d && (!last[c.AccountNumber] || d > last[c.AccountNumber])) last[c.AccountNumber] = d; });
    const t = iso(today);
    const firstClean = CLEAN.reduce((m, c) => { const d = String(c.VisitDate || "").slice(0, 10); return d && (!m || d < m) ? d : m; }, "");
    // Not visited (no tap survey or cleaning) in the last 14 days.
    const lastAny = {}; const bump = (k, d) => { if (d && (!lastAny[k] || d > lastAny[k])) lastAny[k] = d; };
    VISITS.forEach(v => bump(v.AccountNumber, String(v.VisitDate || "").slice(0, 10))); CLEAN.forEach(c => bump(c.AccountNumber, String(c.VisitDate || "").slice(0, 10)));
    const hideMonthly = $("#nvmonthly").checked, hideUpcoming = $("#nvweek").checked;
    // Stops still coming up this week: their rotation runs this week and their day is today or later.
    const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
    const todayIdx = (today.getDay() + 6) % 7;
    const phaseNow = (Math.floor((Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) - Date.UTC(2024, 11, 9)) / (7 * 864e5)) % 4 + 4) % 4 + 1;
    const upcoming = s => { const ph = String(s.Phase || ""); const runs = /all/i.test(ph) || ph.includes(String(phaseNow)); const di = DAYS.indexOf(String(s.DayOfWeek || "").trim()); return runs && di >= todayIdx; };
    const nv = Object.values(STOPS).filter(okStop).filter(s => !(hideMonthly && EVERY(s.Phase) === 28)).filter(s => !(hideUpcoming && upcoming(s))).map(s => ({ s, l: lastAny[s.Title] || null, ago: lastAny[s.Title] ? daysBetween(lastAny[s.Title], t) : null }))
      .filter(x => x.ago == null || x.ago > 14).sort((a, b) => (b.ago == null ? 1e9 : b.ago) - (a.ago == null ? 1e9 : a.ago));
    const sched = Object.values(STOPS).filter(okStop).filter(s => !(hideMonthly && EVERY(s.Phase) === 28)).length;
    const upCount = Object.values(STOPS).filter(okStop).filter(upcoming).length;
    $("#nvweeknote").textContent = `(week ${phaseNow} of the rotation · ${fmt(upCount)} stops still to come)`;
    $("#nvcount").textContent = fmt(nv.length) + " of " + fmt(sched);
    $("#notvisited").innerHTML = nv.slice(0, 500).map(x => `<tr><td>${esc(titleCase(x.s.CustomerName))} <span class="muted">#${esc(x.s.Title)}</span></td><td>${esc(x.s.Location || "")}</td><td>${esc(x.s.RouteNum || "")}${x.s.LineCleaner ? " · " + esc(x.s.LineCleaner) : ""}</td><td>${esc(x.s.DayOfWeek || "")}</td><td>every ${EVERY(x.s.Phase)} days</td><td class="num">${x.l ? esc(x.l) : "No record"}</td><td class="num">${x.ago == null ? "–" : x.ago}</td></tr>`).join("") || `<tr><td colspan="7" class="muted">Every scheduled account has been visited in the last 14 days.</td></tr>`;
    $("#nvnote").textContent = firstClean ? `Cleaning records start ${firstClean}, so "No record" can also mean the account simply hasn't been cleaned since then.` : "";
  }

  /* ---------- tap changes: each visit vs the previous visit at the same account ---------- */
  let chBrand = null;
  function tapCounts(v) { let t = []; try { t = JSON.parse(v.TapList || "[]"); } catch (e) {} const m = {}; t.forEach(x => { if (x.status === "Pending") return; const k = String(x.brand || "").toUpperCase(); if (!k) return; const y = m[k] || (m[k] = { n: 0, ours: !!x.ours }); y.n++; }); return m; }
  function renderChanges() {
    const byAcct = {};
    VISITS.forEach(v => { if (st.src === "app" && v.DataSource === "Encompass import") return; if (st.src === "import" && v.DataSource !== "Encompass import") return;
      const d = String(v.VisitDate || "").slice(0, 10); if (d) (byAcct[v.AccountNumber] = byAcct[v.AccountNumber] || []).push({ d, v }); });
    const changes = [];
    Object.entries(byAcct).forEach(([acct, list]) => {
      const s = STOPS[acct] || {}; list.sort((x, y) => x.d < y.d ? -1 : 1);
      for (let i = 1; i < list.length; i++) {
        const cur = list[i], prev = list[i - 1]; if (st.from && cur.d < st.from) continue; if (st.to && cur.d > st.to) continue;
        const loc = cur.v.Market || cur.v.Location || s.Location || "";
        if (st.loc !== "All" && loc !== st.loc) continue; if (st.route && String(s.RouteNum || "") !== st.route) continue;
        const A = tapCounts(prev.v), B = tapCounts(cur.v), added = [], removed = [];
        new Set([...Object.keys(A), ...Object.keys(B)]).forEach(k => { const a = A[k] ? A[k].n : 0, b = B[k] ? B[k].n : 0; const ours = (B[k] || A[k]).ours;
          if (b > a) added.push({ k, n: b - a, ours }); if (a > b) removed.push({ k, n: a - b, ours }); });
        if (added.length || removed.length) changes.push({ acct, name: titleCase(cur.v.AccountName || s.CustomerName), loc, route: s.RouteNum || "", d: cur.d, prevD: prev.d, added, removed });
      }
    });
    const sum = (arr, f) => arr.reduce((t, c) => t + f(c), 0), cnt = (l, o) => l.filter(x => x.ours === o).reduce((t, x) => t + x.n, 0);
    const oA = sum(changes, c => cnt(c.added, true)), oR = sum(changes, c => cnt(c.removed, true)), cA = sum(changes, c => cnt(c.added, false)), cR = sum(changes, c => cnt(c.removed, false));
    const swing = n => (n > 0 ? "+" : "") + fmt(n);
    $("#chkpis").innerHTML = [[fmt(changes.length), "Accounts with tap changes"], [fmt(oA), OUR() + " taps added"], [fmt(oR), OUR() + " taps removed"], [swing(oA - oR), "Net " + OUR() + " taps"], [fmt(cA), "Competitor taps added"], [fmt(cR), "Competitor taps removed"], [swing(cA - cR), "Net competitor taps"]]
      .map(([b, s], i) => `<div class="kpi ${i === 3 ? (oA - oR >= 0 ? "up" : "down") : i === 6 ? (cA - cR > 0 ? "down" : "up") : ""}"><b>${b}</b><span>${s}</span></div>`).join("");
    const B = {}; changes.forEach(c => { c.added.forEach(x => { const y = B[x.k] || (B[x.k] = { k: x.k, ours: x.ours, a: 0, r: 0 }); y.a += x.n; }); c.removed.forEach(x => { const y = B[x.k] || (B[x.k] = { k: x.k, ours: x.ours, a: 0, r: 0 }); y.r += x.n; }); });
    const bl = Object.values(B).sort((x, y) => Math.abs(y.a - y.r) - Math.abs(x.a - x.r) || (y.a + y.r) - (x.a + x.r));
    $("#chbrands").innerHTML = bl.length ? bl.slice(0, 40).map(b => `<tr data-k="${esc(b.k)}" class="${chBrand === b.k ? "selrow" : ""}"><td>${esc(b.k)}</td><td>${b.ours ? `<span class="pill o">${OUR()}</span>` : '<span class="pill c">Competitor</span>'}</td><td class="num">${b.a}</td><td class="num">${b.r}</td><td class="num"><b class="${b.a - b.r > 0 ? (b.ours ? "gain" : "loss") : b.a - b.r < 0 ? (b.ours ? "loss" : "gain") : ""}">${swing(b.a - b.r)}</b></td></tr>`).join("") : "";
    $("#chbrands").querySelectorAll("tr").forEach(tr => tr.onclick = () => { chBrand = chBrand === tr.dataset.k ? null : tr.dataset.k; renderChanges(); });
    let list = chBrand ? changes.filter(c => c.added.some(x => x.k === chBrand) || c.removed.some(x => x.k === chBrand)) : changes;
    list = list.sort((x, y) => x.d < y.d ? 1 : -1);
    const pills = (l, cls) => l.map(x => `<span class="pill ${x.ours ? "o" : "c"}">${cls}${esc(x.k)}${x.n > 1 ? " ×" + x.n : ""}</span>`).join("");
    $("#chlist").innerHTML = list.slice(0, 300).map(c => `<tr><td class="num">${esc(c.d)}</td><td>${esc(c.name)} <span class="muted">#${esc(c.acct)}</span></td><td>${esc(c.loc)}${c.route ? " · " + esc(c.route) : ""}</td><td>${pills(c.added, "+ ")}</td><td>${pills(c.removed, "− ")}</td><td class="num muted">${esc(c.prevD)}</td></tr>`).join("");
    $("#chfilter").innerHTML = chBrand ? `<span class="chip">${esc(chBrand)}<button aria-label="Clear brand">×</button></span>` : "";
    if (chBrand) $("#chfilter button").onclick = () => { chBrand = null; renderChanges(); };
    const multi = Object.values(byAcct).filter(l => l.length > 1).length;
    $("#chempty").hidden = changes.length > 0;
    $("#chempty").textContent = multi ? "No tap changes between visits in this date range." : "No changes to show yet. Changes appear once an account has been surveyed at least twice — the Encompass import has one survey per account, so this fills in as app visits come in.";
    $("#chbody").hidden = changes.length === 0;
  }

  /* ---------- this week's stops for one route ---------- */
  let schedRoute = "";
  function renderSchedule() {
    const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
    const todayIdx = (today.getDay() + 6) % 7;
    const phaseNow = (Math.floor((Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) - Date.UTC(2024, 11, 9)) / (7 * 864e5)) % 4 + 4) % 4 + 1;
    const runs = s => { const ph = String(s.Phase || ""); return /all/i.test(ph) || ph.includes(String(phaseNow)); };
    const route = schedRoute || st.route || $("#schedRoute").value;
    if ($("#schedRoute").value !== route) $("#schedRoute").value = route;
    const lastV = {}, lastC = {}; const up = (m, k, d, extra) => { if (d && (!m[k] || d > m[k].d)) m[k] = Object.assign({ d }, extra || {}); };
    VISITS.forEach(v => up(lastV, v.AccountNumber, String(v.VisitDate || "").slice(0, 10)));
    CLEAN.forEach(c => up(lastC, c.AccountNumber, String(c.VisitDate || "").slice(0, 10), { type: c.CleanQ1 || "" }));
    const t = iso(today); let html = "", total = 0;
    for (let i = todayIdx; i < 7; i++) {
      const day = new Date(today); day.setDate(today.getDate() + (i - todayIdx));
      const list = Object.values(STOPS).filter(s => String(s.RouteNum || "") === route && String(s.DayOfWeek || "").trim() === DAYS[i] && runs(s))
        .sort((x, y) => (Number(x.Sequence) || 999) - (Number(y.Sequence) || 999) || String(x.CustomerName).localeCompare(String(y.CustomerName)));
      if (!list.length) continue; total += list.length;
      const done = list.filter(s => (lastV[s.Title] && lastV[s.Title].d === t) || (lastC[s.Title] && lastC[s.Title].d === t)).length;
      html += `<details ${i === todayIdx || !html ? "open" : ""}><summary class="dayhead">${DAYS[i]}, ${day.getMonth() + 1}/${day.getDate()}${i === todayIdx ? ' <span class="pill o">Today</span>' : ""} <span class="muted">· ${list.length} stops${i === todayIdx ? " · " + done + " done" : ""}</span></summary>
        <div class="tablewrap" style="max-height:none"><table><thead><tr><th class="num">#</th><th>Account</th><th>Location</th><th>Schedule</th><th class="num">Last tap survey</th><th class="num">Last cleaning</th><th></th></tr></thead><tbody>` +
        list.map((s, n) => { const v = lastV[s.Title], c = lastC[s.Title]; const isDone = i === todayIdx && ((v && v.d === t) || (c && c.d === t));
          return `<tr><td class="num">${esc(s.Sequence || n + 1)}</td><td>${esc(titleCase(s.CustomerName))} <span class="muted">#${esc(s.Title)}</span></td><td>${esc(s.Location || "")}</td><td>every ${EVERY(s.Phase)} days</td><td class="num">${v ? esc(v.d) : '<span class="muted">none</span>'}</td><td class="num">${c ? esc(c.d) + (c.type ? " · " + esc(c.type) : "") : '<span class="muted">none</span>'}</td><td>${isDone ? '<span class="pill o">Done</span>' : ""}</td></tr>`; }).join("") + `</tbody></table></div></details>`;
    }
    const cl = (Object.values(STOPS).find(s => String(s.RouteNum || "") === route) || {}).LineCleaner || "";
    $("#schedsum").textContent = route ? `Route ${route}${cl ? " · " + cl : ""} · week ${phaseNow} of the rotation · ${total} stops left this week` : "";
    $("#sched").innerHTML = html || `<p class="muted">No more stops scheduled on this route this week.</p>`;
  }

  function fillRoutes() {
    const rs = [...new Set(Object.values(STOPS).map(s => String(s.RouteNum || "")).filter(Boolean))].sort();
    const names = {}; Object.values(STOPS).forEach(s => { if (s.RouteNum && s.LineCleaner) names[s.RouteNum] = s.LineCleaner; });
    const opts = rs.map(r => `<option value="${esc(r)}">${esc(r)}${names[r] ? " · " + esc(names[r]) : ""}</option>`).join("");
    $("#schedRoute").innerHTML = opts; if (!schedRoute && rs.length) schedRoute = rs[0];
    $("#route").innerHTML = `<option value="">All routes</option>` + rs.map(r => `<option value="${esc(r)}">${esc(r)}${names[r] ? " · " + esc(names[r]) : ""}</option>`).join("");
  }
  function wire() {
    document.querySelectorAll("#loc button").forEach(b => b.onclick = () => { st.loc = b.dataset.v; document.querySelectorAll("#loc button").forEach(x => x.setAttribute("aria-pressed", x === b)); render(); });
    $("#route").onchange = e => { st.route = e.target.value; if (st.route) schedRoute = st.route; render(); };
    $("#range").onchange = e => { setRange(e.target.value); $("#custom").hidden = e.target.value !== "custom"; $("#from").value = st.from; $("#to").value = st.to; render(); };
    $("#from").onchange = e => { st.from = e.target.value; render(); };
    $("#to").onchange = e => { st.to = e.target.value || iso(today); render(); };
    $("#src").onchange = e => { st.src = e.target.value; render(); };
    $("#q").oninput = e => { st.q = e.target.value.trim(); render(); };
    document.querySelectorAll("th[data-k]").forEach(th => th.onclick = () => { const k = th.dataset.k; st.sortD = st.sortK === k ? -st.sortD : (k === "s" ? 1 : -1); st.sortK = k; render(); });
    $("#refresh").onclick = async () => { await start(true); };
    $("#nvmonthly").onchange = () => render();
    $("#nvweek").onchange = () => render();
    $("#schedRoute").onchange = e => { schedRoute = e.target.value; renderSchedule(); };
    $("#signout").onclick = () => msalApp.logoutRedirect({ account });
  }
  function status(msg, isErr) { $("#status").hidden = !msg; $("#status").className = "status" + (isErr ? " err" : ""); $("#status").innerHTML = msg || ""; }

  async function start(reload) {
    try {
      await loadAll(); fillRoutes(); if (!reload) wire();
      $("#main").hidden = false; status(""); render();
      $("#who").textContent = account ? account.name || account.username : "";
    } catch (e) {
      if (e.message === "redirecting") return;
      if (e.message === "noaccess") return status("You're signed in, but you don't have access to the Draft Survey data. Ask Matthew Hessell to add you.", true);
      status("Couldn't load the data (" + esc(e.message) + "). Try Refresh in a minute.", true);
    }
  }

  async function boot() {
    if (window.__DEMO) { account = { name: "Demo" }; VISITS = window.__DEMO.visits; STOPS = window.__DEMO.stops; BRANDSUP = window.__DEMO.brands; CLEAN = window.__DEMO.clean || []; fillRoutes(); wire(); $("#signin").hidden = true; $("#main").hidden = false; render(); return; }
    if (!window.msal) return status("The Microsoft sign-in library didn't load. Check your connection and reload.", true);
    msalApp = new msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: location.origin + "/" }, cache: { cacheLocation: "localStorage" } });
    const res = await msalApp.handleRedirectPromise().catch(e => { status("Sign-in didn't finish: " + esc(e.errorMessage || e.message), true); return null; });
    account = (res && res.account) || msalApp.getAllAccounts()[0] || null;
    if (!account) { $("#signin").hidden = false; $("#signbtn").onclick = () => msalApp.loginRedirect({ scopes: CFG.scopes }); return; }
    msalApp.setActiveAccount(account); $("#signin").hidden = true;
    start(false);
  }
  boot();
})();

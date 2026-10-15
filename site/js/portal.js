// Staff portal: live lots from the Excel file on OneDrive, sales and holds written straight into it, and the inventory agent.
(() => {
  const CFG = window.STONEFLOW_CONFIG || {};
  const { heldOf, normCode } = window.StoneSheet;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const S = { started: false, mode: null, auth: null, store: null, meta: null, lots: [], log: [], who: "", loadedAt: null,
    filter: "in", sec: "", q: "", sel: null, chat: [], busy: false, ctl: null, writing: 0, sample: null, downloads: null };

  /* ---------- helpers ---------- */
  const stockOf = (l) => l.stock;
  const holdOf = (l) => (l.heldTotal != null ? l.heldTotal : heldOf(l));
  const freeOf = (l) => Math.max(0, l.stock - holdOf(l));
  const viewOnly = () => S.mode === "view";
  const sizeOf = (l) => (l.L || l.H) ? `${l.L || "?"} x ${l.H || "?"}` : "";
  const lotName = (l) => `${l.lot || "(no lot #)"} ${l.name}`.trim();
  const holdText = (l) => [...Object.entries(l.holds || {}).filter(([, n]) => n).map(([p, n]) => `${n} ${p}`), ...(l.mgrHold ? [`${l.mgrHold} (sales mgr)`] : [])].join(", ");
  const fmtTime = (iso) => { try { return new Date(iso).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }); } catch { return ""; } };
  const toks = (s) => String(s || "").toLowerCase().replace(/(\d)\s*[x×*]\s*(\d)/g, "$1 $2").split(/[\s,]+/).filter(Boolean);
  function matches(l, ts) {
    if (!ts.length) return true;
    const hay = [l.lot, l.name, l.section, l.loc, l.L, l.H, sizeOf(l), l.bm ? "bookmatch bm" : "", l.outside ? "outside" : ""].join(" ").toLowerCase();
    return ts.every((t) => hay.includes(t) || (t.length > 2 && normCode(l.lot) === normCode(t)));
  }
  function busy(on, text) {
    S.writing += on ? 1 : -1;
    $("busy").hidden = S.writing <= 0; if (on) $("busy").textContent = text || "Saving to Excel…";
  }
  function note(html, cls) { const n = $("note"); n.hidden = !html; n.className = "notice" + (cls ? " " + cls : ""); n.innerHTML = html || ""; }
  function flash(t) { note(`<span>${esc(t)}</span>`); clearTimeout(flash.t); flash.t = setTimeout(renderNote, 6000); }

  /* ---------- start ---------- */
  async function open() {
    if (S.started) return; S.started = true;
    const use = (n) => (window.claude && typeof claude.use === "function") ? claude.use(n).catch(() => null) : Promise.resolve(null);
    const [sample, downloads] = await Promise.all([use("sample"), use("downloads")]);
    if (sample) { const lim = await sample.limits().catch(() => null); S.sample = lim?.tools ? sample : null; }
    S.downloads = downloads;

    if (CFG.microsoft?.clientId && !window.claude && CFG.serverMode !== false) return openServer();
    if (CFG.microsoft?.clientId && !window.claude) {
      S.mode = "live";
      S.auth = window.MsAuth.create({ clientId: CFG.microsoft.clientId, tenant: CFG.microsoft.tenant });
      if (!S.auth.signedIn()) return showGate("signin", S.pendingErr);
      try {
        busy(true, "Opening the Excel file…");
        S.store = await window.InventoryStores.connect(CFG.excel || {}, S.auth.getToken);
        const me = await window.InventoryStores.graph(S.auth.getToken)("/me?$select=displayName,mail,userPrincipalName");
        S.who = me.displayName || me.mail || me.userPrincipalName || "";
        S.email = String(me.mail || me.userPrincipalName || "").toLowerCase();
      } catch (e) { busy(false); return showGate(e.code === "auth" ? "signin" : "error", e.message); }
      busy(false);
      return start();
    }
    if (window.STONEFLOW_DEMO_XLSX) {
      const bytes = Uint8Array.from(atob(window.STONEFLOW_DEMO_XLSX), (c) => c.charCodeAt(0));
      return startDemo(bytes, "LIVE_INVENTORY.xlsx (demo copy)");
    }
    showGate("demo-pick");
  }
  // Server mode: anyone can see the stock (read-only). Staff on the editor list sign in to make changes.
  async function openServer() {
    S.auth = window.MsAuth.create({ clientId: CFG.microsoft.clientId, tenant: CFG.microsoft.tenant });
    const getToken = () => (S.auth.signedIn() ? S.auth.getToken() : Promise.resolve(null));
    busy(true, "Reading the inventory…");
    let me = { role: "viewer", connected: true };
    try { me = await (await fetch("/api/me", { headers: S.auth.signedIn() ? { authorization: "Bearer " + await S.auth.getToken() } : {} })).json(); } catch {}
    S.me = me; S.who = me.name || me.email || ""; S.email = String(me.email || "").toLowerCase();
    if (!me.connected) { busy(false); return showGate("connect"); }
    try {
      if (me.role === "editor") { S.mode = "live"; S.store = await window.InventoryStores.connectServer(CFG.excel || {}, getToken); }
      else { S.mode = "view"; S.store = window.InventoryStores.view(getToken); }
    } catch (e) { busy(false); return showGate("error", e.message); }
    busy(false);
    return start();
  }
  function startDemo(bytes, name) { S.mode = "demo"; S.store = window.InventoryStores.demo(bytes, name); S.who = "Demo user"; start(); }

  async function start() {
    $("gate").hidden = true; $("desk-wrap").hidden = false; $("pActions").hidden = false;
    const signedIn = !!S.auth?.signedIn?.();
    $("signOut").hidden = !signedIn; $("demoDownload").hidden = S.mode !== "demo";
    $("signIn").hidden = !(S.auth && !signedIn);
    $("whoami").textContent = S.who + (S.mode === "view" && signedIn ? " · view only" : "");
    document.body.classList.toggle("view-only", viewOnly());
    $("stSyncL").textContent = viewOnly() ? "Last updated" : "Read from Excel";
    $("invFoot").textContent = viewOnly() ? "Stock is read live from the yard's inventory. To buy or reserve slabs, contact us or talk to your salesperson." : "Tap a lot to record a sale or change holds. Every change is written straight into the Excel file.";
    await load();
    if (!viewOnly()) S.store.log.read().then((l) => { S.log = l || []; renderLog(); }).catch(() => {});
    if (!viewOnly()) S.store.orders.read().then((o) => { S.orderRows = o || []; renderOrders(); }).catch(() => {});
    setInterval(async () => {
      if (document.visibilityState !== "visible" || S.busy || S.writing || !S.store) return;
      if (await S.store.changed()) load(undefined, true);
    }, 30000);
  }

  function showGate(kind, msg) {
    $("desk-wrap").hidden = true; $("pActions").hidden = kind === "signin" || kind === "demo-pick" ? true : $("pActions").hidden;
    const g = $("gate"); g.hidden = false;
    if (kind === "signin") {
      g.innerHTML = `<h2>Sign in to the inventory</h2><p>Use the Microsoft account that can open the LIVE_INVENTORY Excel file on OneDrive. Your sales and holds are written straight into that file.</p>
        ${msg ? `<p style="color:var(--sold)">${esc(msg)}</p>` : ""}<div><button class="btn primary" type="button" id="msSignIn">Sign in with Microsoft</button></div>`;
      $("msSignIn").onclick = () => S.auth.signIn();
    } else if (kind === "connect") {
      g.innerHTML = `<h2>Inventory not connected yet</h2><p>The owner needs to connect the site to the OneDrive that holds the Excel file. This is done once.</p>
        <div><a class="btn primary" href="/api/connect">Owner: connect OneDrive</a></div>`;
    } else if (kind === "error" && S.me) {
      g.innerHTML = `<h2>Couldn't read the inventory</h2><p>${esc(msg)}</p><div><button class="btn" type="button" id="gRetry">Try again</button></div>`;
      $("gRetry").onclick = () => location.reload();
    } else if (kind === "error") {
      g.innerHTML = `<h2>Couldn't open the Excel file</h2><p>${esc(msg)}</p><p>Check that the sharing link in <code>js/config.js</code> is a "Can edit" link and that you are signed in with an account it was shared with.</p>
        <div><button class="btn" type="button" id="gRetry">Try again</button> <button class="btn" type="button" id="gOut">Sign in as someone else</button></div>`;
      $("gRetry").onclick = () => location.reload();
      $("gOut").onclick = () => { S.auth.signOut(); location.reload(); };
    } else {
      g.innerHTML = `<h2>Demo mode</h2><p>This site isn't connected to OneDrive yet. Open a copy of your LIVE_INVENTORY Excel file to try the portal and the agent. The file stays in this browser.</p>
        <ol><li>Fill in <code>js/config.js</code> to connect OneDrive (see README).</li><li>Until then, test here with a copy of the file.</li></ol>
        <div><input type="file" id="demoFile" accept=".xlsx" hidden><button class="btn primary" type="button" id="demoPick">Open a copy of the Excel file</button></div>`;
      $("demoPick").onclick = () => $("demoFile").click();
      $("demoFile").onchange = async (e) => { const f = e.target.files?.[0]; if (f) startDemo(new Uint8Array(await f.arrayBuffer()), f.name + " (demo copy)"); };
    }
  }

  async function load(sheet, quiet) {
    if (!quiet) busy(true, "Reading the Excel file…");
    try {
      const m = await S.store.load(sheet);
      S.meta = m.meta; S.lots = m.lots; S.loadedAt = new Date();
      if (S.sel && !S.lots.some((l) => l.id === S.sel)) S.sel = null;
      renderAll();
    } catch (e) {
      if (e.code === "auth") return S.me ? (S.auth.signOut(), location.reload()) : showGate("signin", e.message);
      if (e.code === "not_connected") return showGate("connect");
      flash(e.message);
    } finally { if (!quiet) busy(false); }
  }

  /* ---------- rendering ---------- */
  function renderNote() {
    const m = S.meta; if (!m) return note("");
    // Sheet notes are only for the people listed in adminEmails (config.js); everyone else sees no notice.
    const admins = (CFG.adminEmails || []).map((x) => String(x).trim().toLowerCase()).filter(Boolean);
    const isAdmin = S.mode === "demo" || (S.email && admins.includes(S.email));
    if (!isAdmin) return note("");
    const w = m.warnings || [];
    const src = S.mode === "demo"
      ? `<span><strong>Demo copy</strong> of ${esc(m.name.replace(/ \(demo copy\)$/, ""))}, sheet ${esc(m.sheet)}. Changes stay in this browser; use <strong>Download changed file</strong> to see them in Excel.</span>`
      : `<span>Connected to <strong>${esc(m.name)}</strong> on ${m.driveType === "personal" ? "OneDrive" : "OneDrive for Business / SharePoint"}, sheet <strong>${esc(m.sheet)}</strong> · ${m.mode === "workbook" ? "live Excel edits" : "file edits with version check"}</span>`;
    note(src + (w.length ? `<details class="warns"><summary>${w.length} thing${w.length > 1 ? "s" : ""} to check in the sheet</summary><ul>${w.map((x) => `<li>${esc(x)}</li>`).join("")}</ul></details>` : ""), S.mode === "demo" ? "demo" : "");
  }
  function renderHead() {
    const m = S.meta;
    $("pSource").textContent = S.mode === "demo" ? "Demo · Excel copy in this browser" : viewOnly() ? "Live stock · updated from the yard" : "Live inventory · OneDrive Excel · you can make changes";
    $("sheetSel").innerHTML = (m.sheets || [m.sheet]).map((n) => `<option ${n === m.sheet ? "selected" : ""}>${esc(n)}</option>`).join("");
    $("sheetSel").hidden = (m.sheets || []).length < 2;
    const oe = $("openExcel"); oe.hidden = !m.webUrl || viewOnly(); if (m.webUrl) oe.href = m.webUrl;
    const sel = $("secSel"), cur = S.sec;
    sel.innerHTML = `<option value="">All sections</option>` + (m.sections || []).map((s) => `<option ${s === cur ? "selected" : ""}>${esc(s)}</option>`).join("");
  }
  function renderStats() {
    let free = 0, hold = 0, stock = 0, lotsIn = 0;
    for (const l of S.lots) if (l.stock > 0) { lotsIn++; stock += l.stock; free += freeOf(l); hold += Math.min(holdOf(l), l.stock); }
    $("stFree").textContent = free.toLocaleString(); $("stHold").textContent = hold.toLocaleString();
    $("stStock").textContent = stock.toLocaleString(); $("stLots").textContent = lotsIn.toLocaleString();
    $("stSync").textContent = S.loadedAt ? S.loadedAt.toLocaleTimeString([], { hour: "numeric", minute: "2-digit", second: "2-digit" }) : "–";
  }
  function renderTable() {
    const box = $("tablebox"), ts = toks(S.q);
    const rows = S.lots.filter((l) => (S.filter === "all" || (S.filter === "in" ? l.stock > 0 : l.stock <= 0)) && (!S.sec || l.section === S.sec) && matches(l, ts));
    $("invCount").textContent = `${rows.length} of ${S.lots.length} lots`;
    if (!rows.length) { box.innerHTML = `<div class="empty">No lots match.</div>`; return; }
    const rack = !viewOnly(), ncol = rack ? 7 : 6;
    let h = `<table><thead><tr><th>Lot #</th><th>Color</th><th>Size</th><th class="n">Stock</th><th class="n">Hold</th><th class="n">Free</th>${rack ? "<th>Rack</th>" : ""}</tr></thead><tbody>`, sec = null;
    for (const l of rows) {
      if (l.section !== sec) { sec = l.section; h += `<tr class="sec"><td colspan="${ncol}">${esc(sec || "—")}</td></tr>`; }
      const hd = holdOf(l), fr = freeOf(l);
      h += `<tr class="lotrow ${l.stock <= 0 ? "out" : ""} ${S.sel === l.id ? "sel" : ""}" data-id="${esc(l.id)}" tabindex="0">
        <td class="code">${esc(l.lot || "—")}</td><td class="color">${esc(l.name)} ${l.bm ? '<span class="pill bm">BM</span>' : ""}</td>
        <td>${esc(sizeOf(l))}</td><td class="n">${l.stock < 0 ? `<b class="neg">${l.stock}</b>` : l.stock}</td>
        <td class="n">${hd ? `<span class="pill hold">${hd}</span>` : ""}</td><td class="n">${fr ? `<b class="free">${fr}</b>` : `<b class="zero">0</b>`}</td>
        ${rack ? `<td class="loc" title="${esc(l.loc)}">${esc(l.loc)}</td>` : ""}</tr>`;
    }
    box.innerHTML = h + "</tbody></table>";
  }
  function renderCard() {
    const box = $("lotCard"), l = S.lots.find((x) => x.id === S.sel);
    if (!l) { box.hidden = true; box.innerHTML = ""; return; }
    const drafts = box.dataset.lot === l.id ? [...box.querySelectorAll("input")].filter((i) => i.dataset.dirty === "1").map((i) => [i.id, i.value]) : [];
    box.dataset.lot = l.id; box.hidden = false;
    const people = S.meta.holdPeople || [], hd = holdOf(l), fr = freeOf(l);
    if (viewOnly()) {
      box.innerHTML = `<div class="lotcard"><div class="topline"><h3>${esc(lotName(l))}</h3><button class="link" type="button" id="cardClose">Close</button></div>
        <div class="meta">${esc([l.section, sizeOf(l) && sizeOf(l) + " in", l.bm && "bookmatched"].filter(Boolean).join(" · "))}</div>
        <div class="nums"><div><b>${l.stock}</b><span>In stock</span></div><div><b>${hd}</b><span>On hold</span></div><div><b style="color:var(--ok)">${fr}</b><span>Free to sell</span></div></div>
        <div class="meta">To buy or reserve these slabs, contact us or talk to your salesperson.</div></div>`;
      $("cardClose").onclick = () => { S.sel = null; renderCard(); renderTable(); };
      return;
    }
    const recent = S.log.filter((e) => e.lot === l.lot && e.color === l.name).slice(-5).reverse();
    box.innerHTML = `<div class="lotcard">
      <div class="topline"><h3>${esc(lotName(l))}</h3><button class="link" type="button" id="cardClose">Close</button></div>
      <div class="meta">${esc([l.section, sizeOf(l) && sizeOf(l) + " in", l.bm && "bookmatched", l.loc && "rack " + l.loc, "Excel row " + l.row].filter(Boolean).join(" · "))}</div>
      <div class="nums"><div><b>${l.stock}</b><span>In stock</span></div><div><b>${hd}</b><span>On hold${hd ? " · " + esc(holdText(l)) : ""}</span></div><div><b style="color:var(--ok)">${fr}</b><span>Free to sell</span></div></div>
      <div class="forms">
        <div class="fbox"><h4>Record sale</h4>
          <div class="frow"><input type="number" id="saleQty" min="1" step="1" value="1" aria-label="Slabs"><input class="note" id="saleNote" placeholder="Customer / job (optional)" aria-label="Note"></div>
          <div class="frow"><button class="btn sell" type="button" id="doSale">Record sale</button><button class="btn" type="button" id="doReturn">Record return</button></div>
          <span class="eyebrow" id="saleMsg"></span></div>
        <div class="fbox"><h4>Holds</h4>
          ${l.outside ? `<span class="eyebrow">Outside-warehouse lots have no hold columns.</span>` : `<div class="holdgrid">${people.map((p, i) => `<label for="h${i}">${esc(p)}<input type="number" min="0" step="1" id="h${i}" data-p="${esc(p)}" value="${Number(l.holds?.[p]) || 0}"></label>`).join("")}</div>
          ${l.mgrHold ? `<span class="eyebrow">Sales mgr hold: ${esc(l.mgrHold)} (edit in Excel)</span>` : ""}
          <div class="frow"><button class="btn" type="button" id="saveHolds">Save holds</button><span class="eyebrow" id="holdMsg"></span></div>`}</div>
      </div>
      ${recent.length ? `<div class="meta">Recent: ${recent.map((e) => `${esc(fmtTime(e.at))} ${esc(e.action)} ${e.qty || ""} ${e.who ? "· " + esc(e.who) : ""}`).join(" · ")}</div>` : ""}
    </div>`;
    for (const [id, v] of drafts) { const i = $(id); if (i) { i.value = v; i.dataset.dirty = "1"; } }
    box.querySelectorAll("input").forEach((i) => i.addEventListener("input", () => { i.dataset.dirty = "1"; }));
    const cm = S.cardMsg && S.cardMsg.lot === l.id ? S.cardMsg : {};
    $("saleMsg").textContent = cm.sale || ""; if ($("holdMsg")) $("holdMsg").textContent = cm.hold || "";
    const say = (k, t) => { S.cardMsg = { lot: l.id, [k]: t }; const el = $(k === "sale" ? "saleMsg" : "holdMsg"); if (el) el.textContent = t; };
    $("cardClose").onclick = () => { S.sel = null; renderCard(); renderTable(); };
    const doSale = async (sign) => {
      const q = Math.floor(Number($("saleQty").value)); if (!(q > 0)) return say("sale", "Enter how many slabs.");
      const n = $("saleNote").value; $("saleNote").value = ""; $("saleNote").dataset.dirty = ""; $("saleQty").dataset.dirty = "";
      try { await recordSale(l, sign * q, n, "by hand"); say("sale", sign > 0 ? `Recorded ${q} sold in Excel.` : `Recorded ${q} returned in Excel.`); }
      catch (e) { say("sale", e.message); }
    };
    $("doSale").onclick = () => doSale(1); $("doReturn").onclick = () => doSale(-1);
    const sh = $("saveHolds");
    if (sh) sh.onclick = async () => {
      const want = {}; box.querySelectorAll("[data-p]").forEach((i) => { want[i.dataset.p] = Math.max(0, Math.floor(Number(i.value) || 0)); });
      const order = Object.keys(want).filter((p) => (Number(l.holds?.[p]) || 0) !== want[p]).sort((a, b) => (want[a] - (l.holds?.[a] || 0)) - (want[b] - (l.holds?.[b] || 0)));
      box.querySelectorAll("[data-p]").forEach((i) => { i.dataset.dirty = ""; });
      try { for (const p of order) await setHold(l, p, want[p], "by hand"); say("hold", order.length ? "Holds saved in Excel." : "No changes."); }
      catch (e) { say("hold", e.message); }
    };
  }
  function renderLog() {
    $("logWhere").textContent = S.mode === "demo" ? "this browser" : (CFG.excel?.salesLogName || "sales log");
    const items = S.log.slice(-15).reverse();
    $("log").innerHTML = items.length ? items.map((e) => `<div><time>${esc(fmtTime(e.at))}</time><span>${esc(`${e.action} ${e.qty ? Math.abs(e.qty) + " · " : ""}${e.lot} ${e.color}${e.note ? " · " + e.note : ""}${e.who ? " · " + e.who : ""}`)}</span></div>`).join("")
      : `<div><span style="color:var(--muted)">No changes yet.</span></div>`;
  }
  function renderAll() { renderHead(); renderNote(); renderStats(); renderTable(); renderCard(); renderLog(); renderOrders(); renderMsgs(); }

  /* ---------- writes (all go straight into the Excel file) ---------- */
  async function logIt(action, l, qty, note, cell, via) {
    const line = { at: new Date().toISOString(), who: S.who + (via === "agent" ? " (agent)" : ""), action, lot: l.lot, color: l.name, qty, note: note || "", cell: cell || "" };
    S.log.push(line); renderLog();
    try { await S.store.log.append([line]); } catch {}
  }
  async function recordSale(l, qty, note, via) {
    if (!qty) throw new Error("Quantity must be a whole number.");
    busy(true);
    try {
      const res = await S.store.sale(l, qty);
      l.stock = res.stockAfter;
      await logIt(qty > 0 ? "Sold" : "Returned", l, qty, note, res.change.ref, via);
      renderAll(); load(undefined, true);
      return res;
    } finally { busy(false); }
  }
  async function setHold(l, person, qty, via) {
    busy(true);
    try {
      const res = await S.store.hold(l, person, qty);
      l.holds = { ...(l.holds || {}), [person]: qty }; if (!qty) delete l.holds[person];
      await logIt(`Hold ${person}`, l, qty, `${res.change.prev} → ${qty}`, res.change.ref, via);
      renderAll(); load(undefined, true);
      return res;
    } finally { busy(false); }
  }
  async function undo(a) {
    if (a.kind === "order") return cancelOrder(a.order);
    const l = S.lots.find((x) => x.id === a.lotId) || S.lots.find((x) => x.lot === a.lot && x.name === a.name);
    if (!l) throw new Error("That lot isn't in the sheet anymore.");
    busy(true);
    try { await S.store.undo(l, a.change); await logIt("Undo", l, 0, a.label, a.change.ref); await load(undefined, true); }
    finally { busy(false); }
  }

  /* ---------- sales orders (stock deducted in Excel + a line per slab lot in "StoneFlow sales orders.csv") ---------- */
  S.orderRows = [];
  const money = (n) => n == null || n === "" || isNaN(n) ? "" : Number(n).toLocaleString(undefined, { style: "currency", currency: CFG.currency || "USD" });
  const sqftOf = (l, qty) => { const a = parseFloat(l.L), b = parseFloat(l.H); return isFinite(a) && isFinite(b) ? Math.round((a * b / 144) * qty * 10) / 10 : null; };
  function ordersFromRows(rows) {
    const map = new Map();
    for (const r of rows || []) {
      if (!r.order) continue;
      let o = map.get(r.order);
      if (!o) { o = { order: r.order, date: r.date, client: r.client, phone: r.phone, staff: r.staff, notes: r.notes, status: "Open", lines: [] }; map.set(r.order, o); }
      if (r.status === "Cancelled" && !r.lot) { o.status = "Cancelled"; o.cancelledAt = r.date; continue; }
      o.lines.push(r);
    }
    return [...map.values()].reverse();
  }
  async function nextOrderNo() {
    try { S.orderRows = await S.store.orders.read(); } catch {}
    const d = new Date(), prefix = `SO-${String(d.getFullYear()).slice(2)}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}-`;
    const n = Math.max(0, ...S.orderRows.filter((r) => String(r.order).startsWith(prefix)).map((r) => Number(String(r.order).slice(prefix.length)) || 0));
    return prefix + String(n + 1).padStart(2, "0");
  }
  async function createOrder({ client, phone, lines, notes, via }) {
    client = String(client || "").trim();
    if (!client) throw new Error("Which client is this sales order for?");
    if (!Array.isArray(lines) || !lines.length) throw new Error("Add at least one lot and quantity.");
    const items = lines.map((ln) => {
      const l = S.lots.find((x) => x.id === String(ln.lot_id));
      if (!l) throw new Error(`No lot with id ${ln.lot_id}. Use ids from find_lots.`);
      const qty = Math.floor(Number(ln.qty));
      if (!(qty > 0)) throw new Error(`Quantity for ${lotName(l)} must be 1 or more.`);
      const price = ln.price_per_slab == null || ln.price_per_slab === "" ? null : Number(ln.price_per_slab);
      return { l, qty, price: isFinite(price) ? price : null };
    });
    const need = {}; for (const it of items) need[it.l.id] = (need[it.l.id] || 0) + it.qty;
    for (const it of items) if (need[it.l.id] > it.l.stock) throw new Error(`Only ${Math.max(it.l.stock, 0)} slab(s) of ${lotName(it.l)} in stock, but the order needs ${need[it.l.id]}. Nothing was recorded.`);
    busy(true, "Creating the sales order…");
    const order = await nextOrderNo(), done = [];
    try {
      for (const it of items) {                                     // each line re-checks the live Excel row before writing
        const res = await S.store.sale(it.l, it.qty);
        it.l.stock = res.stockAfter; done.push({ it, change: res.change, held: res.heldInvolved || 0 });
      }
    } catch (e) {
      for (const d of done.slice().reverse()) { try { await S.store.undo(d.it.l, d.change); } catch {} }
      busy(false); load(undefined, true);
      throw new Error(`${e.message} The order was not created and no stock was changed.`);
    }
    const date = new Date().toISOString(), staff = S.who + (via === "agent" ? " (agent)" : "");
    const rows = done.map((d) => ({ order, date, client, phone: phone || "", lot: d.it.l.lot, color: d.it.l.name, size: sizeOf(d.it.l), qty: d.it.qty,
      sqft: sqftOf(d.it.l, d.it.qty), price: d.it.price, total: d.it.price != null ? Math.round(d.it.price * d.it.qty * 100) / 100 : null,
      cell: d.change.ref, status: "Open", staff, notes: notes || "" }));
    let saved = false; try { saved = await S.store.orders.append(rows); } catch {}
    S.orderRows = [...S.orderRows.filter((r) => r.order !== order), ...rows];
    for (const d of done) await logIt("Sold", d.it.l, d.it.qty, `${order} · ${client}`, d.change.ref, via);
    busy(false); renderAll(); load(undefined, true);
    return { order, client, rows, saved, held: done.reduce((t, d) => t + d.held, 0) };
  }
  async function cancelOrder(orderNo, via) {
    const o = ordersFromRows(S.orderRows).find((x) => x.order === orderNo);
    if (!o) throw new Error(`No sales order ${orderNo}.`);
    if (o.status === "Cancelled") throw new Error(`${orderNo} is already cancelled.`);
    busy(true, "Cancelling the order…");
    try {
      for (const ln of o.lines) {
        const l = S.lots.find((x) => x.lot === ln.lot && x.name === ln.color);
        if (!l) throw new Error(`Lot ${ln.lot} isn't in the sheet anymore, so return it by hand.`);
        await S.store.undo(l, { kind: l.outside ? "outside" : "sale", ref: ln.cell, qty: Number(ln.qty), row: l.row });
        await logIt("Returned", l, -Number(ln.qty), `${orderNo} cancelled`, ln.cell, via);
      }
    } finally { busy(false); }
    const row = { order: orderNo, date: new Date().toISOString(), client: o.client, status: "Cancelled", staff: S.who + (via === "agent" ? " (agent)" : "") };
    try { await S.store.orders.append([row]); } catch {}
    S.orderRows.push(row); renderAll(); load(undefined, true);
    return { cancelled: orderNo, slabs_back_in_stock: o.lines.reduce((t, l) => t + Number(l.qty), 0) };
  }
  function renderOrders() {
    const box = $("orders"); if (!box) return;
    $("ordersWhere").textContent = S.mode === "demo" ? "this browser" : (CFG.excel?.ordersFileName || "StoneFlow sales orders.csv");
    const list = ordersFromRows(S.orderRows).slice(0, 12);
    box.innerHTML = list.length ? list.map((o) => {
      const slabs = o.lines.reduce((t, l) => t + Number(l.qty || 0), 0), total = o.lines.reduce((t, l) => t + (Number(l.total) || 0), 0);
      return `<div><time>${esc(fmtTime(o.date))}</time><span><b>${esc(o.order)}</b> · ${esc(o.client)} · ${slabs} slab${slabs === 1 ? "" : "s"}${total ? " · " + money(total) : ""} ${o.status === "Cancelled" ? '<span class="pill out">Cancelled</span>' : ""}</span><button class="link" type="button" data-view-order="${esc(o.order)}">View</button></div>`;
    }).join("") : `<div><span style="color:var(--muted)">No sales orders yet. Ask the agent, e.g. "Create a sales order for ABC Granite: 3 slabs of IQ466".</span></div>`;
  }
  function orderHtml(o) {
    const C = CFG.companyDetails || {};
    const lines = o.lines, slabs = lines.reduce((t, l) => t + Number(l.qty || 0), 0);
    const sqft = Math.round(lines.reduce((t, l) => t + (Number(l.sqft) || 0), 0) * 10) / 10, total = lines.reduce((t, l) => t + (Number(l.total) || 0), 0);
    return `<div class="so">
      <div class="so-top"><div><b class="so-co">${esc(CFG.company || "")}</b><div>${esc(C.address || "")}</div><div>${esc([C.phone, C.email].filter(Boolean).join(" · "))}</div></div>
        <div class="so-id"><h2>Sales order</h2><div><b>${esc(o.order)}</b></div><div>${esc(new Date(o.date).toLocaleDateString())}</div>${o.status === "Cancelled" ? '<div class="pill out">Cancelled</div>' : ""}</div></div>
      <div class="so-party"><div><span class="eyebrow">Sold to</span><div><b>${esc(o.client)}</b></div><div>${esc(o.phone || "")}</div></div><div><span class="eyebrow">Salesperson</span><div>${esc(o.staff || "")}</div></div></div>
      <div class="so-tw"><table class="so-t"><thead><tr><th>#</th><th>Lot #</th><th>Color</th><th>Size (in)</th><th class="n">Slabs</th><th class="n">Sq ft</th><th class="n">Price / slab</th><th class="n">Amount</th></tr></thead><tbody>
      ${lines.map((l, i) => `<tr><td>${i + 1}</td><td>${esc(l.lot)}</td><td>${esc(l.color)}</td><td>${esc(l.size)}</td><td class="n">${esc(l.qty)}</td><td class="n">${l.sqft ?? ""}</td><td class="n">${money(l.price)}</td><td class="n">${money(l.total)}</td></tr>`).join("")}
      </tbody><tfoot><tr><td colspan="4">Total</td><td class="n">${slabs}</td><td class="n">${sqft || ""}</td><td></td><td class="n">${total ? money(total) : ""}</td></tr></tfoot></table></div>
      ${o.notes ? `<p><span class="eyebrow">Notes</span><br>${esc(o.notes)}</p>` : ""}
      <p class="so-fine">Slabs are reserved for this order and taken out of the yard inventory. Bring this order number at pick-up.</p>
      <div class="so-sign"><div>Customer signature</div><div>Date</div></div></div>`;
  }
  function showOrder(orderNo) {
    const o = ordersFromRows(S.orderRows).find((x) => x.order === orderNo); if (!o) return flash(`No sales order ${orderNo}.`);
    let m = $("soModal");
    if (!m) { m = document.createElement("div"); m.id = "soModal"; m.className = "modal"; document.body.appendChild(m); m.addEventListener("click", (e) => { if (e.target === m) m.hidden = true; }); }
    const canPrint = !window.claude;
    m.hidden = false;
    m.innerHTML = `<div class="modal-box" role="dialog" aria-modal="true" aria-label="Sales order ${esc(o.order)}">${orderHtml(o)}
      <div class="modal-actions">${canPrint ? '<button class="btn primary" type="button" id="soPdf">Save as PDF</button><button class="btn" type="button" id="soPrint">Print</button>' : ""}<button class="btn" type="button" id="soDl">Download</button>
      ${o.status !== "Cancelled" ? '<button class="btn" type="button" id="soCancel">Cancel order</button>' : ""}<button class="btn" type="button" id="soClose">Close</button><span class="eyebrow" id="soMsg"></span></div></div>`;
    $("soClose").onclick = () => { m.hidden = true; };
    if ($("soPrint")) $("soPrint").onclick = () => window.print();
    // PDF: the browser's print dialog with "Save as PDF"; the page title becomes the file name (e.g. "SO-261015-01 ABC Granite.pdf").
    if ($("soPdf")) $("soPdf").onclick = () => {
      const old = document.title; document.title = `${o.order} ${o.client}`.replace(/[\\/:*?"<>|]/g, "");
      $("soMsg").textContent = "Choose \"Save as PDF\" as the printer";
      const back = () => { document.title = old; window.removeEventListener("afterprint", back); };
      window.addEventListener("afterprint", back); window.print(); setTimeout(back, 60000);
    };
    $("soDl").onclick = async () => {
      const css = [...document.styleSheets].map((ss) => { try { return [...ss.cssRules].map((r) => r.cssText).join("\n"); } catch { return ""; } }).join("\n");
      const html = `<!doctype html><meta charset="utf-8"><title>${esc(o.order)}</title><style>${css}</style><body style="padding:24px">${orderHtml(o)}</body>`;
      const name = `${o.order} ${o.client}.html`.replace(/[\\/:*?"<>|]/g, "");
      if (S.downloads) { try { await S.downloads.save({ filename: name, data: html }); } catch (e) { if (e?.code !== "declined") $("soMsg").textContent = "Couldn't save the file here."; } return; }
      const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([html], { type: "text/html" })); a.download = name; a.click();
    };
    if ($("soCancel")) $("soCancel").onclick = async (e) => {
      e.target.disabled = true;
      try { await cancelOrder(o.order); S.chat.forEach((c) => (c.acts || []).forEach((a) => { if (a.kind === "order" && a.order === o.order) a.undone = true; })); renderMsgs(); showOrder(o.order); }
      catch (err) { $("soMsg").textContent = err.message; e.target.disabled = false; }
    };
  }

  /* ---------- agent ---------- */
  function lotView(l) {
    return { id: l.id, lot: l.lot || "(none)", color: l.name, section: l.section, size_in: sizeOf(l) || "not listed", bookmatched: !!l.bm,
      ...(viewOnly() ? {} : { racks: l.loc || "", hold_detail: holdText(l) || undefined }), in_stock: l.stock, on_hold: holdOf(l), free_to_sell: freeOf(l),
      ...(l.outside ? { outside_warehouse: true } : {}), ...(l.stock < 0 ? { data_problem: "negative stock in sheet" } : {}) };
  }
  function rules() {
    const m = S.meta, inStock = S.lots.filter((l) => l.stock > 0).length;
    return `You are the inventory agent for a stone countertop slab yard. You keep the LIVE lot inventory correct and answer stock questions for sales staff. The inventory is the Excel file "${m.name}" (sheet ${m.sheet}) on OneDrive; your changes are written straight into it.
Today is ${new Date().toDateString()}. You are talking with ${S.who || "a staff member"}. ${S.lots.length} lots (${inStock} with stock). Each LOT # (like IQ466 or IA4826) is a batch of identical slabs; stock is a slab COUNT. Size is L x H in inches. Section names tell the type (GRANITE, CALACATTA / CARRARA quartz, JUMBO vs REGULAR size, PRINTED, 2CM, ONYX). In color names -J = jumbo, -SJ = super jumbo, -R = regular, PFB = printed full body. "Outside" lots are counts at an outside warehouse. Salespeople who hold slabs: ${(m.holdPeople || []).join(", ")}.

Tools: find_lots (always call before answering about stock; never guess), record_sale, set_hold, create_sales_order, cancel_sales_order, find_orders.
When asked if something is available: search by lot # or color words. Answer with total free-to-sell slabs, then each matching lot on its own line: lot #, color, size, free count (and hold detail if any), racks. If several colors match, group by color. If nothing is in stock, say so.
When told slabs were sold: identify the ONE lot (by lot #, or color plus size). If more than one lot could match, ask which lot before recording. Then call record_sale with that lot id and the quantity (default 1). If the quantity is more than free_to_sell but not more than in_stock, ask whether the sale is for the person holding them before recording. Never record a sale larger than in_stock.
Returns or cancelled sales: record_sale with type "return".
Holds: "hold 5 IQ1466J for Puru" → set_hold with that person's name and the NEW total for that person (0 to release).
Sales orders: when the user is selling to a client ("create a sales order for ABC Granite: 3 IQ466 and 2 IQ1466J at $450 a slab"), call find_lots for each item, then call create_sales_order ONCE with the client name and every line (lot id, slabs, price per slab if given). Ask only for what is missing: the client name, the quantity, or which lot when several match. It checks stock for every line first; if a line is short nothing is recorded, so tell the user what is free instead. After creating, reply with the order number, client, each line (lot #, color, size, slabs, price) and total slabs, sq ft and amount. To cancel an order use cancel_sales_order. Use record_sale only for a quick sale with no client order. Use find_orders for questions about past orders.
Language: staff may write or speak in English or Spanish (spoken messages can have small transcription mistakes, e.g. "I Q 466" means IQ466, "for sixty six" may mean 466). Always reply in the language of the user's latest message. Keep lot numbers, color names and sizes exactly as in the inventory.
Reply briefly for staff on a phone: short lines, "•" for list items, **bold** only for the key number. No tables or headings. Never invent lots, sizes or counts.`;
  }
  const TOOL_DEFS = [
    { name: "find_lots", description: "Search the live lot inventory. Returns up to 40 lots (most free slabs first) with id, lot #, color, section, size, racks, in_stock, on_hold, free_to_sell, plus totals.",
      input_schema: { type: "object", properties: { lot: { type: "string", description: "Lot number such as IQ466. Matched loosely." }, text: { type: "string", description: "Words that must all appear: color, section, size like 127x64, rack." }, in_stock_only: { type: "boolean", description: "Default true." } } } },
    { name: "record_sale", description: "Record slabs sold (type 'sale') or returned (type 'return') for one lot id from find_lots. Writes the sale into the Excel file and returns the new stock.",
      input_schema: { type: "object", required: ["lot_id", "qty"], properties: { lot_id: { type: "string" }, qty: { type: "integer", minimum: 1 }, type: { type: "string", enum: ["sale", "return"] }, note: { type: "string", description: "Customer, job or salesperson mentioned." } } } },
    { name: "set_hold", description: "Set how many slabs of one lot are held for one salesperson (the new total; 0 releases). Person must be one of the hold names.",
      input_schema: { type: "object", required: ["lot_id", "person", "qty"], properties: { lot_id: { type: "string" }, person: { type: "string" }, qty: { type: "integer", minimum: 0 } } } },
    { name: "create_sales_order", description: "Create a sales order for a client: checks every line against live stock, deducts the slabs in the Excel file, and saves the order. Returns the order number and lines.",
      input_schema: { type: "object", required: ["client", "lines"], properties: {
        client: { type: "string", description: "Client / company name." }, client_phone: { type: "string" },
        lines: { type: "array", items: { type: "object", required: ["lot_id", "qty"], properties: { lot_id: { type: "string", description: "Lot id from find_lots." }, qty: { type: "integer", minimum: 1 }, price_per_slab: { type: "number" } } } },
        notes: { type: "string", description: "Job name, delivery or pick-up details." } } } },
    { name: "cancel_sales_order", description: "Cancel a sales order by its number: puts its slabs back in stock in the Excel file.",
      input_schema: { type: "object", required: ["order"], properties: { order: { type: "string", description: "e.g. SO-261013-01" } } } },
    { name: "find_orders", description: "List recent sales orders, optionally for one client or one order number.",
      input_schema: { type: "object", properties: { client: { type: "string" }, order: { type: "string" } } } },
  ];
  function pushAct(a) { const m = S.chat[S.chat.length - 1]; if (m && m.role === "assistant") { (m.acts = m.acts || []).push({ ...a, undone: false }); renderMsgs(); } }
  const EXEC = {
    async create_sales_order(input) {
      const o = await createOrder({ client: input.client, phone: input.client_phone, lines: input.lines, notes: input.notes, via: "agent" });
      const slabs = o.rows.reduce((t, r) => t + r.qty, 0), total = o.rows.reduce((t, r) => t + (r.total || 0), 0);
      pushAct({ kind: "order", order: o.order, label: `Sales order ${o.order} · ${o.client} · ${slabs} slab${slabs === 1 ? "" : "s"}` });
      return { order_number: o.order, client: o.client, saved_to_onedrive: o.saved,
        lines: o.rows.map((r) => ({ lot: r.lot, color: r.color, size_in: r.size, slabs: r.qty, sqft: r.sqft, price_per_slab: r.price, amount: r.total, excel_cell: r.cell })),
        total_slabs: slabs, total_sqft: Math.round(o.rows.reduce((t, r) => t + (r.sqft || 0), 0) * 10) / 10, total_amount: total || null,
        ...(o.held ? { warning: `${o.held} slab(s) came from held stock` } : {}),
        ...(!o.saved && S.mode !== "demo" ? { warning_file: "Stock was updated, but the order file couldn't be saved in OneDrive (no permission to the folder)." } : {}) };
    },
    async cancel_sales_order(input) {
      const r = await cancelOrder(String(input.order || "").trim().toUpperCase(), "agent");
      S.chat.forEach((c) => (c.acts || []).forEach((a) => { if (a.kind === "order" && a.order === r.cancelled) a.undone = true; }));
      return r;
    },
    async find_orders(input) {
      try { S.orderRows = await S.store.orders.read(); } catch {}
      const c = String(input.client || "").toLowerCase(), n = String(input.order || "").toUpperCase();
      return ordersFromRows(S.orderRows).filter((o) => (!c || String(o.client).toLowerCase().includes(c)) && (!n || o.order === n)).slice(0, 15)
        .map((o) => ({ order: o.order, date: String(o.date).slice(0, 10), client: o.client, status: o.status, lines: o.lines.map((l) => ({ lot: l.lot, color: l.color, slabs: l.qty, amount: l.total })) }));
    },
    find_lots(input) {
      const code = input.lot ? normCode(input.lot) : "", ts = toks(input.text), only = input.in_stock_only !== false;
      let hits = S.lots.filter((l) => (!code || normCode(l.lot) === code) && matches(l, ts));
      if (code && !hits.length) hits = S.lots.filter((l) => normCode(l.lot).includes(code) && matches(l, ts));
      const shown = hits.filter((l) => !only || l.stock > 0).sort((a, b) => freeOf(b) - freeOf(a));
      return { lots_matched: hits.length, lots_in_stock: hits.filter((l) => l.stock > 0).length, total_in_stock: hits.reduce((t, l) => t + Math.max(0, l.stock), 0),
        total_free: hits.reduce((t, l) => t + freeOf(l), 0), showing: Math.min(shown.length, 40), lots: shown.slice(0, 40).map(lotView) };
    },
    async record_sale(input) {
      const l = S.lots.find((x) => x.id === String(input.lot_id)); if (!l) throw new Error(`No lot with id ${input.lot_id}. Use ids from find_lots.`);
      const qty = Math.floor(Number(input.qty)); if (!(qty > 0)) throw new Error("qty must be 1 or more.");
      const sign = input.type === "return" ? -1 : 1;
      const res = await recordSale(l, sign * qty, input.note, "agent");
      const label = `${sign > 0 ? qty + " sold" : qty + " returned"} · ${lotName(l)} · cell ${res.change.ref}`;
      pushAct({ kind: "sale", lotId: l.id, lot: l.lot, name: l.name, change: res.change, label });
      return { recorded_in_excel: true, cell: res.change.ref, lot: lotName(l), qty, type: sign > 0 ? "sale" : "return", in_stock_now: l.stock, free_now: freeOf(l),
        ...(res.heldInvolved && sign > 0 ? { warning: `${res.heldInvolved} of these came from held slabs` } : {}) };
    },
    async set_hold(input) {
      const people = S.meta.holdPeople || [];
      const key = (x) => String(x).toUpperCase().replace(/[^A-Z]/g, ""), want = key(input.person || "");
      const p = people.find((x) => key(x) === want) || people.filter((x) => key(x).startsWith(want)).sort((a, b) => a.length - b.length)[0];
      if (!want || !p) throw new Error(`Unknown person "${input.person}". Hold names: ${people.join(", ")}.`);
      const l = S.lots.find((x) => x.id === String(input.lot_id)); if (!l) throw new Error(`No lot with id ${input.lot_id}.`);
      const qty = Math.max(0, Math.floor(Number(input.qty) || 0));
      const res = await setHold(l, p, qty, "agent");
      pushAct({ kind: "hold", lotId: l.id, lot: l.lot, name: l.name, change: res.change, label: `Hold ${p}: ${res.change.prev} → ${qty} · ${lotName(l)}` });
      return { saved_in_excel: true, person: p, held_for_person: qty, on_hold_total: holdOf(l), free_now: freeOf(l) };
    },
  };

  // Talk to Claude: through the viewer's Claude account in a Claude preview, otherwise through the site's /api/agent function.
  async function askClaude(turns, onText, signal) {
    if (S.sample) {
      const tools = TOOL_DEFS.map((t) => ({ name: t.name, description: t.description, inputSchema: t.input_schema, execute: (inp) => EXEC[t.name](inp) }));
      const r = await S.sample([{ role: "user", content: rules() }, ...turns], { modelTier: "quick", tools, signal, onText: ({ text }) => onText(text) });
      return r.text;
    }
    if (S.mode !== "live" && S.mode !== "view") throw new Error("The agent needs the site to be connected to Microsoft (see README), or open the preview in Claude.");
    const tools = viewOnly() ? TOOL_DEFS.filter((t) => t.name === "find_lots") : TOOL_DEFS;
    const messages = turns.map((t) => ({ role: t.role, content: t.content }));
    let last = "";
    for (let round = 0; round < 6; round++) {
      const token = S.auth?.signedIn?.() ? await S.auth.getToken() : null;
      const r = await fetch(CFG.agentUrl || "/api/agent", { method: "POST", signal,
        headers: { "content-type": "application/json", ...(token ? { authorization: "Bearer " + token } : {}) }, body: JSON.stringify({ system: viewOnly() ? "" : rules(), messages, tools }) });
      const d = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(d.error || `Agent error ${r.status}`);
      const said = (d.content || []).filter((b) => b.type === "text").map((b) => b.text).join("").trim();
      if (said) { last = said; onText(said); }
      const uses = (d.content || []).filter((b) => b.type === "tool_use");
      if (d.stop_reason !== "tool_use" || !uses.length) break;
      messages.push({ role: "assistant", content: d.content });
      const results = [];
      for (const u of uses) {
        try { if (viewOnly() && u.name !== "find_lots") throw new Error("read-only"); results.push({ type: "tool_result", tool_use_id: u.id, content: JSON.stringify(await EXEC[u.name](u.input || {})) }); }
        catch (e) { results.push({ type: "tool_result", tool_use_id: u.id, content: "Error: " + e.message, is_error: true }); }
      }
      messages.push({ role: "user", content: results });
    }
    return last || "Done.";
  }

  function fmtReply(t) { return esc(t).replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>").replace(/\n/g, "<br>"); }
  function renderMsgs() {
    const box = $("msgs");
    if (!S.chat.length) {
      const top = S.lots.filter((l) => l.lot && !l.outside && freeOf(l) > 5).sort((a, b) => freeOf(b) - freeOf(a));
      const a = top[0], b = top.find((l) => l.name !== a?.name) || top[1], person = (S.meta?.holdPeople || [])[4] || (S.meta?.holdPeople || [])[0] || "Puru";
      const ask = a ? [`Is ${a.lot} available?`, `Do we have ${b ? b.name.replace(/[-(].*$/, "").trim().toLowerCase() : "calacatta"} in jumbo size?`, `Sold 2 slabs of ${a.lot}`, `Hold 3 ${b?.lot || a.lot} for ${person}`, `Create a sales order for ABC Granite: 2 slabs of ${a.lot} at $450 each`] : [];
      const can = S.sample || S.mode === "live" || viewOnly();
      if (viewOnly()) ask.splice(2);
      box.innerHTML = `<div class="starters"><p>${viewOnly() ? "Ask what's in stock, in English or Spanish. For example, a lot number, a color or a size." : can ? "Ask about stock, tell me what sold, or put slabs on hold. I write every change into the Excel file." : "The agent needs the site connected to Microsoft, or open this preview in Claude. You can still record sales and holds in the lot list."}</p>
        ${can ? ask.map((x) => `<button class="chip" type="button" data-ask="${esc(x)}">${esc(x)}</button>`).join("") : ""}</div>`;
      return;
    }
    box.innerHTML = S.chat.map((m, i) => {
      if (m.role === "user") return `<div class="msg user"><div class="b">${esc(m.text)}</div></div>`;
      const acts = (m.acts || []).map((a, j) => `<div class="act ${a.undone ? "undone" : ""}"><span class="pill ${a.kind === "hold" ? "hold" : "out"}">${a.kind === "order" ? "Order" : a.kind === "sale" ? "Excel" : "Hold"}</span><span class="lbl">${esc(a.label)}</span>${a.kind === "order" ? `<button class="link" type="button" data-view-order="${esc(a.order)}">View</button>` : ""}<button class="link" type="button" data-undo="${i}:${j}" ${a.undone ? "disabled" : ""}>${a.undone ? (a.kind === "order" ? "Cancelled" : "Undone") : (a.kind === "order" ? "Cancel order" : "Undo")}</button></div>`).join("");
      return `<div class="msg agent ${m.err ? "err" : ""}"><div class="b ${m.thinking && !m.text ? "thinking" : ""}">${m.thinking && !m.text ? "Checking the Excel file…" : fmtReply(m.text)}</div>${acts}</div>`;
    }).join("");
    box.scrollTop = box.scrollHeight;
  }
  const ERR = { not_granted: "Claude isn't allowed on this page for you, so the agent can't answer. You can still record sales in the lot list.",
    rate_limited: "Too many requests right now. Wait a moment and send again.", tools_unavailable: "This view can't run the agent's tools. Try the Claude website or desktop app.",
    refused: "The agent couldn't answer that. Try rephrasing.", prompt_too_large: "That conversation got too long. Start a new chat." };
  async function send(text) {
    text = text.trim(); if (!text || S.busy || !S.meta) return;
    S.busy = true; $("sendBtn").textContent = "Stop"; $("ask").value = "";
    S.chat.push({ role: "user", text });
    const turns = S.chat.slice(-12).filter((m) => !m.err).map((m) => ({ role: m.role, content: m.text }));
    const reply = { role: "assistant", text: "", thinking: true, acts: [] };
    S.chat.push(reply); renderMsgs();
    S.ctl = new AbortController();
    try { reply.text = await askClaude(turns, (t) => { reply.text = t; reply.thinking = false; renderMsgs(); }, S.ctl.signal); }
    catch (e) {
      if (e?.code === "cancelled" || e?.name === "AbortError") reply.text = (reply.text ? reply.text + "\n" : "") + "(stopped)";
      else { const m = ERR[e?.code] || e?.message || "The agent was interrupted. Send again."; reply.text = e?.text ? e.text + "\n\n" + m : (reply.text ? reply.text + "\n\n" + m : m); reply.err = !e?.text && !reply.acts.length; }
    } finally { reply.thinking = false; S.busy = false; $("sendBtn").textContent = "Send"; renderMsgs(); if (!reply.err) Voice.say(reply.text); }
  }

  /* ---------- voice: speak to the agent (English / Spanish) and optional read-aloud ---------- */
  const Voice = (() => {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    const LANGS = { en: { code: "en-US", label: "EN", ph: "e.g. Is IQ466 available?" }, es: { code: "es-US", label: "ES", ph: "p. ej. ¿Hay IQ466 disponible?" } };
    const store = (k, v) => { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch { return null; } };
    let lang = store("sf_voice_lang") === "es" ? "es" : "en", speak = store("sf_voice_speak") === "1", rec = null, base = "";
    const mic = $("micBtn"), spk = $("spkBtn"), lb = $("langBtn"), ask = $("ask");
    function paint() {
      lb.textContent = LANGS[lang].label; lb.title = lang === "en" ? "Voice language: English (tap for Español)" : "Idioma de voz: Español (toca para English)";
      ask.placeholder = LANGS[lang].ph; spk.setAttribute("aria-pressed", String(speak));
      mic.setAttribute("aria-pressed", String(!!rec)); mic.title = rec ? (lang === "es" ? "Escuchando… toca para parar" : "Listening… tap to stop") : (lang === "es" ? "Habla con el agente" : "Speak to the agent");
    }
    function stop() { try { rec?.stop(); } catch {} }
    function start() {
      if (rec) return stop();
      window.speechSynthesis?.cancel();
      const r = new SR(); rec = r; base = ask.value ? ask.value.trim() + " " : "";
      r.lang = LANGS[lang].code; r.interimResults = true; r.continuous = false; r.maxAlternatives = 1;
      let finalText = "";
      r.onresult = (e) => {
        let interim = ""; finalText = "";
        for (const res of e.results) (res.isFinal ? (finalText += res[0].transcript) : (interim += res[0].transcript));
        ask.value = base + (finalText || interim);
      };
      r.onerror = (e) => {
        const m = { "not-allowed": "Microphone blocked. Allow the microphone for this site (lock icon in the address bar).", "service-not-allowed": "Microphone blocked. Allow the microphone for this site (lock icon in the address bar).",
          "no-speech": lang === "es" ? "No te escuché. Toca el micrófono e intenta de nuevo." : "Didn't hear anything. Tap the mic and try again.", "audio-capture": "No microphone found.", network: "Voice needs an internet connection." }[e.error];
        if (m) flash(m);
      };
      r.onend = () => { rec = null; paint(); const t = ask.value.trim(); if (finalText.trim() && t && !S.busy) send(t); else ask.focus(); };
      try { r.start(); } catch { rec = null; }
      paint();
    }
    function say(text) {
      if (!speak || !text || !window.speechSynthesis) return;
      const plain = String(text).replace(/\*\*/g, "").replace(/•/g, ",").replace(/\s+x\s+/gi, " by ").replace(/\n+/g, ". ");
      const u = new SpeechSynthesisUtterance(plain);
      u.lang = /[¿¡ñáéíóú]|\b(hay|disponible|losas|sí|está|tenemos)\b/i.test(plain) ? "es-US" : "en-US";
      speechSynthesis.cancel(); speechSynthesis.speak(u);
    }
    lb.onclick = () => { lang = lang === "en" ? "es" : "en"; store("sf_voice_lang", lang); if (rec) { stop(); } paint(); };
    spk.onclick = () => { speak = !speak; store("sf_voice_speak", speak ? "1" : "0"); if (!speak) window.speechSynthesis?.cancel(); paint(); };
    if (SR) { mic.hidden = false; mic.onclick = start; } else { lb.title = "Voice typing isn't supported in this browser. Use Chrome, Edge or Safari."; }
    if (window.speechSynthesis) spk.hidden = false;
    paint();
    return { say };
  })();

  /* ---------- events ---------- */
  $("composer").onsubmit = (e) => { e.preventDefault(); if (S.busy) { S.ctl?.abort(); return; } send($("ask").value); };
  $("clearChat").onclick = () => { if (S.busy) S.ctl?.abort(); S.chat = []; renderMsgs(); };
  $("msgs").addEventListener("click", async (e) => {
    const chip = e.target.closest("[data-ask]"); if (chip) return send(chip.dataset.ask);
    const u = e.target.closest("[data-undo]"); if (!u) return;
    const [i, j] = u.dataset.undo.split(":").map(Number), a = S.chat[i]?.acts?.[j];
    if (!a || a.undone) return; u.disabled = true;
    try { await undo(a); a.undone = true; renderMsgs(); } catch (err) { u.disabled = false; flash(err.message); }
  });
  const openLot = (id) => { S.sel = S.sel === id ? null : id; renderCard(); renderTable(); if (S.sel) $("lotCard").scrollIntoView({ block: "nearest" }); };
  $("tablebox").addEventListener("click", (e) => { const r = e.target.closest("tr[data-id]"); if (r) openLot(r.dataset.id); });
  $("tablebox").addEventListener("keydown", (e) => { const r = e.target.closest("tr[data-id]"); if (r && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); openLot(r.dataset.id); } });
  $("search").addEventListener("input", (e) => { S.q = e.target.value; renderTable(); });
  $("secSel").addEventListener("change", (e) => { S.sec = e.target.value; renderTable(); });
  document.querySelectorAll(".seg button").forEach((b) => (b.onclick = () => { S.filter = b.dataset.f; document.querySelectorAll(".seg button").forEach((x) => x.setAttribute("aria-pressed", String(x === b))); renderTable(); }));
  const setTab = (t) => { $("desk").dataset.tab = t; $("tabAgent").setAttribute("aria-selected", String(t === "agent")); $("tabInv").setAttribute("aria-selected", String(t === "inventory")); };
  $("tabAgent").onclick = () => setTab("agent"); $("tabInv").onclick = () => setTab("inventory");
  $("refreshBtn").onclick = () => load();
  $("sheetSel").onchange = (e) => load(e.target.value);
  $("signOut").onclick = () => { S.auth?.signOut(); location.reload(); };
  $("signIn").onclick = () => S.auth?.signIn();
  $("demoDownload").onclick = async () => {
    const bytes = S.store?.drive?.bytes?.(); if (!bytes) return;
    const name = `LIVE_INVENTORY-demo-${new Date().toISOString().slice(0, 10)}.xlsx`;
    if (S.downloads) { try { await S.downloads.save({ filename: name, data: new Blob([bytes]) }); } catch (e) { if (e?.code !== "declined") flash("The file couldn't be saved here."); } return; }
    const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([bytes])); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  };

  document.addEventListener("click", (e) => { const v = e.target.closest("[data-view-order]"); if (v) showOrder(v.dataset.viewOrder); });
  window.Portal = { open };

  // Coming back from Microsoft sign-in: finish it, then open the portal.
  (async () => {
    if (CFG.microsoft?.clientId && /[?&](code|error)=/.test(location.search)) {
      const auth = window.MsAuth.create({ clientId: CFG.microsoft.clientId, tenant: CFG.microsoft.tenant });
      try { await auth.handleRedirect(); } catch (e) { S.pendingErr = e.message; }
      if (location.hash !== "#staff") location.hash = "#staff"; else window.dispatchEvent(new HashChangeEvent("hashchange"));
      return;
    }
    if (location.hash === "#staff") open();
  })();
})();

// Public site: generated stone swatches, collections, page routing, contact form.
(() => {
  const CFG = window.STONEFLOW_CONFIG || {};
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  if (CFG.company) { $("brandName").textContent = CFG.company; $("footName").textContent = CFG.company; document.title = CFG.company; }

  /* ---------- stone renderer (seeded, so every visit draws the same slab) ---------- */
  function rng(seed) { let s = seed >>> 0; return () => { s = (s + 0x6D2B79F5) >>> 0; let t = s; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }
  const STONES = {
    calacattaGold: { base: ["#F7F5F0", "#ECE9E2"], veins: ["rgba(150,120,70,.55)", "rgba(110,105,100,.45)"], n: 5, w: 7, speck: 0 },
    calacattaGrey: { base: ["#F6F6F4", "#E7E8E6"], veins: ["rgba(95,100,105,.5)", "rgba(140,140,140,.35)"], n: 6, w: 6, speck: 0 },
    carrara: { base: ["#E9EBEC", "#DADEE0"], veins: ["rgba(110,118,125,.35)"], n: 14, w: 2.5, speck: 0 },
    printed: { base: ["#F4F1EC", "#E6E1D9"], veins: ["rgba(120,95,70,.6)", "rgba(60,60,60,.45)"], n: 4, w: 11, speck: 0 },
    granite: { base: ["#26272A", "#1A1B1D"], veins: [], n: 0, w: 0, speck: 2600, speckColors: ["#5A5C61", "#8C8E92", "#3B3D42", "#A7A096"] },
    onyx: { base: ["#E3C48E", "#C99A57"], veins: ["rgba(255,245,220,.55)", "rgba(150,95,40,.35)"], n: 18, w: 5, speck: 0, bands: true },
  };
  function drawStone(canvas, kind, seed) {
    const st = STONES[kind], ctx = canvas.getContext("2d"), W = canvas.width, H = canvas.height, R = rng(seed);
    const g = ctx.createLinearGradient(0, 0, W, H); g.addColorStop(0, st.base[0]); g.addColorStop(1, st.base[1]);
    ctx.fillStyle = g; ctx.fillRect(0, 0, W, H);
    if (st.speck) {
      for (let i = 0; i < st.speck * (W * H) / (1016 * 512); i++) {
        ctx.fillStyle = st.speckColors[Math.floor(R() * st.speckColors.length)];
        ctx.globalAlpha = .35 + R() * .5; const r = .6 + R() * 2.6;
        ctx.beginPath(); ctx.arc(R() * W, R() * H, r, 0, Math.PI * 2); ctx.fill();
      }
      ctx.globalAlpha = 1; return;
    }
    const vein = (color, width, blur) => {
      let x = st.bands ? -20 : R() * W * .3 - W * .1, y = st.bands ? R() * H : R() * H;
      const ang0 = st.bands ? (R() - .5) * .25 : -0.35 + (R() - .5) * .9;
      ctx.save(); ctx.filter = blur ? `blur(${blur}px)` : "none"; ctx.strokeStyle = color; ctx.lineCap = "round";
      let ang = ang0, w = width * (.5 + R());
      ctx.beginPath(); ctx.moveTo(x, y);
      for (let i = 0; i < 220 && x < W + 40; i++) {
        ang += (R() - .5) * .35; ang = ang0 + (ang - ang0) * .92;
        x += Math.cos(ang) * 9; y += Math.sin(ang) * 9 + (R() - .5) * 3;
        ctx.lineTo(x, y);
        if (!blur && i % 6 === 5) { ctx.lineWidth = w; ctx.stroke(); ctx.beginPath(); ctx.moveTo(x, y); w = Math.max(.4, w * (.75 + R() * .5)); }
      }
      ctx.lineWidth = w; ctx.stroke(); ctx.restore();
    };
    for (let i = 0; i < st.n; i++) { const c = st.veins[i % st.veins.length]; vein(c, st.w * 2.2, 6); vein(c, st.w * .5, 0); }
    for (let i = 0; i < st.n * 3; i++) vein(st.veins[0].replace(/[\d.]+\)$/, ".18)"), 1, 0);
  }

  const COLLECTIONS = [
    { kind: "granite", title: "Granite", spec: "Up to 135 × 79 in · polished and leathered", names: "Black Pearl, Brown Fantasy, Steel Grey" },
    { kind: "calacattaGold", title: "Calacatta Jumbo", spec: "130–138 × 79 in · jumbo and super jumbo", names: "Arabia Grey, Arctic Grey, Zircon, Zoya" },
    { kind: "calacattaGrey", title: "Calacatta Regular", spec: "127 × 64 in · many bookmatched", names: "Arabescato, Astro, Sienna, Zircon" },
    { kind: "carrara", title: "Carrara & Basics", spec: "127 × 64 and jumbo · 2 cm and 3 cm", names: "Bologna, Bianco Pepper, Pure White, Pure Black" },
    { kind: "printed", title: "Printed Quartz", spec: "126–131 in · printed full body", names: "Aries, Bion, Clarina, Taj Mahal" },
    { kind: "onyx", title: "Quartz Onyx", spec: "107 × 72 in · 1.5 cm and 2.5 cm", names: "Champagne Onyx" },
  ];
  $("collGrid").innerHTML = COLLECTIONS.map((c, i) => `<article><div class="sw"><canvas width="640" height="360" data-kind="${c.kind}" data-seed="${101 + i * 17}" aria-hidden="true"></canvas></div>
    <div class="body"><h3>${esc(c.title)}</h3><span class="spec">${esc(c.spec)}</span><span class="names">${esc(c.names)}</span></div></article>`).join("");
  document.querySelectorAll("canvas[data-kind]").forEach((cv) => drawStone(cv, cv.dataset.kind, Number(cv.dataset.seed)));
  drawStone($("heroSlab"), "calacattaGold", 7);

  /* ---------- routing: #staff opens the portal, everything else is the public site ---------- */
  function route() {
    const staff = location.hash === "#staff";
    $("home-view").hidden = staff; $("staff-view").hidden = !staff;
    $("staffLink").textContent = staff ? "Back to site" : "Live Inventory";
    $("staffLink").setAttribute("href", staff ? "#home" : "#staff");
    if (staff) { window.Portal?.open(); window.scrollTo(0, 0); }
    else if (location.hash && location.hash !== "#home") document.querySelector(location.hash)?.scrollIntoView();
  }
  window.addEventListener("hashchange", route);
  route();

  /* ---------- contact form (Netlify Forms when deployed) ---------- */
  $("contactForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const st = $("cStatus");
    if (window.claude) { st.textContent = "Demo site: messages aren't sent from this preview."; return; }
    st.textContent = "Sending…";
    try {
      const r = await fetch("/", { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams(new FormData(e.target)).toString() });
      if (!r.ok) throw new Error();
      e.target.reset(); st.textContent = "Sent. We'll call you back today.";
    } catch { st.textContent = "Couldn't send. Please call us instead."; }
  });
})();

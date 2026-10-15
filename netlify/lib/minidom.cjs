// A tiny XML DOM, just enough for site/js/xlsxlite.js to read .xlsx files on the server
// (Node has no DOMParser). Supports elements, attributes, text, entities, getElementsByTagName,
// childNodes, localName, getAttribute(NS), textContent and closest(tag).
const ENT = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };
const decode = (s) => s.indexOf("&") < 0 ? s : s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (m, e) =>
  e[0] === "#" ? String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : (ENT[e] ?? m));

class Text { constructor(t, parent) { this.nodeType = 3; this.data = t; this.parentNode = parent; } get textContent() { return this.data; } }
class El {
  constructor(name, attrs, parent) {
    this.nodeType = 1; this.tagName = name; this.localName = name.includes(":") ? name.split(":").pop() : name;
    this.attrs = attrs; this.childNodes = []; this.parentNode = parent;
  }
  getAttribute(n) { return Object.prototype.hasOwnProperty.call(this.attrs, n) ? this.attrs[n] : null; }
  getAttributeNS(ns, local) {
    for (const [k, v] of Object.entries(this.attrs)) {
      if (!k.includes(":")) continue;
      const [pre, loc] = k.split(":");
      if (loc === local && this.lookupNS(pre) === ns) return v;
    }
    return null;
  }
  lookupNS(pre) { for (let e = this; e && e.nodeType === 1; e = e.parentNode) { const v = e.attrs["xmlns:" + pre]; if (v != null) return v; } return null; }
  getElementsByTagName(name) {
    const out = [], stack = [...this.childNodes].reverse();
    while (stack.length) {
      const n = stack.pop(); if (n.nodeType !== 1) continue;
      if (name === "*" || n.tagName === name || n.localName === name) out.push(n);
      for (let i = n.childNodes.length - 1; i >= 0; i--) stack.push(n.childNodes[i]);
    }
    return out;
  }
  get textContent() { let s = ""; const st = [...this.childNodes].reverse(); while (st.length) { const n = st.pop(); if (n.nodeType === 3) s += n.data; else for (let i = n.childNodes.length - 1; i >= 0; i--) st.push(n.childNodes[i]); } return s; }
  closest(tag) { for (let e = this; e && e.nodeType === 1; e = e.parentNode) if (e.tagName === tag || e.localName === tag) return e; return null; }
}

function parse(xml) {
  const doc = new El("#document", {}, null); doc.nodeType = 9;
  let cur = doc, i = 0;
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<!DOCTYPE[^>]*>|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
  const attrRe = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m.index > i) { const t = xml.slice(i, m.index); if (cur !== doc) cur.childNodes.push(new Text(decode(t), cur)); }
    i = re.lastIndex;
    if (m[1] != null) cur.childNodes.push(new Text(m[1], cur));
    else if (m[2]) { if (cur.parentNode) cur = cur.parentNode; }
    else if (m[3]) {
      const attrs = {}; let a; attrRe.lastIndex = 0;
      while ((a = attrRe.exec(m[4] || ""))) attrs[a[1]] = decode(a[2] ?? a[3] ?? "");
      const el = new El(m[3], attrs, cur); cur.childNodes.push(el);
      if (!m[5]) cur = el;
    }
  }
  doc.documentElement = doc.childNodes.find((n) => n.nodeType === 1) || null;
  return doc;
}

class DOMParser { parseFromString(s) { return parse(String(s)); } }
module.exports = { DOMParser };

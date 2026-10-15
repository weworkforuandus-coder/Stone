// /api/agent: the only place the Claude API key lives.
// Editors (OWNER_EMAIL + ALLOWED_EMAILS, signed in with Microsoft) get every inventory tool.
// Everyone else, signed in or not, gets a read-only agent: it can only look up stock, uses a fixed
// set of instructions written here on the server, and is rate limited.
// Environment variables:
//   ANTHROPIC_API_KEY   your Claude API key
//   ALLOWED_EMAILS      staff who may change the inventory
//   PUBLIC_AGENT        optional: "off" turns the read-only agent off for visitors
//   CLAUDE_MODEL        optional, default claude-haiku-4-5-20251001
import { json, caller, rateLimited, ipOf, env } from "../lib/common.mjs";

const EDIT_TOOLS = new Set(["find_lots", "record_sale", "set_hold", "create_sales_order", "cancel_sales_order", "find_orders"]);

const VIEWER_RULES = `You are the stock assistant on a stone slab yard's website. Visitors ask whether slabs are available.
You can ONLY look up stock with find_lots. You cannot sell, hold, reserve or price anything, and you have no information about prices, customers, orders or yard locations.
Always call find_lots before answering about stock; never guess. Search by lot number (like IQ466) or by color words and size (like "super white" or "127x64").
Answer with the total slabs free to sell, then each matching lot on its own line: lot #, color, size, slabs free. If several colors match, group them. If nothing is free, say it isn't available right now.
Section names tell the type (GRANITE, CALACATTA / CARRARA quartz, JUMBO vs REGULAR, PRINTED, 2CM, ONYX). In color names -J = jumbo, -SJ = super jumbo, -R = regular, PFB = printed full body. Size is length x height in inches.
To buy, reserve or get a price, tell them to contact the yard using the Visit section of the website or talk to a salesperson.
Only discuss the yard's slab stock. Politely decline anything else.
Visitors may write or speak in English or Spanish (spoken messages can have small transcription mistakes, e.g. "I Q 466" means IQ466). Always reply in the language of the visitor's latest message. Keep lot numbers, color names and sizes exactly as in the inventory.
Reply briefly: short lines, "•" for list items, **bold** only for the key number. No tables or headings. Never invent lots, sizes or counts.`;

export default async (req, context) => {
  if (req.method !== "POST") return json({ error: "Use POST." }, 405);
  const key = env("ANTHROPIC_API_KEY");
  if (!key) return json({ error: "The agent isn't set up: add ANTHROPIC_API_KEY in Netlify." }, 500);

  const user = await caller(req).catch(() => null);
  const editor = user?.role === "editor";
  if (!editor) {
    if (env("PUBLIC_AGENT").toLowerCase() === "off") return json({ error: "The stock assistant is for staff only. Use the lot list, or contact us." }, 403);
    if (rateLimited("agent:" + (user?.email || ipOf(req, context)), 30, 10 * 60_000)) return json({ error: "Too many questions in a short time. Please wait a few minutes." }, 429);
  }

  let body;
  try { body = await req.json(); } catch { return json({ error: "Bad request." }, 400); }
  const { system, messages, tools } = body || {};
  if (!Array.isArray(messages) || !messages.length || messages.length > (editor ? 60 : 24)) return json({ error: "Conversation too long. Start a new chat." }, 400);
  if (JSON.stringify(messages).length > (editor ? 200000 : 60000)) return json({ error: "Conversation too long. Start a new chat." }, 400);
  if (!Array.isArray(tools)) return json({ error: "Bad request." }, 400);

  let useSystem, useTools;
  if (editor) {
    if (typeof system !== "string" || system.length > 20000) return json({ error: "Bad request." }, 400);
    if (tools.some((t) => !EDIT_TOOLS.has(t?.name))) return json({ error: "Bad request." }, 400);
    useSystem = system; useTools = tools;
  } else {
    // Visitors: only the lookup tool, and only our own instructions.
    useTools = tools.filter((t) => t?.name === "find_lots").slice(0, 1);
    if (!useTools.length) return json({ error: "Bad request." }, 400);
    const usedOther = messages.some((m) => Array.isArray(m.content) && m.content.some((b) => b?.type === "tool_use" && b.name !== "find_lots"));
    if (usedOther) return json({ error: "Bad request." }, 400);
    useSystem = VIEWER_RULES + `\nToday is ${new Date().toDateString()}.`;
  }

  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
    body: JSON.stringify({ model: env("CLAUDE_MODEL") || "claude-haiku-4-5-20251001", max_tokens: editor ? 1024 : 700, system: useSystem, messages, tools: useTools }),
  });
  const d = await r.json().catch(() => ({}));
  if (!r.ok) return json({ error: r.status === 429 ? "The agent is busy. Try again in a moment." : `Claude error ${r.status}: ${d?.error?.message || "unknown"}` }, 502);
  return json({ content: d.content || [], stop_reason: d.stop_reason || "end_turn", role: editor ? "editor" : "viewer" });
};

export const config = { path: "/api/agent" };

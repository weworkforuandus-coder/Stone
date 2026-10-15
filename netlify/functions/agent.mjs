// /api/agent: the only place the Claude API key lives.
// Checks the caller's Microsoft sign-in, allows only the emails in ALLOWED_EMAILS,
// then forwards one round of the agent conversation to Claude.
// Environment variables (Netlify > Site configuration > Environment variables):
//   ANTHROPIC_API_KEY   your Claude API key
//   ALLOWED_EMAILS      comma-separated staff emails, e.g. "you@gmail.com, ravi@yourco.com"
//   CLAUDE_MODEL        optional, default claude-haiku-4-5-20251001

const TOOL_NAMES = new Set(["find_lots", "record_sale", "set_hold", "create_sales_order", "cancel_sales_order", "find_orders"]);
const json = (data, status = 200) => new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json" } });

export default async (req) => {
  if (req.method !== "POST") return json({ error: "Use POST." }, 405);
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return json({ error: "The agent isn't set up: add ANTHROPIC_API_KEY in Netlify." }, 500);

  // 1. Who is calling? Ask Microsoft with the caller's own token.
  const auth = req.headers.get("authorization") || "";
  if (!/^Bearer\s+\S+/.test(auth)) return json({ error: "Sign in with Microsoft first." }, 401);
  const me = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName", { headers: { authorization: auth } });
  if (!me.ok) return json({ error: "Your Microsoft sign-in expired. Sign in again." }, 401);
  const u = await me.json();
  const emails = [u.mail, u.userPrincipalName].filter(Boolean).map((e) => String(e).toLowerCase());
  const allowed = String(process.env.ALLOWED_EMAILS || "").toLowerCase().split(/[\s,;]+/).filter(Boolean);
  if (!allowed.length) return json({ error: "No staff are allowed yet: set ALLOWED_EMAILS in Netlify." }, 403);
  if (!emails.some((e) => allowed.includes(e))) return json({ error: `${emails[0] || "This account"} isn't on the staff list. Ask the owner to add it.` }, 403);

  // 2. Check the request is an inventory-agent turn, nothing else.
  let body;
  try { body = await req.json(); } catch { return json({ error: "Bad request." }, 400); }
  const { system, messages, tools } = body || {};
  if (typeof system !== "string" || system.length > 20000) return json({ error: "Bad request." }, 400);
  if (!Array.isArray(messages) || !messages.length || messages.length > 60) return json({ error: "Conversation too long. Start a new chat." }, 400);
  if (!Array.isArray(tools) || tools.some((t) => !TOOL_NAMES.has(t?.name))) return json({ error: "Bad request." }, 400);
  if (JSON.stringify(messages).length > 200000) return json({ error: "Conversation too long. Start a new chat." }, 400);

  // 3. One Claude call. The browser runs the tools against the Excel file and sends the results back.
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
    body: JSON.stringify({ model: process.env.CLAUDE_MODEL || "claude-haiku-4-5-20251001", max_tokens: 1024, system, messages, tools }),
  });
  const d = await r.json().catch(() => ({}));
  if (!r.ok) return json({ error: r.status === 429 ? "The agent is busy. Try again in a moment." : `Claude error ${r.status}: ${d?.error?.message || "unknown"}` }, 502);
  return json({ content: d.content || [], stop_reason: d.stop_reason || "end_turn" });
};

export const config = { path: "/api/agent" };

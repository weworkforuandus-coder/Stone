// /api/me: who is this visitor (editor / viewer / not signed in), and is OneDrive connected?
import { json, caller, kvGet } from "../lib/common.mjs";

export default async (req) => {
  const u = await caller(req).catch(() => null);
  const owner = await kvGet("owner").catch(() => null);
  return json({ signedIn: !!u, name: u?.name || "", email: u?.email || "", role: u?.role || "viewer", isOwner: !!u?.owner, connected: !!owner?.refresh_token });
};

export const config = { path: "/api/me" };

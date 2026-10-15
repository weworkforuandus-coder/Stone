/*
 * Microsoft sign-in for a single-page site (OAuth 2.0 authorization code + PKCE). No libraries.
 * Works for work/school (Microsoft 365) and personal Microsoft accounts.
 * Register the site's address as a "Single-page application" redirect URI in Azure (see README).
 */
(function (root) {
  const KEY = "sf_ms_tokens", TMP = "sf_ms_pkce";
  const SCOPES = "openid profile offline_access User.Read Files.ReadWrite.All";
  const b64url = (bytes) => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
    del(k) { try { localStorage.removeItem(k); } catch {} },
  };

  function create(cfg) {
    const tenant = cfg.tenant || "common";
    const base = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0`;
    const redirectUri = cfg.redirectUri || (location.origin + location.pathname);

    async function signIn() {
      const verifier = b64url(crypto.getRandomValues(new Uint8Array(48)));
      const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
      const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
      try { sessionStorage.setItem(TMP, JSON.stringify({ verifier, state, hash: location.hash })); } catch {}
      const q = new URLSearchParams({
        client_id: cfg.clientId, response_type: "code", redirect_uri: redirectUri, response_mode: "query",
        scope: SCOPES, state, code_challenge: challenge, code_challenge_method: "S256", prompt: "select_account",
      });
      location.assign(`${base}/authorize?${q}`);
    }

    async function tokenRequest(params) {
      const r = await fetch(`${base}/token`, {
        method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ client_id: cfg.clientId, scope: SCOPES, ...params }),
      });
      const d = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(d.error_description ? d.error_description.split("\r\n")[0] : "Microsoft sign-in failed.");
      const t = { access: d.access_token, refresh: d.refresh_token, exp: Date.now() + (d.expires_in - 60) * 1000 };
      store.set(KEY, t);
      return t;
    }

    // Call once on page load: finishes a sign-in that just came back from Microsoft.
    async function handleRedirect() {
      const p = new URLSearchParams(location.search);
      if (!p.has("code") && !p.has("error")) return null;
      let tmp = null; try { tmp = JSON.parse(sessionStorage.getItem(TMP) || "null"); sessionStorage.removeItem(TMP); } catch {}
      history.replaceState(null, "", location.pathname + (tmp?.hash || ""));
      if (p.has("error")) throw new Error(p.get("error_description") || "Sign-in was cancelled.");
      if (!tmp || tmp.state !== p.get("state")) throw new Error("Sign-in expired. Please sign in again.");
      await tokenRequest({ grant_type: "authorization_code", code: p.get("code"), redirect_uri: redirectUri, code_verifier: tmp.verifier });
      return true;
    }

    async function getToken() {
      const t = store.get(KEY);
      if (!t) return null;
      if (t.exp > Date.now()) return t.access;
      if (!t.refresh) { store.del(KEY); return null; }
      try { return (await tokenRequest({ grant_type: "refresh_token", refresh_token: t.refresh })).access; }
      catch { store.del(KEY); return null; }
    }

    function signOut() { store.del(KEY); }
    const signedIn = () => !!store.get(KEY);
    return { signIn, handleRedirect, getToken, signOut, signedIn };
  }

  root.MsAuth = { create };
})(typeof window !== "undefined" ? window : globalThis);

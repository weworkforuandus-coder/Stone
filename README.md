# StoneFlow Surfaces: website + OneDrive inventory agent

## Who can do what

| Person | Sign-in | Can |
|---|---|---|
| Anyone (customers, public) | none | See live stock (lot #, color, size, stock, on hold, free) and ask the agent "is IQ466 available?" in English or Spanish. No rack locations, no names of who holds slabs, no sales orders. |
| Signed in, not on the staff list | Microsoft | Same as anyone (read-only). |
| Staff (owner + `ALLOWED_EMAILS`) | Microsoft | Everything: sales, holds, sales orders (PDF), undo, sales log. |

The site's server opens the Excel file with the **owner's** OneDrive connection (connected once at `/api/connect`).
Staff don't need the Excel file shared with them, and nobody's browser ever gets the owner's access.

### Netlify environment variables

| Variable | Value |
|---|---|
| `MS_CLIENT_ID` | Azure Application (client) ID |
| `MS_CLIENT_SECRET` | Azure client secret **Value** (mark as secret) |
| `OWNER_EMAIL` | the Microsoft account that owns the Excel file |
| `EXCEL_PATH` | path inside the owner's OneDrive, e.g. `/Stone/LIVE INVENTORY.xlsx` |
| `ALLOWED_EMAILS` | staff who may change the inventory, comma-separated |
| `ANTHROPIC_API_KEY` | Claude API key (mark as secret) |
| `PUBLIC_AGENT` | optional: `off` hides the agent from visitors |

Azure: keep the Single-page application redirect `https://YOUR-SITE.netlify.app/`, and add a **Web** platform redirect
`https://YOUR-SITE.netlify.app/api/connect` plus a client secret. After deploying, the owner opens
`https://YOUR-SITE.netlify.app/api/connect` once and signs in.

Older setup notes below describe the first version, where each person opened the file with their own OneDrive access.

A company website with a **Staff portal** that works directly on your **LIVE_INVENTORY.xlsx in OneDrive**. Nothing gets copied into a separate database.

- Staff sign in with their Microsoft account.
- They ask the agent: **"Is IQ466 available?"**, **"Sold 2 slabs of IQ466"**, **"Hold 5 IQ1466J for Puru"**.
- The agent writes each sale into the next empty sale column (X, Y, Z… on that lot's row), so STOCK, BALANCE and SALES update with your own formulas. Holds go into the salesperson's column (SUNNY.C … DEV).
- Staff can also record sales, returns and holds by hand from the lot list. Every change has **Undo**.
- A **StoneFlow sales log.csv** is kept next to the Excel file, recording who sold what, with notes.
- **Sales orders:** say *"Create a sales order for ABC Granite: 3 slabs of IQ466 at $450 and 2 of IQ1466J"*. The agent:
  1. checks every line against live stock (if any line is short, nothing is recorded and it tells you what is free),
  2. deducts the slabs in the Excel file,
  3. saves the order in **StoneFlow sales orders.csv** next to the Excel file (order number like SO-261013-01, client, lots, sizes, slabs, sq ft, prices),
  4. shows a printable sales order you can print or download.

  You can cancel an order from the chat, from the order itself, or by asking the agent ("cancel SO-261013-01"). Cancelling puts the slabs back.

Only the changed cells are touched. Formatting, merged cells, colors and formulas stay as they are.

---

## What's in this folder

| Path | What it is |
|---|---|
| `site/index.html` | The website (public pages + Staff portal at `#staff`) |
| `site/js/config.js` | **Your settings**: Microsoft app ID and the Excel file link |
| `site/js/portal.js` | Staff portal and the agent's tools |
| `site/js/stores.js` | Reads and writes the Excel file on OneDrive |
| `site/js/xlsxlite.js` | Edits single cells inside the .xlsx without changing anything else |
| `site/js/stonesheet.js` | Understands your LIVE_INVENTORY layout (lots, holds, sections, outside colors) |
| `site/js/auth.js` | Microsoft sign-in |
| `netlify/functions/agent.mjs` | Server function that holds the Claude API key and checks the staff list |
| `netlify.toml` | Hosting settings and security headers |

## How it works

```
Staff browser ──sign in──▶ Microsoft
     │  reads/writes cells ──▶ OneDrive: LIVE_INVENTORY.xlsx (+ sales log)
     │  asks the agent ──▶ /api/agent (your Netlify function) ──▶ Claude
     ◀── Claude asks for find_lots / record_sale / set_hold; the browser runs them on the Excel file
```

The Claude API key never reaches the browser. The function only answers people whose Microsoft email is on your `ALLOWED_EMAILS` list.

### Personal OneDrive vs OneDrive for Business

| Your OneDrive | How the site writes | Good to know |
|---|---|---|
| **Personal** (outlook.com, gmail sign-in) | Downloads the file, changes the cells, and uploads it with a version check. If someone else saved first, it re-reads and retries. | If the file is open in **desktop Excel without AutoSave**, saves are blocked until it's closed. Excel for the web shows "newer version available"; click refresh. |
| **Business / SharePoint** (Microsoft 365) | Writes cells live through Microsoft's Excel API | Works while others have the file open. |

Microsoft's live Excel API only works on Business/SharePoint, so personal OneDrive uses the file method. The site picks the right one automatically.

---

## Setup (about 45 minutes, on a computer)

You need:
- a **Microsoft account** that owns the Excel file
- a free **Netlify** account (https://netlify.com)
- an **Anthropic Console** account with a little credit (https://console.anthropic.com)
- **Node.js LTS** (https://nodejs.org)

### Step 1: Put the Excel file in OneDrive and share it

1. Upload **LIVE_INVENTORY.xlsx** to your OneDrive.
2. Right-click it → **Share** → choose **Specific people** → add each staff member's Microsoft email → set **Can edit**.
3. Click **Copy link**. Save this link; it's your `shareUrl`.

Each staff member needs a Microsoft account with the same email you shared with. A free one at https://account.microsoft.com is fine.

### Step 2: Register the app with Microsoft (Azure, free)

1. Go to https://portal.azure.com and sign in with the Microsoft account that owns the file. If it asks you to create a free Azure account, do that; there's no charge for app registrations.
2. Search for **App registrations** → **New registration**.
   - Name: `StoneFlow inventory`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: platform **Single-page application (SPA)**. Use `https://YOUR-SITE.netlify.app/` (you get this address in step 4, so you can add it afterwards).
3. Click **Register**. Copy the **Application (client) ID**.
4. Open **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** → tick `User.Read`, `Files.ReadWrite.All`, `offline_access` → **Add permissions**.

### Step 3: Get a Claude API key

1. In the Anthropic Console, under **Billing**, add credit (for example $5). Under **Limits**, set a monthly spend limit (for example $20).
2. **API Keys** → **Create Key**. Copy it (starts with `sk-ant-`).

### Step 4: Put the site online (Netlify)

Open a terminal in this folder:

```
npx netlify-cli login
npx netlify-cli init          # choose "Create & configure a new site", pick a site name
npx netlify-cli env:set ANTHROPIC_API_KEY sk-ant-YOUR-KEY
npx netlify-cli env:set ALLOWED_EMAILS "you@gmail.com, ravi@yourcompany.com"
```

Now your site address is `https://YOUR-SITE.netlify.app`. Add it as the SPA redirect URI in Azure (step 2.2), with the trailing `/`.

Netlify's drag-and-drop upload doesn't include server functions, so use the commands above (or connect a GitHub repo).

### Step 5: Fill in the settings

Edit `site/js/config.js`:

```js
microsoft: { clientId: "PASTE-APPLICATION-CLIENT-ID", tenant: "common" },
excel:     { shareUrl: "PASTE-THE-ONEDRIVE-SHARE-LINK", ... },
```

Also put your company details in `companyDetails` (printed on sales orders), and replace the sample address and phone number in `site/index.html` (the **Visit the yard** section).

Then deploy:

```
npx netlify-cli deploy --prod
```

### Step 6: Test

1. Open your site → **Staff portal** → **Sign in with Microsoft** → accept the permissions.
2. The portal shows your lots and the sheet name (e.g. AUG 2026).
3. In the agent, try:
   - `Is IQ466 available?`
   - `Sold 2 slabs of IQ466`: open the Excel file; the 2 is in the next empty cell of row 219, and STOCK dropped by 2. Then tap **Undo** in the chat.
   - `Hold 3 IQ1466J for Puru`, then `release Puru's hold on IQ1466J`.
   - `Sold 999 of IQ466`: it refuses.
   - `Create a sales order for Test Client: 1 slab of IQ466 at $400`: you get an order number. Check the **Sales orders** panel and **StoneFlow sales orders.csv** in OneDrive, then click **Cancel order**.

### Step 7: Contact form (optional)

The **Visit** form uses Netlify Forms. In Netlify: **Forms** → **Enable form detection** → redeploy. Messages appear under Forms, and you can turn on email notifications there.

---

## Day to day

- **New month:** add the new month's sheet (e.g. "SEP 2026") as usual. The portal picks the newest month sheet automatically. You can also switch sheets in the portal.
- **Adding staff:** share the Excel file with them (step 1), and add their email to `ALLOWED_EMAILS` (`npx netlify-cli env:set …`, then redeploy).
- **Removing staff:** remove them from the file's sharing and from `ALLOWED_EMAILS`.
- **Rows full:** each row holds sales in columns X–BL (41 entries). If a lot's row fills up, the agent tells you; start the new month sheet.
- **Sales-manager holds** ("10 puru" in HOLD BY SALES.MGR) count as held but stay editable only in Excel.
- **Outside colors:** a sale lowers the STK number in that list.

## Security

- Staff sign in with Microsoft. Nobody can open the file without your OneDrive sharing.
- The agent function checks every request against Microsoft and `ALLOWED_EMAILS`, and only allows the three inventory tools.
- The Claude key exists only in Netlify's environment variables.
- Security headers stop other sites from embedding yours and only allow scripts from your own site.
- OneDrive keeps **version history** on the Excel file: right-click → **Version history** to restore any earlier copy.
- Turn on two-step sign-in for your Microsoft, Netlify and Anthropic accounts.

## Costs

| Part | Cost |
|---|---|
| Netlify (free plan) | $0 |
| OneDrive / Azure app registration | $0 (uses your existing OneDrive) |
| Claude API (Haiku model) | about 1 cent per agent question; your spend limit caps it |

## Troubleshooting

| You see | Fix |
|---|---|
| Sign-in error "redirect URI mismatch" | The site address in Azure must match exactly, including `https://` and the trailing `/`, under **Single-page application**. |
| "Couldn't open the Excel file" | `shareUrl` must be a link to the file, shared with the account you signed in with, with **Can edit**. |
| "the order file couldn't be saved" | The account must be able to create files in the Excel file's folder: share the **folder** (not only the file) with Can edit. Stock was still updated. |
| "isn't on the staff list" | Add that email to `ALLOWED_EMAILS` and redeploy. |
| "The Excel file is locked" | Someone has it open in desktop Excel without AutoSave. Close it or turn AutoSave on. |
| "The Excel sheet changed and lot … moved" | Someone inserted or deleted rows in Excel. Tap **Refresh**. |
| Agent says it's busy / Claude error | Check your Anthropic credit and spend limit. |
| Excel shows old totals | Open the file in Excel; it recalculates on open. The portal always computes stock from OP + PURCHASE − SALES. |

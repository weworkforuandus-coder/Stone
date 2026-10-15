// Site settings. Fill these in after the Azure and Netlify steps in README.md.
window.STONEFLOW_CONFIG = {
  company: "StoneFlow Surfaces",
  // Printed on sales orders.
  companyDetails: { address: "", phone: "(999) 999-9999", email: "weworkforuandus@gmail.com" },
  currency: "USD",

  // Azure app registration (README step 2). Leave clientId empty to run the portal in demo mode.
  microsoft: {
    clientId: "68948113-4da7-4aac-86f1-eb2ae369e9aa",                 // "Application (client) ID" from Azure
    tenant: "60025a4f-7cc3-4607-aa4b-a762755d2f11",             // "common" = work and personal Microsoft accounts
  },

  // Where LIVE_INVENTORY.xlsx lives (README step 3). Use ONE of these:
  excel: {
    shareUrl: "https://onedrive.live.com/personal/41561d8edc467c53/_layouts/15/Doc.aspx?sourcedoc=%7B8361661A-B706-4A61-9830-9E816D505CA0%7D&file=LIVE%20INVENTORY.xlsx&action=default&mobileredirect=true",                 // a OneDrive "Can edit" sharing link to the file (works for everyone you share with)
    filePath: "https://onedrive.live.com/personal/41561d8edc467c53/Documents/Stone/LIVE%20INVENTORY.xlsx",                 // or a path in YOUR OneDrive, e.g. "/Documents/LIVE_INVENTORY.xlsx" (only you)
    sheet: "AUG 2026",                    // optional; empty = newest month sheet (e.g. "AUG 2026")
    mode: "auto",                 // auto | file | workbook (auto: Excel API on Business/SharePoint, file mode on personal OneDrive)
    salesLogName: "StoneFlow sales log.csv",       // written next to the Excel file; who sold what, with notes
    ordersFileName: "StoneFlow sales orders.csv",  // sales orders created by the agent (one line per lot)
  },

  // The agent's server function (netlify/functions/agent.mjs). Keeps the Claude API key off the browser.
  agentUrl: "/api/agent",
};

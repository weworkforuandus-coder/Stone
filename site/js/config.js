// Site settings. Fill these in after the Azure and Netlify steps in README.md.
window.STONEFLOW_CONFIG = {
  company: "StoneFlow Surfaces",
  // Printed on sales orders.
  companyDetails: { address: "Sample address, City, State ZIP", phone: "(000) 000-0000", email: "sales@example.com" },
  currency: "USD",

  // Azure app registration (README step 2). Leave clientId empty to run the portal in demo mode.
  microsoft: {
    clientId: "",                 // "Application (client) ID" from Azure
    tenant: "common",             // "common" = work and personal Microsoft accounts
  },

  // Where LIVE_INVENTORY.xlsx lives (README step 3). Use ONE of these:
  excel: {
    shareUrl: "",                 // a OneDrive "Can edit" sharing link to the file (works for everyone you share with)
    filePath: "",                 // or a path in YOUR OneDrive, e.g. "/Documents/LIVE_INVENTORY.xlsx" (only you)
    sheet: "",                    // optional; empty = newest month sheet (e.g. "AUG 2026")
    mode: "auto",                 // auto | file | workbook (auto: Excel API on Business/SharePoint, file mode on personal OneDrive)
    salesLogName: "StoneFlow sales log.csv",       // written next to the Excel file; who sold what, with notes
    ordersFileName: "StoneFlow sales orders.csv",  // sales orders created by the agent (one line per lot)
  },

  // The agent's server function (netlify/functions/agent.mjs). Keeps the Claude API key off the browser.
  agentUrl: "/api/agent",
};

// Site settings.
window.STONEFLOW_CONFIG = {
  company: "StoneFlow Surfaces",
  // Printed on sales orders.
  companyDetails: { address: "", phone: "(999) 999-9999", email: "weworkforuandus@gmail.com" },
  currency: "USD",
  // Only these sign-ins see the "things to check in the sheet" notes.
  adminEmails: ["weworkforuandus@gmail.com"],

  // Microsoft sign-in for staff (Azure app registration).
  microsoft: {
    clientId: "68948113-4da7-4aac-86f1-eb2ae369e9aa",
    tenant: "consumers",
  },

  // The Excel file is opened by the site's server with the owner's OneDrive connection
  // (Netlify variable EXCEL_PATH). Anyone can view stock; staff in ALLOWED_EMAILS can make changes.
  excel: {
    sheet: "",                    // empty = newest month sheet (e.g. "AUG 2026")
    salesLogName: "StoneFlow sales log.csv",
    ordersFileName: "StoneFlow sales orders.csv",
  },

  agentUrl: "/api/agent",
};

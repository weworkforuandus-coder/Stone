// Site settings.
window.STONEFLOW_CONFIG = {
  company: "StoneFlow Surfaces",
  // Printed on sales orders.
  companyDetails: { address: "YOUR ADDRESS, City, State ZIP", phone: "(999) 999-9999", email: "weworkforuandus@gmail.com" },
  currency: "USD",
  // Only these sign-ins see the "things to check in the sheet" notes.
  adminEmails: ["weworkforuandus@gmail.com"],

  microsoft: {
    clientId: "68948113-4da7-4aac-86f1-eb2ae369e9aa",
    tenant: "consumers",
  },

  excel: {
    shareUrl: "",
    filePath: "/Stone/LIVE INVENTORY.xlsx",
    sheet: "AUG 2026",
    mode: "auto",
    salesLogName: "StoneFlow sales log.csv",
    ordersFileName: "StoneFlow sales orders.csv",
  },

  agentUrl: "/api/agent",
};
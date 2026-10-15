// Loads the same sheet reader the website uses (site/js), so the server and browser read the Excel file identically.
if (typeof globalThis.DOMParser === "undefined") globalThis.DOMParser = require("./minidom.cjs").DOMParser;
globalThis.StoneSheet = require("../../site/js/stonesheet.js");
globalThis.XlsxLite = require("../../site/js/xlsxlite.js");
require("../../site/js/stores.js");
module.exports = globalThis.InventoryStores;

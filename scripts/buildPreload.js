"use strict";
const fs   = require("fs");
const path = require("path");

const webappDir = path.join(__dirname, "..", "app", "atc-ui", "webapp");
const files = {};

function walk(dir, base) {
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    const rel  = (base ? base + "/" + f : f);
    // Skip the preload file itself to avoid recursive bundling
    if (rel === "Component-preload.js") continue;
    if (fs.statSync(full).isDirectory()) {
      walk(full, rel);
    } else if (f.endsWith(".js") || f.endsWith(".xml") || f.endsWith(".json") || f.endsWith(".properties")) {
      const moduleName = "com/sap/atcanalysis/" + rel;
      files[moduleName] = fs.readFileSync(full, "utf8");
    }
  }
}

walk(webappDir, "");

let out = "sap.ui.require.preload({\n";
const entries = Object.entries(files);
entries.forEach(([name, content], i) => {
  out += "  " + JSON.stringify(name) + ": " + JSON.stringify(content);
  if (i < entries.length - 1) out += ",";
  out += "\n";
});
out += "}, \"com.sap.atcanalysis.Component-preload\");\n";

const dest = path.join(webappDir, "Component-preload.js");
fs.writeFileSync(dest, out);
console.log("Written Component-preload.js with", entries.length, "modules:");
entries.forEach(([n]) => console.log(" ", n));

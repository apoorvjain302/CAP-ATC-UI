"use strict";
const fs   = require("fs");
const path = require("path");

const webapp = path.join(__dirname, "app", "atc-ui", "webapp");
const map = {};

function walk(dir, base) {
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    const rel  = (base + "/" + f);
    if (fs.statSync(full).isDirectory()) {
      walk(full, rel);
    } else if (/\.(js|xml|json|properties)$/.test(f)) {
      if (f === "Component-preload.js") continue;
      map[rel] = fs.readFileSync(full, "utf8");
    }
  }
}
walk(webapp, "com/sap/atcanalysis");

const entries = Object.entries(map)
  .map(([k, v]) => JSON.stringify(k) + ": " + JSON.stringify(v))
  .join(",\n  ");
const out = "sap.ui.require.preload({\n  " + entries + "\n}, \"com.sap.atcanalysis.Component-preload\");\n";
fs.writeFileSync(path.join(webapp, "Component-preload.js"), out);
console.log("Written", Object.keys(map).length, "entries to Component-preload.js");

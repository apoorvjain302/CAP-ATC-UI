"use strict";
/**
 * bufferUtil.js — shared binary buffer helper for CAP LargeBinary round-trips.
 *
 * All file content is stored in SQLite as a hex string (via .toString("hex")).
 * CAP / better-sqlite3 may return it as:
 *   - string  → hex string  (normal path)
 *   - Buffer  → raw bytes already decoded by newer driver versions
 *   - Uint8Array / stream → other CAP v8 driver variants
 *
 * Use toBuffer() when reading content back from DB.
 * Use toHex()   when writing content to DB.
 */

async function _streamToBuffer(stream) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    stream.on("data", c => chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c)));
    stream.on("end",  () => resolve(Buffer.concat(chunks)));
    stream.on("error", reject);
  });
}

/**
 * Convert a DB-returned LargeBinary value to a Node.js Buffer.
 * Safe to call on any value shape CAP / better-sqlite3 may return.
 */
async function toBuffer(val) {
  if (!val) return null;
  if (typeof val === "string") return Buffer.from(val, "hex");
  if (Buffer.isBuffer(val))   return val;
  if (val instanceof Uint8Array) return Buffer.from(val);
  if (typeof val.pipe === "function" || typeof val.on === "function") return _streamToBuffer(val);
  if (val.buffer) return Buffer.from(val.buffer);
  return Buffer.from(String(val), "hex");
}

/**
 * Convert a Buffer to a hex string for safe DB storage.
 * Use this on every INSERT/UPDATE of a LargeBinary column.
 */
function toHex(buf) {
  if (!buf) return null;
  if (Buffer.isBuffer(buf)) return buf.toString("hex");
  if (buf instanceof Uint8Array) return Buffer.from(buf).toString("hex");
  return String(buf);
}

module.exports = { toBuffer, toHex };

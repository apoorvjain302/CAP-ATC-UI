"use strict";
/**
 * pptxGen.js — Node.js port of tool_pptx.py::write_pptx
 *
 * Uses a lossless ZIP patching strategy: reads the template ZIP, decompresses
 * only the entries that need modification, patches the XML, recompresses with
 * the same method, and reconstructs the ZIP.  All other entries are copied
 * byte-for-byte (raw compressed data preserved), matching Python's zipfile
 * behaviour and avoiding JSZip's full-recompression artefacts.
 */

const zlib = require("zlib");
const JSZip = require("jszip");   // still used to READ the template easily

const _GREEN  = "92D050";
const _YELLOW = "FFFF00";
const _RED    = "FF0000";

function _escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function _replaceChartCategories(xml, labels) {
  const catM = xml.match(/<c:cat>([\s\S]*?)<\/c:cat>/);
  if (!catM) return xml;
  let cache = `<c:ptCount val="${labels.length}"/>`;
  labels.forEach((lbl, i) => {
    cache += `<c:pt idx="${i}"><c:v>${_escapeXml(lbl)}</c:v></c:pt>`;
  });
  const newCat = catM[0].replace(
    /<c:strCache>[\s\S]*?<\/c:strCache>/,
    `<c:strCache>${cache}</c:strCache>`,
  );
  return xml.slice(0, catM.index) + newCat + xml.slice(catM.index + catM[0].length);
}

function _replaceSeriesValues(xml, seriesIndex, values) {
  const serPattern = /<c:ser[\s>]/g;
  const serEnd     = "</c:ser>";
  const positions  = [];
  let m;
  while ((m = serPattern.exec(xml)) !== null) {
    const endPos = xml.indexOf(serEnd, m.index);
    if (endPos === -1) break;
    positions.push([m.index, endPos + serEnd.length]);
  }
  if (seriesIndex >= positions.length) return xml;

  const [s, e] = positions[seriesIndex];
  let serXml = xml.slice(s, e);
  let cache = `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>`;
  values.forEach((v, i) => { cache += `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`; });

  serXml = serXml.replace(
    /(<c:val>[\s\S]*?<c:numRef>[\s\S]*?)<c:numCache>[\s\S]*?<\/c:numCache>/,
    `$1<c:numCache>${cache}</c:numCache>`,
  );
  return xml.slice(0, s) + serXml + xml.slice(e);
}

function _replaceSeriesColors(xml, seriesIndex, colors) {
  const serPattern = /<c:ser[\s>]/g;
  const serEnd     = "</c:ser>";
  const positions  = [];
  let m;
  while ((m = serPattern.exec(xml)) !== null) {
    const endPos = xml.indexOf(serEnd, m.index);
    if (endPos === -1) break;
    positions.push([m.index, endPos + serEnd.length]);
  }
  if (seriesIndex >= positions.length) return xml;

  const [s, e] = positions[seriesIndex];
  let serXml = xml.slice(s, e);

  let dptBlock = "";
  colors.forEach((color, i) => {
    dptBlock +=
      `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/>` +
      `<c:spPr><a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:dPt>`;
  });

  serXml = serXml.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/g, "");
  const insertM = serXml.match(/<c:cat>|<c:val>/);
  if (insertM) {
    serXml = serXml.slice(0, insertM.index) + dptBlock + serXml.slice(insertM.index);
  }
  return xml.slice(0, s) + serXml + xml.slice(e);
}

function _removePivotSource(xml) {
  xml = xml.replace(/<c:pivotSource>[\s\S]*?<\/c:pivotSource>/g, "");
  // Remove full element: <c:externalData ...>...</c:externalData> or self-closing <c:externalData .../>
  xml = xml.replace(/<c:externalData\b[\s\S]*?<\/c:externalData>/g, "");
  xml = xml.replace(/<c:externalData\b[^>]*\/>/g, "");
  return xml;
}

function _updateChartXml(xml, categories, values, colors) {
  xml = _removePivotSource(xml);
  xml = _replaceChartCategories(xml, categories);
  xml = _replaceSeriesValues(xml, 0, values);
  if (colors) xml = _replaceSeriesColors(xml, 0, colors);
  return xml;
}

async function _getChartPaths(zip, slideIndex) {
  const relsName = `ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`;
  const relsFile = zip.file(relsName);
  if (!relsFile) return [];
  const relsXml = await relsFile.async("string");

  const paths = [];
  const re    = /Target="([^"]*chart[^"]*)"/gi;
  let mm;
  while ((mm = re.exec(relsXml)) !== null) {
    let target = mm[1];
    if (!target.startsWith("ppt/")) {
      target = _resolvePath("ppt/slides/", target);
    }
    paths.push(target);
  }
  return paths;
}

function _resolvePath(base, relative) {
  const parts = (base + relative).split("/");
  const out   = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p !== ".") out.push(p);
  }
  return out.join("/");
}

function _replaceImpactCount(xml, label, value) {
  const idx = xml.indexOf(label);
  if (idx === -1) return xml;
  const afterLabel = xml.indexOf(": ", idx + label.length);
  if (afterLabel === -1) return xml;
  const runEnd = xml.indexOf("</a:t>", afterLabel);
  if (runEnd === -1) return xml;
  const runText = xml.slice(afterLabel, runEnd);
  const newRun  = runText.replace(/\d+(?= impacts)/, value);
  if (newRun === runText) return xml;
  return xml.slice(0, afterLabel) + newRun + xml.slice(runEnd);
}

// ── Low-level ZIP builder ─────────────────────────────────────────────────────
// Builds a PPTX ZIP where we patch only the entries listed in `patches`
// (Map<name, Buffer>) and copy everything else raw from the template buffer.
// This avoids JSZip's full-recompression which changes the ZIP structure.

function _uint32LE(buf, offset) {
  return buf[offset] | (buf[offset+1] << 8) | (buf[offset+2] << 16) | (buf[offset+3] * 0x1000000);
}
function _uint16LE(buf, offset) {
  return buf[offset] | (buf[offset+1] << 8);
}
function _writeUint32LE(buf, offset, val) {
  buf[offset]   = val & 0xff;
  buf[offset+1] = (val >>> 8)  & 0xff;
  buf[offset+2] = (val >>> 16) & 0xff;
  buf[offset+3] = (val >>> 24) & 0xff;
}
function _writeUint16LE(buf, offset, val) {
  buf[offset]   = val & 0xff;
  buf[offset+1] = (val >>> 8) & 0xff;
}

// Parse all local file headers from a ZIP buffer.
// Returns array of { name, method, crc32, compressedSize, uncompressedSize,
//                    headerStart, dataStart, dataEnd }
function _parseZipEntries(buf) {
  const entries = [];
  let pos = 0;
  while (pos + 30 <= buf.length) {
    const sig = _uint32LE(buf, pos);
    if (sig !== 0x04034b50) break;  // local file header signature
    const method     = _uint16LE(buf, pos + 8);
    const crc32      = _uint32LE(buf, pos + 14);
    let   compSize   = _uint32LE(buf, pos + 18);
    let   uncompSize = _uint32LE(buf, pos + 22);
    const nameLen    = _uint16LE(buf, pos + 26);
    const extraLen   = _uint16LE(buf, pos + 28);
    const name       = buf.slice(pos + 30, pos + 30 + nameLen).toString("utf8");
    const headerStart = pos;
    const dataStart   = pos + 30 + nameLen + extraLen;

    // Handle data descriptor (bit 3 of general purpose flag)
    const flags = _uint16LE(buf, pos + 6);
    let dataEnd;
    if (flags & 0x0008) {
      // sizes in local header are 0; scan for data descriptor
      // data descriptor: optional sig 0x08074b50 then crc32, compSize, uncompSize
      let ddPos = dataStart;
      // We'll compute the actual compressed size after we parse the central dir
      // For now, store what's in the header (may be 0) and fix up later
      dataEnd = dataStart + compSize; // placeholder
    } else {
      dataEnd = dataStart + compSize;
    }

    entries.push({ name, method, crc32, compressedSize: compSize, uncompressedSize: uncompSize,
                   headerStart, dataStart, dataEnd, flags });
    pos = dataEnd;
  }
  return entries;
}

// Finds the End-of-Central-Directory record in a ZIP buffer.
function _findEOCD(buf) {
  // Search backwards from end
  for (let i = buf.length - 22; i >= 0; i--) {
    if (_uint32LE(buf, i) === 0x06054b50) return i;
  }
  return -1;
}

// Parse central directory entries.
function _parseCentralDir(buf) {
  const eocdPos = _findEOCD(buf);
  if (eocdPos === -1) return null;
  const cdOffset = _uint32LE(buf, eocdPos + 16);
  const cdSize   = _uint32LE(buf, eocdPos + 12);
  const cdCount  = _uint16LE(buf, eocdPos + 8);

  const entries = [];
  let pos = cdOffset;
  for (let i = 0; i < cdCount; i++) {
    if (_uint32LE(buf, pos) !== 0x02014b50) break;
    const method       = _uint16LE(buf, pos + 10);
    const crc32        = _uint32LE(buf, pos + 16);
    const compSize     = _uint32LE(buf, pos + 20);
    const uncompSize   = _uint32LE(buf, pos + 24);
    const nameLen      = _uint16LE(buf, pos + 28);
    const extraLen     = _uint16LE(buf, pos + 30);
    const commentLen   = _uint16LE(buf, pos + 32);
    const localOffset  = _uint32LE(buf, pos + 42);
    const name         = buf.slice(pos + 46, pos + 46 + nameLen).toString("utf8");
    const extraField   = buf.slice(pos + 46 + nameLen, pos + 46 + nameLen + extraLen);
    const fileAttr     = buf.slice(pos + 38, pos + 42);  // external file attributes
    const internalAttr = _uint16LE(buf, pos + 36);
    const versionMade  = _uint16LE(buf, pos + 4);
    const versionNeeded= _uint16LE(buf, pos + 6);
    const flags        = _uint16LE(buf, pos + 8);
    const dosTime      = _uint16LE(buf, pos + 12);
    const dosDate      = _uint16LE(buf, pos + 14);
    const cdEntry      = buf.slice(pos, pos + 46 + nameLen + extraLen + commentLen);
    entries.push({ name, method, crc32, compSize, uncompSize, localOffset,
                   nameLen, extraLen, commentLen, extraField, versionMade,
                   versionNeeded, flags, dosTime, dosDate, internalAttr,
                   fileAttr, cdEntry, cdEntryStart: pos });
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return { entries, eocdPos, cdOffset, cdSize };
}

// Build CRC-32 table
const _CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function _crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) crc = _CRC32_TABLE[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

/**
 * Rebuild a ZIP buffer, replacing entries listed in `patches` (Map<name, Buffer>)
 * with new deflated content.  All other entries are copied verbatim (raw bytes).
 * Returns a Buffer.
 */
async function _rebuildZip(templateBuf, patches) {
  const cd = _parseCentralDir(templateBuf);
  if (!cd) throw new Error("Could not parse ZIP central directory");

  const outParts = [];
  const newCdEntries = [];
  let outOffset = 0;

  for (const cdEnt of cd.entries) {
    const localHeader = templateBuf.slice(cdEnt.localOffset, cdEnt.localOffset + 30);
    const nameLen  = _uint16LE(localHeader, 26);
    const extraLen = _uint16LE(localHeader, 28);
    const dataStart = cdEnt.localOffset + 30 + nameLen + extraLen;
    const dataEnd   = dataStart + cdEnt.compSize;

    let newLocalHeader, newData;
    const patchBuf = patches.get(cdEnt.name);

    if (patchBuf !== undefined) {
      // Re-compress the patched content
      const deflated  = zlib.deflateRawSync(patchBuf, { level: 6 });
      const crc       = _crc32(patchBuf);
      const compSize  = deflated.length;
      const uncompSize = patchBuf.length;

      // Clone local header and update sizes/crc
      newLocalHeader = Buffer.from(localHeader);
      _writeUint16LE(newLocalHeader, 6,  cdEnt.flags & ~0x0008); // clear data-descriptor flag
      _writeUint16LE(newLocalHeader, 8,  8);   // method = deflate
      _writeUint32LE(newLocalHeader, 14, crc);
      _writeUint32LE(newLocalHeader, 18, compSize);
      _writeUint32LE(newLocalHeader, 22, uncompSize);

      const nameAndExtra = templateBuf.slice(cdEnt.localOffset + 30, cdEnt.localOffset + 30 + nameLen + extraLen);
      newData = deflated;

      // Build new CD entry (copy template CD entry, update sizes/crc/offset)
      const newCd = Buffer.from(cdEnt.cdEntry);
      _writeUint32LE(newCd, 16, crc);
      _writeUint32LE(newCd, 20, compSize);
      _writeUint32LE(newCd, 24, uncompSize);
      _writeUint32LE(newCd, 42, outOffset);
      _writeUint16LE(newCd, 10, 8); // method = deflate

      outParts.push(newLocalHeader, nameAndExtra, newData);
      newCdEntries.push(newCd);
      outOffset += newLocalHeader.length + nameAndExtra.length + newData.length;
    } else {
      // Copy raw — entire local entry verbatim
      const rawEntry = templateBuf.slice(cdEnt.localOffset, dataEnd);

      // Clone CD entry with updated offset only
      const newCd = Buffer.from(cdEnt.cdEntry);
      _writeUint32LE(newCd, 42, outOffset);

      outParts.push(rawEntry);
      newCdEntries.push(newCd);
      outOffset += rawEntry.length;
    }
  }

  // Central directory
  const cdStart = outOffset;
  const cdBufs  = newCdEntries;
  const cdTotal  = cdBufs.reduce((s, b) => s + b.length, 0);
  outParts.push(...cdBufs);

  // EOCD
  const eocd = Buffer.alloc(22);
  _writeUint32LE(eocd, 0,  0x06054b50);
  _writeUint16LE(eocd, 4,  0);
  _writeUint16LE(eocd, 6,  0);
  _writeUint16LE(eocd, 8,  cd.entries.length);
  _writeUint16LE(eocd, 10, cd.entries.length);
  _writeUint32LE(eocd, 12, cdTotal);
  _writeUint32LE(eocd, 16, cdStart);
  _writeUint16LE(eocd, 20, 0);
  outParts.push(eocd);

  return Buffer.concat(outParts);
}

/**
 * Generate PPTX from template buffer.
 * Returns a Buffer containing the new PPTX file.
 */
async function writePptx(templateBuf, atcData, spddCount, spauCount, customerName) {
  // Use JSZip only for READ operations (convenient API)
  const zip  = await JSZip.loadAsync(templateBuf);
  const data = atcData;

  const _readXml = (name) => {
    const f = zip.file(name);
    return f ? f.async("string") : Promise.resolve("");
  };

  // Collect all patches: Map<zipEntryName, Buffer>
  const patches = new Map();
  const _patch = (name, xml) => patches.set(name, Buffer.from(xml, "utf8"));

  // ── Slide 1: customer name + date placeholder ────────────────────────────
  {
    let s1 = await _readXml("ppt/slides/slide1.xml");
    if (customerName) {
      // Template contains XML-encoded placeholder: &lt;Customer Name&gt; S/4HANA Conversion
      // Replace it with the actual customer name
      s1 = s1.replace(
        /&lt;Customer Name&gt;(?=\s+S\/4HANA Conversion)/,
        _escapeXml(customerName),
      );
      // Fallback: any text ending with "- S/4HANA Conversion " (old template style)
      if (!s1.includes(customerName)) {
        s1 = s1.replace(
          /[^<>"']+(?=- S\/4HANA Conversion )/,
          _escapeXml(customerName),
        );
      }
    }
    // Replace any date text (e.g. "June 04, 2026") with <Date> placeholder
    s1 = s1.replace(
      /(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s*\d{4}/g,
      "&lt;Date&gt;",
    );
    _patch("ppt/slides/slide1.xml", s1);
  }

  // ── Slide 4: object type bar chart ───────────────────────────────────────
  if (data) {
  const chart4Paths = await _getChartPaths(zip, 3);
  if (chart4Paths.length) {
    let xml = await _readXml(chart4Paths[0]);
    xml = _updateChartXml(xml, data.uniqueObjTypes, data.uniqueObjCounts);
    _patch(chart4Paths[0], xml);
  }

  // ── Slide 5: remediation category bar chart ───────────────────────────────
  const slide5Colors = [
    _RED, _RED, _RED, _RED, _GREEN, _YELLOW, _RED, _RED, _RED, _RED,
  ];
  const chart5Paths = await _getChartPaths(zip, 4);
  if (chart5Paths.length && data.slide5Values.some(v => v > 0)) {
    let xml = await _readXml(chart5Paths[0]);
    xml = _updateChartXml(xml, data.slide5Categories, data.slide5Values, slide5Colors);
    _patch(chart5Paths[0], xml);
  }

  // ── Slide 6: impact breakdown bar chart ───────────────────────────────────
  const chart6Paths = await _getChartPaths(zip, 5);
  if (chart6Paths.length && data.slide6Values.some(v => v > 0)) {
    let xml = await _readXml(chart6Paths[0]);
    xml = _updateChartXml(xml, data.slide6Categories, data.slide6Values);
    _patch(chart6Paths[0], xml);
  }

  // ── Slide 7: HCA + S/4H waterfall ────────────────────────────────────────
  const hcaColors = [_RED,_RED,_RED,_GREEN,_RED,_RED,_RED,_RED,_YELLOW];
  const s4hColors = [_RED,_RED,_RED,_RED,_GREEN,_RED,_RED,_GREEN,_YELLOW];
  const chart7Paths = await _getChartPaths(zip, 6);
  if (chart7Paths.length > 0) {
    let xml = await _readXml(chart7Paths[0]);
    xml = _updateChartXml(xml, data.chart8Categories, data.chart8Values, s4hColors);
    _patch(chart7Paths[0], xml);
  }
  if (chart7Paths.length > 1) {
    let xml = await _readXml(chart7Paths[1]);
    xml = _updateChartXml(xml, data.chart7Categories, data.chart7Values, hcaColors);
    _patch(chart7Paths[1], xml);
  }

  // ── Slide 4 XML: unique object count text ─────────────────────────────────
  let s4xml = await _readXml("ppt/slides/slide4.xml");
  const totalObj = data.uniqueObjCounts.reduce((a, b) => a + b, 0);
  s4xml = s4xml.replace(/Unique Object Count \(\d+\)/, `Unique Object Count (${totalObj})`);
  _patch("ppt/slides/slide4.xml", s4xml);

  // ── Slide 6 XML: impact count text boxes ──────────────────────────────────
  let s6xml = await _readXml("ppt/slides/slide6.xml");
  s6xml = _replaceImpactCount(s6xml, "Total Impacts reported by ATC Report",  String(data.totalCount));
  s6xml = _replaceImpactCount(s6xml, "Pre-existing Error Impact",              String(data.preExistingErrorCount));
  s6xml = _replaceImpactCount(s6xml, "HCA Impact (HANA Code Adaptation)",      String(data.hcaCount));
  s6xml = _replaceImpactCount(s6xml, "S/4HANA Impact",                         String(data.s4Count));
  s6xml = _replaceImpactCount(s6xml, "Third party objects",                    String(data.thirdPartyCount));
  _patch("ppt/slides/slide6.xml", s6xml);

  // ── Slide 7 XML: HCA + S/4H title count labels ────────────────────────────
  let s7xml = await _readXml("ppt/slides/slide7.xml");
  s7xml = s7xml.replace(/Break-up for HCA Impacts \(\d+\)/,    `Break-up for HCA Impacts (${data.hcaWaterfallCount || data.hcaCount})`);
  s7xml = s7xml.replace(/Break-up for S\/4HANA Impacts \(\d+\)/, `Break-up for S/4HANA Impacts (${data.s4Count})`);
  _patch("ppt/slides/slide7.xml", s7xml);
  } // end if (data) — slides 4-7 only populated when ATC data is available

  // ── Slide 8 XML: SPDD/SPAU + HCA + S4H summary table ─────────────────────
  // The template table has value cells in their own <a:t> run (separate from label rows).
  // Template values (fixed): "34 objects", "58 objects", "61 impacts", "16* impacts",
  //                           "128 impacts", "194 impacts", "414* impacts"
  // CRITICAL: do NOT use sequential .replace() for these — each call with no /g flag
  // replaces only the FIRST match, so the 2nd call re-replaces the already-substituted
  // first cell instead of the second cell. Replace ALL 7 in a single pass using their
  // exact template literal values so each matches exactly one cell.
  const hcaMandatory  = data ? data.chart7Values[3] : 0;
  const hcaOptional   = data ? data.chart7Values[8] : 0;
  const s4Remediation = data ? data.chart8Values[4] : 0;
  const s4FalsePos    = data ? data.chart8Values[7] : 0;
  const s4Optional    = data ? data.chart8Values[8] : 0;

  let s8xml = await _readXml("ppt/slides/slide8.xml");
  if (s8xml) {
    // Replace each cell by its exact template literal value — guaranteed unique in the table.
    s8xml = s8xml.replace(/<a:t>34 objects<\/a:t>/, `<a:t>${spddCount} objects</a:t>`);
    s8xml = s8xml.replace(/<a:t>58 objects<\/a:t>/, `<a:t>${spauCount} objects</a:t>`);
    s8xml = s8xml.replace(/<a:t>61 impacts<\/a:t>/, `<a:t>${hcaMandatory} impacts</a:t>`);
    s8xml = s8xml.replace(/<a:t>16\* impacts<\/a:t>/, `<a:t>${hcaOptional}* impacts</a:t>`);
    s8xml = s8xml.replace(/<a:t>128 impacts<\/a:t>/, `<a:t>${s4Remediation} impacts</a:t>`);
    s8xml = s8xml.replace(/<a:t>194 impacts<\/a:t>/, `<a:t>${s4FalsePos} impacts</a:t>`);
    s8xml = s8xml.replace(/<a:t>414\* impacts<\/a:t>/, `<a:t>${s4Optional}* impacts</a:t>`);
    _patch("ppt/slides/slide8.xml", s8xml);
  }

  // ── Slide 10 XML: Key Takeaway — replace hardcoded counts + yellow highlight table texts ────────────
  let s10xml = await _readXml("ppt/slides/slide10.xml");
  if (s10xml && data) {
    const preExistingObjCount = (data.syntaxErrorObjCount || 0) + (data.inconsistentObjCount || 0);
    const fitGapObjCount      = data.fitGapObjCount || 0;
    const cloneObjCount       = data.cloneObjCount  || 0;

    // Replace "Review 98 custom code objects" with actual pre-existing object count
    s10xml = s10xml.replace(
      /Review \d+ custom code objects which are reported to have pre-existing issues/,
      `Review ${preExistingObjCount} custom code objects which are reported to have pre-existing issues`
    );

    // Replace "Review 563 custom code objects which involves deprecated functionalities"
    s10xml = s10xml.replace(
      /Review \d+ custom code objects which involves deprecated functionalities/,
      `Review ${fitGapObjCount} custom code objects which involves deprecated functionalities`
    );

    // Replace the two hardcoded FG example bullets with actual top FG categories
    const cats = (data.fitGapCategories && data.fitGapCategories.length > 0)
      ? data.fitGapCategories
      : [];
    const bullet1 = cats[0] ? _escapeXml(cats[0]) : "Fit Gap objects identified in the ATC report";
    const bullet2 = cats[1] ? _escapeXml(cats[1]) : "";

    s10xml = s10xml.replace(
      /<a:t>Credit management changes in FI\/SD<\/a:t>/,
      `<a:t>${bullet1}</a:t>`
    );
    s10xml = s10xml.replace(
      /<a:t>GENERAL LEDGER: INCOMPATIBLE CHANGES IN S\/4HANA COMPARED TO CLASSIC ERP RELEASES<\/a:t>/,
      bullet2 ? `<a:t>${bullet2}</a:t>` : `<a:t><\/a:t>`
    );

    // Replace "Review 9246 clone programs"
    s10xml = s10xml.replace(
      /Review \d+ clone programs/,
      `Review ${cloneObjCount} clone programs`
    );

    // Replace standalone "3454" (clone objects impacted in ATC) — it sits in its own <a:t>
    // The number of clone rows in ATC report = rows where Clone?=Yes
    const cloneImpactRows = (data.cloneImpactCount !== undefined) ? data.cloneImpactCount : cloneObjCount;
    s10xml = s10xml.replace(
      /(<a:t>)\s*3454\s*(<\/a:t>)/,
      `$1${cloneImpactRows}$2`
    );
  }

  // ── Slide 10: yellow highlight all text runs inside the table (<a:tbl>) ────
  if (s10xml) {
    const tblStart = s10xml.indexOf("<a:tbl>");
    const tblEnd   = s10xml.indexOf("</a:tbl>") + "</a:tbl>".length;
    if (tblStart !== -1 && tblEnd > tblStart) {
      let tblXml = s10xml.slice(tblStart, tblEnd);
      // Add yellow highlight to every <a:rPr> inside the table.
      // If <a:highlight> already present, skip; otherwise insert after the opening <a:rPr ...> tag.
      tblXml = tblXml.replace(/(<a:rPr\b[^>]*>)(?![\s\S]*?<a:highlight)/g, (match) => {
        return match + `<a:highlight><a:srgbClr val="${_YELLOW}"/></a:highlight>`;
      });
      s10xml = s10xml.slice(0, tblStart) + tblXml + s10xml.slice(tblEnd);
    }
  }
  _patch("ppt/slides/slide10.xml", s10xml);

  // ── Fix Content_Types.xml: remove stale oleObject + customXml references ─────
  // The template references customXml/itemN.xml files that do not exist in the ZIP.
  // PowerPoint detects these missing parts and shows the repair dialog on open.
  // Strip both oleObject and customXml <Override> entries to prevent this.
  let ctXml = await _readXml("[Content_Types].xml");
  if (ctXml) {
    ctXml = ctXml.replace(/<Override[^>]*oleObject[^>]*\/>/g, "");
    ctXml = ctXml.replace(/<Override[^>]*customXml[^>]*\/>/g, "");
    _patch("[Content_Types].xml", ctXml);
  }

  // ── Strip customXml relationships from ALL slide .rels files ──────────────
  // Every slide's _rels file references customXml/itemN.xml files not in the ZIP.
  // Remove these dangling relationships so PowerPoint does not try to load them.
  const allSlideRels = Object.keys(zip.files).filter(
    k => k.startsWith("ppt/slides/_rels/") && k.endsWith(".xml.rels")
  );
  for (const relsPath of allSlideRels) {
    const relsEntry = zip.file(relsPath);
    if (!relsEntry) continue;
    let relsXml = patches.has(relsPath)
      ? patches.get(relsPath).toString("utf8")
      : await relsEntry.async("string");
    const cleaned = relsXml.replace(/<Relationship[^>]*customXml[^>]*\/>/g, "");
    if (cleaned !== relsXml) _patch(relsPath, cleaned);
  }

  // ── Fix slide11: remove broken OLE graphicFrame objects ───────────────────
  // The template has 3 embedded Excel OLE objects on slide 11 (rId5/rId7/rId9)
  // whose backing .bin files don't exist in the ZIP — PowerPoint repairs/removes them.
  // Strip all <p:graphicFrame> elements containing <p:oleObj> to prevent repair dialog.
  let s11xml = await _readXml("ppt/slides/slide11.xml");
  if (s11xml && s11xml.includes("<p:oleObj")) {
    // Collect rIds used by oleObj elements before removing them
    const oleRids = [...s11xml.matchAll(/<p:oleObj[^>]*\br:id="([^"]+)"/g)].map(m => m[1]);
    s11xml = s11xml.replace(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g, (match) => {
      return match.includes("<p:oleObj") ? "" : match;
    });
    _patch("ppt/slides/slide11.xml", s11xml);

    // Remove dangling rId relationships for the stripped OLE objects
    const s11RelsPath = "ppt/slides/_rels/slide11.xml.rels";
    let s11rels = await _readXml(s11RelsPath);
    if (s11rels && oleRids.length) {
      for (const rid of oleRids) {
        s11rels = s11rels.replace(
          new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*/>`), ""
        );
      }
      _patch(s11RelsPath, s11rels);
    }
  }

  // ── Strip pivotSource + externalData from ALL chart files ─────────────────
  const allChartFiles = Object.keys(zip.files).filter(
    k => k.startsWith("ppt/charts/chart") && k.endsWith(".xml")
  );
  for (const chartPath of allChartFiles) {
    if (patches.has(chartPath)) {
      // Already patched above — apply strip to the patched version
      let cxml = patches.get(chartPath).toString("utf8");
      cxml = cxml.replace(/<c:pivotSource>[\s\S]*?<\/c:pivotSource>/g, "");
      cxml = cxml.replace(/<c:externalData\b[\s\S]*?<\/c:externalData>/g, "");
      cxml = cxml.replace(/<c:externalData\b[^>]*\/>/g, "");
      _patch(chartPath, cxml);
    } else {
      const cf = zip.file(chartPath);
      if (cf) {
        let cxml = await cf.async("string");
        if (cxml.includes("<c:pivotSource>") || cxml.includes("<c:externalData")) {
          cxml = cxml.replace(/<c:pivotSource>[\s\S]*?<\/c:pivotSource>/g, "");
          cxml = cxml.replace(/<c:externalData\b[\s\S]*?<\/c:externalData>/g, "");
          cxml = cxml.replace(/<c:externalData\b[^>]*\/>/g, "");
          _patch(chartPath, cxml);
        }
      }
    }
  }

  // ── Slide 12 (Thank you / Key Takeaway): replace numbers with yellow-highlighted XX + subtitle ───
  let s12xml = await _readXml("ppt/slides/slide12.xml");
  if (s12xml) {
    // Replace "D@S - App Dev & Int - India" with "BTP , Global Delivery hub"
    s12xml = s12xml.replace(
      /D@S - App Dev &amp; Int - India/g,
      "BTP , Global Delivery hub",
    );

    // Replace the 4 numbered sentences: highlight the number as yellow "XX".
    // Approach: match the ENTIRE <a:r>...<a:rPr...>...</a:rPr><a:t>TEXT</a:t></a:r> block,
    // inject <a:highlight> into the existing <a:rPr>, and replace the number with XX inside
    // the existing <a:t> — NO run splitting, so XML structure stays perfectly balanced.
    const PATTERNS_12 = [
      { pat: /Review \d+ custom code objects which are reported to have pre-existing issues[^<]*/, numRe: /\d+/ },
      { pat: /Review \d+ custom code objects which involves deprecated functionalities[^<]*/,      numRe: /\d+/ },
      { pat: /Review \d+ clone programs[^<]*/,                                                     numRe: /\d+/ },
      { pat: /\d+ clone objects are impacted[^<]*/,                                                numRe: /\d+/ },
    ];
    for (const { pat, numRe } of PATTERNS_12) {
      // Match <a:r> ... <a:rPr (attrs)> (children) </a:rPr> <a:t>TARGET_TEXT</a:t> </a:r>
      // rprBody uses (?:(?!</?a:r[\\s>])[\\s\\S])*? to prevent crossing <a:r>/<\/a:r> boundaries,
      // which would cause the greedy match to span multiple runs and corrupt unrelated runs.
      s12xml = s12xml.replace(
        new RegExp(`(<a:r>\\s*<a:rPr)(\\b[^>]*>)((?:(?!</?a:r[\\s>])[\\s\\S])*?)(</a:rPr>\\s*<a:t>)(${pat.source})(</a:t>\\s*</a:r>)`),
        (fullMatch, rprOpen, rprAttrs, rprBody, midTag, text, closeTag) => {
          // Inject yellow highlight into <a:rPr> if not already present
          let newRprBody = rprBody;
          if (!newRprBody.includes("<a:highlight>")) {
            newRprBody = `<a:highlight><a:srgbClr val="${_YELLOW}"/></a:highlight>` + newRprBody;
          }
          // Replace number with XX inside the text
          const newText = text.replace(numRe, "XX");
          return `${rprOpen}${rprAttrs}${newRprBody}${midTag}${_escapeXml(newText)}${closeTag}`;
        }
      );
    }

    _patch("ppt/slides/slide12.xml", s12xml);
  }

  // ── Slide 14: remove image (p:pic element with rId2 = image) ─────────────
  let s14xml = await _readXml("ppt/slides/slide14.xml");
  if (s14xml && s14xml.includes("<p:pic")) {
    // Collect rIds of all pic elements before removing them
    const picRids = [...s14xml.matchAll(/<p:blipFill>[\s\S]*?r:embed="([^"]+)"/g)].map(m => m[1]);
    // Remove all <p:pic> elements
    s14xml = s14xml.replace(/<p:pic\b[\s\S]*?<\/p:pic>/g, "");
    _patch("ppt/slides/slide14.xml", s14xml);

    // Remove dangling image relationships from slide14.xml.rels
    const s14RelsPath = "ppt/slides/_rels/slide14.xml.rels";
    let s14rels = await _readXml(s14RelsPath);
    if (s14rels && picRids.length) {
      for (const rid of picRids) {
        s14rels = s14rels.replace(
          new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*/>`), ""
        );
      }
      _patch(s14RelsPath, s14rels);
    }
  }

  // ── Slide master footer: 2024 → 2026 ─────────────────────────────────────
  let smXml = await _readXml("ppt/slideMasters/slideMaster1.xml");
  if (smXml && smXml.includes("2024 SAP SE")) {
    smXml = smXml.replace(
      /2024 SAP SE or an SAP affiliate company\./g,
      "2026 SAP SE or an SAP affiliate company.",
    );
    _patch("ppt/slideMasters/slideMaster1.xml", smXml);
  }

  // ── Rebuild ZIP losslessly ────────────────────────────────────────────────
  return _rebuildZip(templateBuf, patches);
}

module.exports = { writePptx };

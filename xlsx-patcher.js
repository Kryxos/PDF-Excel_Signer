(function (global) {
  'use strict';

  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const DRAWING_REL = REL_NS + '/drawing';
  const IMAGE_REL = REL_NS + '/image';

  function getJSZip() {
    if (global.JSZip) return global.JSZip;
    if (typeof require === 'function') {
      try { return require('jszip'); } catch (_) {}
    }
    throw new Error('Brak biblioteki JSZip.');
  }

  function escapeXml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function decodeXml(value) {
    return String(value)
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&');
  }

  function nextRelationshipID(xmlText) {
    let max = 0;
    for (const match of xmlText.matchAll(/\bId="rId([0-9]+)"/g)) {
      max = Math.max(max, Number(match[1]) || 0);
    }
    return 'rId' + (max + 1);
  }

  function appendRelationshipXML(xmlText, id, relType, target) {
    if (!xmlText.includes('</Relationships>')) {
      throw new Error('Uszkodzony plik relacji XLSX.');
    }
    const entry = `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(relType)}" Target="${escapeXml(target)}"/>`;
    return xmlText.replace('</Relationships>', entry + '</Relationships>');
  }

  function dirname(path) {
    const p = path.replace(/\\/g, '/');
    const i = p.lastIndexOf('/');
    return i >= 0 ? p.slice(0, i) : '.';
  }

  function basename(path) {
    const p = path.replace(/\\/g, '/');
    const i = p.lastIndexOf('/');
    return i >= 0 ? p.slice(i + 1) : p;
  }

  function cleanPath(path) {
    const out = [];
    for (const part of path.replace(/\\/g, '/').split('/')) {
      if (!part || part === '.') continue;
      if (part === '..') out.pop();
      else out.push(part);
    }
    return out.join('/');
  }

  function zipRelationshipPath(partPath) {
    return cleanPath(dirname(partPath) + '/_rels/' + basename(partPath) + '.rels');
  }

  function resolveZipTarget(basePart, target) {
    target = String(target).replace(/\\/g, '/');
    if (target.startsWith('/')) return cleanPath(target.slice(1));
    return cleanPath(dirname(basePart) + '/' + target);
  }

  function parseCellZeroBased(cell) {
    const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(String(cell).trim().toUpperCase());
    if (!match) throw new Error('Nieprawidłowy adres komórki.');
    let col = 0;
    for (const ch of match[1]) col = col * 26 + ch.charCodeAt(0) - 64;
    const row = Number(match[2]);
    if (col < 1 || col > 16384 || row < 1 || row > 1048576) {
      throw new Error('Adres komórki jest poza zakresem Excela.');
    }
    return { col: col - 1, row: row - 1 };
  }

  function xmlFloatAttr(tag, attr, xmlText, fallback) {
    const re = new RegExp('<' + tag + '\\b[^>]*\\b' + attr + '="([^"]+)"');
    const m = re.exec(xmlText);
    if (!m) return fallback;
    const value = Number(m[1]);
    return Number.isFinite(value) ? value : fallback;
  }

  function attrFloat(tagText, attr, fallback) {
    const re = new RegExp('\\b' + attr + '="([^"]+)"');
    const m = re.exec(tagText);
    if (!m) return fallback;
    const value = Number(m[1]);
    return Number.isFinite(value) ? value : fallback;
  }

  function attrInt(tagText, attr, fallback) {
    const re = new RegExp('\\b' + attr + '="([0-9]+)"');
    const m = re.exec(tagText);
    if (!m) return fallback;
    const value = Number(m[1]);
    return Number.isFinite(value) ? Math.trunc(value) : fallback;
  }

  function cellPixelSize(sheetXML, col0, row0, googleStyle) {
    const defaultCol = xmlFloatAttr('sheetFormatPr', 'defaultColWidth', sheetXML, 8.43);
    const defaultRow = xmlFloatAttr('sheetFormatPr', 'defaultRowHeight', sheetXML, 15.0);

    let colWidth = defaultCol;
    for (const match of sheetXML.matchAll(/<col\b[^>]*>/g)) {
      const tag = match[0];
      const min = attrInt(tag, 'min', -1);
      const max = attrInt(tag, 'max', -1);
      if (min <= col0 + 1 && col0 + 1 <= max) {
        colWidth = attrFloat(tag, 'width', colWidth);
        break;
      }
    }

    let rowHeight = defaultRow;
    const rowRe = new RegExp('<row\\b[^>]*\\br="' + (row0 + 1) + '"[^>]*>');
    const rowMatch = rowRe.exec(sheetXML);
    if (rowMatch) rowHeight = attrFloat(rowMatch[0], 'ht', rowHeight);

    let widthPx = googleStyle ? Math.trunc(colWidth * 7.0) - 1 : Math.trunc(colWidth * 7.0 + 5.0);
    let heightPx = Math.trunc(rowHeight * 4.0 / 3.0);
    widthPx = Math.max(1, widthPx);
    heightPx = Math.max(1, heightPx);
    return { width: widthPx, height: heightPx };
  }

  function pngDimensions(data) {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    const sig = [137,80,78,71,13,10,26,10];
    if (bytes.length < 24 || !sig.every((v, i) => bytes[i] === v)) {
      throw new Error('Nieprawidłowy plik PNG.');
    }
    const w = ((bytes[16] << 24) >>> 0) + (bytes[17] << 16) + (bytes[18] << 8) + bytes[19];
    const h = ((bytes[20] << 24) >>> 0) + (bytes[21] << 16) + (bytes[22] << 8) + bytes[23];
    if (w <= 0 || h <= 0) throw new Error('Nieprawidłowe wymiary PNG.');
    return { width: w, height: h };
  }

  function ensureWorksheetRNamespace(sheetXML) {
    if (sheetXML.includes(`xmlns:r="${REL_NS}"`)) return sheetXML;
    return sheetXML.replace(/<worksheet\b[^>]*>/, root => root.slice(0, -1) + ` xmlns:r="${REL_NS}">`);
  }

  function insertWorksheetDrawingReference(sheetXML, rid) {
    sheetXML = ensureWorksheetRNamespace(sheetXML);
    const escapedRid = rid.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const refRe = new RegExp('<drawing\\b[^>]*\\br:id="' + escapedRid + '"[^>]*/>');
    if (refRe.test(sheetXML)) return sheetXML;

    const ref = `<drawing r:id="${escapeXml(rid)}"/>`;
    const following = [
      '<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects',
      '<controls', '<webPublishItems', '<tableParts', '<extLst', '</worksheet>'
    ];
    let insertAt = -1;
    for (const token of following) {
      const idx = sheetXML.indexOf(token);
      if (idx >= 0 && (insertAt < 0 || idx < insertAt)) insertAt = idx;
    }
    if (insertAt >= 0) return sheetXML.slice(0, insertAt) + ref + sheetXML.slice(insertAt);
    return sheetXML;
  }

  function ensureDrawingContentType(contentTypesXML, drawingPath) {
    let s = contentTypesXML;
    if (!/<Default\b[^>]*\bExtension="png"[^>]*\/>/.test(s)) {
      s = s.replace('</Types>', '<Default ContentType="image/png" Extension="png"/></Types>');
    }
    const partName = '/' + drawingPath;
    if (!s.includes(`PartName="${partName}"`)) {
      const entry = `<Override ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" PartName="${escapeXml(partName)}"/>`;
      s = s.replace('</Types>', entry + '</Types>');
    }
    return s;
  }

  function findDrawingRelationship(relsText) {
    for (const match of relsText.matchAll(/<Relationship\b([^>]*)\/>/g)) {
      const attrs = match[1];
      const id = /\bId="([^"]+)"/.exec(attrs)?.[1] || '';
      const type = /\bType="([^"]+)"/.exec(attrs)?.[1] || '';
      const target = /\bTarget="([^"]+)"/.exec(attrs)?.[1] || '';
      if (type.endsWith('/drawing')) return { id, target };
    }
    return { id: '', target: '' };
  }

  function maxNumberedPart(names, regex) {
    let max = 0;
    for (const name of names) {
      const m = regex.exec(name);
      regex.lastIndex = 0;
      if (m) max = Math.max(max, Number(m[1]) || 0);
    }
    return max;
  }

  function ensureDrawingRelationshipNamespace(drawingXML) {
    if (drawingXML.includes(`xmlns:r="${REL_NS}"`)) return drawingXML;
    const rootRe = /<xdr:wsDr\b[^>]*?\/?>/;
    const root = rootRe.exec(drawingXML)?.[0];
    if (!root) return drawingXML;
    const selfClosing = root.trim().endsWith('/>');
    let repl;
    if (selfClosing) {
      const base = root.trim().slice(0, -2).trimEnd();
      repl = base + ` xmlns:r="${REL_NS}"/>`;
    } else {
      repl = root.slice(0, -1) + ` xmlns:r="${REL_NS}">`;
    }
    return drawingXML.replace(root, repl);
  }

  function appendAnchorToDrawing(drawingXML, anchor) {
    if (drawingXML.includes('</xdr:wsDr>')) {
      return drawingXML.replace('</xdr:wsDr>', anchor + '</xdr:wsDr>');
    }
    const root = /<xdr:wsDr\b[^>]*?\/>/.exec(drawingXML)?.[0];
    if (root) {
      const open = root.trim().slice(0, -2).trimEnd() + '>';
      return drawingXML.replace(root, open + anchor + '</xdr:wsDr>');
    }
    throw new Error('Uszkodzony drawing XML: brak elementu xdr:wsDr.');
  }

  async function readText(zip, path) {
    const file = zip.file(path);
    if (!file) return null;
    return await file.async('string');
  }

  function resolveSheets(workbookXML, workbookRelsXML) {
    const ridTarget = new Map();
    for (const match of workbookRelsXML.matchAll(/<Relationship\b([^>]*)\/>/g)) {
      const attrs = match[1];
      const id = /\bId="([^"]+)"/.exec(attrs)?.[1];
      const target = /\bTarget="([^"]+)"/.exec(attrs)?.[1];
      if (id && target) ridTarget.set(id, decodeXml(target));
    }

    const out = new Map();
    for (const match of workbookXML.matchAll(/<sheet\b([^>]*)\/?\s*>/g)) {
      const attrs = match[1];
      const name = /\bname="([^"]*)"/.exec(attrs)?.[1];
      const rid = /\br:id="([^"]+)"/.exec(attrs)?.[1];
      if (!name || !rid) continue;
      let target = ridTarget.get(rid);
      if (!target) continue;
      target = target.replace(/\\/g, '/');
      if (target.startsWith('/')) target = target.slice(1);
      else if (!target.startsWith('xl/')) target = 'xl/' + target;
      out.set(decodeXml(name), cleanPath(target));
    }
    return out;
  }

  async function ensureSheetDrawing(zip, names, sheetPath) {
    let sheetXML = await readText(zip, sheetPath);
    if (sheetXML == null) throw new Error('Brak XML arkusza.');

    const relsPath = zipRelationshipPath(sheetPath);
    let relsText = await readText(zip, relsPath);
    if (relsText == null) relsText = `<Relationships xmlns="${PKG_REL_NS}"></Relationships>`;

    let { id: drawRID, target: drawTarget } = findDrawingRelationship(relsText);
    let drawingPath;

    if (drawRID) {
      drawingPath = resolveZipTarget(sheetPath, drawTarget);
      const escapedRid = drawRID.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const refRe = new RegExp('<drawing\\b[^>]*\\br:id="' + escapedRid + '"[^>]*/>');
      if (!refRe.test(sheetXML)) {
        sheetXML = insertWorksheetDrawingReference(sheetXML, drawRID);
        zip.file(sheetPath, sheetXML);
      }
    } else {
      const drawingNo = maxNumberedPart(names, /^xl\/drawings\/drawing([0-9]+)\.xml$/) + 1;
      drawingPath = `xl/drawings/drawing${drawingNo}.xml`;
      drawRID = nextRelationshipID(relsText);
      relsText = appendRelationshipXML(relsText, drawRID, DRAWING_REL, `../drawings/drawing${drawingNo}.xml`);
      zip.file(relsPath, relsText);
      if (!names.includes(relsPath)) names.push(relsPath);

      sheetXML = insertWorksheetDrawingReference(sheetXML, drawRID);
      zip.file(sheetPath, sheetXML);

      const drawingXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r` +
        `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:x3Unk="http://schemas.microsoft.com/office/drawing/2010/slicer" xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer"></xdr:wsDr>`;
      zip.file(drawingPath, drawingXML);
      if (!names.includes(drawingPath)) names.push(drawingPath);

      const drawingRelsPath = zipRelationshipPath(drawingPath);
      zip.file(drawingRelsPath, `<Relationships xmlns="${PKG_REL_NS}"></Relationships>`);
      if (!names.includes(drawingRelsPath)) names.push(drawingRelsPath);

      let contentTypes = await readText(zip, '[Content_Types].xml');
      if (contentTypes == null) throw new Error('Brak [Content_Types].xml.');
      contentTypes = ensureDrawingContentType(contentTypes, drawingPath);
      zip.file('[Content_Types].xml', contentTypes);
    }

    const drawingRelsPath = zipRelationshipPath(drawingPath);
    let drawingXML = await readText(zip, drawingPath);
    if (drawingXML == null) throw new Error(`Relacja arkusza wskazuje brakujący drawing: ${drawingPath}`);
    let drawingRels = await readText(zip, drawingRelsPath);
    if (drawingRels == null) {
      drawingRels = `<Relationships xmlns="${PKG_REL_NS}"></Relationships>`;
      zip.file(drawingRelsPath, drawingRels);
      if (!names.includes(drawingRelsPath)) names.push(drawingRelsPath);
    }

    drawingXML = ensureDrawingRelationshipNamespace(drawingXML);
    zip.file(drawingPath, drawingXML);
    return { drawingPath, drawingRelsPath };
  }

  function dataUrlToBytes(dataUrl) {
    const comma = dataUrl.indexOf(',');
    const raw = comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl;
    if (typeof atob === 'function') {
      const bin = atob(raw);
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(raw, 'base64'));
    throw new Error('Brak dekodera Base64.');
  }

  async function verifyInsertions(zipBytes, requests) {
    const JSZip = getJSZip();
    const zip = await JSZip.loadAsync(zipBytes);
    const workbookXML = await readText(zip, 'xl/workbook.xml');
    const workbookRelsXML = await readText(zip, 'xl/_rels/workbook.xml.rels');
    if (!workbookXML || !workbookRelsXML) throw new Error('Walidacja: brak struktury skoroszytu.');
    const sheets = resolveSheets(workbookXML, workbookRelsXML);

    for (const rq of requests) {
      const sheetPath = sheets.get(rq.sheet);
      if (!sheetPath) throw new Error(`Walidacja: brak arkusza "${rq.sheet}".`);
      const relsText = await readText(zip, zipRelationshipPath(sheetPath)) || '';
      const drawingRel = findDrawingRelationship(relsText);
      if (!drawingRel.id || !drawingRel.target) throw new Error(`Walidacja: arkusz ${rq.sheet} nie ma relacji drawing.`);
      const drawingPath = resolveZipTarget(sheetPath, drawingRel.target);
      const drawingXML = await readText(zip, drawingPath);
      if (!drawingXML) throw new Error(`Walidacja: brak ${drawingPath}.`);
      if (drawingXML.includes('r:embed=') && !drawingXML.includes(`xmlns:r="${REL_NS}"`)) {
        throw new Error(`Walidacja: ${drawingPath} używa r:embed bez deklaracji xmlns:r.`);
      }
      const { col, row } = parseCellZeroBased(rq.cell);
      const want = `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`;
      if (!drawingXML.includes(want)) throw new Error(`Walidacja: obraz nie został zakotwiczony w ${rq.sheet}!${rq.cell}.`);
    }
  }

  async function patchXlsxInCell(inputBytes, operations) {
    const JSZip = getJSZip();
    const zip = await JSZip.loadAsync(inputBytes);
    const names = Object.keys(zip.files).filter(name => !zip.files[name].dir);

    const contentTypes = await readText(zip, '[Content_Types].xml');
    const workbookXML = await readText(zip, 'xl/workbook.xml');
    const workbookRelsXML = await readText(zip, 'xl/_rels/workbook.xml.rels');
    if (!contentTypes || !workbookXML || !workbookRelsXML) {
      throw new Error('XLSX nie ma wymaganej struktury skoroszytu.');
    }

    const sheetMap = resolveSheets(workbookXML, workbookRelsXML);
    const googleStyle = workbookXML.includes('GoogleSheetsCustomDataVersion');
    let maxImage = maxNumberedPart(names, /^xl\/media\/image([0-9]+)\.[A-Za-z0-9]+$/);

    const normalizedOps = operations.map(op => ({
      sheet: String(op.sheet || '').trim(),
      cell: String(op.cell || '').trim().toUpperCase(),
      png: op.png instanceof Uint8Array ? op.png : dataUrlToBytes(op.pngDataUrl || op.png || '')
    }));

    for (const rq of normalizedOps) {
      if (!rq.sheet) throw new Error('Nie wybrano arkusza.');
      const sheetPath = sheetMap.get(rq.sheet);
      if (!sheetPath) throw new Error(`Nie znaleziono arkusza "${rq.sheet}".`);
      const { col, row } = parseCellZeroBased(rq.cell);
      const dims = pngDimensions(rq.png);
      const sheetXML = await readText(zip, sheetPath);
      if (sheetXML == null) throw new Error(`Nie znaleziono XML arkusza "${rq.sheet}".`);
      const cell = cellPixelSize(sheetXML, col, row, googleStyle);

      const scale = Math.min(cell.width / dims.width, cell.height / dims.height);
      const displayW = Math.max(1, Math.trunc(dims.width * scale));
      const displayH = Math.max(1, Math.trunc(dims.height * scale));
      const cx = displayW * 9525;
      const cy = displayH * 9525;

      const { drawingPath, drawingRelsPath } = await ensureSheetDrawing(zip, names, sheetPath);

      maxImage += 1;
      const imagePath = `xl/media/image${maxImage}.png`;
      zip.file(imagePath, rq.png, { binary: true });
      if (!names.includes(imagePath)) names.push(imagePath);

      let drawRels = await readText(zip, drawingRelsPath);
      if (drawRels == null) drawRels = `<Relationships xmlns="${PKG_REL_NS}"></Relationships>`;
      const imageRID = nextRelationshipID(drawRels);
      drawRels = appendRelationshipXML(drawRels, imageRID, IMAGE_REL, `../media/image${maxImage}.png`);
      zip.file(drawingRelsPath, drawRels);

      let drawingXML = await readText(zip, drawingPath);
      if (drawingXML == null) throw new Error('Uszkodzony drawing XML.');
      let picID = 0;
      for (const m of drawingXML.matchAll(/<xdr:cNvPr\b[^>]*\bid="([0-9]+)"/g)) {
        const n = Number(m[1]) || 0;
        if (n >= picID) picID = n + 1;
      }

      const anchor = `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${picID}" name="image${maxImage}.png"/><xdr:cNvPicPr preferRelativeResize="0"/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${REL_NS}" cstate="print" r:embed="${imageRID}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr></xdr:pic><xdr:clientData fLocksWithSheet="0"/></xdr:oneCellAnchor>`;
      drawingXML = appendAnchorToDrawing(drawingXML, anchor);
      zip.file(drawingPath, drawingXML);

      let ct = await readText(zip, '[Content_Types].xml');
      ct = ensureDrawingContentType(ct, drawingPath);
      zip.file('[Content_Types].xml', ct);
    }

    const out = await zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    await verifyInsertions(out, normalizedOps);
    return out;
  }

  global.XLSXInCell = {
    patchXlsxInCell,
    parseCellZeroBased,
    dataUrlToBytes
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = global.XLSXInCell;
  }
})(typeof globalThis !== 'undefined' ? globalThis : this);

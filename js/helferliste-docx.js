/**
 * SCHULHELFER · Shared Helferliste .docx builder.
 *
 * A .docx file is a ZIP of OOXML parts. This module packages one
 * client-side using a minimal STORE-only ZIP writer (no compression,
 * no external library). Word, Pages, LibreOffice and Google Docs all
 * open the result natively.
 *
 * Used by:
 *   - Public site (js/app.js) when an admin opens with ?admin=KEY.
 *   - Admin dashboard (admin/admin.js) inside the per-event Helfer
 *     drawer.
 *
 * Public surface (attached to `window.HelferListeDocx`):
 *   build(event, helpers)         → Uint8Array (.docx bytes)
 *   download(bytes, filename)     → triggers browser file save
 *   filenameFor(event)            → "Helferliste_<name>_<stamp>.docx"
 *
 * `event` shape (only the listed fields are read):
 *   { name, datum, zeit, beschreibung, maxHelfer }
 * `helpers` is an array of { name, telefon } objects (other fields
 * ignored).
 */
(function () {
  'use strict';

  // ---- CRC-32 (IEEE 802.3 polynomial) ---------------------------------
  const CRC32_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) {
      c = CRC32_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // ---- ZIP store-only writer ------------------------------------------
  // files: [{ name: 'path/in/zip', data: string|Uint8Array }]
  function zipStore(files) {
    const encoder = new TextEncoder();
    const localChunks = [];
    const centralChunks = [];
    let offset = 0;

    for (const f of files) {
      const nameBytes = encoder.encode(f.name);
      const data = typeof f.data === 'string' ? encoder.encode(f.data) : f.data;
      const crc = crc32(data);
      const size = data.length;

      const lfh = new Uint8Array(30);
      const lv = new DataView(lfh.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0, true);
      lv.setUint16(8, 0, true);
      lv.setUint16(10, 0, true);
      lv.setUint16(12, 0x21, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, size, true);
      lv.setUint32(22, size, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      localChunks.push(lfh, nameBytes, data);

      const cdh = new Uint8Array(46);
      const cv = new DataView(cdh.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0x21, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      centralChunks.push(cdh, nameBytes);
      offset += 30 + nameBytes.length + size;
    }

    const localTotal = offset;
    let centralTotal = 0;
    for (const c of centralChunks) centralTotal += c.length;

    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, centralTotal, true);
    ev.setUint32(16, localTotal, true);
    ev.setUint16(20, 0, true);

    const out = new Uint8Array(localTotal + centralTotal + 22);
    let pos = 0;
    for (const c of localChunks) { out.set(c, pos); pos += c.length; }
    for (const c of centralChunks) { out.set(c, pos); pos += c.length; }
    out.set(eocd, pos);
    return out;
  }

  // ---- WordprocessingML helpers ---------------------------------------
  function xmlEscape(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  const DEFAULT_FONT_RPR = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>';

  function rPrBody(opts) {
    const parts = [DEFAULT_FONT_RPR];
    if (opts.bold)        parts.push('<w:b/><w:bCs/>');
    if (opts.italic)      parts.push('<w:i/><w:iCs/>');
    if (opts.caps)        parts.push('<w:caps/>');
    if (opts.size)        parts.push('<w:sz w:val="' + opts.size + '"/><w:szCs w:val="' + opts.size + '"/>');
    if (opts.charSpacing) parts.push('<w:spacing w:val="' + opts.charSpacing + '"/>');
    if (opts.color)       parts.push('<w:color w:val="' + opts.color + '"/>');
    return parts.join('');
  }

  function wPara(text, opts) {
    opts = opts || {};
    const rPr = '<w:rPr>' + rPrBody(opts) + '</w:rPr>';
    const pProps = [];
    if (opts.align) pProps.push('<w:jc w:val="' + opts.align + '"/>');
    const spacingAttrs = [];
    if (opts.spacingBefore != null) spacingAttrs.push('w:before="' + opts.spacingBefore + '"');
    if (opts.spacingAfter != null)  spacingAttrs.push('w:after="' + opts.spacingAfter + '"');
    if (opts.lineHeight != null) {
      spacingAttrs.push('w:line="' + opts.lineHeight + '"');
      spacingAttrs.push('w:lineRule="auto"');
    }
    if (spacingAttrs.length) pProps.push('<w:spacing ' + spacingAttrs.join(' ') + '/>');
    const pPr = pProps.length ? '<w:pPr>' + pProps.join('') + '</w:pPr>' : '';
    return '<w:p>' + pPr + '<w:r>' + rPr +
      '<w:t xml:space="preserve">' + xmlEscape(text) + '</w:t>' +
      '</w:r></w:p>';
  }

  function wCell(text, opts) {
    opts = opts || {};
    const width = opts.width || 2000;
    const shading = opts.shading
      ? '<w:shd w:val="clear" w:color="auto" w:fill="' + opts.shading + '"/>'
      : '';
    const vAlign = '<w:vAlign w:val="' + (opts.vAlign || 'center') + '"/>';
    const tcBorders = opts.borders || '';
    const tcPr = '<w:tcPr><w:tcW w:w="' + width + '" w:type="dxa"/>' +
      shading + tcBorders + vAlign + '</w:tcPr>';
    const runOpts = {
      bold: opts.bold, italic: opts.italic, caps: opts.caps,
      charSpacing: opts.charSpacing, size: opts.size,
      color: opts.textColor || opts.color
    };
    const rPr = '<w:rPr>' + rPrBody(runOpts) + '</w:rPr>';
    const pProps = [];
    if (opts.align) pProps.push('<w:jc w:val="' + opts.align + '"/>');
    pProps.push('<w:spacing w:before="0" w:after="0"/>');
    const pPr = '<w:pPr>' + pProps.join('') + '</w:pPr>';
    return '<w:tc>' + tcPr + '<w:p>' + pPr + '<w:r>' + rPr +
      '<w:t xml:space="preserve">' + xmlEscape(text || '') + '</w:t>' +
      '</w:r></w:p></w:tc>';
  }

  // ---- Document body --------------------------------------------------
  function formatTimestamp(date) {
    const d = date instanceof Date ? date : new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '.' + d.getFullYear() +
      ', ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ' Uhr';
  }

  function buildDocumentXml(event, helpers) {
    helpers = Array.isArray(helpers) ? helpers : [];
    const maxHelfer = parseInt(event.maxHelfer, 10) || helpers.length;
    const rowCount = Math.max(maxHelfer, helpers.length);

    const COL_NR   = 700;
    const COL_NAME = 5600;
    const COL_TEL  = 3338;

    const headerRow = '<w:tr>' +
      '<w:trPr><w:tblHeader/><w:trHeight w:val="520" w:hRule="atLeast"/></w:trPr>' +
      wCell('Nr.',     { width: COL_NR,   bold: true, size: 22, caps: true, charSpacing: 20, align: 'center' }) +
      wCell('Name',    { width: COL_NAME, bold: true, size: 22, caps: true, charSpacing: 20 }) +
      wCell('Telefon', { width: COL_TEL,  bold: true, size: 22, caps: true, charSpacing: 20 }) +
      '</w:tr>';

    const dataRows = [];
    for (let i = 0; i < rowCount; i++) {
      const h = helpers[i] || {};
      dataRows.push('<w:tr>' +
        '<w:trPr><w:trHeight w:val="520" w:hRule="atLeast"/></w:trPr>' +
        wCell(String(i + 1), { width: COL_NR,   size: 22, align: 'center' }) +
        wCell(h.name || '',   { width: COL_NAME, size: 22 }) +
        wCell(h.telefon || '', { width: COL_TEL,  size: 22 }) +
        '</w:tr>');
    }

    const tableProps = '<w:tblPr>' +
      '<w:tblW w:w="5000" w:type="pct"/>' +
      '<w:tblInd w:w="0" w:type="dxa"/>' +
      '<w:tblBorders>' +
        '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
        '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="B8B8B8"/>' +
      '</w:tblBorders>' +
      '<w:tblCellMar>' +
        '<w:top w:w="120" w:type="dxa"/>' +
        '<w:bottom w:w="120" w:type="dxa"/>' +
        '<w:left w:w="160" w:type="dxa"/>' +
        '<w:right w:w="160" w:type="dxa"/>' +
      '</w:tblCellMar>' +
      '</w:tblPr>';
    const tableGrid = '<w:tblGrid>' +
      '<w:gridCol w:w="' + COL_NR + '"/>' +
      '<w:gridCol w:w="' + COL_NAME + '"/>' +
      '<w:gridCol w:w="' + COL_TEL + '"/>' +
      '</w:tblGrid>';

    const metaParts = [];
    if (event.datum) metaParts.push(event.datum);
    if (event.zeit)  metaParts.push(event.zeit);
    const metaLine = metaParts.join('  ·  ');

    const bodyParts = [];
    bodyParts.push(wPara('Helferliste', {
      bold: true, size: 22, caps: true, charSpacing: 40, spacingAfter: 80
    }));
    bodyParts.push(wPara(event.name || '', {
      bold: true, size: 48, lineHeight: 260, spacingAfter: 160
    }));
    if (metaLine) {
      bodyParts.push(wPara(metaLine, {
        italic: true, size: 24, spacingAfter: 100
      }));
    }
    if (event.beschreibung) {
      bodyParts.push(wPara(String(event.beschreibung), {
        size: 22, lineHeight: 300, spacingAfter: 200
      }));
    } else {
      bodyParts.push(wPara('', { size: 8, spacingAfter: 80 }));
    }
    bodyParts.push(wPara(
      'Angemeldet: ' + helpers.length + ' von ' + maxHelfer +
        '  ·  Primarstufe Rittergasse Basel',
      { size: 18, spacingAfter: 40 }
    ));
    bodyParts.push(wPara(
      'Stand: ' + formatTimestamp(new Date()),
      { size: 18, italic: true, spacingAfter: 360 }
    ));
    bodyParts.push('<w:tbl>' + tableProps + tableGrid + headerRow + dataRows.join('') + '</w:tbl>');
    bodyParts.push('<w:p/>');
    bodyParts.push(
      '<w:sectPr>' +
        '<w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>' +
      '</w:sectPr>'
    );

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body>' + bodyParts.join('') + '</w:body>' +
      '</w:document>';
  }

  // ---- Public API -----------------------------------------------------
  function build(event, helpers) {
    const contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>';

    const rootRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>';

    return zipStore([
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'word/document.xml', data: buildDocumentXml(event, helpers) }
    ]);
  }

  function download(bytes, filename) {
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.rel = 'noopener';
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    // Defer cleanup: revoking synchronously can cancel the download
    // before the browser dispatches it (notably iOS Safari).
    setTimeout(function () {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 1000);
  }

  function filenameFor(event) {
    const safeName = String((event && event.name) || 'Anlass').replace(/[^a-z0-9äöüÄÖÜß]/gi, '_');
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const stamp = now.getFullYear() + pad(now.getMonth() + 1) + pad(now.getDate()) +
      '_' + pad(now.getHours()) + pad(now.getMinutes());
    return 'Helferliste_' + safeName + '_' + stamp + '.docx';
  }

  window.HelferListeDocx = { build: build, download: download, filenameFor: filenameFor };
})();

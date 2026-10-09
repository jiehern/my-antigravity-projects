/**
 * timesheet-template.js
 * 100% Authentic Coach Timesheet Excel (.xlsx) Generator
 * Faithful replica of 'JieHern_CoachTimesheet September2026.xlsx'
 * Preserves exact XML structures, cell borders (thin all-around, double bottom on fees),
 * Aptos Narrow & Arial fonts, numFmts (d-mmm, h:mm am/pm), column widths, and formulas.
 */

// Embedded static XML parts from the authentic template
const TEMPLATE_PARTS = {
  "xl/drawings/drawing1.xml": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:x3Unk=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" xmlns:sle15=\"http://schemas.microsoft.com/office/drawing/2012/slicer\"/>"
  },
  "xl/worksheets/_rels/sheet1.xml.rels": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/></Relationships>"
  },
  "docProps/core.xml": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dcterms:created xsi:type=\"dcterms:W3CDTF\">2026-01-08T22:35:10Z</dcterms:created><dc:creator>Christine Chow Pek Har</dc:creator></cp:coreProperties>"
  },
  "xl/theme/theme1.xml": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" name=\"Sheets\"><a:themeElements><a:clrScheme name=\"Sheets\"><a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"000000\"/></a:dk2><a:lt2><a:srgbClr val=\"FFFFFF\"/></a:lt2><a:accent1><a:srgbClr val=\"156082\"/></a:accent1><a:accent2><a:srgbClr val=\"E97132\"/></a:accent2><a:accent3><a:srgbClr val=\"196B24\"/></a:accent3><a:accent4><a:srgbClr val=\"0F9ED5\"/></a:accent4><a:accent5><a:srgbClr val=\"A02B93\"/></a:accent5><a:accent6><a:srgbClr val=\"4EA72E\"/></a:accent6><a:hlink><a:srgbClr val=\"467886\"/></a:hlink><a:folHlink><a:srgbClr val=\"467886\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Sheets\"><a:majorFont><a:latin typeface=\"Aptos Narrow\"/><a:ea typeface=\"Aptos Narrow\"/><a:cs typeface=\"Aptos Narrow\"/></a:majorFont><a:minorFont><a:latin typeface=\"Aptos Narrow\"/><a:ea typeface=\"Aptos Narrow\"/><a:cs typeface=\"Aptos Narrow\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"
  },
  "xl/styles.xml": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"d-mmm\"/><numFmt numFmtId=\"165\" formatCode=\"h:mm am/pm\"/></numFmts><fonts count=\"6\"><font><sz val=\"11.0\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/><scheme val=\"minor\"/></font><font><color theme=\"1\"/><name val=\"Aptos Narrow\"/><scheme val=\"minor\"/></font><font><sz val=\"11.0\"/><color theme=\"1\"/><name val=\"Arial\"/></font><font><sz val=\"11.0\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/></font><font><b/><sz val=\"11.0\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/></font><font><b/><sz val=\"11.0\"/><color theme=\"1\"/><name val=\"Arial\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"lightGray\"/></fill></fills><borders count=\"3\"><border/><border><left style=\"thin\"><color rgb=\"FF000000\"/></left><right style=\"thin\"><color rgb=\"FF000000\"/></right><top style=\"thin\"><color rgb=\"FF000000\"/></top><bottom style=\"thin\"><color rgb=\"FF000000\"/></bottom></border><border><top style=\"thin\"><color rgb=\"FF000000\"/></top><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border></borders><cellStyleXfs count=\"1\"><xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" applyAlignment=\"1\" applyFont=\"1\"/></cellStyleXfs><cellXfs count=\"24\"><xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyFont=\"1\"><alignment readingOrder=\"0\" shrinkToFit=\"0\" vertical=\"bottom\" wrapText=\"0\"/></xf><xf borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf borderId=\"1\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" readingOrder=\"0\"/></xf><xf borderId=\"0\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf borderId=\"0\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyFont=\"1\"><alignment horizontal=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"4\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"4\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"2\" numFmtId=\"164\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\" applyNumberFormat=\"1\"><alignment horizontal=\"center\" readingOrder=\"0\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"2\" numFmtId=\"165\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\" applyNumberFormat=\"1\"><alignment horizontal=\"center\" readingOrder=\"0\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyBorder=\"1\" applyFont=\"1\"/><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\"/></xf><xf borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf borderId=\"1\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment readingOrder=\"0\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" readingOrder=\"0\" vertical=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment readingOrder=\"0\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" readingOrder=\"0\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" readingOrder=\"0\" vertical=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment vertical=\"bottom\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"16\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\" applyNumberFormat=\"1\"><alignment horizontal=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"3\" numFmtId=\"16\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\" applyNumberFormat=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf borderId=\"1\" fillId=\"0\" fontId=\"5\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment readingOrder=\"0\"/></xf><xf borderId=\"0\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" readingOrder=\"0\"/></xf><xf borderId=\"2\" fillId=\"0\" fontId=\"3\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyBorder=\"1\" applyFont=\"1\"><alignment horizontal=\"center\"/></xf></cellXfs><cellStyles count=\"1\"><cellStyle xfId=\"0\" name=\"Normal\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/></styleSheet>"
  },
  "xl/workbook.xml": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr/><sheets><sheet state=\"visible\" name=\"Master\" sheetId=\"1\" r:id=\"rId4\"/></sheets><definedNames/><calcPr/></workbook>"
  },
  "xl/_rels/workbook.xml.rels": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"
  },
  "_rels/.rels": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
  },
  "[Content_Types].xml": {
    "type": "utf8",
    "data": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default ContentType=\"application/xml\" Extension=\"xml\"/><Default ContentType=\"application/vnd.openxmlformats-officedocument.obfuscatedFont\" Extension=\"odttf\"/><Default ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" Extension=\"rels\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" PartName=\"/xl/worksheets/sheet1.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" PartName=\"/xl/sharedStrings.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\" PartName=\"/xl/drawings/drawing1.xml\"/><Override ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" PartName=\"/docProps/core.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" PartName=\"/xl/styles.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\" PartName=\"/xl/theme/theme1.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" PartName=\"/xl/workbook.xml\"/></Types>"
  }
};

/**
 * Pure Web-standard ZIP Builder (RFC 1951 / PKZIP 2.0)
 * Runs entirely in browser without external dependencies.
 */
export function createZipBlob(filesMap) {
  const textEncoder = new TextEncoder();
  let offset = 0;

  // Pre-computed CRC-32 table
  const crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    crcTable[i] = c >>> 0;
  }

  function getCrc32(bytes) {
    let crc = 0 ^ (-1);
    for (let i = 0; i < bytes.length; i++) {
      crc = (crc >>> 8) ^ crcTable[(crc ^ bytes[i]) & 0xFF];
    }
    return (crc ^ (-1)) >>> 0;
  }

  const localParts = [];
  const cdParts = [];

  for (const [name, content] of Object.entries(filesMap)) {
    let dataBytes;
    if (content instanceof Uint8Array) {
      dataBytes = content;
    } else if (typeof content === 'string') {
      dataBytes = textEncoder.encode(content);
    } else {
      continue;
    }

    const nameBytes = textEncoder.encode(name);
    const crc = getCrc32(dataBytes);
    const size = dataBytes.length;

    // Local Header (30 bytes + name length)
    const lh = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(lh.buffer);
    lv.setUint32(0, 0x04034b50, true); // signature
    lv.setUint16(4, 20, true);         // version needed
    lv.setUint16(6, 0, true);          // flags
    lv.setUint16(8, 0, true);          // compression (store)
    lv.setUint16(10, 0, true);         // time
    lv.setUint16(12, 0, true);         // date
    lv.setUint32(14, crc, true);       // crc32
    lv.setUint32(18, size, true);      // compressed size
    lv.setUint32(22, size, true);      // uncompressed size
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);         // extra len
    lh.set(nameBytes, 30);

    localParts.push(lh, dataBytes);

    // Central Directory Header (46 bytes + name length)
    const cd = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true); // signature
    cv.setUint16(4, 20, true);         // version made by
    cv.setUint16(6, 20, true);         // version needed
    cv.setUint16(8, 0, true);          // flags
    cv.setUint16(10, 0, true);         // compression
    cv.setUint16(12, 0, true);         // time
    cv.setUint16(14, 0, true);         // date
    cv.setUint32(16, crc, true);       // crc32
    cv.setUint32(20, size, true);      // compressed size
    cv.setUint32(24, size, true);      // uncompressed size
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);         // extra len
    cv.setUint16(32, 0, true);         // comment len
    cv.setUint16(34, 0, true);         // disk
    cv.setUint16(36, 0, true);         // internal attr
    cv.setUint32(38, 0, true);         // external attr
    cv.setUint32(42, offset, true);    // relative offset
    cd.set(nameBytes, 46);

    cdParts.push(cd);
    offset += lh.length + size;
  }

  const cdTotalSize = cdParts.reduce((acc, p) => acc + p.length, 0);
  const cdOffset = offset;
  const numEntries = localParts.length / 2;

  // EOCD (22 bytes)
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, numEntries, true);
  ev.setUint16(10, numEntries, true);
  ev.setUint32(12, cdTotalSize, true);
  ev.setUint32(16, cdOffset, true);
  ev.setUint16(20, 0, true);

  return new Blob([...localParts, ...cdParts, eocd], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

/**
 * Calculate the class credit units / hours for a session:
 * - Pre Comp classes (or 90m / 1.5h): 1.5 units (= RM60 at RM40 rate)
 * - LTS classes (50m): 1.0 unit (= RM40 at RM40 rate)
 * - 60m: 1.0 unit
 * - 120m: 2.0 units
 * - 30m: 0.5 units
 * - 45m: 0.75 units
 * - 75m: 1.25 units
 * - General: duration / 60
 */
export function getClassCreditUnits(c) {
  if (!c) return 1.0;
  const note = (c.note || '').toLowerCase();
  const cat = (c.category || '').toLowerCase();
  const dur = Number(c.duration) || 60;

  if (note.includes('pre comp') || note.includes('precomp') || cat.includes('pre comp') || cat.includes('precomp') || dur === 90) {
    return 1.5;
  }
  if (note.includes('lts') || dur === 50) {
    return 1.0;
  }
  if (dur === 120) return 2.0;
  if (dur === 30) return 0.5;
  if (dur === 45) return 0.75;
  if (dur === 75) return 1.25;
  if (dur === 60) return 1.0;
  return Math.round((dur / 60) * 10) / 10;
}

/**
 * Manager Bonus System:
 * 24 hrs or class → RM100
 * 30 hrs or class → RM150
 * 40 hrs or class → RM250
 * 50 hrs or class → RM350
 */
export function calculateManagerBonus(totalUnits) {
  const units = Number(totalUnits) || 0;
  if (units >= 50) {
    return {
      bonus: 350,
      tier: 4,
      threshold: 50,
      nextThreshold: null,
      needed: 0,
      label: '≥ 50 hrs tier (RM350 max)',
      tierName: 'Tier 4'
    };
  }
  if (units >= 40) {
    return {
      bonus: 250,
      tier: 3,
      threshold: 40,
      nextThreshold: 50,
      needed: Math.round((50 - units) * 10) / 10,
      label: '≥ 40 hrs tier (RM250)',
      tierName: 'Tier 3'
    };
  }
  if (units >= 30) {
    return {
      bonus: 150,
      tier: 2,
      threshold: 30,
      nextThreshold: 40,
      needed: Math.round((40 - units) * 10) / 10,
      label: '≥ 30 hrs tier (RM150)',
      tierName: 'Tier 2'
    };
  }
  if (units >= 24) {
    return {
      bonus: 100,
      tier: 1,
      threshold: 24,
      nextThreshold: 30,
      needed: Math.round((30 - units) * 10) / 10,
      label: '≥ 24 hrs tier (RM100)',
      tierName: 'Tier 1'
    };
  }
  return {
    bonus: 0,
    tier: 0,
    threshold: 0,
    nextThreshold: 24,
    needed: Math.round((24 - units) * 10) / 10,
    label: 'Below tier threshold (< 24 hrs)',
    tierName: 'No Tier'
  };
}

/**
 * Generate 100% Authentic Coach Timesheet .xlsx
 */
export function exportCoachTimesheetXLSX(classes, options = {}) {
  const coachName = options.coachName !== undefined ? options.coachName : '';
  const ratePerClass = options.ratePerClass !== undefined ? Number(options.ratePerClass) : 40.0;

  // Sort chronologically
  const sortedClasses = classes.slice().sort((a, b) => {
    const tsA = a.timestamp || new Date(a.date + 'T' + (a.time || '00:00')).getTime();
    const tsB = b.timestamp || new Date(b.date + 'T' + (b.time || '00:00')).getTime();
    return tsA - tsB;
  });

  // Track shared strings
  const sharedStrings = [];
  const stringIndexMap = new Map();
  let totalStringRefs = 0;

  function getStrIdx(str) {
    totalStringRefs++;
    const s = String(str);
    if (!stringIndexMap.has(s)) {
      stringIndexMap.set(s, sharedStrings.length);
      sharedStrings.push(s);
    }
    return stringIndexMap.get(s);
  }

  // Pre-register standard strings in identical order to template
  const idxName = getStrIdx('Name');
  const idxColon = getStrIdx(':');
  let c1CellXml = `<c r="C1" s="2"/>`;
  if (coachName) {
    const idxCoach = getStrIdx(coachName);
    c1CellXml = `<c r="C1" s="2" t="s"><v>${idxCoach}</v></c>`;
  }
  const idxDate = getStrIdx('Date');
  const idxDay = getStrIdx('Day');
  const idxTime = getStrIdx('Time');
  const idxDuration = getStrIdx('Duration');
  const idxDescription = getStrIdx('Description');
  const idxTotalClass = getStrIdx('Total Class');
  const idxRemarks = getStrIdx('Remarks');

  const rowsXml = [];

  // Row 1: Coach Name
  rowsXml.push(`<row r="1" ht="14.25" customHeight="1"><c r="A1" s="1" t="s"><v>${idxName}</v></c><c r="B1" s="1" t="s"><v>${idxColon}</v></c>${c1CellXml}<c r="D1" s="3"/><c r="F1" s="4"/></row>`);

  // Rows 2 & 3: Empty spacer rows matching reference template
  rowsXml.push(`<row r="2" ht="14.25" customHeight="1"><c r="C2" s="4"/><c r="D2" s="3"/><c r="F2" s="4"/></row>`);
  rowsXml.push(`<row r="3" ht="14.25" customHeight="1"><c r="C3" s="4"/><c r="D3" s="3"/><c r="F3" s="4"/></row>`);

  // Row 4: Table Header with Aptos Narrow bold and thin borders
  rowsXml.push(`<row r="4" ht="14.25" customHeight="1"><c r="A4" s="5" t="s"><v>${idxDate}</v></c><c r="B4" s="5" t="s"><v>${idxDay}</v></c><c r="C4" s="5" t="s"><v>${idxTime}</v></c><c r="D4" s="6" t="s"><v>${idxDuration}</v></c><c r="E4" s="5" t="s"><v>${idxDescription}</v></c><c r="F4" s="5" t="s"><v>${idxTotalClass}</v></c><c r="G4" s="5" t="s"><v>${idxRemarks}</v></c></row>`);

  const epoch = new Date(Date.UTC(1899, 11, 30));
  let detectedMonth = '';
  let detectedYear = '';

  // Data rows starting from Row 5
  sortedClasses.forEach((c, idx) => {
    const r = 5 + idx;

    // 1. Date Serial
    const [y, m, d] = (c.date || '2026-09-01').split('-').map(Number);
    const dateObj = new Date(Date.UTC(y, m - 1, d));
    const serial = Math.round((dateObj - epoch) / 86400000);

    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    detectedMonth = monthNames[m - 1] || 'September';
    detectedYear = String(y || 2026);

    // Day of week abbreviation
    const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const dayStr = dayNames[dateObj.getUTCDay()] || 'Sun';

    // 2. Time fraction
    const timeParts = (c.time || '12:00').split(':');
    const hours = parseInt(timeParts[0], 10) || 12;
    const mins = parseInt(timeParts[1], 10) || 0;
    const timeFraction = (hours + mins / 60) / 24;

    // 3. Duration formatted string
    let durStr = '60mins';
    const durMins = Number(c.duration) || 60;
    if (durMins === 50) durStr = '50mins';
    else if (durMins === 60) durStr = '1hour';
    else if (durMins === 90) durStr = '1.5hour';
    else if (durMins === 120) durStr = '2hours';
    else if (durMins === 45) durStr = '45mins';
    else if (durMins === 30) durStr = '30mins';
    else if (durMins === 75) durStr = '75mins';
    else durStr = `${durMins}mins`;
    const durIdx = getStrIdx(durStr);

    // 4. Description & Remarks parsing
    let desc = 'LTS';
    let remarks = '';
    const note = (c.note || '').trim();
    if (note.includes(' - ')) {
      const parts = note.split(' - ');
      desc = parts[0].trim() || 'LTS';
      remarks = parts.slice(1).join(' - ').trim();
    } else if (note) {
      desc = note;
    } else if (c.category && c.category.toLowerCase() !== 'aerosplash') {
      desc = c.category;
    }
    const descIdx = getStrIdx(desc);

    // 5. Total Class Credit Units (LTS 50m = 1.0, Pre Comp 90m = 1.5)
    const units = getClassCreditUnits(c);
    const unitsStr = units.toFixed(1);

    let rowXml = `<row r="${r}" ht="14.25" customHeight="1">`;
    rowXml += `<c r="A${r}" s="7"><v>${serial}.0</v></c>`;
    rowXml += `<c r="B${r}" s="2" t="str"><f>TEXT(A${r}, &quot;ddd&quot;)</f><v>${dayStr}</v></c>`;
    rowXml += `<c r="C${r}" s="8"><v>${timeFraction}</v></c>`;
    rowXml += `<c r="D${r}" s="14" t="s"><v>${durIdx}</v></c>`;
    rowXml += `<c r="E${r}" s="15" t="s"><v>${descIdx}</v></c>`;
    rowXml += `<c r="F${r}" s="16"><v>${unitsStr}</v></c>`;

    if (remarks) {
      const remIdx = getStrIdx(remarks);
      rowXml += `<c r="G${r}" s="13" t="s"><v>${remIdx}</v></c>`;
    } else {
      rowXml += `<c r="G${r}" s="10"/>`;
    }
    rowXml += `</row>`;

    rowsXml.push(rowXml);
  });

  const count = sortedClasses.length;
  const totalUnits = sortedClasses.reduce((sum, c) => sum + getClassCreditUnits(c), 0);
  const totalUnitsStr = totalUnits.toFixed(1);

  // The template table has a standard 29-row grid (rows 5 to 33)
  // If count < 29, pad empty rows up to row 33 so Total Row is row 34, exactly matching reference sheet
  const minTableLastRow = 33;
  const lastDataRow = 4 + count;
  const tableLastRow = Math.max(minTableLastRow, lastDataRow);

  for (let r = lastDataRow + 1; r <= tableLastRow; r++) {
    if (r === 33) {
      rowsXml.push(`<row r="${r}" ht="14.25" customHeight="1"><c r="A${r}" s="19"/><c r="B${r}" s="19"/><c r="C${r}" s="19"/><c r="D${r}" s="20"/><c r="E${r}" s="10"/><c r="F${r}" s="11"/><c r="G${r}" s="10"/></row>`);
    } else {
      rowsXml.push(`<row r="${r}" ht="14.25" customHeight="1"><c r="A${r}" s="19"/><c r="B${r}" s="19"/><c r="C${r}" s="11"/><c r="D${r}" s="9"/><c r="E${r}" s="10"/><c r="F${r}" s="11"/><c r="G${r}" s="10"/></row>`);
    }
  }

  // Row totalRow (Row 34 when count <= 29): Total Class row with exact SUM formula and full cell borders
  const totalRow = tableLastRow + 1;
  const sumFormula = `SUM(F5:F${tableLastRow})`;

  rowsXml.push(
    `<row r="${totalRow}" ht="14.25" customHeight="1">` +
    `<c r="A${totalRow}" s="10"/>` +
    `<c r="B${totalRow}" s="10"/>` +
    `<c r="C${totalRow}" s="11"/>` +
    `<c r="D${totalRow}" s="9"/>` +
    `<c r="E${totalRow}" s="21"/>` +
    `<c r="F${totalRow}" s="5"><f>${sumFormula}</f><v>${totalUnitsStr}</v></c>` +
    `<c r="G${totalRow}" s="10"/>` +
    `</row>`
  );

  // Row totalRow + 1 (Row 35 when count <= 29): Separator row matching reference sheet
  const rSep = totalRow + 1;
  rowsXml.push(`<row r="${rSep}" ht="14.25" customHeight="1"><c r="C${rSep}" s="4"/><c r="D${rSep}" s="3"/><c r="F${rSep}" s="4"/></row>`);

  // Row totalRow + 2 (Row 36 when count <= 29): Currency symbol "RM"
  const rRm = totalRow + 2;
  const idxRm = getStrIdx('RM');
  rowsXml.push(`<row r="${rRm}" ht="14.25" customHeight="1"><c r="C${rRm}" s="4" t="s"><v>${idxRm}</v></c><c r="D${rRm}" s="3"/><c r="F${rRm}" s="4"/></row>`);

  // Row totalRow + 3 (Row 37 when count <= 29): "Rate per class :" and Rate value (e.g. 40.0)
  const rRate = totalRow + 3;
  const idxRate = getStrIdx('Rate per class :');
  rowsXml.push(`<row r="${rRate}" ht="14.25" customHeight="1"><c r="A${rRate}" s="1" t="s"><v>${idxRate}</v></c><c r="C${rRate}" s="22"><v>${ratePerClass.toFixed(1)}</v></c><c r="D${rRate}" s="3"/><c r="F${rRate}" s="4"/></row>`);

  // Row totalRow + 4 (Row 38 when count <= 29): "Total class :" and formula =F{totalRow} (=F34)
  const rTotal = totalRow + 4;
  const idxTotal = getStrIdx('Total class :');
  rowsXml.push(`<row r="${rTotal}" ht="14.25" customHeight="1"><c r="A${rTotal}" s="1" t="s"><v>${idxTotal}</v></c><c r="C${rTotal}" s="4"><f>F${totalRow}</f><v>${totalUnitsStr}</v></c><c r="D${rTotal}" s="3"/><c r="F${rTotal}" s="4"/></row>`);

  // Row totalRow + 5 (Row 39 when count <= 29): Base fees label and formula =C{rRate}*C{rTotal} (=C37*C38)
  const rFees = totalRow + 5;
  const feesLabel = options.feesLabel || (detectedMonth && detectedYear ? `${detectedMonth} ${detectedYear} base fees :` : 'Base fees :');
  const idxFees = getStrIdx(feesLabel);
  const baseFees = totalUnits * ratePerClass;
  rowsXml.push(`<row r="${rFees}" ht="14.25" customHeight="1"><c r="A${rFees}" s="1" t="s"><v>${idxFees}</v></c><c r="C${rFees}" s="4"><f>C${rRate}*C${rTotal}</f><v>${baseFees.toFixed(2)}</v></c><c r="D${rFees}" s="3"/><c r="F${rFees}" s="4"/></row>`);

  // Row totalRow + 6 (Row 40 when count <= 29): Manager bonus row with nested IF formula
  const rBonus = totalRow + 6;
  const idxBonus = getStrIdx('Manager bonus :');
  const bonusInfo = calculateManagerBonus(totalUnits);
  const bonusAmount = bonusInfo.bonus;
  const bonusFormula = `IF(C${rTotal}&gt;=50,350,IF(C${rTotal}&gt;=40,250,IF(C${rTotal}&gt;=30,150,IF(C${rTotal}&gt;=24,100,0))))`;
  rowsXml.push(`<row r="${rBonus}" ht="14.25" customHeight="1"><c r="A${rBonus}" s="1" t="s"><v>${idxBonus}</v></c><c r="C${rBonus}" s="4"><f>${bonusFormula}</f><v>${bonusAmount.toFixed(2)}</v></c><c r="D${rBonus}" s="3"/><c r="F${rBonus}" s="4"/></row>`);

  // Row totalRow + 7 (Row 41 when count <= 29): "Total fees to be paid :" with formula =C{rFees}+C{rBonus} (style 23 double underline)
  const rGrandTotal = totalRow + 7;
  const idxGrandTotal = getStrIdx('Total fees to be paid :');
  const grandTotal = baseFees + bonusAmount;
  rowsXml.push(`<row r="${rGrandTotal}" ht="14.25" customHeight="1"><c r="A${rGrandTotal}" s="1" t="s"><v>${idxGrandTotal}</v></c><c r="C${rGrandTotal}" s="23"><f>C${rFees}+C${rBonus}</f><v>${grandTotal.toFixed(2)}</v></c><c r="D${rGrandTotal}" s="3"/><c r="F${rGrandTotal}" s="4"/></row>`);

  // Rows trailing spacer rows matching template
  for (let r = rGrandTotal + 1; r <= rGrandTotal + 6; r++) {
    rowsXml.push(`<row r="${r}" ht="14.25" customHeight="1"><c r="C${r}" s="4"/><c r="D${r}" s="3"/><c r="F${r}" s="4"/></row>`);
  }

  // Build sheet1.xml
  const sheet1Xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
    `<sheetPr><pageSetUpPr/></sheetPr>` +
    `<sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.0"/>` +
    `<cols>` +
    `<col customWidth="1" min="1" max="2" width="8.63"/>` +
    `<col customWidth="1" min="3" max="4" width="8.75"/>` +
    `<col customWidth="1" min="5" max="5" width="13.25"/>` +
    `<col customWidth="1" min="6" max="6" width="9.88"/>` +
    `<col customWidth="1" min="7" max="7" width="16.88"/>` +
    `<col customWidth="1" min="8" max="26" width="8.63"/>` +
    `</cols>` +
    `<sheetData>${rowsXml.join('')}</sheetData>` +
    `<printOptions/><pageMargins bottom="0.75" footer="0.0" header="0.0" left="0.7" right="0.7" top="0.75"/><pageSetup orientation="portrait"/>` +
    `<drawing r:id="rId1"/>` +
    `</worksheet>`;

  // Build sharedStrings.xml
  const siElements = sharedStrings.map(s => {
    const escaped = s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    return `<si><t>${escaped}</t></si>`;
  }).join('');

  const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totalStringRefs}" uniqueCount="${sharedStrings.length}">${siElements}</sst>`;

  // Prepare full files map
  const outFiles = {};
  for (const [name, meta] of Object.entries(TEMPLATE_PARTS)) {
    if (meta.type === 'base64') {
      const binStr = atob(meta.data);
      const bytes = new Uint8Array(binStr.length);
      for (let i = 0; i < binStr.length; i++) bytes[i] = binStr.charCodeAt(i);
      outFiles[name] = bytes;
    } else {
      outFiles[name] = meta.data;
    }
  }

  outFiles['xl/worksheets/sheet1.xml'] = sheet1Xml;
  outFiles['xl/sharedStrings.xml'] = sharedStringsXml;

  return createZipBlob(outFiles);
}

import { unzipSync, zipSync } from 'fflate';
import * as fs from 'fs';

// src/zip/index.ts
var encoder = new TextEncoder();
var decoder = new TextDecoder();
function zip_read(data) {
  const unzipped = unzipSync(data);
  return { files: unzipped };
}
function zip_write(archive, compress) {
  const zippable = {};
  for (const [name, data] of Object.entries(archive.files)) {
    zippable[name] = compress ? data : [data, { level: 0 }];
  }
  return zipSync(zippable);
}
function zip_read_str(archive, path) {
  let data = archive.files[path];
  if (!data) {
    const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
    data = archive.files[normalized];
  }
  if (!data) return null;
  return decoder.decode(data);
}
function zip_add_str(archive, path, content) {
  archive.files[path] = encoder.encode(content);
}
function zip_new() {
  return { files: {} };
}
function zip_has(archive, path) {
  if (archive.files[path]) return true;
  const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
  if (archive.files[normalized]) return true;
  const lpath = path.toLowerCase();
  for (const k of Object.keys(archive.files)) {
    if (k.toLowerCase() === lpath) return true;
  }
  return false;
}

// src/xml/parser.ts
var attregexg = /\s([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
var tagregex1 = /<[\/\?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?<>\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'"<>\s=]+))*\s*[\/\?]?>/gm;
var tagregex2 = /<[^<>]*>/g;
var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
var tagregex = XML_HEADER.match(tagregex1) ? tagregex1 : tagregex2;
var nsregex2 = /<(\/?)\w+:/;
function parsexmltag(tag, skip_root, skip_LC) {
  const z = {};
  let eq = 0;
  let c = 0;
  for (; eq !== tag.length; ++eq) if ((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
  z[0] = tag.slice(0, eq);
  if (eq === tag.length) return z;
  const m = tag.match(attregexg);
  if (m) {
    for (let i = 0; i < m.length; ++i) {
      const cc = m[i].slice(1);
      let c2 = 0;
      for (c2 = 0; c2 < cc.length; ++c2) if (cc.charCodeAt(c2) === 61) break;
      let q = cc.slice(0, c2).trim();
      while (cc.charCodeAt(c2 + 1) === 32) ++c2;
      const quot = (eq = cc.charCodeAt(c2 + 1)) === 34 || eq === 39 ? 1 : 0;
      const v = cc.slice(c2 + 1 + quot, cc.length - quot);
      let j = 0;
      for (j = 0; j < q.length; ++j) if (q.charCodeAt(j) === 58) break;
      if (j === q.length) {
        if (q.indexOf("_") > 0) q = q.slice(0, q.indexOf("_"));
        z[q] = v;
        z[q.toLowerCase()] = v;
      } else {
        const k = (j === 5 && q.slice(0, 5) === "xmlns" ? "xmlns" : "") + q.slice(j + 1);
        if (z[k] && q.slice(j - 3, j) === "ext") continue;
        z[k] = v;
        z[k.toLowerCase()] = v;
      }
    }
  }
  return z;
}
function strip_ns(x) {
  return x.replace(nsregex2, "<$1");
}
function parsexmlbool(value) {
  switch (value) {
    case 1:
    case true:
    case "1":
    case "true":
      return true;
    case 0:
    case false:
    case "0":
    case "false":
      return false;
  }
  return false;
}

// src/xml/escape.ts
var encodings = {
  "&quot;": '"',
  "&apos;": "'",
  "&gt;": ">",
  "&lt;": "<",
  "&amp;": "&"
};
var rencoding = {
  '"': "&quot;",
  "'": "&apos;",
  ">": "&gt;",
  "<": "&lt;",
  "&": "&amp;"
};
var encregex = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/gi;
var coderegex = /_x([\da-fA-F]{4})_/g;
var decregex = /[&<>'"]/g;
var charegex = /[\u0000-\u0008\u000b-\u001f\uFFFE-\uFFFF]/g;
function raw_unescapexml(text) {
  const s = text + "";
  const i = s.indexOf("<![CDATA[");
  if (i === -1) {
    return s.replace(encregex, ($$, $1) => {
      return encodings[$$] || String.fromCharCode(parseInt($1, $$.indexOf("x") > -1 ? 16 : 10)) || $$;
    }).replace(coderegex, (_m, c) => {
      return String.fromCharCode(parseInt(c, 16));
    });
  }
  const j = s.indexOf("]]>");
  return raw_unescapexml(s.slice(0, i)) + s.slice(i + 9, j) + raw_unescapexml(s.slice(j + 3));
}
function unescapexml(text, xlsx) {
  const out = raw_unescapexml(text);
  return xlsx ? out.replace(/\r\n/g, "\n") : out;
}
function escapexml(text) {
  const s = text + "";
  return s.replace(decregex, (y) => rencoding[y]).replace(charegex, (s2) => "_x" + ("000" + s2.charCodeAt(0).toString(16)).slice(-4) + "_");
}
var htmlcharegex = /[\u0000-\u001f]/g;
function escapehtml(text) {
  const s = text + "";
  return s.replace(decregex, (y) => rencoding[y]).replace(/\n/g, "<br/>").replace(htmlcharegex, (s2) => "&#x" + ("000" + s2.charCodeAt(0).toString(16)).slice(-4) + ";");
}
[
  ["nbsp", " "],
  ["middot", "\xB7"],
  ["quot", '"'],
  ["apos", "'"],
  ["gt", ">"],
  ["lt", "<"],
  ["amp", "&"]
].map(([name, ch]) => [new RegExp("&" + name + ";", "gi"), ch]);

// src/utils/helpers.ts
function fill(c, l) {
  let o = "";
  while (o.length < l) o += c;
  return o;
}
function keys(o) {
  return Object.keys(o);
}
function dup(o) {
  if (typeof o === "object" && o !== null) {
    if (Array.isArray(o)) return o.slice();
    const out = {};
    for (const k of Object.keys(o)) out[k] = o[k];
    return out;
  }
  return o;
}
function str_match_xml_ns_g(str, tag) {
  const re = new RegExp("<(?:\\w+:)?" + tag + "[\\s>][\\s\\S]*?<\\/(?:\\w+:)?" + tag + ">", "g");
  return str.match(re);
}
function str_match_xml_ns(str, tag) {
  const m = str_match_xml_ns_g(str, tag);
  return m ? m[0] : null;
}

// src/xml/writer.ts
var wtregex = /(^\s|\s$|\n)/;
function writetag(f, g) {
  return "<" + f + (g.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + g + "</" + f + ">";
}
function wxt_helper(h) {
  return keys(h).map((k) => " " + k + '="' + h[k] + '"').join("");
}
function writextag(f, g, h) {
  return "<" + f + (h != null ? wxt_helper(h) : "") + (g != null ? (g.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + g + "</" + f : "/") + ">";
}
function write_w3cdtf(d, t) {
  try {
    return d.toISOString().replace(/\.\d*/, "");
  } catch (e) {
    if (t) throw e;
  }
  return "";
}
function write_vt(s, xlsx) {
  switch (typeof s) {
    case "string": {
      let o = writextag("vt:lpwstr", escapexml(s));
      o = o.replace(/&quot;/g, "_x0022_");
      return o;
    }
    case "number":
      return writextag((s | 0) === s ? "vt:i4" : "vt:r8", escapexml(String(s)));
    case "boolean":
      return writextag("vt:bool", s ? "true" : "false");
  }
  if (s instanceof Date) return writextag("vt:filetime", write_w3cdtf(s));
  throw new Error("Unable to serialize " + s);
}

// src/xml/namespaces.ts
var XMLNS = {
  CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  CUST_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
  EXT_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  CT: "http://schemas.openxmlformats.org/package/2006/content-types",
  RELS: "http://schemas.openxmlformats.org/package/2006/relationships",
  TCMNT: "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
  dc: "http://purl.org/dc/elements/1.1/",
  dcterms: "http://purl.org/dc/terms/",
  dcmitype: "http://purl.org/dc/dcmitype/",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
  xsi: "http://www.w3.org/2001/XMLSchema-instance",
  xsd: "http://www.w3.org/2001/XMLSchema"
};
var XMLNS_main = [
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://purl.oclc.org/ooxml/spreadsheetml/main",
  "http://schemas.microsoft.com/office/excel/2006/main",
  "http://schemas.microsoft.com/office/excel/2006/2"
];
var RELS = {
  WB: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  SHEET: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  CHARTSHEET: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
  HLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  VML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
  CMNT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  EXT_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  CUST_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
  SST: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
  STY: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  THEME: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  CHART: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  CCHAIN: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
  TCMNT: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
  PEOPLE: "http://schemas.microsoft.com/office/2017/10/relationships/person",
  DRAWING: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
  META: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
  XLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
};

// src/opc/content-types.ts
var nsregex = /<(\w+):/;
var ct2type = {
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
  "application/vnd.ms-excel.sheet.macroEnabled.main+xml": "workbooks",
  "application/vnd.ms-excel.sheet.binary.macroEnabled.main": "workbooks",
  "application/vnd.ms-excel.addin.macroEnabled.main+xml": "workbooks",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "workbooks",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml": "sheets",
  "application/vnd.ms-excel.worksheet": "sheets",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml": "charts",
  "application/vnd.ms-excel.chartsheet": "charts",
  "application/vnd.ms-excel.macrosheet+xml": "macros",
  "application/vnd.ms-excel.macrosheet": "macros",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml": "dialogs",
  "application/vnd.ms-excel.dialogsheet": "dialogs",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",
  "application/vnd.ms-excel.sharedStrings": "strs",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": "styles",
  "application/vnd.ms-excel.styles": "styles",
  "application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
  "application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
  "application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments",
  "application/vnd.ms-excel.comments": "comments",
  "application/vnd.ms-excel.threadedcomments+xml": "threadedcomments",
  "application/vnd.ms-excel.person+xml": "people",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml": "metadata",
  "application/vnd.ms-excel.sheetMetadata": "metadata",
  "application/vnd.ms-excel.calcChain": "calcchains",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",
  "application/vnd.openxmlformats-officedocument.theme+xml": "themes",
  "application/vnd.ms-office.vbaProject": "vba",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml": "links",
  "application/vnd.ms-excel.externalLink": "links",
  "application/vnd.openxmlformats-officedocument.drawing+xml": "drawings",
  "application/vnd.openxmlformats-package.relationships+xml": "rels"
};
var CT_LIST = {
  workbooks: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
  },
  strs: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
  },
  comments: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
  },
  sheets: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  },
  charts: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
  },
  dialogs: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"
  },
  macros: {
    xlsx: "application/vnd.ms-excel.macrosheet+xml"
  },
  metadata: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
  },
  styles: {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
  }
};
function new_ct() {
  return {
    workbooks: [],
    sheets: [],
    charts: [],
    dialogs: [],
    macros: [],
    rels: [],
    strs: [],
    comments: [],
    threadedcomments: [],
    links: [],
    coreprops: [],
    extprops: [],
    custprops: [],
    themes: [],
    styles: [],
    calcchains: [],
    vba: [],
    drawings: [],
    metadata: [],
    people: [],
    xmlns: ""
  };
}
function parse_ct(data) {
  const ct = new_ct();
  if (!data) return ct;
  const ctext = {};
  const matches = data.match(tagregex) || [];
  for (const x of matches) {
    const y = parsexmltag(x);
    switch (y[0].replace(nsregex, "<")) {
      case "<?xml":
        break;
      case "<Types":
        ct.xmlns = y["xmlns" + (y[0].match(/<(\w+):/) || ["", ""])[1]];
        break;
      case "<Default":
        ctext[y.Extension.toLowerCase()] = y.ContentType;
        break;
      case "<Override":
        if (ct2type[y.ContentType] && ct[ct2type[y.ContentType]] !== void 0) {
          ct[ct2type[y.ContentType]].push(y.PartName);
        }
        break;
    }
  }
  if (ct.xmlns !== XMLNS.CT) throw new Error("Unknown Namespace: " + ct.xmlns);
  ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
  ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
  ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
  ct.defaults = ctext;
  return ct;
}
function evert_arr(obj) {
  const o = {};
  for (const [k, v] of Object.entries(obj)) {
    if (!o[v]) o[v] = [];
    o[v].push(k);
  }
  return o;
}
function write_ct(ct, opts) {
  const type2ct = evert_arr(ct2type);
  const o = [];
  o.push(XML_HEADER);
  o.push(
    writextag("Types", null, {
      xmlns: XMLNS.CT,
      "xmlns:xsd": XMLNS.xsd,
      "xmlns:xsi": XMLNS.xsi
    })
  );
  const defaults = [
    ["xml", "application/xml"],
    ["bin", "application/vnd.ms-excel.sheet.binary.macroEnabled.main"],
    ["vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"],
    ["data", "application/vnd.openxmlformats-officedocument.model+data"],
    ["bmp", "image/bmp"],
    ["png", "image/png"],
    ["gif", "image/gif"],
    ["emf", "image/x-emf"],
    ["wmf", "image/x-wmf"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["tif", "image/tiff"],
    ["tiff", "image/tiff"],
    ["pdf", "application/pdf"],
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"]
  ];
  for (const [ext, contentType] of defaults) {
    o.push(writextag("Default", null, { Extension: ext, ContentType: contentType }));
  }
  const f1 = (w) => {
    if (ct[w] && ct[w].length > 0) {
      const v = ct[w][0];
      o.push(
        writextag("Override", null, {
          PartName: (v[0] === "/" ? "" : "/") + v,
          ContentType: CT_LIST[w]?.[opts.bookType || "xlsx"] || CT_LIST[w]?.["xlsx"]
        })
      );
    }
  };
  const f2 = (w) => {
    for (const v of ct[w] || []) {
      o.push(
        writextag("Override", null, {
          PartName: (v[0] === "/" ? "" : "/") + v,
          ContentType: CT_LIST[w]?.[opts.bookType || "xlsx"] || CT_LIST[w]?.["xlsx"]
        })
      );
    }
  };
  const f3 = (t) => {
    for (const v of ct[t] || []) {
      o.push(
        writextag("Override", null, {
          PartName: (v[0] === "/" ? "" : "/") + v,
          ContentType: type2ct[t]?.[0]
        })
      );
    }
  };
  f1("workbooks");
  f2("sheets");
  f2("charts");
  f3("themes");
  f1("strs");
  f1("styles");
  f3("coreprops");
  f3("extprops");
  f3("custprops");
  f3("vba");
  f3("comments");
  f3("threadedcomments");
  f3("drawings");
  f2("metadata");
  f3("people");
  if (o.length > 2) {
    o.push("</Types>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}

// src/opc/relationships.ts
function resolve_path(target, basePath) {
  if (target.charAt(0) === "/") return target;
  const base = basePath.slice(0, basePath.lastIndexOf("/") + 1);
  const parts = (base + target).split("/");
  const resolved = [];
  for (const p of parts) {
    if (p === "..") resolved.pop();
    else if (p !== ".") resolved.push(p);
  }
  return resolved.join("/");
}
function get_rels_path(file) {
  const n = file.lastIndexOf("/");
  return file.slice(0, n + 1) + "_rels/" + file.slice(n + 1) + ".rels";
}
function parse_rels(data, currentFilePath) {
  const rels = { "!id": {} };
  if (!data) return rels;
  if (currentFilePath.charAt(0) !== "/") {
    currentFilePath = "/" + currentFilePath;
  }
  const matches = data.match(tagregex) || [];
  for (const x of matches) {
    const y = parsexmltag(x);
    if (y[0] === "<Relationship") {
      const rel = {
        Type: y.Type,
        Target: unescapexml(y.Target),
        Id: y.Id
      };
      if (y.TargetMode) rel.TargetMode = y.TargetMode;
      const canonictarget = y.TargetMode === "External" ? y.Target : resolve_path(y.Target, currentFilePath);
      rels[canonictarget] = rel;
      rels["!id"][y.Id] = rel;
    }
  }
  return rels;
}
function write_rels(rels) {
  const o = [
    XML_HEADER,
    writextag("Relationships", null, {
      xmlns: XMLNS.RELS
    })
  ];
  for (const rid of Object.keys(rels["!id"])) {
    o.push(writextag("Relationship", null, rels["!id"][rid]));
  }
  if (o.length > 2) {
    o.push("</Relationships>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}
function add_rels(rels, rId, f, type, targetmode) {
  if (!rels["!id"]) rels["!id"] = {};
  if (!rels["!idx"]) rels["!idx"] = 1;
  if (rId < 0) {
    for (rId = rels["!idx"]; rels["!id"]["rId" + rId]; ++rId) {
    }
  }
  rels["!idx"] = rId + 1;
  const relobj = {
    Id: "rId" + rId,
    Type: type,
    Target: f
  };
  if ([RELS.HLINK].indexOf(type) > -1) relobj.TargetMode = "External";
  if (rels["!id"][relobj.Id]) throw new Error("Cannot rewrite rId " + rId);
  rels["!id"][relobj.Id] = relobj;
  rels[("/" + relobj.Target).replace("//", "/")] = relobj;
  return rId;
}

// src/opc/core-properties.ts
var CORE_PROPS = [
  ["cp:category", "Category"],
  ["cp:contentStatus", "ContentStatus"],
  ["cp:keywords", "Keywords"],
  ["cp:lastModifiedBy", "LastAuthor"],
  ["cp:lastPrinted", "LastPrinted"],
  ["cp:revision", "Revision"],
  ["cp:version", "Version"],
  ["dc:creator", "Author"],
  ["dc:description", "Comments"],
  ["dc:identifier", "Identifier"],
  ["dc:language", "Language"],
  ["dc:subject", "Subject"],
  ["dc:title", "Title"],
  ["dcterms:created", "CreatedDate", "date"],
  ["dcterms:modified", "ModifiedDate", "date"]
];
function xml_extract(data, tag) {
  const open = "<" + tag;
  const close = "</" + tag + ">";
  let si = data.indexOf(open);
  if (si === -1) {
    return null;
  }
  const gt = data.indexOf(">", si);
  if (gt === -1) return null;
  const ei = data.indexOf(close, gt);
  if (ei === -1) return null;
  return data.slice(gt + 1, ei);
}
function parse_core_props(data) {
  const p = {};
  for (const f of CORE_PROPS) {
    const content = xml_extract(data, f[0]);
    if (content != null && content.length > 0) {
      p[f[1]] = unescapexml(content);
    }
    if (f[2] === "date" && p[f[1]]) {
      p[f[1]] = new Date(p[f[1]]);
    }
  }
  return p;
}
function cp_doit(f, g, h, o, p) {
  if (p[f] != null || g == null || g === "") return;
  p[f] = g;
  g = escapexml(g);
  o.push(h ? writextag(f, g, h) : writetag(f, g));
}
function write_core_props(cp, opts) {
  const o = [
    XML_HEADER,
    writextag("cp:coreProperties", null, {
      "xmlns:cp": XMLNS.CORE_PROPS,
      "xmlns:dc": XMLNS.dc,
      "xmlns:dcterms": XMLNS.dcterms,
      "xmlns:dcmitype": XMLNS.dcmitype,
      "xmlns:xsi": XMLNS.xsi
    })
  ];
  const p = {};
  if (!cp && !opts?.Props) return o.join("");
  if (cp) {
    if (cp.CreatedDate != null) {
      cp_doit(
        "dcterms:created",
        typeof cp.CreatedDate === "string" ? cp.CreatedDate : write_w3cdtf(cp.CreatedDate, opts?.WTF),
        { "xsi:type": "dcterms:W3CDTF" },
        o,
        p
      );
    }
    if (cp.ModifiedDate != null) {
      cp_doit(
        "dcterms:modified",
        typeof cp.ModifiedDate === "string" ? cp.ModifiedDate : write_w3cdtf(cp.ModifiedDate, opts?.WTF),
        { "xsi:type": "dcterms:W3CDTF" },
        o,
        p
      );
    }
  }
  for (const f of CORE_PROPS) {
    let v = opts?.Props?.[f[1]] != null ? opts.Props[f[1]] : cp ? cp[f[1]] : null;
    if (v === true) v = "1";
    else if (v === false) v = "0";
    else if (typeof v === "number") v = String(v);
    if (v != null) cp_doit(f[0], v, null, o, p);
  }
  if (o.length > 2) {
    o.push("</cp:coreProperties>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}

// src/opc/extended-properties.ts
var EXT_PROPS = [
  ["Application", "Application", "string"],
  ["AppVersion", "AppVersion", "string"],
  ["Company", "Company", "string"],
  ["DocSecurity", "DocSecurity", "string"],
  ["Manager", "Manager", "string"],
  ["HyperlinksChanged", "HyperlinksChanged", "bool"],
  ["SharedDoc", "SharedDoc", "bool"],
  ["LinksUpToDate", "LinksUpToDate", "bool"],
  ["ScaleCrop", "ScaleCrop", "bool"]
];
function xml_extract_ns(data, tag) {
  const re = new RegExp("<(?:\\w+:)?" + tag + "[\\s>]([\\s\\S]*?)<\\/(?:\\w+:)?" + tag + ">");
  const m = data.match(re);
  return m ? m[1] : null;
}
function parse_ext_props(data, p) {
  if (!p) p = {};
  for (const f of EXT_PROPS) {
    const xml = xml_extract_ns(data, f[0]);
    switch (f[2]) {
      case "string":
        if (xml) p[f[1]] = unescapexml(xml);
        break;
      case "bool":
        p[f[1]] = xml === "true";
        break;
    }
  }
  const hpMatch = data.match(/<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/);
  const topMatch = data.match(/<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/);
  if (hpMatch && topMatch) {
    const lpstrs = topMatch[1].match(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g);
    if (lpstrs) {
      const parts = lpstrs.map((s) => {
        const m = s.match(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/);
        return m ? unescapexml(m[1]) : "";
      });
      const i4match = hpMatch[1].match(/<vt:i4>(\d+)<\/vt:i4>/);
      if (i4match) {
        p.Worksheets = parseInt(i4match[1], 10);
        p.SheetNames = parts.slice(0, p.Worksheets);
      }
    }
  }
  return p;
}
function write_ext_props(cp) {
  const o = [];
  const W = writextag;
  if (!cp) cp = {};
  cp.Application = "xlsx-format";
  o.push(XML_HEADER);
  o.push(
    writextag("Properties", null, {
      xmlns: XMLNS.EXT_PROPS,
      "xmlns:vt": XMLNS.vt
    })
  );
  for (const f of EXT_PROPS) {
    if (cp[f[1]] === void 0) continue;
    let v;
    switch (f[2]) {
      case "string":
        v = escapexml(String(cp[f[1]]));
        break;
      case "bool":
        v = cp[f[1]] ? "true" : "false";
        break;
    }
    if (v !== void 0) o.push(W(f[0], v));
  }
  o.push(
    W(
      "HeadingPairs",
      W(
        "vt:vector",
        W("vt:variant", "<vt:lpstr>Worksheets</vt:lpstr>") + W("vt:variant", W("vt:i4", String(cp.Worksheets))),
        { size: "2", baseType: "variant" }
      )
    )
  );
  o.push(
    W(
      "TitlesOfParts",
      W(
        "vt:vector",
        cp.SheetNames.map((s) => "<vt:lpstr>" + escapexml(s) + "</vt:lpstr>").join(""),
        { size: String(cp.Worksheets), baseType: "lpstr" }
      )
    )
  );
  if (o.length > 2) {
    o.push("</Properties>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}

// src/opc/custom-properties.ts
var custregex = /<[^<>]+>[^<]*/g;
function parse_cust_props(data, opts) {
  const p = {};
  let name = "";
  const m = data.match(custregex);
  if (m) {
    for (let i = 0; i < m.length; ++i) {
      const x = m[i];
      const y = parsexmltag(x);
      switch (strip_ns(y[0])) {
        case "<?xml":
          break;
        case "<Properties":
          break;
        case "<property":
          name = unescapexml(y.name);
          break;
        case "</property>":
          name = "";
          break;
        default:
          if (x.indexOf("<vt:") === 0) {
            const toks = x.split(">");
            const type = toks[0].slice(4);
            const text = toks[1];
            switch (type) {
              case "lpstr":
              case "bstr":
              case "lpwstr":
                p[name] = unescapexml(text);
                break;
              case "bool":
                p[name] = parsexmlbool(text);
                break;
              case "i1":
              case "i2":
              case "i4":
              case "i8":
              case "int":
              case "uint":
                p[name] = parseInt(text, 10);
                break;
              case "r4":
              case "r8":
              case "decimal":
                p[name] = parseFloat(text);
                break;
              case "filetime":
              case "date":
                p[name] = new Date(text);
                break;
              case "cy":
              case "error":
                p[name] = unescapexml(text);
                break;
              default:
                if (type.slice(-1) === "/") break;
                if (opts?.WTF && typeof console !== "undefined")
                  console.warn("Unexpected", x, type, toks);
            }
          }
      }
    }
  }
  return p;
}
function write_cust_props(cp) {
  const o = [
    XML_HEADER,
    writextag("Properties", null, {
      xmlns: XMLNS.CUST_PROPS,
      "xmlns:vt": XMLNS.vt
    })
  ];
  if (!cp) return o.join("");
  let pid = 1;
  for (const k of Object.keys(cp)) {
    ++pid;
    o.push(
      writextag("property", write_vt(cp[k]), {
        fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
        pid: String(pid),
        name: escapexml(k)
      })
    );
  }
  if (o.length > 2) {
    o.push("</Properties>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}

// src/utils/buffer.ts
new TextEncoder();
new TextDecoder();
function utf8read(orig) {
  let out = "";
  let i = 0;
  let c = 0;
  let d = 0;
  let e = 0;
  let f = 0;
  let w = 0;
  while (i < orig.length) {
    c = orig.charCodeAt(i++);
    if (c < 128) {
      out += String.fromCharCode(c);
      continue;
    }
    d = orig.charCodeAt(i++);
    if (c > 191 && c < 224) {
      f = (c & 31) << 6 | d & 63;
      out += String.fromCharCode(f);
      continue;
    }
    e = orig.charCodeAt(i++);
    if (c < 240) {
      out += String.fromCharCode((c & 15) << 12 | (d & 63) << 6 | e & 63);
      continue;
    }
    f = orig.charCodeAt(i++);
    w = ((c & 7) << 18 | (d & 63) << 12 | (e & 63) << 6 | f & 63) - 65536;
    out += String.fromCharCode(55296 + (w >>> 10 & 1023));
    out += String.fromCharCode(56320 + (w & 1023));
  }
  return out;
}

// src/xlsx/shared-strings.ts
function parse_rpr(rpr) {
  const font = {};
  const m = rpr.match(tagregex);
  let pass = false;
  if (m) {
    for (let i = 0; i < m.length; ++i) {
      const y = parsexmltag(m[i]);
      switch (y[0].replace(/<\w*:/g, "<")) {
        case "<condense":
        case "<extend":
          break;
        case "<shadow":
          if (!y.val) break;
        case "<shadow>":
        case "<shadow/>":
          font.shadow = 1;
          break;
        case "</shadow>":
          break;
        case "<rFont":
          font.name = y.val;
          break;
        case "<sz":
          font.sz = y.val;
          break;
        case "<strike":
          if (!y.val) break;
        case "<strike>":
        case "<strike/>":
          font.strike = 1;
          break;
        case "</strike>":
          break;
        case "<u":
          if (!y.val) break;
          switch (y.val) {
            case "double":
              font.uval = "double";
              break;
            case "singleAccounting":
              font.uval = "single-accounting";
              break;
            case "doubleAccounting":
              font.uval = "double-accounting";
              break;
          }
        case "<u>":
        case "<u/>":
          font.u = 1;
          break;
        case "</u>":
          break;
        case "<b":
          if (y.val === "0") break;
        case "<b>":
        case "<b/>":
          font.b = 1;
          break;
        case "</b>":
          break;
        case "<i":
          if (y.val === "0") break;
        case "<i>":
        case "<i/>":
          font.i = 1;
          break;
        case "</i>":
          break;
        case "<color":
          if (y.rgb) font.color = y.rgb.slice(2, 8);
          break;
        case "<color>":
        case "<color/>":
        case "</color>":
          break;
        case "<family":
          font.family = y.val;
          break;
        case "<vertAlign":
          font.valign = y.val;
          break;
        case "<scheme":
          break;
        case "<extLst":
        case "<extLst>":
        case "</extLst>":
          break;
        case "<ext":
          pass = true;
          break;
        case "</ext>":
          pass = false;
          break;
        default:
          if (y[0].charCodeAt(1) !== 47 && !pass)
            throw new Error("Unrecognized rich format " + y[0]);
      }
    }
  }
  return font;
}
var rregex = /<(?:\w+:)?r>/g;
var rend = /<\/(?:\w+:)?r>/;
function str_match_xml_ns_local(str, tag) {
  const re = new RegExp("<(?:\\w+:)?" + tag + "\\b[^<>]*>", "g");
  const reEnd = new RegExp("<\\/(?:\\w+:)?" + tag + ">", "g");
  const m = re.exec(str);
  if (!m) return null;
  const si = m.index;
  const sf = re.lastIndex;
  reEnd.lastIndex = re.lastIndex;
  const m2 = reEnd.exec(str);
  if (!m2) return null;
  const ei = m2.index;
  const ef = reEnd.lastIndex;
  return [str.slice(si, ef), str.slice(sf, ei)];
}
function str_remove_xml_ns_g_local(str, tag) {
  const re = new RegExp("<(?:\\w+:)?" + tag + "\\b[^<>]*>", "g");
  const reEnd = new RegExp("<\\/(?:\\w+:)?" + tag + ">", "g");
  const out = [];
  let lastEnd = 0;
  let m;
  while (m = re.exec(str)) {
    out.push(str.slice(lastEnd, m.index));
    reEnd.lastIndex = re.lastIndex;
    const m2 = reEnd.exec(str);
    if (!m2) break;
    lastEnd = reEnd.lastIndex;
    re.lastIndex = reEnd.lastIndex;
  }
  out.push(str.slice(lastEnd));
  return out.join("");
}
function parse_r(r) {
  const t = str_match_xml_ns_local(r, "t");
  if (!t) return { t: "s", v: "" };
  const o = { t: "s", v: unescapexml(t[1]) };
  const rpr = str_match_xml_ns_local(r, "rPr");
  if (rpr) o.s = parse_rpr(rpr[1]);
  return o;
}
function parse_rs(rs) {
  return rs.replace(rregex, "").split(rend).map(parse_r).filter((r) => r.v);
}
function rs_to_html(rs) {
  const nlregex = /(\r\n|\n)/g;
  return rs.map((r) => {
    if (!r.v) return "";
    const intro = [];
    const outro = [];
    if (r.s) {
      const font = r.s;
      const style = [];
      if (font.u) style.push("text-decoration: underline;");
      if (font.uval) style.push("text-underline-style:" + font.uval + ";");
      if (font.sz) style.push("font-size:" + font.sz + "pt;");
      if (font.outline) style.push("text-effect: outline;");
      if (font.shadow) style.push("text-shadow: auto;");
      intro.push('<span style="' + style.join("") + '">');
      if (font.b) {
        intro.push("<b>");
        outro.push("</b>");
      }
      if (font.i) {
        intro.push("<i>");
        outro.push("</i>");
      }
      if (font.strike) {
        intro.push("<s>");
        outro.push("</s>");
      }
      let align = font.valign || "";
      if (align === "superscript" || align === "super") align = "sup";
      else if (align === "subscript") align = "sub";
      if (align !== "") {
        intro.push("<" + align + ">");
        outro.push("</" + align + ">");
      }
      outro.push("</span>");
    }
    return intro.join("") + r.v.replace(nlregex, "<br/>") + outro.join("");
  }).join("");
}
var sitregex = /<(?:\w+:)?t\b[^<>]*>([^<]*)<\/(?:\w+:)?t>/g;
var sirregex = /<(?:\w+:)?r\b[^<>]*>/;
function parse_si(x, opts) {
  const html = opts ? opts.cellHTML !== false : true;
  const z = {};
  if (!x) return { t: "" };
  if (x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
    z.t = unescapexml(utf8read(x.slice(x.indexOf(">") + 1).split(/<\/(?:\w+:)?t>/)[0] || ""), true);
    z.r = utf8read(x);
    if (html) z.h = escapehtml(z.t);
  } else if (x.match(sirregex)) {
    z.r = utf8read(x);
    const stripped = str_remove_xml_ns_g_local(x, "rPh");
    sitregex.lastIndex = 0;
    const matches = stripped.match(sitregex) || [];
    z.t = unescapexml(utf8read(matches.join("").replace(tagregex, "")), true);
    if (html) z.h = rs_to_html(parse_rs(z.r));
  }
  return z;
}
var sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g;
var sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/;
function parse_sst_xml(data, opts) {
  const s = [];
  if (!data) return s;
  const sst = str_match_xml_ns_local(data, "sst");
  if (sst) {
    const ss = sst[1].replace(sstr1, "").split(sstr2);
    for (let i = 0; i < ss.length; ++i) {
      const o = parse_si(ss[i].trim(), opts);
      if (o != null) s[s.length] = o;
    }
    const tag = parsexmltag(sst[0].slice(0, sst[0].indexOf(">")));
    s.Count = tag.count;
    s.Unique = tag.uniquecount;
  }
  return s;
}
var straywsregex = /^\s|\s$|[\t\n\r]/;
function write_sst_xml(sst, opts) {
  if (!opts.bookSST) return "";
  const o = [XML_HEADER];
  o.push(
    writextag("sst", null, {
      xmlns: XMLNS_main[0],
      count: String(sst.Count),
      uniqueCount: String(sst.Unique)
    })
  );
  for (let i = 0; i !== sst.length; ++i) {
    if (sst[i] == null) continue;
    const s = sst[i];
    let sitag = "<si>";
    if (s.r) sitag += s.r;
    else {
      sitag += "<t";
      if (!s.t) s.t = "";
      if (typeof s.t !== "string") s.t = String(s.t);
      if (s.t.match(straywsregex)) sitag += ' xml:space="preserve"';
      sitag += ">" + escapexml(s.t) + "</t>";
    }
    sitag += "</si>";
    o.push(sitag);
  }
  if (o.length > 2) {
    o.push("</sst>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}

// src/ssf/table.ts
function SSF_init_table(t) {
  if (!t) t = {};
  t[0] = "General";
  t[1] = "0";
  t[2] = "0.00";
  t[3] = "#,##0";
  t[4] = "#,##0.00";
  t[9] = "0%";
  t[10] = "0.00%";
  t[11] = "0.00E+00";
  t[12] = "# ?/?";
  t[13] = "# ??/??";
  t[14] = "m/d/yy";
  t[15] = "d-mmm-yy";
  t[16] = "d-mmm";
  t[17] = "mmm-yy";
  t[18] = "h:mm AM/PM";
  t[19] = "h:mm:ss AM/PM";
  t[20] = "h:mm";
  t[21] = "h:mm:ss";
  t[22] = "m/d/yy h:mm";
  t[37] = "#,##0 ;(#,##0)";
  t[38] = "#,##0 ;[Red](#,##0)";
  t[39] = "#,##0.00;(#,##0.00)";
  t[40] = "#,##0.00;[Red](#,##0.00)";
  t[45] = "mm:ss";
  t[46] = "[h]:mm:ss";
  t[47] = "mmss.0";
  t[48] = "##0.0E+0";
  t[49] = "@";
  t[56] = '"\u4E0A\u5348/\u4E0B\u5348 "hh"\u6642"mm"\u5206"ss"\u79D2 "';
  return t;
}
var table_fmt = SSF_init_table();
var SSF_default_map = {
  5: 37,
  6: 38,
  7: 39,
  8: 40,
  23: 0,
  24: 0,
  25: 0,
  26: 0,
  27: 14,
  28: 14,
  29: 14,
  30: 14,
  31: 14,
  50: 14,
  51: 14,
  52: 14,
  53: 14,
  54: 14,
  55: 14,
  56: 14,
  57: 14,
  58: 14,
  59: 1,
  60: 2,
  61: 3,
  62: 4,
  67: 9,
  68: 10,
  69: 12,
  70: 13,
  71: 14,
  72: 14,
  73: 15,
  74: 16,
  75: 17,
  76: 20,
  77: 21,
  78: 22,
  79: 45,
  80: 46,
  81: 47,
  82: 0
};
var SSF_default_str = {
  5: '"$"#,##0_);\\("$"#,##0\\)',
  63: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  64: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  65: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  66: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
};
function SSF_load(fmt, idx) {
  if (typeof idx !== "number") {
    idx = +idx || -1;
    for (let i = 0; i < 392; ++i) {
      if (table_fmt[i] === void 0) {
        if (idx < 0) idx = i;
        continue;
      }
      if (table_fmt[i] === fmt) {
        idx = i;
        break;
      }
    }
    if (idx < 0) idx = 391;
  }
  table_fmt[idx] = fmt;
  return idx;
}
function SSF_load_table(tbl) {
  for (let i = 0; i < 392; ++i) if (tbl[i] !== void 0) SSF_load(tbl[i], i);
}
function make_ssf() {
  table_fmt = SSF_init_table();
}

// src/xlsx/styles.ts
function parse_numFmts(t, styles, opts) {
  const m = t.match(tagregex);
  if (!m) return;
  for (let i = 0; i < m.length; ++i) {
    const y = parsexmltag(m[i]);
    switch (strip_tag(y[0])) {
      case "<numFmt": {
        const f = unescapexml(y.formatCode);
        const j = parseInt(y.numFmtId, 10);
        styles.NumberFmt[j] = f;
        if (j > 0) {
          SSF_load(f, j);
        }
        break;
      }
    }
  }
}
function parse_cellXfs(t, styles) {
  const m = t.match(tagregex);
  if (!m) return;
  let xf = null;
  for (let i = 0; i < m.length; ++i) {
    const y = parsexmltag(m[i]);
    switch (strip_tag(y[0])) {
      case "<xf":
        xf = {
          numFmtId: parseInt(y.numFmtId, 10) || 0,
          fontId: parseInt(y.fontId, 10) || 0,
          fillId: parseInt(y.fillId, 10) || 0,
          borderId: parseInt(y.borderId, 10) || 0,
          xfId: parseInt(y.xfId, 10) || 0
        };
        if (y.applyNumberFormat) xf.applyNumberFormat = y.applyNumberFormat === "1";
        styles.CellXf.push(xf);
        break;
    }
  }
}
function strip_tag(tag) {
  return tag.replace(/<\w+:/, "<");
}
function parse_sty_xml(data, _themes, opts) {
  const styles = {
    NumberFmt: {},
    CellXf: [],
    Fonts: [],
    Fills: [],
    Borders: []
  };
  if (!data) return styles;
  const numFmts = data.match(/<(?:\w+:)?numFmts[^>]*>([\s\S]*?)<\/(?:\w+:)?numFmts>/);
  if (numFmts) parse_numFmts(numFmts[1], styles);
  const cellXfs = data.match(/<(?:\w+:)?cellXfs[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/);
  if (cellXfs) parse_cellXfs(cellXfs[1], styles);
  return styles;
}
function write_sty_xml(_wb, _opts) {
  const o = [XML_HEADER];
  o.push(
    writextag("styleSheet", null, {
      xmlns: XMLNS_main[0],
      "xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    })
  );
  o.push('<numFmts count="1"><numFmt numFmtId="164" formatCode="General"/></numFmts>');
  o.push(
    '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>'
  );
  o.push(
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
  );
  o.push(
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
  );
  o.push(
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  );
  o.push(
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
  );
  o.push(
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
  );
  o.push("</styleSheet>");
  o[1] = o[1].replace("/>", ">");
  return o.join("");
}

// src/xlsx/theme.ts
function parse_theme_xml(data) {
  const theme = { themeElements: { clrScheme: [] } };
  const colors = [];
  const clrMatch = data.match(/<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/);
  if (clrMatch) {
    const valRegex = /<a:(?:sysClr|srgbClr)[^>]*(?:val|lastClr)="([0-9A-Fa-f]{6})"/g;
    let m;
    while (m = valRegex.exec(clrMatch[1])) {
      colors.push(m[1]);
    }
  }
  theme.themeElements.clrScheme = colors;
  return theme;
}
function write_theme_xml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`;
}

// src/xlsx/workbook.ts
var WBPropsDef = [
  ["allowRefreshQuery", false, "bool"],
  ["autoCompressPictures", true, "bool"],
  ["backupFile", false, "bool"],
  ["checkCompatibility", false, "bool"],
  ["CodeName", ""],
  ["date1904", false, "bool"],
  ["defaultThemeVersion", 0, "int"],
  ["filterPrivacy", false, "bool"],
  ["hidePivotFieldList", false, "bool"],
  ["promptedSolutions", false, "bool"],
  ["publishItems", false, "bool"],
  ["refreshAllConnections", false, "bool"],
  ["saveExternalLinkValues", true, "bool"],
  ["showBorderUnselectedTables", true, "bool"],
  ["showInkAnnotation", true, "bool"],
  ["showObjects", "all"],
  ["showPivotChartFilter", false, "bool"],
  ["updateLinks", "userSet"]
];
var WBViewDef = [
  ["activeTab", 0, "int"],
  ["autoFilterDateGrouping", true, "bool"],
  ["firstSheet", 0, "int"],
  ["minimized", false, "bool"],
  ["showHorizontalScroll", true, "bool"],
  ["showSheetTabs", true, "bool"],
  ["showVerticalScroll", true, "bool"],
  ["tabRatio", 600, "int"],
  ["visibility", "visible"]
];
var SheetDef = [];
var CalcPrDef = [
  ["calcCompleted", "true"],
  ["calcMode", "auto"],
  ["calcOnSave", "true"],
  ["concurrentCalc", "true"],
  ["fullCalcOnLoad", "false"],
  ["fullPrecision", "true"],
  ["iterate", "false"],
  ["iterateCount", "100"],
  ["iterateDelta", "0.001"],
  ["refMode", "A1"]
];
function push_defaults_array(target, defaults) {
  for (let j = 0; j < target.length; ++j) {
    const w = target[j];
    for (let i = 0; i < defaults.length; ++i) {
      const z = defaults[i];
      if (w[z[0]] == null) w[z[0]] = z[1];
      else
        switch (z[2]) {
          case "bool":
            if (typeof w[z[0]] === "string") w[z[0]] = parsexmlbool(w[z[0]]);
            break;
          case "int":
            if (typeof w[z[0]] === "string") w[z[0]] = parseInt(w[z[0]], 10);
            break;
        }
    }
  }
}
function push_defaults(target, defaults) {
  for (let i = 0; i < defaults.length; ++i) {
    const z = defaults[i];
    if (target[z[0]] == null) target[z[0]] = z[1];
    else
      switch (z[2]) {
        case "bool":
          if (typeof target[z[0]] === "string") target[z[0]] = parsexmlbool(target[z[0]]);
          break;
        case "int":
          if (typeof target[z[0]] === "string") target[z[0]] = parseInt(target[z[0]], 10);
          break;
      }
  }
}
function parse_wb_defaults(wb) {
  push_defaults(wb.WBProps, WBPropsDef);
  push_defaults(wb.CalcPr, CalcPrDef);
  push_defaults_array(wb.WBView, WBViewDef);
  push_defaults_array(wb.Sheets, SheetDef);
}
var badchars = ":][*?/\\".split("");
function check_ws_name(n, safe) {
  try {
    if (n === "") throw new Error("Sheet name cannot be blank");
    if (n.length > 31) throw new Error("Sheet name cannot exceed 31 chars");
    if (n.charCodeAt(0) === 39 || n.charCodeAt(n.length - 1) === 39)
      throw new Error("Sheet name cannot start or end with apostrophe (')");
    if (n.toLowerCase() === "history") throw new Error("Sheet name cannot be 'History'");
    for (const c of badchars) {
      if (n.indexOf(c) !== -1) throw new Error("Sheet name cannot contain : \\ / ? * [ ]");
    }
  } catch (e) {
    throw e;
  }
  return true;
}
function check_wb_names(N, S) {
  for (let i = 0; i < N.length; ++i) {
    check_ws_name(N[i]);
    for (let j = 0; j < i; ++j) {
      if (N[i] === N[j]) throw new Error("Duplicate Sheet Name: " + N[i]);
    }
  }
}
function check_wb(wb) {
  if (!wb || !wb.SheetNames || !wb.Sheets) throw new Error("Invalid Workbook");
  if (!wb.SheetNames.length) throw new Error("Workbook is empty");
  wb.Workbook && wb.Workbook.Sheets || [];
  check_wb_names(wb.SheetNames);
}
var wbnsregex = /<\w+:workbook/;
function parse_wb_xml(data, opts) {
  if (!data) throw new Error("Could not find file");
  const wb = {
    AppVersion: {},
    WBProps: {},
    WBView: [],
    Sheets: [],
    CalcPr: {},
    Names: [],
    xmlns: ""
  };
  let xmlns = "xmlns";
  let dname = {};
  let dnstart = 0;
  data.replace(tagregex, function xml_wb(x, idx) {
    const y = parsexmltag(x);
    switch (strip_ns(y[0])) {
      case "<?xml":
        break;
      case "<workbook":
        if (x.match(wbnsregex)) xmlns = "xmlns" + x.match(/<(\w+):/)?.[1];
        wb.xmlns = y[xmlns];
        break;
      case "</workbook>":
        break;
      case "<fileVersion":
        delete y[0];
        wb.AppVersion = y;
        break;
      case "<fileVersion/>":
      case "</fileVersion>":
        break;
      case "<fileSharing":
      case "<fileSharing/>":
        break;
      case "<workbookPr":
      case "<workbookPr/>":
        WBPropsDef.forEach((w) => {
          if (y[w[0]] == null) return;
          switch (w[2]) {
            case "bool":
              wb.WBProps[w[0]] = parsexmlbool(y[w[0]]);
              break;
            case "int":
              wb.WBProps[w[0]] = parseInt(y[w[0]], 10);
              break;
            default:
              wb.WBProps[w[0]] = y[w[0]];
          }
        });
        if (y.codeName) wb.WBProps.CodeName = utf8read(y.codeName);
        break;
      case "</workbookPr>":
        break;
      case "<workbookProtection":
      case "<workbookProtection/>":
        break;
      case "<bookViews":
      case "<bookViews>":
      case "</bookViews>":
        break;
      case "<workbookView":
      case "<workbookView/>":
        delete y[0];
        wb.WBView.push(y);
        break;
      case "</workbookView>":
        break;
      case "<sheets":
      case "<sheets>":
      case "</sheets>":
        break;
      case "<sheet":
        switch (y.state) {
          case "hidden":
            y.Hidden = 1;
            break;
          case "veryHidden":
            y.Hidden = 2;
            break;
          default:
            y.Hidden = 0;
        }
        delete y.state;
        y.name = unescapexml(utf8read(y.name));
        delete y[0];
        wb.Sheets.push(y);
        break;
      case "</sheet>":
        break;
      case "<functionGroups":
      case "<functionGroups/>":
      case "<functionGroup":
        break;
      case "<externalReferences":
      case "</externalReferences>":
      case "<externalReferences>":
      case "<externalReference":
        break;
      case "<definedNames/>":
        break;
      case "<definedNames>":
      case "<definedNames":
        break;
      case "</definedNames>":
        break;
      case "<definedName": {
        dname = {};
        dname.Name = utf8read(y.name);
        if (y.comment) dname.Comment = y.comment;
        if (y.localSheetId) dname.Sheet = +y.localSheetId;
        if (parsexmlbool(y.hidden || "0")) dname.Hidden = true;
        dnstart = idx + x.length;
        break;
      }
      case "</definedName>": {
        dname.Ref = unescapexml(utf8read(data.slice(dnstart, idx)));
        wb.Names.push(dname);
        break;
      }
      case "<definedName/>":
        break;
      case "<calcPr":
      case "<calcPr/>":
        delete y[0];
        wb.CalcPr = y;
        break;
    }
    return x;
  });
  if (XMLNS_main.indexOf(wb.xmlns) === -1) throw new Error("Unknown Namespace: " + wb.xmlns);
  parse_wb_defaults(wb);
  return wb;
}
function write_wb_xml(wb) {
  const o = [XML_HEADER];
  o.push(
    writextag("workbook", null, {
      xmlns: XMLNS_main[0],
      "xmlns:r": XMLNS.r
    })
  );
  const write_names = !!(wb.Workbook && (wb.Workbook.Names || []).length > 0);
  const workbookPr = { codeName: "ThisWorkbook" };
  if (wb.Workbook && wb.Workbook.WBProps) {
    WBPropsDef.forEach((x) => {
      if (!wb.Workbook || !wb.Workbook.WBProps) return;
      const wbp = wb.Workbook.WBProps;
      if (wbp[x[0]] == null) return;
      if (wbp[x[0]] === x[1]) return;
      workbookPr[x[0]] = wbp[x[0]];
    });
    if (wb.Workbook.WBProps.CodeName) {
      workbookPr.codeName = wb.Workbook.WBProps.CodeName;
      delete workbookPr.CodeName;
    }
  }
  o.push(writextag("workbookPr", null, workbookPr));
  const sheets = wb.Workbook && wb.Workbook.Sheets || [];
  if (sheets[0] && !!sheets[0].Hidden) {
    o.push("<bookViews>");
    let i = 0;
    for (i = 0; i < wb.SheetNames.length; ++i) {
      if (!sheets[i]) break;
      if (!sheets[i].Hidden) break;
    }
    if (i === wb.SheetNames.length) i = 0;
    o.push('<workbookView firstSheet="' + i + '" activeTab="' + i + '"/>');
    o.push("</bookViews>");
  }
  o.push("<sheets>");
  for (let i = 0; i < wb.SheetNames.length; ++i) {
    const sht = { name: escapexml(wb.SheetNames[i].slice(0, 31)) };
    sht.sheetId = "" + (i + 1);
    sht["r:id"] = "rId" + (i + 1);
    if (sheets[i])
      switch (sheets[i].Hidden) {
        case 1:
          sht.state = "hidden";
          break;
        case 2:
          sht.state = "veryHidden";
          break;
      }
    o.push(writextag("sheet", null, sht));
  }
  o.push("</sheets>");
  if (write_names) {
    o.push("<definedNames>");
    if (wb.Workbook && wb.Workbook.Names)
      wb.Workbook.Names.forEach((n) => {
        const d = { name: n.Name };
        if (n.Comment) d.comment = n.Comment;
        if (n.Sheet != null) d.localSheetId = "" + n.Sheet;
        if (n.Hidden) d.hidden = "1";
        if (!n.Ref) return;
        o.push(writextag("definedName", escapexml(n.Ref), d));
      });
    o.push("</definedNames>");
  }
  if (o.length > 2) {
    o.push("</workbook>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}

// src/utils/cell.ts
function decode_row(rowstr) {
  return parseInt(unfix_row(rowstr), 10) - 1;
}
function encode_row(row) {
  return "" + (row + 1);
}
function unfix_row(cstr) {
  return cstr.replace(/\$(\d+)$/, "$1");
}
function decode_col(colstr) {
  const c = unfix_col(colstr);
  let d = 0;
  for (let i = 0; i < c.length; ++i) d = 26 * d + c.charCodeAt(i) - 64;
  return d - 1;
}
function encode_col(col) {
  if (col < 0) throw new Error("invalid column " + col);
  let s = "";
  for (++col; col; col = Math.floor((col - 1) / 26))
    s = String.fromCharCode((col - 1) % 26 + 65) + s;
  return s;
}
function unfix_col(cstr) {
  return cstr.replace(/^\$([A-Z])/, "$1");
}
function decode_cell(cstr) {
  let R = 0, C = 0;
  for (let i = 0; i < cstr.length; ++i) {
    const cc = cstr.charCodeAt(i);
    if (cc >= 48 && cc <= 57) R = 10 * R + (cc - 48);
    else if (cc >= 65 && cc <= 90) C = 26 * C + (cc - 64);
  }
  return { c: C - 1, r: R - 1 };
}
function encode_cell(cell) {
  let col = cell.c + 1;
  let s = "";
  for (; col; col = (col - 1) / 26 | 0)
    s = String.fromCharCode((col - 1) % 26 + 65) + s;
  return s + (cell.r + 1);
}
function decode_range(range) {
  const idx = range.indexOf(":");
  if (idx === -1) return { s: decode_cell(range), e: decode_cell(range) };
  return { s: decode_cell(range.slice(0, idx)), e: decode_cell(range.slice(idx + 1)) };
}
function encode_range(cs, ce) {
  if (typeof ce === "undefined" || typeof ce === "number") {
    return encode_range(cs.s, cs.e);
  }
  const s = typeof cs === "string" ? cs : encode_cell(cs);
  const e = typeof ce === "string" ? ce : encode_cell(ce);
  return s === e ? s : s + ":" + e;
}
function safe_decode_range(range) {
  const o = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
  let idx = 0, i = 0, cc = 0;
  const len = range.length;
  for (idx = 0; i < len; ++i) {
    if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) break;
    idx = 26 * idx + cc;
  }
  o.s.c = --idx;
  for (idx = 0; i < len; ++i) {
    if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) break;
    idx = 10 * idx + cc;
  }
  o.s.r = --idx;
  if (i === len || cc !== 10) {
    o.e.c = o.s.c;
    o.e.r = o.s.r;
    return o;
  }
  ++i;
  for (idx = 0; i !== len; ++i) {
    if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) break;
    idx = 26 * idx + cc;
  }
  o.e.c = --idx;
  for (idx = 0; i !== len; ++i) {
    if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) break;
    idx = 10 * idx + cc;
  }
  o.e.r = --idx;
  return o;
}

// src/utils/date.ts
function datenum(v, date1904) {
  let epoch = v.getTime();
  if (date1904) epoch -= 1462 * 24 * 60 * 60 * 1e3;
  const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
  const result = (epoch - dnthresh) / (24 * 60 * 60 * 1e3);
  if (result < 60) return result;
  if (result >= 60) return result + 1;
  return result;
}
function numdate(v, date1904) {
  let date = v;
  if (date > 60) --date;
  const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
  return new Date(dnthresh + date * 24 * 60 * 60 * 1e3);
}
function local_to_utc(d) {
  const off = d.getTimezoneOffset();
  return new Date(d.getTime() + off * 60 * 1e3);
}
function utc_to_local(d) {
  const off = d.getTimezoneOffset();
  return new Date(d.getTime() - off * 60 * 1e3);
}

// src/ssf/format.ts
function _strrev(x) {
  let o = "";
  let i = x.length - 1;
  while (i >= 0) o += x.charAt(i--);
  return o;
}
function pad0(v, d) {
  const t = "" + v;
  return t.length >= d ? t : fill("0", d - t.length) + t;
}
function pad_(v, d) {
  const t = "" + v;
  return t.length >= d ? t : fill(" ", d - t.length) + t;
}
function rpad_(v, d) {
  const t = "" + v;
  return t.length >= d ? t : t + fill(" ", d - t.length);
}
function pad0r1(v, d) {
  const t = "" + Math.round(v);
  return t.length >= d ? t : fill("0", d - t.length) + t;
}
function pad0r2(v, d) {
  const t = "" + v;
  return t.length >= d ? t : fill("0", d - t.length) + t;
}
var p2_32 = Math.pow(2, 32);
function pad0r(v, d) {
  if (v > p2_32 || v < -p2_32) return pad0r1(v, d);
  const i = Math.round(v);
  return pad0r2(i, d);
}
function SSF_isgeneral(s, i) {
  i = i || 0;
  return s.length >= 7 + i && (s.charCodeAt(i) | 32) === 103 && (s.charCodeAt(i + 1) | 32) === 101 && (s.charCodeAt(i + 2) | 32) === 110 && (s.charCodeAt(i + 3) | 32) === 101 && (s.charCodeAt(i + 4) | 32) === 114 && (s.charCodeAt(i + 5) | 32) === 97 && (s.charCodeAt(i + 6) | 32) === 108;
}
var days = [
  ["Sun", "Sunday"],
  ["Mon", "Monday"],
  ["Tue", "Tuesday"],
  ["Wed", "Wednesday"],
  ["Thu", "Thursday"],
  ["Fri", "Friday"],
  ["Sat", "Saturday"]
];
var months = [
  ["J", "Jan", "January"],
  ["F", "Feb", "February"],
  ["M", "Mar", "March"],
  ["A", "Apr", "April"],
  ["M", "May", "May"],
  ["J", "Jun", "June"],
  ["J", "Jul", "July"],
  ["A", "Aug", "August"],
  ["S", "Sep", "September"],
  ["O", "Oct", "October"],
  ["N", "Nov", "November"],
  ["D", "Dec", "December"]
];
function SSF_normalize_xl_unsafe(v) {
  const s = v.toPrecision(16);
  if (s.indexOf("e") > -1) {
    const m = s.slice(0, s.indexOf("e"));
    const ml = m.indexOf(".") > -1 ? m.slice(0, m.slice(0, 2) === "0." ? 17 : 16) : m.slice(0, 15) + fill("0", m.length - 15);
    return +ml + +("1" + s.slice(s.indexOf("e"))) - 1 || +s;
  }
  const n = s.indexOf(".") > -1 ? s.slice(0, s.slice(0, 2) === "0." ? 17 : 16) : s.slice(0, 15) + fill("0", s.length - 15);
  return Number(n);
}
function SSF_fix_hijri(_date, o) {
  o[0] -= 581;
  const dow = _date.getDay();
  if (_date.getTime() < -22038912e5) return (dow + 6) % 7;
  return dow;
}
function SSF_parse_date_code(v, opts, b2) {
  if (v > 2958465 || v < 0) return null;
  v = SSF_normalize_xl_unsafe(v);
  let date = v | 0;
  let time = Math.floor(86400 * (v - date));
  const out = {
    D: date,
    T: time,
    u: 86400 * (v - date) - time,
    y: 0,
    m: 0,
    d: 0,
    H: 0,
    M: 0,
    S: 0,
    q: 0
  };
  if (Math.abs(out.u) < 1e-6) out.u = 0;
  if (opts && opts.date1904) date += 1462;
  if (out.u > 0.9999) {
    out.u = 0;
    if (++time === 86400) {
      out.T = time = 0;
      ++date;
      ++out.D;
    }
  }
  let dout;
  let dow = 0;
  if (date === 60) {
    dout = b2 ? [1317, 10, 29] : [1900, 2, 29];
    dow = 3;
  } else if (date === 0) {
    dout = b2 ? [1317, 8, 29] : [1900, 1, 0];
    dow = 6;
  } else {
    if (date > 60) --date;
    const d = new Date(1900, 0, 1);
    d.setDate(d.getDate() + date - 1);
    dout = [d.getFullYear(), d.getMonth() + 1, d.getDate()];
    dow = d.getDay();
    if (date < 60) dow = (dow + 6) % 7;
    if (b2) dow = SSF_fix_hijri(d, dout);
  }
  out.y = dout[0];
  out.m = dout[1];
  out.d = dout[2];
  out.S = time % 60;
  time = Math.floor(time / 60);
  out.M = time % 60;
  time = Math.floor(time / 60);
  out.H = time;
  out.q = dow;
  return out;
}
function SSF_strip_decimal(o) {
  return o.indexOf(".") === -1 ? o : o.replace(/(?:\.0*|(\.\d*[1-9])0+)$/, "$1");
}
function SSF_normalize_exp(o) {
  if (o.indexOf("E") === -1) return o;
  return o.replace(/(?:\.0*|(\.\d*[1-9])0+)[Ee]/, "$1E").replace(/(E[+-])(\d)$/, "$10$2");
}
function SSF_small_exp(v) {
  const w = v < 0 ? 12 : 11;
  let o = SSF_strip_decimal(v.toFixed(12));
  if (o.length <= w) return o;
  o = v.toPrecision(10);
  if (o.length <= w) return o;
  return v.toExponential(5);
}
function SSF_large_exp(v) {
  const o = SSF_strip_decimal(v.toFixed(11));
  return o.length > (v < 0 ? 12 : 11) || o === "0" || o === "-0" ? v.toPrecision(6) : o;
}
function SSF_general_num(v) {
  if (!isFinite(v)) return isNaN(v) ? "#NUM!" : "#DIV/0!";
  const V = Math.floor(Math.log(Math.abs(v)) * Math.LOG10E);
  let o;
  if (V >= -4 && V <= -1) o = v.toPrecision(10 + V);
  else if (Math.abs(V) <= 9) o = SSF_small_exp(v);
  else if (V === 10) o = v.toFixed(10).substr(0, 12);
  else o = SSF_large_exp(v);
  return SSF_strip_decimal(SSF_normalize_exp(o.toUpperCase()));
}
function SSF_general(v, opts) {
  switch (typeof v) {
    case "string":
      return v;
    case "boolean":
      return v ? "TRUE" : "FALSE";
    case "number":
      return (v | 0) === v ? v.toString(10) : SSF_general_num(v);
    case "undefined":
      return "";
    case "object":
      if (v == null) return "";
      if (v instanceof Date) return SSF_format(14, datenum(v, opts && opts.date1904), opts);
  }
  throw new Error("unsupported value in General format: " + v);
}
function SSF_write_date(type, fmt, val, ss0) {
  let o = "";
  let ss = 0;
  let tt = 0;
  let y = val.y;
  let out = 0;
  let outl = 0;
  switch (type) {
    case 98:
      y = val.y + 543;
    /* falls through */
    case 121:
      switch (fmt.length) {
        case 1:
        case 2:
          out = y % 100;
          outl = 2;
          break;
        default:
          out = y % 1e4;
          outl = 4;
          break;
      }
      break;
    case 109:
      switch (fmt.length) {
        case 1:
        case 2:
          out = val.m;
          outl = fmt.length;
          break;
        case 3:
          return months[val.m - 1][1];
        case 5:
          return months[val.m - 1][0];
        default:
          return months[val.m - 1][2];
      }
      break;
    case 100:
      switch (fmt.length) {
        case 1:
        case 2:
          out = val.d;
          outl = fmt.length;
          break;
        case 3:
          return days[val.q][0];
        default:
          return days[val.q][1];
      }
      break;
    case 104:
      switch (fmt.length) {
        case 1:
        case 2:
          out = 1 + (val.H + 11) % 12;
          outl = fmt.length;
          break;
        default:
          throw "bad hour format: " + fmt;
      }
      break;
    case 72:
      switch (fmt.length) {
        case 1:
        case 2:
          out = val.H;
          outl = fmt.length;
          break;
        default:
          throw "bad hour format: " + fmt;
      }
      break;
    case 77:
      switch (fmt.length) {
        case 1:
        case 2:
          out = val.M;
          outl = fmt.length;
          break;
        default:
          throw "bad minute format: " + fmt;
      }
      break;
    case 115:
      if (fmt !== "s" && fmt !== "ss" && fmt !== ".0" && fmt !== ".00" && fmt !== ".000")
        throw "bad second format: " + fmt;
      if (val.u === 0 && (fmt === "s" || fmt === "ss")) return pad0(val.S, fmt.length);
      if (ss0 >= 2) tt = ss0 === 3 ? 1e3 : 100;
      else tt = ss0 === 1 ? 10 : 1;
      ss = Math.round(tt * (val.S + val.u));
      if (ss >= 60 * tt) ss = 0;
      if (fmt === "s") return ss === 0 ? "0" : "" + ss / tt;
      o = pad0(ss, 2 + ss0);
      if (fmt === "ss") return o.substr(0, 2);
      return "." + o.substr(2, fmt.length - 1);
    case 90:
      switch (fmt) {
        case "[h]":
        case "[hh]":
          out = val.D * 24 + val.H;
          break;
        case "[m]":
        case "[mm]":
          out = (val.D * 24 + val.H) * 60 + val.M;
          break;
        case "[s]":
        case "[ss]":
          out = ((val.D * 24 + val.H) * 60 + val.M) * 60 + (ss0 === 0 ? Math.round(val.S + val.u) : val.S);
          break;
        default:
          throw "bad abstime format: " + fmt;
      }
      outl = fmt.length === 3 ? 1 : 2;
      break;
    case 101:
      out = y;
      outl = 1;
      break;
  }
  return outl > 0 ? pad0(out, outl) : "";
}
function commaify(s) {
  const w = 3;
  if (s.length <= w) return s;
  const j = s.length % w;
  let o = s.substr(0, j);
  for (let i = j; i !== s.length; i += w) o += (o.length > 0 ? "," : "") + s.substr(i, w);
  return o;
}
var pct1 = /%/g;
function write_num_pct(type, fmt, val) {
  const sfmt = fmt.replace(pct1, "");
  const mul = fmt.length - sfmt.length;
  return write_num(type, sfmt, val * Math.pow(10, 2 * mul)) + fill("%", mul);
}
function write_num_cm(type, fmt, val) {
  let idx = fmt.length - 1;
  while (fmt.charCodeAt(idx - 1) === 44) --idx;
  return write_num(type, fmt.substr(0, idx), val / Math.pow(10, 3 * (fmt.length - idx)));
}
function write_num_exp(fmt, val) {
  let o;
  const idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
  if (fmt.match(/^#+0.0E\+0$/)) {
    if (val === 0) return "0.0E+0";
    if (val < 0) return "-" + write_num_exp(fmt, -val);
    const period = fmt.indexOf(".");
    const ee = Math.floor(Math.log(val) * Math.LOG10E) % period < 0 ? Math.floor(Math.log(val) * Math.LOG10E) % period + period : Math.floor(Math.log(val) * Math.LOG10E) % period;
    o = (val / Math.pow(10, ee)).toPrecision(idx + 1 + (period + ee) % period);
    if (o.indexOf("e") === -1) {
      const fakee = Math.floor(Math.log(val) * Math.LOG10E);
      if (o.indexOf(".") === -1)
        o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length + ee);
      else o += "E+" + (fakee - ee);
      while (o.substr(0, 2) === "0.") {
        o = o.charAt(0) + o.substr(2, period) + "." + o.substr(2 + period);
        o = o.replace(/^0+([1-9])/, "$1").replace(/^0+\./, "0.");
      }
      o = o.replace(/\+-/, "-");
    }
    o = o.replace(
      /^([+-]?)(\d*)\.(\d*)[Ee]/,
      ($$, $1, $2, $3) => $1 + $2 + $3.substr(0, (period + ee) % period) + "." + $3.substr(ee) + "E"
    );
  } else o = val.toExponential(idx);
  if (fmt.match(/E\+00$/) && o.match(/e[+-]\d$/))
    o = o.substr(0, o.length - 1) + "0" + o.charAt(o.length - 1);
  if (fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/, "e");
  return o.replace("e", "E");
}
function SSF_frac(x, D, mixed) {
  const sgn = x < 0 ? -1 : 1;
  let B = x * sgn;
  let P_2 = 0, P_1 = 1, P = 0;
  let Q_2 = 1, Q_1 = 0, Q = 0;
  let A = Math.floor(B);
  while (Q_1 < D) {
    A = Math.floor(B);
    P = A * P_1 + P_2;
    Q = A * Q_1 + Q_2;
    if (B - A < 5e-8) break;
    B = 1 / (B - A);
    P_2 = P_1;
    P_1 = P;
    Q_2 = Q_1;
    Q_1 = Q;
  }
  if (Q > D) {
    if (Q_1 > D) {
      Q = Q_2;
      P = P_2;
    } else {
      Q = Q_1;
      P = P_1;
    }
  }
  if (!mixed) return [0, sgn * P, Q];
  const q = Math.floor(sgn * P / Q);
  return [q, sgn * P - q * Q, Q];
}
var frac1 = /# (\?+)( ?)\/( ?)(\d+)/;
function write_num_f1(r, aval, sign) {
  const den = parseInt(r[4], 10);
  const rr = Math.round(aval * den);
  const base = Math.floor(rr / den);
  const myn = rr - base * den;
  const myd = den;
  return sign + (base === 0 ? "" : "" + base) + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad_(myn, r[1].length) + r[2] + "/" + r[3] + pad0(myd, r[4].length));
}
function write_num_f2(r, aval, sign) {
  return sign + (aval === 0 ? "" : "" + aval) + fill(" ", r[1].length + 2 + r[4].length);
}
var dec1 = /^#*0*\.([0#]+)/;
var closeparen = /\)[^)]*[0#]/;
var phone = /\(###\) ###\\?-####/;
function hashq(str) {
  let o = "";
  for (let i = 0; i !== str.length; ++i) {
    const cc = str.charCodeAt(i);
    switch (cc) {
      case 35:
        break;
      case 63:
        o += " ";
        break;
      case 48:
        o += "0";
        break;
      default:
        o += String.fromCharCode(cc);
    }
  }
  return o;
}
function rnd(val, d) {
  const sgn = val < 0 ? -1 : 1;
  const dd = Math.pow(10, d);
  return "" + sgn * (Math.round(sgn * val * dd) / dd);
}
function dec(val, d) {
  const _frac = val - Math.floor(val);
  const dd = Math.pow(10, d);
  if (d < ("" + Math.round(_frac * dd)).length) return 0;
  return Math.round(_frac * dd);
}
function carry(val, d) {
  if (d < ("" + Math.round((val - Math.floor(val)) * Math.pow(10, d))).length) return 1;
  return 0;
}
function flr(val) {
  if (val < 2147483647 && val > -2147483648) return "" + (val >= 0 ? val | 0 : val - 1 | 0);
  return "" + Math.floor(val);
}
function write_num_flt(type, fmt, val) {
  if (type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
    const ffmt = fmt.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
    if (val >= 0) return write_num_flt("n", ffmt, val);
    return "(" + write_num_flt("n", ffmt, -val) + ")";
  }
  if (fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm(type, fmt, val);
  if (fmt.indexOf("%") !== -1) return write_num_pct(type, fmt, val);
  if (fmt.indexOf("E") !== -1) return write_num_exp(fmt, val);
  if (fmt.charCodeAt(0) === 36)
    return "$" + write_num_flt(type, fmt.substr(fmt.charAt(1) === " " ? 2 : 1), val);
  let o;
  let r;
  let ri;
  let ff;
  const aval = Math.abs(val);
  const sign = val < 0 ? "-" : "";
  if (fmt.match(/^00+$/)) return sign + pad0r(aval, fmt.length);
  if (fmt.match(/^[#?]+$/)) {
    o = pad0r(val, 0);
    if (o === "0") o = "";
    return o.length > fmt.length ? o : hashq(fmt.substr(0, fmt.length - o.length)) + o;
  }
  if (r = fmt.match(frac1)) return write_num_f1(r, aval, sign);
  if (fmt.match(/^#+0+$/)) return sign + pad0r(aval, fmt.length - fmt.indexOf("0"));
  if (r = fmt.match(dec1)) {
    o = rnd(val, r[1].length).replace(/^([^\.]+)$/, "$1." + hashq(r[1])).replace(/\.$/, "." + hashq(r[1])).replace(/\.(\d*)$/, ($$, $1) => "." + $1 + fill("0", hashq(r[1]).length - $1.length));
    return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./, ".");
  }
  fmt = fmt.replace(/^#+([0.])/, "$1");
  if (r = fmt.match(/^(0*)\.(#*)$/)) {
    return sign + rnd(aval, r[2].length).replace(/\.(\d*[1-9])0*$/, ".$1").replace(/^(-?\d*)$/, "$1.").replace(/^0\./, r[1].length ? "0." : ".");
  }
  if (fmt.match(/^#{1,3},##0(\.?)$/)) return sign + commaify(pad0r(aval, 0));
  if (r = fmt.match(/^#,##0\.([#0]*0)$/)) {
    return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify("" + (Math.floor(val) + carry(val, r[1].length))) + "." + pad0(dec(val, r[1].length), r[1].length);
  }
  if (r = fmt.match(/^#,#*,#0/)) return write_num_flt(type, fmt.replace(/^#,#*,/, ""), val);
  if (r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/)) {
    o = _strrev(write_num_flt(type, fmt.replace(/[\\-]/g, ""), val));
    ri = 0;
    return _strrev(
      _strrev(fmt.replace(/\\/g, "")).replace(/[0#]/g, (x) => {
        return ri < o.length ? o.charAt(ri++) : x === "0" ? "0" : "";
      })
    );
  }
  if (fmt.match(phone)) {
    o = write_num_flt(type, "##########", val);
    return "(" + o.substr(0, 3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
  }
  let oa = "";
  if (r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/)) {
    ri = Math.min(r[4].length, 7);
    ff = SSF_frac(aval, Math.pow(10, ri) - 1, false);
    o = "" + sign;
    oa = write_num("n", r[1], ff[1]);
    if (oa.charAt(oa.length - 1) === " ") oa = oa.substr(0, oa.length - 1) + "0";
    o += oa + r[2] + "/" + r[3];
    oa = rpad_(ff[2], ri);
    if (oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length - oa.length)) + oa;
    o += oa;
    return o;
  }
  if (r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/)) {
    ri = Math.min(Math.max(r[1].length, r[4].length), 7);
    ff = SSF_frac(aval, Math.pow(10, ri) - 1, true);
    return sign + (ff[0] || (ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1], ri) + r[2] + "/" + r[3] + rpad_(ff[2], ri) : fill(" ", 2 * ri + 1 + r[2].length + r[3].length));
  }
  if (r = fmt.match(/^[#0?]+$/)) {
    o = pad0r(val, 0);
    if (fmt.length <= o.length) return o;
    return hashq(fmt.substr(0, fmt.length - o.length)) + o;
  }
  if (r = fmt.match(/^([#0?]+)\.([#0]+)$/)) {
    o = "" + val.toFixed(Math.min(r[2].length, 10)).replace(/([^0])0+$/, "$1");
    ri = o.indexOf(".");
    const lres = fmt.indexOf(".") - ri;
    const rres = fmt.length - o.length - lres;
    return hashq(fmt.substr(0, lres) + o + fmt.substr(fmt.length - rres));
  }
  if (r = fmt.match(/^00,000\.([#0]*0)$/)) {
    ri = dec(val, r[1].length);
    return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(flr(val)).replace(/^\d,\d{3}$/, "0$&").replace(/^\d*$/, ($$) => "00," + ($$.length < 3 ? pad0(0, 3 - $$.length) : "") + $$) + "." + pad0(ri, r[1].length);
  }
  switch (fmt) {
    case "###,##0.00":
      return write_num_flt(type, "#,##0.00", val);
    case "###,###":
    case "##,###":
    case "#,###": {
      const x = commaify(pad0r(aval, 0));
      return x !== "0" ? sign + x : "";
    }
    case "###,###.00":
      return write_num_flt(type, "###,##0.00", val).replace(/^0\./, ".");
    case "#,###.00":
      return write_num_flt(type, "#,##0.00", val).replace(/^0\./, ".");
  }
  throw new Error("unsupported format |" + fmt + "|");
}
function write_num_int(type, fmt, val) {
  if (type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
    const ffmt = fmt.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
    if (val >= 0) return write_num_int("n", ffmt, val);
    return "(" + write_num_int("n", ffmt, -val) + ")";
  }
  if (fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm(type, fmt, val);
  if (fmt.indexOf("%") !== -1) return write_num_pct(type, fmt, val);
  if (fmt.indexOf("E") !== -1) return write_num_exp(fmt, val);
  if (fmt.charCodeAt(0) === 36)
    return "$" + write_num_int(type, fmt.substr(fmt.charAt(1) === " " ? 2 : 1), val);
  let o;
  let r;
  let ri;
  let ff;
  const aval = Math.abs(val);
  const sign = val < 0 ? "-" : "";
  if (fmt.match(/^00+$/)) return sign + pad0(aval, fmt.length);
  if (fmt.match(/^[#?]+$/)) {
    o = "" + val;
    if (val === 0) o = "";
    return o.length > fmt.length ? o : hashq(fmt.substr(0, fmt.length - o.length)) + o;
  }
  if (r = fmt.match(frac1)) return write_num_f2(r, aval, sign);
  if (fmt.match(/^#+0+$/)) return sign + pad0(aval, fmt.length - fmt.indexOf("0"));
  if (r = fmt.match(dec1)) {
    o = ("" + val).replace(/^([^\.]+)$/, "$1." + hashq(r[1])).replace(/\.$/, "." + hashq(r[1]));
    o = o.replace(/\.(\d*)$/, ($$, $1) => "." + $1 + fill("0", hashq(r[1]).length - $1.length));
    return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./, ".");
  }
  fmt = fmt.replace(/^#+([0.])/, "$1");
  if (r = fmt.match(/^(0*)\.(#*)$/)) {
    return sign + ("" + aval).replace(/\.(\d*[1-9])0*$/, ".$1").replace(/^(-?\d*)$/, "$1.").replace(/^0\./, r[1].length ? "0." : ".");
  }
  if (fmt.match(/^#{1,3},##0(\.?)$/)) return sign + commaify("" + aval);
  if (r = fmt.match(/^#,##0\.([#0]*0)$/)) {
    return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify("" + val) + "." + fill("0", r[1].length);
  }
  if (r = fmt.match(/^#,#*,#0/)) return write_num_int(type, fmt.replace(/^#,#*,/, ""), val);
  if (r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/)) {
    o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g, ""), val));
    ri = 0;
    return _strrev(
      _strrev(fmt.replace(/\\/g, "")).replace(/[0#]/g, (x) => {
        return ri < o.length ? o.charAt(ri++) : x === "0" ? "0" : "";
      })
    );
  }
  if (fmt.match(phone)) {
    o = write_num_int(type, "##########", val);
    return "(" + o.substr(0, 3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
  }
  let oa = "";
  if (r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/)) {
    ri = Math.min(r[4].length, 7);
    ff = SSF_frac(aval, Math.pow(10, ri) - 1, false);
    o = "" + sign;
    oa = write_num("n", r[1], ff[1]);
    if (oa.charAt(oa.length - 1) === " ") oa = oa.substr(0, oa.length - 1) + "0";
    o += oa + r[2] + "/" + r[3];
    oa = rpad_(ff[2], ri);
    if (oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length - oa.length)) + oa;
    o += oa;
    return o;
  }
  if (r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/)) {
    ri = Math.min(Math.max(r[1].length, r[4].length), 7);
    ff = SSF_frac(aval, Math.pow(10, ri) - 1, true);
    return sign + (ff[0] || (ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1], ri) + r[2] + "/" + r[3] + rpad_(ff[2], ri) : fill(" ", 2 * ri + 1 + r[2].length + r[3].length));
  }
  if (r = fmt.match(/^[#0?]+$/)) {
    o = "" + val;
    if (fmt.length <= o.length) return o;
    return hashq(fmt.substr(0, fmt.length - o.length)) + o;
  }
  if (r = fmt.match(/^([#0]+)\.([#0]+)$/)) {
    o = "" + val.toFixed(Math.min(r[2].length, 10)).replace(/([^0])0+$/, "$1");
    ri = o.indexOf(".");
    const lres = fmt.indexOf(".") - ri;
    const rres = fmt.length - o.length - lres;
    return hashq(fmt.substr(0, lres) + o + fmt.substr(fmt.length - rres));
  }
  if (r = fmt.match(/^00,000\.([#0]*0)$/)) {
    return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify("" + val).replace(/^\d,\d{3}$/, "0$&").replace(/^\d*$/, ($$) => "00," + ($$.length < 3 ? pad0(0, 3 - $$.length) : "") + $$) + "." + pad0(0, r[1].length);
  }
  switch (fmt) {
    case "###,###":
    case "##,###":
    case "#,###": {
      const x = commaify("" + aval);
      return x !== "0" ? sign + x : "";
    }
    default:
      if (fmt.match(/\.[0#?]*$/))
        return write_num_int(type, fmt.slice(0, fmt.lastIndexOf(".")), val) + hashq(fmt.slice(fmt.lastIndexOf(".")));
  }
  throw new Error("unsupported format |" + fmt + "|");
}
function write_num(type, fmt, val) {
  return (val | 0) === val ? write_num_int(type, fmt, val) : write_num_flt(type, fmt, val);
}
function SSF_split_fmt(fmt) {
  const out = [];
  let in_str = false;
  let j = 0;
  for (let i = 0; i < fmt.length; ++i) {
    switch (fmt.charCodeAt(i)) {
      case 34:
        in_str = !in_str;
        break;
      case 95:
      case 42:
      case 92:
        ++i;
        break;
      case 59:
        out[out.length] = fmt.substr(j, i - j);
        j = i + 1;
    }
  }
  out[out.length] = fmt.substr(j);
  if (in_str === true) throw new Error("Format |" + fmt + "| unterminated string ");
  return out;
}
var SSF_abstime = /\[[HhMmSs\u0E0A\u0E19\u0E17]*\]/;
function fmt_is_date(fmt) {
  let i = 0;
  let c = "";
  let o = "";
  while (i < fmt.length) {
    switch (c = fmt.charAt(i)) {
      case "G":
        if (SSF_isgeneral(fmt, i)) i += 6;
        i++;
        break;
      case '"':
        for (; fmt.charCodeAt(++i) !== 34 && i < fmt.length; ) {
        }
        ++i;
        break;
      case "\\":
        i += 2;
        break;
      case "_":
        i += 2;
        break;
      case "@":
        ++i;
        break;
      case "B":
      case "b":
        if (fmt.charAt(i + 1) === "1" || fmt.charAt(i + 1) === "2") return true;
      /* falls through */
      case "M":
      case "D":
      case "Y":
      case "H":
      case "S":
      case "E":
      case "m":
      case "d":
      case "y":
      case "h":
      case "s":
      case "e":
      case "g":
        return true;
      case "A":
      case "a":
      case "\u4E0A":
        if (fmt.substr(i, 3).toUpperCase() === "A/P") return true;
        if (fmt.substr(i, 5).toUpperCase() === "AM/PM") return true;
        if (fmt.substr(i, 5).toUpperCase() === "\u4E0A\u5348/\u4E0B\u5348") return true;
        ++i;
        break;
      case "[":
        o = c;
        while (fmt.charAt(i++) !== "]" && i < fmt.length) o += fmt.charAt(i);
        if (o.match(SSF_abstime)) return true;
        break;
      case ".":
      case "0":
      case "#":
        while (i < fmt.length && ("0#?.,E+-%".indexOf(c = fmt.charAt(++i)) > -1 || c === "\\" && fmt.charAt(i + 1) === "-" && "0#".indexOf(fmt.charAt(i + 2)) > -1)) {
        }
        break;
      case "?":
        while (fmt.charAt(++i) === c) {
        }
        break;
      case "*":
        ++i;
        if (fmt.charAt(i) === " " || fmt.charAt(i) === "*") ++i;
        break;
      case "(":
      case ")":
        ++i;
        break;
      case "1":
      case "2":
      case "3":
      case "4":
      case "5":
      case "6":
      case "7":
      case "8":
      case "9":
        while (i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) {
        }
        break;
      case " ":
        ++i;
        break;
      default:
        ++i;
        break;
    }
  }
  return false;
}
function eval_fmt(fmt, v, opts, flen) {
  const out = [];
  let o = "";
  let i = 0;
  let c = "";
  let lst = "t";
  let dt = null;
  let j;
  let cc;
  let hr = "H";
  while (i < fmt.length) {
    switch (c = fmt.charAt(i)) {
      case "G":
        if (!SSF_isgeneral(fmt, i)) throw new Error("unrecognized character " + c + " in " + fmt);
        out[out.length] = { t: "G", v: "General" };
        i += 7;
        break;
      case '"':
        for (o = ""; (cc = fmt.charCodeAt(++i)) !== 34 && i < fmt.length; )
          o += String.fromCharCode(cc);
        out[out.length] = { t: "t", v: o };
        ++i;
        break;
      case "\\": {
        const w = fmt.charAt(++i);
        const t2 = w === "(" || w === ")" ? w : "t";
        out[out.length] = { t: t2, v: w };
        ++i;
        break;
      }
      case "_":
        out[out.length] = { t: "t", v: " " };
        i += 2;
        break;
      case "@":
        out[out.length] = { t: "T", v };
        ++i;
        break;
      case "B":
      case "b":
        if (fmt.charAt(i + 1) === "1" || fmt.charAt(i + 1) === "2") {
          if (dt == null) {
            dt = SSF_parse_date_code(v, opts, fmt.charAt(i + 1) === "2");
            if (dt == null) return "";
          }
          out[out.length] = { t: "X", v: fmt.substr(i, 2) };
          lst = c;
          i += 2;
          break;
        }
      /* falls through */
      case "M":
      case "D":
      case "Y":
      case "H":
      case "S":
      case "E":
        c = c.toLowerCase();
      /* falls through */
      case "m":
      case "d":
      case "y":
      case "h":
      case "s":
      case "e":
      case "g":
        if (v < 0) return "";
        if (dt == null) {
          dt = SSF_parse_date_code(v, opts);
          if (dt == null) return "";
        }
        o = c;
        while (++i < fmt.length && fmt.charAt(i).toLowerCase() === c) o += c;
        if (c === "m" && lst.toLowerCase() === "h") c = "M";
        if (c === "h") c = hr;
        out[out.length] = { t: c, v: o };
        lst = c;
        break;
      case "A":
      case "a":
      case "\u4E0A": {
        const q = { t: c, v: c };
        if (dt == null) dt = SSF_parse_date_code(v, opts);
        if (fmt.substr(i, 3).toUpperCase() === "A/P") {
          if (dt != null) q.v = dt.H >= 12 ? fmt.charAt(i + 2) : c;
          q.t = "T";
          hr = "h";
          i += 3;
        } else if (fmt.substr(i, 5).toUpperCase() === "AM/PM") {
          if (dt != null) q.v = dt.H >= 12 ? "PM" : "AM";
          q.t = "T";
          i += 5;
          hr = "h";
        } else if (fmt.substr(i, 5).toUpperCase() === "\u4E0A\u5348/\u4E0B\u5348") {
          if (dt != null) q.v = dt.H >= 12 ? "\u4E0B\u5348" : "\u4E0A\u5348";
          q.t = "T";
          i += 5;
          hr = "h";
        } else {
          q.t = "t";
          ++i;
        }
        if (dt == null && q.t === "T") return "";
        out[out.length] = q;
        lst = c;
        break;
      }
      case "[":
        o = c;
        while (fmt.charAt(i++) !== "]" && i < fmt.length) o += fmt.charAt(i);
        if (o.slice(-1) !== "]") throw 'unterminated "[" block: |' + o + "|";
        if (o.match(SSF_abstime)) {
          if (dt == null) {
            dt = SSF_parse_date_code(v, opts);
            if (dt == null) return "";
          }
          out[out.length] = { t: "Z", v: o.toLowerCase() };
          lst = o.charAt(1);
        } else if (o.indexOf("$") > -1) {
          o = (o.match(/\$([^-\[\]]*)/) || [])[1] || "$";
          if (!fmt_is_date(fmt)) out[out.length] = { t: "t", v: o };
        }
        break;
      case ".":
        if (dt != null) {
          o = c;
          while (++i < fmt.length && (c = fmt.charAt(i)) === "0") o += c;
          out[out.length] = { t: "s", v: o };
          break;
        }
      /* falls through */
      case "0":
      case "#":
        o = c;
        while (++i < fmt.length && "0#?.,E+-%".indexOf(c = fmt.charAt(i)) > -1) o += c;
        out[out.length] = { t: "n", v: o };
        break;
      case "?":
        o = c;
        while (fmt.charAt(++i) === c) o += c;
        out[out.length] = { t: c, v: o };
        lst = c;
        break;
      case "*":
        ++i;
        if (fmt.charAt(i) === " " || fmt.charAt(i) === "*") ++i;
        break;
      case "(":
      case ")":
        out[out.length] = { t: flen === 1 ? "t" : c, v: c };
        ++i;
        break;
      case "1":
      case "2":
      case "3":
      case "4":
      case "5":
      case "6":
      case "7":
      case "8":
      case "9":
        o = c;
        while (i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) o += fmt.charAt(i);
        out[out.length] = { t: "D", v: o };
        break;
      case " ":
        out[out.length] = { t: c, v: c };
        ++i;
        break;
      case "$":
        out[out.length] = { t: "t", v: "$" };
        ++i;
        break;
      default:
        if (",$-+/():!^&'~{}<>=\u20ACacfijklopqrtuvwxzP".indexOf(c) === -1)
          throw new Error("unrecognized character " + c + " in " + fmt);
        out[out.length] = { t: "t", v: c };
        ++i;
        break;
    }
  }
  let bt = 0;
  let ss0 = 0;
  let ssm;
  for (i = out.length - 1, lst = "t"; i >= 0; --i) {
    if (!out[i]) continue;
    switch (out[i].t) {
      case "h":
      case "H":
        out[i].t = hr;
        lst = "h";
        if (bt < 1) bt = 1;
        break;
      case "s":
        if (ssm = out[i].v.match(/\.0+$/)) {
          ss0 = Math.max(ss0, ssm[0].length - 1);
          bt = 4;
        }
        if (bt < 3) bt = 3;
      /* falls through */
      case "d":
      case "y":
      case "e":
        lst = out[i].t;
        break;
      case "M":
        lst = out[i].t;
        if (bt < 2) bt = 2;
        break;
      case "m":
        if (lst === "s") {
          out[i].t = "M";
          if (bt < 2) bt = 2;
        }
        break;
      case "X":
        break;
      case "Z":
        if (bt < 1 && out[i].v.match(/[Hh]/)) bt = 1;
        if (bt < 2 && out[i].v.match(/[Mm]/)) bt = 2;
        if (bt < 3 && out[i].v.match(/[Ss]/)) bt = 3;
    }
  }
  if (dt) {
    let _dt;
    switch (bt) {
      case 0:
        break;
      case 1:
      case 2:
      case 3:
        if (dt.u >= 0.5) {
          dt.u = 0;
          ++dt.S;
        }
        if (dt.S >= 60) {
          dt.S = 0;
          ++dt.M;
        }
        if (dt.M >= 60) {
          dt.M = 0;
          ++dt.H;
        }
        if (dt.H >= 24) {
          dt.H = 0;
          ++dt.D;
          _dt = SSF_parse_date_code(dt.D);
          if (_dt) {
            _dt.u = dt.u;
            _dt.S = dt.S;
            _dt.M = dt.M;
            _dt.H = dt.H;
            dt = _dt;
          }
        }
        break;
      case 4:
        switch (ss0) {
          case 1:
            dt.u = Math.round(dt.u * 10) / 10;
            break;
          case 2:
            dt.u = Math.round(dt.u * 100) / 100;
            break;
          case 3:
            dt.u = Math.round(dt.u * 1e3) / 1e3;
            break;
        }
        if (dt.u >= 1) {
          dt.u = 0;
          ++dt.S;
        }
        if (dt.S >= 60) {
          dt.S = 0;
          ++dt.M;
        }
        if (dt.M >= 60) {
          dt.M = 0;
          ++dt.H;
        }
        if (dt.H >= 24) {
          dt.H = 0;
          ++dt.D;
          _dt = SSF_parse_date_code(dt.D);
          if (_dt) {
            _dt.u = dt.u;
            _dt.S = dt.S;
            _dt.M = dt.M;
            _dt.H = dt.H;
            dt = _dt;
          }
        }
        break;
    }
  }
  let nstr = "";
  let jj;
  for (i = 0; i < out.length; ++i) {
    if (!out[i]) continue;
    switch (out[i].t) {
      case "t":
      case "T":
      case " ":
      case "D":
        break;
      case "X":
        out[i].v = "";
        out[i].t = ";";
        break;
      case "d":
      case "m":
      case "y":
      case "h":
      case "H":
      case "M":
      case "s":
      case "e":
      case "b":
      case "Z":
        out[i].v = SSF_write_date(out[i].t.charCodeAt(0), out[i].v, dt, ss0);
        out[i].t = "t";
        break;
      case "n":
      case "?":
        jj = i + 1;
        while (out[jj] != null && ((c = out[jj].t) === "?" || c === "D" || (c === " " || c === "t") && out[jj + 1] != null && (out[jj + 1].t === "?" || out[jj + 1].t === "t" && out[jj + 1].v === "/") || out[i].t === "(" && (c === " " || c === "n" || c === ")") || c === "t" && (out[jj].v === "/" || out[jj].v === " " && out[jj + 1] != null && out[jj + 1].t === "?"))) {
          out[i].v += out[jj].v;
          out[jj] = { v: "", t: ";" };
          ++jj;
        }
        nstr += out[i].v;
        i = jj - 1;
        break;
      case "G":
        out[i].t = "t";
        out[i].v = SSF_general(v, opts);
        break;
    }
  }
  let vv = "";
  let myv;
  let ostr;
  if (nstr.length > 0) {
    if (nstr.charCodeAt(0) === 40) {
      myv = v < 0 && nstr.charCodeAt(0) === 45 ? -v : v;
      ostr = write_num("n", nstr, myv);
    } else {
      myv = v < 0 && flen > 1 ? -v : v;
      ostr = write_num("n", nstr, myv);
      if (myv < 0 && out[0] && out[0].t === "t") {
        ostr = ostr.substr(1);
        out[0].v = "-" + out[0].v;
      }
    }
    jj = ostr.length - 1;
    let decpt = out.length;
    for (i = 0; i < out.length; ++i)
      if (out[i] != null && out[i].t !== "t" && out[i].v.indexOf(".") > -1) {
        decpt = i;
        break;
      }
    let lasti = out.length;
    if (decpt === out.length && ostr.indexOf("E") === -1) {
      for (i = out.length - 1; i >= 0; --i) {
        if (out[i] == null || "n?".indexOf(out[i].t) === -1) continue;
        if (jj >= out[i].v.length - 1) {
          jj -= out[i].v.length;
          out[i].v = ostr.substr(jj + 1, out[i].v.length);
        } else if (jj < 0) out[i].v = "";
        else {
          out[i].v = ostr.substr(0, jj + 1);
          jj = -1;
        }
        out[i].t = "t";
        lasti = i;
      }
      if (jj >= 0 && lasti < out.length) out[lasti].v = ostr.substr(0, jj + 1) + out[lasti].v;
    } else if (decpt !== out.length && ostr.indexOf("E") === -1) {
      jj = ostr.indexOf(".") - 1;
      for (i = decpt; i >= 0; --i) {
        if (out[i] == null || "n?".indexOf(out[i].t) === -1) continue;
        j = out[i].v.indexOf(".") > -1 && i === decpt ? out[i].v.indexOf(".") - 1 : out[i].v.length - 1;
        vv = out[i].v.substr(j + 1);
        for (; j >= 0; --j) {
          if (jj >= 0 && (out[i].v.charAt(j) === "0" || out[i].v.charAt(j) === "#"))
            vv = ostr.charAt(jj--) + vv;
        }
        out[i].v = vv;
        out[i].t = "t";
        lasti = i;
      }
      if (jj >= 0 && lasti < out.length) out[lasti].v = ostr.substr(0, jj + 1) + out[lasti].v;
      jj = ostr.indexOf(".") + 1;
      for (i = decpt; i < out.length; ++i) {
        if (out[i] == null || "n?(".indexOf(out[i].t) === -1 && i !== decpt) continue;
        j = out[i].v.indexOf(".") > -1 && i === decpt ? out[i].v.indexOf(".") + 1 : 0;
        vv = out[i].v.substr(0, j);
        for (; j < out[i].v.length; ++j) {
          if (jj < ostr.length) vv += ostr.charAt(jj++);
        }
        out[i].v = vv;
        out[i].t = "t";
        lasti = i;
      }
    }
  }
  for (i = 0; i < out.length; ++i) {
    if (out[i] != null && "n?".indexOf(out[i].t) > -1) {
      myv = flen > 1 && v < 0 && i > 0 && out[i - 1].v === "-" ? -v : v;
      out[i].v = write_num(out[i].t, out[i].v, myv);
      out[i].t = "t";
    }
  }
  let retval = "";
  for (i = 0; i !== out.length; ++i) if (out[i] != null) retval += out[i].v;
  return retval;
}
var cfregex2 = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
function chkcond(v, rr) {
  if (rr == null) return false;
  const thresh = parseFloat(rr[2]);
  switch (rr[1]) {
    case "=":
      if (v == thresh) return true;
      break;
    case ">":
      if (v > thresh) return true;
      break;
    case "<":
      if (v < thresh) return true;
      break;
    case "<>":
      if (v != thresh) return true;
      break;
    case ">=":
      if (v >= thresh) return true;
      break;
    case "<=":
      if (v <= thresh) return true;
      break;
  }
  return false;
}
function choose_fmt(f, v) {
  let fmt = SSF_split_fmt(f);
  const l = fmt.length;
  const lat = fmt[l - 1].indexOf("@");
  let ll = l;
  if (l < 4 && lat > -1) --ll;
  if (fmt.length > 4) throw new Error("cannot find right format for |" + fmt.join("|") + "|");
  if (typeof v !== "number")
    return [4, fmt.length === 4 || lat > -1 ? fmt[fmt.length - 1] : "@"];
  if (typeof v === "number" && !isFinite(v)) v = 0;
  switch (fmt.length) {
    case 1:
      fmt = lat > -1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"];
      break;
    case 2:
      fmt = lat > -1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"];
      break;
    case 3:
      fmt = lat > -1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"];
      break;
  }
  const ff = v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
  if (fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1) return [ll, ff];
  if (fmt[0].match(/\[[=<>]/) != null || fmt[1].match(/\[[=<>]/) != null) {
    const m1 = fmt[0].match(cfregex2);
    const m2 = fmt[1].match(cfregex2);
    return chkcond(v, m1) ? [ll, fmt[0]] : chkcond(v, m2) ? [ll, fmt[1]] : [ll, fmt[m1 != null && m2 != null ? 2 : 1]];
  }
  return [ll, ff];
}
function SSF_format(fmt, v, o) {
  if (o == null) o = {};
  let sfmt = "";
  switch (typeof fmt) {
    case "string":
      if (fmt === "m/d/yy" && o.dateNF) sfmt = o.dateNF;
      else sfmt = fmt;
      break;
    case "number":
      if (fmt === 14 && o.dateNF) sfmt = o.dateNF;
      else sfmt = (o.table != null ? o.table : table_fmt)[fmt];
      if (sfmt == null) sfmt = o.table && o.table[SSF_default_map[fmt]] || table_fmt[SSF_default_map[fmt]];
      if (sfmt == null) sfmt = SSF_default_str[fmt] || "General";
      break;
  }
  if (SSF_isgeneral(sfmt, 0)) return SSF_general(v, o);
  if (v instanceof Date) v = datenum(v, o.date1904);
  const f = choose_fmt(sfmt, v);
  if (SSF_isgeneral(f[1])) return SSF_general(v, o);
  if (v === true) v = "TRUE";
  else if (v === false) v = "FALSE";
  else if (v === "" || v == null) return "";
  else if (isNaN(v) && f[1].indexOf("0") > -1) return "#NUM!";
  else if (!isFinite(v) && f[1].indexOf("0") > -1) return "#DIV/0!";
  return eval_fmt(f[1], v, o, f[0]);
}

// src/xlsx/worksheet.ts
var mergecregex = /<(?:\w+:)?mergeCell ref=["'][A-Z0-9:]+['"]\s*[\/]?>/g;
var hlinkregex = /<(?:\w+:)?hyperlink [^<>]*>/gm;
var dimregex = /"(\w*:\w*)"/;
var colregex = /<(?:\w+:)?col\b[^<>]*[\/]?>/g;
var afregex = /<(?:\w:)?autoFilter[^>]*([\/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
var marginregex = /<(?:\w+:)?pageMargins[^<>]*\/>/g;
function parse_ws_xml_dim(ws, s) {
  const d = safe_decode_range(s);
  if (d.s.r <= d.e.r && d.s.c <= d.e.c && d.s.r >= 0 && d.s.c >= 0)
    ws["!ref"] = encode_range(d);
}
function parse_ws_xml_margins(tag) {
  return {
    left: parseFloat(tag.left) || 0.7,
    right: parseFloat(tag.right) || 0.7,
    top: parseFloat(tag.top) || 0.75,
    bottom: parseFloat(tag.bottom) || 0.75,
    header: parseFloat(tag.header) || 0.3,
    footer: parseFloat(tag.footer) || 0.3
  };
}
function parse_ws_xml_autofilter(data) {
  const tag = parsexmltag(data.match(/<[^>]*>/)?.[0] || "");
  return { ref: tag.ref || "" };
}
function parse_ws_xml_cols(columns, cols) {
  for (let i = 0; i < cols.length; ++i) {
    const tag = parsexmltag(cols[i]);
    if (!tag.min || !tag.max) continue;
    const min = parseInt(tag.min, 10) - 1;
    const max = parseInt(tag.max, 10) - 1;
    const width = tag.width ? parseFloat(tag.width) : void 0;
    const hidden = tag.hidden === "1";
    for (let j = min; j <= max; ++j) {
      if (!columns[j]) columns[j] = {};
      if (width !== void 0) columns[j].width = width;
      if (hidden) columns[j].hidden = true;
    }
  }
}
function parse_ws_xml_hlinks(s, hlinks, rels) {
  for (let i = 0; i < hlinks.length; ++i) {
    const tag = parsexmltag(hlinks[i]);
    if (!tag.ref) continue;
    const rng = safe_decode_range(tag.ref);
    for (let R = rng.s.r; R <= rng.e.r; ++R) {
      for (let C = rng.s.c; C <= rng.e.c; ++C) {
        const addr = encode_cell({ r: R, c: C });
        const dense = s["!data"] != null;
        let cell;
        if (dense) {
          if (!s["!data"][R]) s["!data"][R] = [];
          cell = s["!data"][R][C];
        } else {
          cell = s[addr];
        }
        if (!cell) {
          cell = { t: "z", v: void 0 };
          if (dense) s["!data"][R][C] = cell;
          else s[addr] = cell;
        }
        let target = "";
        if (tag.id) {
          const rel = rels["!id"]?.[tag.id];
          if (rel) target = rel.Target;
        }
        if (tag.location) target += "#" + tag.location;
        cell.l = { Target: target };
        if (tag.tooltip) cell.l.Tooltip = tag.tooltip;
      }
    }
  }
}
var rowregex = /<(?:\w+:)?row\b[^>]*>/g;
var cellregex = /<(?:\w+:)?c\b[^>]*(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
function parse_ws_xml_data(sdata, s, opts, refguess, _themes, styles, wb) {
  const dense = s["!data"] != null;
  const date1904 = wb?.WBProps?.date1904;
  sdata.match(rowregex) || [];
  const rows = sdata.split(/<\/(?:\w+:)?row>/);
  for (let ri = 0; ri < rows.length; ++ri) {
    const rowStr = rows[ri];
    if (!rowStr) continue;
    const rowTagMatch = rowStr.match(/<(?:\w+:)?row\b[^>]*>/);
    if (!rowTagMatch) continue;
    const rowTag = parsexmltag(rowTagMatch[0]);
    const R = parseInt(rowTag.r, 10) - 1;
    if (isNaN(R)) continue;
    if (rowTag.ht || rowTag.hidden) {
      if (!s["!rows"]) s["!rows"] = [];
      if (!s["!rows"][R]) s["!rows"][R] = {};
      if (rowTag.ht) s["!rows"][R].hpt = parseFloat(rowTag.ht);
      if (rowTag.hidden === "1") s["!rows"][R].hidden = true;
    }
    if (opts.sheetRows && R >= opts.sheetRows) continue;
    cellregex.lastIndex = 0;
    let cellMatch;
    while (cellMatch = cellregex.exec(rowStr)) {
      const cellTag = parsexmltag(cellMatch[0].match(/<(?:\w+:)?c\b[^>]*/)?.[0] + ">" || "");
      const ref = cellTag.r;
      if (!ref) continue;
      let C = 0;
      for (let ci = 0; ci < ref.length; ++ci) {
        const cc = ref.charCodeAt(ci);
        if (cc >= 65 && cc <= 90) C = 26 * C + (cc - 64);
        else break;
      }
      C -= 1;
      if (R < refguess.s.r) refguess.s.r = R;
      if (R > refguess.e.r) refguess.e.r = R;
      if (C < refguess.s.c) refguess.s.c = C;
      if (C > refguess.e.c) refguess.e.c = C;
      const cellType = cellTag.t || "n";
      cellTag.s ? parseInt(cellTag.s, 10) : 0;
      const cellValue = cellMatch[1] || "";
      let cell;
      const vMatch = cellValue.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
      const fMatch = cellValue.match(/<(?:\w+:)?f[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/);
      const isMatch = cellValue.match(/<(?:\w+:)?is>([\s\S]*?)<\/(?:\w+:)?is>/);
      const v = vMatch ? vMatch[1] : null;
      switch (cellType) {
        case "s":
          if (v !== null) {
            const idx = parseInt(v, 10);
            cell = { t: "s", v: "" };
            cell._sstIdx = idx;
          } else {
            cell = { t: "z" };
          }
          break;
        case "str":
          cell = { t: "s", v: v ? unescapexml(v) : "" };
          break;
        case "inlineStr":
          if (isMatch) {
            const tMatch = isMatch[1].match(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/);
            cell = { t: "s", v: tMatch ? unescapexml(tMatch[1]) : "" };
          } else {
            cell = { t: "s", v: "" };
          }
          break;
        case "b":
          cell = { t: "b", v: v === "1" };
          break;
        case "e":
          cell = { t: "e", v: v ? parseInt(v, 10) || 0 : 0 };
          cell.w = v || "";
          break;
        case "d":
          if (v) {
            cell = { t: "d", v: new Date(v) };
          } else {
            cell = { t: "z" };
          }
          break;
        default:
          if (v !== null) {
            cell = { t: "n", v: parseFloat(v) };
          } else {
            if (!opts.sheetStubs) continue;
            cell = { t: "z" };
          }
          break;
      }
      if (fMatch && opts.cellFormula !== false) {
        cell.f = unescapexml(fMatch[1]);
        const fTag = parsexmltag(cellValue.match(/<(?:\w+:)?f[^>]*/)?.[0] + ">" || "");
        if (fTag.t === "shared" && fTag.si != null) ;
        if (fTag.t === "array" && fTag.ref) {
          cell.F = fTag.ref;
          cell.D = fTag.dt === "1";
        }
      }
      if (opts.cellText !== false) {
        if (cell.t === "n") {
          const nfmt = cell.z || cell.XF && cell.XF.numFmtId != null && styles?.NumberFmt[cell.XF.numFmtId] || table_fmt[cell.XF && cell.XF.numFmtId || 0];
          if (nfmt) {
            try {
              cell.w = SSF_format(nfmt, cell.v, { date1904 });
            } catch {
            }
          }
          if (opts.cellDates && cell.XF) {
            const fmtStr = nfmt || table_fmt[cell.XF.numFmtId || 0] || "";
            if (typeof fmtStr === "string" && fmt_is_date(fmtStr) && typeof cell.v === "number") {
              cell.t = "d";
              cell.v = numdate(cell.v);
            }
          }
        }
      }
      if (dense) {
        if (!s["!data"][R]) s["!data"][R] = [];
        s["!data"][R][C] = cell;
      } else {
        s[ref] = cell;
      }
    }
  }
}
function resolve_sst(s, sst, opts) {
  const dense = s["!data"] != null;
  if (dense) {
    const data = s["!data"];
    for (let R = 0; R < data.length; ++R) {
      if (!data[R]) continue;
      for (let C = 0; C < data[R].length; ++C) {
        const cell = data[R][C];
        if (!cell || cell._sstIdx === void 0) continue;
        const idx = cell._sstIdx;
        delete cell._sstIdx;
        if (sst[idx]) {
          cell.v = sst[idx].t;
          if (opts.cellHTML !== false && sst[idx].h) cell.h = sst[idx].h;
          if (sst[idx].r) cell.r = sst[idx].r;
        }
      }
    }
  } else {
    for (const ref of Object.keys(s)) {
      if (ref.charAt(0) === "!") continue;
      const cell = s[ref];
      if (!cell || cell._sstIdx === void 0) continue;
      const idx = cell._sstIdx;
      delete cell._sstIdx;
      if (sst[idx]) {
        cell.v = sst[idx].t;
        if (opts.cellHTML !== false && sst[idx].h) cell.h = sst[idx].h;
        if (sst[idx].r) cell.r = sst[idx].r;
      }
    }
  }
}
function parse_ws_xml(data, opts, _idx, rels, wb, _themes, styles) {
  if (!data) return {};
  if (!opts) opts = {};
  if (!rels) rels = { "!id": {} };
  const s = opts.dense ? { "!data": [] } : {};
  const refguess = { s: { r: 2e6, c: 2e6 }, e: { r: 0, c: 0 } };
  let data1 = "";
  let data2 = "";
  const sdMatch = data.match(/<(?:\w+:)?sheetData[^>]*>([\s\S]*?)<\/(?:\w+:)?sheetData>/);
  if (sdMatch) {
    data1 = data.slice(0, sdMatch.index);
    data2 = data.slice(sdMatch.index + sdMatch[0].length);
  } else {
    data1 = data2 = data;
  }
  const ridx = (data1.match(/<(?:\w*:)?dimension/) || { index: -1 }).index;
  if (ridx > 0) {
    const ref = data1.slice(ridx, ridx + 50).match(dimregex);
    if (ref && !opts.nodim) parse_ws_xml_dim(s, ref[1]);
  }
  const columns = [];
  if (opts.cellStyles) {
    const cols = data1.match(colregex);
    if (cols) parse_ws_xml_cols(columns, cols);
  }
  if (sdMatch) parse_ws_xml_data(sdMatch[1], s, opts, refguess, _themes, styles, wb);
  const afilter = data2.match(afregex);
  if (afilter) s["!autofilter"] = parse_ws_xml_autofilter(afilter[0]);
  const merges = [];
  const _merge = data2.match(mergecregex);
  if (_merge) {
    for (let i = 0; i < _merge.length; ++i)
      merges[i] = safe_decode_range(_merge[i].slice(_merge[i].indexOf("=") + 2));
  }
  const hlink = data2.match(hlinkregex);
  if (hlink) parse_ws_xml_hlinks(s, hlink, rels);
  const margins = data2.match(marginregex);
  if (margins) s["!margins"] = parse_ws_xml_margins(parsexmltag(margins[0]));
  const legm = data2.match(/legacyDrawing r:id="(.*?)"/);
  if (legm) s["!legrel"] = legm[1];
  if (opts.nodim) {
    refguess.s.c = refguess.s.r = 0;
  }
  if (!s["!ref"] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) {
    s["!ref"] = encode_range(refguess);
  }
  if (opts.sheetRows > 0 && s["!ref"]) {
    const tmpref = safe_decode_range(s["!ref"]);
    if (opts.sheetRows <= tmpref.e.r) {
      tmpref.e.r = opts.sheetRows - 1;
      if (tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
      if (tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
      if (tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
      if (tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
      s["!fullref"] = s["!ref"];
      s["!ref"] = encode_range(tmpref);
    }
  }
  if (columns.length > 0) s["!cols"] = columns;
  if (merges.length > 0) s["!merges"] = merges;
  return s;
}
function write_ws_xml_merges(merges) {
  if (merges.length === 0) return "";
  const o = ['<mergeCells count="' + merges.length + '">'];
  for (let i = 0; i < merges.length; ++i)
    o.push('<mergeCell ref="' + encode_range(merges[i]) + '"/>');
  o.push("</mergeCells>");
  return o.join("");
}
function write_ws_xml(ws, opts, _idx, _rels, _wb) {
  const o = [XML_HEADER];
  o.push(
    writextag("worksheet", null, {
      xmlns: XMLNS_main[0],
      "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    })
  );
  const ref = ws["!ref"] || "A1";
  o.push('<dimension ref="' + ref + '"/>');
  o.push('<sheetViews><sheetView workbookViewId="0"');
  if (_idx === 0) o.push(' tabSelected="1"');
  o.push("/></sheetViews>");
  o.push('<sheetFormatPr defaultRowHeight="15"/>');
  if (ws["!cols"]) {
    o.push("<cols>");
    for (let i = 0; i < ws["!cols"].length; ++i) {
      if (!ws["!cols"][i]) continue;
      const col = ws["!cols"][i];
      const attrs = {
        min: String(i + 1),
        max: String(i + 1)
      };
      if (col.width) attrs.width = String(col.width);
      else attrs.width = "9.140625";
      if (col.hidden) attrs.hidden = "1";
      attrs.customWidth = "1";
      o.push(writextag("col", null, attrs));
    }
    o.push("</cols>");
  }
  o.push("<sheetData>");
  const dense = ws["!data"] != null;
  const range = safe_decode_range(ref);
  for (let R = range.s.r; R <= range.e.r; ++R) {
    const row_cells = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let cell;
      if (dense) {
        cell = ws["!data"]?.[R]?.[C];
      } else {
        const addr2 = encode_cell({ r: R, c: C });
        cell = ws[addr2];
      }
      if (!cell || cell.t === "z") continue;
      const addr = encode_cell({ r: R, c: C });
      let v = "";
      let t = "";
      switch (cell.t) {
        case "b":
          v = cell.v ? "1" : "0";
          t = "b";
          break;
        case "n":
          v = String(cell.v);
          break;
        case "e":
          v = String(cell.v);
          t = "e";
          break;
        case "d":
          if (opts.cellDates) {
            v = cell.v.toISOString();
            t = "d";
          } else {
            v = String(datenum(cell.v));
          }
          break;
        case "s":
          v = escapexml(String(cell.v));
          t = "str";
          break;
      }
      let cellXml = '<c r="' + addr + '"';
      if (t) cellXml += ' t="' + t + '"';
      cellXml += ">";
      if (cell.f) {
        cellXml += "<f";
        if (cell.F) cellXml += ' ref="' + cell.F + '" t="array"';
        cellXml += ">" + escapexml(cell.f) + "</f>";
      }
      if (v !== "") cellXml += "<v>" + v + "</v>";
      cellXml += "</c>";
      row_cells.push(cellXml);
    }
    if (row_cells.length > 0) {
      let rowTag = '<row r="' + (R + 1) + '"';
      if (ws["!rows"]?.[R]) {
        if (ws["!rows"][R].hpt) rowTag += ' ht="' + ws["!rows"][R].hpt + '" customHeight="1"';
        if (ws["!rows"][R].hidden) rowTag += ' hidden="1"';
      }
      rowTag += ">";
      o.push(rowTag);
      o.push(row_cells.join(""));
      o.push("</row>");
    }
  }
  o.push("</sheetData>");
  if (ws["!merges"] && ws["!merges"].length > 0) {
    o.push(write_ws_xml_merges(ws["!merges"]));
  }
  if (ws["!autofilter"]) {
    o.push('<autoFilter ref="' + ws["!autofilter"].ref + '"/>');
  }
  if (ws["!margins"]) {
    const m = ws["!margins"];
    o.push(
      writextag("pageMargins", null, {
        left: String(m.left || 0.7),
        right: String(m.right || 0.7),
        top: String(m.top || 0.75),
        bottom: String(m.bottom || 0.75),
        header: String(m.header || 0.3),
        footer: String(m.footer || 0.3)
      })
    );
  }
  o.push("</worksheet>");
  o[1] = o[1].replace("/>", ">");
  return o.join("");
}

// src/xlsx/comments.ts
function sheet_insert_comments(sheet, comments, threaded, people) {
  const dense = sheet["!data"] != null;
  for (const comment of comments) {
    const r = decode_cell(comment.ref);
    if (r.r < 0 || r.c < 0) continue;
    let cell;
    if (dense) {
      if (!sheet["!data"][r.r]) sheet["!data"][r.r] = [];
      cell = sheet["!data"][r.r][r.c];
    } else {
      cell = sheet[comment.ref];
    }
    if (!cell) {
      cell = { t: "z" };
      if (dense) sheet["!data"][r.r][r.c] = cell;
      else sheet[comment.ref] = cell;
      const range = safe_decode_range(sheet["!ref"] || "BDWGO1000001:A1");
      if (range.s.r > r.r) range.s.r = r.r;
      if (range.e.r < r.r) range.e.r = r.r;
      if (range.s.c > r.c) range.s.c = r.c;
      if (range.e.c < r.c) range.e.c = r.c;
      sheet["!ref"] = encode_range(range);
    }
    if (!cell.c) cell.c = [];
    const o = { a: comment.author, t: comment.t, r: comment.r, T: threaded };
    if (comment.h) o.h = comment.h;
    for (let i = cell.c.length - 1; i >= 0; --i) {
      if (!threaded && cell.c[i].T) return;
      if (threaded && !cell.c[i].T) cell.c.splice(i, 1);
    }
    if (threaded && people) {
      for (let i = 0; i < people.length; ++i) {
        if (o.a === people[i].id) {
          o.a = people[i].name || o.a;
          break;
        }
      }
    }
    cell.c.push(o);
  }
}
function parse_si_simple(x) {
  if (!x) return { t: "", r: "", h: "" };
  const tMatch = x.match(/<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/);
  const t = tMatch ? unescapexml(tMatch[1]) : "";
  return { t, r: x, h: t };
}
function parse_comments_xml(data, opts) {
  if (data.match(/<(?:\w+:)?comments\s*\/>/)) return [];
  const authors = [];
  const commentList = [];
  const authtag = str_match_xml_ns(data, "authors");
  if (authtag) {
    authtag.split(/<\/\w*:?author>/).forEach((x) => {
      if (x === "" || x.trim() === "") return;
      const a = x.match(/<(?:\w+:)?author[^<>]*>(.*)/);
      if (a) authors.push(a[1]);
    });
  }
  const cmnttag = str_match_xml_ns(data, "commentList");
  if (cmnttag) {
    cmnttag.split(/<\/\w*:?comment>/).forEach((x) => {
      if (x === "" || x.trim() === "") return;
      const cm = x.match(/<(?:\w+:)?comment[^<>]*>/);
      if (!cm) return;
      const y = parsexmltag(cm[0]);
      const comment = {
        author: y.authorId && authors[y.authorId] || "sheetjsghost",
        ref: y.ref,
        guid: y.guid,
        t: ""
      };
      const cell = decode_cell(y.ref);
      if (opts && opts.sheetRows && opts.sheetRows <= cell.r) return;
      const textMatch = str_match_xml_ns(x, "text");
      const rt = textMatch ? parse_si_simple(textMatch) : { r: "", t: "", h: "" };
      comment.r = rt.r;
      if (rt.r === "<t></t>") {
        rt.t = "";
        rt.h = "";
      }
      comment.t = (rt.t || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      if (opts && opts.cellHTML) comment.h = rt.h;
      commentList.push(comment);
    });
  }
  return commentList;
}
function write_comments_xml(data) {
  const o = [XML_HEADER, writextag("comments", null, { xmlns: XMLNS_main[0] })];
  const iauthor = [];
  o.push("<authors>");
  data.forEach((x) => {
    x[1].forEach((w) => {
      const a = escapexml(w.a);
      if (iauthor.indexOf(a) === -1) {
        iauthor.push(a);
        o.push("<author>" + a + "</author>");
      }
      if (w.T && w.ID && iauthor.indexOf("tc=" + w.ID) === -1) {
        iauthor.push("tc=" + w.ID);
        o.push("<author>tc=" + w.ID + "</author>");
      }
    });
  });
  if (iauthor.length === 0) {
    iauthor.push("SheetJ5");
    o.push("<author>SheetJ5</author>");
  }
  o.push("</authors>");
  o.push("<commentList>");
  data.forEach((d) => {
    let lastauthor = 0;
    const ts = [];
    let tcnt = 0;
    if (d[1][0] && d[1][0].T && d[1][0].ID) lastauthor = iauthor.indexOf("tc=" + d[1][0].ID);
    d[1].forEach((c) => {
      if (c.a) lastauthor = iauthor.indexOf(escapexml(c.a));
      if (c.T) ++tcnt;
      ts.push(c.t == null ? "" : escapexml(c.t));
    });
    if (tcnt === 0) {
      d[1].forEach((c) => {
        o.push(
          '<comment ref="' + d[0] + '" authorId="' + iauthor.indexOf(escapexml(c.a)) + '"><text>'
        );
        o.push(writetag("t", c.t == null ? "" : escapexml(c.t)));
        o.push("</text></comment>");
      });
    } else {
      if (d[1][0] && d[1][0].T && d[1][0].ID)
        lastauthor = iauthor.indexOf("tc=" + d[1][0].ID);
      o.push('<comment ref="' + d[0] + '" authorId="' + lastauthor + '"><text>');
      let t = "Comment:\n    " + ts[0] + "\n";
      for (let i = 1; i < ts.length; ++i) t += "Reply:\n    " + ts[i] + "\n";
      o.push(writetag("t", escapexml(t)));
      o.push("</text></comment>");
    }
  });
  o.push("</commentList>");
  if (o.length > 2) {
    o.push("</comments>");
    o[1] = o[1].replace("/>", ">");
  }
  return o.join("");
}
function parse_tcmnt_xml(data, opts) {
  const out = [];
  let comment = {};
  let tidx = 0;
  data.replace(tagregex, function xml_tcmnt(x, idx) {
    const y = parsexmltag(x);
    switch (strip_ns(y[0])) {
      case "<?xml":
        break;
      case "<ThreadedComments":
      case "</ThreadedComments>":
        break;
      case "<threadedComment":
        comment = { author: y.personId, guid: y.id, ref: y.ref, T: 1 };
        break;
      case "</threadedComment>":
        if (comment.t != null) out.push(comment);
        break;
      case "<text>":
      case "<text":
        tidx = idx + x.length;
        break;
      case "</text>":
        comment.t = data.slice(tidx, idx).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
        break;
    }
    return x;
  });
  return out;
}
function write_tcmnt_xml(comments, people, opts) {
  const o = [
    XML_HEADER,
    writextag("ThreadedComments", null, { xmlns: XMLNS.TCMNT }).replace(/[/]>/, ">")
  ];
  comments.forEach((carr) => {
    let rootid = "";
    (carr[1] || []).forEach((c, idx) => {
      if (!c.T) {
        delete c.ID;
        return;
      }
      if (c.a && people.indexOf(c.a) === -1) people.push(c.a);
      const tcopts = {
        ref: carr[0],
        id: "{54EE7951-7262-4200-6969-" + ("000000000000" + opts.tcid++).slice(-12) + "}"
      };
      if (idx === 0) rootid = tcopts.id;
      else tcopts.parentId = rootid;
      c.ID = tcopts.id;
      if (c.a)
        tcopts.personId = "{54EE7950-7262-4200-6969-" + ("000000000000" + people.indexOf(c.a)).slice(-12) + "}";
      o.push(writextag("threadedComment", writetag("text", c.t || ""), tcopts));
    });
  });
  o.push("</ThreadedComments>");
  return o.join("");
}
function parse_people_xml(data) {
  const out = [];
  data.replace(tagregex, function xml_people(x) {
    const y = parsexmltag(x);
    switch (strip_ns(y[0])) {
      case "<?xml":
        break;
      case "<personList":
      case "</personList>":
        break;
      case "<person":
        out.push({ name: y.displayname, id: y.id });
        break;
    }
    return x;
  });
  return out;
}
function write_people_xml(people) {
  const o = [
    XML_HEADER,
    writextag("personList", null, {
      xmlns: XMLNS.TCMNT,
      "xmlns:x": XMLNS_main[0]
    }).replace(/[/]>/, ">")
  ];
  people.forEach((person, idx) => {
    o.push(
      writextag("person", null, {
        displayName: person,
        id: "{54EE7950-7262-4200-6969-" + ("000000000000" + idx).slice(-12) + "}",
        userId: person,
        providerId: "None"
      })
    );
  });
  o.push("</personList>");
  return o.join("");
}

// src/xlsx/vml.ts
var XLMLNS = {
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  x: "urn:schemas-microsoft-com:office:excel",
  mv: "http://macVmlSchemaUri"
};
function parse_vml(data, sheet, comments) {
  let cidx = 0;
  (str_match_xml_ns_g(data, "(?:shape|rect)") || []).forEach((m) => {
    let type = "";
    let hidden = true;
    let aidx = -1;
    let R = -1, C = -1;
    m.replace(tagregex, function(x, idx) {
      const y = parsexmltag(x);
      switch (strip_ns(y[0])) {
        case "<ClientData":
          if (y.ObjectType) type = y.ObjectType;
          break;
        case "<Visible":
        case "<Visible/>":
          hidden = false;
          break;
        case "<Row":
        case "<Row>":
          aidx = idx + x.length;
          break;
        case "</Row>":
          R = +m.slice(aidx, idx).trim();
          break;
        case "<Column":
        case "<Column>":
          aidx = idx + x.length;
          break;
        case "</Column>":
          C = +m.slice(aidx, idx).trim();
          break;
      }
      return "";
    });
    switch (type) {
      case "Note": {
        const ref = R >= 0 && C >= 0 ? encode_cell({ r: R, c: C }) : comments[cidx]?.ref;
        const dense = sheet["!data"] != null;
        let cell;
        if (dense) {
          const rows = sheet["!data"];
          cell = rows?.[R]?.[C];
        } else {
          cell = sheet[ref];
        }
        if (cell && cell.c) {
          cell.c.hidden = hidden;
        }
        ++cidx;
        break;
      }
    }
  });
}
function wxt_helper2(h) {
  return keys(h).map((k) => " " + k + '="' + h[k] + '"').join("");
}
function write_vml_comment(x, _shapeid) {
  const c = decode_cell(x[0]);
  const fillopts = { color2: "#BEFF82", type: "gradient" };
  if (fillopts.type === "gradient") fillopts.angle = "-180";
  const fillparm = fillopts.type === "gradient" ? writextag("o:fill", null, { type: "gradientUnscaled", "v:ext": "view" }) : null;
  const fillxml = writextag("v:fill", fillparm, fillopts);
  const shadata = { on: "t", obscured: "t" };
  return [
    "<v:shape" + wxt_helper2({
      id: "_x0000_s" + _shapeid,
      type: "#_x0000_t202",
      style: "position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10" + (x[1].hidden ? ";visibility:hidden" : ""),
      fillcolor: "#ECFAD4",
      strokecolor: "#edeaa1"
    }) + ">",
    fillxml,
    writextag("v:shadow", null, shadata),
    writextag("v:path", null, { "o:connecttype": "none" }),
    '<v:textbox><div style="text-align:left"></div></v:textbox>',
    '<x:ClientData ObjectType="Note">',
    "<x:MoveWithCells/>",
    "<x:SizeWithCells/>",
    writetag(
      "x:Anchor",
      [c.c + 1, 0, c.r + 1, 0, c.c + 3, 20, c.r + 5, 20].join(",")
    ),
    writetag("x:AutoFill", "False"),
    writetag("x:Row", String(c.r)),
    writetag("x:Column", String(c.c)),
    x[1].hidden ? "" : "<x:Visible/>",
    "</x:ClientData>",
    "</v:shape>"
  ].join("");
}
function write_vml(rId, comments) {
  const csize = [21600, 21600];
  const bbox = ["m0,0l0", csize[1], csize[0], csize[1], csize[0], "0xe"].join(",");
  const o = [
    writextag("xml", null, {
      "xmlns:v": XLMLNS.v,
      "xmlns:o": XLMLNS.o,
      "xmlns:x": XLMLNS.x,
      "xmlns:mv": XLMLNS.mv
    }).replace(/\/>/, ">"),
    writextag(
      "o:shapelayout",
      writextag("o:idmap", null, { "v:ext": "edit", data: String(rId) }),
      { "v:ext": "edit" }
    )
  ];
  let _shapeid = 65536 * rId;
  const _comments = comments || [];
  if (_comments.length > 0)
    o.push(
      writextag(
        "v:shapetype",
        [
          writextag("v:stroke", null, { joinstyle: "miter" }),
          writextag("v:path", null, { gradientshapeok: "t", "o:connecttype": "rect" })
        ].join(""),
        {
          id: "_x0000_t202",
          coordsize: csize.join(","),
          "o:spt": "202",
          path: bbox
        }
      )
    );
  _comments.forEach((x) => {
    ++_shapeid;
    o.push(write_vml_comment(x, _shapeid));
  });
  o.push("</xml>");
  return o.join("");
}

// src/xlsx/metadata.ts
function parse_xlmeta_xml(data, opts) {
  const out = { Types: [], Cell: [], Value: [] };
  if (!data) return out;
  let metatype = 2;
  let lastmeta;
  data.replace(tagregex, function(x) {
    const y = parsexmltag(x);
    switch (strip_ns(y[0])) {
      case "<?xml":
        break;
      case "<metadata":
      case "</metadata>":
        break;
      case "<metadataTypes":
      case "</metadataTypes>":
        break;
      case "<metadataType":
        out.Types.push({ name: y.name });
        break;
      case "</metadataType>":
        break;
      case "<futureMetadata":
        for (let j = 0; j < out.Types.length; ++j)
          if (out.Types[j].name === y.name) lastmeta = out.Types[j];
        break;
      case "</futureMetadata>":
        break;
      case "<bk>":
      case "</bk>":
        break;
      case "<rc":
        if (metatype === 1)
          out.Cell.push({ type: out.Types[y.t - 1].name, index: +y.v });
        else if (metatype === 0)
          out.Value.push({ type: out.Types[y.t - 1].name, index: +y.v });
        break;
      case "</rc>":
        break;
      case "<cellMetadata":
        metatype = 1;
        break;
      case "</cellMetadata>":
        metatype = 2;
        break;
      case "<valueMetadata":
        metatype = 0;
        break;
      case "</valueMetadata>":
        metatype = 2;
        break;
      case "<extLst":
      case "<extLst>":
      case "</extLst>":
      case "<extLst/>":
        break;
      case "<ext":
        break;
      case "</ext>":
        break;
      case "<rvb":
        if (!lastmeta) break;
        if (!lastmeta.offsets) lastmeta.offsets = [];
        lastmeta.offsets.push(+y.i);
        break;
    }
    return x;
  });
  return out;
}
function write_xlmeta_xml() {
  const o = [XML_HEADER];
  o.push(
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">\n  <metadataTypes count="1">\n    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>\n  </metadataTypes>\n  <futureMetadata name="XLDAPR" count="1">\n    <bk>\n      <extLst>\n        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">\n          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>\n        </ext>\n      </extLst>\n    </bk>\n  </futureMetadata>\n  <cellMetadata count="1">\n    <bk>\n      <rc t="1" v="0"/>\n    </bk>\n  </cellMetadata>\n</metadata>'
  );
  return o.join("");
}

// src/xlsx/calc-chain.ts
function parse_cc_xml(data) {
  const d = [];
  if (!data) return d;
  let i = 1;
  (data.match(tagregex) || []).forEach((x) => {
    const y = parsexmltag(x);
    switch (y[0]) {
      case "<?xml":
        break;
      case "<calcChain":
      case "<calcChain>":
      case "</calcChain>":
        break;
      case "<c":
        delete y[0];
        if (y.i) i = y.i;
        else y.i = i;
        d.push(y);
        break;
    }
  });
  return d;
}

// src/xlsx/parse-zip.ts
function strip_front_slash(x) {
  return x.charAt(0) === "/" ? x.slice(1) : x;
}
function resolve_path2(target, basePath) {
  if (target.charAt(0) === "/") return target;
  const base = basePath.slice(0, basePath.lastIndexOf("/") + 1);
  const parts = (base + target).split("/");
  const resolved = [];
  for (const p of parts) {
    if (p === "..") resolved.pop();
    else if (p !== ".") resolved.push(p);
  }
  return resolved.join("/");
}
function getzipstr(zip, path, safe) {
  const p = zip_read_str(zip, path);
  if (p == null && !safe) throw new Error("Could not find " + path);
  return p;
}
function getzipdata(zip, path, safe) {
  return getzipstr(zip, path, safe);
}
var RELS_WS = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
];
function get_sheet_type(n) {
  if (RELS_WS.indexOf(n) > -1) return "sheet";
  return n && n.length ? n : "sheet";
}
function safe_parse_wbrels(wbrels, sheets) {
  if (!wbrels) return null;
  try {
    const result = sheets.map((w) => {
      const id = w.id || w.strRelID;
      return [w.name, wbrels["!id"][id].Target, get_sheet_type(wbrels["!id"][id].Type)];
    });
    return result.length === 0 ? null : result;
  } catch {
    return null;
  }
}
function safe_parse_sheet(zip, path, relsPath, sheetName, idx, sheetRels, sheets, stype, opts, wb, themes, styles, strs) {
  try {
    sheetRels[sheetName] = parse_rels(getzipstr(zip, relsPath, true), path);
    const data = getzipdata(zip, path);
    if (!data) return;
    let _ws;
    switch (stype) {
      case "sheet":
        _ws = parse_ws_xml(data, opts);
        break;
      default:
        return;
    }
    if (!_ws) return;
    resolve_sst(_ws, strs, opts);
    sheets[sheetName] = _ws;
    const comments = [];
    let tcomments = [];
    if (sheetRels[sheetName]) {
      for (const n of keys(sheetRels[sheetName])) {
        if (n === "!id" || n === "!idx") continue;
        const rel = sheetRels[sheetName][n];
        if (!rel || !rel.Type) continue;
        if (rel.Type === RELS.CMNT) {
          const dfile = resolve_path2(rel.Target, path);
          const cmntData = getzipdata(zip, dfile, true);
          if (cmntData) {
            const parsedComments = parse_comments_xml(cmntData, opts);
            if (parsedComments && parsedComments.length > 0) {
              sheet_insert_comments(_ws, parsedComments, false);
            }
          }
        }
        if (rel.Type === RELS.TCMNT) {
          const dfile = resolve_path2(rel.Target, path);
          const tcData = getzipdata(zip, dfile, true);
          if (tcData) {
            tcomments = tcomments.concat(parse_tcmnt_xml(tcData, opts));
          }
        }
      }
    }
    if (tcomments.length > 0) {
      sheet_insert_comments(_ws, tcomments, true, opts.people || []);
    }
    if (_ws["!legdrawel"] && sheetRels[sheetName]) {
      const dfile = resolve_path2(_ws["!legdrawel"].Target, path);
      const draw = getzipstr(zip, dfile, true);
      if (draw) parse_vml(utf8read(draw), _ws, comments);
    }
  } catch (e) {
    if (opts.WTF) throw e;
  }
}
function parse_zip(zip, opts) {
  make_ssf();
  const o = opts || {};
  if (!zip_has(zip, "[Content_Types].xml")) {
    throw new Error("Unsupported ZIP file");
  }
  const dir = parse_ct(getzipstr(zip, "[Content_Types].xml"));
  if (dir.workbooks.length === 0) {
    const binname = "xl/workbook.xml";
    if (getzipdata(zip, binname, true)) dir.workbooks.push(binname);
  }
  if (dir.workbooks.length === 0) {
    throw new Error("Could not find workbook");
  }
  const themes = { themeElements: { clrScheme: [] } };
  let styles = { };
  let strs = [];
  if (!o.bookSheets && !o.bookProps) {
    if (dir.sst) {
      try {
        const sstData = getzipdata(zip, strip_front_slash(dir.sst));
        if (sstData) strs = parse_sst_xml(sstData, o);
      } catch (e) {
        if (o.WTF) throw e;
      }
    }
    if (dir.themes.length) {
      const themeData = getzipstr(zip, dir.themes[0].replace(/^\//, ""), true);
      if (themeData) {
        const parsed = parse_theme_xml(themeData);
        Object.assign(themes, parsed);
      }
    }
    if (dir.style) {
      const styData = getzipdata(zip, strip_front_slash(dir.style));
      if (styData) styles = parse_sty_xml(styData);
    }
  }
  const wb = parse_wb_xml(
    getzipdata(zip, strip_front_slash(dir.workbooks[0])));
  const props = {};
  if (dir.coreprops.length) {
    const propdata = getzipdata(zip, strip_front_slash(dir.coreprops[0]), true);
    if (propdata) Object.assign(props, parse_core_props(propdata));
    if (dir.extprops.length) {
      const extdata = getzipdata(zip, strip_front_slash(dir.extprops[0]), true);
      if (extdata) parse_ext_props(extdata, props);
    }
  }
  let custprops = {};
  if (!o.bookSheets || o.bookProps) {
    if (dir.custprops.length) {
      const custdata = getzipstr(zip, strip_front_slash(dir.custprops[0]), true);
      if (custdata) custprops = parse_cust_props(custdata, o);
    }
  }
  const out = {};
  if (o.bookSheets || o.bookProps) {
    let sheets2;
    if (wb.Sheets) sheets2 = wb.Sheets.map((x) => x.name);
    else if (props.Worksheets && props.SheetNames?.length > 0) sheets2 = props.SheetNames;
    if (o.bookProps) {
      out.Props = props;
      out.Custprops = custprops;
    }
    if (o.bookSheets && sheets2) out.SheetNames = sheets2;
    if (o.bookSheets ? out.SheetNames : o.bookProps) return out;
  }
  const sheets = {};
  if (o.bookDeps && dir.calcchain) {
    parse_cc_xml(getzipdata(zip, strip_front_slash(dir.calcchain), true) || "");
  }
  const sheetRels = {};
  const wbsheets = wb.Sheets;
  props.Worksheets = wbsheets.length;
  props.SheetNames = [];
  for (let j = 0; j < wbsheets.length; ++j) {
    props.SheetNames[j] = wbsheets[j].name;
  }
  const wbrelsi = dir.workbooks[0].lastIndexOf("/");
  let wbrelsfile = (dir.workbooks[0].slice(0, wbrelsi + 1) + "_rels/" + dir.workbooks[0].slice(wbrelsi + 1) + ".rels").replace(/^\//, "");
  if (!zip_has(zip, wbrelsfile)) wbrelsfile = "xl/_rels/workbook.xml.rels";
  let wbrels = parse_rels(
    getzipstr(zip, wbrelsfile, true),
    wbrelsfile.replace(/_rels.*/, "s5s")
  );
  if ((dir.metadata || []).length >= 1) {
    o.xlmeta = parse_xlmeta_xml(
      getzipdata(zip, strip_front_slash(dir.metadata[0]), true) || "");
  }
  if ((dir.people || []).length >= 1) {
    o.people = parse_people_xml(
      getzipdata(zip, strip_front_slash(dir.people[0]), true) || ""
    );
  }
  const wbrelsArr = wbrels ? safe_parse_wbrels(wbrels, wb.Sheets) : null;
  const nmode = getzipdata(zip, "xl/worksheets/sheet.xml", true) ? 1 : 0;
  for (let i = 0; i < props.Worksheets; ++i) {
    let stype = "sheet";
    let path;
    if (wbrelsArr && wbrelsArr[i]) {
      path = "xl/" + wbrelsArr[i][1].replace(/[/]?xl\//, "");
      if (!zip_has(zip, path)) path = wbrelsArr[i][1];
      if (!zip_has(zip, path))
        path = wbrelsfile.replace(/_rels\/[\S\s]*$/, "") + wbrelsArr[i][1];
      stype = wbrelsArr[i][2];
    } else {
      path = "xl/worksheets/sheet" + (i + 1 - nmode) + ".xml";
      path = path.replace(/sheet0\./, "sheet.");
    }
    if (o.sheets != null) {
      if (typeof o.sheets === "number" && i !== o.sheets) continue;
      if (typeof o.sheets === "string" && props.SheetNames[i].toLowerCase() !== o.sheets.toLowerCase()) continue;
      if (Array.isArray(o.sheets)) {
        let seen = false;
        for (const s of o.sheets) {
          if (typeof s === "number" && s === i) seen = true;
          if (typeof s === "string" && s.toLowerCase() === props.SheetNames[i].toLowerCase()) seen = true;
        }
        if (!seen) continue;
      }
    }
    const relsPath = path.replace(/^(.*)(\/)([^/]*)$/, "$1/_rels/$3.rels");
    safe_parse_sheet(zip, path, relsPath, props.SheetNames[i], i, sheetRels, sheets, stype, o, wb, themes, styles, strs);
  }
  const result = {
    Sheets: sheets,
    SheetNames: props.SheetNames,
    Props: props,
    Custprops: custprops,
    bookType: "xlsx"
  };
  if (wb.WBProps) {
    result.Workbook = {
      WBProps: wb.WBProps,
      Sheets: wb.Sheets,
      Names: wb.Names
    };
  }
  return result;
}

// src/utils/base64.ts
function base64decode(input) {
  let str = input;
  if (str.slice(0, 5) === "data:") {
    const i = str.slice(0, 1024).indexOf(";base64,");
    if (i > -1) str = str.slice(i + 8);
  }
  const binaryStr = atob(str);
  const len = binaryStr.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryStr.charCodeAt(i);
  }
  return bytes;
}
function base64encode(data) {
  let binaryStr = "";
  for (let i = 0; i < data.length; i++) {
    binaryStr += String.fromCharCode(data[i]);
  }
  return btoa(binaryStr);
}
function to_uint8array(data, opts) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(data)) return new Uint8Array(data.buffer, data.byteOffset, data.length);
  if (typeof data === "string") {
    if (opts.type === "base64") {
      return base64decode(data);
    }
    const u8 = new Uint8Array(data.length);
    for (let i = 0; i < data.length; ++i) u8[i] = data.charCodeAt(i);
    return u8;
  }
  if (Array.isArray(data)) return new Uint8Array(data);
  throw new Error("Unsupported data type for read()");
}
function detect_type(data) {
  if (data instanceof Uint8Array || data instanceof ArrayBuffer) return "array";
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(data)) return "buffer";
  if (typeof data === "string") return "base64";
  return "array";
}
function read(data, opts) {
  make_ssf();
  const o = opts ? dup(opts) : {};
  if (!o.type) o.type = detect_type(data);
  const u8 = to_uint8array(data, o);
  if (u8[0] === 80 && u8[1] === 75) {
    const zip = zip_read(u8);
    return parse_zip(zip, o);
  }
  if (u8[0] === 37 && u8[1] === 80 && u8[2] === 68 && u8[3] === 70) {
    throw new Error("PDF File is not a spreadsheet");
  }
  if (u8[0] === 137 && u8[1] === 80 && u8[2] === 78 && u8[3] === 71) {
    throw new Error("PNG Image File is not a spreadsheet");
  }
  throw new Error("Unsupported file format. xlsx-format only supports XLSX files.");
}
function readFile(filename, opts) {
  const data = fs.readFileSync(filename);
  return read(new Uint8Array(data), opts);
}

// src/xlsx/write-zip.ts
function write_zip_xlsx(wb, opts) {
  if (wb && !wb.SSF) {
    wb.SSF = dup(table_fmt);
  }
  if (wb && wb.SSF) {
    make_ssf();
    SSF_load_table(wb.SSF);
  }
  opts.rels = { "!id": {} };
  opts.wbrels = { "!id": {} };
  opts.Strings = [];
  opts.Strings.Count = 0;
  opts.Strings.Unique = 0;
  opts.revStrings = /* @__PURE__ */ new Map();
  const ct = new_ct();
  const zip = zip_new();
  let f = "";
  opts.cellXfs = [];
  if (!wb.Props) wb.Props = {};
  f = "docProps/core.xml";
  zip_add_str(zip, f, write_core_props(wb.Props, opts));
  ct.coreprops.push(f);
  add_rels(opts.rels, 2, f, RELS.CORE_PROPS);
  f = "docProps/app.xml";
  if (wb.Props && wb.Props.SheetNames) ; else if (!wb.Workbook || !wb.Workbook.Sheets) {
    wb.Props.SheetNames = wb.SheetNames;
  } else {
    const _sn = [];
    for (let _i = 0; _i < wb.SheetNames.length; ++_i) {
      if ((wb.Workbook.Sheets[_i] || {}).Hidden !== 2) _sn.push(wb.SheetNames[_i]);
    }
    wb.Props.SheetNames = _sn;
  }
  wb.Props.Worksheets = wb.Props.SheetNames.length;
  zip_add_str(zip, f, write_ext_props(wb.Props));
  ct.extprops.push(f);
  add_rels(opts.rels, 3, f, RELS.EXT_PROPS);
  if (wb.Custprops !== wb.Props && keys(wb.Custprops || {}).length > 0) {
    f = "docProps/custom.xml";
    zip_add_str(zip, f, write_cust_props(wb.Custprops));
    ct.custprops.push(f);
    add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
  }
  const people = ["SheetJ5"];
  opts.tcid = 0;
  for (let rId = 1; rId <= wb.SheetNames.length; ++rId) {
    const wsrels = { "!id": {} };
    const ws = wb.Sheets[wb.SheetNames[rId - 1]];
    f = "xl/worksheets/sheet" + rId + ".xml";
    zip_add_str(zip, f, write_ws_xml(ws || {}, opts, rId - 1));
    ct.sheets.push(f);
    add_rels(opts.wbrels, -1, "worksheets/sheet" + rId + ".xml", RELS.SHEET);
    if (ws) {
      const comments = ws["!comments"];
      let need_vml = false;
      if (comments && comments.length > 0) {
        let needtc = false;
        comments.forEach((carr) => {
          carr[1].forEach((c) => {
            if (c.T === true) needtc = true;
          });
        });
        if (needtc) {
          const cf = "xl/threadedComments/threadedComment" + rId + ".xml";
          zip_add_str(zip, cf, write_tcmnt_xml(comments, people, opts));
          ct.threadedcomments.push(cf);
          add_rels(wsrels, -1, "../threadedComments/threadedComment" + rId + ".xml", RELS.TCMNT);
        }
        const cf2 = "xl/comments" + rId + ".xml";
        zip_add_str(zip, cf2, write_comments_xml(comments));
        ct.comments.push(cf2);
        add_rels(wsrels, -1, "../comments" + rId + ".xml", RELS.CMNT);
        need_vml = true;
      }
      if (ws["!legacy"]) {
        if (need_vml) {
          zip_add_str(
            zip,
            "xl/drawings/vmlDrawing" + rId + ".vml",
            write_vml(rId, ws["!comments"])
          );
        }
      }
      delete ws["!comments"];
      delete ws["!legacy"];
    }
    if (wsrels["!id"].rId1) {
      zip_add_str(zip, get_rels_path(f), write_rels(wsrels));
    }
  }
  if (opts.Strings != null && opts.Strings.length > 0) {
    f = "xl/sharedStrings.xml";
    zip_add_str(zip, f, write_sst_xml(opts.Strings, opts));
    ct.strs.push(f);
    add_rels(opts.wbrels, -1, "sharedStrings.xml", RELS.SST);
  }
  f = "xl/workbook.xml";
  zip_add_str(zip, f, write_wb_xml(wb));
  ct.workbooks.push(f);
  add_rels(opts.rels, 1, f, RELS.WB);
  f = "xl/theme/theme1.xml";
  zip_add_str(zip, f, write_theme_xml());
  ct.themes.push(f);
  add_rels(opts.wbrels, -1, "theme/theme1.xml", RELS.THEME);
  f = "xl/styles.xml";
  zip_add_str(zip, f, write_sty_xml());
  ct.styles.push(f);
  add_rels(opts.wbrels, -1, "styles.xml", RELS.STY);
  f = "xl/metadata.xml";
  zip_add_str(zip, f, write_xlmeta_xml());
  ct.metadata.push(f);
  add_rels(opts.wbrels, -1, "metadata.xml", RELS.META);
  if (people.length > 1) {
    f = "xl/persons/person.xml";
    zip_add_str(zip, f, write_people_xml(people));
    ct.people.push(f);
    add_rels(opts.wbrels, -1, "persons/person.xml", RELS.PEOPLE);
  }
  zip_add_str(zip, "[Content_Types].xml", write_ct(ct, opts));
  zip_add_str(zip, "_rels/.rels", write_rels(opts.rels));
  zip_add_str(zip, "xl/_rels/workbook.xml.rels", write_rels(opts.wbrels));
  return zip;
}
function write(wb, opts) {
  make_ssf();
  if (!opts || !opts.unsafe) check_wb(wb);
  const o = dup(opts || {});
  if (o.cellStyles) {
    o.cellNF = true;
    o.sheetStubs = true;
  }
  const zip = write_zip_xlsx(wb, o);
  const compressed = zip_write(zip, !!o.compression);
  switch (o.type) {
    case "base64":
      return base64encode_u8(compressed);
    case "buffer":
      if (typeof Buffer !== "undefined") return Buffer.from(compressed.buffer, compressed.byteOffset, compressed.byteLength);
      return compressed;
    case "array":
      return compressed;
    default:
      return compressed;
  }
}
function base64encode_u8(data) {
  return base64encode(data);
}
function writeFile(wb, filename, opts) {
  const o = opts ? dup(opts) : {};
  o.type = "buffer";
  const data = write(wb, o);
  fs.writeFileSync(filename, data instanceof Uint8Array ? Buffer.from(data) : data);
}

// src/api/book.ts
function book_new(ws, wsname) {
  const wb = { SheetNames: [], Sheets: {} };
  if (ws) book_append_sheet(wb, ws, wsname || "Sheet1");
  return wb;
}
function book_append_sheet(wb, ws, name, roll) {
  let i = 1;
  if (!name) {
    for (; i <= 65535; ++i, name = void 0) {
      if (wb.SheetNames.indexOf(name = "Sheet" + i) === -1) break;
    }
  }
  if (!name || wb.SheetNames.length >= 65535) throw new Error("Too many worksheets");
  if (roll && wb.SheetNames.indexOf(name) >= 0 && name.length < 32) {
    const m = name.match(/\d+$/);
    i = m && +m[0] || 0;
    const root = m && name.slice(0, m.index) || name;
    for (++i; i <= 65535; ++i) {
      if (wb.SheetNames.indexOf(name = root + i) === -1) break;
    }
  }
  check_ws_name(name);
  if (wb.SheetNames.indexOf(name) >= 0) throw new Error("Worksheet with name |" + name + "| already exists!");
  wb.SheetNames.push(name);
  wb.Sheets[name] = ws;
  return name;
}
function sheet_new(opts) {
  const out = {};
  if (opts?.dense) out["!data"] = [];
  return out;
}
function wb_sheet_idx(wb, sh) {
  if (typeof sh === "number") {
    if (sh >= 0 && wb.SheetNames.length > sh) return sh;
    throw new Error("Cannot find sheet # " + sh);
  } else if (typeof sh === "string") {
    const idx = wb.SheetNames.indexOf(sh);
    if (idx > -1) return idx;
    throw new Error("Cannot find sheet name |" + sh + "|");
  }
  throw new Error("Cannot find sheet |" + sh + "|");
}
function book_set_sheet_visibility(wb, sh, vis) {
  if (!wb.Workbook) wb.Workbook = {};
  if (!wb.Workbook.Sheets) wb.Workbook.Sheets = [];
  const idx = wb_sheet_idx(wb, sh);
  if (!wb.Workbook.Sheets[idx]) wb.Workbook.Sheets[idx] = {};
  switch (vis) {
    case 0:
    case 1:
    case 2:
      break;
    default:
      throw new Error("Bad sheet visibility setting " + vis);
  }
  wb.Workbook.Sheets[idx].Hidden = vis;
}
function cell_set_number_format(cell, fmt) {
  cell.z = fmt;
  return cell;
}
function cell_set_hyperlink(cell, target, tooltip) {
  if (!target) {
    delete cell.l;
  } else {
    cell.l = { Target: target };
    if (tooltip) cell.l.Tooltip = tooltip;
  }
  return cell;
}
function cell_set_internal_link(cell, range, tooltip) {
  return cell_set_hyperlink(cell, "#" + range, tooltip);
}
function cell_add_comment(cell, text, author) {
  if (!cell.c) cell.c = [];
  cell.c.push({ t: text, a: author || "SheetJS" });
}
function sheet_set_array_formula(ws, range, formula, dynamic) {
  const rng = typeof range !== "string" ? range : safe_decode_range(range);
  const rngstr = typeof range === "string" ? range : encode_range(range);
  for (let R = rng.s.r; R <= rng.e.r; ++R) {
    for (let C = rng.s.c; C <= rng.e.c; ++C) {
      const ref = encode_col(C) + encode_row(R);
      const dense = ws["!data"] != null;
      let cell;
      if (dense) {
        if (!ws["!data"][R]) ws["!data"][R] = [];
        cell = ws["!data"][R][C] || (ws["!data"][R][C] = { t: "z" });
      } else {
        cell = ws[ref] || (ws[ref] = { t: "z" });
      }
      cell.t = "n";
      cell.F = rngstr;
      delete cell.v;
      if (R === rng.s.r && C === rng.s.c) {
        cell.f = formula;
        if (dynamic) cell.D = true;
      }
    }
  }
  if (ws["!ref"]) {
    const wsr = decode_range(ws["!ref"]);
    if (wsr.s.r > rng.s.r) wsr.s.r = rng.s.r;
    if (wsr.s.c > rng.s.c) wsr.s.c = rng.s.c;
    if (wsr.e.r < rng.e.r) wsr.e.r = rng.e.r;
    if (wsr.e.c < rng.e.c) wsr.e.c = rng.e.c;
    ws["!ref"] = encode_range(wsr);
  }
  return ws;
}
function sheet_to_formulae(ws) {
  if (ws == null || ws["!ref"] == null) return [];
  const r = safe_decode_range(ws["!ref"]);
  const cols = [];
  const cmds = [];
  const dense = ws["!data"] != null;
  for (let C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
  for (let R = r.s.r; R <= r.e.r; ++R) {
    const rr = encode_row(R);
    for (let C = r.s.c; C <= r.e.c; ++C) {
      const y = cols[C] + rr;
      const x = dense ? (ws["!data"][R] || [])[C] : ws[y];
      if (x === void 0) continue;
      let val = "";
      let ref = y;
      if (x.F != null) {
        ref = x.F;
        if (!x.f) continue;
        val = x.f;
        if (ref.indexOf(":") === -1) ref = ref + ":" + ref;
      }
      if (x.f != null) val = x.f;
      else if (x.t === "z") continue;
      else if (x.t === "n" && x.v != null) val = "" + x.v;
      else if (x.t === "b") val = x.v ? "TRUE" : "FALSE";
      else if (x.w !== void 0) val = "'" + x.w;
      else if (x.v === void 0) continue;
      else if (x.t === "s") val = "'" + x.v;
      else val = "" + x.v;
      cmds.push(ref + "=" + val);
    }
  }
  return cmds;
}

// src/api/aoa.ts
function sheet_add_aoa(_ws, data, opts) {
  const o = opts || {};
  const dense = _ws ? _ws["!data"] != null : !!o.dense;
  const ws = _ws || (dense ? { "!data": [] } : {});
  if (dense && !ws["!data"]) ws["!data"] = [];
  let _R = 0, _C = 0;
  if (ws && o.origin != null) {
    if (typeof o.origin === "number") _R = o.origin;
    else {
      const _origin = typeof o.origin === "string" ? decode_cell(o.origin) : o.origin;
      _R = _origin.r;
      _C = _origin.c;
    }
  }
  const range = { s: { c: 1e7, r: 1e7 }, e: { c: 0, r: 0 } };
  if (ws["!ref"]) {
    const _range = safe_decode_range(ws["!ref"]);
    range.s.c = _range.s.c;
    range.s.r = _range.s.r;
    range.e.c = Math.max(range.e.c, _range.e.c);
    range.e.r = Math.max(range.e.r, _range.e.r);
    if (_R === -1) range.e.r = _R = ws["!ref"] ? _range.e.r + 1 : 0;
  } else {
    range.s.c = range.e.c = range.s.r = range.e.r = 0;
  }
  let row = [];
  let seen = false;
  for (let R = 0; R < data.length; ++R) {
    if (!data[R]) continue;
    if (!Array.isArray(data[R])) throw new Error("aoa_to_sheet expects an array of arrays");
    const __R = _R + R;
    if (dense) {
      if (!ws["!data"][__R]) ws["!data"][__R] = [];
      row = ws["!data"][__R];
    }
    const data_R = data[R];
    for (let C = 0; C < data_R.length; ++C) {
      if (typeof data_R[C] === "undefined") continue;
      let cell = { v: data_R[C], t: "" };
      const __C = _C + C;
      if (range.s.r > __R) range.s.r = __R;
      if (range.s.c > __C) range.s.c = __C;
      if (range.e.r < __R) range.e.r = __R;
      if (range.e.c < __C) range.e.c = __C;
      seen = true;
      if (data_R[C] && typeof data_R[C] === "object" && !Array.isArray(data_R[C]) && !(data_R[C] instanceof Date)) {
        cell = data_R[C];
      } else {
        if (Array.isArray(cell.v)) {
          cell.f = data_R[C][1];
          cell.v = cell.v[0];
        }
        if (cell.v === null) {
          if (cell.f) cell.t = "n";
          else if (o.nullError) {
            cell.t = "e";
            cell.v = 0;
          } else if (!o.sheetStubs) continue;
          else cell.t = "z";
        } else if (typeof cell.v === "number") {
          if (isFinite(cell.v)) cell.t = "n";
          else if (isNaN(cell.v)) {
            cell.t = "e";
            cell.v = 15;
          } else {
            cell.t = "e";
            cell.v = 7;
          }
        } else if (typeof cell.v === "boolean") {
          cell.t = "b";
        } else if (cell.v instanceof Date) {
          cell.z = o.dateNF || table_fmt[14];
          if (!o.UTC) cell.v = local_to_utc(cell.v);
          if (o.cellDates) {
            cell.t = "d";
            cell.w = SSF_format(cell.z, datenum(cell.v, o.date1904));
          } else {
            cell.t = "n";
            cell.v = datenum(cell.v, o.date1904);
            cell.w = SSF_format(cell.z, cell.v);
          }
        } else {
          cell.t = "s";
        }
      }
      if (dense) {
        if (row[__C] && row[__C].z) cell.z = row[__C].z;
        row[__C] = cell;
      } else {
        const cell_ref = encode_col(__C) + (__R + 1);
        if (ws[cell_ref] && ws[cell_ref].z) cell.z = ws[cell_ref].z;
        ws[cell_ref] = cell;
      }
    }
  }
  if (seen && range.s.c < 104e5) ws["!ref"] = encode_range(range);
  return ws;
}
function aoa_to_sheet(data, opts) {
  return sheet_add_aoa(null, data, opts);
}

// src/types.ts
var BErr = {
  0: "#NULL!",
  7: "#DIV/0!",
  15: "#VALUE!",
  23: "#REF!",
  29: "#NAME?",
  36: "#NUM!",
  42: "#N/A",
  43: "#GETTING_DATA"
};
for (const [k, v] of Object.entries(BErr)) {
}

// src/api/format.ts
function safe_format_cell(cell, v) {
  const q = cell.t === "d" && v instanceof Date;
  if (cell.z != null) {
    try {
      return cell.w = SSF_format(cell.z, q ? datenum(v) : v);
    } catch {
    }
  }
  try {
    return cell.w = SSF_format((cell.XF || {}).numFmtId || (q ? 14 : 0), q ? datenum(v) : v);
  } catch {
    return "" + v;
  }
}
function format_cell(cell, v, o) {
  if (cell == null || cell.t == null || cell.t === "z") return "";
  if (cell.w !== void 0) return cell.w;
  if (cell.t === "d" && !cell.z && o && o.dateNF) cell.z = o.dateNF;
  if (cell.t === "e") return BErr[cell.v] || String(cell.v);
  if (v == null) return safe_format_cell(cell, cell.v);
  return safe_format_cell(cell, v);
}

// src/api/json.ts
function make_json_row(sheet, r, R, cols, header, hdr, o) {
  const rr = encode_row(R);
  const defval = o.defval;
  const raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
  let isempty = true;
  const dense = sheet["!data"] != null;
  const row = header === 1 ? [] : {};
  if (header !== 1) {
    try {
      Object.defineProperty(row, "__rowNum__", { value: R, enumerable: false });
    } catch {
      row.__rowNum__ = R;
    }
  }
  if (!dense || sheet["!data"][R]) {
    for (let C = r.s.c; C <= r.e.c; ++C) {
      const val = dense ? (sheet["!data"][R] || [])[C] : sheet[cols[C] + rr];
      if (val == null || val.t === void 0) {
        if (defval === void 0) continue;
        if (hdr[C] != null) row[hdr[C]] = defval;
        continue;
      }
      let v = val.v;
      switch (val.t) {
        case "z":
          if (v == null) break;
          continue;
        case "e":
          v = v === 0 ? null : void 0;
          break;
        case "s":
        case "b":
          break;
        case "n":
          if (!val.z || !fmt_is_date(String(val.z))) break;
          v = numdate(v);
          if (typeof v === "number") break;
        /* falls through */
        case "d":
          if (!(o && (o.UTC || o.raw === false))) v = utc_to_local(new Date(v));
          break;
        default:
          throw new Error("unrecognized type " + val.t);
      }
      if (hdr[C] != null) {
        if (v == null) {
          if (val.t === "e" && v === null) row[hdr[C]] = null;
          else if (defval !== void 0) row[hdr[C]] = defval;
          else if (raw && v === null) row[hdr[C]] = null;
          else continue;
        } else {
          row[hdr[C]] = (val.t === "n" && typeof o.rawNumbers === "boolean" ? o.rawNumbers : raw) ? v : format_cell(val, v, o);
        }
        if (v != null) isempty = false;
      }
    }
  }
  return { row, isempty };
}
function sheet_to_json(sheet, opts) {
  if (sheet == null || sheet["!ref"] == null) return [];
  let header = 0, offset = 1;
  const hdr = [];
  const o = opts || {};
  const range = o.range != null ? o.range : sheet["!ref"];
  if (o.header === 1) header = 1;
  else if (o.header === "A") header = 2;
  else if (Array.isArray(o.header)) header = 3;
  else if (o.header == null) header = 0;
  let r;
  switch (typeof range) {
    case "string":
      r = safe_decode_range(range);
      break;
    case "number":
      r = safe_decode_range(sheet["!ref"]);
      r.s.r = range;
      break;
    default:
      r = range;
  }
  if (header > 0) offset = 0;
  const rr = encode_row(r.s.r);
  const cols = [];
  const out = [];
  let outi = 0;
  const dense = sheet["!data"] != null;
  let R = r.s.r;
  const header_cnt = {};
  if (dense && !sheet["!data"][R]) sheet["!data"][R] = [];
  const colinfo = o.skipHidden && sheet["!cols"] || [];
  const rowinfo = o.skipHidden && sheet["!rows"] || [];
  for (let C = r.s.c; C <= r.e.c; ++C) {
    if ((colinfo[C] || {}).hidden) continue;
    cols[C] = encode_col(C);
    const val = dense ? sheet["!data"][R][C] : sheet[cols[C] + rr];
    let v, vv;
    switch (header) {
      case 1:
        hdr[C] = C - r.s.c;
        break;
      case 2:
        hdr[C] = cols[C];
        break;
      case 3:
        hdr[C] = o.header[C - r.s.c];
        break;
      default: {
        const _val = val == null ? { w: "__EMPTY", t: "s" } : val;
        vv = v = format_cell(_val, null, o);
        let counter = header_cnt[v] || 0;
        if (!counter) header_cnt[v] = 1;
        else {
          do {
            vv = v + "_" + counter++;
          } while (header_cnt[vv]);
          header_cnt[v] = counter;
          header_cnt[vv] = 1;
        }
        hdr[C] = vv;
      }
    }
  }
  for (R = r.s.r + offset; R <= r.e.r; ++R) {
    if ((rowinfo[R] || {}).hidden) continue;
    const row = make_json_row(sheet, r, R, cols, header, hdr, o);
    if (row.isempty === false || (header === 1 ? o.blankrows !== false : !!o.blankrows))
      out[outi++] = row.row;
  }
  out.length = outi;
  return out;
}
function sheet_add_json(_ws, js, opts) {
  const o = opts || {};
  const dense = _ws ? _ws["!data"] != null : !!o.dense;
  const offset = +!o.skipHeader;
  const ws = _ws || {};
  if (!_ws && dense) ws["!data"] = [];
  let _R = 0, _C = 0;
  if (ws && o.origin != null) {
    if (typeof o.origin === "number") _R = o.origin;
    else {
      const _origin = typeof o.origin === "string" ? decode_cell(o.origin) : o.origin;
      _R = _origin.r;
      _C = _origin.c;
    }
  }
  const range = { s: { c: 0, r: 0 }, e: { c: _C, r: _R + js.length - 1 + offset } };
  if (ws["!ref"]) {
    const _range = safe_decode_range(ws["!ref"]);
    range.e.c = Math.max(range.e.c, _range.e.c);
    range.e.r = Math.max(range.e.r, _range.e.r);
    if (_R === -1) {
      _R = _range.e.r + 1;
      range.e.r = _R + js.length - 1 + offset;
    }
  } else {
    if (_R === -1) {
      _R = 0;
      range.e.r = js.length - 1 + offset;
    }
  }
  const hdr = o.header || [];
  let C = 0;
  js.forEach((JS, R) => {
    if (dense && !ws["!data"][_R + R + offset]) ws["!data"][_R + R + offset] = [];
    const ROW = dense ? ws["!data"][_R + R + offset] : null;
    keys(JS).forEach((k) => {
      if ((C = hdr.indexOf(k)) === -1) hdr[C = hdr.length] = k;
      let v = JS[k];
      let t = "z";
      let z = "";
      const ref = dense ? "" : encode_col(_C + C) + encode_row(_R + R + offset);
      const cell = dense ? ROW[_C + C] : ws[ref];
      if (v && typeof v === "object" && !(v instanceof Date)) {
        if (dense) ROW[_C + C] = v;
        else ws[ref] = v;
      } else {
        if (typeof v === "number") t = "n";
        else if (typeof v === "boolean") t = "b";
        else if (typeof v === "string") t = "s";
        else if (v instanceof Date) {
          t = "d";
          if (!o.UTC) v = local_to_utc(v);
          if (!o.cellDates) {
            t = "n";
            v = datenum(v);
          }
          z = cell != null && cell.z && fmt_is_date(String(cell.z)) ? String(cell.z) : o.dateNF || table_fmt[14];
        } else if (v === null && o.nullError) {
          t = "e";
          v = 0;
        }
        if (!cell) {
          const newCell = { t, v };
          if (z) newCell.z = z;
          if (dense) ROW[_C + C] = newCell;
          else ws[ref] = newCell;
        } else {
          cell.t = t;
          cell.v = v;
          delete cell.w;
          if (z) cell.z = z;
        }
      }
    });
  });
  range.e.c = Math.max(range.e.c, _C + hdr.length - 1);
  const __R = encode_row(_R);
  if (dense && !ws["!data"][_R]) ws["!data"][_R] = [];
  if (offset) {
    for (C = 0; C < hdr.length; ++C) {
      if (dense) ws["!data"][_R][C + _C] = { t: "s", v: hdr[C] };
      else ws[encode_col(C + _C) + __R] = { t: "s", v: hdr[C] };
    }
  }
  ws["!ref"] = encode_range(range);
  return ws;
}
function json_to_sheet(js, opts) {
  return sheet_add_json(null, js, opts);
}

// src/api/csv.ts
var qreg = /"/g;
function make_csv_row(sheet, r, R, cols, fs3, rs, FS, w, o) {
  let isempty = true;
  const row = [];
  const rr = encode_row(R);
  const dense = sheet["!data"] != null;
  const datarow = dense ? sheet["!data"][R] || [] : [];
  for (let C = r.s.c; C <= r.e.c; ++C) {
    if (!cols[C]) continue;
    const val = dense ? datarow[C] : sheet[cols[C] + rr];
    let txt = "";
    if (val == null) txt = "";
    else if (val.v != null) {
      isempty = false;
      txt = "" + (o.rawNumbers && val.t === "n" ? val.v : format_cell(val, null, o));
      for (let i = 0, cc = 0; i !== txt.length; ++i) {
        if ((cc = txt.charCodeAt(i)) === fs3 || cc === rs || cc === 10 || cc === 13 || cc === 34 || o.forceQuotes) {
          txt = '"' + txt.replace(qreg, '""') + '"';
          break;
        }
      }
      if (txt === "ID" && w === 0 && row.length === 0) txt = '"ID"';
    } else if (val.f != null && !val.F) {
      isempty = false;
      txt = "=" + val.f;
      if (txt.indexOf(",") >= 0) txt = '"' + txt.replace(qreg, '""') + '"';
    } else txt = "";
    row.push(txt);
  }
  if (o.strip) while (row[row.length - 1] === "") --row.length;
  if (o.blankrows === false && isempty) return null;
  return row.join(FS);
}
function sheet_to_csv(sheet, opts) {
  const out = [];
  const o = opts == null ? {} : opts;
  if (sheet == null || sheet["!ref"] == null) return "";
  const r = safe_decode_range(sheet["!ref"]);
  const FS = o.FS !== void 0 ? o.FS : ",";
  const fs3 = FS.charCodeAt(0);
  const RS = o.RS !== void 0 ? o.RS : "\n";
  const rs = RS.charCodeAt(0);
  const cols = [];
  const colinfo = o.skipHidden && sheet["!cols"] || [];
  const rowinfo = o.skipHidden && sheet["!rows"] || [];
  for (let C = r.s.c; C <= r.e.c; ++C) {
    if (!(colinfo[C] || {}).hidden) cols[C] = encode_col(C);
  }
  let w = 0;
  for (let R = r.s.r; R <= r.e.r; ++R) {
    if ((rowinfo[R] || {}).hidden) continue;
    const row = make_csv_row(sheet, r, R, cols, fs3, rs, FS, w, o);
    if (row == null) continue;
    if (row || o.blankrows !== false) out.push((w++ ? RS : "") + row);
  }
  return out.join("");
}
function sheet_to_txt(sheet, opts) {
  const o = opts || {};
  o.FS = "	";
  o.RS = "\n";
  return sheet_to_csv(sheet, o);
}

// src/api/html.ts
var HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
var HTML_END = "</body></html>";
function make_html_row(ws, r, R, o) {
  const M = ws["!merges"] || [];
  const oo = [];
  const dense = ws["!data"] != null;
  for (let C = r.s.c; C <= r.e.c; ++C) {
    let RS = 0, CS = 0;
    for (let j = 0; j < M.length; ++j) {
      if (M[j].s.r > R || M[j].s.c > C) continue;
      if (M[j].e.r < R || M[j].e.c < C) continue;
      if (M[j].s.r < R || M[j].s.c < C) {
        RS = -1;
        break;
      }
      RS = M[j].e.r - M[j].s.r + 1;
      CS = M[j].e.c - M[j].s.c + 1;
      break;
    }
    if (RS < 0) continue;
    const coord = encode_col(C) + encode_row(R);
    let cell = dense ? (ws["!data"][R] || [])[C] : ws[coord];
    if (cell && cell.t === "n" && cell.v != null && !isFinite(cell.v)) {
      if (isNaN(cell.v)) cell = { t: "e", v: 36, w: BErr[36] };
      else cell = { t: "e", v: 7, w: BErr[7] };
    }
    let w = cell && cell.v != null && (cell.h || escapehtml(cell.w || (format_cell(cell), cell.w) || "")) || "";
    const sp = {};
    if (RS > 1) sp.rowspan = String(RS);
    if (CS > 1) sp.colspan = String(CS);
    if (o.editable) {
      w = '<span contenteditable="true">' + w + "</span>";
    } else if (cell) {
      sp["data-t"] = cell && cell.t || "z";
      if (cell.v != null) sp["data-v"] = escapehtml(cell.v instanceof Date ? cell.v.toISOString() : String(cell.v));
      if (cell.z != null) sp["data-z"] = String(cell.z);
      if (cell.f != null) sp["data-f"] = escapehtml(cell.f);
      if (cell.l && (cell.l.Target || "#").charAt(0) !== "#" && (!o.sanitizeLinks || (cell.l.Target || "").slice(0, 11).toLowerCase() !== "javascript:")) {
        w = '<a href="' + escapehtml(cell.l.Target) + '">' + w + "</a>";
      }
    }
    sp.id = (o.id || "sjs") + "-" + coord;
    oo.push(writextag("td", w, sp));
  }
  return "<tr>" + oo.join("") + "</tr>";
}
function make_html_preamble(_ws, _r, o) {
  return "<table" + (o && o.id ? ' id="' + o.id + '"' : "") + ">";
}
function sheet_to_html(ws, opts) {
  const o = opts || {};
  const header = o.header != null ? o.header : HTML_BEGIN;
  const footer = o.footer != null ? o.footer : HTML_END;
  const out = [header];
  const r = decode_range(ws["!ref"] || "A1");
  out.push(make_html_preamble(ws, r, o));
  if (ws["!ref"]) {
    for (let R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
  }
  out.push("</table>" + footer);
  return out.join("");
}

// src/index.ts
var version = "1.0.0-alpha.0";

export { SSF_format, aoa_to_sheet, book_append_sheet, book_new, book_set_sheet_visibility, cell_add_comment, cell_set_hyperlink, cell_set_internal_link, cell_set_number_format, decode_cell, decode_col, decode_range, decode_row, encode_cell, encode_col, encode_range, encode_row, format_cell, json_to_sheet, read, readFile, sheet_add_aoa, sheet_add_json, sheet_new, sheet_set_array_formula, sheet_to_csv, sheet_to_formulae, sheet_to_html, sheet_to_json, sheet_to_txt, version, wb_sheet_idx, write, writeFile };
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map
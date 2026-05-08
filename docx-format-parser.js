const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");

function xmlDecode(s) {
  if (!s) return "";
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

function wAttr(node, localName) {
  if (!node || !node.attributes) return null;
  for (let i = 0; i < node.attributes.length; i++) {
    const a = node.attributes[i];
    if (a.localName === localName || a.nodeName === `w:${localName}`) return a.nodeValue;
  }
  return null;
}

function rAttr(node, localName) {
  if (!node || !node.attributes) return null;
  for (let i = 0; i < node.attributes.length; i++) {
    const a = node.attributes[i];
    if (a.localName === localName || a.nodeName === `r:${localName}`) return a.nodeValue;
  }
  return null;
}

function aAttr(node, localName) {
  if (!node || !node.attributes) return null;
  for (let i = 0; i < node.attributes.length; i++) {
    const a = node.attributes[i];
    if (a.localName === localName || a.nodeName === `a:${localName}`) return a.nodeValue;
  }
  return null;
}

function relAttr(node, localName) {
  if (!node || !node.attributes) return null;
  for (let i = 0; i < node.attributes.length; i++) {
    const a = node.attributes[i];
    if (a.localName === localName) return a.nodeValue;
  }
  return null;
}

function findChild(node, localName) {
  if (!node || !node.childNodes) return null;
  for (let i = 0; i < node.childNodes.length; i++) {
    const c = node.childNodes[i];
    if (c.localName === localName || c.nodeName === `w:${localName}`) return c;
  }
  return null;
}

function findDeep(node, localName) {
  if (!node) return null;
  const direct = findChild(node, localName);
  if (direct) return direct;
  if (node.childNodes) {
    for (let i = 0; i < node.childNodes.length; i++) {
      const found = findDeep(node.childNodes[i], localName);
      if (found) return found;
    }
  }
  return null;
}

function findChildren(node, localName) {
  const results = [];
  if (!node || !node.childNodes) return results;
  for (let i = 0; i < node.childNodes.length; i++) {
    const c = node.childNodes[i];
    if (c.localName === localName || c.nodeName === `w:${localName}`) results.push(c);
  }
  return results;
}

function isElem(node, localName) {
  if (!node || node.nodeType !== 1) return false;
  return node.localName === localName || node.nodeName === `w:${localName}`;
}

function isText(node) {
  return node && (node.nodeType === 3 || node.nodeType === 4);
}

class DocxFormatParser {
  constructor(buffer) {
    this.buffer = buffer;
    this.zip = null;
    this.styles = {};
    this.numbering = {};
    this.themeColors = {};
    this.relImages = {};
    this.doc = null;
    this.body = null;
  }

  async parse() {
    this.zip = await JSZip.loadAsync(this.buffer);
    await this._parseTheme();
    await this._parseStyles();
    await this._parseNumbering();
    await this._parseRels();
    const docXml = await this._readFile("word/document.xml");
    this.doc = new DOMParser().parseFromString(docXml, "text/xml");
    this.body = findChild(this.doc.documentElement, "body");
    let docBgStyle = "";
    const bgEl = findChild(this.doc.documentElement, "background");
    if (bgEl) {
      const color = wAttr(bgEl, "color");
      if (color && color !== "auto") docBgStyle = `background-color:#${color}`;
    }
    let html = this._convertBodyToHtml();
    html = await this._resolveImages(html);
    const wrappedHtml = docBgStyle ? `<div style="${docBgStyle}">${html}</div>` : html;
    return { html: wrappedHtml, text: this._stripHtml(wrappedHtml) };
  }

  async _readFile(name) {
    const file = this.zip.file(name);
    if (!file) return "";
    return file.async("string");
  }

  async _parseTheme() {
    const xml = await this._readFile("word/theme/theme1.xml");
    if (!xml) return;
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const clrScheme = findDeep(doc.documentElement, "clrScheme");
    if (!clrScheme) return;
    const colorMap = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    for (const key of colorMap) {
      const section = findChild(clrScheme, key);
      if (!section) continue;
      let srgb = null;
      for (let i = 0; i < section.childNodes.length; i++) {
        const c = section.childNodes[i];
        if (c.localName === "srgbClr" || c.nodeName === "a:srgbClr") {
          srgb = aAttr(c, "val");
          break;
        }
      }
      if (srgb && /^[0-9a-fA-F]{6}$/.test(srgb)) this.themeColors[key] = `#${srgb}`;
    }
  }

  async _parseStyles() {
    const xml = await this._readFile("word/styles.xml");
    if (!xml) return;
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const root = doc.documentElement;
    for (let i = 0; i < root.childNodes.length; i++) {
      const node = root.childNodes[i];
      if (node.localName === "style" || node.nodeName === "w:style") {
        const id = wAttr(node, "styleId");
        if (id) this.styles[id] = this._parseStyleProps(node);
      }
    }
    const docDefaults = findChild(root, "docDefaults");
    if (docDefaults) this.docDefaultStyle = this._parseStyleProps(docDefaults);
  }

  _parseStyleProps(node) {
    const props = {};
    const rPr = findDeep(node, "rPr");
    if (rPr) {
      const szEl = findDeep(rPr, "sz");
      if (szEl) props.sz = wAttr(szEl, "val");
      const szCsEl = findDeep(rPr, "szCs");
      if (szCsEl) props.szCs = wAttr(szCsEl, "val");
      const rFonts = findDeep(rPr, "rFonts");
      if (rFonts) {
        props.ascii = wAttr(rFonts, "ascii") || wAttr(rFonts, "eastAsia") || wAttr(rFonts, "hAnsi");
        props.eastAsia = wAttr(rFonts, "eastAsia") || wAttr(rFonts, "ascii");
        props.hAnsi = wAttr(rFonts, "hAnsi") || wAttr(rFonts, "ascii");
      }
      const color = findDeep(rPr, "color");
      if (color) props.color = wAttr(color, "val");
      const bold = findDeep(rPr, "b");
      if (bold) {
        const val = wAttr(bold, "val");
        props.bold = val !== "0" && val !== "false";
      }
      const italic = findDeep(rPr, "i");
      if (italic) {
        const val = wAttr(italic, "val");
        props.italic = val !== "0" && val !== "false";
      }
      const underline = findDeep(rPr, "u");
      if (underline) {
        const val = wAttr(underline, "val");
        if (val && val !== "none") props.underline = val;
      }
      const strike = findDeep(rPr, "strike");
      if (strike) {
        const val = wAttr(strike, "val");
        props.strike = val !== "0" && val !== "false";
      }
      const shd = findDeep(rPr, "shd");
      if (shd) {
        const fill = wAttr(shd, "fill");
        if (fill && fill !== "auto") props.highlight = `#${fill}`;
      }
      const vertAlign = findDeep(rPr, "vertAlign");
      if (vertAlign) props.vertAlign = wAttr(vertAlign, "val");
    }

    const pPr = findDeep(node, "pPr");
    if (pPr) {
      const jc = findDeep(pPr, "jc");
      if (jc) props.jc = wAttr(jc, "val");
      const ind = findDeep(pPr, "ind");
      if (ind) {
        props.indLeft = wAttr(ind, "left") || wAttr(ind, "start");
        props.indRight = wAttr(ind, "right") || wAttr(ind, "end");
        props.indFirstLine = wAttr(ind, "firstLine") || wAttr(ind, "firstLineChars");
        props.indHanging = wAttr(ind, "hanging");
      }
      const spacing = findDeep(pPr, "spacing");
      if (spacing) {
        props.spBefore = wAttr(spacing, "before");
        props.spAfter = wAttr(spacing, "after");
        props.lineSpacing = wAttr(spacing, "line");
        props.lineRule = wAttr(spacing, "lineRule");
      }
      const pBdr = findDeep(pPr, "pBdr");
      if (pBdr) {
        const bottom = findChild(pBdr, "bottom");
        if (bottom) {
          const val = wAttr(bottom, "val");
          if (val && val !== "none") {
            props.pBorderBottom = { style: val, sz: wAttr(bottom, "sz") || "4", color: wAttr(bottom, "color") || "000000" };
          }
        }
      }
      const pShd = findDeep(pPr, "shd");
      if (pShd) {
        const fill = wAttr(pShd, "fill");
        if (fill && fill !== "auto") props.pBackground = `#${fill}`;
      }
      const numPr = findDeep(pPr, "numPr");
      if (numPr) {
        const ilvlEl = findChild(numPr, "ilvl");
        const numIdEl = findChild(numPr, "numId");
        if (ilvlEl) props.ilvl = wAttr(ilvlEl, "val");
        if (numIdEl) props.numId = wAttr(numIdEl, "val");
      }
    }
    return props;
  }

  async _parseNumbering() {
    const xml = await this._readFile("word/numbering.xml");
    if (!xml) return;
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const root = doc.documentElement;
    const abstractNums = {};

    for (let i = 0; i < root.childNodes.length; i++) {
      const node = root.childNodes[i];
      if (node.localName === "abstractNum" || node.nodeName === "w:abstractNum") {
        const id = wAttr(node, "abstractNumId");
        if (!id) continue;
        abstractNums[id] = {};
        for (let j = 0; j < node.childNodes.length; j++) {
          const lvl = node.childNodes[j];
          if (lvl.localName === "lvl" || lvl.nodeName === "w:lvl") {
            const ilvl = wAttr(lvl, "ilvl");
            if (ilvl == null) continue;
            const fmtEl = findChild(lvl, "numFmt");
            const lvlTextEl = findChild(lvl, "lvlText");
            const startEl = findChild(lvl, "start");
            abstractNums[id][ilvl] = {
              fmt: fmtEl ? wAttr(fmtEl, "val") : "decimal",
              text: lvlTextEl ? wAttr(lvlTextEl, "val") : null,
              start: startEl ? parseInt(wAttr(startEl, "val") || "1", 10) : 1,
            };
          }
        }
      }
    }

    for (let i = 0; i < root.childNodes.length; i++) {
      const node = root.childNodes[i];
      if (node.localName === "num" || node.nodeName === "w:num") {
        const numId = wAttr(node, "numId");
        if (!numId) continue;
        const abstractIdEl = findChild(node, "abstractNumId");
        const abstractId = abstractIdEl ? wAttr(abstractIdEl, "val") : null;
        const levels = {};
        for (let j = 0; j < node.childNodes.length; j++) {
          const lvl = node.childNodes[j];
          if (lvl.localName === "lvl" || lvl.nodeName === "w:lvl") {
            const ilvl = wAttr(lvl, "ilvl");
            if (ilvl == null) continue;
            const fmtEl = findChild(lvl, "numFmt");
            const lvlTextEl = findChild(lvl, "lvlText");
            const startEl = findChild(lvl, "start");
            levels[ilvl] = {
              fmt: fmtEl ? wAttr(fmtEl, "val") : "decimal",
              text: lvlTextEl ? wAttr(lvlTextEl, "val") : null,
              start: startEl ? parseInt(wAttr(startEl, "val") || "1", 10) : 1,
            };
          }
        }
        if (abstractId && abstractNums[abstractId]) {
          for (const [ilvl, lvl] of Object.entries(abstractNums[abstractId])) {
            if (!levels[ilvl]) levels[ilvl] = lvl;
          }
        }
        this.numbering[numId] = { abstractId, levels };
      }
    }
  }

  async _parseRels() {
    const xml = await this._readFile("word/_rels/document.xml.rels");
    if (!xml) return;
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const root = doc.documentElement;
    for (let i = 0; i < root.childNodes.length; i++) {
      const node = root.childNodes[i];
      if (node.localName === "Relationship") {
        const id = relAttr(node, "Id");
        const target = relAttr(node, "Target");
        const type = relAttr(node, "Type");
        if (id && target && type && type.includes("image")) {
          this.relImages[id] = target.replace(/^\/?/, "");
        }
      }
    }
  }

  async _getImageDataUri(relId) {
    const target = this.relImages[relId];
    if (!target) return null;
    const filePath = `word/${target}`;
    const file = this.zip.file(filePath);
    if (!file) return null;
    const data = await file.async("base64");
    const ext = target.split(".").pop().toLowerCase();
    const mimeMap = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp", emf: "image/emf", wmf: "image/x-wmf", tiff: "image/tiff", tif: "image/tiff" };
    return `data:${mimeMap[ext] || "image/png"};base64,${data}`;
  }

  _resolveStyle(styleId, visited = new Set()) {
    if (!styleId || visited.has(styleId)) return {};
    visited.add(styleId);
    const style = this.styles[styleId];
    if (!style) return {};
    const baseId = style.basedOn;
    const base = baseId ? this._resolveStyle(baseId, visited) : {};
    return { ...base, ...style };
  }

  _twipsToPt(val) { return val ? (parseInt(val, 10) / 20).toFixed(1) : null; }
  _twipsToEm(val) { return val ? (parseInt(val, 10) / 240).toFixed(3) : null; }
  _halfPointsToPt(val) { return val ? (parseInt(val, 10) / 2).toFixed(1) : null; }
  _ptToFontSize(sz) { const pt = this._halfPointsToPt(sz); return pt ? `${pt}pt` : null; }

  _convertBodyToHtml() {
    if (!this.body) return "";
    const parts = [];
    this._listCounters = {};
    for (let i = 0; i < this.body.childNodes.length; i++) {
      const child = this.body.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (isElem(child, "p")) {
        parts.push(this._convertPara(child));
      } else if (isElem(child, "tbl")) {
        parts.push(this._convertTable(child));
      } else if (isElem(child, "sdt")) {
        const sdtContent = findChild(child, "sdtContent");
        if (sdtContent) {
          for (let j = 0; j < sdtContent.childNodes.length; j++) {
            const c = sdtContent.childNodes[j];
            if (c.nodeType !== 1) continue;
            if (isElem(c, "p")) parts.push(this._convertPara(c));
            else if (isElem(c, "tbl")) parts.push(this._convertTable(c));
          }
        }
      }
    }
    return parts.join("\n");
  }

  _convertPara(node) {
    const pPr = findChild(node, "pPr");
    const styleId = pPr ? wAttr(findChild(pPr, "pStyle"), "val") : null;
    const resolved = this._resolveStyle(styleId);
    const merged = { ...resolved };

    if (pPr) {
      const jcEl = findChild(pPr, "jc");
      if (jcEl) merged.jc = wAttr(jcEl, "val");
      const ind = findChild(pPr, "ind");
      if (ind) {
        merged.indLeft = wAttr(ind, "left") || wAttr(ind, "start") || merged.indLeft;
        merged.indRight = wAttr(ind, "right") || wAttr(ind, "end") || merged.indRight;
        merged.indFirstLine = wAttr(ind, "firstLine") || wAttr(ind, "firstLineChars") || merged.indFirstLine;
        merged.indHanging = wAttr(ind, "hanging") || merged.indHanging;
      }
      const spacing = findChild(pPr, "spacing");
      if (spacing) {
        merged.spBefore = wAttr(spacing, "before") || merged.spBefore;
        merged.spAfter = wAttr(spacing, "after") || merged.spAfter;
        merged.lineSpacing = wAttr(spacing, "line") || merged.lineSpacing;
        merged.lineRule = wAttr(spacing, "lineRule") || merged.lineRule;
      }
      const pShd = findChild(pPr, "shd");
      if (pShd) {
        const fill = wAttr(pShd, "fill");
        if (fill && fill !== "auto") merged.pBackground = `#${fill}`;
      }
      const numPr = findChild(pPr, "numPr");
      if (numPr) {
        const ilvlEl = findChild(numPr, "ilvl");
        const numIdEl = findChild(numPr, "numId");
        if (ilvlEl) merged.ilvl = wAttr(ilvlEl, "val") || "0";
        if (numIdEl) merged.numId = wAttr(numIdEl, "val");
      }
    }

    const pStyle = [];
    if (merged.jc === "center") pStyle.push("text-align:center");
    else if (merged.jc === "right") pStyle.push("text-align:right");
    else if (merged.jc === "both" || merged.jc === "distribute") pStyle.push("text-align:justify");
    if (merged.indLeft) pStyle.push(`margin-left:${this._twipsToEm(merged.indLeft)}em`);
    if (merged.indRight) pStyle.push(`margin-right:${this._twipsToEm(merged.indRight)}em`);
    if (merged.indFirstLine) pStyle.push(`text-indent:${this._twipsToEm(merged.indFirstLine)}em`);
    else if (merged.indHanging) pStyle.push(`text-indent:-${this._twipsToEm(merged.indHanging)}em`);
    if (merged.spBefore) pStyle.push(`margin-top:${this._twipsToPt(merged.spBefore)}pt`);
    if (merged.spAfter) pStyle.push(`margin-bottom:${this._twipsToPt(merged.spAfter)}pt`);
    if (merged.lineSpacing && merged.lineRule !== "exact") {
      const lineEm = this._twipsToEm(merged.lineSpacing);
      if (lineEm) pStyle.push(`line-height:${lineEm}`);
    } else if (merged.lineSpacing && merged.lineRule === "exact") {
      const linePt = this._twipsToPt(merged.lineSpacing);
      if (linePt) pStyle.push(`line-height:${linePt}pt`);
    }
    if (merged.pBackground) pStyle.push(`background-color:${merged.pBackground}`);
    if (merged.pBorderBottom) {
      const border = merged.pBorderBottom;
      const w = Math.round(parseInt(border.sz, 10) / 8);
      pStyle.push(`border-bottom:${w}px solid #${border.color}`);
    }

    let prefix = "";
    if (merged.numId && merged.ilvl !== undefined) {
      const num = this.numbering[merged.numId];
      if (num) {
        const ilvl = merged.ilvl || "0";
        const level = num.levels[ilvl];
        if (level) {
          const counterKey = `${merged.numId}_${ilvl}`;
          if (!this._listCounters[counterKey]) this._listCounters[counterKey] = level.start || 1;
          const numText = (level.text || "%1.").replace("%1", String(this._listCounters[counterKey]));
          this._listCounters[counterKey]++;
          const indent = parseInt(ilvl, 10) * 24;
          pStyle.push(`margin-left:${24 + indent}pt`, "text-indent:-24pt");
          prefix = `<span style="color:inherit;white-space:pre">${this._esc(numText)} </span>`;
        }
      }
    }

    const styleAttr = pStyle.length ? ` style="${pStyle.join(";")}"` : "";
    const runs = this._convertInline(node);
    const tag = styleId && /^heading/i.test(styleId) ? (styleId === "Heading1" ? "h1" : styleId === "Heading2" ? "h2" : styleId === "Heading3" ? "h3" : "h4") : "p";

    if (!runs.trim() && !prefix) return `<${tag}${styleAttr}><br></${tag}>`;
    return `<${tag}${styleAttr}>${prefix}${runs}</${tag}>`;
  }

  _convertInline(parentNode) {
    const parts = [];
    for (let i = 0; i < parentNode.childNodes.length; i++) {
      const child = parentNode.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (isElem(child, "r")) {
        parts.push(this._convertRun(child));
      } else if (isElem(child, "hyperlink")) {
        parts.push(`<a>${this._convertInline(child)}</a>`);
      } else if (isElem(child, "smartTag")) {
        parts.push(this._convertInline(child));
      } else if (isElem(child, "sdt")) {
        const sdtContent = findChild(child, "sdtContent");
        if (sdtContent) parts.push(this._convertInline(sdtContent));
      } else if (isElem(child, "ins")) {
        parts.push(this._convertInline(child));
      } else if (isElem(child, "del")) {
      } else if (isElem(child, "bookmarkStart") || isElem(child, "bookmarkEnd") || isElem(child, "proofErr") || isElem(child, "rPr")) {
      }
    }
    return parts.join("");
  }

  _convertRun(node) {
    const rPr = findChild(node, "rPr");
    const style = [];

    let isBold = false, isItalic = false, isUnderline = false, isStrike = false;
    let color = null, highlight = null, fontSize = null, fontFamily = null, vertAlign = null;

    const resolvedRpr = this._resolveStyle(rPr ? wAttr(findChild(rPr, "rStyle"), "val") : null);

    if (rPr) {
      const boldEl = findChild(rPr, "b");
      if (boldEl) { const v = wAttr(boldEl, "val"); isBold = v !== "0" && v !== "false"; }
      else if (resolvedRpr.bold) isBold = true;

      const italicEl = findChild(rPr, "i");
      if (italicEl) { const v = wAttr(italicEl, "val"); isItalic = v !== "0" && v !== "false"; }
      else if (resolvedRpr.italic) isItalic = true;

      const underlineEl = findChild(rPr, "u");
      if (underlineEl) { const v = wAttr(underlineEl, "val"); isUnderline = v && v !== "none"; }
      else if (resolvedRpr.underline) isUnderline = true;

      const strikeEl = findChild(rPr, "strike");
      if (strikeEl) { const v = wAttr(strikeEl, "val"); isStrike = v !== "0" && v !== "false"; }
      else if (resolvedRpr.strike) isStrike = true;

      const colorEl = findChild(rPr, "color");
      color = colorEl ? wAttr(colorEl, "val") : resolvedRpr.color;
      if (color === "auto") color = null;

      const shdEl = findChild(rPr, "shd");
      if (shdEl) {
        const fill = wAttr(shdEl, "fill");
        if (fill && fill !== "auto") highlight = `#${fill}`;
      }
      if (!highlight && resolvedRpr.highlight) highlight = resolvedRpr.highlight;

      const szEl = findChild(rPr, "sz");
      fontSize = szEl ? wAttr(szEl, "val") : resolvedRpr.sz;

      const rFonts = findChild(rPr, "rFonts");
      if (rFonts) fontFamily = wAttr(rFonts, "ascii") || wAttr(rFonts, "eastAsia") || wAttr(rFonts, "hAnsi");
      else fontFamily = resolvedRpr.ascii;

      const vertAlignEl = findChild(rPr, "vertAlign");
      if (vertAlignEl) vertAlign = wAttr(vertAlignEl, "val");
      else if (resolvedRpr.vertAlign) vertAlign = resolvedRpr.vertAlign;
    }

    if (isBold) style.push("font-weight:bold");
    if (isItalic) style.push("font-style:italic");
    if (isUnderline) style.push("text-decoration:underline");
    if (isStrike) style.push(style.some(s => s.startsWith("text-decoration")) ? " text-decoration:line-through" : "text-decoration:line-through");
    if (color) style.push(`color:${color.startsWith("#") ? color : `#${color}`}`);
    if (highlight) style.push(`background-color:${highlight}`);
    if (fontSize) style.push(`font-size:${this._ptToFontSize(fontSize)}`);
    if (fontFamily) style.push(`font-family:"${fontFamily}",serif`);
    if (vertAlign === "superscript") style.push("vertical-align:super;font-size:0.8em");
    else if (vertAlign === "subscript") style.push("vertical-align:sub;font-size:0.8em");

    const styleAttr = style.length ? ` style="${style.join(";")}"` : "";
    let content = "";

    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (isElem(child, "t")) {
        const text = child.textContent || "";
        content += this._esc(xmlDecode(text));
      } else if (isElem(child, "delText")) {
      } else if (isElem(child, "tab")) {
        content += "&nbsp;&nbsp;&nbsp;&nbsp;";
      } else if (isElem(child, "br")) {
        content += "<br>";
      } else if (isElem(child, "cr")) {
        content += "<br>";
      } else if (isElem(child, "sym")) {
        const charCode = wAttr(child, "char");
        if (charCode) {
          const code = parseInt(charCode, 16);
          if (!isNaN(code)) content += String.fromCharCode(code);
        }
      } else if (isElem(child, "drawing")) {
        content += this._convertDrawing(child);
      } else if (isElem(child, "pict")) {
        content += this._convertPict(child);
      } else if (isElem(child, "noBreakHyphen")) {
        content += "\u2011";
      } else if (isElem(child, "softHyphen")) {
        content += "\u00AD";
      }
    }

    if (!content) return "";
    return `<span${styleAttr}>${content}</span>`;
  }

  _convertDrawing(node) {
    let content = "";
    const traverse = (n) => {
      if (n.nodeType === 1) {
        if (n.localName === "blip" || n.nodeName === "a:blip") {
          const embed = rAttr(n, "embed");
          if (embed && this.relImages[embed]) {
            const target = this.relImages[embed];
            const filePath = `word/${target}`;
            const file = this.zip.file(filePath);
            if (file) {
              const ext = target.split(".").pop().toLowerCase();
              const mimeMap = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp", emf: "image/emf", wmf: "image/x-wmf", tiff: "image/tiff" };
              const mime = mimeMap[ext] || "image/png";
              content += `__IMG_${embed}__`;
            }
          }
        }
        if (n.childNodes) {
          for (let i = 0; i < n.childNodes.length; i++) traverse(n.childNodes[i]);
        }
      }
    };
    traverse(node);
    if (content) {
      const embedMatch = content.match(/__IMG_([^_]+)__/);
      if (embedMatch) return `<img data-rel="${embedMatch[1]}" style="max-width:100%;height:auto" />`;
    }
    return "";
  }

  _convertPict(node) {
    let content = "";
    const traverse = (n) => {
      if (n.nodeType === 1) {
        const embed = rAttr(n, "embed") || rAttr(n, "id");
        if (embed && this.relImages[embed]) {
          content = `<img data-rel="${embed}" style="max-width:100%;height:auto" />`;
        }
        if (n.childNodes) {
          for (let i = 0; i < n.childNodes.length; i++) traverse(n.childNodes[i]);
        }
      }
    };
    traverse(node);
    return content;
  }

  async _resolveImages(html) {
    const imgRe = /<img\s+data-rel="([^"]+)"\s*\/?>/g;
    let result = html;
    let m;
    const matches = [];
    while ((m = imgRe.exec(html)) !== null) {
      matches.push({ full: m[0], relId: m[1] });
    }
    for (const match of matches) {
      const dataUri = await this._getImageDataUri(match.relId);
      if (dataUri) {
        result = result.replace(match.full, `<img src="${dataUri}" style="max-width:100%;height:auto" />`);
      } else {
        result = result.replace(match.full, "");
      }
    }
    return result;
  }

  _convertTable(node) {
    const tblPr = findChild(node, "tblPr");
    const tableStyle = [];

    if (tblPr) {
      const tblW = findChild(tblPr, "tblW");
      if (tblW) {
        const type = wAttr(tblW, "type");
        const val = wAttr(tblW, "w");
        if (type === "pct") tableStyle.push("width:100%");
        else if (val) tableStyle.push(`width:${this._twipsToPt(val)}pt`);
      }
      const jc = findChild(tblPr, "jc");
      if (jc) {
        const val = wAttr(jc, "val");
        if (val === "center") tableStyle.push("margin:0 auto");
        else if (val === "right") tableStyle.push("margin:0 0 0 auto");
      }
      const tblBorders = findChild(tblPr, "tblBorders");
      if (tblBorders) {
        const borders = this._parseBorders(tblBorders);
        if (borders) tableStyle.push(borders);
      }
      const tblShd = findChild(tblPr, "shd");
      if (tblShd) {
        const fill = wAttr(tblShd, "fill");
        if (fill && fill !== "auto") tableStyle.push(`background-color:#${fill}`);
      }
    }

    const tblCellMar = tblPr ? findChild(tblPr, "tblCellMar") : null;
    let cellPadding = null;
    if (tblCellMar) {
      const top = findChild(tblCellMar, "top");
      if (top) cellPadding = wAttr(top, "w");
    }

    const colWidths = [];
    const tblGrid = findChild(node, "tblGrid");
    if (tblGrid) {
      const gridCols = findChildren(tblGrid, "gridCol");
      for (const col of gridCols) {
        const w = wAttr(col, "w");
        colWidths.push(w ? this._twipsToPt(w) : null);
      }
    }

    const rowEls = findChildren(node, "tr");
    const rows = [];
    for (const rowEl of rowEls) {
      const trPr = findChild(rowEl, "trPr");
      const isHeader = trPr ? !!findChild(trPr, "tblHeader") : false;
      const cellEls = findChildren(rowEl, "tc");
      const cells = [];
      let cellIdx = 0;
      for (const cellEl of cellEls) {
        const cellResult = this._convertCell(cellEl, isHeader, colWidths[cellIdx], cellPadding);
        cells.push(cellResult);
        const gridSpanEl = findChild(cellEl, "gridSpan");
        if (gridSpanEl) {
          cellIdx += parseInt(wAttr(gridSpanEl, "val") || "1", 10);
        } else {
          cellIdx++;
        }
      }
      rows.push(cells);
    }

    const styleAttr = tableStyle.length ? ` style="${tableStyle.join(";")}"` : "";
    let html = `<table${styleAttr}>`;
    for (const row of rows) {
      html += "<tr>";
      for (const cell of row) html += cell;
      html += "</tr>";
    }
    html += "</table>";
    return html;
  }

  _convertCell(node, isHeader, colWidth, cellPadding) {
    const tcPr = findChild(node, "tcPr");
    const cellStyle = ["border:1px solid #000", "padding:6px 8px", "vertical-align:top"];
    if (colWidth) cellStyle.push(`width:${colWidth}pt`);
    if (cellPadding) cellStyle.push(`padding:${this._twipsToPt(cellPadding)}pt`);

    if (tcPr) {
      const vAlign = findChild(tcPr, "vAlign");
      if (vAlign) {
        const val = wAttr(vAlign, "val");
        if (val === "center") cellStyle.push("vertical-align:middle");
        else if (val === "bottom") cellStyle.push("vertical-align:bottom");
      }
      const shd = findChild(tcPr, "shd");
      if (shd) {
        const fill = wAttr(shd, "fill");
        if (fill && fill !== "auto") cellStyle.push(`background-color:#${fill}`);
      }
      const tcBorders = findChild(tcPr, "tcBorders");
      if (tcBorders) {
        const borders = this._parseBorders(tcBorders);
        if (borders) cellStyle.push(borders);
      }
      const gridSpan = findChild(tcPr, "gridSpan");
      if (gridSpan) {
        const span = parseInt(wAttr(gridSpan, "val") || "1", 10);
        if (span > 1) cellStyle.push(`colspan:${span}`);
      }
    }

    const tag = isHeader ? "th" : "td";
    const styleAttr = ` style="${cellStyle.join(";")}"`;
    const paras = [];
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (isElem(child, "p")) paras.push(this._convertPara(child));
      else if (isElem(child, "tbl")) paras.push(this._convertTable(child));
    }
    if (!paras.length) return `<${tag}${styleAttr}><br></${tag}>`;
    return `<${tag}${styleAttr}>${paras.join("\n")}</${tag}>`;
  }

  _parseBorders(borderNode) {
    if (!borderNode) return null;
    const styles = [];
    const sides = ["top", "left", "bottom", "right"];
    for (const side of sides) {
      const el = findChild(borderNode, side);
      if (!el) continue;
      const val = wAttr(el, "val");
      if (!val || val === "none" || val === "nil") continue;
      const sz = wAttr(el, "sz") || "4";
      const color = wAttr(el, "color") || "000000";
      const w = Math.round(parseInt(sz, 10) / 8);
      styles.push(`border-${side}:${w}px solid #${color}`);
    }
    return styles.length ? styles.join(";") : null;
  }

  _stripHtml(html) {
    return html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  }

  _esc(text) {
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
}

async function parseDocx(buffer) {
  const parser = new DocxFormatParser(buffer);
  const result = await parser.parse();
  return result;
}

module.exports = { DocxFormatParser, parseDocx };

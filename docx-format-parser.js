const JSZip = require("jszip");

function parseXml(xml) {
  const DOMParser = require("xmldom").DOMParser || null;
  if (DOMParser) return new DOMParser().parseFromString(xml, "text/xml");
  return null;
}

function xmlDecode(s) {
  if (!s) return "";
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}

function extractAttrValue(xml, tagName, attrName) {
  const re = new RegExp(`<${tagName}[^>]*\\s${attrName}="([^"]*)"`, "i");
  const m = xml.match(re);
  return m ? m[1] : null;
}

function getElementsByTagName(xml, tagName) {
  const results = [];
  const re = new RegExp(`<${tagName}(?:\\s[^>]*)?>[\\s\\S]*?<\\/${tagName}>|<${tagName}\\s[^>]*\\/?>`, "gi");
  let m;
  while ((m = re.exec(xml)) !== null) {
    results.push(m[0]);
  }
  return results;
}

function getAttr(el, name) {
  const re = new RegExp(`${name}="([^"]*)"`, "i");
  const m = el.match(re);
  return m ? m[1] : null;
}

function resolveColor(val, themeColors) {
  if (!val || val === "auto") return null;
  if (val.startsWith("#")) return val;
  if (themeColors[val]) return themeColors[val];
  return null;
}

class DocxFormatParser {
  constructor(buffer) {
    this.buffer = buffer;
    this.zip = null;
    this.styles = {};
    this.numbering = {};
    this.themeColors = {};
    this.relImages = {};
    this.docXml = "";
    this.bodyXml = "";
  }

  async parse() {
    this.zip = await JSZip.loadAsync(this.buffer);
    await this._parseTheme();
    await this._parseStyles();
    await this._parseNumbering();
    await this._parseRels();
    this.docXml = await this._readFile("word/document.xml");
    this.bodyXml = this._extractTag(this.docXml, "w:body");
    const docBgEl = getElementsByTagName(this.docXml, "w:background")[0];
    let docBgStyle = "";
    if (docBgEl) {
      const bgFill = getAttr(docBgEl, "w:color");
      if (bgFill && bgFill !== "auto") docBgStyle = `background-color:#${bgFill}`;
    }
    const html = await this._convertBodyToHtml();
    const wrappedHtml = docBgStyle ? `<div style="${docBgStyle}">${html}</div>` : html;
    return { html: wrappedHtml, text: this._stripHtml(wrappedHtml) };
  }

  async _readFile(name) {
    const file = this.zip.file(name);
    if (!file) return "";
    return file.async("string");
  }

  _extractTag(xml, tag) {
    const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, "i");
    const m = xml.match(re);
    return m ? m[1] : "";
  }

  async _parseTheme() {
    const xml = await this._readFile("word/theme/theme1.xml");
    if (!xml) return;
    const clrScheme = this._extractTag(xml, "a:clrScheme");
    if (!clrScheme) return;
    const colorMap = {
      "a:dk1": "dk1", "a:lt1": "lt1", "a:dk2": "dk2", "a:lt2": "lt2",
      "a:accent1": "accent1", "a:accent2": "accent2", "a:accent3": "accent3",
      "a:accent4": "accent4", "a:accent5": "accent5", "a:accent6": "accent6",
      "a:hlink": "hlink", "a:folHlink": "folHlink",
    };
    for (const [tag, key] of Object.entries(colorMap)) {
      const section = this._extractTag(clrScheme, tag);
      const srgb = getAttr(section, "a:srgbClr") || getAttr(section, "val");
      if (srgb && /^[0-9a-fA-F]{6}$/.test(srgb)) {
        this.themeColors[key] = `#${srgb}`;
      }
    }
  }

  async _parseStyles() {
    const xml = await this._readFile("word/styles.xml");
    if (!xml) return;
    const styleEls = getElementsByTagName(xml, "w:style");
    for (const el of styleEls) {
      const id = getAttr(el, "w:styleId");
      if (!id) continue;
      this.styles[id] = this._parseStyleProps(el);
    }
    const docDefaults = this._extractTag(xml, "w:docDefaults");
    if (docDefaults) {
      this.docDefaultStyle = this._parseStyleProps(docDefaults);
    }
  }

  _parseStyleProps(el) {
    const props = {};
    const rPr = this._extractTag(el, "w:rPr");
    if (rPr) {
      const sz = getAttr(rPr, "w:val", "w:sz");
      const szEl = getElementsByTagName(rPr, "w:sz")[0];
      if (szEl) props.sz = getAttr(szEl, "w:val");
      const szCsEl = getElementsByTagName(rPr, "w:szCs")[0];
      if (szCsEl) props.szCs = getAttr(szCsEl, "w:val");

      const rFonts = getElementsByTagName(rPr, "w:rFonts")[0];
      if (rFonts) {
        props.ascii = getAttr(rFonts, "w:ascii") || getAttr(rFonts, "w:eastAsia") || getAttr(rFonts, "w:hAnsi");
        props.eastAsia = getAttr(rFonts, "w:eastAsia") || getAttr(rFonts, "w:ascii");
        props.hAnsi = getAttr(rFonts, "w:hAnsi") || getAttr(rFonts, "w:ascii");
      }

      const color = getElementsByTagName(rPr, "w:color")[0];
      if (color) props.color = getAttr(color, "w:val");

      const bold = getElementsByTagName(rPr, "w:b")[0];
      if (bold) {
        const val = getAttr(bold, "w:val");
        props.bold = val !== "0" && val !== "false";
      }

      const italic = getElementsByTagName(rPr, "w:i")[0];
      if (italic) {
        const val = getAttr(italic, "w:val");
        props.italic = val !== "0" && val !== "false";
      }

      const underline = getElementsByTagName(rPr, "w:u")[0];
      if (underline) {
        const val = getAttr(underline, "w:val");
        if (val && val !== "none") props.underline = val;
      }

      const strike = getElementsByTagName(rPr, "w:strike")[0];
      if (strike) {
        const val = getAttr(strike, "w:val");
        props.strike = val !== "0" && val !== "false";
      }

      const shd = getElementsByTagName(rPr, "w:shd")[0];
      if (shd) {
        const fill = getAttr(shd, "w:fill");
        if (fill && fill !== "auto") props.highlight = `#${fill}`;
      }

      const vertAlign = getElementsByTagName(rPr, "w:vertAlign")[0];
      if (vertAlign) props.vertAlign = getAttr(vertAlign, "w:val");
    }

    const pPr = this._extractTag(el, "w:pPr");
    if (pPr) {
      const jc = getElementsByTagName(pPr, "w:jc")[0];
      if (jc) props.jc = getAttr(jc, "w:val");

      const ind = getElementsByTagName(pPr, "w:ind")[0];
      if (ind) {
        props.indLeft = getAttr(ind, "w:left") || getAttr(ind, "w:start");
        props.indRight = getAttr(ind, "w:right") || getAttr(ind, "w:end");
        props.indFirstLine = getAttr(ind, "w:firstLine") || getAttr(ind, "w:firstLineChars");
        props.indHanging = getAttr(ind, "w:hanging");
      }

      const spacing = getElementsByTagName(pPr, "w:spacing")[0];
      if (spacing) {
        props.spBefore = getAttr(spacing, "w:before");
        props.spAfter = getAttr(spacing, "w:after");
        props.lineSpacing = getAttr(spacing, "w:line");
        props.lineRule = getAttr(spacing, "w:lineRule");
      }

      const pBdr = this._extractTag(pPr, "w:pBdr");
      if (pBdr) {
        const bottom = getElementsByTagName(pBdr, "w:bottom")[0];
        if (bottom) {
          const val = getAttr(bottom, "w:val");
          if (val && val !== "none") {
            props.pBorderBottom = {
              style: val,
              sz: getAttr(bottom, "w:sz") || "4",
              color: getAttr(bottom, "w:color") || "000000",
            };
          }
        }
      }

      const pShd = getElementsByTagName(pPr, "w:shd")[0];
      if (pShd) {
        const fill = getAttr(pShd, "w:fill");
        if (fill && fill !== "auto") props.pBackground = `#${fill}`;
      }

      const numPr = this._extractTag(pPr, "w:numPr");
      if (numPr) {
        const ilvlEl = getElementsByTagName(numPr, "w:ilvl")[0];
        const numIdEl = getElementsByTagName(numPr, "w:numId")[0];
        if (ilvlEl) props.ilvl = getAttr(ilvlEl, "w:val");
        if (numIdEl) props.numId = getAttr(numIdEl, "w:val");
      }
    }

    return props;
  }

  async _parseNumbering() {
    const xml = await this._readFile("word/numbering.xml");
    if (!xml) return;
    const numEls = getElementsByTagName(xml, "w:num");
    for (const el of numEls) {
      const numId = getAttr(el, "w:numId");
      if (!numId) continue;
      const abstractIdEl = getElementsByTagName(el, "w:abstractNumId")[0];
      const abstractId = abstractIdEl ? getAttr(abstractIdEl, "w:val") : null;
      this.numbering[numId] = { abstractId, levels: {} };
      const lvlEls = getElementsByTagName(el, "w:lvl");
      for (const lvl of lvlEls) {
        const ilvl = getAttr(lvl, "w:ilvl");
        if (ilvl == null) continue;
        const fmtEl = getElementsByTagName(lvl, "w:numFmt")[0];
        const lvlTextEl = getElementsByTagName(lvl, "w:lvlText")[0];
        const startEl = getElementsByTagName(lvl, "w:start")[0];
        this.numbering[numId].levels[ilvl] = {
          fmt: fmtEl ? getAttr(fmtEl, "w:val") : "decimal",
          text: lvlTextEl ? getAttr(lvlTextEl, "w:val") : null,
          start: startEl ? parseInt(getAttr(startEl, "w:val") || "1", 10) : 1,
        };
      }
    }
    const abstractNumEls = getElementsByTagName(xml, "w:abstractNum");
    for (const el of abstractNumEls) {
      const id = getAttr(el, "w:abstractNumId");
      if (!id) continue;
      const lvlEls = getElementsByTagName(el, "w:lvl");
      for (const num of Object.values(this.numbering)) {
        if (num.abstractId === id) {
          for (const lvl of lvlEls) {
            const ilvl = getAttr(lvl, "w:ilvl");
            if (ilvl == null) continue;
            if (!num.levels[ilvl]) {
              const fmtEl = getElementsByTagName(lvl, "w:numFmt")[0];
              const lvlTextEl = getElementsByTagName(lvl, "w:lvlText")[0];
              const startEl = getElementsByTagName(lvl, "w:start")[0];
              num.levels[ilvl] = {
                fmt: fmtEl ? getAttr(fmtEl, "w:val") : "decimal",
                text: lvlTextEl ? getAttr(lvlTextEl, "w:val") : null,
                start: startEl ? parseInt(getAttr(startEl, "w:val") || "1", 10) : 1,
              };
            }
          }
        }
      }
    }
  }

  async _parseRels() {
    const xml = await this._readFile("word/_rels/document.xml.rels");
    if (!xml) return;
    const relEls = getElementsByTagName(xml, "Relationship");
    for (const el of relEls) {
      const id = getAttr(el, "Id");
      const target = getAttr(el, "Target");
      const type = getAttr(el, "Type");
      if (id && target && type && type.includes("image")) {
        this.relImages[id] = target.replace(/^\/?/, "");
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
    const mime = mimeMap[ext] || "image/png";
    return `data:${mime};base64,${data}`;
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

  _twipsToPt(val) {
    if (!val) return null;
    return (parseInt(val, 10) / 20).toFixed(1);
  }

  _twipsToEm(val) {
    if (!val) return null;
    return (parseInt(val, 10) / 240).toFixed(3);
  }

  _halfPointsToPt(val) {
    if (!val) return null;
    return (parseInt(val, 10) / 2).toFixed(1);
  }

  _ptToFontSize(sz) {
    if (!sz) return null;
    const pt = this._halfPointsToPt(sz);
    return pt ? `${pt}pt` : null;
  }

  async _convertBodyToHtml() {
    const parts = [];
    const children = this.bodyXml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/gi) || [];
    this._listCounters = {};
    for (const child of children) {
      if (child.startsWith("<w:tbl")) {
        parts.push(await this._convertTable(child));
      } else {
        parts.push(await this._convertPara(child));
      }
    }
    return parts.join("\n");
  }

  async _convertPara(xml) {
    const pPr = this._extractTag(xml, "w:pPr");
    const styleId = pPr ? getAttr(getElementsByTagName(pPr, "w:pStyle")[0], "w:val") : null;
    const resolved = this._resolveStyle(styleId);
    const merged = { ...resolved };

    if (pPr) {
      const jcEl = getElementsByTagName(pPr, "w:jc")[0];
      if (jcEl) merged.jc = getAttr(jcEl, "w:val");

      const ind = getElementsByTagName(pPr, "w:ind")[0];
      if (ind) {
        merged.indLeft = getAttr(ind, "w:left") || getAttr(ind, "w:start") || merged.indLeft;
        merged.indRight = getAttr(ind, "w:right") || getAttr(ind, "w:end") || merged.indRight;
        merged.indFirstLine = getAttr(ind, "w:firstLine") || getAttr(ind, "w:firstLineChars") || merged.indFirstLine;
        merged.indHanging = getAttr(ind, "w:hanging") || merged.indHanging;
      }

      const spacing = getElementsByTagName(pPr, "w:spacing")[0];
      if (spacing) {
        merged.spBefore = getAttr(spacing, "w:before") || merged.spBefore;
        merged.spAfter = getAttr(spacing, "w:after") || merged.spAfter;
        merged.lineSpacing = getAttr(spacing, "w:line") || merged.lineSpacing;
        merged.lineRule = getAttr(spacing, "w:lineRule") || merged.lineRule;
      }

      const pShd = getElementsByTagName(pPr, "w:shd")[0];
      if (pShd) {
        const fill = getAttr(pShd, "w:fill");
        if (fill && fill !== "auto") merged.pBackground = `#${fill}`;
      }

      const numPr = this._extractTag(pPr, "w:numPr");
      if (numPr) {
        const ilvlEl = getElementsByTagName(numPr, "w:ilvl")[0];
        const numIdEl = getElementsByTagName(numPr, "w:numId")[0];
        if (ilvlEl) merged.ilvl = getAttr(ilvlEl, "w:val") || "0";
        if (numIdEl) merged.numId = getAttr(numIdEl, "w:val");
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
          prefix = `<span style="color:inherit;white-space:pre">${this._escapeHtml(numText)} </span>`;
        }
      }
    }

    const styleAttr = pStyle.length ? ` style="${pStyle.join(";")}"` : "";
    const runs = await this._convertRuns(xml);
    const tag = styleId && /^heading/i.test(styleId) ? (styleId === "Heading1" ? "h1" : styleId === "Heading2" ? "h2" : styleId === "Heading3" ? "h3" : "h4") : "p";

    if (!runs.trim() && !prefix) return `<${tag}${styleAttr}><br></${tag}>`;
    return `<${tag}${styleAttr}>${prefix}${runs}</${tag}>`;
  }

  async _convertRuns(paraXml) {
    const results = [];
    const runRe = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/gi;
    let m;
    while ((m = runRe.exec(paraXml)) !== null) {
      results.push(await this._convertRun(m[0]));
    }
    const hyperlinkRe = /<w:hyperlink(?:\s[^>]*)?>[\s\S]*?<\/w:hyperlink>/gi;
    while ((m = hyperlinkRe.exec(paraXml)) !== null) {
      const runs = [];
      let rm;
      const runRe2 = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/gi;
      while ((rm = runRe2.exec(m[0])) !== null) {
        runs.push(await this._convertRun(rm[0]));
      }
      const rId = getAttr(m[0], "r:id");
      results.push(`<a>${runs.join("")}</a>`);
    }
    return results.join("");
  }

  async _convertRun(xml) {
    const rPr = this._extractTag(xml, "w:rPr");
    const style = [];

    let isBold = false;
    let isItalic = false;
    let isUnderline = false;
    let isStrike = false;
    let color = null;
    let highlight = null;
    let fontSize = null;
    let fontFamily = null;
    let vertAlign = null;

    const resolvedRpr = { ...this._resolveStyle(rPr ? getAttr(getElementsByTagName(rPr, "w:rStyle")[0], "w:val") : null) };
    if (rPr) {
      const boldEl = getElementsByTagName(rPr, "w:b")[0];
      if (boldEl) { const v = getAttr(boldEl, "w:val"); isBold = v !== "0" && v !== "false"; }
      else if (resolvedRpr.bold) isBold = true;

      const italicEl = getElementsByTagName(rPr, "w:i")[0];
      if (italicEl) { const v = getAttr(italicEl, "w:val"); isItalic = v !== "0" && v !== "false"; }
      else if (resolvedRpr.italic) isItalic = true;

      const underlineEl = getElementsByTagName(rPr, "w:u")[0];
      if (underlineEl) { const v = getAttr(underlineEl, "w:val"); isUnderline = v && v !== "none"; }
      else if (resolvedRpr.underline) isUnderline = true;

      const strikeEl = getElementsByTagName(rPr, "w:strike")[0];
      if (strikeEl) { const v = getAttr(strikeEl, "w:val"); isStrike = v !== "0" && v !== "false"; }
      else if (resolvedRpr.strike) isStrike = true;

      const colorEl = getElementsByTagName(rPr, "w:color")[0];
      color = colorEl ? getAttr(colorEl, "w:val") : resolvedRpr.color;
      if (color === "auto") color = null;

      const shdEl = getElementsByTagName(rPr, "w:shd")[0];
      if (shdEl) {
        const fill = getAttr(shdEl, "w:fill");
        if (fill && fill !== "auto") highlight = `#${fill}`;
      }
      if (!highlight && resolvedRpr.highlight) highlight = resolvedRpr.highlight;

      const szEl = getElementsByTagName(rPr, "w:sz")[0];
      fontSize = szEl ? getAttr(szEl, "w:val") : resolvedRpr.sz;

      const rFonts = getElementsByTagName(rPr, "w:rFonts")[0];
      if (rFonts) fontFamily = getAttr(rFonts, "w:ascii") || getAttr(rFonts, "w:eastAsia") || getAttr(rFonts, "w:hAnsi");
      else fontFamily = resolvedRpr.ascii;

      const vertAlignEl = getElementsByTagName(rPr, "w:vertAlign")[0];
      if (vertAlignEl) vertAlign = getAttr(vertAlignEl, "w:val");
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
    const drawingRe = /<w:drawing[\s\S]*?<\/w:drawing>/gi;
    const drawingMatch = xml.match(drawingRe);
    if (drawingMatch) {
      for (const drawing of drawingMatch) {
        const embed = getAttr(drawing, "r:embed");
        if (embed) {
          const dataUri = await this._getImageDataUri(embed);
          if (dataUri) content += `<img src="${dataUri}" style="max-width:100%;height:auto" />`;
        }
      }
    }

    const pictRe = /<w:pict[\s\S]*?<\/w:pict>/gi;
    const pictMatch = xml.match(pictRe);
    if (pictMatch) {
      for (const pict of pictMatch) {
        const embed = getAttr(pict, "r:embed");
        if (embed) {
          const dataUri = await this._getImageDataUri(embed);
          if (dataUri) content += `<img src="${dataUri}" style="max-width:100%;height:auto" />`;
        }
      }
    }

    const textRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/gi;
    let tm;
    while ((tm = textRe.exec(xml)) !== null) {
      content += this._escapeHtml(xmlDecode(tm[1]));
    }

    const tabRe = /<w:tab\s*\/?>/gi;
    if (tabRe.test(xml)) content += "&nbsp;&nbsp;&nbsp;&nbsp;";

    const brRe = /<w:br\s*\/?>/gi;
    if (brRe.test(xml)) content += "<br>";

    if (!content) return "";
    return `<span${styleAttr}>${content}</span>`;
  }

  async _convertTable(xml) {
    const tblPr = this._extractTag(xml, "w:tblPr");
    const tableStyle = [];

    const tblW = getElementsByTagName(tblPr || "", "w:tblW")[0];
    const tblWType = tblW ? getAttr(tblW, "w:type") : null;
    const tblWVal = tblW ? getAttr(tblW, "w:w") : null;
    if (tblWType === "pct") tableStyle.push("width:100%");
    else if (tblWVal) tableStyle.push(`width:${this._twipsToPt(tblWVal)}pt`);

    const jc = getElementsByTagName(tblPr || "", "w:jc")[0];
    if (jc) {
      const val = getAttr(jc, "w:val");
      if (val === "center") tableStyle.push("margin:0 auto");
      else if (val === "right") tableStyle.push("margin:0 0 0 auto");
    }

    const tblBorders = this._extractTag(tblPr || "", "w:tblBorders");
    const borders = this._parseBorders(tblBorders);
    if (borders) tableStyle.push(borders);

    const tblShd = getElementsByTagName(tblPr || "", "w:shd")[0];
    if (tblShd) {
      const fill = getAttr(tblShd, "w:fill");
      if (fill && fill !== "auto") tableStyle.push(`background-color:#${fill}`);
    }

    const tblCellMar = this._extractTag(tblPr || "", "w:tblCellMar");
    let cellPadding = null;
    if (tblCellMar) {
      const top = getElementsByTagName(tblCellMar, "w:top")[0];
      if (top) cellPadding = getAttr(top, "w:w");
    }

    const colWidths = [];
    const tblGrid = this._extractTag(xml, "w:tblGrid");
    if (tblGrid) {
      const gridCols = getElementsByTagName(tblGrid, "w:gridCol");
      for (const col of gridCols) {
        const w = getAttr(col, "w:w");
        colWidths.push(w ? this._twipsToPt(w) : null);
      }
    }

    const rowEls = getElementsByTagName(xml, "w:tr");
    const rows = [];
    for (const rowEl of rowEls) {
      const isHeader = /<w:tblHeader\s*\/?>/.test(this._extractTag(rowEl, "w:trPr") || "");
      const cellEls = getElementsByTagName(rowEl, "w:tc");
      const cells = [];
      let cellIdx = 0;
      for (const cellEl of cellEls) {
        const cellResult = await this._convertCell(cellEl, isHeader, colWidths[cellIdx], cellPadding);
        cells.push(cellResult);
        const gridSpanEl = getElementsByTagName(cellEl, "w:gridSpan")[0];
        if (gridSpanEl) {
          cellIdx += parseInt(getAttr(gridSpanEl, "w:val") || "1", 10);
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

  async _convertCell(cellEl, isHeader, colWidth, cellPadding) {
    const tcPr = this._extractTag(cellEl, "w:tcPr");
    const cellStyle = ["border:1px solid #000", "padding:6px 8px", "vertical-align:top"];

    if (colWidth) cellStyle.push(`width:${colWidth}pt`);
    if (cellPadding) cellStyle.push(`padding:${this._twipsToPt(cellPadding)}pt`);

    const vAlign = getElementsByTagName(tcPr || "", "w:vAlign")[0];
    if (vAlign) {
      const val = getAttr(vAlign, "w:val");
      if (val === "center") cellStyle.push("vertical-align:middle");
      else if (val === "bottom") cellStyle.push("vertical-align:bottom");
    }

    const shd = getElementsByTagName(tcPr || "", "w:shd")[0];
    if (shd) {
      const fill = getAttr(shd, "w:fill");
      if (fill && fill !== "auto") cellStyle.push(`background-color:#${fill}`);
    }

    const tcBorders = this._extractTag(tcPr || "", "w:tcBorders");
    const borders = this._parseBorders(tcBorders);
    if (borders) cellStyle.push(borders);

    const tag = isHeader ? "th" : "td";
    const styleAttr = ` style="${cellStyle.join(";")}"`;

    const paras = [];
    const childParas = cellEl.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/gi) || [];
    for (const p of childParas) {
      paras.push(await this._convertPara(p));
    }

    if (!paras.length) return `<${tag}${styleAttr}><br></${tag}>`;
    return `<${tag}${styleAttr}>${paras.join("\n")}</${tag}>`;
  }

  _parseBorders(borderXml) {
    if (!borderXml) return null;
    const styles = [];
    const sides = ["top", "left", "bottom", "right"];
    for (const side of sides) {
      const el = getElementsByTagName(borderXml, `w:${side}`)[0];
      if (!el) continue;
      const val = getAttr(el, "w:val");
      if (!val || val === "none" || val === "nil") continue;
      const sz = getAttr(el, "w:sz") || "4";
      const color = getAttr(el, "w:color") || "000000";
      const w = Math.round(parseInt(sz, 10) / 8);
      styles.push(`border-${side}:${w}px solid #${color}`);
    }
    return styles.length ? styles.join(";") : null;
  }

  _stripHtml(html) {
    return html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  }

  _escapeHtml(text) {
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
}

module.exports = { DocxFormatParser };

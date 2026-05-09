const JSZip = require("jszip");

function esc(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

class DocxFormatParser {
  constructor(buffer) {
    this.buffer = buffer;
    this.zip = null;
  }

  async parse() {
    this.zip = await JSZip.loadAsync(this.buffer);
    const docXml = await this._readFile("word/document.xml");
    const stylesXml = await this._readFile("word/styles.xml");

    const defaultFont = this._extractDefaultFont(stylesXml);
    const defaultSize = this._extractDefaultSize(stylesXml);

    const html = this._buildHtml(docXml, defaultFont, defaultSize);
    const text = html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    return { html, text };
  }

  async _readFile(name) {
    const file = this.zip.file(name);
    if (!file) return null;
    return file.async("string");
  }

  _extractDefaultFont(stylesXml) {
    if (!stylesXml) return null;
    const m = stylesXml.match(/<w:docDefaults>[\s\S]*?<w:rFonts[^>]*w:ascii="([^"]+)"/);
    return m ? m[1] : null;
  }

  _extractDefaultSize(stylesXml) {
    if (!stylesXml) return null;
    const m = stylesXml.match(/<w:docDefaults>[\s\S]*?<w:sz[^>]*w:val="(\d+)"/);
    return m ? (parseInt(m[1], 10) / 2).toFixed(1) : null;
  }

  _twipsToEm(val) { return val ? (parseInt(val, 10) / 240).toFixed(3) : null; }
  _twipsToPt(val) { return val ? (parseInt(val, 10) / 20).toFixed(1) : null; }
  _halfToPt(val) { return val ? (parseInt(val, 10) / 2).toFixed(1) : null; }

  _buildHtml(docXml, defaultFont, defaultSize) {
    if (!docXml) return "";
    const bodyMatch = docXml.match(/<w:body>([\s\S]*?)<\/w:body>/);
    if (!bodyMatch) return "";
    const bodyXml = bodyMatch[1];

    const bgMatch = docXml.match(/<w:background[^>]*w:color="([^"]+)"/);
    const docBg = bgMatch && bgMatch[1] !== "auto" ? `background-color:#${bgMatch[1]}` : "";

    const parts = [];
    const elRe = /<w:(p|tbl|sdt)\b/g;
    let m;
    const usedRanges = [];

    while ((m = elRe.exec(bodyXml)) !== null) {
      const tag = m[1];
      const start = m.index;

      let skip = false;
      for (const range of usedRanges) {
        if (start >= range[0] && start < range[1]) { skip = true; break; }
      }
      if (skip) continue;

      const endTag = `</w:${tag}>`;
      const endIdx = bodyXml.indexOf(endTag, start);
      if (endIdx === -1) continue;
      const xml = bodyXml.slice(start, endIdx + endTag.length);
      usedRanges.push([start, endIdx + endTag.length]);

      if (tag === "p") {
        parts.push(this._convertPara(xml));
      } else if (tag === "tbl") {
        parts.push(this._convertTable(xml));
      } else if (tag === "sdt") {
        const contentMatch = xml.match(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/);
        if (contentMatch) {
          const inner = contentMatch[1];
          let innerM;
          const innerRe = /<w:(p|tbl)\b/g;
          while ((innerM = innerRe.exec(inner)) !== null) {
            const innerTag = innerM[1];
            const innerStart = innerM.index;
            const innerEndTag = `</w:${innerTag}>`;
            const innerEndIdx = inner.indexOf(innerEndTag, innerStart);
            if (innerEndIdx === -1) continue;
            const innerXml = inner.slice(innerStart, innerEndIdx + innerEndTag.length);
            if (innerTag === "p") parts.push(this._convertPara(innerXml));
            else if (innerTag === "tbl") parts.push(this._convertTable(innerXml));
          }
        }
      }
    }

    let html = parts.join("\n");
    if (defaultFont || defaultSize) {
      const s = [];
      if (defaultFont) s.push(`font-family:"${defaultFont}",serif`);
      if (defaultSize) s.push(`font-size:${defaultSize}pt`);
      html = `<div style="${s.join(";")}">${html}</div>`;
    }
    if (docBg) {
      html = `<div style="${docBg}">${html}</div>`;
    }
    return html;
  }

  _convertPara(paraXml) {
    const pPr = paraXml.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
    const style = [];
    let isHeading = false;
    let headingLevel = "p";

    if (pPr) {
      const pPrXml = pPr[1];
      const styleM = pPrXml.match(/<w:pStyle[^>]*w:val="([^"]+)"/);
      if (styleM) {
        const styleName = styleM[1];
        if (/^Heading(\d)/i.test(styleName)) {
          isHeading = true;
          headingLevel = `h${Math.min(parseInt(RegExp.$1, 10), 6)}`;
        }
      }

      const jcM = pPrXml.match(/<w:jc[^>]*w:val="([^"]+)"/);
      if (jcM) {
        const v = jcM[1];
        if (v === "center") style.push("text-align:center");
        else if (v === "right") style.push("text-align:right");
        else if (v === "both" || v === "distribute") style.push("text-align:justify");
      }

      const indLeftM = pPrXml.match(/<w:ind[^>]*(?:w:left|w:start)="(\d+)"/);
      if (indLeftM) style.push(`margin-left:${this._twipsToEm(indLeftM[1])}em`);
      const indFirstM = pPrXml.match(/<w:ind[^>]*(?:w:firstLine|w:firstLineChars)="(\d+)"/);
      if (indFirstM) style.push(`text-indent:${this._twipsToEm(indFirstM[1])}em`);

      const spBeforeM = pPrXml.match(/<w:spacing[^>]*w:before="(\d+)"/);
      if (spBeforeM) style.push(`margin-top:${this._twipsToPt(spBeforeM[1])}pt`);
      const spAfterM = pPrXml.match(/<w:spacing[^>]*w:after="(\d+)"/);
      if (spAfterM) style.push(`margin-bottom:${this._twipsToPt(spAfterM[1])}pt`);

      const lineM = pPrXml.match(/<w:spacing[^>]*w:line="(\d+)"/);
      const lineRuleM = pPrXml.match(/<w:spacing[^>]*w:lineRule="(\w+)"/);
      if (lineM) {
        if (lineRuleM && lineRuleM[1] === "exact") {
          style.push(`line-height:${this._twipsToPt(lineM[1])}pt`);
        } else {
          style.push(`line-height:${this._twipsToEm(lineM[1])}`);
        }
      }

      const shdM = pPrXml.match(/<w:shd[^>]*w:fill="([^"]+)"/);
      if (shdM && shdM[1] !== "auto") style.push(`background-color:#${shdM[1]}`);

      const borderM = pPrXml.match(/<w:pBdr>[\s\S]*?<w:bottom[^>]*w:val="([^"]*)"[^>]*(?:w:sz="([^"]*)")?[^>]*(?:w:color="([^"]*)")?/);
      if (borderM && borderM[1] !== "none" && borderM[1] !== "nil") {
        const w = Math.round(parseInt(borderM[2] || "4", 10) / 8);
        const c = borderM[3] || "000000";
        style.push(`border-bottom:${w}px solid #${c}`);
      }
    }

    const runs = this._convertRuns(paraXml);
    const tag = isHeading ? headingLevel : "p";
    const styleAttr = style.length ? ` style="${style.join(";")}"` : "";

    if (!runs.trim()) return `<${tag}${styleAttr}><br></${tag}>`;
    return `<${tag}${styleAttr}>${runs}</${tag}>`;
  }

  _convertRuns(parentXml) {
    const parts = [];
    let pos = 0;
    const content = parentXml;

    while (pos < content.length) {
      const runIdx = content.indexOf("<w:r>", pos);
      const hyperIdx = content.indexOf("<w:hyperlink", pos);
      const sdtIdx = content.indexOf("<w:sdt>", pos);

      let nextIdx = -1;
      let nextType = "";
      if (runIdx !== -1) { nextIdx = runIdx; nextType = "run"; }
      if (hyperIdx !== -1 && (nextIdx === -1 || hyperIdx < nextIdx)) { nextIdx = hyperIdx; nextType = "hyperlink"; }
      if (sdtIdx !== -1 && (nextIdx === -1 || sdtIdx < nextIdx)) { nextIdx = sdtIdx; nextType = "sdt"; }

      if (nextIdx === -1) break;

      if (nextType === "run") {
        const endIdx = content.indexOf("</w:r>", nextIdx);
        if (endIdx === -1) break;
        const runXml = content.slice(nextIdx, endIdx + 6);
        parts.push(this._convertRun(runXml));
        pos = endIdx + 6;
      } else if (nextType === "hyperlink") {
        const endIdx = this._findClosingTag(content, nextIdx, "w:hyperlink");
        if (endIdx === -1) break;
        const hyperXml = content.slice(nextIdx, endIdx);
        parts.push(`<a>${this._convertRuns(hyperXml)}</a>`);
        pos = endIdx;
      } else if (nextType === "sdt") {
        const endIdx = this._findClosingTag(content, nextIdx, "w:sdt");
        if (endIdx === -1) break;
        const sdtXml = content.slice(nextIdx, endIdx);
        const contentM = sdtXml.match(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/);
        if (contentM) parts.push(this._convertRuns(contentM[1]));
        pos = endIdx;
      }
    }

    return parts.join("");
  }

  _findClosingTag(xml, start, tagName) {
    const openTag = `<${tagName}`;
    const closeTag = `</${tagName}>`;
    let depth = 0;
    let i = start;
    while (i < xml.length) {
      if (xml.slice(i, i + openTag.length) === openTag &&
          (xml[i + openTag.length] === '>' || xml[i + openTag.length] === ' ' || xml[i + openTag.length] === '/')) {
        depth++;
        const selfClose = xml.indexOf("/>", i);
        const nextOpen = xml.indexOf(">", i);
        if (nextOpen !== -1) i = nextOpen + 1;
        else break;
      } else if (xml.slice(i, i + closeTag.length) === closeTag) {
        depth--;
        if (depth === 0) return i + closeTag.length;
        i += closeTag.length;
      } else {
        i++;
      }
    }
    return -1;
  }

  _convertRun(runXml) {
    const rPrM = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    const style = [];

    if (rPrM) {
      const rPr = rPrM[1];
      if ((rPr.includes("<w:b/>") || rPr.match(/<w:b\s/)) && !rPr.match(/<w:b[^>]*w:val="(0|false)"/)) {
        style.push("font-weight:bold");
      }

      if ((rPr.includes("<w:i/>") || rPr.match(/<w:i\s/)) && !rPr.match(/<w:i[^>]*w:val="(0|false)"/)) {
        style.push("font-style:italic");
      }

      if ((rPr.includes("<w:u/>") || rPr.match(/<w:u\s/)) && !rPr.match(/<w:u[^>]*w:val="none"/)) {
        style.push("text-decoration:underline");
      }

      if ((rPr.includes("<w:strike/>") || rPr.match(/<w:strike\s/)) && !rPr.match(/<w:strike[^>]*w:val="(0|false)"/)) {
        style.push("text-decoration:line-through");
      }

      const colorM = rPr.match(/<w:color[^>]*w:val="([^"]+)"/);
      if (colorM && colorM[1] !== "auto") style.push(`color:#${colorM[1]}`);

      const shdM = rPr.match(/<w:shd[^>]*w:fill="([^"]+)"/);
      if (shdM && shdM[1] !== "auto") style.push(`background-color:#${shdM[1]}`);

      const szM = rPr.match(/<w:sz[^>]*w:val="(\d+)"/);
      if (szM) style.push(`font-size:${this._halfToPt(szM[1])}pt`);

      const fontM = rPr.match(/<w:rFonts[^>]*(?:w:ascii|w:eastAsia|w:hAnsi)="([^"]+)"/);
      if (fontM) style.push(`font-family:"${fontM[1]}",serif`);

      const vertM = rPr.match(/<w:vertAlign[^>]*w:val="(\w+)"/);
      if (vertM) {
        if (vertM[1] === "superscript") style.push("vertical-align:super;font-size:0.8em");
        else if (vertM[1] === "subscript") style.push("vertical-align:sub;font-size:0.8em");
      }
    }

    const content = this._extractRunContent(runXml);
    if (!content) return "";

    const styleAttr = style.length ? ` style="${style.join(";")}"` : "";
    return `<span${styleAttr}>${content}</span>`;
  }

  _extractRunContent(runXml) {
    const parts = [];
    const childRe = /<w:(t|br|tab|cr|sym|drawing|noBreakHyphen)(\s[^>]*)?\/?>/g;
    let m;
    while ((m = childRe.exec(runXml)) !== null) {
      const tag = m[1];
      if (tag === "t") {
        const closeIdx = runXml.indexOf("</w:t>", m.index);
        if (closeIdx === -1) continue;
        const textStart = runXml.indexOf(">", m.index) + 1;
        const text = runXml.slice(textStart, closeIdx);
        parts.push(esc(text));
      } else if (tag === "br" || tag === "cr") {
        parts.push("<br>");
      } else if (tag === "tab") {
        parts.push("&nbsp;&nbsp;&nbsp;&nbsp;");
      } else if (tag === "sym") {
        const charM = m[0].match(/w:char="([^"]+)"/);
        if (charM) {
          const code = parseInt(charM[1], 16);
          if (!isNaN(code)) parts.push(String.fromCharCode(code));
        }
      } else if (tag === "noBreakHyphen") {
        parts.push("\u2011");
      } else if (tag === "drawing") {
        parts.push(this._extractDrawing(runXml));
      }
    }
    return parts.join("");
  }

  _extractDrawing(runXml) {
    const blipM = runXml.match(/<a:blip[^>]*r:embed="([^"]+)"/);
    if (blipM) return `<img data-embed="${blipM[1]}" style="max-width:100%;height:auto" />`;
    return "";
  }

  _convertTable(tblXml) {
    const tblPr = tblXml.match(/<w:tblPr>([\s\S]*?)<\/w:tblPr>/);
    const style = [];

    if (tblPr) {
      const tblPrXml = tblPr[1];
      const bordersM = tblPrXml.match(/<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/);
      if (bordersM) {
        const borderXml = bordersM[1];
        const sides = ["top", "left", "bottom", "right"];
        for (const side of sides) {
          const sideM = borderXml.match(new RegExp(`<w:${side}[^>]*w:val="([^"]*)"[^>]*(?:w:sz="([^"]*)")?[^>]*(?:w:color="([^"]*)")?`));
          if (sideM && sideM[1] !== "none" && sideM[1] !== "nil") {
            const w = Math.round(parseInt(sideM[2] || "4", 10) / 8);
            const c = sideM[3] || "000000";
            style.push(`border-${side}:${w}px solid #${c}`);
          }
        }
      }

      const tblShd = tblPrXml.match(/<w:shd[^>]*w:fill="([^"]+)"/);
      if (tblShd && tblShd[1] !== "auto") style.push(`background-color:#${tblShd[1]}`);

      const tblW = tblPrXml.match(/<w:tblW[^>]*(?:w:w="(\d+)")?/);
      if (tblW && tblW[1]) style.push(`width:${this._twipsToPt(tblW[1])}pt`);
    }

    const styleAttr = style.length ? ` style="${style.join(";")}"` : "";
    const rows = [];
    const trRe = /<w:tr\b[\s\S]*?<\/w:tr>/g;
    let trM;
    while ((trM = trRe.exec(tblXml)) !== null) {
      rows.push(this._convertRow(trM[0]));
    }

    return `<table${styleAttr}>${rows.join("")}</table>`;
  }

  _convertRow(trXml) {
    const cells = [];
    const tcRe = /<w:tc\b[\s\S]*?<\/w:tc>/g;
    let tcM;
    while ((tcM = tcRe.exec(trXml)) !== null) {
      cells.push(this._convertCell(tcM[0]));
    }
    return `<tr>${cells.join("")}</tr>`;
  }

  _convertCell(tcXml) {
    const style = ["border:1px solid #000", "padding:6px 8px", "vertical-align:top"];
    const tcPrM = tcXml.match(/<w:tcPr>([\s\S]*?)<\/w:tcPr>/);

    if (tcPrM) {
      const tcPr = tcPrM[1];
      const shdM = tcPr.match(/<w:shd[^>]*w:fill="([^"]+)"/);
      if (shdM && shdM[1] !== "auto") style.push(`background-color:#${shdM[1]}`);

      const vAlignM = tcPr.match(/<w:vAlign[^>]*w:val="(\w+)"/);
      if (vAlignM) {
        if (vAlignM[1] === "center") style.push("vertical-align:middle");
        else if (vAlignM[1] === "bottom") style.push("vertical-align:bottom");
      }

      const gridSpanM = tcPr.match(/<w:gridSpan[^>]*w:val="(\d+)"/);
      if (gridSpanM && parseInt(gridSpanM[1], 10) > 1) {
        style.push(`colspan:${gridSpanM[1]}`);
      }
    }

    const paras = [];
    const paraRe = /<w:p\b[\s\S]*?<\/w:p>/g;
    let pM;
    while ((pM = paraRe.exec(tcXml)) !== null) {
      paras.push(this._convertPara(pM[0]));
    }

    const styleAttr = ` style="${style.join(";")}"`;
    return `<td${styleAttr}>${paras.join("\n") || "<br>"}</td>`;
  }
}

async function parseDocx(buffer) {
  const parser = new DocxFormatParser(buffer);
  return parser.parse();
}

module.exports = { DocxFormatParser, parseDocx };

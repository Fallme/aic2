const WordExtractor = require("word-extractor");

function esc(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

class DocFormatParser {
  constructor(buffer) {
    this.buffer = buffer;
  }

  async parse() {
    const extractor = new WordExtractor();
    const doc = await extractor.extract(this.buffer);
    const body = doc.getBody() || "";
    const headers = doc.getHeaders() || "";
    const footers = doc.getFooters() || "";

    const fullText = [headers, body, footers].filter(Boolean).join("\n");
    const html = this._buildHtml(body, headers, footers);
    const text = fullText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
    return { html, text };
  }

  _buildHtml(body, headers, footers) {
    const parts = [];

    if (headers.trim()) {
      parts.push(`<div style="border-bottom:1px solid #ccc;padding-bottom:8px;margin-bottom:16px;color:#666;font-size:0.9em">${this._textToHtml(headers)}</div>`);
    }

    parts.push(this._textToHtml(body));

    if (footers.trim()) {
      parts.push(`<div style="border-top:1px solid #ccc;padding-top:8px;margin-top:16px;color:#666;font-size:0.9em">${this._textToHtml(footers)}</div>`);
    }

    return parts.join("\n");
  }

  _textToHtml(text) {
    if (!text || !text.trim()) return "";
    const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    const blocks = [];
    let firstText = true;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) {
        blocks.push('<p><br></p>');
        continue;
      }

      const safe = esc(line);

      if (firstText && line.length <= 80 && /合同|协议|确认书|承诺书|订单|补充协议/.test(line)) {
        blocks.push(`<h1 style="text-align:center;font-size:1.5em;font-weight:bold;margin:16px 0">${safe}</h1>`);
        firstText = false;
        continue;
      }
      firstText = false;

      if (/^第[一二三四五六七八九十百千]+[章节篇]\s*/.test(line) || /^[一二三四五六七八九十]+[、.．]\s*/.test(line)) {
        blocks.push(`<h2 style="font-size:1.2em;font-weight:bold;margin:12px 0 8px">${safe}</h2>`);
      } else if (/^第[一二三四五六七八九十百千]+条\s*/.test(line) || /^\d+[、.．]\s*/.test(line)) {
        blocks.push(`<p style="text-indent:2em;margin:4px 0;line-height:1.8">${safe}</p>`);
      } else if (/^（[一二三四五六七八九十\d]+）/.test(line) || /^\([一二三四五六七八九十\d]+\)/.test(line)) {
        blocks.push(`<p style="text-indent:2em;margin:4px 0;line-height:1.8">${safe}</p>`);
      } else if (/^甲方[：:]/.test(line) || /^乙方[：:]/.test(line) || /^丙方[：:]/.test(line)) {
        blocks.push(`<p style="margin:8px 0;line-height:1.8;font-weight:bold">${safe}</p>`);
      } else if (/签[字名]|盖章|日期|年\s*月\s*日/.test(line) && line.length <= 30) {
        blocks.push(`<p style="margin:4px 0;text-align:right">${safe}</p>`);
      } else {
        blocks.push(`<p style="text-indent:2em;margin:4px 0;line-height:1.8">${safe}</p>`);
      }
    }

    return blocks.join("\n");
  }
}

async function parseDoc(buffer) {
  const parser = new DocFormatParser(buffer);
  return parser.parse();
}

module.exports = { DocFormatParser, parseDoc };

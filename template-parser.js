/**
 * Template Parser — 解析 .docx 模板中的占位符和结构
 *
 * 使用 JSZip 读取 .docx ZIP，解析 word/document.xml 提取：
 * - 段落和表格结构
 * - 占位符字段（【xxx】、{xxx}、下划线空白等）
 * - 字段类型推断
 */

const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * 占位符正则匹配模式
 */
const PLACEHOLDER_PATTERNS = [
  // 【xxx】中文方括号
  { regex: /【([^】]{1,60})】/g, type: "bracket", priority: 1 },
  // {xxx} 花括号
  { regex: /\{([^}{]{1,60})\}/g, type: "brace", priority: 2 },
  // <xxx> 尖括号
  { regex: /<([^<>{]{1,60})>/g, type: "angle", priority: 3 },
  // ____连续下划线（至少4个）
  { regex: /_{4,}/g, type: "blank", priority: 4 },
  // （待填）圆括号内含"待填""填写""请填写"等
  { regex: /（(待填|填写|请填写|请输入|请补充|待补充|待定|暂缺)）/g, type: "pending", priority: 5 },
];

/**
 * 字段类型推断关键词
 */
const FIELD_TYPE_HINTS = {
  甲方: "party_a",
  乙方: "party_b",
  丙方: "party_c",
  买方: "buyer",
  卖方: "seller",
  采购方: "buyer",
  供应商: "supplier",
  出租方: "lessor",
  承租方: "lessee",
  用人单位: "employer",
  劳动者: "employee",
  名称: "name",
  地址: "address",
  法定代表人: "legal_rep",
  联系人: "contact",
  联系电话: "phone",
  电话: "phone",
  手机: "mobile",
  邮箱: "email",
  开户行: "bank",
  账号: "account",
  统一社会信用代码: "credit_code",
  合同编号: "contract_no",
  签订日期: "sign_date",
  签署日期: "sign_date",
  日期: "date",
  金额: "amount",
  总价: "total_price",
  单价: "unit_price",
  数量: "quantity",
  付款方式: "payment_method",
  交付日期: "delivery_date",
  交货日期: "delivery_date",
  质保期: "warranty",
  违约金: "penalty",
  产品: "product",
  服务: "service",
  项目: "project",
};

/**
 * 从 .docx Buffer 解析模板结构
 * @param {Buffer} buffer - .docx 文件 Buffer
 * @returns {Promise<object>} 模板结构
 */
async function parseTemplate(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  // 读取 document.xml
  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("Invalid .docx: missing word/document.xml");

  // 读取 styles.xml（用于样式解析）
  let stylesXml = "";
  try {
    stylesXml = (await zip.file("word/styles.xml")?.async("string")) || "";
  } catch (_) {}

  // 读取 media 列表
  const mediaFiles = [];
  zip.forEach((relPath, entry) => {
    if (relPath.startsWith("word/media/") && !entry.dir) {
      mediaFiles.push(relPath);
    }
  });

  const parser = new DOMParser();
  const doc = parser.parseFromString(docXml, "text/xml");

  // 提取段落
  const paragraphs = extractParagraphs(doc);
  // 提取表格
  const tables = extractTables(doc);
  // 提取所有占位符
  const placeholders = extractPlaceholders(paragraphs, tables);
  // 推断字段类型
  const fields = placeholders.map(inferFieldType);

  return {
    fileName: "",
    paragraphCount: paragraphs.length,
    tableCount: tables.length,
    mediaCount: mediaFiles.length,
    mediaFiles,
    paragraphs: paragraphs.map(summarizeParagraph),
    tables: tables.map(summarizeTable),
    fields,
    fieldCount: fields.length,
    requiredFields: fields.filter((f) => f.importance !== "optional"),
    optionalFields: fields.filter((f) => f.importance === "optional"),
    hasImages: mediaFiles.length > 0,
    hasHeaders: docXml.includes("<w:hdr") || docXml.includes("<w:ftr"),
  };
}

/**
 * 提取段落信息
 */
function extractParagraphs(doc) {
  const paragraphs = [];
  const pNodes = doc.getElementsByTagName("w:p");

  for (let i = 0; i < pNodes.length; i++) {
    const p = pNodes[i];
    const runs = [];
    const runNodes = p.getElementsByTagName("w:r");

    for (let j = 0; j < runNodes.length; j++) {
      const r = runNodes[j];
      const tNode = r.getElementsByTagName("w:t")[0];
      const text = tNode ? tNode.textContent : "";
      const rPr = r.getElementsByTagName("w:rPr")[0];
      runs.push({
        text,
        bold: hasProperty(rPr, "w:b"),
        italic: hasProperty(rPr, "w:i"),
        underline: hasProperty(rPr, "w:u"),
        font: getAttrValue(rPr, "w:rFonts", "w:ascii") || "",
        size: getAttrValue(rPr, "w:sz", "w:val") || "",
        color: getAttrValue(rPr, "w:color", "w:val") || "",
      });
    }

    const fullText = runs.map((r) => r.text).join("");
    const pPr = p.getElementsByTagName("w:pPr")[0];
    const style = pPr ? getAttrValue(pPr, "w:pStyle", "w:val") : "";

    paragraphs.push({
      index: i,
      text: fullText,
      runs,
      style,
      isHeading: /^heading|标题/i.test(style || ""),
    });
  }

  return paragraphs;
}

/**
 * 提取表格信息
 */
function extractTables(doc) {
  const tables = [];
  const tblNodes = doc.getElementsByTagName("w:tbl");

  for (let i = 0; i < tblNodes.length; i++) {
    const tbl = tblNodes[i];
    const rows = [];
    const trNodes = tbl.getElementsByTagName("w:tr");

    for (let j = 0; j < trNodes.length; j++) {
      const tr = trNodes[j];
      const cells = [];
      const tcNodes = tr.getElementsByTagName("w:tc");

      for (let k = 0; k < tcNodes.length; k++) {
        const tc = tcNodes[k];
        const cellText = getNodeText(tc);
        const gridSpan = getAttrValue(tc, "w:tcPr", "w:gridSpan");
        cells.push({
          text: cellText,
          colspan: gridSpan ? parseInt(gridSpan) : 1,
        });
      }

      rows.push({ cells });
    }

    tables.push({ index: i, rows, rowCount: rows.length });
  }

  return tables;
}

/**
 * 从段落和表格中提取占位符
 */
function extractPlaceholders(paragraphs, tables) {
  const placeholders = [];
  const seen = new Set();

  // 段落中的占位符
  for (const p of paragraphs) {
    const found = findPlaceholdersInText(p.text);
    for (const ph of found) {
      const key = `${ph.type}:${ph.placeholder}`;
      if (seen.has(key)) continue;
      seen.add(key);
      placeholders.push({
        ...ph,
        location: "paragraph",
        locationIndex: p.index,
        context: p.text.slice(0, 100),
        style: p.style,
        isHeading: p.isHeading,
      });
    }
  }

  // 表格中的占位符
  for (const tbl of tables) {
    for (let ri = 0; ri < tbl.rows.length; ri++) {
      const row = tbl.rows[ri];
      for (let ci = 0; ci < row.cells.length; ci++) {
        const cell = row.cells[ci];
        const found = findPlaceholdersInText(cell.text);
        for (const ph of found) {
          const key = `${ph.type}:${ph.placeholder}`;
          if (seen.has(key)) continue;
          seen.add(key);
          placeholders.push({
            ...ph,
            location: "table",
            tableIndex: tbl.index,
            rowIndex: ri,
            cellIndex: ci,
            context: cell.text.slice(0, 100),
          });
        }
      }
    }
  }

  return placeholders;
}

/**
 * 在文本中查找所有占位符
 */
function findPlaceholdersInText(text) {
  const results = [];
  for (const pattern of PLACEHOLDER_PATTERNS) {
    const regex = new RegExp(pattern.regex.source, pattern.regex.flags);
    let match;
    while ((match = regex.exec(text))) {
      results.push({
        placeholder: match[0],
        label: match[1] || match[0],
        type: pattern.type,
        priority: pattern.priority,
        start: match.index,
        end: match.index + match[0].length,
      });
    }
  }
  // 按优先级排序
  results.sort((a, b) => a.priority - b.priority);
  return results;
}

/**
 * 推断字段类型和重要性
 */
function inferFieldType(ph) {
  const label = ph.label || ph.placeholder;
  let fieldType = "unknown";
  let importance = "required";

  // 匹配字段类型关键词
  for (const [keyword, type] of Object.entries(FIELD_TYPE_HINTS)) {
    if (label.includes(keyword)) {
      fieldType = type;
      break;
    }
  }

  // 判断重要性
  if (/选填|可选|备注|说明|附注/.test(label)) {
    importance = "optional";
  }
  if (/签章|盖章|签字/.test(label)) {
    importance = "signature";
  }

  return {
    ...ph,
    fieldType,
    importance,
    aiSuggestable: importance !== "signature", // 签章类不可AI建议
  };
}

/**
 * 辅助函数
 */
function hasProperty(parent, tagName) {
  if (!parent) return false;
  return parent.getElementsByTagName(tagName).length > 0;
}

function getAttrValue(parent, tagName, attrName) {
  if (!parent) return "";
  const nodes = parent.getElementsByTagName(tagName);
  if (!nodes.length) return "";
  return nodes[0].getAttribute(attrName) || "";
}

function getNodeText(node) {
  const texts = [];
  const tNodes = node.getElementsByTagName("w:t");
  for (let i = 0; i < tNodes.length; i++) {
    texts.push(tNodes[i].textContent || "");
  }
  return texts.join("");
}

function summarizeParagraph(p) {
  return {
    index: p.index,
    text: p.text.slice(0, 200),
    style: p.style,
    isHeading: p.isHeading,
    runCount: p.runs.length,
    hasPlaceholder: PLACEHOLDER_PATTERNS.some((pat) => pat.regex.test(p.text)),
  };
}

function summarizeTable(t) {
  return {
    index: t.index,
    rowCount: t.rowCount,
    cells: t.rows.slice(0, 3).map((r) => r.cells.map((c) => c.text.slice(0, 50))),
  };
}

module.exports = { parseTemplate, findPlaceholdersInText, FIELD_TYPE_HINTS, PLACEHOLDER_PATTERNS };

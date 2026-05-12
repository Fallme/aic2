require("dotenv").config();
const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const zlib = require("zlib");
const { execFile, spawn } = require("child_process");
const { DocxFormatParser } = require("./docx-format-parser");
const { DocFormatParser } = require("./doc-format-parser");
const { parseTemplate } = require("./template-parser");

const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data");
const UPLOAD_DIR = path.join(DATA_DIR, "uploads");
const STORE_PATH = path.join(DATA_DIR, "kb-store.json");
const DIMENSIONS_PATH = path.join(ROOT, "contract-knowledge-dimensions.json");
const PRESET_RULES_PATH = path.join(ROOT, "preset-contract-rules.json");
const PRESET_RULES_VERSION = "contract-industry-core-2026-05-v2";
const PORT = Number(process.env.PORT || 5173);
const DEFAULT_XIAOMI_BASE_URL = "https://token-plan-cn.xiaomimimo.com/v1";
const LEGACY_XIAOMI_BASE_URL = "https://api.xiaomimimo.com/v1";
const DEFAULT_XIAOMI_MODEL = "mimo-v2.5-pro";
const XIAOMI_BASE_URL = process.env.XIAOMI_BASE_URL || process.env.MIMO_BASE_URL || DEFAULT_XIAOMI_BASE_URL;
const XIAOMI_API_KEY = process.env.XIAOMI_API_KEY || process.env.MIMO_API_KEY || "";
const XIAOMI_MODEL = process.env.XIAOMI_MODEL || process.env.MIMO_MODEL || DEFAULT_XIAOMI_MODEL;
const OPENAI_BASE_URL = process.env.OPENAI_BASE_URL || "https://api.openai.com/v1";
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || process.env.GPT_API_KEY || "";
const OPENAI_MODEL = process.env.OPENAI_MODEL || process.env.GPT_MODEL || "gpt-4o-mini";
const DASHSCOPE_BASE_URL =
  process.env.DASHSCOPE_BASE_URL || "https://dashscope.aliyuncs.com/compatible-mode/v1";
const DASHSCOPE_API_KEY = process.env.DASHSCOPE_API_KEY || "";
const DASHSCOPE_MODEL = process.env.DASHSCOPE_MODEL || "qwen-plus";
const AI_PROVIDER_PRIORITY = (process.env.AI_PROVIDER_PRIORITY || "xiaomi,openai,dashscope")
  .split(",")
  .map((item) => item.trim().toLowerCase())
  .filter(Boolean);
const MODEL = XIAOMI_API_KEY ? XIAOMI_MODEL : OPENAI_API_KEY ? OPENAI_MODEL : DASHSCOPE_MODEL;
const DATABASE_URL = process.env.DATABASE_URL || "";
const DATABASE_SSL = process.env.DATABASE_SSL !== "false";
const STORE_DB_KEY = "main";
let pgPool = null;
let storeCache = null;
let storeBackend = "file";
let dbWriteChain = Promise.resolve();
let presetRulesCache = null;

const EMPLOYMENT_TEMPLATE_TEXT = `# 劳动合同

甲方（用人单位）：
名称：【用人单位】
统一社会信用代码：【统一社会信用代码】
住所地：【用人单位住所地】
法定代表人：【法定代表人】
联系电话：【用人单位联系电话】

乙方（劳动者）：
姓名：【劳动者】
身份证号码：【身份证号码】
住址：【劳动者住址】
联系电话：【劳动者联系电话】
紧急联系人：【紧急联系人】
联系电话：【紧急联系人电话】

甲乙双方根据《中华人民共和国劳动合同法》及相关法律、法规的规定，遵循平等自愿、协商一致、诚实信用的原则，签订本合同，共同遵守。

第一条 合同期限
1.1 本合同期限类型为【合同期限类型】劳动合同。
1.2 本合同有效期自【合同开始日期】起至【合同结束日期】止，共计【合同期限】。
1.3 本合同包含试用期【试用期】。试用期自【试用期开始日期】起至【试用期结束日期】止。试用期工资为人民币【试用期工资】元/月，且不得低于本合同约定工资的80%并不得低于用人单位所在地最低工资标准。

第二条 工作内容与工作地点
2.1 乙方同意根据甲方工作需要，担任【岗位】工作。具体工作职责包括：
（一）【岗位职责一】
（二）【岗位职责二】
（三）【岗位职责三】
2.2 乙方的工作地点为：【工作地点】。
2.3 甲方可根据生产经营需要及乙方的业务能力、工作表现，经与乙方协商一致，可以调整乙方的工作岗位、工作内容和工作地点。

第三条 工作时间与休息休假
3.1 甲方实行【工时制度】工时制度。乙方每日工作时间不超过8小时，每周工作时间不超过40小时；依法实行特殊工时制度的，按主管部门审批或备案要求执行。
3.2 甲方因生产经营需要安排乙方延长工作时间的，应依法与工会和乙方协商，并按照法律规定安排补休或支付加班工资。
3.3 甲方保证乙方依法享有休息日、法定节假日、年休假、婚假、产假、陪产假、病假等国家和地方规定的假期。

第四条 劳动报酬
4.1 乙方正常工作时间的工资标准为每月人民币【月工资】元（大写：【工资大写】）。
4.2 甲方每月【工资支付日】日前以货币形式或银行转账方式足额支付乙方工资。工资支付日遇节假日的，提前至最近的工作日支付。
4.3 乙方的工资构成包括基本工资【基本工资】元、岗位工资【岗位工资】元、绩效工资【绩效工资】元；绩效工资按照甲方依法制定并向乙方公示的考核制度执行。
4.4 甲方安排乙方延长工作时间、休息日工作或法定休假日工作的，应依法支付加班工资或安排补休。
4.5 甲方根据公司经营效益、岗位变化及乙方工作表现调整乙方工资的，应依法履行协商、告知或制度程序。

第五条 社会保险与福利待遇
5.1 甲乙双方必须依法参加社会保险，甲方按照国家和地方规定为乙方缴纳养老保险、医疗保险、失业保险、工伤保险和生育保险，乙方个人应缴纳的部分由甲方从工资中代扣代缴。
5.2 乙方患病或非因工负伤，其医疗期待遇和病假工资按照国家及当地有关规定执行。
5.3 乙方患职业病或因工负伤的，其待遇按照国家及当地有关规定执行。
5.4 甲方为乙方提供的其他福利待遇为：【福利待遇】。

第六条 劳动保护、劳动条件和职业危害防护
6.1 甲方应建立健全劳动安全卫生制度，严格执行国家劳动安全规程和标准，对乙方进行劳动安全卫生教育，为乙方提供符合国家规定的劳动安全卫生条件和必要的劳动防护用品。
6.2 乙方岗位可能存在职业危害的，甲方应向乙方如实告知，并依法采取职业危害防护措施。
6.3 乙方在劳动过程中应严格遵守安全操作规程。乙方对甲方管理人员违章指挥、强令冒险作业，有权拒绝执行。

第七条 规章制度与劳动纪律
7.1 甲方依法制定并公示或告知乙方的规章制度，包括员工手册、岗位职责、劳动纪律、绩效考核、考勤管理等，作为本合同履行依据。
7.2 乙方应遵守甲方劳动纪律，按时完成工作任务，提高职业技能，遵守职业道德。
7.3 乙方违反劳动纪律或规章制度的，甲方可依据依法制定的规章制度和本合同约定处理。

第八条 劳动合同的变更、解除、终止和续订
8.1 经甲乙双方协商一致，可以变更本合同约定的内容。变更劳动合同，应当采用书面形式。
8.2 本合同的解除按照《中华人民共和国劳动合同法》及相关法律法规执行。
8.3 有下列情形之一的，本合同终止：
（一）本合同期满的；
（二）乙方开始依法享受基本养老保险待遇的；
（三）乙方死亡，或者被人民法院宣告死亡或者宣告失踪的；
（四）甲方被依法宣告破产的；
（五）甲方被吊销营业执照、责令关闭、撤销或者甲方决定提前解散的；
（六）法律、行政法规规定的其他情形。
8.4 本合同期满前，甲乙双方应依法协商劳动合同续订事宜。

第九条 经济补偿
9.1 符合《中华人民共和国劳动合同法》第四十六条规定情形的，甲方应依法向乙方支付经济补偿。
9.2 经济补偿按乙方在甲方工作的年限，每满一年支付一个月工资的标准向乙方支付；六个月以上不满一年的，按一年计算；不满六个月的，向乙方支付半个月工资的经济补偿。法律法规另有规定的，从其规定。

第十条 保密与竞业限制
10.1 乙方在职期间及离职后，均有义务保守甲方的商业秘密、个人信息、客户资料、技术资料以及与知识产权相关的保密事项。
10.2 保密范围、保密期限及违约责任为：【保密/竞业】。
10.3 如乙方属于高级管理人员、高级技术人员或其他负有保密义务的人员，双方可另行签订竞业限制协议，明确竞业限制范围、期限、补偿标准及违约责任。

第十一条 培训服务期与违约金
11.1 甲方为乙方提供专项培训费用并进行专业技术培训的，可以与乙方另行订立培训服务期协议。
11.2 乙方违反服务期约定的，应按照约定向甲方支付违约金。违约金不得超过甲方提供的培训费用，且不得超过服务期尚未履行部分所应分摊的培训费用。
11.3 除法律规定的培训服务期和竞业限制情形外，甲方不得与乙方约定由乙方承担违约金。

第十二条 劳动争议处理
12.1 甲乙双方因履行本合同发生劳动争议的，可以协商解决；协商不成的，可以依法向劳动争议调解组织申请调解，或向有管辖权的劳动争议仲裁委员会申请仲裁。对仲裁裁决不服的，除法律另有规定外，可以依法向人民法院提起诉讼。

第十三条 通知与送达
13.1 甲乙双方在本合同中载明的地址、联系方式为有效联系方式和送达地址。任何一方变更联系方式或地址的，应提前【通知提前天数】日书面通知对方。
13.2 甲方向乙方送达与劳动关系有关的文件，可以通过当面交付、邮寄、电子邮件、短信、企业通讯工具等方式送达；因乙方未及时更新联系方式导致无法送达的，由乙方承担相应后果。

第十四条 其他事项
14.1 本合同未尽事宜，按国家及地方有关法律、法规、规章和甲方依法制定的规章制度执行。
14.2 甲方依法制定的规章制度作为本合同附件，与本合同具有同等约束力，但其内容不得违反法律法规强制性规定。
14.3 本合同一式两份，甲乙双方各执一份，具有同等法律效力。自双方签字或盖章之日起生效。

（以下无正文）

签订时间：【签订日期】

甲方（盖章）：
法定代表人或授权代表（签字）：

乙方（签名）：

确认：本人已详细阅读并理解本合同所有条款，自愿与甲方签订本合同，遵守合同约定。

附件一：员工手册
附件二：岗位说明书
附件三：保密协议（如适用）`;

function defaultTemplatePlaceholder(template = {}, key = "", fallback = "") {
  const matched = (template.requiredFields || []).find((field) => field.key === key || field.label === fallback);
  return `【${matched?.label || fallback || key || "待补充"}】`;
}

function buildOfficialStyleTemplateText(template = {}) {
  const partyA = defaultTemplatePlaceholder(template, "partyA", "甲方名称");
  const partyB = defaultTemplatePlaceholder(template, "partyB", "乙方名称");
  const subject = defaultTemplatePlaceholder(template, "subject", "合同标的");
  const scope = defaultTemplatePlaceholder(template, "serviceScope", "服务/履行范围");
  const amount = defaultTemplatePlaceholder(template, "amount", "合同金额");
  const payment = defaultTemplatePlaceholder(template, "payment", "付款方式");
  const delivery = defaultTemplatePlaceholder(template, "delivery", "交付安排");
  const acceptance = defaultTemplatePlaceholder(template, "acceptance", "验收标准");
  const term = defaultTemplatePlaceholder(template, "term", "合同期限");
  const confidentiality = defaultTemplatePlaceholder(template, "confidentiality", "保密要求");
  const ipOwnership = defaultTemplatePlaceholder(template, "ipOwnership", "知识产权");
  const liability = defaultTemplatePlaceholder(template, "liability", "违约责任");
  const coreFields = (template.requiredFields || [])
    .map((field) => `${field.label}：【${field.label}】`)
    .join("\n");
  return `${template.name}

甲方：${partyA}
乙方：${partyB}

鉴于甲乙双方具有签订和履行本合同的合法资格，双方根据《中华人民共和国民法典》及相关法律法规，在平等、自愿、公平、诚实信用的基础上，就${subject}相关事项协商一致，订立本合同。

第一条 合同核心信息
${coreFields}

第二条 合同标的与履行范围
2.1 本合同标的为：${subject}。
2.2 履行范围、技术要求、质量要求、数量规格、服务边界或附件清单为：${scope}。
2.3 双方确认，附件、报价单、订单、需求说明、技术方案、图纸或验收标准经双方盖章、签字或书面确认后，与本合同具有同等效力。

第三条 合同期限与履行安排
3.1 合同期限或项目周期为：${term}。
3.2 交付、运输、实施、服务或阶段成果提交安排为：${delivery}。
3.3 任何一方需变更履行时间、地点、联系人、交付方式或其他关键安排的，应提前书面通知对方并取得确认。

第四条 价款、付款与发票
4.1 本合同价款为：${amount}。价款是否含税、适用税率、发票类型和结算依据应在本合同或附件中明确。
4.2 付款安排为：${payment}。
4.3 收款方应按照约定提交合法有效发票、付款申请资料、验收证明或其他无争议结算资料；付款方在收到完整资料并确认无争议后按约付款。

第五条 交付、验收与整改
5.1 验收标准、验收期限、异议提出方式和整改要求为：${acceptance}。
5.2 验收不合格或交付成果不符合约定的，履约方应在合理期限内完成整改并重新提交验收；因此造成延期或损失的，按本合同违约责任处理。
5.3 付款节点与交付、验收、开票及无争议资料提交条件相互衔接；未完成对应条件的，对方有权暂缓支付相应款项。

第六条 双方权利义务
6.1 甲方应按约提供必要资料、现场条件、确认意见、付款和协作事项。
6.2 乙方应按约完成合同标的相关交付、服务、质量保障、售后支持和资料留存义务。
6.3 双方应指定联系人负责合同履行沟通；联系人变更应及时书面通知对方。

第七条 保密、数据与知识产权
7.1 双方对合作过程中知悉的商业秘密、技术资料、客户资料、个人信息及未公开信息承担保密义务。保密要求为：${confidentiality}。
7.2 合同成果、既有技术、第三方素材、源文件、文档、商标、专利、著作权或其他知识产权归属及使用范围为：${ipOwnership}。
7.3 任何一方使用对方资料、标识、数据或成果，不得超出本合同目的和授权范围。

第八条 违约责任
8.1 任一方违反本合同约定的，应承担继续履行、采取补救措施、赔偿损失、支付违约金或解除合同等责任。
8.2 违约责任具体约定为：${liability}。
8.3 违约金、赔偿范围、责任上限或免责事由应公平合理，不得排除对方主要权利、免除自身主要责任或加重对方责任。

第九条 合同变更、解除与终止
9.1 本合同的变更、补充应采用书面形式，经双方签字或盖章后生效。
9.2 一方严重违约、迟延履行经催告后仍未改正，或发生导致合同目的不能实现的情形，守约方有权依法解除合同并要求违约方承担责任。
9.3 合同解除或终止不影响结算、保密、知识产权、违约责任和争议解决条款的效力。

第十条 不可抗力与通知送达
10.1 因不可抗力导致不能履行合同的，受影响方应及时通知对方并在合理期限内提供证明，双方根据影响程度协商处理。
10.2 双方在本合同中载明的地址、联系人、电话、电子邮箱为有效通知送达信息；任何一方变更的，应及时书面通知对方。

第十一条 争议解决
11.1 因本合同产生的争议，双方应先友好协商；协商不成的，提交【争议解决方式】处理。
11.2 本合同适用中华人民共和国法律。

第十二条 附则
12.1 本合同未尽事宜，双方可另行签订补充协议。
12.2 本合同一式【份数】份，甲乙双方各执【持有份数】份，自双方签字或盖章之日起生效。

甲方（盖章）：____________________
授权代表：____________________
签署日期：【签订日期】

乙方（盖章）：____________________
授权代表：____________________
签署日期：【签订日期】`;
}

function enrichOfficialTemplate(template = {}) {
  if (template.templateText) return template;
  return {
    ...template,
    keywords: [...new Set([...(template.keywords || []), "合同范本", "示范文本", "市场监管"])],
    templateText: buildOfficialStyleTemplateText(template),
  };
}

const CONTRACT_TEMPLATES = [
  {
    id: "purchase",
    name: "采购合同",
    keywords: ["采购", "买设备", "买材料", "供货", "供应商", "采购订单"],
    outline: ["合同主体", "采购标的", "数量质量", "价格与付款", "交付验收", "违约责任", "争议解决"],
    requiredFields: [
      { key: "partyA", label: "采购方", question: "采购方/甲方是谁？" },
      { key: "partyB", label: "供应商", question: "供应商/乙方是谁？" },
      { key: "subject", label: "采购标的", question: "采购什么产品、设备或材料？" },
      { key: "amount", label: "合同金额", question: "合同金额、单价或计价方式是什么？" },
      { key: "delivery", label: "交付安排", question: "交付时间、地点和方式是什么？" },
      { key: "payment", label: "付款方式", question: "付款节点和付款条件是什么？" },
      { key: "acceptance", label: "验收标准", question: "验收标准和验收期限是什么？" },
    ],
  },
  {
    id: "sales",
    name: "销售合同",
    keywords: ["销售", "出售", "卖给", "客户购买", "经销", "销售订单"],
    outline: ["合同主体", "销售产品", "价格税费", "发货交付", "验收与售后", "付款", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "销售方", question: "销售方/甲方是谁？" },
      { key: "partyB", label: "购买方", question: "购买方/乙方是谁？" },
      { key: "subject", label: "销售标的", question: "销售的产品或服务是什么？" },
      { key: "amount", label: "价款", question: "价款、税费和结算方式是什么？" },
      { key: "delivery", label: "交付安排", question: "发货、交付地点和风险转移如何约定？" },
      { key: "payment", label: "付款方式", question: "客户付款节点和账期是什么？" },
    ],
  },
  {
    id: "service",
    name: "服务合同",
    keywords: ["服务", "咨询", "运维", "外包服务", "顾问", "支持"],
    outline: ["服务范围", "服务期限", "服务标准", "费用与付款", "交付成果", "保密", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "委托方", question: "委托方/甲方是谁？" },
      { key: "partyB", label: "服务方", question: "服务方/乙方是谁？" },
      { key: "serviceScope", label: "服务范围", question: "需要提供哪些具体服务？" },
      { key: "term", label: "服务期限", question: "服务期限或关键里程碑是什么？" },
      { key: "amount", label: "服务费用", question: "服务费用和计费方式是什么？" },
      { key: "payment", label: "付款安排", question: "付款节点和付款条件是什么？" },
      { key: "acceptance", label: "验收/考核", question: "服务成果如何验收或考核？" },
    ],
  },
  {
    id: "software",
    name: "软件开发合同",
    keywords: ["软件", "系统开发", "小程序", "APP", "平台开发", "定制开发", "源代码"],
    outline: ["需求范围", "开发计划", "交付物", "验收", "知识产权", "费用付款", "维护支持", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "委托方", question: "委托开发方/甲方是谁？" },
      { key: "partyB", label: "开发方", question: "开发方/乙方是谁？" },
      { key: "serviceScope", label: "开发范围", question: "要开发什么系统、模块或功能？" },
      { key: "delivery", label: "交付物", question: "交付物包括源码、文档、部署包还是其他内容？" },
      { key: "term", label: "开发周期", question: "开发周期和里程碑是什么？" },
      { key: "amount", label: "开发费用", question: "开发费用是多少？" },
      { key: "ipOwnership", label: "知识产权", question: "软件著作权、源码和成果归属如何约定？" },
    ],
  },
  {
    id: "nda",
    name: "保密协议",
    keywords: ["保密", "NDA", "商业秘密", "秘密信息", "尽调", "合作洽谈"],
    outline: ["保密信息范围", "保密义务", "例外情形", "使用限制", "保密期限", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "披露方", question: "披露方是谁？如果双方互相披露，请说明。" },
      { key: "partyB", label: "接收方", question: "接收方是谁？" },
      { key: "subject", label: "合作/披露背景", question: "保密信息用于什么合作、项目或洽谈？" },
      { key: "confidentiality", label: "保密信息范围", question: "保密信息包括哪些内容？" },
      { key: "term", label: "保密期限", question: "保密期限是多久？" },
      { key: "liability", label: "违约责任", question: "泄密后的违约金或赔偿方式如何约定？" },
    ],
  },
  {
    id: "employment",
    name: "劳动合同",
    keywords: ["劳动", "员工", "入职", "岗位", "试用期", "薪资", "工资", "社保", "用工", "劳动者", "用人单位"],
    outline: ["主体信息", "合同期限", "工作内容与地点", "工作时间与休假", "劳动报酬", "社保福利", "劳动保护", "规章制度", "解除终止", "经济补偿", "保密竞业", "争议处理", "签署附件"],
    requiredFields: [
      { key: "partyA", label: "用人单位", question: "用人单位名称是什么？" },
      { key: "employerCreditCode", label: "统一社会信用代码", question: "用人单位统一社会信用代码是什么？" },
      { key: "employerAddress", label: "用人单位住所地", question: "用人单位住所地在哪里？" },
      { key: "legalRep", label: "法定代表人", question: "用人单位法定代表人是谁？" },
      { key: "employerPhone", label: "用人单位联系电话", question: "用人单位联系电话是多少？" },
      { key: "partyB", label: "劳动者", question: "劳动者姓名是什么？" },
      { key: "employeeId", label: "身份证号码", question: "劳动者身份证号码是多少？" },
      { key: "employeeAddress", label: "劳动者住址", question: "劳动者住址在哪里？" },
      { key: "employeePhone", label: "劳动者联系电话", question: "劳动者联系电话是多少？" },
      { key: "emergencyContact", label: "紧急联系人", question: "紧急联系人是谁？" },
      { key: "emergencyPhone", label: "紧急联系人电话", question: "紧急联系人电话是多少？" },
      { key: "termType", label: "合同期限类型", question: "合同期限类型是固定期限、无固定期限还是以完成一定工作任务为期限？" },
      { key: "contractStart", label: "合同开始日期", question: "劳动合同从哪天开始？" },
      { key: "contractEnd", label: "合同结束日期", question: "劳动合同到哪天结束？" },
      { key: "term", label: "合同期限", question: "合同期限共多久？" },
      { key: "probation", label: "试用期", question: "试用期多久？如无试用期请填写无。" },
      { key: "probationStart", label: "试用期开始日期", question: "试用期从哪天开始？" },
      { key: "probationEnd", label: "试用期结束日期", question: "试用期到哪天结束？" },
      { key: "probationSalary", label: "试用期工资", question: "试用期工资是多少元/月？" },
      { key: "subject", label: "岗位", question: "岗位、职级和工作地点是什么？" },
      { key: "jobDuty1", label: "岗位职责一", question: "第一项主要岗位职责是什么？" },
      { key: "jobDuty2", label: "岗位职责二", question: "第二项主要岗位职责是什么？" },
      { key: "jobDuty3", label: "岗位职责三", question: "第三项主要岗位职责是什么？" },
      { key: "workLocation", label: "工作地点", question: "工作地点在哪里？" },
      { key: "workHours", label: "工时制度", question: "适用标准工时、综合工时还是不定时工时？" },
      { key: "amount", label: "月工资", question: "正常工作时间工资是多少元/月？" },
      { key: "salaryChinese", label: "工资大写", question: "工资大写金额是什么？" },
      { key: "payDay", label: "工资支付日", question: "每月几日前支付工资？" },
      { key: "baseSalary", label: "基本工资", question: "基本工资是多少？" },
      { key: "positionSalary", label: "岗位工资", question: "岗位工资是多少？" },
      { key: "performanceSalary", label: "绩效工资", question: "绩效工资是多少？" },
      { key: "welfare", label: "福利待遇", question: "除社保外还有哪些福利待遇？" },
      { key: "confidentiality", label: "保密/竞业", question: "是否需要保密、竞业限制或知识产权约定？" },
      { key: "noticeDays", label: "通知提前天数", question: "联系方式或地址变更需提前几日通知？" },
      { key: "signingDate", label: "签订日期", question: "合同签订日期是哪天？" },
    ],
    templateText: EMPLOYMENT_TEMPLATE_TEXT,
  },
  {
    id: "lease",
    name: "租赁合同",
    keywords: ["租赁", "出租", "租房", "租金", "房屋", "场地", "设备租赁"],
    outline: ["租赁物", "租赁期限", "租金押金", "交付使用", "维修维护", "转租限制", "违约解除"],
    requiredFields: [
      { key: "partyA", label: "出租方", question: "出租方是谁？" },
      { key: "partyB", label: "承租方", question: "承租方是谁？" },
      { key: "subject", label: "租赁物", question: "租赁物是什么，地址或编号是什么？" },
      { key: "term", label: "租赁期限", question: "租赁期限从何时到何时？" },
      { key: "amount", label: "租金/押金", question: "租金、押金和支付周期是什么？" },
      { key: "delivery", label: "交付状态", question: "交付标准、附属设施和使用用途是什么？" },
    ],
  },
  {
    id: "agency",
    name: "委托代理合同",
    keywords: ["委托", "代理", "代办", "授权", "代表", "居间"],
    outline: ["委托事项", "授权范围", "代理期限", "费用报酬", "报告义务", "禁止行为", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "委托方", question: "委托方是谁？" },
      { key: "partyB", label: "受托方/代理方", question: "受托方或代理方是谁？" },
      { key: "subject", label: "委托事项", question: "具体委托或代理事项是什么？" },
      { key: "term", label: "委托期限", question: "委托期限是多久？" },
      { key: "amount", label: "报酬/费用", question: "代理报酬、费用承担和支付方式是什么？" },
      { key: "liability", label: "责任边界", question: "越权代理、未完成事项的责任如何约定？" },
    ],
  },
  {
    id: "cooperation",
    name: "合作协议",
    keywords: ["合作", "联合", "战略合作", "渠道合作", "共同", "分成"],
    outline: ["合作目标", "合作内容", "分工职责", "费用收益", "知识产权", "保密", "退出机制", "争议解决"],
    requiredFields: [
      { key: "partyA", label: "合作方一", question: "合作方一是谁？" },
      { key: "partyB", label: "合作方二", question: "合作方二是谁？" },
      { key: "subject", label: "合作事项", question: "合作项目、产品或业务目标是什么？" },
      { key: "serviceScope", label: "分工职责", question: "各方分别承担哪些工作和资源投入？" },
      { key: "amount", label: "费用/收益", question: "成本、收益分配或分成机制是什么？" },
      { key: "term", label: "合作期限", question: "合作期限、续约和退出机制如何约定？" },
    ],
  },
  {
    id: "loan",
    name: "借款合同",
    keywords: ["借款", "贷款", "还款", "利息", "出借", "借贷"],
    outline: ["借款金额", "借款用途", "放款", "利息", "还款", "担保", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "出借方", question: "出借方是谁？" },
      { key: "partyB", label: "借款方", question: "借款方是谁？" },
      { key: "amount", label: "借款金额", question: "借款本金是多少？" },
      { key: "subject", label: "借款用途", question: "借款用途是什么？" },
      { key: "term", label: "借款期限", question: "借款期限、放款日和到期日是什么？" },
      { key: "payment", label: "还款/利息", question: "利率、还款方式和还款节点是什么？" },
      { key: "liability", label: "担保/违约", question: "是否有担保、逾期利息或违约责任？" },
    ],
  },
  {
    id: "education_exam_cooperation",
    name: "教育考试合作协议",
    keywords: ["教育", "考试", "报考", "培训", "证书", "职业资格", "报名", "UI设计师", "考试合作"],
    outline: ["合作内容", "报名资料", "报名费用", "缴费安排", "考试与证书", "保密规定", "违约责任", "争议解决"],
    requiredFields: [
      { key: "partyA", label: "甲方名称", question: "甲方/服务协助方是谁？" },
      { key: "partyB", label: "乙方名称", question: "乙方/报考方是谁？" },
      { key: "examDepartment", label: "报考部门", question: "报考部门或考试组织机构是什么？" },
      { key: "examType", label: "报考工种", question: "报考工种、等级或证书名称是什么？" },
      { key: "applicantCount", label: "报名人数", question: "报名人数是多少？" },
      { key: "amount", label: "报名费用", question: "报名费用合计多少，是否含税？" },
      { key: "payment", label: "缴费方式", question: "费用缴纳节点、账户和凭证要求是什么？" },
      { key: "term", label: "考试/证书期限", question: "考试安排、证书取得或结果通知期限是什么？" },
    ],
    templateText: `教育考试合作协议书

甲方：【甲方名称】
乙方：【乙方名称】

第一章 总则
甲乙双方本着诚实守信、互惠互利、相互支持、公平公正的原则，就乙方报考【报考工种】等事项达成合作共识，明确双方权利义务，经过平等协商，自愿达成本协议。

第二章 合作内容
一、合作内容
【报考部门】：【报考部门】
【报考工种】：【报考工种】
【报名数量】：【报名人数】人

二、报名费用
报名费用合计：【报名费用】。
乙方应按照本协议约定及时向甲方缴纳报名费用。甲方收到费用及乙方完整报名资料后，按照考试报名规则协助乙方办理报名、考试安排或证书取得相关事项。

第三章 双方权利义务
1. 甲方应按照约定流程协助乙方完成报名、学习或考试事项安排，并对乙方提交的报名资料承担保密义务。
2. 乙方应保证其提供的身份、学历、工作经历、联系方式及其他报名资料真实、准确、完整。因乙方资料不真实、不完整或逾期提交导致无法报名、无法考试或证书结果受影响的，由乙方自行承担责任。
3. 乙方应按照甲方通知准时参加考试、学习或相关流程，并遵守考试组织机构及当地相关规定。

第四章 保密规定
甲乙双方对在合作过程中知悉的身份信息、联系方式、报名资料、考试资料、商业信息及未公开资料负有保密义务，未经对方书面同意不得向第三方披露，但法律法规另有规定或监管机构要求披露的除外。

第五章 违约责任
任何一方违反本协议约定，给对方造成损失的，应承担继续履行、采取补救措施或赔偿损失等违约责任。因不可归责于甲方的考试政策调整、考试机构安排变化或乙方个人原因导致考试、证书结果变化的，双方应根据实际情况协商处理。

第六章 争议解决
因本协议产生的争议，双方应先友好协商；协商不成的，提交甲方所在地有管辖权的人民法院处理。`,
  },
  {
    id: "system_integration",
    name: "系统集成合同",
    keywords: ["系统集成", "软硬件", "联调", "设备安装", "信息化项目", "接口对接", "等保"],
    outline: ["集成范围", "软硬件清单", "实施计划", "联调测试", "验收", "质保维保", "信息安全", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "甲方名称", question: "甲方/建设方是谁？" },
      { key: "partyB", label: "乙方名称", question: "乙方/集成方是谁？" },
      { key: "subject", label: "项目名称", question: "系统集成项目名称是什么？" },
      { key: "serviceScope", label: "集成范围", question: "集成范围、设备清单和系统边界是什么？" },
      { key: "delivery", label: "实施交付", question: "实施计划、联调和交付安排是什么？" },
      { key: "acceptance", label: "验收标准", question: "初验、试运行、终验标准是什么？" },
      { key: "amount", label: "合同金额", question: "合同金额和税费安排是什么？" },
      { key: "payment", label: "付款节点", question: "付款节点是否与交付验收绑定？" },
    ],
  },
  {
    id: "manufacturing_supply",
    name: "制造供货合同",
    keywords: ["制造", "供货", "设备", "零部件", "生产", "质量标准", "质保", "交货"],
    outline: ["产品规格", "技术标准", "数量价格", "生产交付", "检验验收", "质保售后", "违约责任", "争议解决"],
    requiredFields: [
      { key: "partyA", label: "采购方", question: "采购方是谁？" },
      { key: "partyB", label: "供货方", question: "供货方是谁？" },
      { key: "subject", label: "产品/设备", question: "产品、设备或零部件名称及规格是什么？" },
      { key: "amount", label: "价款", question: "总价、单价和税费如何约定？" },
      { key: "delivery", label: "交货安排", question: "交货地点、批次、运输和风险转移如何约定？" },
      { key: "acceptance", label: "质量验收", question: "检验标准、异议期和整改机制是什么？" },
      { key: "term", label: "质保期", question: "质保期、维修更换和售后责任如何约定？" },
    ],
  },
  {
    id: "processing_oem",
    name: "加工/OEM合同",
    keywords: ["加工", "代工", "OEM", "ODM", "委托生产", "来料加工", "贴牌"],
    outline: ["加工产品", "图纸工艺", "材料供应", "质量检验", "交付结算", "知识产权", "保密", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "委托方", question: "委托加工方是谁？" },
      { key: "partyB", label: "加工方", question: "加工生产方是谁？" },
      { key: "subject", label: "加工产品", question: "加工产品、规格型号和数量是什么？" },
      { key: "serviceScope", label: "加工要求", question: "图纸、工艺、BOM、封样或质量要求是什么？" },
      { key: "amount", label: "加工费用", question: "加工费、材料费和结算方式是什么？" },
      { key: "delivery", label: "交付安排", question: "交付批次、地点和包装运输要求是什么？" },
      { key: "ipOwnership", label: "知识产权/模具", question: "图纸、模具、商标和成果归属如何约定？" },
    ],
  },
  {
    id: "technical_service",
    name: "技术服务合同",
    keywords: ["技术服务", "技术支持", "实施", "运维", "调试", "测试", "技术咨询"],
    outline: ["服务范围", "人员安排", "服务标准", "成果交付", "验收考核", "费用付款", "保密与数据", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "委托方", question: "委托方是谁？" },
      { key: "partyB", label: "服务方", question: "技术服务方是谁？" },
      { key: "serviceScope", label: "服务范围", question: "具体技术服务内容是什么？" },
      { key: "term", label: "服务期限", question: "服务期限、响应时限或里程碑是什么？" },
      { key: "delivery", label: "服务成果", question: "需要提交哪些报告、文档或成果？" },
      { key: "amount", label: "服务费用", question: "服务费用和计费方式是什么？" },
      { key: "acceptance", label: "验收考核", question: "服务如何验收或考核？" },
    ],
  },
  {
    id: "consulting",
    name: "咨询顾问合同",
    keywords: ["咨询", "顾问", "方案", "报告", "调研", "管理咨询", "财务顾问"],
    outline: ["咨询事项", "工作方式", "交付成果", "费用付款", "资料提供", "保密", "知识产权", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "委托方", question: "委托方是谁？" },
      { key: "partyB", label: "顾问方", question: "顾问或咨询服务方是谁？" },
      { key: "serviceScope", label: "咨询范围", question: "咨询项目、目标和范围是什么？" },
      { key: "delivery", label: "交付成果", question: "交付报告、方案或会议成果是什么？" },
      { key: "term", label: "服务周期", question: "咨询周期和关键节点是什么？" },
      { key: "amount", label: "咨询费用", question: "咨询费用和付款节点是什么？" },
    ],
  },
  {
    id: "distribution",
    name: "经销/代理销售合同",
    keywords: ["经销", "代理销售", "渠道", "分销", "授权区域", "销售指标", "返利"],
    outline: ["授权范围", "产品价格", "销售指标", "订单交付", "结算返利", "市场合规", "终止退出", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "授权方", question: "授权方/供货方是谁？" },
      { key: "partyB", label: "经销方", question: "经销商或代理销售方是谁？" },
      { key: "subject", label: "产品/服务", question: "授权销售的产品或服务是什么？" },
      { key: "serviceScope", label: "授权区域", question: "授权区域、渠道和客户范围是什么？" },
      { key: "amount", label: "价格政策", question: "价格、折扣、返利和费用政策是什么？" },
      { key: "term", label: "授权期限", question: "授权期限和续约条件是什么？" },
      { key: "liability", label: "销售合规", question: "串货、虚假宣传、价格管控等责任如何约定？" },
    ],
  },
  {
    id: "logistics",
    name: "物流运输合同",
    keywords: ["物流", "运输", "承运", "配送", "仓储", "货物运输", "到货"],
    outline: ["运输货物", "运输路线", "交付签收", "运费结算", "货损货差", "保险", "时效责任", "争议解决"],
    requiredFields: [
      { key: "partyA", label: "托运方", question: "托运方是谁？" },
      { key: "partyB", label: "承运方", question: "承运方是谁？" },
      { key: "subject", label: "货物信息", question: "运输货物名称、数量、包装和价值是什么？" },
      { key: "delivery", label: "运输安排", question: "起运地、目的地、运输方式和时效是什么？" },
      { key: "amount", label: "运费", question: "运费、装卸费和结算方式是什么？" },
      { key: "acceptance", label: "签收验收", question: "签收、异议和货损货差处理如何约定？" },
    ],
  },
  {
    id: "construction_decoration",
    name: "工程/装修合同",
    keywords: ["工程", "装修", "施工", "安装", "改造", "工程款", "竣工验收"],
    outline: ["工程范围", "图纸预算", "工期", "材料设备", "工程款", "竣工验收", "质保维修", "安全责任"],
    requiredFields: [
      { key: "partyA", label: "发包方", question: "发包方/业主是谁？" },
      { key: "partyB", label: "承包方", question: "承包方/施工方是谁？" },
      { key: "subject", label: "工程项目", question: "工程或装修项目名称、地点是什么？" },
      { key: "serviceScope", label: "施工范围", question: "施工范围、图纸和预算内容是什么？" },
      { key: "term", label: "工期", question: "开工、竣工和延期规则是什么？" },
      { key: "amount", label: "工程款", question: "工程款、付款节点和结算方式是什么？" },
      { key: "acceptance", label: "竣工验收", question: "竣工验收、整改和质保期如何约定？" },
    ],
  },
  {
    id: "saas_license",
    name: "SaaS/软件许可合同",
    keywords: ["SaaS", "软件许可", "账号", "订阅", "云服务", "软件使用", "许可"],
    outline: ["许可范围", "账号权限", "服务等级", "费用续费", "数据安全", "可用性", "终止退出", "知识产权"],
    requiredFields: [
      { key: "partyA", label: "客户方", question: "客户方是谁？" },
      { key: "partyB", label: "服务提供方", question: "SaaS或软件许可提供方是谁？" },
      { key: "subject", label: "软件/服务", question: "软件、模块、账号或云服务名称是什么？" },
      { key: "serviceScope", label: "许可范围", question: "授权用户、功能范围和使用限制是什么？" },
      { key: "term", label: "订阅期限", question: "订阅或许可期限是什么？" },
      { key: "amount", label: "许可费用", question: "费用、续费和计费方式是什么？" },
      { key: "confidentiality", label: "数据安全", question: "数据存储、备份、删除和安全责任如何约定？" },
    ],
  },
  {
    id: "data_processing",
    name: "数据处理委托协议",
    keywords: ["数据处理", "个人信息", "委托处理", "数据安全", "隐私", "删除返还", "处理者"],
    outline: ["处理目的", "数据范围", "处理方式", "安全措施", "个人权利响应", "删除返还", "审计监督", "违约责任"],
    requiredFields: [
      { key: "partyA", label: "委托方", question: "数据处理委托方是谁？" },
      { key: "partyB", label: "受托方", question: "数据处理受托方是谁？" },
      { key: "subject", label: "数据类型", question: "涉及哪些数据或个人信息类型？" },
      { key: "serviceScope", label: "处理目的/方式", question: "处理目的、方式、系统和权限是什么？" },
      { key: "term", label: "处理期限", question: "处理期限和删除返还时间是什么？" },
      { key: "confidentiality", label: "安全措施", question: "加密、访问控制、日志、备份和泄露通知如何约定？" },
    ],
  },
  {
    id: "ip_license",
    name: "知识产权许可合同",
    keywords: ["知识产权", "许可", "商标", "专利", "著作权", "授权使用", "版权"],
    outline: ["许可标的", "许可范围", "地域期限", "许可费用", "质量控制", "权利保证", "侵权处理", "终止返还"],
    requiredFields: [
      { key: "partyA", label: "许可方", question: "知识产权许可方是谁？" },
      { key: "partyB", label: "被许可方", question: "被许可方是谁？" },
      { key: "subject", label: "许可标的", question: "许可的商标、专利、著作权或作品是什么？" },
      { key: "serviceScope", label: "许可范围", question: "许可方式、地域、用途和是否可转授权如何约定？" },
      { key: "term", label: "许可期限", question: "许可期限是什么？" },
      { key: "amount", label: "许可费用", question: "许可费、结算和税费如何约定？" },
      { key: "liability", label: "侵权责任", question: "权利瑕疵、第三方侵权和违约责任如何约定？" },
    ],
  },
].map(enrichOfficialTemplate);

const COMMON_REVIEW_ISSUES = [
  "主体名称、证照号码、签约权限不完整",
  "标的、数量、质量标准或服务范围描述不清",
  "价款、税费、付款节点、发票要求不明确",
  "付款、交付、验收、开票、违约责任之间的业务逻辑不匹配",
  "交付、验收、异议期和整改机制缺失",
  "违约责任过轻、过重或缺少可执行标准",
  "知识产权、成果归属、源代码交付约定不清",
  "保密、数据安全、个人信息保护条款不足",
  "解除、终止、续约和退出机制不完整",
  "不可抗力、通知送达、争议解决条款缺失",
  "表述存在绝对化、单方解释、无限责任、免责过宽、期限不明或对象不清",
  "用词错误、错别字、语病、前后称谓不一致或同一概念多种表述",
  "权利义务明显失衡，只有一方义务或一方可任意变更、解除、暂停履行",
  "内部审批、授权、盖章或金额阈值规则未满足",
  "合同标题、编号、条款层级、附件引用、签署栏或日期格式不规范",
  "正文与附件、报价单、订单、验收单等文件引用不一致",
];

fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
};

function normalizeStore(store = {}) {
  return {
    ...store,
    documents: Array.isArray(store.documents) ? store.documents : [],
    rules: Array.isArray(store.rules) ? store.rules : [],
    contractMemory: Array.isArray(store.contractMemory) ? store.contractMemory : [],
    customTemplates: Array.isArray(store.customTemplates) ? store.customTemplates : [],
  };
}

function loadPresetRules() {
  if (presetRulesCache) return presetRulesCache;
  if (!fs.existsSync(PRESET_RULES_PATH)) {
    presetRulesCache = [];
    return presetRulesCache;
  }
  const raw = JSON.parse(fs.readFileSync(PRESET_RULES_PATH, "utf8"));
  presetRulesCache = (Array.isArray(raw) ? raw : [])
    .map((rule) => ({
      ...normalizeRule(rule),
      id: rule.id,
      reviewStatus: "active",
      ruleSource: "平台预设",
      ruleBasis: rule.ruleBasis || "通用法规",
      useScope: rule.useScope || ["生成", "审查"],
      sourceUrl: rule.sourceUrl || "",
      check: rule.check || null,
      presetVersion: PRESET_RULES_VERSION,
    }))
    .filter((rule) => rule.id && rule.ruleName);
  return presetRulesCache;
}

function ensurePresetRules(store = {}) {
  const normalized = normalizeStore(store);
  const presets = loadPresetRules();
  if (!presets.length) return { store: normalized, inserted: [] };

  const existingIds = new Set(normalized.rules.map((rule) => rule.id).filter(Boolean));
  const existingKeys = new Set(normalized.rules.map(ruleFingerprint));
  const inserted = [];
  for (const preset of presets) {
    const key = ruleFingerprint(preset);
    if (existingIds.has(preset.id) || existingKeys.has(key)) continue;
    inserted.push(JSON.parse(JSON.stringify(preset)));
    existingIds.add(preset.id);
    existingKeys.add(key);
  }
  if (inserted.length) normalized.rules = [...inserted, ...normalized.rules];
  normalized.presetRuleVersion = PRESET_RULES_VERSION;
  return { store: normalized, inserted };
}

function cloneStore(store = {}) {
  return normalizeStore(JSON.parse(JSON.stringify(normalizeStore(store))));
}

function hasStoreData(store = {}) {
  const normalized = normalizeStore(store);
  return Boolean(normalized.documents.length || normalized.rules.length || normalized.contractMemory.length || normalized.customTemplates.length);
}

function loadFileStore() {
  if (!fs.existsSync(STORE_PATH)) return normalizeStore();
  return normalizeStore(JSON.parse(fs.readFileSync(STORE_PATH, "utf8")));
}

function saveFileStore(store) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(STORE_PATH, JSON.stringify(normalizeStore(store), null, 2), "utf8");
}

function tryCreatePgPool() {
  if (!DATABASE_URL) return null;
  try {
    const { Pool } = require("pg");
    return new Pool({
      connectionString: DATABASE_URL,
      ssl: DATABASE_SSL ? { rejectUnauthorized: false } : false,
    });
  } catch (error) {
    console.warn(`PostgreSQL dependency unavailable, falling back to file store: ${error.message}`);
    return null;
  }
}

async function initDatabaseStore(localStore) {
  pgPool = tryCreatePgPool();
  if (!pgPool) return localStore;
  await pgPool.query(`
    create table if not exists app_store (
      key text primary key,
      value jsonb not null,
      updated_at timestamptz not null default now()
    )
  `);
  const result = await pgPool.query("select value from app_store where key = $1", [STORE_DB_KEY]);
  if (result.rows[0]?.value) {
    storeBackend = "postgres";
    return normalizeStore(result.rows[0].value);
  }
  const initialStore = normalizeStore(localStore);
  await pgPool.query(
    `insert into app_store (key, value, updated_at)
     values ($1, $2::jsonb, now())
     on conflict (key) do update set value = excluded.value, updated_at = now()`,
    [STORE_DB_KEY, JSON.stringify(initialStore)]
  );
  storeBackend = "postgres";
  return initialStore;
}

async function initializeStore() {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
  const localStore = loadFileStore();
  try {
    const seeded = ensurePresetRules(await initDatabaseStore(localStore));
    storeCache = normalizeStore(seeded.store);
    if (seeded.inserted.length) {
      console.log(`Preset contract rules seeded: ${seeded.inserted.length}`);
      saveFileStore(storeCache);
      scheduleDatabaseSave(storeCache);
    } else if (storeBackend === "postgres" && hasStoreData(storeCache)) {
      saveFileStore(storeCache);
    }
  } catch (error) {
    console.warn(`PostgreSQL store init failed, falling back to file store: ${error.message}`);
    pgPool = null;
    storeBackend = "file";
    const seeded = ensurePresetRules(localStore);
    storeCache = normalizeStore(seeded.store);
    if (seeded.inserted.length) saveFileStore(storeCache);
  }
}

function scheduleDatabaseSave(store) {
  if (!pgPool) return;
  const snapshot = JSON.stringify(normalizeStore(store));
  dbWriteChain = dbWriteChain
    .then(() =>
      pgPool.query(
        `insert into app_store (key, value, updated_at)
         values ($1, $2::jsonb, now())
         on conflict (key) do update set value = excluded.value, updated_at = now()`,
        [STORE_DB_KEY, snapshot]
      )
    )
    .catch((error) => {
      console.error(`PostgreSQL store save failed: ${error.message}`);
    });
}

function loadStore() {
  if (!storeCache) storeCache = ensurePresetRules(loadFileStore()).store;
  return cloneStore(storeCache);
}

function saveStore(store) {
  storeCache = normalizeStore(store);
  saveFileStore(storeCache);
  scheduleDatabaseSave(storeCache);
}

function sendJson(res, status, body) {
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "GET,POST,OPTIONS",
    "access-control-allow-headers": "content-type,authorization",
  });
  res.end(JSON.stringify(body));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function splitBuffer(buffer, delimiter) {
  const parts = [];
  let start = 0;
  let index;
  while ((index = buffer.indexOf(delimiter, start)) !== -1) {
    parts.push(buffer.subarray(start, index));
    start = index + delimiter.length;
  }
  parts.push(buffer.subarray(start));
  return parts;
}

function parseMultipart(body, contentType) {
  const boundaryMatch = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType || "");
  if (!boundaryMatch) return { fields: {}, files: [] };
  const boundary = Buffer.from(`--${boundaryMatch[1] || boundaryMatch[2]}`);
  const fields = {};
  const files = [];

  for (const rawPart of splitBuffer(body, boundary)) {
    let part = rawPart;
    if (part.length < 8) continue;
    if (part.subarray(0, 2).toString() === "\r\n") part = part.subarray(2);
    if (part.subarray(0, 2).toString() === "--") continue;
    const headerEnd = part.indexOf(Buffer.from("\r\n\r\n"));
    if (headerEnd === -1) continue;
    const header = part.subarray(0, headerEnd).toString("utf8");
    let value = part.subarray(headerEnd + 4);
    if (value.subarray(value.length - 2).toString() === "\r\n") value = value.subarray(0, value.length - 2);
    const name = /name="([^"]+)"/i.exec(header)?.[1];
    const filename = /filename="([^"]*)"/i.exec(header)?.[1];
    const fileType = /content-type:\s*([^\r\n]+)/i.exec(header)?.[1] || "application/octet-stream";
    if (!name) continue;
    if (filename) files.push({ field: name, filename: path.basename(filename), contentType: fileType, buffer: value });
    else fields[name] = value.toString("utf8");
  }
  return { fields, files };
}

function xmlDecode(value) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function escapeHtmlForDocument(value = "") {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function unzipEntry(buffer, entryName) {
  let eocd = -1;
  for (let i = buffer.length - 22; i >= 0; i -= 1) {
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd === -1) throw new Error("DOCX zip directory not found");
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let cursor = buffer.readUInt32LE(eocd + 16);
  for (let i = 0; i < totalEntries; i += 1) {
    if (buffer.readUInt32LE(cursor) !== 0x02014b50) break;
    const method = buffer.readUInt16LE(cursor + 10);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const nameLen = buffer.readUInt16LE(cursor + 28);
    const extraLen = buffer.readUInt16LE(cursor + 30);
    const commentLen = buffer.readUInt16LE(cursor + 32);
    const localOffset = buffer.readUInt32LE(cursor + 42);
    const name = buffer.subarray(cursor + 46, cursor + 46 + nameLen).toString("utf8");
    if (name === entryName) {
      const localNameLen = buffer.readUInt16LE(localOffset + 26);
      const localExtraLen = buffer.readUInt16LE(localOffset + 28);
      const dataStart = localOffset + 30 + localNameLen + localExtraLen;
      const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
      if (method === 0) return compressed;
      if (method === 8) return zlib.inflateRawSync(compressed);
      throw new Error(`Unsupported zip compression method: ${method}`);
    }
    cursor += 46 + nameLen + extraLen + commentLen;
  }
  return Buffer.alloc(0);
}

function readZipEntries(buffer) {
  let eocd = -1;
  for (let i = buffer.length - 22; i >= 0; i -= 1) {
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd === -1) throw new Error("DOCX zip directory not found");
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let cursor = buffer.readUInt32LE(eocd + 16);
  const entries = [];
  for (let i = 0; i < totalEntries; i += 1) {
    if (buffer.readUInt32LE(cursor) !== 0x02014b50) break;
    const method = buffer.readUInt16LE(cursor + 10);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const uncompressedSize = buffer.readUInt32LE(cursor + 24);
    const nameLen = buffer.readUInt16LE(cursor + 28);
    const extraLen = buffer.readUInt16LE(cursor + 30);
    const commentLen = buffer.readUInt16LE(cursor + 32);
    const localOffset = buffer.readUInt32LE(cursor + 42);
    const name = buffer.subarray(cursor + 46, cursor + 46 + nameLen).toString("utf8");
    const localNameLen = buffer.readUInt16LE(localOffset + 26);
    const localExtraLen = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
    let data;
    if (method === 0) data = Buffer.from(compressed);
    else if (method === 8) data = zlib.inflateRawSync(compressed);
    else throw new Error(`Unsupported zip compression method: ${method}`);
    entries.push({ name, data, uncompressedSize });
    cursor += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date = new Date()) {
  const year = Math.max(1980, date.getFullYear());
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const day = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, day };
}

function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const { time, day } = dosDateTime();
  for (const entry of entries) {
    const nameBuffer = Buffer.from(entry.name, "utf8");
    const raw = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data || "");
    const compressed = entry.name.endsWith("/") ? Buffer.alloc(0) : zlib.deflateRawSync(raw);
    const method = entry.name.endsWith("/") ? 0 : 8;
    const crc = crc32(raw);
    const local = Buffer.alloc(30 + nameBuffer.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(time, 10);
    local.writeUInt16LE(day, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuffer.length, 26);
    local.writeUInt16LE(0, 28);
    nameBuffer.copy(local, 30);
    localParts.push(local, compressed);

    const central = Buffer.alloc(46 + nameBuffer.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(method, 10);
    central.writeUInt16LE(time, 12);
    central.writeUInt16LE(day, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBuffer.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(entry.name.endsWith("/") ? 0x10 : 0, 38);
    central.writeUInt32LE(offset, 42);
    nameBuffer.copy(central, 46);
    centralParts.push(central);
    offset += local.length + compressed.length;
  }
  const centralStart = offset;
  const centralBuffer = Buffer.concat(centralParts);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralBuffer.length, 12);
  eocd.writeUInt32LE(centralStart, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...localParts, centralBuffer, eocd]);
}

function xmlEncodeText(value = "") {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function nonEmptyDocLines(text = "") {
  return String(text || "")
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

function replaceParagraphTextXml(paraXml = "", nextText = "") {
  const matches = Array.from(paraXml.matchAll(/<w:t([^>]*)>([\s\S]*?)<\/w:t>/g));
  if (!matches.length) return paraXml;
  let output = paraXml;
  for (let i = matches.length - 1; i >= 0; i -= 1) {
    const match = matches[i];
    const attrs = match[1] || "";
    const content = i === 0 ? xmlEncodeText(nextText) : "";
    const start = match.index;
    const end = start + match[0].length;
    const needsSpace = /^\s|\s$/.test(content);
    const normalizedAttrs = needsSpace && !/xml:space=/.test(attrs) ? `${attrs} xml:space="preserve"` : attrs;
    output = `${output.slice(0, start)}<w:t${normalizedAttrs}>${content}</w:t>${output.slice(end)}`;
  }
  return output;
}

function patchDocxDocumentXml(documentXml = "", currentText = "") {
  const lines = nonEmptyDocLines(currentText);
  let lineIndex = 0;
  const patched = documentXml.replace(/<w:p[\s\S]*?<\/w:p>/g, (paraXml) => {
    const texts = Array.from(paraXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)).map((match) => xmlDecode(match[1]));
    const oldText = texts.join("").replace(/\s+/g, " ").trim();
    if (!oldText) return paraXml;
    const next = lines[lineIndex++];
    if (typeof next !== "string" || next === oldText) return paraXml;
    return replaceParagraphTextXml(paraXml, next);
  });
  return patched;
}

function patchDocxBuffer(originalBuffer, currentText = "") {
  const entries = readZipEntries(originalBuffer);
  const documentEntry = entries.find((entry) => entry.name === "word/document.xml");
  if (!documentEntry) throw new Error("未找到 Word 正文 XML，无法生成保真修改版");
  documentEntry.data = Buffer.from(patchDocxDocumentXml(documentEntry.data.toString("utf8"), currentText), "utf8");
  return buildZip(entries);
}

function extractDocxText(buffer) {
  const xml = unzipEntry(buffer, "word/document.xml").toString("utf8");
  const paragraphs = [];
  for (const para of xml.split(/<\/w:p>/)) {
    const texts = [];
    for (const match of para.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) {
      texts.push(xmlDecode(match[1]));
    }
    const line = texts.join("").trim();
    if (line) paragraphs.push(line);
  }
  return paragraphs.join("\n");
}

function extractDocxHtml(buffer) {
  const xml = unzipEntry(buffer, "word/document.xml").toString("utf8");
  const blocks = [];
  let firstText = true;
  for (const para of xml.split(/<\/w:p>/)) {
    const texts = [];
    for (const match of para.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) {
      texts.push(xmlDecode(match[1]));
    }
    const line = texts.join("").trim();
    if (!line) {
      blocks.push('<p class="contract-blank"><br></p>');
      continue;
    }
    const style = /<w:pStyle[^>]*w:val="([^"]+)"/.exec(para)?.[1] || "";
    const safe = escapeHtmlForDocument(line);
    const isTitle = firstText && line.length <= 80 && /合同|协议|确认书|承诺书|订单|补充协议/.test(line);
    firstText = false;
    if (isTitle || /Title|标题|Heading1|Heading 1/i.test(style)) blocks.push(`<h1>${safe}</h1>`);
    else if (/Heading2|Heading 2|标题2|标题 2/i.test(style) || /^第[一二三四五六七八九十百]+[章节篇]\s*/.test(line) || /^[一二三四五六七八九十]+[、.．]\s*/.test(line)) blocks.push(`<h2>${safe}</h2>`);
    else if (/^第[一二三四五六七八九十百]+条\s*/.test(line) || /^\d+[、.．]\s*/.test(line)) blocks.push(`<p class="clause-line">${safe}</p>`);
    else if (/^（[一二三四五六七八九十\d]+）/.test(line) || /^\([一二三四五六七八九十\d]+\)/.test(line)) blocks.push(`<p class="list-line">${safe}</p>`);
    else blocks.push(`<p>${safe}</p>`);
  }
  return blocks.join("");
}

function runPdfToText(filePath) {
  return new Promise((resolve) => {
    execFile("pdftotext", ["-layout", filePath, "-"], { timeout: 20000 }, (err, stdout) => {
      resolve(err ? "" : stdout);
    });
  });
}

function extractPdfFallback(buffer) {
  const text = buffer.toString("latin1");
  const matches = text.match(/[A-Za-z0-9\u00a0-\uffff][A-Za-z0-9\s.,;:()_\-\/\u00a0-\uffff]{8,}/g) || [];
  return matches.join("\n").replace(/\s{3,}/g, " ").slice(0, 120000);
}

async function extractText(filePath, originalName, buffer) {
  const ext = path.extname(originalName).toLowerCase();
  if ([".txt", ".md", ".csv", ".json", ".html", ".htm", ".xml"].includes(ext)) return buffer.toString("utf8");
  if (ext === ".docx") return extractDocxText(buffer);
  if (ext === ".doc") {
    try {
      const docParser = new DocFormatParser(buffer);
      const parsed = await docParser.parse();
      return (parsed.text || "").slice(0, 120000);
    } catch (e) { console.error("[DocFormatParser] extractText error:", e.message); }
  }
  if (ext === ".pdf") return (await runPdfToText(filePath)) || extractPdfFallback(buffer);
  return buffer.toString("utf8").replace(/\0/g, "").slice(0, 120000);
}

function simpleDocType(filename, text) {
  const haystack = `${filename}\n${text.slice(0, 4000)}`;
  if (/法规|条例|办法|法律|司法解释|规定/.test(haystack)) return "法律法规";
  if (/制度|规章|审批|权限|流程|管理办法/.test(haystack)) return "规章制度";
  if (/模板|范本|示范文本/.test(haystack)) return "合同模板";
  if (/合同|协议|甲方|乙方/.test(haystack)) return "历史合同";
  if (/习惯|惯例|口径|实践|做法/.test(haystack)) return "交易习惯";
  return "其他资料";
}

function chunkText(text, maxLen = 1800) {
  const lines = text.split(/\n+/).map((item) => item.trim()).filter(Boolean);
  const chunks = [];
  let current = "";
  for (const line of lines) {
    if ((current + "\n" + line).length > maxLen && current) {
      chunks.push(current);
      current = "";
    }
    current += `${current ? "\n" : ""}${line}`;
  }
  if (current) chunks.push(current);
  return chunks.slice(0, 30);
}

function summarize(text) {
  const compact = text.replace(/\s+/g, " ").trim();
  return compact.slice(0, 220) + (compact.length > 220 ? "..." : "");
}

function safeJsonFromModel(text) {
  const cleaned = text.replace(/```json|```/g, "").trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    const arrayStart = cleaned.indexOf("[");
    const arrayEnd = cleaned.lastIndexOf("]");
    if (arrayStart !== -1 && arrayEnd !== -1 && arrayEnd > arrayStart) return JSON.parse(cleaned.slice(arrayStart, arrayEnd + 1));
    const objStart = cleaned.indexOf("{");
    const objEnd = cleaned.lastIndexOf("}");
    if (objStart !== -1 && objEnd !== -1 && objEnd > objStart) return JSON.parse(cleaned.slice(objStart, objEnd + 1));
    throw new Error("AI response is not valid JSON");
  }
}

function buildJsonRepairPrompt(content, label = "模型结果") {
  return `请把下面${label}修复为严格可解析的 JSON，只输出 JSON，不要输出解释、Markdown 或代码块。

原始内容：
${String(content || "").slice(0, 120000)}

要求：
1. 保留原有字段含义，不新增无依据内容。
2. 所有字符串必须正确转义。
3. 输出必须是一个 JSON 对象。`;
}

async function callJsonModel(prompt, options = {}) {
  const modelCall = await callChatModel(prompt);
  const raw = modelCall.content;
  try {
    return {
      data: safeJsonFromModel(raw),
      raw,
      repaired: false,
      provider: modelCall.provider,
      model: modelCall.model,
    };
  } catch (parseError) {
    if (!options.repair) throw parseError;
    const repairedCall = await callChatModel(buildJsonRepairPrompt(raw, options.label || "模型结果"));
    const repairedRaw = repairedCall.content;
    return {
      data: safeJsonFromModel(repairedRaw),
      raw,
      repairedRaw,
      repaired: true,
      parseError: parseError.message,
      provider: repairedCall.provider,
      model: repairedCall.model,
      originalProvider: modelCall.provider,
      originalModel: modelCall.model,
    };
  }
}

function loadDimensions() {
  return JSON.parse(fs.readFileSync(DIMENSIONS_PATH, "utf8"));
}

function buildPrompt(documents, dimensions) {
  const dimensionSummary = Object.keys(dimensions)
    .map((name) => `${name}: ${(dimensions[name].categories || []).map((cat) => cat["条款名称"]).slice(0, 16).join("、")}`)
    .join("\n");
  const source = documents
    .map((doc) => {
      const sample = chunkText(doc.text || "", 2200).slice(0, 4).join("\n---\n");
      return `文档ID:${doc.id}\n文件名:${doc.name}\n文档类型:${doc.docType}\n业务领域:${doc.domain || "未指定"}\n内容:\n${sample}`;
    })
    .join("\n\n==========\n\n");

  return `你是企业合同知识库与规则库抽取专家。请从上传资料中提取可用于合同生成和合同审查的规则。

规则分类维度只能从以下四大维度中选择：
${dimensionSummary}

请输出严格 JSON 数组，不要输出解释文本。每个元素包含：
{
  "rule_name": "规则名称",
  "dimension": "四大维度之一",
  "rule_type": "必备条款规则/禁止条款规则/审批规则/风险提示规则/信息追问规则/条款推荐规则/通用规则",
  "contract_type": ["适用合同类型"],
  "business_domain": "业务领域",
  "trigger_condition": "触发条件，用自然语言或简单表达式描述",
  "action": "命中后的动作",
  "risk_level": "高/中/低",
  "priority": 1到100的数字，法规和制度更高",
  "rule_basis": "通用法规/行业惯例/企业自定",
  "source_quote": "不超过120字的原文依据",
  "source_doc_id": "来源文档ID",
  "review_status": "pending_review"
}

抽取要求：
1. 只抽取能用于合同生成、审查判断、审批提醒、风险提示或信息追问的内容。
2. 不要把普通摘要当规则。
3. 每条规则必须有 source_quote。
4. 如果资料内容不足，返回少量高置信规则，不要编造。
5. 优先抽取金额阈值、审批权限、必备条款、禁止承诺、验收付款、违约责任、知识产权、保密、交付、争议解决。
6. rule_basis 必须按来源和内容判断：法律法规、监管规定、国家标准归为通用法规；交易习惯、行业惯例、行业标准归为行业惯例；公司制度、审批口径、自定义规则、内部模板归为企业自定。

上传资料如下：
${source}`;
}

function normalizeModelBaseUrl(baseUrl) {
  return String(baseUrl || "").replace(/\/+$/, "");
}

function buildXiaomiProviders() {
  if (!XIAOMI_API_KEY) return [];
  const configs = [
    { baseUrl: XIAOMI_BASE_URL, model: XIAOMI_MODEL, useBearer: true },
    { baseUrl: DEFAULT_XIAOMI_BASE_URL, model: "mimo-v2.5", useBearer: true },
    { baseUrl: LEGACY_XIAOMI_BASE_URL, model: XIAOMI_MODEL, useBearer: true },
    { baseUrl: LEGACY_XIAOMI_BASE_URL, model: "mimo-v2.5", useBearer: true },
    { baseUrl: "https://api.mimo-v2.com/v1", model: "mimo-v2-pro", useBearer: false },
    { baseUrl: "https://api.mimo-v2.com/v1", model: DEFAULT_XIAOMI_MODEL, useBearer: false },
  ];
  const seen = new Set();
  return configs
    .filter((config) => {
      const key = `${normalizeModelBaseUrl(config.baseUrl)}|${config.model}|${config.useBearer ? "bearer" : "api-key"}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .map((config, index) => ({
      id: "xiaomi",
      name: index === 0 ? "Xiaomi MiMo" : "Xiaomi MiMo",
      baseUrl: config.baseUrl,
      apiKey: XIAOMI_API_KEY,
      model: config.model,
      useBearer: config.useBearer,
      extraHeaders: config.useBearer ? {} : { "api-key": XIAOMI_API_KEY },
    }));
}

function getModelProviders() {
  const xiaomiProviders = buildXiaomiProviders();
  const catalog = {
    xiaomi: xiaomiProviders,
    openai: OPENAI_API_KEY
      ? [{
          id: "openai",
          name: "OpenAI",
          baseUrl: OPENAI_BASE_URL,
          apiKey: OPENAI_API_KEY,
          model: OPENAI_MODEL,
        }]
      : null,
    dashscope: DASHSCOPE_API_KEY
      ? [{
          id: "dashscope",
          name: "DashScope",
          baseUrl: DASHSCOPE_BASE_URL,
          apiKey: DASHSCOPE_API_KEY,
          model: DASHSCOPE_MODEL,
        }]
      : null,
  };
  const providers = [];
  const seen = new Set();
  const addProviderGroup = (id, mode = "push") => {
    const group = Array.isArray(catalog[id]) ? catalog[id] : catalog[id] ? [catalog[id]] : [];
    const list = group.filter(Boolean);
    if (!list.length) return;
    const target = [];
    for (const provider of list) {
      const key = `${provider.id}|${normalizeModelBaseUrl(provider.baseUrl)}|${provider.model}|${provider.useBearer === false ? "api-key" : "bearer"}`;
      if (seen.has(key)) continue;
      seen.add(key);
      target.push(provider);
    }
    if (mode === "unshift") providers.unshift(...target);
    else providers.push(...target);
  };
  for (const id of AI_PROVIDER_PRIORITY) addProviderGroup(id);
  addProviderGroup("openai");
  addProviderGroup("dashscope");
  addProviderGroup("xiaomi", "unshift");
  return providers;
}

function activeModelInfo() {
  const [provider] = getModelProviders();
  if (!provider) return { provider: "None", model: MODEL, hasApiKey: false };
  return { provider: provider.name, model: provider.model, hasApiKey: true };
}

async function callCompatibleModel(provider, prompt) {
  const headers = {
    "content-type": "application/json",
    ...(provider.useBearer === false ? {} : { authorization: `Bearer ${provider.apiKey}` }),
    ...(provider.extraHeaders || {}),
  };
  const response = await fetch(`${normalizeModelBaseUrl(provider.baseUrl)}/chat/completions`, {
    method: "POST",
    headers,
    body: JSON.stringify({
      model: provider.model,
      temperature: 0.1,
      messages: [
        { role: "system", content: "你只输出可解析的 JSON，不输出 Markdown。" },
        { role: "user", content: prompt },
      ],
    }),
  });
  const responseText = await response.text();
  let data = {};
  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch {
    data = {};
  }
  if (!response.ok) throw new Error(data?.error?.message || `${provider.name} request failed: ${response.status}`);
  return {
    content: data.choices?.[0]?.message?.content || "",
    provider: provider.name,
    model: provider.model,
  };
}

async function callChatModel(prompt) {
  const providers = getModelProviders();
  if (!providers.length) throw new Error("Missing XIAOMI_API_KEY, OPENAI_API_KEY or DASHSCOPE_API_KEY environment variable");
  const errors = [];
  for (const provider of providers) {
    try {
      return await callCompatibleModel(provider, prompt);
    } catch (error) {
      errors.push(`${provider.name}: ${error.message}`);
    }
  }
  throw new Error(errors.join(" | "));
}

async function callDashScope(prompt) {
  const call = await callChatModel(prompt);
  return call.content;
}

function normalizeRule(rule, docLookup = new Map()) {
  const sourceDocId = rule.source_doc_id || rule.sourceDocId || "";
  const doc = docLookup.get(sourceDocId);
  const contractType = rule.contract_type || rule.contractType || "通用合同";
  const rawRuleType = rule.rule_type || rule.ruleType || "风险提示规则";
  const ruleBasis = rule.rule_basis || rule.ruleBasis || inferRuleBasis(doc, rule);
  return {
    id: rule.id || `RULE_${Date.now()}_${crypto.randomBytes(3).toString("hex")}`,
    ruleName: rule.rule_name || rule.ruleName || "未命名规则",
    dimension: rule.dimension || "通用必备条款",
    ruleType: rawRuleType === "生成约束规则" ? "通用规则" : rawRuleType,
    scenario: rule.scenario || inferScenario(rule.rule_name || rule.ruleName || "", rule.action || "", rule.trigger_condition || rule.triggerCondition || ""),
    contractType: Array.isArray(contractType) ? contractType : [contractType],
    businessDomain: rule.business_domain || rule.businessDomain || doc?.domain || "通用",
    triggerCondition: rule.trigger_condition || rule.triggerCondition || "",
    action: rule.action || "",
    riskLevel: rule.risk_level || rule.riskLevel || "中",
    priority: Number(rule.priority || 50),
    sourceQuote: rule.source_quote || rule.sourceQuote || "",
    sourceDocId: sourceDocId || doc?.id || "",
    sourceDocName: rule.source_doc_name || rule.sourceDocName || doc?.name || "",
    sourceUrl: rule.source_url || rule.sourceUrl || "",
    reviewStatus: rule.review_status || rule.reviewStatus || "pending_review",
    createdAt: rule.created_at || rule.createdAt || new Date().toISOString(),
    importedAt: rule.importedAt,
    ruleSource: rule.rule_source || rule.ruleSource || (doc || rule.source_doc_name || rule.sourceDocName ? "文档抽取" : "公司规则"),
    useScope: rule.use_scope || rule.useScope || ["生成", "审查"],
    ruleBasis,
    check: rule.check || null,
    presetVersion: rule.presetVersion,
  };
}

function inferRuleBasis(doc, rule = {}) {
  const text = [doc?.docType, doc?.name, rule.rule_name, rule.ruleName, rule.source_doc_name, rule.sourceDocName, rule.source_quote, rule.sourceQuote]
    .filter(Boolean)
    .join(" ");
  if (/法律|法规|条例|办法|司法解释|国家标准|监管|规章/.test(text)) return "通用法规";
  if (/行业|惯例|习惯|标准|协会|实践|做法/.test(text)) return "行业惯例";
  return "企业自定";
}

function inferScenario(...parts) {
  const text = parts.join(" ");
  if (/审批|批准|授权|权限|盖章|签批/.test(text)) return "审批规则";
  if (/付款|支付|价款|发票|结算|预付款|尾款|账期/.test(text)) return "付款规则";
  if (/交付|交货|交接|交付物|里程碑|期限|延期/.test(text)) return "交付规则";
  if (/验收|测试|试运行|确认|验收标准/.test(text)) return "验收规则";
  if (/违约|赔偿|责任|损失|处罚|补偿/.test(text)) return "违约责任规则";
  if (/保密|秘密|数据|隐私|个人信息/.test(text)) return "保密规则";
  if (/知识产权|著作权|专利|商标|源码|许可/.test(text)) return "知识产权规则";
  if (/解除|终止|续签|期限|生效/.test(text)) return "期限终止规则";
  return "通用规则";
}

function ruleFingerprint(rule) {
  return [
    rule.ruleName,
    rule.dimension,
    rule.ruleType,
    (rule.contractType || []).join(","),
    rule.businessDomain,
    rule.triggerCondition,
    rule.action,
    rule.sourceQuote,
  ]
    .join("|")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function addUniqueRules(store, rules) {
  const seen = new Set(store.rules.map(ruleFingerprint));
  const inserted = [];
  const skipped = [];
  for (const rule of rules) {
    const key = ruleFingerprint(rule);
    if (seen.has(key)) {
      skipped.push(rule);
      continue;
    }
    seen.add(key);
    inserted.push(rule);
  }
  store.rules.unshift(...inserted);
  return { inserted, skipped };
}

function fallbackRules(documents, dimensions) {
  const dimensionNames = Object.keys(dimensions);
  const rules = [];
  for (const doc of documents) {
    const text = doc.text || "";
    const snippets = [
      { re: /审批|批准|授权|权限|总经理|法务/g, type: "审批规则", risk: "高" },
      { re: /不得|禁止|严禁|不应|不能/g, type: "禁止条款规则", risk: "高" },
      { re: /必须|应当|需|需要|包含|明确/g, type: "必备条款规则", risk: "中" },
      { re: /违约|赔偿|风险|责任|损失/g, type: "风险提示规则", risk: "中" },
    ];
    for (const item of snippets) {
      const match = item.re.exec(text);
      if (!match) continue;
      const start = Math.max(0, match.index - 60);
      const quote = text.slice(start, match.index + 120).replace(/\s+/g, " ").trim();
      rules.push({
        rule_name: `${doc.docType}${item.type}`,
        dimension: dimensionNames.find((name) => text.includes(name.slice(0, 2))) || dimensionNames[0],
        rule_type: item.type,
        contract_type: [doc.contractType || "通用合同"],
        business_domain: doc.domain || "通用",
        trigger_condition: `合同内容涉及${match[0]}要求时`,
        action: "生成或审查合同时提示并要求人工确认。",
        risk_level: item.risk,
        priority: item.risk === "高" ? 80 : 50,
        source_quote: quote,
        source_doc_id: doc.id,
        rule_basis: inferRuleBasis(doc),
        review_status: "pending_review",
      });
    }
  }
  return rules;
}

function toReusableRule(rule) {
  return {
    rule_id: rule.id,
    rule_name: rule.ruleName,
    dimension: rule.dimension,
    rule_type: rule.ruleType,
    scenario: rule.scenario || inferScenario(rule.ruleName, rule.action, rule.triggerCondition),
    rule_category: ruleManagementCategory(rule),
    contract_type: rule.contractType,
    business_domain: rule.businessDomain,
    trigger_condition: rule.triggerCondition,
    action: rule.action,
    risk_level: rule.riskLevel,
    priority: rule.priority,
    rule_basis: rule.ruleBasis || "企业自定",
    source_quote: rule.sourceQuote,
    source_doc_name: rule.sourceDocName,
    source_url: rule.sourceUrl,
    review_status: rule.reviewStatus,
    use_scope: ["生成", "审查"],
  };
}

function publicTemplate(template) {
  return {
    id: template.id,
    name: template.name,
    keywords: template.keywords,
    outline: template.outline,
    requiredFields: template.requiredFields,
    templateText: template.templateText || "",
    custom: Boolean(template.custom),
  };
}

function fieldKeyFromLabel(label = "") {
  const text = String(label || "").trim();
  if (/^(甲方|甲方名称|甲方全称|采购方|委托方|出租方|授权方|发包方|客户方|许可方|托运方|用人单位)$/.test(text)) return "partyA";
  if (/^(乙方|乙方名称|乙方全称|供应商|服务方|承租方|被许可方|承包方|经销方|承运方|劳动者|受托方)$/.test(text)) return "partyB";
  if (/金额|费用|价款|租金|报酬|工程款|许可费|运费|报名费用/.test(text)) return "amount";
  if (/付款|支付|缴费|结算|还款/.test(text)) return "payment";
  if (/期限|周期|工期|租期|服务期|订阅期限|许可期限/.test(text)) return "term";
  if (/交付|交货|运输|实施|施工|成果|服务成果/.test(text)) return "delivery";
  if (/验收|签收|考核|竣工/.test(text)) return "acceptance";
  if (/保密|数据|安全|隐私/.test(text)) return "confidentiality";
  if (/知识产权|权属|许可范围|模具/.test(text)) return "ipOwnership";
  if (/责任|违约|担保|侵权/.test(text)) return "liability";
  if (/范围|服务|内容|职责|授权区域|处理目的|加工要求|咨询范围/.test(text)) return "serviceScope";
  return text
    .replace(/[^\w\u4e00-\u9fa5]/g, "")
    .slice(0, 24) || "field";
}

function placeholdersFromTemplateText(templateText = "") {
  const labels = [];
  const seen = new Set();
  const re = /{{\s*([^{}]{1,40})\s*}}|【\s*([^【】]{1,40})\s*】/g;
  let match;
  while ((match = re.exec(String(templateText || "")))) {
    const label = String(match[1] || match[2] || "").trim();
    if (!label || seen.has(label)) continue;
    seen.add(label);
    labels.push(label);
  }
  return labels.map((label) => ({
    key: fieldKeyFromLabel(label),
    label,
    question: `请填写${label}`,
  }));
}

function normalizeTemplateFields(fields = [], templateText = "") {
  const sourceFields = Array.isArray(fields) && fields.length ? fields : placeholdersFromTemplateText(templateText);
  const seen = new Set();
  return sourceFields
    .map((field) => {
      if (typeof field === "string") return { key: fieldKeyFromLabel(field), label: field, question: `请填写${field}` };
      const label = String(field.label || field.name || field.key || "").trim();
      if (!label) return null;
      return {
        key: String(field.key || fieldKeyFromLabel(label)).trim(),
        label,
        question: String(field.question || `请填写${label}`).trim(),
      };
    })
    .filter(Boolean)
    .filter((field) => {
      const key = `${field.key}|${field.label}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 24);
}

function normalizeCustomTemplate(template = null) {
  if (!template || typeof template !== "object") return null;
  const templateText = String(template.templateText || template.content || "").trim();
  const name = String(template.name || "自定义模板").trim();
  if (!templateText && !name) return null;
  const requiredFields = normalizeTemplateFields(template.requiredFields || template.fields, templateText);
  return {
    id: String(template.id || `custom_${crypto.createHash("md5").update(`${name}\n${templateText}`).digest("hex").slice(0, 10)}`),
    name,
    keywords: Array.isArray(template.keywords) && template.keywords.length ? template.keywords : [name, "自定义模板"],
    outline: Array.isArray(template.outline) && template.outline.length
      ? template.outline
      : (templateText.match(/第[一二三四五六七八九十]+章[^\n]{0,24}|[一二三四五六七八九十]+、[^\n]{0,24}/g) || ["合同主体", "合同内容", "费用付款", "权利义务", "违约责任", "争议解决"]).slice(0, 12),
    requiredFields,
    templateText,
    custom: true,
  };
}

function templateFingerprint(template = {}) {
  return crypto
    .createHash("md5")
    .update([template.name, template.templateText].filter(Boolean).join("\n"))
    .digest("hex")
    .slice(0, 12);
}

function normalizeStoredTemplate(template = null) {
  const normalized = normalizeCustomTemplate(template);
  if (!normalized) return null;
  const id = String(template.id || normalized.id || `custom_${templateFingerprint(normalized)}`).replace(/[^\w-]/g, "_");
  return {
    ...normalized,
    id,
    custom: true,
    imported: Boolean(template.imported || template.sourceFileName),
    sourceFileName: String(template.sourceFileName || ""),
    createdAt: template.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

function templateTitleFromText(text = "", fileName = "") {
  const lines = String(text || "")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  const title = lines.find((line) => line.length <= 40 && /合同|协议|订单|确认书|承诺书/.test(line));
  if (title) return title.replace(/[【】\[\]()（）]/g, "").trim();
  return path.basename(fileName || "", path.extname(fileName || "")) || "导入合同模板";
}

function inferTemplateFieldsFromText(text = "") {
  const source = String(text || "");
  const fields = [];
  const add = (label) => {
    if (fields.some((field) => field.label === label)) return;
    fields.push({ key: fieldKeyFromLabel(label), label, question: `请填写${label}` });
  };
  for (const label of placeholdersFromTemplateText(source).map((field) => field.label)) add(label);
  const categoryFields = [
    { label: "甲方名称", re: /甲方|采购方|委托方|出租方|发包方|客户方|许可方|托运方/ },
    { label: "乙方名称", re: /乙方|供应商|服务方|承租方|承包方|受托方|承运方/ },
    { label: "统一社会信用代码", re: /统一社会信用代码|社会信用代码|营业执照/ },
    { label: "法定代表人", re: /法定代表人|负责人/ },
    { label: "联系人", re: /联系人|经办人/ },
    { label: "联系电话", re: /联系电话|电话|手机号/ },
    { label: "地址", re: /住所|地址|联系地址|通讯地址/ },
    { label: "合同标的", re: /标的|项目名称|货物|产品|设备|服务内容|合作内容|委托事项/ },
    { label: "服务范围", re: /服务范围|服务内容|履行范围|开发范围|授权范围|施工范围|处理目的/ },
    { label: "数量", re: /数量|规格|型号|清单/ },
    { label: "质量标准", re: /质量|标准|技术要求|规格要求/ },
    { label: "合同金额", re: /金额|费用|价款|租金|报酬|工程款|运费|许可费|人民币|￥|¥/ },
    { label: "付款方式", re: /付款|支付|结算|账期|发票|收款/ },
    { label: "交付安排", re: /交付|交货|运输|实施|提交|发货|配送/ },
    { label: "验收标准", re: /验收|签收|考核|确认|试运行/ },
    { label: "合同期限", re: /期限|周期|工期|服务期|租期|授权期限|有效期/ },
    { label: "保密要求", re: /保密|商业秘密|数据安全|个人信息|隐私/ },
    { label: "知识产权", re: /知识产权|著作权|专利|商标|源码|成果归属/ },
    { label: "违约责任", re: /违约|赔偿|责任|损失|解除|终止/ },
    { label: "争议解决", re: /争议|管辖|仲裁|诉讼|法院/ },
    { label: "签订日期", re: /签订日期|签署日期|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日/ },
  ];
  for (const item of categoryFields) {
    if (item.re.test(source)) add(item.label);
  }
  if (!fields.some((field) => field.key === "partyA")) add("甲方名称");
  if (!fields.some((field) => field.key === "partyB")) add("乙方名称");
  return fields.slice(0, 28);
}

function fallbackTemplateTextFromContract(text = "", fileName = "") {
  let template = String(text || "").replace(/\r/g, "").trim().slice(0, 120000);
  if (!template) return "";
  template = template
    .replace(/(甲方(?:（[^）]{0,12}）)?[：:])\s*[^\n]{2,80}/g, "$1【甲方名称】")
    .replace(/(乙方(?:（[^）]{0,12}）)?[：:])\s*[^\n]{2,80}/g, "$1【乙方名称】")
    .replace(/(统一社会信用代码[：:])\s*[0-9A-Z]{8,30}/g, "$1【统一社会信用代码】")
    .replace(/(法定代表人[：:])\s*[^\n]{1,30}/g, "$1【法定代表人】")
    .replace(/(联系人[：:])\s*[^\n]{1,40}/g, "$1【联系人】")
    .replace(/(联系电话|电话)[：:]\s*[0-9\-+\s]{6,30}/g, "$1：【联系电话】")
    .replace(/(地址|住所地)[：:]\s*[^\n]{2,120}/g, "$1：【地址】")
    .replace(/((?:服务内容|服务范围|合作内容|委托事项|开发范围|施工范围|授权范围)[：:])\s*[^\n]{2,200}/g, "$1【服务范围】")
    .replace(/((?:合同标的|标的|项目名称|产品|设备|货物)[：:])\s*[^\n]{2,200}/g, "$1【合同标的】")
    .replace(/((?:付款|付款方式|支付|结算方式)[：:])\s*[^\n]{2,200}/g, "$1【付款方式】")
    .replace(/((?:交付|交付安排|交货|运输|实施安排)[：:])\s*[^\n]{2,200}/g, "$1【交付安排】")
    .replace(/((?:验收|验收标准|质量标准|考核标准)[：:])\s*[^\n]{2,200}/g, "$1【验收标准】")
    .replace(/((?:合同期限|服务期限|租赁期限|工期|授权期限|有效期)[：:])\s*[^\n]{1,120}/g, "$1【合同期限】")
    .replace(/((?:保密|保密要求|数据安全)[：:])\s*[^\n]{2,200}/g, "$1【保密要求】")
    .replace(/((?:知识产权|成果归属|源码归属)[：:])\s*[^\n]{2,200}/g, "$1【知识产权】")
    .replace(/((?:违约责任|赔偿责任|责任承担)[：:])\s*[^\n]{2,220}/g, "$1【违约责任】")
    .replace(/((?:争议解决|管辖|仲裁|诉讼)[：:])\s*[^\n]{2,200}/g, "$1【争议解决】")
    .replace(/人民币\s*\d[\d,]*(?:\.\d+)?\s*(?:万元|元|亿元)?|￥\s*\d[\d,]*(?:\.\d+)?/g, "【合同金额】")
    .replace(/\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日/g, "【签订日期】");
  if (!/【甲方名称】/.test(template)) template = `甲方：【甲方名称】\n乙方：【乙方名称】\n\n${template}`;
  if (!/合同|协议|订单|确认书/.test(template.slice(0, 80))) template = `${templateTitleFromText(text, fileName)}\n\n${template}`;
  return template;
}

function buildTemplateImportPrompt(text = "", fileName = "") {
  return `你是合同模板工程师。请把用户导入的合同抽取成可复用的合同生成模板，不要保留具体交易方、具体金额、具体日期等一次性信息。

文件名：${fileName || "未命名"}

合同原文：
${String(text || "").slice(0, 120000)}

请输出严格 JSON：
{
  "name": "模板名称",
  "keywords": ["关键词"],
  "outline": ["章节或模块"],
  "requiredFields": [
    { "key": "partyA", "label": "甲方名称", "question": "甲方是谁？" }
  ],
  "templateText": "带【字段名】占位符的完整模板正文"
}

要求：
1. templateText 保留原合同的章节结构、条款层级和常用合同格式。
2. 将甲方、乙方、金额、日期、标的、服务范围、交付、付款、验收、期限、联系人、账户、附件等可变信息替换为【字段名】。
3. requiredFields 必须覆盖 templateText 中所有关键占位符，控制在 8-28 个字段。
4. key 优先使用 partyA、partyB、subject、serviceScope、amount、payment、delivery、acceptance、term、confidentiality、ipOwnership、liability。
5. 不要输出 Markdown，不要输出 JSON 之外的文字。`;
}

function fallbackImportedTemplate(text = "", fileName = "") {
  const templateText = fallbackTemplateTextFromContract(text, fileName);
  const requiredFields = normalizeTemplateFields(inferTemplateFieldsFromText(templateText), templateText);
  const outline = (templateText.match(/第[一二三四五六七八九十]+章[^\n]{0,24}|第[一二三四五六七八九十]+条[^\n]{0,24}|[一二三四五六七八九十]+、[^\n]{0,24}/g) || ["合同主体", "合同标的", "费用付款", "交付验收", "权利义务", "违约责任", "争议解决"]).slice(0, 12);
  const name = templateTitleFromText(text, fileName);
  return normalizeStoredTemplate({
    name: name.endsWith("模板") ? name : `${name}模板`,
    keywords: [name, "导入模板", ...outline.slice(0, 3)],
    outline,
    requiredFields,
    templateText,
    imported: true,
    sourceFileName: fileName,
  });
}

function normalizeImportedTemplate(modelData = {}, text = "", fileName = "") {
  const fallback = fallbackImportedTemplate(text, fileName);
  const templateText = fallbackTemplateTextFromContract(
    String(modelData.templateText || modelData.template_text || fallback.templateText || "").trim(),
    fileName
  );
  const requiredFields = normalizeTemplateFields(modelData.requiredFields || modelData.fields || fallback.requiredFields, templateText);
  const outline = Array.isArray(modelData.outline) && modelData.outline.length
    ? modelData.outline.map((item) => String(item || "").trim()).filter(Boolean).slice(0, 12)
    : fallback.outline;
  return normalizeStoredTemplate({
    name: String(modelData.name || fallback.name || "导入合同模板").trim(),
    keywords: Array.isArray(modelData.keywords) && modelData.keywords.length ? modelData.keywords : fallback.keywords,
    outline,
    requiredFields,
    templateText,
    imported: true,
    sourceFileName: fileName,
  });
}

function templatesForPayload(payload = {}) {
  const custom = normalizeCustomTemplate(payload.customTemplate);
  if (!custom) return CONTRACT_TEMPLATES;
  const filtered = CONTRACT_TEMPLATES.filter((template) => template.id !== custom.id);
  return [...filtered, custom];
}

function scoreContractTemplates(description = "", templates = CONTRACT_TEMPLATES) {
  const text = String(description || "").toLowerCase();
  return templates.map((template) => {
    let score = 0;
    if (text.includes(template.name.toLowerCase())) score += 10;
    for (const keyword of template.keywords) {
      if (text.includes(keyword.toLowerCase())) score += 3;
    }
    for (const item of template.outline) {
      if (text.includes(item.toLowerCase())) score += 1;
    }
    return { template, score };
  }).sort((a, b) => b.score - a.score);
}

function matchContractTemplate(description = "", templates = CONTRACT_TEMPLATES) {
  const scored = scoreContractTemplates(description, templates);
  return scored[0].score > 0 ? scored[0].template : templates[0];
}

function templateMatchConfidence(score = 0) {
  if (!score) return 0;
  return Math.max(42, Math.min(96, 48 + score * 8));
}

function publicTemplateCandidate(item) {
  return {
    ...publicTemplate(item.template),
    score: item.score,
    confidence: templateMatchConfidence(item.score),
  };
}

function chooseContractTemplate(description = "", requestedId = "", templates = CONTRACT_TEMPLATES) {
  const requestedTemplate = templates.find((item) => item.id === requestedId);
  const scored = scoreContractTemplates(description, templates);
  const [best] = scored;
  const candidates = scored
    .filter((item) => item.score > 0)
    .slice(0, 3)
    .map(publicTemplateCandidate);
  if (best?.score > 0) {
    return {
      template: best.template,
      requestedTemplate,
      templateSwitched: Boolean(requestedTemplate && requestedTemplate.id !== best.template.id),
      matchedScore: best.score,
      matchConfidence: templateMatchConfidence(best.score),
      templateCandidates: candidates,
    };
  }
  return {
    template: requestedTemplate || templates[0],
    requestedTemplate,
    templateSwitched: false,
    matchedScore: 0,
    matchConfidence: 0,
    templateCandidates: requestedTemplate ? [publicTemplateCandidate({ template: requestedTemplate, score: 0 })] : [],
  };
}

const INDUSTRY_RULE_DOMAINS = ["软件外包", "系统集成", "制造业"];
const INDUSTRY_RULE_DOMAIN_ALIASES = ["软件外包", "系统集成", "制造业", "工业制造业"];
const MANAGEMENT_CATEGORIES = ["通用规则", "行业预设", "企业自定"];

function isEnterpriseRule(rule = {}) {
  return rule.ruleBasis === "企业自定" || rule.ruleSource === "公司规则" || rule.sourceDocName === "公司自己的规则";
}

function isScopedIndustryPresetRule(rule = {}) {
  const id = String(rule.id || "");
  return (
    /^PRESET_INDUSTRY_(SOFTWARE|INTEGRATION|MANUFACTURING)/.test(id) ||
    INDUSTRY_RULE_DOMAIN_ALIASES.includes(rule.businessDomain)
  );
}

function ruleManagementCategory(rule = {}) {
  if (isEnterpriseRule(rule)) return "企业自定";
  if (isScopedIndustryPresetRule(rule)) return "行业预设";
  if (MANAGEMENT_CATEGORIES.includes(rule.ruleCategory)) return rule.ruleCategory;
  return "通用规则";
}

function publicRule(rule = {}) {
  return {
    ...rule,
    ruleCategory: ruleManagementCategory(rule),
  };
}

function enterpriseRulesOnly(rules = []) {
  return (rules || []).filter((rule) => ruleManagementCategory(rule) === "企业自定" || isEnterpriseRule(rule));
}

function ruleSourceIsEnterprise(source = "", rules = []) {
  const text = String(source || "");
  if (/企业|公司|自定|自定义/.test(text)) return true;
  return enterpriseRulesOnly(rules).some((rule) => {
    const candidates = [rule.ruleName, rule.sourceDocName, rule.ruleBasis].filter(Boolean);
    return candidates.some((value) => text && text.includes(String(value)));
  });
}

function inferContractIndustry(text = "") {
  const source = String(text || "");
  if (/系统集成|集成项目|信息化项目|软硬件|接口对接|联调|等保|等级保护|关键信息基础设施|关基|网络设备|服务器|弱电|割接/.test(source)) {
    return "系统集成";
  }
  if (/制造|生产|加工|OEM|ODM|工业|设备制造|零部件|BOM|图纸|工艺|质检|质量标准|特种设备|原材料|来料|封样|首件/.test(source)) {
    return "制造业";
  }
  if (/软件外包|软件开发|定制开发|系统开发|APP|小程序|平台开发|源代码|源码|软著|开源|SLA|UAT|运维|部署上线/.test(source)) {
    return "软件外包";
  }
  return "";
}

function inferContractType(text = "", template = null) {
  const source = String(text || "");
  if (/劳动合同|用人单位|劳动者|试用期|社保|竞业限制/.test(source)) return "劳动合同";
  if (/系统集成|信息化项目|软硬件集成|联调/.test(source)) return "系统集成合同";
  if (/软件开发|软件外包|定制开发|源代码|源码|APP|小程序/.test(source)) return "软件开发合同";
  if (/制造|生产|加工|OEM|ODM|设备制造|零部件/.test(source)) return "制造合同";
  if (template?.name) return template.name;
  const matched = scoreContractTemplates(source)[0];
  return matched?.score > 0 ? matched.template.name : "通用合同";
}

function inferContractProfile(text = "", template = null) {
  const contractType = inferContractType(text, template);
  const industry = inferContractIndustry(text);
  return {
    contractType,
    industry,
    industryOptions: INDUSTRY_RULE_DOMAINS,
    ruleCategories: ["通用规则", "行业预设", "企业自定"],
  };
}

function ruleIndustry(rule = {}) {
  const id = String(rule.id || "");
  if (/PRESET_INDUSTRY_SOFTWARE/.test(id) || rule.businessDomain === "软件外包") return "软件外包";
  if (/PRESET_INDUSTRY_INTEGRATION/.test(id) || rule.businessDomain === "系统集成") return "系统集成";
  if (/PRESET_INDUSTRY_MANUFACTURING/.test(id) || rule.businessDomain === "工业制造业" || rule.businessDomain === "制造业") return "制造业";
  return "";
}

function ruleContractTypes(rule = {}) {
  return Array.isArray(rule.contractType) ? rule.contractType.filter(Boolean) : [rule.contractType || "通用合同"];
}

function ruleMatchesProfile(rule = {}, profile = {}) {
  const category = ruleManagementCategory(rule);
  const industry = ruleIndustry(rule);
  const types = ruleContractTypes(rule);
  const commonScope = !industry && (!types.length || types.includes("通用合同"));
  const typeMatches = types.includes("通用合同") || types.includes(profile.contractType);
  const industryMatches = Boolean(profile.industry && industry && profile.industry === industry);
  const laborMatches = profile.contractType === "劳动合同" && /劳动/.test([rule.businessDomain, rule.ruleName, ...types].join(" "));

  if (category === "通用规则") return commonScope || typeMatches || laborMatches;
  if (category === "行业预设") return industryMatches || laborMatches || (commonScope && typeMatches);
  if (category === "企业自定") return commonScope || industryMatches || typeMatches;
  return false;
}

function selectRulesForContract(store, profile = {}, limit = 90) {
  const rules = getUsableRules(store);
  const selected = [];
  const seen = new Set();
  const categoryOrder = { "通用规则": 0, "行业预设": 1, "企业自定": 2 };
  for (const rule of rules) {
    if (!ruleMatchesProfile(rule, profile)) continue;
    const key = rule.id || ruleFingerprint(rule);
    if (seen.has(key)) continue;
    seen.add(key);
    selected.push(rule);
  }
  return selected
    .sort((a, b) => {
      const categoryRank = (categoryOrder[ruleManagementCategory(a)] ?? 9) - (categoryOrder[ruleManagementCategory(b)] ?? 9);
      if (categoryRank) return categoryRank;
      return Number(b.priority || 0) - Number(a.priority || 0);
    })
    .slice(0, limit);
}

function getUsableRules(store) {
  const active = store.rules.filter((rule) => rule.reviewStatus === "active");
  const fallback = store.rules.filter((rule) => !["inactive", "rejected"].includes(rule.reviewStatus));
  return (active.length ? active : fallback)
    .slice()
    .sort((a, b) => Number(b.priority || 0) - Number(a.priority || 0));
}

function summarizeRulesForPrompt(rules, max = 35) {
  return rules
    .slice(0, max)
    .map((rule, index) => {
      const scope = (rule.contractType || []).join("、") || "通用合同";
      return `${index + 1}. ${ruleManagementCategory(rule)}｜${rule.ruleName}｜${rule.ruleBasis || "企业自定"}｜${rule.ruleType}｜${rule.businessDomain || "通用"}｜${scope}｜${rule.action || rule.triggerCondition || ""}`;
    })
    .join("\n");
}

function fieldKeywords(field = {}) {
  const key = String(field.key || "");
  const label = String(field.label || "");
  const text = `${key} ${label}`;
  if (/delivery|交付|交货|履行|供货|清单|设备|产品/.test(text)) return ["交付", "交货", "交付物", "交付清单", "供货清单", "设备清单", "产品清单", "附件", "里程碑"];
  if (/acceptance|验收|质量|标准/.test(text)) return ["验收", "验收标准", "质量标准", "整改", "复验", "异议", "验收单"];
  if (/payment|付款|支付|结算|发票/.test(text)) return ["付款", "支付", "结算", "发票", "账期", "预付款", "尾款", "含税", "不含税"];
  if (/amount|金额|价款|费用|总价/.test(text)) return ["金额", "价款", "总价", "费用", "报价", "人民币", "元", "含税"];
  if (/subject|scope|标的|范围|内容|服务/.test(text)) return ["标的", "服务内容", "服务范围", "产品", "设备", "项目", "工作内容"];
  if (/confidentiality|保密|数据/.test(text)) return ["保密", "商业秘密", "数据", "个人信息", "隐私"];
  if (/ip|知识产权|成果|源码/.test(text)) return ["知识产权", "成果", "源码", "著作权", "专利", "归属"];
  if (/liability|违约|责任|赔偿/.test(text)) return ["违约", "责任", "赔偿", "损失", "违约金"];
  if (/term|期限|周期|生效/.test(text)) return ["期限", "周期", "生效", "终止", "完成时间"];
  return [key, label].filter(Boolean);
}

function lineSnippets(text = "") {
  return String(text || "")
    .split(/\n+|。|；|;/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter((line) => line.length >= 8 && line.length <= 360);
}

function selectKnowledgeSnippetsForGeneration(store = {}, query = "", options = {}) {
  const field = options.field || {};
  const queryTerms = String(query || "").match(/[\u4e00-\u9fa5A-Za-z0-9]{2,}/g) || [];
  const keywords = [...fieldKeywords(field), ...queryTerms]
    .map((item) => String(item || "").trim())
    .filter(Boolean)
    .slice(0, 36);
  const scored = [];
  for (const doc of store.documents || []) {
    const docHeader = [doc.name, doc.docType, doc.contractType, doc.domain, doc.summary].filter(Boolean).join(" ");
    for (const snippet of lineSnippets([doc.summary, doc.text].filter(Boolean).join("\n")).slice(0, 220)) {
      const source = `${doc.name || "上传文档"}${doc.domain ? `｜${doc.domain}` : ""}`;
      const haystack = `${docHeader} ${snippet}`;
      let score = 0;
      for (const keyword of keywords) {
        if (keyword && haystack.includes(keyword)) score += keyword.length >= 3 ? 3 : 1;
      }
      if (/交付清单|供货清单|设备清单|产品清单|附件|报价单|验收单/.test(snippet)) score += 6;
      if (/金额|价款|付款|验收|交付|质量|规格|数量|型号|参数/.test(snippet)) score += 2;
      if (score > 0) scored.push({ score, source, text: snippet.slice(0, 320) });
    }
  }
  const seen = new Set();
  return scored
    .sort((a, b) => b.score - a.score)
    .filter((item) => {
      const key = compactText(item.text).slice(0, 80);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, options.limit || 8);
}

function summarizeKnowledgeSnippetsForPrompt(snippets = [], max = 2400) {
  const text = (snippets || [])
    .map((item, index) => `${index + 1}. 来源：${item.source}\n片段：${item.text}`)
    .join("\n");
  return text.slice(0, max) || "无可用上传附件片段。";
}

function summarizeRuleSelection(profile = {}, rules = []) {
  const counts = rules.reduce((acc, rule) => {
    const category = ruleManagementCategory(rule);
    acc[category] = (acc[category] || 0) + 1;
    return acc;
  }, {});
  return `合同类型：${profile.contractType || "通用合同"}；所属行业：${profile.industry || "未命中特定行业"}；通用规则和行业预设已默认执行；企业自定规则 ${counts["企业自定"] || 0} 条。`;
}

function summarizeReviewPrecheckForPrompt(issues = [], max = 16) {
  const list = Array.isArray(issues) ? issues : [];
  if (!list.length) return "本地预检未命中明确问题，仍需模型按六层审查方法完整审查。";
  return list
    .slice(0, max)
    .map((issue, index) => {
      const risk = issue.risk_level || issue.riskLevel || "中";
      const title = issue.title || "待复核问题";
      const quote = issue.quote ? `｜定位：${String(issue.quote).slice(0, 120)}` : "";
      const problem = issue.problem ? `｜问题：${String(issue.problem).slice(0, 160)}` : "";
      return `${index + 1}. ${risk}｜${issue.category || "风险问题"}｜${title}${quote}${problem}`;
    })
    .join("\n");
}

function normalizeAnswers(answers = {}) {
  return Object.fromEntries(
    Object.entries(answers || {})
      .map(([key, value]) => [key, String(value || "").trim()])
      .filter(([, value]) => value)
  );
}

function inferAnswersFromDescription(description = "") {
  const text = String(description || "");
  const inferred = {};
  const partyA = /(?:甲方|采购方|委托方|出租方|出借方|用人单位)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const partyB = /(?:乙方|供应商|服务方|承租方|借款方|劳动者)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const amount = /(?:金额|总价|价款|费用|租金|借款|薪资)[为是：:]?\s*([0-9,.]+万?元(?:\/[年月日季])?)/.exec(text)?.[1] || /([0-9,.]+万?元)/.exec(text)?.[1];
  const term = /(?:期限|周期|租期|服务期|合同期)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const payment = /(?:付款|支付|还款|结算)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const delivery = /(?:交付|交货|发货|交接)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  if (partyA) inferred.partyA = partyA.trim();
  if (partyB) inferred.partyB = partyB.trim();
  if (amount) inferred.amount = amount.trim();
  if (term) inferred.term = term.trim();
  if (payment) inferred.payment = payment.trim();
  if (delivery) inferred.delivery = delivery.trim();
  return inferred;
}

function detectMissingFields(template, answers) {
  return template.requiredFields.filter((field) => !String(answers[field.key] || "").trim());
}

function memoryMatchesTemplate(item, template) {
  return !item.templateId || item.templateId === template.id || item.templateName === template.name;
}

function recordGenerationMemory(store, template, description, answers) {
  const now = new Date().toISOString();
  const additions = [];
  const addItem = (fieldKey, fieldLabel, value, source) => {
    const clean = String(value || "").trim();
    if (!clean) return;
    const exists = store.contractMemory.some(
      (item) => item.templateId === template.id && item.fieldKey === fieldKey && item.value === clean
    );
    if (exists) return;
    additions.push({
      id: `MEM_${Date.now()}_${crypto.randomBytes(3).toString("hex")}`,
      templateId: template.id,
      templateName: template.name,
      fieldKey,
      fieldLabel,
      value: clean,
      source,
      description: String(description || "").slice(0, 500),
      createdAt: now,
    });
  };

  addItem("description", "业务描述", description, "用户描述");
  for (const field of template.requiredFields) {
    addItem(field.key, field.label, answers[field.key], "用户补充");
  }
  store.contractMemory.unshift(...additions);
  store.contractMemory = store.contractMemory.slice(0, 500);
}

function cleanFieldSuggestionText(text = "", field = {}) {
  let clean = String(text || "")
    .replace(/\s+/g, " ")
    .replace(/^(建议|请|需要|可以|可考虑|应当|系统会|根据上下文|结合规则库|暂无历史记录)[：:，,\s]*/g, "")
    .trim();
  if (!clean || /建议人工|人工补充|实际情况填写|请按|无需填写|待填写/.test(clean)) return "";
  if (/\.(docx|doc|pdf|txt|json|xlsx?)$/i.test(clean) || /^中华人民共和国.+法/.test(clean)) return "";
  if (/<[^>]+>|w:|xml|style=|font-family/i.test(clean)) return "";
  if (/^(按|以)?实际情况/.test(clean)) return "";
  if (field?.label && clean === field.label) return "";
  return clean.replace(/[。；;]+$/g, "").slice(0, 100);
}

function buildFieldSuggestions(store, template, field, answers, description) {
  const buckets = {
    history: [],
    knowledge: [],
    model: [],
  };
  const seen = {
    history: new Set(),
    knowledge: new Set(),
    model: new Set(),
  };
  const add = (bucket, text, source) => {
    const clean = cleanFieldSuggestionText(text, field);
    if (!clean || clean.length < 2) return;
    const key = clean.toLowerCase();
    if (seen[bucket].has(key)) return;
    seen[bucket].add(key);
    buckets[bucket].push({ text: clean, source });
  };

  for (const item of store.contractMemory) {
    if (item.fieldKey === field.key && memoryMatchesTemplate(item, template)) add("history", item.value, "历史常用");
  }

  const relatedRules = store.rules
    .filter((rule) => !["inactive", "rejected"].includes(rule.reviewStatus))
    .sort((a, b) => Number(b.priority || 0) - Number(a.priority || 0));
  const docValues = store.documents
    .flatMap((doc) => [doc.contractType, doc.domain, doc.summary])
    .filter(Boolean);

  if (["subject", "serviceScope"].includes(field.key)) {
    for (const value of docValues) add("knowledge", value, "知识库推荐");
  }
  if (["payment", "amount"].includes(field.key)) {
    for (const rule of relatedRules.filter((rule) => /付款|支付|价款|发票|结算|账期/.test([rule.scenario, rule.ruleName, rule.action].join(" ")))) {
      add("knowledge", rule.action || rule.triggerCondition || rule.ruleName, "知识库推荐");
    }
  }
  if (["delivery", "acceptance"].includes(field.key)) {
    for (const rule of relatedRules.filter((rule) => /交付|验收|测试|确认|交货/.test([rule.scenario, rule.ruleName, rule.action].join(" ")))) {
      add("knowledge", rule.action || rule.triggerCondition || rule.ruleName, "知识库推荐");
    }
  }
  if (field.key === "confidentiality") {
    for (const rule of relatedRules.filter((rule) => /保密|秘密|数据|隐私|个人信息/.test([rule.scenario, rule.ruleName, rule.action].join(" ")))) {
      add("knowledge", rule.action || rule.triggerCondition || rule.ruleName, "知识库推荐");
    }
  }
  if (field.key === "ipOwnership") {
    for (const rule of relatedRules.filter((rule) => /知识产权|著作权|专利|商标|源码/.test([rule.scenario, rule.ruleName, rule.action].join(" ")))) {
      add("knowledge", rule.action || rule.triggerCondition || rule.ruleName, "知识库推荐");
    }
  }
  if (field.key === "liability") {
    for (const rule of relatedRules.filter((rule) => /违约|赔偿|责任|损失|处罚/.test([rule.scenario, rule.ruleName, rule.action].join(" ")))) {
      add("knowledge", rule.action || rule.triggerCondition || rule.ruleName, "知识库推荐");
    }
  }
  if (field.key === "term") {
    for (const rule of relatedRules.filter((rule) => /期限|终止|解除|续签|生效/.test([rule.scenario, rule.ruleName, rule.action].join(" ")))) {
      add("knowledge", rule.action || rule.triggerCondition || rule.ruleName, "知识库推荐");
    }
  }

  const fallbackByField = {
    partyA: ["我方公司（以营业执照登记名称为准）", "本公司", "甲方为需求提出方"],
    partyB: ["对方公司（以营业执照登记名称为准）", "供应商/服务方公司全称", "乙方为实际履约方"],
    amount: ["按报价单或订单确认金额", "固定总价，含税", "分阶段按验收结果付款"],
    payment: ["合同签署后支付预付款，验收后支付尾款", "按月结算，收到合规发票后付款", "付款前需完成验收并提供发票"],
    delivery: ["按项目里程碑分阶段交付", "交付地点以甲方指定地点为准", "逾期交付需承担违约责任"],
    acceptance: ["甲方在收到交付物后进行验收", "验收不合格的，乙方应限期整改", "以双方确认的验收标准为准"],
    confidentiality: ["双方对合作中获悉的商业秘密承担保密义务", "保密期限为合同终止后两年", "未经披露方书面同意不得向第三方披露"],
    ipOwnership: ["项目成果知识产权归甲方所有", "乙方保留原有工具和通用技术权利", "源码、文档和交付成果应一并移交"],
    liability: ["违约方应赔偿守约方实际损失", "逾期履行按日承担违约金", "严重违约时守约方有权解除合同"],
    term: ["自双方盖章或签字之日起生效", "服务期限以双方确认的项目计划为准", "合同履行完毕且双方权利义务结清后终止"],
  };
  const fallbackValues = fallbackByField[field.key] || [`${field.label}以双方书面确认为准`, `${field.label}按合同目的和履行安排确定`, `${field.label}由双方另行书面确认`];
  if (!buckets.history.length) add("history", fallbackValues[0], "历史常用");
  if (!buckets.knowledge.length) add("knowledge", fallbackValues[1] || fallbackValues[0], "知识库推荐");

  const contextBasedByField = {
    partyA: "甲方为本合同项下需求提出方、付款方或成果接收方",
    partyB: "乙方为本合同项下服务、交付或履约义务承担方",
    amount: "合同总价以双方确认的报价单或订单为准，币种、含税口径和发票类型同步列明",
    payment: "价款分为预付款、交付验收款和尾款，并与发票、验收合格及无争议付款资料绑定",
    delivery: "乙方按项目里程碑提交交付成果，交付地点和接收方式以甲方书面确认为准",
    acceptance: "甲方在收到交付物后按约定标准验收，异议期内提出问题的乙方限期整改并复验",
    confidentiality: "双方对商业秘密、技术资料、客户信息和个人信息承担保密义务",
    ipOwnership: "定制开发成果、源码、文档及相关知识产权归甲方所有，乙方既有工具除外",
    liability: "违约方应赔偿守约方实际损失，逾期履行按日承担违约责任",
    term: "合同自双方签字或盖章之日起生效，履行期限以项目计划或订单约定为准",
  };
  add("model", contextBasedByField[field.key] || `${field.label}以合同目的、交易背景和双方书面确认内容为准`, "上下文/规则补全");

  const suggestions = [
    buckets.history[0],
    buckets.knowledge[0],
    buckets.model[0],
  ].filter(Boolean);

  for (const value of fallbackValues) {
    if (suggestions.length >= 3) break;
    add("model", value, "上下文/规则补全");
    if (buckets.model.at(-1)) suggestions.push(buckets.model.at(-1));
  }

  return suggestions.slice(0, 3);
}

function answerForTemplateField(label = "", answers = {}, requiredFields = []) {
  const directKey = fieldKeyFromLabel(label);
  if (answers[directKey]) return answers[directKey];
  const matched = requiredFields.find((field) => field.label === label || field.key === label);
  if (matched && answers[matched.key]) return answers[matched.key];
  return answers[label] || "";
}

function fillTemplateText(templateText = "", answers = {}, requiredFields = []) {
  const replaceOne = (rawLabel) => {
    const label = String(rawLabel || "").trim();
    return answerForTemplateField(label, answers, requiredFields) || `【${label || "待补充"}】`;
  };
  return String(templateText || "")
    .replace(/{{\s*([^{}]{1,40})\s*}}/g, (_, label) => replaceOne(label))
    .replace(/【\s*([^【】]{1,40})\s*】/g, (_, label) => replaceOne(label));
}

function fallbackContractDraft(template, answers, rules) {
  if (template.templateText) {
    return fillTemplateText(template.templateText, answers, template.requiredFields);
  }
  const parties = [answers.partyA, answers.partyB].filter(Boolean).join(" 与 ") || "合同双方";
  const ruleClauses = rules
    .slice(0, 8)
    .map((rule, index) => `（${index + 1}）双方应遵守“${rule.ruleName}”相关要求，${rule.action || rule.triggerCondition || "并在履行过程中留存必要记录。"}`)
    .join("\n");
  const fieldLines = template.requiredFields
    .map((field) => `- ${field.label}：${answers[field.key] || "【待补充】"}`)
    .join("\n");
  return `# ${template.name}\n\n签约主体：${parties}\n\n## 第一章 核心信息\n${fieldLines}\n\n## 第二章 合同标的与履行\n第一条 双方应按照上述核心信息履行本合同项下义务。\n第二条 合同标的、交付范围、服务内容、质量标准和履行期限以双方书面确认内容为准。\n第三条 一方变更履行内容、时间、地点或联系人信息的，应提前书面通知另一方并取得确认。\n\n## 第三章 价款、付款与验收\n第四条 合同金额、计价方式、发票要求和付款节点以本合同核心信息及附件约定为准。\n第五条 收款方应按约定提供合规发票、付款申请资料及验收证明文件。\n第六条 验收不合格的，履行方应在合理期限内整改，并重新提交验收。\n\n## 第四章 合规与履约要求\n${ruleClauses || "第七条 双方应遵守适用法律法规、行业惯例及各自内部审批要求。"}\n\n## 第五章 违约责任与争议解决\n第八条 任一方违反本合同约定造成损失的，应承担继续履行、采取补救措施或赔偿损失等违约责任。\n第九条 双方因本合同产生争议的，应先友好协商；协商不成的，按双方书面约定的管辖方式解决。\n\n## 第六章 签署\n甲方（盖章）：____________________\n授权代表：____________________\n签署日期：____年__月__日\n\n乙方（盖章）：____________________\n授权代表：____________________\n签署日期：____年__月__日`;
}

function cleanGeneratedContractDraft(value = "") {
  return String(value || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, "")
    .replace(/```(?:json|markdown|md|text)?/gi, "")
    .replace(/```/g, "")
    .replace(/(?:^|\n)\s*#{0,6}\s*(?:思考过程|推理过程|分析过程|生成过程|规则命中|规则预检|生成说明|起草说明)[\s\S]*?(?=\n\s*#{1,6}\s*(?:合同|第一|一、|1[.、])|$)/g, "\n")
    .split("\n")
    .filter((line) => !/^\s*(?:思考|分析|推理|生成说明|起草说明|规则命中|规则预检|我将|我会先)[:：]/.test(line))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildContractGenerationPrompt(template, description, answers, rules, profile = {}, options = {}) {
  return `你是企业合同起草助手。请根据用户意图、模板、已补充信息和规则库生成一份结构完整的中文合同草稿。

用户描述：
${description}

匹配模板：${template.name}
模板大纲：${template.outline.join("、")}
${template.templateText ? `\n模板正文（应优先保留章节结构和字段语义）：\n${template.templateText.slice(0, 120000)}\n` : ""}
${options.existingDraft ? `\n用户当前草稿/已编辑版本（重新生成时参考）：\n${String(options.existingDraft).slice(0, 120000)}\n` : ""}

已补充信息：
${JSON.stringify(answers, null, 2)}

上传附件/知识库可参考片段：
${summarizeKnowledgeSnippetsForPrompt(options.knowledgeSnippets || [], 3600)}

合同识别与规则注入：
${summarizeRuleSelection(profile, rules)}

可用规则（已按合同类型、所属行业和规则类别筛选）：
${summarizeRulesForPrompt(rules)}

输出严格 JSON：
{
  "contract_title": "合同标题",
  "draft": "完整合同草稿，使用中文条款编号",
  "applied_rules": ["使用到的规则名称"],
  "missing_warnings": ["仍建议人工确认的信息"]
}

要求：
1. 不要编造公司证照、银行账户、身份证号等敏感信息，缺失处用【待补充】。
2. 条款要覆盖主体、标的、费用、付款、交付/服务、验收、保密、知识产权、违约责任、解除终止、争议解决。
3. 优先遵守规则库中的企业自定规则、匹配行业预设和高优先级通用规则，并将规则要求吸收到合同条款中。
4. 如提供模板正文，必须优先使用该模板的章节结构，将已补充信息替换到对应占位符；缺失字段保留【待补充】。
5. 不要机械复制用户原话。必须先理解用户描述、已补充信息和上传附件片段，再提炼为合同可用的正式条款；口语、碎片信息、附件清单要整合成逻辑完整、权责清楚、可执行的合同正文。
6. 如果用户要求填写交付清单、供货清单、报价单、验收标准、附件内容等，优先从“上传附件/知识库可参考片段”中识别产品/服务、数量、型号、规格、交付物、验收节点等信息，并写入对应条款或附件引用；没有依据时保留【待补充】。
7. 正文格式必须符合常用中文合同模板：标题单独一行；正文使用“第一章/第一条”或“一、/1.”等清晰层级；每条尽量完整成句；末尾保留甲方、乙方签署栏、签署日期和盖章位置；附件、报价单、验收单等应在正文中引用清楚。
8. draft 字段只能输出可直接放入合同正文的条款，不得包含思考过程、推理过程、分析过程、生成说明、规则命中说明或模型自述。
9. applied_rules 单独列出规则名称，不要在 draft 正文中写“规则预检/规则命中/适用规则清单”。`;
}

function normalizeFieldAnswer(field = {}, value = "") {
  const keyText = `${field.key || ""} ${field.label || ""}`.toLowerCase();
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (/amount|price|fee|cost|金额|价款|费用|租金|薪酬|工资|报酬/.test(keyText)) {
    return text.replace(/[，,]\s*/g, ",");
  }
  return text;
}

function normalizeModelAnswers(modelData = {}, template = {}) {
  const raw = modelData.answers && typeof modelData.answers === "object" ? modelData.answers : modelData;
  const allowed = new Set((template.requiredFields || []).map((field) => field.key));
  return Object.fromEntries(
    Object.entries(raw || {})
      .filter(([key, value]) => allowed.has(key) && String(value || "").trim() && !/待补充|未知|不确定|无法判断/.test(String(value || "")))
      .map(([key, value]) => [key, String(value || "").replace(/\s+/g, " ").trim()])
  );
}

function buildAnswerExtractionPrompt(template, description, answers, rules, snippets, profile = {}) {
  const fieldLines = (template.requiredFields || [])
    .map((field) => `- ${field.key}｜${field.label}｜${field.question || ""}`)
    .join("\n");
  return `你是合同生成信息抽取与条款提炼助手。请从用户连续对话、已知答案、规则库和上传附件片段中识别可以填入模板字段的信息，并把口语表达改写成合同正文可直接使用的正式语句。

模板：${template.name}
合同识别：${profile.contractType || "通用合同"} / ${profile.industry || "通用"}
字段清单：
${fieldLines}

用户对话/描述：
${String(description || "").slice(0, 120000)}

已知答案：
${JSON.stringify(answers, null, 2)}

上传附件/知识库片段：
${summarizeKnowledgeSnippetsForPrompt(snippets || [], 3600)}

规则摘要：
${summarizeRulesForPrompt(rules, 24)}

输出严格 JSON：
{
  "answers": {
    "字段key": "提炼后的合同正文语句或字段值"
  },
  "evidence": {
    "字段key": "使用了哪个用户描述或附件片段"
  }
}

要求：
1. 只输出字段清单中存在的 key。
2. 不要照抄用户原话；要总结成合同可用的正式表述，语句要完整、清晰、可执行。
3. 交付清单、供货清单、附件、报价单、验收标准等优先从上传附件/知识库片段中提取，写成“交付物/数量/规格/交付时间/验收方式”一类可直接入合同的表达。
4. 没有明确依据的字段不要填，不要编造公司证照、金额、银行账户、身份证号等敏感信息。
5. 付款、交付、验收、违约等字段必须符合交易闭环和规则库要求。`;
}

function answerHasAmount(value = "") {
  return /(?:人民币|rmb|￥|¥)?\s*\d[\d,]*(?:\.\d+)?\s*(?:元|万元|亿元|%|％)?|[一二三四五六七八九十百千万亿壹贰叁肆伍陆柒捌玖拾佰仟]+(?:元|万元|亿元)/i.test(String(value || ""));
}

function buildAnswerSuggestions(field = {}, answers = {}, template = {}) {
  const keyText = `${field.key || ""} ${field.label || ""}`;
  if (/付款|支付|payment|结算/.test(keyText)) {
    return [
      { source: "大模型推荐", text: "按30%预付款、60%交付验收款、10%尾款分期支付，并以发票和验收通过作为付款条件。" },
      { source: "知识库推荐", text: "验收合格且收到合法有效发票及完整付款资料后十个工作日内支付对应款项。" },
      { source: "历史常用", text: "按双方确认的付款节点和无争议金额结算。" },
    ];
  }
  if (/金额|价款|费用|amount|price|fee/.test(keyText)) {
    return [
      { source: "大模型推荐", text: "固定总价人民币【金额】元，含税/不含税以双方确认为准。" },
      { source: "知识库推荐", text: "人民币【金额】元（大写：【大写金额】），税费及发票类型按合同约定执行。" },
      { source: "历史常用", text: answers.amount || "人民币【金额】元" },
    ];
  }
  if (/交付|履行|验收|delivery|acceptance/.test(keyText)) {
    return [
      { source: "大模型推荐", text: "乙方按约完成交付，甲方在收到成果后____个工作日内完成验收并书面确认。" },
      { source: "知识库推荐", text: "验收不合格的，乙方应在____个工作日内完成整改并重新提交验收。" },
      { source: "历史常用", text: "按双方确认的交付清单、质量标准和验收流程执行。" },
    ];
  }
  return [
    { source: "大模型推荐", text: `${field.label || "该字段"}以双方书面确认内容为准。` },
    { source: "知识库推荐", text: `${field.label || "该字段"}应明确适用范围、责任边界和履行标准。` },
    { source: "历史常用", text: "暂无历史记录，建议人工补充完整信息。" },
  ];
}

function fallbackValidateGenerationAnswer(field = {}, answer = "", answers = {}, template = {}) {
  const normalizedValue = normalizeFieldAnswer(field, answer);
  const keyText = `${field.key || ""} ${field.label || ""}`.toLowerCase();
  if (!normalizedValue) {
    return {
      accepted: false,
      normalizedValue: "",
      reason: "回答为空，不能写入合同字段。",
      warnings: [],
      suggestions: buildAnswerSuggestions(field, answers, template),
    };
  }
  if (/amount|price|fee|cost|金额|价款|费用|租金|薪酬|工资|报酬/.test(keyText) && !answerHasAmount(normalizedValue)) {
    return {
      accepted: false,
      normalizedValue,
      reason: "该字段需要明确金额、币种或计价方式，目前回答缺少可识别的金额信息。",
      warnings: [],
      suggestions: buildAnswerSuggestions(field, answers, template),
    };
  }
  if (/payment|付款|支付|结算/.test(keyText) && /一次性|全款|先款/.test(normalizedValue) && /交付|验收|服务|开发|供货/.test(JSON.stringify({ answers, template }))) {
    return {
      accepted: true,
      normalizedValue,
      reason: "",
      warnings: ["付款安排已写入，但存在先付款或一次性付款特征，生成合同时会优先补入验收、发票和异议处理条件。"],
      suggestions: [],
    };
  }
  if (/party|甲方|乙方|采购方|供应方|用人单位|劳动者/.test(keyText) && /^[甲乙双方]+$/.test(normalizedValue)) {
    return {
      accepted: false,
      normalizedValue,
      reason: "主体字段需要填写具体公司、组织或个人名称，不能只写甲方/乙方。",
      warnings: [],
      suggestions: buildAnswerSuggestions(field, answers, template),
    };
  }
  return {
    accepted: true,
    normalizedValue,
    reason: "",
    warnings: [],
    suggestions: [],
  };
}

function buildAnswerValidationPrompt(template, field, answer, answers, description, rules, profile = {}, snippets = []) {
  return `你是合同生成字段校验与提炼助手。请判断用户对当前追问字段的回答能否写入合同模板，并结合上下文、规则库和上传附件片段，改写为可直接填入合同正文的正式内容。

合同模板：${template.name}
合同识别：${profile.contractType || "通用合同"} / ${profile.industry || "通用"}
字段：${field.label || field.key}（key=${field.key || ""}）
字段追问：${field.question || ""}

用户整体描述：
${String(description || "").slice(0, 120000)}

已知答案：
${JSON.stringify(answers, null, 2)}

本次回答：
${answer}

上传附件/知识库相关片段：
${summarizeKnowledgeSnippetsForPrompt(snippets, 2800)}

规则摘要：
${summarizeRulesForPrompt(rules, 24)}

输出严格 JSON：
{
  "accepted": true,
  "normalized_value": "可直接写入该字段和合同正文的内容",
  "reason": "拒绝原因或空字符串",
  "warnings": ["可接受但生成时需要注意的逻辑风险"],
  "suggestions": [
    { "source": "历史常用/知识库推荐/大模型推荐", "text": "可替代回答" }
  ]
}

要求：
1. 只校验当前字段，不输出合同全文。
2. normalized_value 只能是可填入正文的内容，不得出现“建议、请、可考虑、需要确认、修改建议”等提示语。
3. 不要机械复制用户原话；把用户口语和碎片化信息总结成合同正文需要的语句、清单或条款片段。
4. 若当前字段是交付清单、供货清单、验收标准、附件内容、报价清单等，优先从上传附件/知识库相关片段中抽取可用信息；若附件片段不足，再结合用户回答提炼。
5. 若回答答非所问、与已知答案冲突、金额/期限/主体明显不完整，accepted=false，并给出三条可选补充。
6. 若可以自动修正口语化表述、补齐币种单位或去掉无关内容，accepted=true，并把修正后内容放入 normalized_value。
7. 付款、交付、验收、违约等字段要检查交易闭环；发现一次性付款、先款后货、缺少验收或发票条件时可以 accepted=true，但 warnings 要指出生成时必须补充约束。`;
}

function normalizeAnswerValidationResult(modelData = {}, fallback = {}) {
  const accepted = modelData.accepted !== false && String(modelData.accepted).toLowerCase() !== "false";
  const normalizedValue = normalizeFieldAnswer({}, modelData.normalized_value || modelData.normalizedValue || fallback.normalizedValue || "");
  return {
    accepted: accepted && Boolean(normalizedValue),
    normalizedValue,
    reason: String(modelData.reason || fallback.reason || "").trim(),
    warnings: Array.isArray(modelData.warnings) ? modelData.warnings.map((item) => String(item || "").trim()).filter(Boolean).slice(0, 4) : fallback.warnings || [],
    suggestions: Array.isArray(modelData.suggestions) && modelData.suggestions.length
      ? modelData.suggestions.map((item) => typeof item === "string" ? { source: "大模型推荐", text: item } : { source: item.source || "大模型推荐", text: item.text || item.value || "" }).filter((item) => item.text).slice(0, 3)
      : fallback.suggestions || [],
  };
}

async function handleValidateContractAnswer(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const answer = String(payload.answer || "").trim();
  const availableTemplates = templatesForPayload(payload);
  const selected = availableTemplates.find((item) => item.id === payload.templateId) || chooseContractTemplate(payload.description || answer, payload.templateId, availableTemplates).template;
  const fieldInput = payload.field || {};
  const field = (selected.requiredFields || []).find((item) => item.key === fieldInput.key) || fieldInput;
  if (!field?.key) return sendJson(res, 400, { error: "缺少需要校验的字段" });
  const answers = normalizeAnswers(payload.answers || {});
  const fallback = fallbackValidateGenerationAnswer(field, answer, answers, selected);
  if (!fallback.accepted) return sendJson(res, 200, { ...fallback, usedAI: false, model: activeModelInfo().model, modelProvider: activeModelInfo().provider });
  const store = loadStore();
  const profileText = [payload.description, answer, selected.name, JSON.stringify(answers)].join("\n");
  const profile = inferContractProfile(profileText, selected);
  const rules = selectRulesForContract(store, profile);
  const knowledgeSnippets = selectKnowledgeSnippetsForGeneration(store, profileText, { field, limit: 8 });
  try {
    const modelResult = await callJsonModel(buildAnswerValidationPrompt(selected, field, answer, answers, payload.description || "", rules, profile, knowledgeSnippets), {
      repair: true,
      label: "合同生成字段校验结果",
    });
    const normalized = normalizeAnswerValidationResult(modelResult.data, fallback);
    return sendJson(res, 200, {
      ...normalized,
      knowledgeSources: knowledgeSnippets.slice(0, 3).map((item) => item.source),
      usedAI: true,
      model: modelResult.model || activeModelInfo().model,
      modelProvider: modelResult.provider || activeModelInfo().provider,
    });
  } catch (error) {
    return sendJson(res, 200, {
      ...fallback,
      usedAI: false,
      fallbackReason: error.message,
      knowledgeSources: knowledgeSnippets.slice(0, 3).map((item) => item.source),
      model: activeModelInfo().model,
      modelProvider: activeModelInfo().provider,
    });
  }
}

function escapeRegExp(value = "") {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function termRegex(term = "") {
  return new RegExp(escapeRegExp(term), "i");
}

function textHasAnyTerm(text = "", terms = []) {
  return (terms || []).some((term) => term && termRegex(term).test(text));
}

function missingRequiredGroups(text = "", groups = []) {
  return (groups || []).filter((group) => !textHasAnyTerm(text, group));
}

function matchedTerms(text = "", terms = []) {
  return (terms || []).filter((term) => term && termRegex(term).test(text));
}

function ruleLocationQuote(contractText = "", rule = {}, problemTerms = []) {
  const terms = [
    ...problemTerms,
    ...(rule.check?.triggerAny || []),
    ...(rule.check?.forbidAny || []),
    rule.ruleName,
  ].filter(Boolean);
  for (const term of terms) {
    const quote = findSentence(contractText, termRegex(term));
    if (quote) return quote;
  }
  return contractText.slice(0, 180);
}

function reviewIssueFromPresetRule(contractText = "", rule = {}) {
  const check = rule.check || null;
  if (!check || !Array.isArray(rule.useScope) || !rule.useScope.includes("审查")) return null;
  if (rule.reviewStatus && ["inactive", "rejected"].includes(rule.reviewStatus)) return null;

  const triggers = check.triggerAny || [];
  if (triggers.length && !textHasAnyTerm(contractText, triggers)) return null;

  const forbidden = matchedTerms(contractText, check.forbidAny || []);
  const missingGroups = missingRequiredGroups(contractText, check.requiredAnyGroups || []);
  if (!forbidden.length && !missingGroups.length) return null;

  const missingText = missingGroups
    .map((group) => group.slice(0, 4).join("/"))
    .join("；");
  const problemParts = [];
  if (check.problem) problemParts.push(check.problem);
  if (forbidden.length) problemParts.push(`命中高风险表述：${forbidden.slice(0, 4).join("、")}。`);
  if (missingText) problemParts.push(`缺少或未清晰识别：${missingText}。`);

  return {
    risk_level: rule.riskLevel || (forbidden.length ? "高" : "中"),
    category: "规则库",
    detail_category: check.detailCategory || rule.scenario || "规则库",
    title: rule.ruleName,
    problem: problemParts.join(""),
    suggestion: rule.action || "按平台预设规则补充或修改合同条款。",
    replacement_text: check.replacementText || rule.action || "",
    source_rule: rule.ruleName,
    source_rule_basis: rule.ruleBasis || "通用法规",
    legal_basis: rule.sourceDocName ? `${rule.sourceDocName}：${rule.sourceQuote || rule.ruleBasis || "合同审查规则"}` : (rule.sourceQuote || rule.ruleBasis || ""),
    source_quote: rule.sourceQuote || "",
    source_url: rule.sourceUrl || "",
    location_hint: check.locationHint || (forbidden.length ? "replace" : "supplement"),
    quote: ruleLocationQuote(contractText, rule, forbidden),
  };
}

function buildRuleReviewIssues(contractText = "", rules = []) {
  return (rules || [])
    .slice()
    .sort((a, b) => Number(b.priority || 0) - Number(a.priority || 0))
    .map((rule) => reviewIssueFromPresetRule(contractText, rule))
    .filter(Boolean)
    .slice(0, 12);
}

function hidePresetRuleIssue(issue = {}, rules = []) {
  return {
    ...issue,
    source_rule: issue.source_rule || issue.sourceRule || "",
    source_rule_basis: issue.source_rule_basis || issue.sourceRuleBasis || issue.rule_basis || issue.ruleBasis || "",
    legal_basis: issue.legal_basis || issue.legalBasis || "",
    source_quote: issue.source_quote || issue.sourceQuote || "",
    source_url: issue.source_url || issue.sourceUrl || "",
  };
}

function applyReviewVisibility(issues = [], rules = [], contractText = "") {
  const visible = (Array.isArray(issues) ? issues : []).map((issue) => hidePresetRuleIssue(issue, rules));
  return mergeReviewIssues(visible, [], contractText);
}

function fallbackReview(contractText, rules) {
  const issues = [];
  const enterpriseRules = enterpriseRulesOnly(rules);
  const checks = [
    { re: /甲方|乙方/, title: "合同主体需核验", risk: "中", suggestion: "确认双方名称、统一社会信用代码、地址、联系人和签约权限完整准确。" },
    { re: /付款|支付|价款|费用/, title: "付款条款需明确", risk: "中", suggestion: "明确金额、税费、付款节点、付款条件、发票和逾期责任。" },
    { re: /验收|交付|服务/, title: "交付验收需闭环", risk: "中", suggestion: "明确交付标准、验收期限、异议处理和整改机制。" },
    { re: /违约|赔偿/, title: "违约责任需可执行", risk: "中", suggestion: "确认违约金、损失赔偿、解除权和责任上限是否合理。" },
    { re: /保密|秘密|数据|个人信息/, title: "保密和数据条款需核查", risk: "中", suggestion: "确认保密范围、期限、例外、个人信息处理和泄密责任。" },
  ];
  for (const check of checks) {
    if (!check.re.test(contractText)) {
      issues.push({
        risk_level: check.risk,
        category: "常见问题",
        title: `缺少${check.title.replace("需", "")}`,
        problem: `未明显识别到${check.title.replace("需", "")}相关约定。`,
        suggestion: check.suggestion,
        replacement_text: check.suggestion,
        source_rule: "常见问题清单",
        location_hint: "supplement",
        quote: "",
      });
    }
  }
  for (const rule of enterpriseRules.slice(0, 12)) {
    const needle = [rule.triggerCondition, rule.ruleName].join(" ").slice(0, 12);
    if (needle && !contractText.includes(needle.slice(0, 4))) {
      issues.push({
        risk_level: rule.riskLevel || "中",
        category: rule.ruleBasis || "规则库",
        title: rule.ruleName,
        problem: "合同文本未明显覆盖该规则要求，建议人工核对。",
        suggestion: rule.action || rule.triggerCondition || "按规则库要求补充或修改条款。",
        replacement_text: rule.action || rule.triggerCondition || "按规则库要求补充或修改条款。",
        source_rule: rule.ruleName,
        location_hint: "supplement",
        quote: rule.sourceQuote || "",
      });
    }
  }
  const allRuleIssues = buildRuleReviewIssues(contractText, rules);
  const visibleRuleIssues = buildRuleReviewIssues(contractText, enterpriseRules);
  const mergedIssues = applyReviewVisibility(
    mergeReviewIssues(issues, [...visibleRuleIssues, ...buildLogicReviewIssues(contractText)], contractText),
    rules,
    contractText
  ).slice(0, 16);
  return {
    overall_risk: mergedIssues.some((item) => item.risk_level === "高") ? "高" : mergedIssues.length ? "中" : "低",
    summary: `AI 暂不可用，已启用本地临时审核，基于平台预设法规规则库发现 ${mergedIssues.length} 项待确认问题。`,
    reviewReport: buildReviewReport(mergedIssues, contractText),
    issues: mergedIssues,
    temporaryReview: true,
    presetRuleIssueCount: allRuleIssues.length,
  };
}

function compactText(value = "") {
  return String(value || "").replace(/\s+/g, "").replace(/[，。；;,.、：:]/g, "");
}

function serverSentenceBounds(text, index) {
  let start = 0;
  for (const mark of ["\n", "。", "；", ";", "！", "!", "？", "?"]) {
    const found = text.lastIndexOf(mark, Math.max(0, index - 1));
    if (found >= start) start = found + mark.length;
  }
  let end = text.length;
  for (const mark of ["\n", "。", "；", ";", "！", "!", "？", "?"]) {
    const found = text.indexOf(mark, index);
    if (found >= 0) end = Math.min(end, found + mark.length);
  }
  while (start < end && /\s/.test(text[start])) start += 1;
  while (end > start && /\s/.test(text[end - 1])) end -= 1;
  return text.slice(start, end).trim();
}

function findSentence(text = "", pattern) {
  const match = pattern.exec(text);
  return match ? serverSentenceBounds(text, match.index).slice(0, 260) : "";
}

function findPaymentSentence(text = "", amountList = []) {
  const bestAmount = amountList[0];
  if (bestAmount) {
    const amountSentence = serverSentenceBounds(text, bestAmount.index).slice(0, 260);
    if (/付款|支付|价款|费用|金额|合计|款项|结算|发票|元|万元|亿元/.test(amountSentence)) return amountSentence;
  }
  return findSentence(text, /付款|支付|价款|费用|合同金额|合同价款|款项|结算|发票/);
}

function captureAfterLabel(text = "", label) {
  const match = new RegExp(`${label}\\s*[：:]\\s*([^\\n；;。]{2,80})`).exec(text);
  return match ? match[1].trim() : "";
}

function reviewIssueText(issue = {}) {
  return [issue.title, issue.problem, issue.suggestion, issue.category, issue.detail_category, issue.detailCategory, issue.source_rule, issue.sourceRule, issue.quote]
    .filter(Boolean)
    .join(" ");
}

function normalizeDetailCategory(detail = "", issue = {}) {
  const text = [detail, issue.title, issue.problem, issue.suggestion, issue.category, issue.source_rule, issue.sourceRule, issue.quote]
    .filter(Boolean)
    .join(" ");
  if (/逻辑|闭环|交易|衔接|顺序|先款|先付款|付款条件|权利义务|单边|失衡|矛盾|冲突|验收.*交付|交付.*验收/.test(text)) return "逻辑错误";
  if (/风险语句|风险表述|绝对|最终解释|免责|不承担任何责任|责任无上限|无限责任|单方|可随时|永久|不可撤销|全部损失|所有损失|显失公平/.test(text)) return "风险语句";
  if (/格式|编号|标题|排版|附件|签署|签章|盖章|日期|落款|页码|空白|错别字|错字|语病|称谓|用语/.test(text)) return "格式问题";
  return String(detail || "其他").trim() || "其他";
}

const AMOUNT_PATTERN = /(?:人民币|RMB|￥|¥)?\s*\d[\d,]*(?:\.\d+)?\s*(?:元|万元|亿元)?|[零一二三四五六七八九十百千万亿壹贰叁肆伍陆柒捌玖拾佰仟萬億两]+(?:元|万元|亿元)/g;

function chineseAmountSectionValue(section = "") {
  const numberMap = {
    零: 0,
    〇: 0,
    一: 1,
    壹: 1,
    二: 2,
    贰: 2,
    两: 2,
    三: 3,
    叁: 3,
    四: 4,
    肆: 4,
    五: 5,
    伍: 5,
    六: 6,
    陆: 6,
    七: 7,
    柒: 7,
    八: 8,
    捌: 8,
    九: 9,
    玖: 9,
  };
  const unitMap = { 十: 10, 拾: 10, 百: 100, 佰: 100, 千: 1000, 仟: 1000 };
  let total = 0;
  let current = 0;
  for (const char of String(section || "")) {
    if (Object.prototype.hasOwnProperty.call(numberMap, char)) {
      current = numberMap[char];
    } else if (Object.prototype.hasOwnProperty.call(unitMap, char)) {
      total += (current || 1) * unitMap[char];
      current = 0;
    }
  }
  return total + current;
}

function parseChineseAmount(raw = "") {
  let text = String(raw || "").replace(/[人民币圆元整正￥¥,，\s]/g, "");
  if (!/[零〇一二两三四五六七八九十百千万亿壹贰叁肆伍陆柒捌玖拾佰仟萬億]/.test(text)) return null;
  text = text.replace(/億/g, "亿").replace(/萬/g, "万");
  let total = 0;
  if (text.includes("亿")) {
    const parts = text.split("亿");
    total += chineseAmountSectionValue(parts.shift()) * 100000000;
    text = parts.join("亿");
  }
  if (text.includes("万")) {
    const parts = text.split("万");
    total += chineseAmountSectionValue(parts.shift()) * 10000;
    text = parts.join("万");
  }
  total += chineseAmountSectionValue(text);
  return total > 0 ? total : null;
}

function parseAmountValue(raw = "") {
  const text = String(raw || "").replace(/\s+/g, "");
  const digit = text.match(/\d[\d,]*(?:\.\d+)?/);
  if (digit) {
    let value = Number(digit[0].replace(/,/g, ""));
    if (!Number.isFinite(value)) return null;
    if (/亿元/.test(text)) value *= 100000000;
    else if (/万元/.test(text)) value *= 10000;
    return value;
  }
  return parseChineseAmount(text);
}

function amountMatches(value = "") {
  const text = String(value || "");
  return [...text.matchAll(AMOUNT_PATTERN)].map((match) => {
    const raw = String(match[0] || "").trim();
    const index = match.index || 0;
    const before = text.slice(Math.max(0, index - 20), index);
    const after = text.slice(index + raw.length, index + raw.length + 20);
    const nearby = `${before}${raw}${after}`;
    const compactNearby = nearby.replace(/\s+/g, "");
    const nextChar = String(after || "").trim().charAt(0);
    const hasAmountUnit = /(人民币|RMB|￥|¥|元|万元|亿元)/.test(raw);
    const hasCurrencyPrefix = /(人民币|RMB|￥|¥)\s*$/.test(before);
    const hasStrongSignal = /(合同总价|合同金额|合同价款|费用合计|报名费用合计|费用总计|价款合计|合计|总价|总费用|总金额|应付金额|价款|金额)/.test(compactNearby);
    const hasWeakSignal = /(费用|付款|支付|款项|报价|结算|收款|应付)/.test(compactNearby);
    const isCountLike = /^(人|名|个|项|份|件|套|门|次|天|日|周|月|年|小时|分钟|%|％)/.test(nextChar);
    const isAccountLike = /(账号|账户|银行|开户|卡号|电话|手机|联系方式|身份证|证件|统一社会信用代码|编号|日期|工种|证书|数量|人数|报名数量)$/.test(before.replace(/\s+/g, ""));
    const digitLength = (raw.match(/\d/g) || []).length;
    let score = 0;
    if (hasAmountUnit) score += 40;
    if (hasCurrencyPrefix) score += 20;
    if (hasStrongSignal) score += 28;
    if (hasWeakSignal) score += 10;
    if (/合计|总价|总金额|合同价款|合同金额|费用合计/.test(compactNearby)) score += 20;
    if (isCountLike) score -= 60;
    if (isAccountLike) score -= 80;
    if (digitLength >= 9 && !hasAmountUnit && !hasCurrencyPrefix) score -= 90;
    if (/^(?:19|20)\d{2}$/.test(raw) && /^年/.test(after.trim())) score -= 90;
    return {
      raw,
      index,
      value: parseAmountValue(raw),
      score,
    };
  }).filter((item) => {
    return Number.isFinite(item.value) && item.value > 0 && item.score > 0;
  }).sort((a, b) => b.score - a.score || b.value - a.value || a.index - b.index);
}

function firstAmount(value = "") {
  return amountMatches(value)[0]?.raw || "";
}

function formatMoney(value) {
  if (!Number.isFinite(value) || value <= 0) return "人民币【待补充】元";
  const rounded = Math.round(value * 100) / 100;
  const formatted = rounded.toLocaleString("zh-CN", {
    minimumFractionDigits: Number.isInteger(rounded) ? 0 : 2,
    maximumFractionDigits: 2,
  });
  return `人民币${formatted}元`;
}

function splitPaymentAmount(amountValue, ratio) {
  if (!Number.isFinite(amountValue) || amountValue <= 0) return "";
  return `，即${formatMoney(amountValue * ratio)}`;
}

function extractReviewContext(contractText = "") {
  const amountList = amountMatches(contractText);
  const amount = amountList[0]?.raw || "";
  const amountValue = amountList.find((item) => Number.isFinite(item.value) && item.value > 0)?.value || null;
  const partyA = captureAfterLabel(contractText, "甲方");
  const partyB = captureAfterLabel(contractText, "乙方");
  const amountSentence = amountList[0]
    ? serverSentenceBounds(contractText, amountList[0].index).slice(0, 260)
    : findSentence(contractText, /人民币|RMB|￥|¥|金额|价款|费用|报价|总价|合同金额|合同价款|\d[\d,]*(?:\.\d+)?\s*(?:元|万元|亿元)?/);
  const paymentSentence = findPaymentSentence(contractText, amountList);
  const deliverySentence = findSentence(contractText, /交付|交货|服务|履行|完成|提交|提供/);
  const acceptanceSentence = findSentence(contractText, /验收|确认|审核|异议|整改/);
  const liabilitySentence = findSentence(contractText, /违约|赔偿|责任|损失|解除|终止/);
  const riskSentence = findSentence(contractText, /最终解释权|不承担任何责任|全部损失|所有损失|无限责任|无条件|永久|不可撤销|随时单方|单方变更|单方解除|免责/);
  const subjectSentence = findSentence(contractText, /标的|项目|服务内容|服务范围|产品|货物|课程|委托事项/);
  return {
    amount,
    amountValue,
    amountList,
    amountSentence,
    partyA,
    partyB,
    paymentSentence,
    deliverySentence,
    acceptanceSentence,
    liabilitySentence,
    riskSentence,
    subjectSentence,
  };
}

function isWeakReplacementText(replacement = "", suggestion = "") {
  const text = String(replacement || "").trim();
  if (!text) return true;
  if (compactText(text) && compactText(text) === compactText(suggestion)) return true;
  if (text.length < 18) return true;
  if (/^(建议|需要|需|请|可|可以|宜|建议双方|双方应明确|明确|确认|完善|补充|修改|优化|核对|注意|可考虑)/.test(text)) return true;
  if (/^应当(?:明确|补充|约定|载明|删除|调整|修改|完善)/.test(text)) return true;
  if (/建议|请.*确认|人工复核|进一步明确|进行约定|补充约定|完善条款|修改为更合理|可考虑|需人工/.test(text) && text.length < 90) return true;
  return false;
}

function cleanReplacementText(replacement = "") {
  return String(replacement || "")
    .trim()
    .replace(/^(?:修改建议|处理意见|建议替换为|建议补充为|建议修改为|建议|请|需要|需|可以|可考虑|宜)[:：，,、\s]*/g, "")
    .replace(/^双方应明确[:：，,、\s]*/g, "")
    .replace(/^应当(?:明确|补充|约定|载明|删除|调整|修改|完善)[:：，,、\s]*/g, "")
    .trim();
}

function hasReviewWording(text = "") {
  return /^(建议|请|需要|需|可考虑|宜|明确|完善|补充|修改|优化|核对|注意)|建议|请.*确认|人工复核|需人工|修改建议|处理意见/.test(String(text || "").trim());
}

function contractContainsQuote(contractText = "", quote = "") {
  const raw = String(quote || "").trim();
  if (!raw) return false;
  const text = String(contractText || "");
  return text.includes(raw) || compactText(text).includes(compactText(raw));
}

function quoteLooksSameIssueTopic(issue = {}, quote = "") {
  const text = reviewIssueText(issue);
  const source = String(quote || "");
  return (
    (/付款|支付|价款|金额|费用|发票|结算|款项|税费/.test(text) && /付款|支付|价款|金额|费用|发票|结算|款项|税费|元|万元|\d/.test(source)) ||
    (/交付|交货|履行|服务|验收|整改|异议|质量/.test(text) && /交付|交货|履行|服务|验收|整改|异议|质量/.test(source)) ||
    (/主体|甲方|乙方|签约|授权|证照|名称|统一社会信用代码|联系人|地址|住所/.test(text) && /甲方|乙方|签约|授权|证照|名称|统一社会信用代码|联系人|地址|住所/.test(source)) ||
    (/违约|赔偿|责任|损失|解除|终止|逾期/.test(text) && /违约|赔偿|责任|损失|解除|终止|逾期/.test(source)) ||
    (/格式|编号|标题|附件|签署|签章|盖章|日期|称谓|错别字|用词|语病/.test(text) && /第[一二三四五六七八九十\d]+|附件|签署|签章|盖章|日期|甲方|乙方/.test(source))
  );
}

function inferReviewLocationHint(issue = {}, context = {}, replacement = "") {
  const rawHint = String(issue.location_hint || issue.locationHint || "").trim().toLowerCase();
  const text = reviewIssueText(issue);
  const quote = String(issue.quote || issue.original_text || issue.originalText || "").trim();
  const hasLocatedSource = contractContainsQuote(context.contractText || "", quote);
  if (rawHint === "delete" || rawHint === "remove") return "delete";
  if (/删除|删去|移除|去除|不应保留|多余|重复|连续空行|不必要空行/.test(text) && (hasLocatedSource || /空行/.test(text))) {
    return "delete";
  }
  if (/最终解释权归|解释权归.*所有/.test(text) && hasLocatedSource) return "delete";
  if (rawHint === "supplement" && hasLocatedSource) {
    if (replacement || /修改|替换|改为|不合理|错误|不准确|冲突|不一致|过宽|过高|过低|绝对|单方|风险|语病|错别字|用词/.test(text)) {
      return "replace";
    }
    if (/未明确|未完整|不完整|缺少|缺失|未约定|未载明|缺乏/.test(text) && quoteLooksSameIssueTopic(issue, quote)) {
      return "replace";
    }
  }
  if (rawHint === "replace" || rawHint === "supplement") return rawHint;
  if (hasLocatedSource && replacement) return "replace";
  return "supplement";
}

function isPaymentIssue(issue = {}) {
  return /付款|支付|价款|金额|费用|发票|结算|账期|款项|税费|合同价款|合同金额/.test(reviewIssueText(issue));
}

function amountValueEqual(a, b) {
  if (!Number.isFinite(a) || !Number.isFinite(b)) return false;
  return Math.abs(a - b) <= Math.max(1, Math.abs(b) * 0.005);
}

function paymentReplacementMissesContractAmount(issue = {}, replacement = "", context = {}) {
  if (!isPaymentIssue(issue) || !Number.isFinite(context.amountValue) || context.amountValue <= 0) return false;
  const values = amountMatches(replacement).map((item) => item.value).filter((value) => Number.isFinite(value) && value > 0);
  if (!values.length) return true;
  return !values.some((value) => amountValueEqual(value, context.amountValue));
}

function buildPaymentReplacementClause(issue = {}, context = {}) {
  const issueAmount = firstAmount([issue.quote, issue.problem, issue.suggestion, issue.replacement_text].join(" "));
  const rawAmount = issueAmount || context.amount;
  const amountValue = parseAmountValue(rawAmount) || context.amountValue;
  const partyA = context.partyA || "甲方";
  const partyB = context.partyB || "乙方";
  const totalText = Number.isFinite(amountValue) && amountValue > 0 ? formatMoney(amountValue) : rawAmount || "人民币【待补充】元";
  const deliveryText = context.deliverySentence ? "对应交付成果或服务" : "本合同项下工作";
  return `付款安排：本合同价款为${totalText}，该价款是否含税、适用税率及发票类型以本合同约定及双方确认的结算文件为准。${partyA}按以下节点向${partyB}支付合同价款：（1）预付款：本合同生效且${partyB}提交合法有效发票及完整付款资料后____个工作日内，${partyA}支付合同价款的30%${splitPaymentAmount(amountValue, 0.3)}；（2）交付验收款：${partyB}完成${deliveryText}并经${partyA}验收合格后____个工作日内，${partyA}支付合同价款的60%${splitPaymentAmount(amountValue, 0.6)}；（3）尾款：全部交付成果验收合格、双方完成结算且无未解决异议后____个工作日内，${partyA}支付合同价款的10%${splitPaymentAmount(amountValue, 0.1)}。${partyB}应按约开具合法有效的增值税发票及付款所需资料；${partyA}对金额、发票或付款资料有异议的，应在收到相关资料后五个工作日内书面提出，双方核对确认后按无争议金额付款。`;
}

function clauseReplacementForIssue(issue = {}, context = {}) {
  const text = reviewIssueText(issue);
  const amount = firstAmount([issue.quote, issue.problem, issue.suggestion].join(" ")) || context.amount;
  const partyA = context.partyA || "甲方";
  const partyB = context.partyB || "乙方";
  const subject = context.subjectSentence ? `围绕“${context.subjectSentence.replace(/\s+/g, " ").slice(0, 80)}”` : "按照本合同约定";
  if (/付款|支付|价款|金额|费用|发票|结算|账期|款项|税费/.test(text)) {
    return buildPaymentReplacementClause({ ...issue, replacement_text: amount }, context);
  }
  if (/主体|甲方|乙方|签约|授权|证照|名称|统一社会信用代码/.test(text)) {
    return `合同主体：甲方${partyA && partyA !== "甲方" ? `（${partyA}）` : ""}应补充统一社会信用代码、住所、联系人及联系方式；乙方${partyB && partyB !== "乙方" ? `（${partyB}）` : ""}应补充统一社会信用代码、住所、联系人及联系方式。双方确认，签署本合同的代表已取得合法有效授权，其签署行为对各自一方具有约束力。`;
  }
  if (/交付|交货|履行|服务|验收|整改|异议|质量/.test(text)) {
    return `交付与验收：${partyB}应${subject}完成交付，交付成果应符合本合同及附件约定的范围、标准和期限。${partyA}应在收到交付成果后____个工作日内完成验收；发现不符合约定的，应一次性书面提出异议及整改要求，${partyB}应在____个工作日内完成整改并重新提交验收。${partyA}逾期未提出书面异议且已实际使用交付成果的，视为验收通过，但不免除${partyB}依法及依约应承担的质量保证责任。`;
  }
  if (/违约|赔偿|责任|损失|解除|终止|逾期|无限责任|全部损失|免责/.test(text)) {
    return "违约责任：任何一方违反本合同约定，给守约方造成损失的，应承担继续履行、采取补救措施或赔偿损失等违约责任。违约金不足以弥补守约方实际损失的，违约方还应赔偿差额；但除故意、重大过失、侵犯知识产权、违反保密义务或法律另有规定外，违约方的赔偿责任以本合同已收或应收价款总额为上限。";
  }
  if (/保密|秘密|隐私|个人信息|数据|泄露/.test(text)) {
    return "保密与数据保护：双方应对在合同订立、履行过程中知悉的商业秘密、技术资料、客户信息、个人信息及其他非公开信息承担保密义务。未经信息披露方书面同意，任何一方不得向第三方披露、转让或用于本合同目的之外的用途。保密义务自相关信息披露之日起生效，并在本合同终止后持续____年；法律法规另有更高要求的，从其规定。";
  }
  if (/知识产权|著作权|专利|商标|源码|成果|归属/.test(text)) {
    return "知识产权与成果归属：双方各自在合同签署前已拥有的知识产权仍归原权利人所有。乙方基于本合同为甲方定制开发或交付的成果，在甲方按约支付相应费用后，其可交付成果的使用权归甲方享有；涉及源代码、著作权、专利申请权或其他权利转移的，应以双方另行书面约定为准。任何一方不得侵犯第三方合法知识产权。";
  }
  if (/解除|终止|续约|退出|暂停|单方/.test(text)) {
    return "解除与终止：任何一方严重违反本合同约定，经守约方书面催告后____个工作日内仍未改正的，守约方有权解除本合同，并要求违约方承担相应违约责任。合同解除或终止不影响双方在解除或终止前已产生的付款、结算、保密、知识产权、违约责任及争议解决等条款的效力。";
  }
  if (/争议|管辖|仲裁|诉讼|法律适用/.test(text)) {
    return "争议解决：本合同的订立、效力、解释、履行及争议解决均适用中华人民共和国法律。因本合同引起或与本合同有关的任何争议，双方应先友好协商解决；协商不成的，任一方均可向合同签署地有管辖权的人民法院提起诉讼。";
  }
  if (/错别字|错字|字词|笔误|用词|语病|称谓|表述|不一致/.test(text)) {
    return "双方确认，本合同中同一主体、标的、金额、期限、附件名称及业务概念应保持前后一致；如文本存在错别字、笔误或表述不一致，应以双方真实意思表示及经双方确认的书面文件为准，并由双方签署书面补充或更正文件予以确认。";
  }
  if (/格式|编号|标题|附件|签署|签章|盖章|日期/.test(text)) {
    return "签署与附件：本合同附件、订单、报价单、验收单及双方确认的其他书面文件均为本合同不可分割的组成部分，与正文具有同等法律效力；附件与正文约定不一致的，以双方后签署或盖章确认的文件为准。本合同经双方授权代表签字并加盖公章或合同专用章后生效。";
  }
  if (/风险|不合理|不可执行|无效|违规|违法|冲突|限制|单方|绝对|永久|不可撤销/.test(text)) {
    return `风险控制条款：双方确认，本合同项下任何权利行使、义务履行、责任承担及限制性约定均应以法律法规允许、合同目的实现及公平合理为原则。任何一方不得以单方通知、格式条款或明显不合理安排排除对方主要权利、免除自身主要责任或加重对方责任；涉及${subject}的权利义务调整，应经双方协商一致并以书面形式确认。相关条款被认定无效或不可执行的，不影响本合同其他条款的效力。`;
  }
  return "补充条款：双方应根据本合同目的及实际履行需要，对本事项的适用范围、履行标准、时间节点、交付材料、验收方式、违约责任及争议处理机制作出明确约定。未尽事宜由双方另行签署书面补充协议，补充协议与本合同具有同等法律效力。";
}

function normalizeReviewIssue(issue = {}, context = {}) {
  const rawCategory = String(issue.category || "风险问题").trim() || "风险问题";
  const rawDetailCategory = String(issue.detail_category || issue.detailCategory || rawCategory).trim() || rawCategory;
  const sourceRule = String(issue.source_rule || issue.sourceRule || "").trim();
  const sourceRuleBasis = String(issue.source_rule_basis || issue.sourceRuleBasis || issue.rule_basis || issue.ruleBasis || "").trim();
  const legalBasis = String(issue.legal_basis || issue.legalBasis || "").trim();
  const sourceQuote = String(issue.source_quote || issue.sourceQuote || "").trim();
  const sourceUrl = String(issue.source_url || issue.sourceUrl || "").trim();
  const detailCategory = normalizeDetailCategory(rawDetailCategory, issue);
  const categoryText = `${rawCategory} ${detailCategory} ${sourceRule}`;
  let category = "风险问题";
  if (/规则|法规|法条|制度|企业|行业|惯例/.test(categoryText)) category = "规则库";
  if (/常见问题|格式|编号|排版|标题|附件|签署|签章|盖章|日期|清单/.test(categoryText)) category = "常见问题";
  if (["逻辑错误", "风险语句"].includes(detailCategory)) category = "风险问题";
  if (detailCategory === "格式问题") category = "常见问题";
  const suggestion = String(issue.suggestion || "").trim();
  const replacement = String(issue.replacement_text || issue.replacementText || issue.suggested_text || "").trim();
  const cleanedReplacement = cleanReplacementText(replacement);
  const locationHint = inferReviewLocationHint({ ...issue, category, detail_category: detailCategory }, context, cleanedReplacement);
  const needsGeneratedReplacement =
    isWeakReplacementText(cleanedReplacement, suggestion) ||
    hasReviewWording(cleanedReplacement) ||
    paymentReplacementMissesContractAmount({ ...issue, category, detail_category: detailCategory }, cleanedReplacement, context);
  const replacementText = locationHint === "delete"
    ? ""
    : needsGeneratedReplacement
      ? clauseReplacementForIssue({ ...issue, category, detail_category: detailCategory }, context)
      : cleanedReplacement;
  return {
    ...issue,
    category,
    detail_category: detailCategory,
    source_rule: sourceRule || issue.source_rule || issue.sourceRule || "",
    source_rule_basis: sourceRuleBasis || (sourceRule ? "规则库" : ""),
    legal_basis: legalBasis || (sourceRule ? `依据：${sourceRule}` : ""),
    source_quote: sourceQuote,
    source_url: sourceUrl,
    location_hint: locationHint,
    replacement_text: replacementText,
  };
}

function normalizeReviewIssues(issues = [], contractText = "") {
  const context = extractReviewContext(contractText);
  context.contractText = contractText;
  const list = Array.isArray(issues) ? issues : [];
  return list.map((issue) => normalizeReviewIssue(issue, context));
}

function issueDedupKey(issue = {}) {
  return compactText([issue.title, issue.quote || issue.problem].filter(Boolean).join("|")).slice(0, 120);
}

function issueComparableText(issue = {}) {
  return compactText([issue.title, issue.problem, issue.quote, issue.source_rule || issue.sourceRule]
    .filter(Boolean)
    .join(""))
    .slice(0, 260);
}

function issueTopicKey(issue = {}) {
  const text = reviewIssueText(issue);
  if (/付款|支付|价款|金额|费用|发票|结算|款项|税费|预付款|尾款/.test(text)) return "payment";
  if (/主体|甲方|乙方|签约|名称|证照|授权|统一社会信用代码|法定代表人|联系人|地址/.test(text)) return "subject";
  if (/交付|交货|履行|服务|验收|整改|异议|质量|成果/.test(text)) return "delivery_acceptance";
  if (/违约|赔偿|责任|损失|解除|终止|逾期|免责|责任上限/.test(text)) return "liability";
  if (/保密|秘密|数据|隐私|个人信息/.test(text)) return "confidentiality";
  if (/知识产权|著作权|专利|商标|源码|成果|归属/.test(text)) return "ip";
  if (/争议|管辖|仲裁|诉讼|法律适用/.test(text)) return "dispute";
  if (/格式|编号|标题|附件|签署|盖章|日期|错别字|错字|语病|称谓|用语|排版/.test(text)) return "format";
  return compactText(issue.title || issue.detail_category || issue.category || "general").slice(0, 32);
}

function textSimilarity(a = "", b = "") {
  const left = new Set(String(a || "").match(/[\u4e00-\u9fa5]{2}|[a-zA-Z0-9]{2,}/g) || []);
  const right = new Set(String(b || "").match(/[\u4e00-\u9fa5]{2}|[a-zA-Z0-9]{2,}/g) || []);
  if (!left.size || !right.size) return 0;
  let overlap = 0;
  for (const token of left) if (right.has(token)) overlap += 1;
  return overlap / Math.min(left.size, right.size);
}

function isSimilarIssue(a = {}, b = {}) {
  const qa = compactText(a.quote || "");
  const qb = compactText(b.quote || "");
  if (qa && qb && (qa.includes(qb) || qb.includes(qa))) return true;
  if (issueTopicKey(a) !== issueTopicKey(b)) return false;
  return textSimilarity(issueComparableText(a), issueComparableText(b)) >= 0.62;
}

function betterReviewIssue(a = {}, b = {}) {
  const rank = (issue) => {
    const risk = issue.risk_level || issue.riskLevel || "中";
    const riskScore = risk === "高" ? 30 : risk === "中" ? 20 : 10;
    const quoteScore = String(issue.quote || "").length ? 8 : 0;
    const replacementScore = String(issue.replacement_text || issue.replacementText || "").length ? 8 : 0;
    const deleteScore = String(issue.location_hint || issue.locationHint || "").toLowerCase() === "delete" ? 10 : 0;
    const sourceScore = /企业|公司|自定|自定义/.test([issue.source_rule, issue.sourceRule, issue.category].filter(Boolean).join(" ")) ? 4 : 0;
    return riskScore + quoteScore + replacementScore + deleteScore + sourceScore + Math.min(10, String(issue.problem || "").length / 40);
  };
  return rank(b) > rank(a) ? b : a;
}

function mergeReviewIssues(primary = [], extra = [], contractText = "") {
  const merged = [];
  const seen = new Set();
  const primaryList = Array.isArray(primary) ? primary : [];
  const extraList = Array.isArray(extra) ? extra : [];
  for (const issue of [...primaryList, ...extraList]) {
    const key = issueDedupKey(issue);
    if (!key) continue;
    const similarIndex = merged.findIndex((item) => key === issueDedupKey(item) || isSimilarIssue(item, issue));
    if (similarIndex >= 0) {
      merged[similarIndex] = betterReviewIssue(merged[similarIndex], issue);
      seen.add(key);
      continue;
    }
    if (seen.has(key)) continue;
    seen.add(key);
    merged.push(issue);
  }
  return normalizeReviewIssues(merged, contractText);
}

function reviewDimension(issue = {}) {
  const text = reviewIssueText(issue);
  if (/主体|甲方|乙方|签约|名称|证照|授权|权限|住所|统一社会信用代码|法定代表人|联系人|联系方式/.test(text)) return "subject";
  if (/格式|编号|标题|排版|附件|签署|签章|盖章|日期|落款|页码|空白|错别字|错字|语病|称谓|用语/.test(text)) return "format";
  if (/规则库|规则|法规|法条|制度|企业|行业|惯例/.test(text)) return "rule";
  return "content";
}

function buildReviewReport(issues = [], contractText = "") {
  const counts = {
    content: 0,
    subject: 0,
    format: 0,
    rule: 0,
    high: 0,
    mid: 0,
    low: 0,
    total: issues.length,
  };
  for (const issue of issues) {
    counts[reviewDimension(issue)] += 1;
    if (issue.risk_level === "高") counts.high += 1;
    else if (issue.risk_level === "低") counts.low += 1;
    else counts.mid += 1;
  }
  const context = extractReviewContext(contractText);
  const subjectStatus = context.partyA && context.partyB ? (counts.subject ? "需核验" : "基本完整") : "缺关键主体";
  const contentStatus = counts.high ? "存在高风险" : counts.mid ? "存在中风险" : counts.content ? "需复核" : "未见明显风险";
  const formatStatus = counts.format ? "需整理" : "基本规范";
  return {
    dimensions: [
      { key: "content", label: "内容审查", count: counts.content, status: contentStatus },
      { key: "subject", label: "主体审查", count: counts.subject, status: subjectStatus },
      { key: "format", label: "格式审查", count: counts.format, status: formatStatus },
      { key: "rule", label: "规则库", count: counts.rule, status: counts.rule ? "命中规则" : "未命中" },
    ],
    subjectReview: {
      partyA: context.partyA || "",
      partyB: context.partyB || "",
      status: subjectStatus,
    },
    counts,
  };
}

function captureWorkflowParty(text = "", label = "") {
  const match = new RegExp(`${label}\\s*[：:]\\s*([^\\n；;。]{2,80})`).exec(text);
  return match ? match[1].replace(/[，,。；;].*$/, "").trim() : "";
}

function analyzeRightsObligationsForWorkflow(contractText = "", profile = {}) {
  const partyA = captureWorkflowParty(contractText, "甲方") || "甲方";
  const partyB = captureWorkflowParty(contractText, "乙方") || "乙方";
  const contractType = profile.contractType || "合同";
  const hasPayment = /付款|支付|价款|费用|报酬|合同金额|合同价款|结算/.test(contractText);
  const hasDelivery = /交付|交货|提交|完成|提供|服务|培训|开发|上线|验收/.test(contractText);
  const hasExamTraining = /报名|考试|培训|证书|报考|学习/.test(contractText);
  const hasIpData = /源代码|知识产权|数据|保密|个人信息|资料|成果归属/.test(contractText);
  if (hasExamTraining) {
    return `${contractType}中，${partyA}的核心权利是要求${partyB}提供真实、准确、完整的报名考试信息并按约缴纳相关费用；核心义务是按约为${partyB}完成报名考试安排、学习或考试事项协助、证书取得协助，并对${partyB}提交的个人资料和考试资料承担保密与合规使用义务。${partyB}的核心权利是获得${partyA}协助办理的合法有效报名、考试安排或证书取得服务；核心义务是按约缴纳费用、提供真实资料、配合考试安排并遵守考试规则。审查需重点核对报名信息责任、费用支付与退费、考试结果责任、资料保密和证书有效性。`;
  }
  const partyARight = hasDelivery
    ? `要求${partyB}按约完成合同事项并提交符合约定的成果或服务`
    : `要求${partyB}按约履行合同义务`;
  const partyADuty = hasPayment ? "按约完成验收、接收发票并支付合同价款或报酬" : "按约提供必要配合、确认履约结果并承担约定费用";
  const partyBRight = hasPayment ? "按约取得价款或报酬" : "要求甲方提供必要配合并按约确认履约结果";
  let partyBDuty = hasDelivery ? "按约完成服务、交付成果、配合验收并处理整改" : "按约履行合同事项并留存履约记录";
  if (hasExamTraining) partyBDuty = "保证报名、资料提交、学习或考试安排真实准确，并按约配合甲方完成相关流程";
  if (hasIpData) partyBDuty += "，同时遵守保密、数据安全、个人信息保护和成果权属约定";
  return `${contractType}中，${partyA}的核心权利是${partyARight}；核心义务是${partyADuty}。${partyB}的核心权利是${partyBRight}；核心义务是${partyBDuty}。审查需重点核对权利义务是否对等、付款与交付验收是否闭环、违约责任是否可执行。`;
}

function buildWorkflowKnowledgeBases(profile = {}, rules = []) {
  const contractType = profile.contractType || "通用合同";
  const bases = [`${contractType}法律法规库`, `${contractType}裁判案例库`, `${contractType}审查知识库`];
  if (profile.industry) bases.push(`${profile.industry}行业预设规则库`);
  const hasEnterprise = rules.some((rule) => ruleManagementCategory(rule) === "企业自定");
  if (hasEnterprise) bases.push("企业自定规则库");
  return [...new Set(bases)];
}

function buildReviewWorkflow(contractText = "", rules = [], profile = {}, issues = [], options = {}) {
  const counts = buildReviewReport(issues, contractText).counts;
  const riskCounts = {
    high: issues.filter((issue) => issue.risk_level === "高").length,
    mid: issues.filter((issue) => issue.risk_level === "中").length,
  };
  const replacementCount = issues.filter((issue) => issue.replacement_text || issue.replacementText).length;
  return {
    phase: options.phase || "done",
    rightsObligations: analyzeRightsObligationsForWorkflow(contractText, profile),
    knowledgeBases: buildWorkflowKnowledgeBases(profile, rules),
    ruleSelection: summarizeRuleSelection(profile, rules),
    reviewTasks: [
      { label: "实质风险识别", status: "done", note: `高/中风险 ${riskCounts.high + riskCounts.mid} 项` },
      { label: "文字符号检查", status: "done", note: `格式和用语 ${counts.format || 0} 项` },
      { label: "生成修改建议", status: "done", note: `可应用正文 ${replacementCount} 条` },
      { label: "签约主体审查", status: "done", note: `主体问题 ${counts.subject || 0} 项` },
    ],
    engine: options.engine || "",
  };
}

function uniqueAmountValues(amountList = []) {
  const values = amountList
    .map((item) => item.value)
    .filter((value) => Number.isFinite(value) && value > 0)
    .map((value) => Math.round(value * 100) / 100);
  return [...new Set(values)];
}

function hasPaymentAllocation(contractText = "") {
  return /分期|付款节点|付款比例|阶段款|进度款|预付款|首付款|验收款|尾款|质保金|里程碑|30%|60%|10%|百分之/.test(contractText);
}

function buildAmountReviewIssues(contractText = "", context = {}, hasPayment = false) {
  const issues = [];
  const quote = context.amountSentence || context.paymentSentence || context.subjectSentence || "";
  const amountValues = uniqueAmountValues(context.amountList || []);
  const hasAmount = Boolean(context.amount);
  if (hasPayment && !hasAmount) {
    issues.push({
      risk_level: "高",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "合同价款金额缺失或不可计算",
      problem: "合同存在付款、价款或费用安排，但未识别到明确的合同总价、币种和金额，后续付款、违约金、责任上限及结算依据均可能无法执行。",
      suggestion: "补充合同总价、币种、计价口径、含税/不含税口径及结算依据。",
      replacement_text: "",
      source_rule: "金额审查规则",
      location_hint: "supplement",
      quote,
    });
  }
  if (hasPayment && hasAmount && !/(人民币|RMB|￥|¥|元|万元|亿元)/.test(context.amount)) {
    issues.push({
      risk_level: "中",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "金额币种或单位不明确",
      problem: `合同出现“${context.amount}”等金额数字，但未明确币种或计量单位，可能导致付款金额、结算口径和税费承担产生歧义。`,
      suggestion: "将金额写明为人民币元、万元或其他明确币种单位，并同步明确含税口径。",
      replacement_text: "",
      source_rule: "金额审查规则",
      location_hint: "replace",
      quote: context.amount,
    });
  }
  if (hasPayment && hasAmount && !hasPaymentAllocation(contractText)) {
    issues.push({
      risk_level: "中",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "付款金额未按节点分配",
      problem: "合同已约定金额或价款，但未按预付款、交付验收款、尾款等节点拆分金额比例，无法清楚判断每一节点应付金额和付款条件。",
      suggestion: "按付款规则拆分付款节点，并将每个节点与发票、交付验收和结算异议处理绑定。",
      replacement_text: "",
      source_rule: "付款分配规则",
      location_hint: "supplement",
      quote,
    });
  }
  if (amountValues.length > 1 && !hasPaymentAllocation(contractText)) {
    issues.push({
      risk_level: "中",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "多处金额表述可能不一致",
      problem: `合同中识别到多个金额值（${amountValues.map(formatMoney).join("、")}），但未说明总价、单价、分期款或结算口径之间的关系，可能造成付款依据冲突。`,
      suggestion: "区分合同总价、单价、阶段款、尾款及税费口径，必要时以表格或附件统一列明。",
      replacement_text: "",
      source_rule: "金额一致性规则",
      location_hint: "supplement",
      quote,
    });
  }
  return issues;
}

function buildSubjectReviewIssues(contractText = "", context = {}) {
  const issues = [];
  const partyLine = findSentence(contractText, /甲方|乙方|签约主体|合同主体|委托方|受托方|采购方|供应商/);
  const hasPartyLabel = /甲方|乙方|委托方|受托方|采购方|供应商|出租方|承租方|出借方|借款方/.test(contractText);
  const hasBothParties = Boolean(context.partyA && context.partyB);
  const hasCredential = /统一社会信用代码|营业执照|身份证号|证件号码|住所|地址|法定代表人|负责人|联系人|联系方式|授权代表/.test(partyLine || contractText.slice(0, 800));
  const hasSignature = /签字|签署|盖章|签章|公章|合同专用章|法定代表人|授权代表/.test(contractText);

  if (!hasPartyLabel || !hasBothParties) {
    issues.push({
      risk_level: "高",
      category: "风险问题",
      detail_category: "主体",
      title: "签约主体信息缺失",
      problem: "合同未完整识别到甲乙双方或签约主体信息，权利义务归属、签署权限和后续送达可能无法确认。",
      suggestion: "补充双方完整名称、统一社会信用代码或证件号码、住所、联系人、联系方式及签署授权信息。",
      replacement_text: "",
      source_rule: "主体审查清单",
      location_hint: "supplement",
      quote: partyLine,
    });
  } else if (!hasCredential) {
    issues.push({
      risk_level: "中",
      category: "风险问题",
      detail_category: "主体",
      title: "主体证照和联系信息不完整",
      problem: "合同已列明双方名称，但未明显载明统一社会信用代码、住所、联系人、联系方式或授权代表信息，主体核验和送达执行存在缺口。",
      suggestion: "在主体条款中补充证照号码、注册地址或住所、联系人、联系方式，并确认签署人已取得合法授权。",
      replacement_text: "",
      source_rule: "主体审查清单",
      location_hint: "supplement",
      quote: partyLine,
    });
  }

  if (!hasSignature) {
    issues.push({
      risk_level: "中",
      category: "常见问题",
      detail_category: "格式问题",
      title: "签署栏或盖章信息缺失",
      problem: "合同未明显识别到签字、盖章、授权代表或签署日期信息，可能影响合同生效证据和主体授权证明。",
      suggestion: "在合同末尾补充双方签署栏，列明签署人、盖章位置和签署日期。",
      replacement_text: "",
      source_rule: "格式审查清单",
      location_hint: "supplement",
      quote: partyLine || context.liabilitySentence || "",
    });
  }
  return issues;
}

function buildLogicReviewIssues(contractText = "") {
  const context = extractReviewContext(contractText);
  const issues = [];
  const hasPayment = /付款|支付|价款|费用|合同金额|合同价款|款项|结算/.test(contractText);
  const hasAmount = Boolean(context.amount);
  const hasPaymentCondition = /验收合格后|验收通过后|收到发票后|发票.*后|付款节点|付款条件|分期|阶段款|里程碑|预付款|尾款|结算后|账期|工作日内支付|日内支付/.test(contractText);
  const hasInvoice = /发票|开票|税率|含税|不含税|增值税/.test(contractText);
  const hasDelivery = /交付|交货|服务|履行|提交|完成|提供/.test(contractText);
  const hasAcceptance = /验收|审核|确认|异议|整改/.test(contractText);
  const upfrontPayment = findSentence(contractText, /签订.{0,30}(支付|付清|一次性)|一次性.{0,20}(支付|付清)|预付.{0,20}(全款|全部)|先款后货|先付款后/);
  const hasAcceptanceStandard = /验收标准|合格标准|质量标准|书面验收|验收期限|异议期|整改|复验/.test(contractText);
  const hasDispute = /争议解决|管辖|仲裁|诉讼|人民法院|仲裁委员会|法律适用/.test(contractText);
  const hasPartyALiability = /甲方.{0,30}(违约|赔偿|责任|损失|逾期|解除)/.test(contractText);
  const hasPartyBLiability = /乙方.{0,30}(违约|赔偿|责任|损失|逾期|解除)/.test(contractText);
  const finalInterpretationQuote = findSentence(contractText, /最终解释权归|解释权归.{0,12}所有/);
  const riskQuote =
    context.riskSentence ||
    findSentence(contractText, /最终解释权归|解释权归.*所有|不承担任何责任|概不负责|全部损失|所有损失|无限责任|无条件|永久|不可撤销|可随时|单方.{0,8}(变更|解除|终止|暂停)/);

  issues.push(...buildSubjectReviewIssues(contractText, context));
  issues.push(...buildAmountReviewIssues(contractText, context, hasPayment));

  if ((hasPayment || hasAmount) && !hasPaymentCondition) {
    issues.push({
      risk_level: "高",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "付款条件未与交付验收闭合",
      problem: "合同存在价款或付款安排，但未明确付款触发条件、验收通过、发票资料和付款期限之间的衔接，容易产生先付款后交付或付款争议。",
      suggestion: "将付款节点与交付、验收合格、发票及完整付款资料绑定，并明确异议处理期限。",
      replacement_text: "",
      source_rule: "本地逻辑复核",
      location_hint: "supplement",
      quote: context.paymentSentence || context.deliverySentence || "",
    });
  }
  if (upfrontPayment && hasDelivery) {
    issues.push({
      risk_level: "高",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "付款顺序与交付验收风险不匹配",
      problem: "合同存在签约后一次性付款、先付款后交付或类似安排，但同时存在交付/服务履行事项，未设置验收合格、交付完成或资料齐备作为付款前提，可能造成付款后履约控制不足。",
      suggestion: "将付款拆分或绑定交付、验收、发票及付款资料，至少保留尾款或验收后付款节点。",
      replacement_text: "",
      source_rule: "分层逻辑审查",
      location_hint: "replace",
      quote: upfrontPayment,
    });
  }
  if ((hasPayment || hasAmount) && !hasInvoice) {
    issues.push({
      risk_level: "中",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "价款税费和发票要求不完整",
      problem: "合同涉及价款或费用，但未明确含税/不含税、发票类型、开票时间及付款资料，可能影响付款条件和财务入账。",
      suggestion: "补充税费承担、发票类型、开票节点和付款资料要求。",
      replacement_text: "",
      source_rule: "本地逻辑复核",
      location_hint: "supplement",
      quote: context.paymentSentence || "",
    });
  }
  if (hasDelivery && !hasAcceptance) {
    issues.push({
      risk_level: "高",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "交付后缺少验收和异议闭环",
      problem: "合同包含交付或服务履行内容，但未明确验收期限、验收标准、异议提出方式和整改机制，后续难以判断是否完成履约。",
      suggestion: "补充交付验收标准、验收期限、异议通知和整改复验流程。",
      replacement_text: "",
      source_rule: "本地逻辑复核",
      location_hint: "supplement",
      quote: context.deliverySentence || "",
    });
  }
  if (hasAcceptance && !hasAcceptanceStandard) {
    issues.push({
      risk_level: "中",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "验收条款缺少可执行标准",
      problem: "合同虽提到验收或确认，但未明确验收标准、验收期限、书面确认、异议期或整改复验机制，后续可能无法判断交付成果是否合格。",
      suggestion: "补充验收标准、验收期限、书面验收方式、异议处理和整改复验机制。",
      replacement_text: "",
      source_rule: "分层逻辑审查",
      location_hint: "supplement",
      quote: context.acceptanceSentence || context.deliverySentence || "",
    });
  }
  if (riskQuote) {
    issues.push({
      risk_level: "高",
      category: "风险问题",
      detail_category: "风险语句",
      title: "存在单方或绝对化风险表述",
      problem: "原文存在最终解释权、免责过宽、责任无上限、无条件/永久/不可撤销或单方变更解除等表述，可能被认定为排除对方主要权利、免除自身责任或显失公平。",
      suggestion: "改为双方协商一致、依法合理、责任边界清晰的表述，避免单方决定和绝对化承诺。",
      replacement_text: "",
      source_rule: "本地逻辑复核",
      location_hint: "replace",
      quote: riskQuote,
    });
  }
  if (finalInterpretationQuote) {
    issues.push({
      risk_level: "高",
      category: "风险问题",
      detail_category: "风险语句",
      title: "删除单方最终解释权表述",
      problem: "原文约定最终解释权归一方所有，属于明显单方化表述，可能排除相对方解释和救济空间，存在显失公平和无效风险。",
      suggestion: "删除该单方解释权表述；如确需解释规则，应改由双方协商一致或依法按合同文本、交易习惯和诚实信用原则解释。",
      replacement_text: "",
      source_rule: "本地风险语句审查",
      location_hint: "delete",
      quote: finalInterpretationQuote,
    });
  }
  if (/\n[ \t]*\n[ \t]*\n/.test(contractText)) {
    issues.push({
      risk_level: "低",
      category: "常见问题",
      detail_category: "格式问题",
      title: "存在不必要空行",
      problem: "合同正文存在连续空行或段落间距过大，影响正式合同版式和阅读连续性。",
      suggestion: "删除多余空行，正文段落之间保留必要间距，条款层级保持连续。",
      replacement_text: "",
      source_rule: "格式审查清单",
      location_hint: "delete",
      quote: "",
    });
  }
  if (/甲方/.test(contractText) && /乙方/.test(contractText) && /(委托方|受托方|服务方|供应商|客户|买方|卖方)/.test(contractText) && !/以下简称/.test(contractText)) {
    issues.push({
      risk_level: "中",
      category: "常见问题",
      detail_category: "格式问题",
      title: "主体称谓可能前后不一致",
      problem: "合同同时使用甲方/乙方及委托方、受托方、服务方、供应商等称谓，且未明显定义简称，容易造成权利义务主体不清。",
      suggestion: "在合同开头统一定义各方简称，并在全文保持同一称谓。",
      replacement_text: "",
      source_rule: "本地逻辑复核",
      location_hint: "supplement",
      quote: findSentence(contractText, /甲方|乙方|委托方|受托方|服务方|供应商|客户|买方|卖方/),
    });
  }
  if ((hasPartyALiability || hasPartyBLiability) && hasPartyALiability !== hasPartyBLiability) {
    issues.push({
      risk_level: "中",
      category: "风险问题",
      detail_category: "逻辑错误",
      title: "违约责任可能单边化",
      problem: "合同仅明显约定一方违约责任，另一方逾期付款、逾期交付、质量不合格、提前解除等责任不完整，可能导致权利义务失衡。",
      suggestion: "补充双方对等的违约责任，分别覆盖付款、交付、验收、保密、知识产权、解除终止等核心义务。",
      replacement_text: "",
      source_rule: "分层逻辑审查",
      location_hint: "supplement",
      quote: context.liabilitySentence || "",
    });
  }
  if (!hasDispute) {
    issues.push({
      risk_level: "中",
      category: "常见问题",
      detail_category: "逻辑错误",
      title: "争议解决条款缺失",
      problem: "合同未明显约定法律适用和争议解决方式，发生争议时可能增加管辖和维权成本。",
      suggestion: "补充适用法律、协商机制和明确的诉讼或仲裁管辖。",
      replacement_text: "",
      source_rule: "分层逻辑审查",
      location_hint: "supplement",
      quote: context.liabilitySentence || context.paymentSentence || context.deliverySentence || "",
    });
  }
  return issues;
}

function buildReviewPrompt(contractText, rules, precheckIssues = [], profile = {}) {
  return `你是资深中国合同律师和企业法务审查助手，也是本系统的大模型深度审查层。请结合规则库、常见问题清单、本地预检线索和合同上下文进行严格审查，不只检查缺失条款，还要审查已有条款中的商业逻辑、法律风险、用语错误和不合理表述。

审查方法采用主流智能合同平台的分层扫描思路：
第一层：合同基础要素，检查主体、授权、标的、数量、金额、期限、附件、签署信息是否完整。
第二层：交易闭环，检查付款-开票-交付-验收-整改-违约-解除是否前后衔接。
第三层：法律风险，检查免责过宽、责任无上限、单方变更/解除、格式条款显失公平、保密/数据/知识产权不足。
第四层：规则库和企业口径，检查是否违反规则库、审批权限、行业惯例、企业自定规则。
第五层：用语和格式，检查错别字、称谓不一致、期限/对象/金额不明、附件引用和签署格式问题。
第六层：主体审查，核验甲乙方名称、统一社会信用代码或证件号、住所、联系人、授权代表、签署盖章信息是否完整一致；无法联网查询主体时，不要编造第三方数据，只输出需要核验的缺口。

合同识别与规则调用策略：
${summarizeRuleSelection(profile, rules)}
请先按合同内容判断合同类型和所属行业，再按“通用规则和法律底线 → 匹配行业预设规则 → 企业自定规则”的顺序审查。当前行业枚举暂为：软件外包、系统集成、制造业。所有审查问题都要尽量给出可信依据：可引用通用法律法规原则、平台预设规则、行业预设规则、企业自定义规则或常见问题清单；不要编造具体法条编号，不确定时写“合同法理与交易闭环审查”或对应规则名称。

规则库：
${summarizeRulesForPrompt(rules, 60)}

本地预检线索：
${summarizeReviewPrecheckForPrompt(precheckIssues)}

常见问题：
${COMMON_REVIEW_ISSUES.map((item, index) => `${index + 1}. ${item}`).join("\n")}

合同文本：
${contractText.slice(0, 120000)}

输出严格 JSON：
{
  "overall_risk": "高/中/低",
  "summary": "一句话审查摘要",
  "issues": [
    {
      "risk_level": "高/中/低",
      "category": "风险问题/规则库/常见问题",
      "detail_category": "逻辑错误/风险语句/格式问题/主体/付款/交付/知识产权/保密/违约/争议解决/其他",
      "title": "问题标题",
      "problem": "发现的问题",
      "suggestion": "处理意见",
      "replacement_text": "可以直接替换或补充进合同的正文条款文本",
      "source_rule": "引用的规则名称或常见问题",
      "source_rule_basis": "通用法规/行业惯例/企业自定/平台预设/常见问题",
      "legal_basis": "法律法规或法理依据的简短说明",
      "source_quote": "规则库或上传资料中的依据片段",
      "location_hint": "replace/supplement/delete",
      "quote": "合同中的相关原文片段"
    }
  ]
}

要求：
1. 你必须在本地预检线索基础上独立复核和补充，不要机械照搬；预检成立时要给出更准确的 quote、problem 和 replacement_text，预检不成立时可以忽略。
2. 优先指出高风险、可执行性差、缺失关键条款、与规则库冲突、权利义务失衡、前后逻辑矛盾、用语不准确、风险表述不合理的问题。
3. 一级 category 只能使用“风险问题 / 规则库 / 常见问题”。细分 detail_category 优先使用“逻辑错误 / 风险语句 / 格式问题”；主体、付款、交付等要素类问题可继续使用对应细分。格式问题统一归入“常见问题”，并在 detail_category 写“格式问题”。中风险和高风险都归入“风险问题”，risk_level 保留高/中/低详细等级。source_rule、source_rule_basis、legal_basis 必须尽量填写，用于前端展示可信依据。
4. 审查力度要覆盖：
   - 主体审查：甲乙方名称、证照号码、住所、联系人、授权代表、盖章签署栏、主体称谓一致性。
   - 商业逻辑：付款与交付/验收/开票条件是否闭合，解除与违约责任是否衔接，附件与正文是否冲突，主体与签署权限是否一致。
   - 用语问题：错别字、称谓不一致、同一概念多种叫法、语病、对象不明、期限不明、金额/比例/币种不清。
   - 金额问题：金额缺失、币种或单位不明、大小写金额不一致、多处金额冲突、总价/单价/阶段款关系不清、付款金额未按节点分配。
   - 风险表述：无条件、永久、不可撤销、全部责任、所有损失、最终解释权、不承担任何责任、可随时单方变更/暂停/解除、责任无上限或免责过宽。
   - 合法合规：格式条款排除主要权利、免除主要责任、加重对方责任，个人信息/数据/保密/知识产权约定不足。
   - 交易闭环：先款后货、一次性付款、无验收尾款、只有交付无验收、只有违约无责任上限、只有一方义务无对方义务。
5. replacement_text 必须是一段可以直接放进合同的合理合法正文条款，不能照搬 quote，不能和 suggestion 相同，不能出现“建议、请、可考虑、需要人工确认、应补充、修改建议”等审查提示语。如果 quote 是问题原文，replacement_text 要结合合同真实业务、金额、主体、交付和验收内容改写成完整句子；如果是缺失条款，replacement_text 要写成可新增的完整条款。
6. quote 必须短而准：付款、价款、账期、发票问题定位到含具体金额、付款节点或费用表述的最短原文句；错别字、用词、语病问题定位到包含错误字词的完整原文句；风险表述定位到直接产生风险的原文短句，不要返回整段。
7. 付款类 replacement_text 应按规则自动拆分付款节点；识别到合同总价时要计算每个节点对应金额，默认可采用 30% 预付款、60% 交付验收款、10% 尾款，并绑定发票、验收和异议处理。金额缺失、币种单位不明、大小写不一致或多处金额冲突必须形成独立 issue。
8. 需要新增或补充条款时，location_hint 写 supplement，quote 写“应该插入位置附近的短句”（例如主体缺失定位到甲方/乙方/签约主体附近，付款节点缺失定位到金额或付款条款附近，签署栏缺失定位到合同末尾或签署位置附近）；replacement_text 只写应补充的新条款。
9. 原文已有相关表述但表述错误、逻辑不合理、风险过高、用词不准或格式不对时，location_hint 必须写 replace，quote 写需要被替换的原文短句，replacement_text 写替换后的合同正文，不要写 supplement。
10. 需要删除原文内容时，location_hint 写 delete，quote 写应删除的最短原文；replacement_text 为空字符串。删除类问题包括最终解释权归一方所有、重复条款、无效或明显不公平表述、多余连续空行、无关内容。
11. replacement_text 示例：
   - 不合格：建议明确付款节点。
   - 合格：付款安排：本合同价款为人民币100,000元。甲方按30%预付款、60%交付验收款、10%尾款向乙方付款；乙方应按约开具合法有效发票及完整付款资料，甲方在收到无争议付款资料后十个工作日内支付对应款项。
   - 不合格：建议删除绝对化表述。
   - 合格：双方行使合同权利和承担合同责任应以法律法规允许、合同目的实现及公平合理为原则；任何一方不得以单方条款排除对方主要权利或免除自身主要责任。
12. 同时检查格式常见问题，包括标题、编号、条款层级、附件引用、签署栏、日期、盖章位置、正文与附件一致性、连续空行和不必要空白。格式问题统一归入常见问题。
13. 不要输出 Markdown，不要输出 JSON 之外的文字。
14. 如果没有明显问题，也要列出建议人工复核的关注点。`;
}

async function handleGenerateContract(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const description = String(payload.description || "").trim();
  if (!description) return sendJson(res, 400, { error: "请先输入合同生成意图或业务描述" });
  const store = loadStore();
  const availableTemplates = templatesForPayload(payload);
  const templateChoice = chooseContractTemplate(description, payload.templateId, availableTemplates);
  const { template } = templateChoice;
  const profile = inferContractProfile(description, template);
  let answers = normalizeAnswers({ ...inferAnswersFromDescription(description), ...(payload.answers || {}) });
  const initialProfileText = [description, JSON.stringify(answers), template.name].join("\n");
  const initialRuleProfile = inferContractProfile(initialProfileText, template);
  const initialRules = selectRulesForContract(store, initialRuleProfile);
  const initialSnippets = selectKnowledgeSnippetsForGeneration(store, initialProfileText, { limit: 10 });
  let extractedAnswers = {};
  try {
    const extraction = await callJsonModel(buildAnswerExtractionPrompt(template, description, answers, initialRules, initialSnippets, initialRuleProfile), {
      repair: true,
      label: "合同生成字段提炼结果",
    });
    extractedAnswers = normalizeModelAnswers(extraction.data, template);
    answers = normalizeAnswers({ ...answers, ...(payload.answers || {}), ...extractedAnswers });
  } catch {
    // 不中断生成；后续草稿生成仍会使用本地抽取和用户已确认答案。
  }
  const missingFields = detectMissingFields(template, answers);
  recordGenerationMemory(store, template, description, answers);
  if (missingFields.length && !payload.forceDraft) {
    const field = missingFields[0];
    const suggestions = buildFieldSuggestions(store, template, field, answers, description);
    saveStore(store);
    return sendJson(res, 200, {
      status: "need_more_info",
      template: publicTemplate(template),
      requestedTemplate: templateChoice.requestedTemplate ? publicTemplate(templateChoice.requestedTemplate) : null,
      templateSwitched: templateChoice.templateSwitched,
      knownAnswers: answers,
      missingFields,
      question: {
        key: field.key,
        label: field.label,
        question: field.question,
        suggestions,
      },
      questions: [
        {
          key: field.key,
          label: field.label,
          question: field.question,
          suggestions,
        },
      ],
      matchConfidence: templateChoice.matchConfidence,
      templateCandidates: templateChoice.templateCandidates,
      ruleProfile: profile,
    });
  }

  const profileText = [description, JSON.stringify(answers), template.name].join("\n");
  const ruleProfile = inferContractProfile(profileText, template);
  const rules = selectRulesForContract(store, ruleProfile);
  const knowledgeSnippets = selectKnowledgeSnippetsForGeneration(store, profileText, { limit: 12 });
  saveStore(store);
  try {
    const modelCall = await callChatModel(buildContractGenerationPrompt(template, description, answers, rules, ruleProfile, {
      existingDraft: payload.existingDraft,
      knowledgeSnippets,
    }));
    const result = safeJsonFromModel(modelCall.content);
    const draft = cleanGeneratedContractDraft(result.draft) || fallbackContractDraft(template, answers, rules);
    return sendJson(res, 200, {
      status: "draft_ready",
      template: publicTemplate(template),
      requestedTemplate: templateChoice.requestedTemplate ? publicTemplate(templateChoice.requestedTemplate) : null,
      templateSwitched: templateChoice.templateSwitched,
      matchConfidence: templateChoice.matchConfidence,
      templateCandidates: templateChoice.templateCandidates,
      draft,
      title: result.contract_title || template.name,
      appliedRules: result.applied_rules?.length ? result.applied_rules : rules.slice(0, 8).map((rule) => rule.ruleName),
      missingWarnings: result.missing_warnings || [],
      ruleProfile,
      ruleSelection: summarizeRuleSelection(ruleProfile, rules),
      extractedAnswers,
      knowledgeSources: knowledgeSnippets.slice(0, 5).map((item) => item.source),
      model: modelCall.model,
      modelProvider: modelCall.provider,
    });
  } catch (error) {
    const modelInfo = activeModelInfo();
    return sendJson(res, 200, {
      status: "draft_ready",
      template: publicTemplate(template),
      requestedTemplate: templateChoice.requestedTemplate ? publicTemplate(templateChoice.requestedTemplate) : null,
      templateSwitched: templateChoice.templateSwitched,
      matchConfidence: templateChoice.matchConfidence,
      templateCandidates: templateChoice.templateCandidates,
      draft: fallbackContractDraft(template, answers, rules),
      title: template.name,
      appliedRules: rules.slice(0, 8).map((rule) => rule.ruleName),
      missingWarnings: missingFields.map((field) => field.label),
      ruleProfile,
      ruleSelection: summarizeRuleSelection(ruleProfile, rules),
      extractedAnswers,
      knowledgeSources: knowledgeSnippets.slice(0, 5).map((item) => item.source),
      usedFallback: true,
      fallbackReason: error.message,
      model: modelInfo.model,
      modelProvider: modelInfo.provider,
    });
  }
}

// ═══════════════════════════════════════════════════════════════
// Smart Draft — Multi-Agent Contract Generation System
// ═══════════════════════════════════════════════════════════════

const builtinTemplatesPath = path.join(ROOT, "builtin-templates.json");
let builtinTemplates = [];
try { builtinTemplates = JSON.parse(fs.readFileSync(builtinTemplatesPath, "utf-8")); } catch {}

// In-memory session store for smart draft sessions
const smartDraftSessions = new Map();

function createSmartDraftSession() {
  const id = "sd_" + crypto.randomBytes(8).toString("hex");
  const session = {
    id,
    createdAt: Date.now(),
    description: "",
    knowledgeMode: "industry", // "industry" | "enterprise"
    guidanceMode: "ask_user",   // "ask_user" | "kb_search" | "llm_infer"
    intent: null,      // Agent 1 output
    extractedData: null, // Agent 2 output
    rules: [],          // Agent 3 output
    knowledgeSnippets: [], // Agent 3 output
    template: null,     // matched builtin template
    answers: {},        // accumulated field answers
    missingFields: [],  // fields still needing info
    currentQuestion: null,
    draft: "",          // current draft text
    messages: [],       // chat history
    step: "idle",       // idle|intent|extract|knowledge|generate|iterate|done
    files: [],          // uploaded file info
  };
  smartDraftSessions.set(id, session);
  // Cleanup old sessions (>2h)
  for (const [k, v] of smartDraftSessions) {
    if (Date.now() - v.createdAt > 7200000) smartDraftSessions.delete(k);
  }
  return session;
}

// ── Agent 1: Intent Analyzer ──
async function agentIntentAnalyze(description, fileNames) {
  const prompt = `你是一个合同意图分析专家。根据用户描述和文件列表，分析合同类型和关键信息。

用户描述: ${description || "(未提供)"}
上传文件: ${fileNames && fileNames.length ? fileNames.join(", ") : "(未上传)"}

可选合同类型: goods(采购), service(服务), software(软件开发), labor(劳动), rental(租赁), sales(销售), consulting(咨询), agency(代理), nda(保密), framework(框架)

请返回JSON格式:
{
  "contractType": "goods/service/software/...",
  "contractTypeCn": "采购合同/服务合同/...",
  "industry": "IT行业/制造业/通用/...",
  "parties": ["甲方名称(如有)", "乙方名称(如有)"],
  "keyTerms": ["关键词1", "关键词2"],
  "complexity": "simple/medium/complex",
  "summary": "一句话总结合同意图"
}`;

  try {
    const result = await callJsonModel(prompt, null, { temperature: 0.1, maxTokens: 1000 });
    if (result) return result.data || result;
  } catch (e) {
    console.error("[Agent1] Intent analysis failed:", e.message);
  }

  // Fallback: keyword matching
  const typeMap = {
    goods: ["采购", "购买", "设备", "物料"],
    service: ["服务", "运维", "支持"],
    software: ["软件", "开发", "系统", "平台"],
    labor: ["劳动", "雇佣", "员工"],
    rental: ["租赁", "出租", "房屋"],
    sales: ["销售", "出售", "产品"],
    consulting: ["咨询", "顾问"],
    agency: ["代理", "委托", "经销"],
    nda: ["保密", "秘密", "机密"],
    framework: ["框架", "战略", "合作"],
  };
  let bestType = "service", bestScore = 0;
  for (const [type, keywords] of Object.entries(typeMap)) {
    const score = keywords.filter((k) => (description || "").includes(k)).length;
    if (score > bestScore) { bestScore = score; bestType = type; }
  }
  const tpl = builtinTemplates.find((t) => t.type === bestType);
  return {
    contractType: bestType,
    contractTypeCn: tpl ? tpl.name : "服务合同",
    industry: "通用",
    parties: [],
    keyTerms: (description || "").slice(0, 50).split(/[,，、\s]+/).filter(Boolean),
    complexity: "medium",
    summary: description || "合同起草",
  };
}

// ── Agent 2: Data Extractor ──
async function agentDataExtract(fileTexts, intent) {
  if (!fileTexts || fileTexts.length === 0) return { items: [], fields: {}, summary: "无参考文件" };

  const combined = fileTexts.map((f, i) => `--- 文件${i + 1}: ${f.name} ---\n${f.text.slice(0, 3000)}`).join("\n\n");

  const prompt = `你是一个合同数据提取专家。从以下上传文件中提取与"${intent.contractTypeCn || '合同'}"相关的结构化数据。

${combined}

请返回JSON格式:
{
  "items": [{"name":"品名","spec":"规格","qty":数量,"unit":"单位","price":单价,"amount":金额,...}],
  "fields": {"partyA":"甲方名称","partyB":"乙方名称","totalAmount":"总金额","...":"..."},
  "summary": "提取结果一句话总结"
}

如果没有可提取的数据，items和fields返回空对象。`;

  try {
    const result = await callJsonModel(prompt, null, { temperature: 0.1, maxTokens: 2000 });
    if (result) return result.data || result;
  } catch (e) {
    console.error("[Agent2] Data extraction failed:", e.message);
  }
  return { items: [], fields: {}, summary: "提取失败" };
}

// ── Agent 3: Knowledge Retriever ──
function agentKnowledgeRetrieve(intent, knowledgeMode) {
  const contractType = intent.contractType || "service";
  const industry = intent.industry || "通用";

  // Select rules based on mode
  const currentStore = loadStore();
  const allRules = currentStore.rules || [];
  let rules = allRules.filter((r) => r.reviewStatus === "active" || allRules.every((x) => x.reviewStatus !== "active"));

  if (knowledgeMode === "enterprise") {
    rules = rules.filter((r) => r.ruleBasis === "企业自定");
  } else {
    // industry mode: include 行业惯例 + 通用法规
    rules = rules.filter((r) => r.ruleBasis !== "企业自定");
  }

  // Filter by contract type match
  const profile = { contractType, industry };
  rules = rules.filter((r) => ruleMatchesProfile(r, profile)).slice(0, 20);

  // Select knowledge snippets
  const keyTerms = (intent.keyTerms || []).join(" ");
  const snippets = selectKnowledgeSnippetsForGeneration(currentStore, keyTerms, 10);

  return { rules, snippets };
}

// ── Agent 4: Draft Generator ──
async function agentDraftGenerate(session) {
  const { template, extractedData, rules, knowledgeSnippets, answers, guidanceMode, description } = session;
  const tpl = template || builtinTemplates[0];

  // Build field value summary
  const answerLines = Object.entries(answers).map(([k, v]) => `- ${k}: ${v}`).join("\n") || "（暂无）";

  // Build items summary
  let itemsSummary = "（无附件数据）";
  if (extractedData && extractedData.items && extractedData.items.length > 0) {
    itemsSummary = extractedData.items.map((it, i) =>
      `${i + 1}. ${it.name || ""} ${it.spec || ""} × ${it.qty || ""}${it.unit || ""} @${it.price || ""} = ${it.amount || ""}`
    ).join("\n");
  }

  // Build rules summary
  const ruleLines = rules.slice(0, 10).map((r) => `- [${r.ruleBasis || "通用"}] ${r.ruleName}: ${(r.action || "").slice(0, 100)}`).join("\n") || "（无适用规则）";

  // Guidance mode instruction
  let guidance = "";
  if (guidanceMode === "ask_user") {
    guidance = "请在合同中标注所有不确定的信息为【待确认：xxx】格式，并在最后列出需要用户确认的问题清单。";
  } else if (guidanceMode === "kb_search") {
    guidance = "对于不确定的信息，尝试从知识库内容中找到合理答案填入，并标注信息来源。仍无法确定的标注【待确认：xxx】。";
  } else {
    guidance = "对于不确定的信息，请根据上下文和行业惯例自行推断最合理的答案，直接填入合同，不要留空。";
  }

  const prompt = `你是一名资深中国合同律师。请根据以下信息起草一份完整的${tpl.name}。

## 合同描述
${description || "（用户未提供描述）"}

## 合同类型: ${tpl.name}
## 必备条款: ${tpl.requiredClauses.join("、")}

## 已知信息
${answerLines}

## 附件提取数据
${itemsSummary}

## 适用规则（${rules.length}条）
${ruleLines}

## 知识库参考
${knowledgeSnippets.map((s) => `- ${s}`).join("\n") || "（无）"}

## 起草指引
${guidance}

## 格式要求
1. 使用中文合同格式
2. 条款编号连续
3. 金额使用大写+小写
4. 日期格式: YYYY年MM月DD日
5. 签名栏包含: 甲方(盖章)、法定代表人/授权代表、日期

请直接输出合同全文（纯文本，不加markdown格式）：`;

  try {
    const result = await callChatModel(prompt, null, { temperature: 0.3, maxTokens: 8000 });
    if (result) return result;
  } catch (e) {
    console.error("[Agent4] Draft generation failed:", e.message);
  }

  // Fallback: basic structure
  return fallbackSmartDraft(tpl, answers, extractedData);
}

// ── Agent 5: Dialogue Manager ──
function agentDialogueManage(session) {
  const tpl = session.template || builtinTemplates[0];
  const required = tpl.requiredClauses || [];
  const answers = session.answers || {};

  // Find missing fields
  const missing = [];
  const fieldHints = {
    "合同主体": ["甲方", "乙方"],
    "采购标的": ["品名", "规格", "数量", "单价"],
    "质量标准": ["质量", "验收"],
    "交货": ["交货", "交付", "交付时间"],
    "付款": ["付款", "支付"],
    "违约": ["违约", "赔偿"],
    "保密": ["保密"],
    "期限": ["期限", "有效期", "合同期"],
  };

  for (const clause of required) {
    const matched = Object.keys(fieldHints).find((k) => clause.includes(k));
    if (matched) {
      const hints = fieldHints[matched];
      const hasValue = hints.some((h) =>
        Object.keys(answers).some((k) => k.includes(h) && answers[k])
      );
      if (!hasValue) {
        missing.push({ clause, hint: matched, question: `请提供${clause}相关信息` });
      }
    }
  }

  // Also check extracted data coverage
  if (session.extractedData && session.extractedData.fields) {
    const ed = session.extractedData.fields;
    if (ed.partyA && !answers["甲方"]) answers["甲方"] = ed.partyA;
    if (ed.partyB && !answers["乙方"]) answers["乙方"] = ed.partyB;
    if (ed.totalAmount && !answers["总金额"]) answers["总金额"] = ed.totalAmount;
  }

  session.missingFields = missing;
  session.currentQuestion = missing.length > 0 ? missing[0] : null;

  return {
    missingCount: missing.length,
    currentQuestion: session.currentQuestion,
    totalRequired: required.length,
    filledCount: required.length - missing.length,
    progress: Math.round(((required.length - missing.length) / required.length) * 100),
  };
}

// ── Fallback draft generation ──
function fallbackSmartDraft(tpl, answers, extractedData) {
  const partyA = answers["甲方"] || answers["partyA"] || "【甲方名称】";
  const partyB = answers["乙方"] || answers["partyB"] || "【乙方名称】";
  const today = new Date();
  const dateStr = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;

  let itemsTable = "";
  if (extractedData && extractedData.items && extractedData.items.length > 0) {
    itemsTable = "\n序号 | 品名 | 规格型号 | 数量 | 单位 | 单价(元) | 金额(元)\n";
    itemsTable += "--- | --- | --- | --- | --- | --- | ---\n";
    extractedData.items.forEach((it, i) => {
      itemsTable += `${i + 1} | ${it.name || ""} | ${it.spec || ""} | ${it.qty || ""} | ${it.unit || ""} | ${it.price || ""} | ${it.amount || ""}\n`;
    });
  }

  let draft = `${tpl.name}\n\n`;
  draft += `甲方（买方）：${partyA}\n`;
  draft += `乙方（卖方）：${partyB}\n\n`;
  draft += `根据《中华人民共和国民法典》及相关法律法规，甲乙双方在平等、自愿、公平、诚实信用的原则基础上，经友好协商，就${tpl.description}事宜达成如下协议：\n\n`;

  const clauses = tpl.requiredClauses || [];
  clauses.forEach((clause, i) => {
    draft += `第${["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"][i] || i + 1}条 ${clause}\n`;
    if (clause.includes("主体")) {
      draft += `1.1 甲方：${partyA}\n1.2 乙方：${partyB}\n\n`;
    } else if (clause.includes("标的") || clause.includes("采购")) {
      draft += itemsTable ? `${itemsTable}\n` : `（待补充合同标的详情）\n\n`;
    } else if (clause.includes("付款")) {
      draft += `${answers["付款方式"] || "（待确认付款方式）"}\n\n`;
    } else if (clause.includes("违约")) {
      draft += `任何一方违反本合同约定的，应向守约方支付合同总金额【待确认比例】%的违约金，并赔偿由此造成的全部损失。\n\n`;
    } else {
      draft += `（${clause}具体条款待完善）\n\n`;
    }
  });

  draft += `本合同一式肆份，甲乙双方各执贰份，具有同等法律效力。\n\n`;
  draft += `甲方（盖章）：                    乙方（盖章）：\n`;
  draft += `法定代表人：                      法定代表人：\n`;
  draft += `日期：${dateStr}                    日期：${dateStr}\n`;

  return draft;
}

// ── API Endpoints ──

// POST /api/smart-draft/init — Initialize a smart draft session
async function handleSmartDraftInit(req, res, body) {
  let description = "";
  let fileTexts = [];
  let knowledgeMode = "industry";

  try {
    const data = JSON.parse(body);
    description = data.description || "";
    knowledgeMode = data.knowledgeMode || "industry";
    fileTexts = data.fileTexts || []; // [{name, text}]
  } catch { return sendJson(res, 400, { error: "Invalid JSON" }); }

  const session = createSmartDraftSession();
  session.description = description;
  session.knowledgeMode = knowledgeMode;
  session.step = "intent";

  // Agent 1: Intent analysis
  const fileNames = fileTexts.map((f) => f.name);
  session.intent = await agentIntentAnalyze(description, fileNames);

  // Match builtin template
  session.template = builtinTemplates.find((t) => t.type === session.intent.contractType) || builtinTemplates[0];

  // Agent 2: Data extraction
  session.step = "extract";
  session.extractedData = await agentDataExtract(fileTexts, session.intent);

  // Merge extracted fields into answers
  if (session.extractedData && session.extractedData.fields) {
    Object.assign(session.answers, session.extractedData.fields);
  }

  // Agent 3: Knowledge retrieval
  session.step = "knowledge";
  const kb = agentKnowledgeRetrieve(session.intent, session.knowledgeMode);
  session.rules = kb.rules;
  session.knowledgeSnippets = kb.snippets;

  // Pre-fill answers from memory and extracted data
  const memStore = loadStore();
  const memory = (memStore.contractMemory || []).slice(-10);
  if (session.intent.parties) {
    if (session.intent.parties[0]) session.answers["甲方"] = session.intent.parties[0];
    if (session.intent.parties[1]) session.answers["乙方"] = session.intent.parties[1];
  }

  // Agent 5: Determine missing fields
  const dialogueResult = agentDialogueManage(session);

  session.step = "iterate";
  session.messages.push({
    role: "assistant",
    content: `已识别为「${session.intent.contractTypeCn}」（${session.intent.industry}）\n` +
      `模板: ${session.template.name}\n` +
      (session.extractedData.items.length > 0 ? `已提取 ${session.extractedData.items.length} 条数据\n` : "") +
      `匹配到 ${session.rules.length} 条适用规则\n` +
      (dialogueResult.missingCount > 0
        ? `需要补充 ${dialogueResult.missingCount} 项信息，第一个问题：\n${dialogueResult.currentQuestion?.question || "请描述合同详情"}`
        : "信息齐全，可以直接生成草稿"),
  });

  return sendJson(res, 200, {
    sessionId: session.id,
    intent: session.intent,
    template: { id: session.template.id, name: session.template.name },
    extractedData: session.extractedData,
    rulesCount: session.rules.length,
    dialogue: dialogueResult,
    messages: session.messages,
    answers: session.answers,
  });
}

// POST /api/smart-draft/generate — Generate draft
async function handleSmartDraftGenerate(req, res, body) {
  let sessionId = "";
  try {
    const data = JSON.parse(body);
    sessionId = data.sessionId || "";
    if (data.guidanceMode) {
      const session = smartDraftSessions.get(sessionId);
      if (session) session.guidanceMode = data.guidanceMode;
    }
  } catch { return sendJson(res, 400, { error: "Invalid JSON" }); }

  const session = smartDraftSessions.get(sessionId);
  if (!session) return sendJson(res, 404, { error: "Session not found" });

  session.step = "generate";

  // Agent 4: Generate draft
  session.draft = await agentDraftGenerate(session);
  // Normalize draft to string
  if (session.draft && typeof session.draft === "object") {
    session.draft = session.draft.content || session.draft.draft || JSON.stringify(session.draft);
  }

  session.step = "done";
  session.messages.push({ role: "assistant", content: "草稿已生成，请在左侧编辑器中查看和修改。" });

  // Record to memory
  const memStore = loadStore();
  if (!memStore.contractMemory) memStore.contractMemory = [];
  memStore.contractMemory.push({ ts: Date.now(), type: session.intent?.contractType, answers: session.answers });
  if (memStore.contractMemory.length > 500) memStore.contractMemory = memStore.contractMemory.slice(-500);
  saveStore(memStore);

  return sendJson(res, 200, {
    draft: session.draft,
    sessionId: session.id,
    messages: session.messages,
    appliedRules: session.rules.slice(0, 5).map((r) => r.ruleName),
  });
}

// POST /api/smart-draft/answer — User answers a question
async function handleSmartDraftAnswer(req, res, body) {
  let sessionId = "", field = "", value = "";
  try {
    const data = JSON.parse(body);
    sessionId = data.sessionId || "";
    field = data.field || "";
    value = data.value || "";
  } catch { return sendJson(res, 400, { error: "Invalid JSON" }); }

  const session = smartDraftSessions.get(sessionId);
  if (!session) return sendJson(res, 404, { error: "Session not found" });

  session.answers[field] = value;
  session.messages.push({ role: "user", content: `${field}: ${value}` });

  // Re-evaluate missing fields
  const dialogueResult = agentDialogueManage(session);

  if (dialogueResult.missingCount === 0) {
    session.messages.push({ role: "assistant", content: "所有必要信息已收集完毕！可以生成最终草稿。" });
  } else {
    session.messages.push({
      role: "assistant",
      content: `还需要 ${dialogueResult.missingCount} 项信息。\n${dialogueResult.currentQuestion?.question || "请继续补充"}`,
    });
  }

  return sendJson(res, 200, { dialogue: dialogueResult, messages: session.messages, answers: session.answers });
}

// POST /api/smart-draft/guide — Three guidance modes
async function handleSmartDraftGuide(req, res, body) {
  let sessionId = "", mode = "";
  try {
    const data = JSON.parse(body);
    sessionId = data.sessionId || "";
    mode = data.mode || "llm_infer"; // ask_user | kb_search | llm_infer
  } catch { return sendJson(res, 400, { error: "Invalid JSON" }); }

  const session = smartDraftSessions.get(sessionId);
  if (!session) return sendJson(res, 404, { error: "Session not found" });

  session.guidanceMode = mode;
  const question = session.currentQuestion;

  if (mode === "kb_search" && question) {
    // Search knowledge base for relevant content
    const snippets = selectKnowledgeSnippetsForGeneration(loadStore(), question.clause + " " + question.hint, 5);
    const ruleMatches = (session.rules || []).filter((r) =>
      (r.action || "").includes(question.hint) || (r.ruleName || "").includes(question.hint)
    ).slice(0, 3);

    session.messages.push({
      role: "assistant",
      content: `📚 知识库中关于"${question.hint}"的内容：\n` +
        (snippets.length > 0 ? snippets.map((s) => `• ${s}`).join("\n") : "未找到相关内容") +
        (ruleMatches.length > 0 ? "\n\n相关规则：\n" + ruleMatches.map((r) => `• ${r.ruleName}: ${(r.action || "").slice(0, 120)}`).join("\n") : ""),
    });
    return sendJson(res, 200, { snippets, rules: ruleMatches, messages: session.messages });
  }

  if (mode === "llm_infer" && question) {
    // Use LLM to infer answer
    const prompt = `根据以下合同信息，为"${question.clause}"推断一个合理的答案。
合同类型: ${session.template?.name}
已有信息: ${JSON.stringify(session.answers)}
缺失信息: ${question.clause}
附件数据: ${JSON.stringify(session.extractedData?.items?.slice(0, 3) || [])}

请只返回推断的答案值（简短，不超过50字），不要解释：`;

    try {
      const inferred = await callChatModel(prompt, null, { temperature: 0.3, maxTokens: 200 });
      if (inferred) {
        const cleanAnswer = inferred.replace(/^["']|["']$/g, "").trim();
        session.answers[question.hint] = cleanAnswer;
        session.messages.push({ role: "user", content: `${question.hint}: ${cleanAnswer} (AI推断)` });

        const dialogueResult = agentDialogueManage(session);
        session.messages.push({
          role: "assistant",
          content: dialogueResult.missingCount > 0
            ? `AI已推断"${question.hint}"为"${cleanAnswer}"。\n还需补充 ${dialogueResult.missingCount} 项。\n${dialogueResult.currentQuestion?.question || "请继续"}`
            : `AI已推断所有缺失信息。可以生成最终草稿。`,
        });
        return sendJson(res, 200, { inferred: cleanAnswer, dialogue: dialogueResult, messages: session.messages, answers: session.answers });
      }
    } catch (e) {
      console.error("[Agent5] LLM infer failed:", e.message);
    }
  }

  return sendJson(res, 200, { messages: session.messages });
}

// POST /api/smart-draft/extract-file — Extract data from a single file
async function handleSmartDraftExtractFile(req, res, body) {
  let sessionId = "", fileName = "", fileText = "";
  try {
    const data = JSON.parse(body);
    sessionId = data.sessionId || "";
    fileName = data.fileName || "";
    fileText = data.fileText || "";
  } catch { return sendJson(res, 400, { error: "Invalid JSON" }); }

  const session = smartDraftSessions.get(sessionId);
  if (!session) return sendJson(res, 404, { error: "Session not found" });

  const result = await agentDataExtract([{ name: fileName, text: fileText }], session.intent || { contractTypeCn: "合同" });

  return sendJson(res, 200, result);
}

// GET /api/smart-draft/session — Get session state
async function handleSmartDraftSession(req, res, sessionId) {
  const session = smartDraftSessions.get(sessionId);
  if (!session) return sendJson(res, 404, { error: "Session not found" });
  return sendJson(res, 200, {
    id: session.id,
    step: session.step,
    intent: session.intent,
    template: session.template ? { id: session.template.id, name: session.template.name } : null,
    extractedData: session.extractedData,
    answers: session.answers,
    draft: session.draft,
    messages: session.messages,
    knowledgeMode: session.knowledgeMode,
    guidanceMode: session.guidanceMode,
    rulesCount: (session.rules || []).length,
  });
}

// GET /api/smart-draft/templates — List builtin templates
async function handleSmartDraftTemplates(req, res) {
  const templates = builtinTemplates.map((t) => ({
    id: t.id, name: t.name, type: t.type, category: t.category,
    description: t.description, keywords: t.keywords,
    requiredClausesCount: t.requiredClauses.length,
  }));
  return sendJson(res, 200, templates);
}

async function parseTemplateImportInput(req, body) {
  const contentType = req.headers["content-type"] || "";
  let contractText = "";
  let fileName = "";
  if (contentType.includes("multipart/form-data")) {
    const { fields, files } = parseMultipart(body, contentType);
    contractText = fields.templateText || fields.contractText || "";
    if (files[0]) {
      const id = `TPL_${Date.now()}_${crypto.randomBytes(3).toString("hex")}`;
      fileName = files[0].filename;
      const filePath = path.join(UPLOAD_DIR, `${id}_${fileName.replace(/[^\w.\-\u4e00-\u9fa5]/g, "_")}`);
      fs.writeFileSync(filePath, files[0].buffer);
      contractText = `${contractText}\n${await extractText(filePath, fileName, files[0].buffer)}`.trim();
    }
  } else {
    const payload = JSON.parse(body.toString("utf8") || "{}");
    contractText = payload.templateText || payload.contractText || "";
    fileName = payload.fileName || "";
  }
  return { contractText: String(contractText || "").trim(), fileName };
}

async function handleImportContractTemplate(req, res) {
  const body = await readBody(req);
  const { contractText, fileName } = await parseTemplateImportInput(req, body);
  if (!contractText) return sendJson(res, 400, { error: "请上传合同文件或粘贴合同正文" });
  const store = loadStore();
  let template;
  let usedAI = false;
  let fallbackReason = "";
  try {
    const modelResult = await callJsonModel(buildTemplateImportPrompt(contractText, fileName), {
      repair: true,
      label: "合同模板抽取结果",
    });
    template = normalizeImportedTemplate(modelResult.data, contractText, fileName);
    usedAI = true;
  } catch (error) {
    template = fallbackImportedTemplate(contractText, fileName);
    fallbackReason = error.message;
  }
  if (!template?.templateText) return sendJson(res, 400, { error: "未能从合同中抽取模板正文" });
  store.customTemplates = [
    template,
    ...store.customTemplates.filter((item) => item.id !== template.id && item.name !== template.name),
  ].slice(0, 80);
  saveStore(store);
  sendJson(res, 201, {
    template: publicTemplate(template),
    usedAI,
    usedFallback: !usedAI,
    fallbackReason,
    templateCount: store.customTemplates.length,
  });
}

async function handleDeleteContractTemplate(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const id = String(payload.templateId || payload.id || "").trim();
  const name = String(payload.name || "").trim();
  if (!id && !name) return sendJson(res, 400, { error: "请提供模板 ID 或名称" });
  const store = loadStore();
  const before = store.customTemplates.length;
  store.customTemplates = store.customTemplates.filter((template) => {
    if (id && template.id === id) return false;
    if (name && template.name === name) return false;
    return true;
  });
  saveStore(store);
  sendJson(res, 200, { deleted: before - store.customTemplates.length });
}

async function parseReviewContractInput(req, body) {
  const contentType = req.headers["content-type"] || "";
  let contractText = "";
  let contractHtml = "";
  let fileName = "";
  let templateDocxId = "";
  if (contentType.includes("multipart/form-data")) {
    const { fields, files } = parseMultipart(body, contentType);
    contractText = fields.contractText || "";
    if (files[0]) {
      const id = `REVIEW_${Date.now()}_${crypto.randomBytes(3).toString("hex")}`;
      fileName = files[0].filename;
      const filePath = path.join(UPLOAD_DIR, `${id}_${fileName.replace(/[^\w.\-\u4e00-\u9fa5]/g, "_")}`);
      fs.writeFileSync(filePath, files[0].buffer);
      const ext = path.extname(fileName).toLowerCase();
      let formattedResult = null;
      if (ext === ".docx") {
        try {
          const parser = new DocxFormatParser(files[0].buffer);
          formattedResult = await parser.parse();
        } catch (e) { console.error("[DocxFormatParser] docx parse error:", e.message); formattedResult = null; }
      } else if (ext === ".doc") {
        try {
          const docParser = new DocFormatParser(files[0].buffer);
          const parsed = await docParser.parse();
          formattedResult = { html: parsed.html || "", text: parsed.text || "" };
        } catch (e) { console.error("[DocFormatParser] doc parse error:", e.message); formattedResult = null; }
      }
      if (formattedResult) {
        contractText = `${contractText}\n${formattedResult.text}`.trim();
        contractHtml = formattedResult.html;
        templateDocxId = id;
      } else {
        contractText = `${contractText}\n${await extractText(filePath, fileName, files[0].buffer)}`.trim();
        if (ext === ".docx") {
          templateDocxId = id;
          contractHtml = extractDocxHtml(files[0].buffer);
        }
      }
    }
  } else {
    const payload = JSON.parse(body.toString("utf8") || "{}");
    contractText = payload.contractText || "";
  }
  return { contractText: String(contractText || "").trim(), contractHtml, fileName, templateDocxId };
}

async function handleParseContract(req, res) {
  const body = await readBody(req);
  const { contractText, contractHtml, fileName, templateDocxId } = await parseReviewContractInput(req, body);
  if (!contractText) return sendJson(res, 400, { error: "请上传合同文件或粘贴合同文本" });
  return sendJson(res, 200, {
    fileName,
    contractText: contractText.slice(0, 60000),
    contractHtml: contractHtml ? contractHtml.slice(0, 500000) : "",
    templateDocxId,
    templatePreserved: Boolean(templateDocxId),
    charCount: contractText.length,
  });
}

async function handleParseFormatted(req, res) {
  const body = await readBody(req);
  const contentType = req.headers["content-type"] || "";
  let buffer = null;
  let fileName = "";
  if (contentType.includes("multipart/form-data")) {
    const { files } = parseMultipart(body, contentType);
    if (files[0]) {
      buffer = files[0].buffer;
      fileName = files[0].filename;
    }
  } else {
    return sendJson(res, 400, { error: "请上传 doc 或 docx 文件" });
  }
  if (!buffer) return sendJson(res, 400, { error: "未收到文件" });
  const ext = path.extname(fileName).toLowerCase();
  if (ext !== ".docx" && ext !== ".doc") {
    return sendJson(res, 400, { error: "仅支持 .doc 和 .docx 格式" });
  }
  try {
    let docxBuffer = buffer;
    if (ext === ".doc") {
      docxBuffer = await convertDocToDocx(buffer);
    }
    const parser = new DocxFormatParser(docxBuffer);
    const result = await parser.parse();
    return sendJson(res, 200, {
      fileName,
      html: result.html.slice(0, 500000),
      text: result.text.slice(0, 60000),
      charCount: result.text.length,
      formatPreserved: true,
    });
  } catch (error) {
    return sendJson(res, 500, { error: `文档解析失败: ${error.message}` });
  }
}

function convertDocToDocx(buffer) {
  return new Promise((resolve, reject) => {
    const tmpDir = path.join(ROOT, "tmp");
    fs.mkdirSync(tmpDir, { recursive: true });
    const baseName = `conv_${Date.now()}_${crypto.randomBytes(4).toString("hex")}`;
    const docPath = path.join(tmpDir, `${baseName}.doc`);
    const docxPath = path.join(tmpDir, `${baseName}.docx`);
    fs.writeFileSync(docPath, buffer);
    const candidates = [
      "libreoffice",
      "soffice",
      process.env.LIBREOFFICE_PATH || "",
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ].filter(Boolean);
    const bin = candidates.find((c) => {
      try { fs.accessSync(c); return true; } catch { return false; }
    });
    if (!bin) {
      try { fs.unlinkSync(docPath); } catch {}
      reject(new Error("服务器未安装 LibreOffice，无法解析 .doc 格式文件，请转换为 .docx 后重新上传"));
      return;
    }
    execFile(bin, ["--headless", "--convert-to", "docx", "--outdir", tmpDir, docPath], { timeout: 30000 }, (err) => {
      try { fs.unlinkSync(docPath); } catch {}
      if (err || !fs.existsSync(docxPath)) {
        reject(new Error("DOC 转换失败，请手动转换为 .docx 后重新上传"));
        return;
      }
      const result = fs.readFileSync(docxPath);
      try { fs.unlinkSync(docxPath); } catch {}
      resolve(result);
    });
  });
}

function findUploadedDocxPath(templateDocxId = "") {
  const id = String(templateDocxId || "").replace(/[^\w-]/g, "");
  if (!id || !id.startsWith("REVIEW_")) return "";
  const match = fs.readdirSync(UPLOAD_DIR).find((name) => name.startsWith(`${id}_`) && name.toLowerCase().endsWith(".docx"));
  return match ? path.join(UPLOAD_DIR, match) : "";
}

async function handleExportPatchedDocx(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const filePath = findUploadedDocxPath(payload.templateDocxId);
  if (!filePath) return sendJson(res, 404, { error: "未找到原始 DOCX 模板，无法保真导出" });
  const currentText = String(payload.currentText || payload.contractText || "").trim();
  if (!currentText) return sendJson(res, 400, { error: "当前文档内容为空，无法导出" });
  const original = fs.readFileSync(filePath);
  const patched = patchDocxBuffer(original, currentText);
  const baseName = path.basename(String(payload.fileName || path.basename(filePath)).replace(/[\\/:*?"<>|]/g, "_"), ".docx") || "合同当前版本";
  const downloadName = `${baseName}_修改版.docx`;
  res.writeHead(200, {
    "content-type": contentTypes[".docx"],
    "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(downloadName)}`,
    "cache-control": "no-store",
  });
  res.end(patched);
}

async function handleReviewContract(req, res) {
  const body = await readBody(req);
  const { contractText, fileName } = await parseReviewContractInput(req, body);
  if (!contractText) return sendJson(res, 400, { error: "请上传合同文件或粘贴合同文本" });

  const store = loadStore();
  const ruleProfile = inferContractProfile(contractText);
  const rules = selectRulesForContract(store, ruleProfile);
  const logicIssues = buildLogicReviewIssues(contractText);
  const allRuleIssues = buildRuleReviewIssues(contractText, rules);
  const visibleRuleIssues = buildRuleReviewIssues(contractText, enterpriseRulesOnly(rules));
  const precheckIssues = mergeReviewIssues(allRuleIssues, logicIssues, contractText);
  const localIssues = mergeReviewIssues(visibleRuleIssues, logicIssues, contractText);
  try {
    const modelReview = await callJsonModel(buildReviewPrompt(contractText, rules, precheckIssues, ruleProfile), {
      repair: true,
      label: "合同审查结果",
    });
    const result = modelReview.data && typeof modelReview.data === "object" ? modelReview.data : {};
    const modelIssues = Array.isArray(result.issues) ? result.issues : [];
    const issues = applyReviewVisibility(mergeReviewIssues(modelIssues, localIssues, contractText), rules, contractText);
    const overallRisk = issues.some((item) => item.risk_level === "高") ? "高" : result.overall_risk || "中";
    return sendJson(res, 200, {
      fileName,
      contractText: contractText.slice(0, 60000),
      overallRisk,
      summary: result.summary || "审查完成。",
      reviewReport: buildReviewReport(issues, contractText),
      reviewWorkflow: buildReviewWorkflow(contractText, rules, ruleProfile, issues, { engine: "大模型深度审查 + 本地规则复核" }),
      issues,
      usedAI: true,
      usedFallback: false,
      reviewEngine: "大模型深度审查 + 本地规则复核",
      model: modelReview.model || activeModelInfo().model,
      modelProvider: modelReview.provider || activeModelInfo().provider,
      modelIssueCount: modelIssues.length,
      localIssueCount: localIssues.length,
      presetRuleIssueCount: allRuleIssues.length,
      totalIssueCount: issues.length,
      ruleProfile,
      ruleSelection: summarizeRuleSelection(ruleProfile, rules),
      aiJsonRepaired: modelReview.repaired,
      aiParseRepairReason: modelReview.parseError || "",
    });
  } catch (error) {
    const fallback = fallbackReview(contractText, rules);
    const modelInfo = activeModelInfo();
    return sendJson(res, 200, {
      fileName,
      contractText: contractText.slice(0, 60000),
      ...fallback,
      overallRisk: fallback.overall_risk,
      reviewWorkflow: buildReviewWorkflow(contractText, rules, ruleProfile, fallback.issues || [], { engine: "本地临时审核（平台预设规则库）" }),
      usedAI: false,
      usedFallback: true,
      reviewEngine: "本地临时审核（平台预设规则库）",
      model: modelInfo.model,
      modelProvider: modelInfo.provider,
      modelIssueCount: 0,
      localIssueCount: fallback.issues?.length || localIssues.length,
      presetRuleIssueCount: fallback.presetRuleIssueCount || allRuleIssues.length,
      totalIssueCount: fallback.issues?.length || 0,
      ruleProfile,
      ruleSelection: summarizeRuleSelection(ruleProfile, rules),
      fallbackReason: error.message,
    });
  }
}

async function handleUpload(req, res) {
  const body = await readBody(req);
  const { fields, files } = parseMultipart(body, req.headers["content-type"]);
  if (!files.length) return sendJson(res, 400, { error: "No files uploaded" });

  const store = loadStore();
  const created = [];
  for (const file of files) {
    const id = `DOC_${Date.now()}_${crypto.randomBytes(4).toString("hex")}`;
    const safeName = file.filename.replace(/[^\w.\-\u4e00-\u9fa5]/g, "_");
    const filePath = path.join(UPLOAD_DIR, `${id}_${safeName}`);
    fs.writeFileSync(filePath, file.buffer);
    let text = "";
    let parseStatus = "parsed";
    let parseMessage = "";
    try {
      text = await extractText(filePath, file.filename, file.buffer);
      if (!text.trim()) {
        parseStatus = "partial";
        parseMessage = "未能提取到可读文本，可尝试上传 Word、TXT 或可复制文本的 PDF。";
      }
    } catch (error) {
      parseStatus = "failed";
      parseMessage = error.message;
    }
    const doc = {
      id,
      name: file.filename,
      filePath,
      contentType: file.contentType,
      size: file.buffer.length,
      docType: fields.docType || simpleDocType(file.filename, text),
      domain: fields.domain || "",
      contractType: fields.contractType || "",
      confidentiality: fields.confidentiality || "内部",
      parseStatus,
      parseMessage,
      summary: summarize(text),
      chunkCount: chunkText(text).length,
      text,
      createdAt: new Date().toISOString(),
    };
    store.documents.unshift(doc);
    created.push({ ...doc, text: undefined });
  }
  saveStore(store);
  sendJson(res, 201, { documents: created });
}

async function handleUpdateDocument(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const allowedFields = new Set(["docType", "domain", "contractType", "confidentiality"]);
  if (!allowedFields.has(payload.field)) return sendJson(res, 400, { error: "Unsupported document field" });
  const store = loadStore();
  const doc = store.documents.find((item) => item.id === payload.documentId);
  if (!doc) return sendJson(res, 404, { error: "Document not found" });
  doc[payload.field] = String(payload.value || "").trim();
  doc.updatedAt = new Date().toISOString();
  saveStore(store);
  sendJson(res, 200, { document: publicDocument(doc) });
}

async function handleExtractRules(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const store = loadStore();
  const selectedIds = Array.isArray(payload.documentIds) ? payload.documentIds : [];
  const documents = store.documents.filter((doc) => selectedIds.length === 0 || selectedIds.includes(doc.id));
  if (!documents.length) return sendJson(res, 400, { error: "No documents selected" });

  const dimensions = loadDimensions();
  let modelRules;
  let usedFallback = false;
  let modelCall = null;
  try {
    modelCall = await callChatModel(buildPrompt(documents, dimensions));
    modelRules = safeJsonFromModel(modelCall.content);
  } catch (error) {
    if (payload.allowFallback === false) return sendJson(res, 500, { error: error.message });
    usedFallback = true;
    modelRules = fallbackRules(documents, dimensions);
  }
  const list = Array.isArray(modelRules) ? modelRules : modelRules.rules || [];
  const docLookup = new Map(documents.map((doc) => [doc.id, doc]));
  const normalized = list.map((rule) => normalizeRule(rule, docLookup));
  const { inserted, skipped } = addUniqueRules(store, normalized);
  saveStore(store);
  sendJson(res, 201, {
    rules: inserted,
    skipped: skipped.length,
    usedFallback,
    fallbackReason: usedFallback ? "AI 调用失败，已使用本地关键词规则兜底。请检查 OpenAI/DashScope Key、模型或网络。" : "",
    model: modelCall?.model || activeModelInfo().model,
    modelProvider: modelCall?.provider || activeModelInfo().provider,
  });
}

async function handleReview(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const store = loadStore();
  const rule = store.rules.find((item) => item.id === payload.ruleId);
  if (!rule) return sendJson(res, 404, { error: "Rule not found" });
  rule.reviewStatus = payload.status || rule.reviewStatus;
  rule.reviewComment = payload.comment || "";
  rule.reviewedAt = new Date().toISOString();
  saveStore(store);
  sendJson(res, 200, { rule: publicRule(rule) });
}

async function handleDeleteRule(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const store = loadStore();
  const before = store.rules.length;
  store.rules = store.rules.filter((item) => item.id !== payload.ruleId);
  saveStore(store);
  sendJson(res, 200, { deleted: before - store.rules.length });
}

async function handleUpdateRule(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const store = loadStore();
  const rule = store.rules.find((item) => item.id === payload.ruleId);
  if (!rule) return sendJson(res, 404, { error: "Rule not found" });
  const updates = payload.updates || {};
  const allowed = [
    "ruleName",
    "dimension",
    "ruleType",
    "scenario",
    "contractType",
    "businessDomain",
    "triggerCondition",
    "action",
    "riskLevel",
    "priority",
    "sourceQuote",
    "reviewStatus",
    "ruleBasis",
  ];
  for (const field of allowed) {
    if (!(field in updates)) continue;
    if (field === "contractType") {
      rule.contractType = Array.isArray(updates.contractType) ? updates.contractType.filter(Boolean) : [String(updates.contractType || "通用合同")];
    } else if (field === "priority") {
      rule.priority = Number(updates.priority || rule.priority || 50);
    } else if (field === "ruleType") {
      rule.ruleType = updates.ruleType === "生成约束规则" ? "通用规则" : String(updates.ruleType || "通用规则");
    } else {
      rule[field] = String(updates[field] || "").trim();
    }
  }
  rule.useScope = ["生成", "审查"];
  rule.updatedAt = new Date().toISOString();
  saveStore(store);
  sendJson(res, 200, { rule: publicRule(rule) });
}

async function handleCreateRule(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const store = loadStore();
  const rawText = String(payload.ruleText || payload.action || payload.ruleName || "").trim();
  if (!rawText) return sendJson(res, 400, { error: "自定义规则内容不能为空" });
  const aiClassified = payload.classifyMode === "ai";
  const scenario = aiClassified ? inferScenario(rawText, payload.scope || "") : payload.scenario || inferScenario(rawText);
  const contractType = payload.contractType?.length ? payload.contractType : (payload.scope ? [payload.scope] : ["通用合同"]);
  const rule = normalizeRule({
    ...payload,
    rule_name: payload.ruleName || rawText.slice(0, 40),
    dimension: payload.dimension || "通用必备条款",
    rule_type: payload.ruleType || (scenario === "审批规则" ? "审批规则" : "通用规则"),
    scenario,
    contract_type: contractType,
    business_domain: payload.businessDomain || "通用",
    trigger_condition: payload.triggerCondition || "合同生成或审查时",
    action: rawText,
    source_quote: payload.sourceQuote || "公司自定义规则",
    source_doc_name: "公司自己的规则",
    review_status: payload.reviewStatus || "active",
    rule_source: "公司规则",
    use_scope: ["生成", "审查"],
    rule_basis: "企业自定",
  });
  const { inserted, skipped } = addUniqueRules(store, [rule]);
  saveStore(store);
  sendJson(res, 201, { rules: inserted.map(publicRule), skipped: skipped.length });
}

function handleExportRules(req, res) {
  const store = loadStore();
  sendJson(res, 200, {
    schema: "contract-rule-library/v1",
    exportedAt: new Date().toISOString(),
    ruleCount: store.rules.length,
    rules: store.rules.map(toReusableRule),
  });
}

async function handleImportRules(req, res) {
  const body = await readBody(req);
  const contentType = req.headers["content-type"] || "";
  let raw = "";
  if (contentType.includes("multipart/form-data")) {
    const { files } = parseMultipart(body, contentType);
    if (!files.length) return sendJson(res, 400, { error: "No rule file uploaded" });
    raw = files[0].buffer.toString("utf8");
  } else {
    raw = body.toString("utf8");
  }
  const payload = JSON.parse(raw || "{}");
  const list = Array.isArray(payload) ? payload : payload.rules || [];
  if (!Array.isArray(list) || !list.length) return sendJson(res, 400, { error: "No rules found in import file" });
  const imported = list.map((rule) => ({
    ...normalizeRule(rule),
    id: `RULE_${Date.now()}_${crypto.randomBytes(3).toString("hex")}`,
    importedAt: new Date().toISOString(),
  }));
  const store = loadStore();
  const { inserted, skipped } = addUniqueRules(store, imported);
  saveStore(store);
  sendJson(res, 201, { rules: inserted, skipped: skipped.length });
}

async function handleAcceptAllRules(req, res) {
  const store = loadStore();
  let count = 0;
  for (const rule of store.rules) {
    if (rule.reviewStatus === "pending_review") {
      rule.reviewStatus = "active";
      rule.reviewedAt = new Date().toISOString();
      rule.reviewComment = "批量接受";
      count += 1;
    }
  }
  saveStore(store);
  sendJson(res, 200, { accepted: count });
}

function publicDocument(doc) {
  const { text, filePath, ...rest } = doc;
  return rest;
}


/* ── 联网搜索 ── */
async function webSearch(query, maxResults = 5) {
  const results = [];
  try {
    const url = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}&kl=cn-zh`;
    const res = await fetch(url, {
      headers: { "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" },
      signal: AbortSignal.timeout(10000),
    });
    const html = await res.text();
    const resultRegex = /<a[^>]+class="result__a"[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?<a[^>]+class="result__snippet"[^>]*>([\s\S]*?)<\/a>/gi;
    let m;
    while ((m = resultRegex.exec(html)) && results.length < maxResults) {
      let link = m[1];
      const uMatch = link.match(/[?&]uddg=([^&]+)/);
      if (uMatch) link = decodeURIComponent(uMatch[1]);
      const title = m[2].replace(/<[^>]+>/g, "").trim();
      const snippet = m[3].replace(/<[^>]+>/g, "").trim();
      if (title && link) results.push({ title, url: link, snippet });
    }
  } catch (_) { /* fall through */ }

  if (!results.length) {
    try {
      const url = `https://api.duckduckgo.com/?q=${encodeURIComponent(query)}&format=json&no_html=1&skip_disambig=1`;
      const res = await fetch(url, { headers: { "user-agent": "Mozilla/5.0" }, signal: AbortSignal.timeout(8000) });
      const data = await res.json();
      if (data.AbstractText) results.push({ title: data.Heading || query, url: data.AbstractURL || "", snippet: data.AbstractText });
      (data.RelatedTopics || []).slice(0, maxResults).forEach((t) => {
        if (t.Text && t.FirstURL) results.push({ title: t.Text.split(" ").slice(0, 8).join(" "), url: t.FirstURL, snippet: t.Text });
      });
    } catch (_) { /* ignore */ }
  }
  return results.slice(0, maxResults);
}

async function handleWebSearch(req, res) {
  const body = JSON.parse((await readBody(req)).toString() || "{}");
  const { query } = body;
  if (!query) return sendJson(res, 400, { error: "缺少搜索关键词" });
  try {
    const results = await webSearch(query, 8);
    return sendJson(res, 200, { query, results });
  } catch (err) {
    return sendJson(res, 500, { error: `搜索失败: ${err.message}` });
  }
}

/* ── 合同对话 ── */
async function handleContractChat(req, res) {
  const body = JSON.parse((await readBody(req)).toString() || "{}");
  const { message, contractText, contractHtml, history, enableWebSearch } = body;
  if (!message) return sendJson(res, 400, { error: "缺少消息内容" });

  const providers = getModelProviders();
  if (!providers.length) return sendJson(res, 500, { error: "未配置 AI API Key" });

  const contractContext = contractText
    ? `\n\n以下是用户导入的合同原文：\n${contractText.slice(0, 120000)}`
    : "\n\n用户尚未导入合同。";

  const hasSelectedText = message.includes("【选中文本】");

  let searchResults = [];
  let searchContext = "";
  if (enableWebSearch) {
    try {
      const cleanQuery = message.replace(/【选中文本】[\s\S]*/g, "").replace(/[，。！？、；：""''（）【】]/g, " ").trim().slice(0, 200);
      const searchQuery = cleanQuery + " 合同 法律";
      searchResults = await webSearch(searchQuery, 5);
      if (searchResults.length) {
        searchContext = `\n\n以下是联网搜索到的相关资料，请参考这些信息回答用户问题：\n${searchResults.map((r, i) => `[${i + 1}] ${r.title}\n    来源: ${r.url}\n    摘要: ${r.snippet}`).join("\n\n")}`;
      }
    } catch (_) { /* search failure is non-fatal */ }
  }

  const systemPrompt = `你是一位专业的合同审查助手。你的职责是帮助用户分析、理解和审查合同内容。
${contractContext}
${searchContext}

你的能力包括：
1. 总结合同要点和关键条款
2. 分析合同中的风险点和潜在问题
3. 检查条款是否完整、合理
4. 提供修改建议
5. 解释法律术语和条款含义
6. 对比行业标准条款
7. 润色和优化合同文字表达

${hasSelectedText ? "用户选中了合同中的一段文字进行提问，请重点分析这段文字并给出针对性的建议。\n当需要给出修改建议时，请使用以下格式标记可替换的文字：\n【原文】需要替换的原文内容\n【修改】建议替换的新内容\n这样用户可以一键应用你的修改建议。" : ""}
当需要给出修改建议时，如果涉及具体的条款替换，也请使用【原文】和【修改】的格式。
${searchContext ? "回答时请引用搜索结果中的相关信息，并注明来源编号如[1][2]等。" : ""}

请用中文回复，语言专业但易懂。如果用户询问的内容在合同中找不到，请明确说明。`;

  const messages = [
    { role: "system", content: systemPrompt },
    ...(Array.isArray(history) ? history : []),
    { role: "user", content: message },
  ];

  const errors = [];
  for (const provider of providers) {
    try {
      const headers = {
        "content-type": "application/json",
        ...(provider.useBearer === false ? {} : { authorization: `Bearer ${provider.apiKey}` }),
        ...(provider.extraHeaders || {}),
      };
      const response = await fetch(`${normalizeModelBaseUrl(provider.baseUrl)}/chat/completions`, {
        method: "POST",
        headers,
        body: JSON.stringify({ model: provider.model, temperature: 0.3, max_tokens: 16384, messages }),
      });
      const responseText = await response.text();
      let data = {};
      try { data = responseText ? JSON.parse(responseText) : {}; } catch { data = {}; }
      if (!response.ok) throw new Error(data?.error?.message || `${provider.name} request failed: ${response.status}`);
      const reply = data.choices?.[0]?.message?.content || "";
      return sendJson(res, 200, { reply, provider: provider.name, model: provider.model, searchResults: searchResults.length ? searchResults : undefined });
    } catch (error) {
      errors.push(`${provider.name}: ${error.message}`);
    }
  }
  return sendJson(res, 500, { error: errors.join(" | ") });
}

/* ── 模板文件查找（修复版）── */
function findTemplateDocxPath(docxId = "") {
  const id = String(docxId || "").replace(/[^\w-]/g, "");
  if (!id) return "";
  // 直接查找 {id}.docx
  const direct = path.join(UPLOAD_DIR, `${id}.docx`);
  if (fs.existsSync(direct)) return direct;
  // 兜底：REVIEW_ 前缀格式
  if (id.startsWith("REVIEW_")) {
    const match = fs.readdirSync(UPLOAD_DIR).find((name) => name.startsWith(`${id}_`) && name.toLowerCase().endsWith(".docx"));
    return match ? path.join(UPLOAD_DIR, match) : "";
  }
  return "";
}

async function handleUploadTemplateFixed(req, res) {
  const body = await readBody(req);
  const contentType = req.headers["content-type"] || "";
  const parts = parseMultipart(body, contentType);
  if (!parts || !parts.files.length) return sendJson(res, 400, { error: "请上传文件" });

  const file = parts.files[0];
  if (!/\.docx?$/i.test(file.filename)) return sendJson(res, 400, { error: "仅支持 .doc/.docx 文件" });

  const docxId = crypto.randomUUID();
  const isDoc = /\.doc$/i.test(file.filename) && !/\.docx$/i.test(file.filename);
  const ext = isDoc ? "doc" : "docx";
  const savePath = path.join(UPLOAD_DIR, `${docxId}.${ext}`);
  if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
  fs.writeFileSync(savePath, file.buffer);

  try {
    let result;
    if (isDoc) {
      // .doc file: use word-extractor to get text, then extract placeholders
      const docParser = new DocFormatParser(file.buffer);
      const parsed = await docParser.parse();
      const text = parsed.text || "";
      // Extract placeholder fields from text
      const PLACEHOLDER_PATTERNS = [
        { regex: /【([^】]{1,60})】/g, type: "bracket" },
        { regex: /\{([^}{]{1,60})\}/g, type: "brace" },
        { regex: /<([^<>{]{1,60})>/g, type: "angle" },
        { regex: /_{4,}/g, type: "blank" },
        { regex: /（(待填|填写|请填写|请输入|请补充|待补充|待定|暂缺)）/g, type: "pending" },
      ];
      const fields = [];
      const seen = new Set();
      for (const pat of PLACEHOLDER_PATTERNS) {
        let m;
        const re = new RegExp(pat.regex.source, pat.regex.flags);
        while ((m = re.exec(text)) !== null) {
          const placeholder = m[0];
          const label = m[1] || placeholder;
          if (!seen.has(placeholder)) {
            seen.add(placeholder);
            fields.push({ placeholder, label, fieldType: guessFieldType(label), importance: "optional", location: "text" });
          }
        }
      }
      result = { fields, paragraphs: [{ text: text.slice(0, 5000) }], tables: [], media: [] };
    } else {
      // .docx file: use JSZip-based parser
      result = await parseTemplate(file.buffer);
    }
    result.docxId = docxId;
    result.fileName = file.filename;
    return sendJson(res, 200, result);
  } catch (err) {
    return sendJson(res, 500, { error: `模板解析失败: ${err.message}` });
  }
}

function guessFieldType(label) {
  const l = (label || "").toLowerCase();
  if (l.includes("甲方") || l.includes("买方") || l.includes("委托")) return "partyA";
  if (l.includes("乙方") || l.includes("卖方") || l.includes("受托")) return "partyB";
  if (l.includes("金额") || l.includes("总价") || l.includes("价格") || l.includes("费用")) return "amount";
  if (l.includes("日期") || l.includes("时间")) return "date";
  if (l.includes("期限") || l.includes("有效期")) return "term";
  if (l.includes("地址") || l.includes("住所")) return "address";
  if (l.includes("电话") || l.includes("联系")) return "contact";
  if (l.includes("名称") || l.includes("公司")) return "name";
  return "text";
}

async function handleParseTemplateFixed(req, res) {
  const body = JSON.parse((await readBody(req)).toString() || "{}");
  let buffer;
  if (body.docxId) {
    const filePath = findTemplateDocxPath(body.docxId);
    if (!filePath) return sendJson(res, 404, { error: "未找到文件" });
    buffer = fs.readFileSync(filePath);
  } else {
    return sendJson(res, 400, { error: "请上传 .docx 模板文件或提供 docxId" });
  }
  try {
    const result = await parseTemplate(buffer);
    return sendJson(res, 200, result);
  } catch (err) {
    return sendJson(res, 500, { error: `模板解析失败: ${err.message}` });
  }
}

async function handleFillTemplateFixed(req, res) {
  const body = JSON.parse((await readBody(req)).toString() || "{}");
  const { docxId, values } = body;
  if (!docxId) return sendJson(res, 400, { error: "缺少 docxId" });
  if (!values || typeof values !== "object") return sendJson(res, 400, { error: "缺少填充数据 values" });

  const filePath = findTemplateDocxPath(docxId);
  if (!filePath) return sendJson(res, 404, { error: "未找到模板文件" });

  try {
    const buffer = fs.readFileSync(filePath);
    const filled = await fillTemplate(buffer, values);
    const downloadName = `filled_${Date.now()}.docx`;
    res.writeHead(200, {
      "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(downloadName)}`,
      "cache-control": "no-store",
    });
    res.end(filled);
  } catch (err) {
    return sendJson(res, 500, { error: `填充失败: ${err.message}` });
  }
}

async function handleServeDocxFixed(req, res) {
  const id = req.url.split("/api/documents/file/")[1]?.split("?")[0];
  if (!id) return sendJson(res, 400, { error: "Missing file id" });
  const filePath = findTemplateDocxPath(id);
  if (!filePath || !fs.existsSync(filePath)) return sendJson(res, 404, { error: "File not found" });
  res.writeHead(200, {
    "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "cache-control": "public, max-age=3600",
    "access-control-allow-origin": "*",
  });
  fs.createReadStream(filePath).pipe(res);
}
async function handleAIAnalyzeTemplate(req, res) {
  const body = JSON.parse((await readBody(req)).toString() || "{}");
  const { fields, context, attachments } = body;
  if (!fields || !fields.length) return sendJson(res, 400, { error: "缺少字段列表" });

  const providers = getModelProviders();
  if (!providers.length) return sendJson(res, 500, { error: "未配置 AI API Key" });

  const fieldsDesc = fields.map((f, i) => `${i + 1}. ${f.placeholder} (标签: ${f.label}, 类型: ${f.fieldType}, 位置: ${f.location})`).join("\n");
  const attachDesc = attachments ? `\n\n附件内容：\n${JSON.stringify(attachments).slice(0, 4000)}` : "";

  const prompt = `你是一位专业的合同助手。用户正在填写一份合同模板，模板中包含以下占位符字段：

${fieldsDesc}
${context ? `\n合同背景：${context}` : ""}
${attachDesc}

请完成以下任务：
1. 从附件内容中提取能匹配到模板字段的信息，生成一个 JSON 对象 { "【字段名】": "值" }
2. 列出无法从附件获取的字段，以及你需要追问用户的问题
3. 对关键字段（如金额、日期、违约条款）提供行业标准参考建议

返回格式（JSON）：
{
  "filledValues": { "【甲方名称】": "xxx", ... },
  "filledCount": 3,
  "missingQuestions": [
    { "field": "【合同金额】", "question": "合同总价是多少？", "importance": "required" }
  ],
  "aiSuggestions": [
    { "field": "【违约金比例】", "suggestion": "建议设为合同总额的5%-10%", "reason": "行业惯例" }
  ]
}`;

  const messages = [{ role: "user", content: prompt }];
  const errors = [];
  for (const provider of providers) {
    try {
      const headers = {
        "content-type": "application/json",
        ...(provider.useBearer === false ? {} : { authorization: `Bearer ${provider.apiKey}` }),
        ...(provider.extraHeaders || {}),
      };
      const response = await fetch(`${normalizeModelBaseUrl(provider.baseUrl)}/chat/completions`, {
        method: "POST",
        headers,
        body: JSON.stringify({ model: provider.model, temperature: 0.3, max_tokens: 4096, messages }),
      });
      const responseText = await response.text();
      let data = {};
      try { data = responseText ? JSON.parse(responseText) : {}; } catch { data = {}; }
      if (!response.ok) throw new Error(data?.error?.message || `${provider.name} failed: ${response.status}`);
      const reply = data.choices?.[0]?.message?.content || "";
      let analysis = {};
      try {
        const jsonMatch = reply.match(/\{[\s\S]*\}/);
        if (jsonMatch) analysis = JSON.parse(jsonMatch[0]);
      } catch (_) {}
      return sendJson(res, 200, { analysis, rawReply: reply, provider: provider.name, model: provider.model });
    } catch (error) {
      errors.push(`${provider.name}: ${error.message}`);
    }
  }
  return sendJson(res, 500, { error: errors.join(" | ") });
}
async function handleApi(req, res, pathname) {
  try {
    if (req.method === "GET" && pathname === "/api/health") {
      const modelInfo = activeModelInfo();
      return sendJson(res, 200, {
        ok: true,
        hasApiKey: modelInfo.hasApiKey,
        model: modelInfo.model,
        modelProvider: modelInfo.provider,
        hasXiaomi: Boolean(XIAOMI_API_KEY),
        hasOpenAI: Boolean(OPENAI_API_KEY),
        hasDashScope: Boolean(DASHSCOPE_API_KEY),
        providerPriority: AI_PROVIDER_PRIORITY,
        storage: storeBackend,
        hasDatabase: Boolean(pgPool),
      });
    }
    if (req.method === "GET" && pathname === "/api/dimensions") return sendJson(res, 200, loadDimensions());
    if (req.method === "GET" && pathname === "/api/contract-templates") {
      const store = loadStore();
      const customTemplates = (store.customTemplates || []).map(normalizeStoredTemplate).filter(Boolean);
      const seen = new Set();
      const templates = [...CONTRACT_TEMPLATES, ...customTemplates].filter((template) => {
        if (!template?.id || seen.has(template.id)) return false;
        seen.add(template.id);
        return true;
      });
      return sendJson(res, 200, { templates: templates.map(publicTemplate) });
    }
    if (req.method === "GET" && pathname === "/api/documents") {
      const store = loadStore();
      return sendJson(res, 200, { documents: store.documents.map(publicDocument) });
    }
    if (req.method === "GET" && pathname === "/api/rules") {
      const store = loadStore();
      return sendJson(res, 200, { rules: store.rules.map(publicRule) });
    }
    if (req.method === "GET" && pathname === "/api/rules/export") return handleExportRules(req, res);
    if (req.method === "POST" && pathname === "/api/documents") return handleUpload(req, res);
    if (req.method === "POST" && pathname === "/api/documents/update") return handleUpdateDocument(req, res);
    if (req.method === "POST" && pathname === "/api/extract-rules") return handleExtractRules(req, res);
    if (req.method === "POST" && pathname === "/api/rules/review") return handleReview(req, res);
    if (req.method === "POST" && pathname === "/api/rules/update") return handleUpdateRule(req, res);
    if (req.method === "POST" && pathname === "/api/rules/delete") return handleDeleteRule(req, res);
    if (req.method === "POST" && pathname === "/api/rules/custom") return handleCreateRule(req, res);
    if (req.method === "POST" && pathname === "/api/rules/import") return handleImportRules(req, res);
    if (req.method === "POST" && pathname === "/api/rules/accept-all") return handleAcceptAllRules(req, res);
    if (req.method === "POST" && pathname === "/api/contracts/validate-answer") return handleValidateContractAnswer(req, res);
    if (req.method === "POST" && pathname === "/api/contracts/generate") return handleGenerateContract(req, res);
    if (req.method === "POST" && pathname === "/api/contract-templates/import") return handleImportContractTemplate(req, res);
    if (req.method === "POST" && pathname === "/api/contract-templates/delete") return handleDeleteContractTemplate(req, res);
    if (req.method === "POST" && pathname === "/api/contracts/parse") return handleParseContract(req, res);
    if (req.method === "POST" && pathname === "/api/contracts/export-docx") return handleExportPatchedDocx(req, res);
    if (req.method === "POST" && pathname === "/api/contracts/review") return handleReviewContract(req, res);
    if (req.method === "POST" && pathname === "/api/documents/parse-formatted") return handleParseFormatted(req, res);
    if (req.method === "POST" && pathname === "/api/contracts/chat") return handleContractChat(req, res);
    if (req.method === "POST" && pathname === "/api/web-search") return handleWebSearch(req, res);
    if (req.method === "GET" && pathname.startsWith("/api/documents/file/")) return handleServeDocxFixed(req, res);
    if (req.method === "POST" && pathname === "/api/templates/parse") return handleParseTemplateFixed(req, res);
    if (req.method === "POST" && pathname === "/api/templates/upload") return handleUploadTemplateFixed(req, res);
    if (req.method === "POST" && pathname === "/api/templates/fill") return handleFillTemplateFixed(req, res);
    if (req.method === "POST" && pathname === "/api/templates/ai-analyze") return handleAIAnalyzeTemplate(req, res);

    // Smart Draft endpoints
    if (req.method === "POST" && pathname === "/api/smart-draft/init") {
      let b = ""; req.on("data", (c) => (b += c)); req.on("end", () => handleSmartDraftInit(req, res, b)); return;
    }
    if (req.method === "POST" && pathname === "/api/smart-draft/generate") {
      let b = ""; req.on("data", (c) => (b += c)); req.on("end", () => handleSmartDraftGenerate(req, res, b)); return;
    }
    if (req.method === "POST" && pathname === "/api/smart-draft/answer") {
      let b = ""; req.on("data", (c) => (b += c)); req.on("end", () => handleSmartDraftAnswer(req, res, b)); return;
    }
    if (req.method === "POST" && pathname === "/api/smart-draft/guide") {
      let b = ""; req.on("data", (c) => (b += c)); req.on("end", () => handleSmartDraftGuide(req, res, b)); return;
    }
    if (req.method === "POST" && pathname === "/api/smart-draft/extract-file") {
      let b = ""; req.on("data", (c) => (b += c)); req.on("end", () => handleSmartDraftExtractFile(req, res, b)); return;
    }
    if (req.method === "GET" && pathname.startsWith("/api/smart-draft/session/")) {
      return handleSmartDraftSession(req, res, pathname.split("/").pop());
    }
    if (req.method === "GET" && pathname === "/api/smart-draft/templates") {
      return handleSmartDraftTemplates(req, res);
    }

    sendJson(res, 404, { error: "Not found" });
  } catch (error) {
    sendJson(res, 500, { error: error.message });
  }
}

function serveStatic(req, res, pathname) {
  const requested = pathname === "/" ? "/home.html" : pathname;
  const filePath = path.normalize(path.join(ROOT, requested));
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    return res.end("Forbidden");
  }
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    res.writeHead(404);
    return res.end("Not found");
  }
  res.writeHead(200, { "content-type": contentTypes[path.extname(filePath)] || "application/octet-stream" });
  fs.createReadStream(filePath).pipe(res);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "access-control-allow-origin": "*",
      "access-control-allow-methods": "GET,POST,OPTIONS",
      "access-control-allow-headers": "content-type,authorization",
    });
    return res.end();
  }
  if (url.pathname.startsWith("/api/")) handleApi(req, res, url.pathname);
  else serveStatic(req, res, url.pathname);
});

initializeStore().then(() => {
  server.listen(PORT, () => {
    const modelInfo = activeModelInfo();
    console.log(`ContractAI running at http://localhost:${PORT}`);
    console.log(`Model provider: ${modelInfo.provider} (${modelInfo.model})`);
    console.log(`Xiaomi MiMo API key configured: ${XIAOMI_API_KEY ? "yes" : "no"}`);
    console.log(`OpenAI API key configured: ${OPENAI_API_KEY ? "yes" : "no"}`);
    console.log(`DashScope API key configured: ${DASHSCOPE_API_KEY ? "yes" : "no"}`);
    console.log(`Store backend: ${storeBackend}`);

    // Auto-start token-monitor
    const tokenMonitorPath = path.join(__dirname, "..", "token-monitor", "server.js");
    if (fs.existsSync(tokenMonitorPath)) {
      const tm = spawn(process.execPath, [tokenMonitorPath], {
        stdio: "inherit",
        detached: false,
      });
      tm.on("error", (err) => console.error("[token-monitor] start failed:", err.message));
      console.log(`Token Monitor started at http://localhost:3001`);
    }
  });
});

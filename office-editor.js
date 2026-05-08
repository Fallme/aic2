class OfficeEditor {
  constructor(container, options = {}) {
    this.container = typeof container === "string" ? document.getElementById(container) : container;
    if (!this.container) throw new Error("OfficeEditor: container not found");
    this.options = Object.assign({
      mode: "word",
      placeholder: "在此输入内容…",
      onChange: null,
      onSelectionChange: null,
      readOnly: false,
      spellcheck: false,
      minZoom: 60,
      maxZoom: 200,
      defaultZoom: 100,
      defaultFontFamily: "SimSun",
      defaultFontSize: "12pt",
      contractMode: true,
    }, options);

    this.zoom = this.options.defaultZoom;
    this.history = [];
    this.historyIndex = -1;
    this.historyLocked = false;
    this.maxHistory = 100;
    this.findState = { query: "", caseSensitive: false, matches: [], current: -1 };
    this._selectionChangeTimer = null;
    this._inputTimer = null;
    this._tableMenuOpen = false;

    this._build();
    this._bindToolbar();
    this._bindEditor();
    this._bindKeyboard();
    this._pushHistory();
    this._updateToolbarState();
  }

  _build() {
    this.container.classList.add("oe-wrapper");
    this.container.innerHTML = "";

    this.toolbar = this._createToolbar();
    this.container.appendChild(this.toolbar);

    this.scrollWrap = document.createElement("div");
    this.scrollWrap.className = "oe-scroll-wrap";

    this.pageWrap = document.createElement("div");
    this.pageWrap.className = "oe-page-wrap";
    this.pageWrap.style.fontSize = this.zoom + "%";

    this.editor = document.createElement("div");
    this.editor.className = "oe-editor";
    this.editor.contentEditable = !this.options.readOnly;
    this.editor.spellcheck = this.options.spellcheck;
    this.editor.setAttribute("data-placeholder", this.options.placeholder);
    if (this.options.contractMode) this.editor.classList.add("oe-contract");

    this.pageWrap.appendChild(this.editor);
    this.scrollWrap.appendChild(this.pageWrap);
    this.container.appendChild(this.scrollWrap);

    this.statusBar = this._createStatusBar();
    this.container.appendChild(this.statusBar);

    this.findBar = this._createFindBar();
    this.container.appendChild(this.findBar);
  }

  _createToolbar() {
    const bar = document.createElement("div");
    bar.className = "oe-toolbar";

    bar.innerHTML = `
      <div class="oe-tb-group" data-group="history">
        <button class="oe-btn" data-cmd="undo" title="撤销 (Ctrl+Z)"><svg viewBox="0 0 16 16" width="16" height="16"><path d="M3.5 6.5L1 4l2.5-2.5" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M1 4h9a4 4 0 0 1 0 8H7" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="redo" title="重做 (Ctrl+Y)"><svg viewBox="0 0 16 16" width="16" height="16"><path d="M12.5 6.5L15 4l-2.5-2.5" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M15 4H6a4 4 0 0 0 0 8h3" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
      </div>
      <div class="oe-tb-sep"></div>
      <div class="oe-tb-group" data-group="font">
        <select class="oe-select" data-cmd="fontName" title="字体">
          <option value="SimSun">宋体</option>
          <option value="SimHei">黑体</option>
          <option value="KaiTi">楷体</option>
          <option value="FangSong">仿宋</option>
          <option value="Microsoft YaHei">微软雅黑</option>
          <option value="Noto Sans SC">思源黑体</option>
          <option value="Arial">Arial</option>
          <option value="Times New Roman">Times New Roman</option>
        </select>
        <select class="oe-select oe-font-size" data-cmd="fontSize" title="字号">
          <option value="1">八号</option>
          <option value="2">七号</option>
          <option value="3">小六</option>
          <option value="4">六号</option>
          <option value="5">小五</option>
          <option value="6" selected>五号</option>
          <option value="7">小四</option>
          <option value="8">四号</option>
          <option value="9">小三</option>
          <option value="10">三号</option>
          <option value="11">小二</option>
          <option value="12">二号</option>
          <option value="13">小一</option>
          <option value="14">一号</option>
          <option value="15">小初</option>
          <option value="16">初号</option>
        </select>
      </div>
      <div class="oe-tb-sep"></div>
      <div class="oe-tb-group" data-group="format">
        <button class="oe-btn" data-cmd="bold" title="加粗 (Ctrl+B)"><strong>B</strong></button>
        <button class="oe-btn" data-cmd="italic" title="斜体 (Ctrl+I)"><em>I</em></button>
        <button class="oe-btn" data-cmd="underline" title="下划线 (Ctrl+U)"><u>U</u></button>
        <button class="oe-btn" data-cmd="strikeThrough" title="删除线"><s>S</s></button>
        <button class="oe-btn oe-color-btn" data-cmd="foreColor" title="字体颜色">
          <span class="oe-color-icon">A</span>
          <input type="color" class="oe-color-input" data-cmd="foreColor" value="#000000" />
        </button>
        <button class="oe-btn oe-color-btn" data-cmd="hiliteColor" title="高亮背景">
          <span class="oe-color-icon oe-hilite-icon">ab</span>
          <input type="color" class="oe-color-input" data-cmd="hiliteColor" value="#ffff00" />
        </button>
      </div>
      <div class="oe-tb-sep"></div>
      <div class="oe-tb-group" data-group="heading">
        <select class="oe-select oe-heading-select" data-cmd="formatBlock" title="段落样式">
          <option value="p">正文</option>
          <option value="h1">标题 1</option>
          <option value="h2">标题 2</option>
          <option value="h3">标题 3</option>
          <option value="h4">标题 4</option>
        </select>
      </div>
      <div class="oe-tb-sep"></div>
      <div class="oe-tb-group" data-group="align">
        <button class="oe-btn" data-cmd="justifyLeft" title="左对齐"><svg viewBox="0 0 16 16" width="15" height="15"><path d="M1 2h14M1 5h8M1 8h14M1 11h8M1 14h14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="justifyCenter" title="居中"><svg viewBox="0 0 16 16" width="15" height="15"><path d="M1 2h14M4 5h8M1 8h14M4 11h8M1 14h14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="justifyRight" title="右对齐"><svg viewBox="0 0 16 16" width="15" height="15"><path d="M1 2h14M7 5h8M1 8h14M7 11h8M1 14h14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="justifyFull" title="两端对齐"><svg viewBox="0 0 16 16" width="15" height="15"><path d="M1 2h14M1 5h14M1 8h14M1 11h14M1 14h14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
      </div>
      <div class="oe-tb-sep"></div>
      <div class="oe-tb-group" data-group="list">
        <button class="oe-btn" data-cmd="insertUnorderedList" title="无序列表"><svg viewBox="0 0 16 16" width="15" height="15"><circle cx="2.5" cy="3" r="1.5" fill="currentColor"/><circle cx="2.5" cy="8" r="1.5" fill="currentColor"/><circle cx="2.5" cy="13" r="1.5" fill="currentColor"/><path d="M6 3h9M6 8h9M6 13h9" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="insertOrderedList" title="有序列表"><svg viewBox="0 0 16 16" width="15" height="15"><text x="1" y="5" fill="currentColor" font-size="5" font-family="sans-serif">1.</text><text x="1" y="10" fill="currentColor" font-size="5" font-family="sans-serif">2.</text><text x="1" y="15" fill="currentColor" font-size="5" font-family="sans-serif">3.</text><path d="M6 3h9M6 8h9M6 13h9" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="indent" title="增加缩进"><svg viewBox="0 0 16 16" width="15" height="15"><path d="M1 2h14M5 5l4 3-4 3M1 14h14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg></button>
        <button class="oe-btn" data-cmd="outdent" title="减少缩进"><svg viewBox="0 0 16 16" width="15" height="15"><path d="M1 2h14M9 5L5 8l4 3M1 14h14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg></button>
      </div>
      <div class="oe-tb-sep"></div>
      <div class="oe-tb-group" data-group="insert">
        <button class="oe-btn" data-cmd="insertTable" title="插入表格"><svg viewBox="0 0 16 16" width="15" height="15"><rect x="1" y="1" width="14" height="14" rx="1" fill="none" stroke="currentColor" stroke-width="1.2"/><line x1="1" y1="5.5" x2="15" y2="5.5" stroke="currentColor" stroke-width="1"/><line x1="1" y1="10.5" x2="15" y2="10.5" stroke="currentColor" stroke-width="1"/><line x1="5.5" y1="1" x2="5.5" y2="15" stroke="currentColor" stroke-width="1"/><line x1="10.5" y1="1" x2="10.5" y2="15" stroke="currentColor" stroke-width="1"/></svg></button>
        <button class="oe-btn" data-cmd="insertHR" title="水平线"><svg viewBox="0 0 16 16" width="15" height="15"><line x1="1" y1="8" x2="15" y2="8" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="insertImage" title="插入图片"><svg viewBox="0 0 16 16" width="15" height="15"><rect x="1" y="2" width="14" height="12" rx="1" fill="none" stroke="currentColor" stroke-width="1.2"/><circle cx="5" cy="6" r="1.5" fill="currentColor"/><path d="M1 12l4-4 3 3 2-2 5 5" fill="none" stroke="currentColor" stroke-width="1" stroke-linejoin="round"/></svg></button>
        <button class="oe-btn" data-cmd="insertLink" title="插入链接"><svg viewBox="0 0 16 16" width="15" height="15"><path d="M6.5 9.5a3 3 0 0 0 4.2 0l2-2a3 3 0 0 0-4.2-4.2L7.5 4.3" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/><path d="M9.5 6.5a3 3 0 0 0-4.2 0l-2 2a3 3 0 0 0 4.2 4.2l1-1" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg></button>
      </div>
      <div class="oe-tb-sep"></div>
      <div class="oe-tb-group" data-group="tools">
        <button class="oe-btn" data-cmd="find" title="查找替换 (Ctrl+H)"><svg viewBox="0 0 16 16" width="15" height="15"><circle cx="7" cy="7" r="4" fill="none" stroke="currentColor" stroke-width="1.3"/><line x1="10" y1="10" x2="14" y2="14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="clearFormat" title="清除格式"><svg viewBox="0 0 16 16" width="15" height="15"><text x="1" y="12" fill="currentColor" font-size="12" font-weight="bold" font-family="serif">A</text><line x1="4" y1="3" x2="13" y2="14" stroke="#d00" stroke-width="1.5" stroke-linecap="round"/></svg></button>
        <button class="oe-btn" data-cmd="print" title="打印 (Ctrl+P)"><svg viewBox="0 0 16 16" width="15" height="15"><rect x="3" y="1" width="10" height="5" fill="none" stroke="currentColor" stroke-width="1.2"/><rect x="1" y="5" width="14" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="1.2"/><rect x="3" y="10" width="10" height="5" fill="none" stroke="currentColor" stroke-width="1.2"/></svg></button>
      </div>
      <div class="oe-tb-spacer"></div>
      <div class="oe-tb-group oe-zoom-group" data-group="zoom">
        <button class="oe-btn" data-cmd="zoomOut" title="缩小"><svg viewBox="0 0 16 16" width="14" height="14"><line x1="3" y1="8" x2="13" y2="8" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
        <span class="oe-zoom-label" data-cmd="zoomReset" title="重置缩放">${this.zoom}%</span>
        <button class="oe-btn" data-cmd="zoomIn" title="放大"><svg viewBox="0 0 16 16" width="14" height="14"><line x1="3" y1="8" x2="13" y2="8" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="8" y1="3" x2="8" y2="13" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
      </div>
    `;
    return bar;
  }

  _createStatusBar() {
    const bar = document.createElement("div");
    bar.className = "oe-statusbar";
    bar.innerHTML = `
      <span class="oe-status-chars">字数：0</span>
      <span class="oe-status-words">段落：0</span>
      <span class="oe-status-sel"></span>
      <span class="oe-status-spacer"></span>
      <span class="oe-status-pos">第 1 行，第 1 列</span>
    `;
    return bar;
  }

  _createFindBar() {
    const bar = document.createElement("div");
    bar.className = "oe-findbar";
    bar.hidden = true;
    bar.innerHTML = `
      <div class="oe-find-row">
        <input class="oe-find-input" type="text" placeholder="查找…" />
        <label class="oe-find-label"><input type="checkbox" class="oe-find-case" /> 区分大小写</label>
        <span class="oe-find-count">0 / 0</span>
        <button class="oe-btn oe-find-prev" title="上一个">▲</button>
        <button class="oe-btn oe-find-next" title="下一个">▼</button>
      </div>
      <div class="oe-find-row">
        <input class="oe-replace-input" type="text" placeholder="替换…" />
        <button class="oe-btn oe-replace-one">替换</button>
        <button class="oe-btn oe-replace-all">全部替换</button>
        <button class="oe-btn oe-find-close" title="关闭">✕</button>
      </div>
    `;
    return bar;
  }

  _bindToolbar() {
    const bar = this.toolbar;
    bar.addEventListener("mousedown", (e) => {
      const btn = e.target.closest("[data-cmd]");
      if (!btn) return;
      e.preventDefault();
    });
    bar.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-cmd]");
      if (!btn) return;
      const cmd = btn.dataset.cmd;
      this._execToolbarCommand(cmd, btn);
    });
    bar.addEventListener("change", (e) => {
      const sel = e.target.closest("[data-cmd]");
      if (!sel) return;
      const cmd = sel.dataset.cmd;
      this._execToolbarCommand(cmd, sel);
    });
    bar.querySelectorAll("input[type=color].oe-color-input").forEach((input) => {
      input.addEventListener("input", () => {
        const cmd = input.dataset.cmd;
        this._applyColor(cmd, input.value);
      });
    });

    this._bindFindBar();
  }

  _execToolbarCommand(cmd, el) {
    this.editor.focus();
    switch (cmd) {
      case "undo":
        this._undo();
        break;
      case "redo":
        this._redo();
        break;
      case "fontName":
        document.execCommand("fontName", false, el.value);
        this._ensureContractFont();
        break;
      case "fontSize":
        document.execCommand("fontSize", false, el.value);
        this._normalizeFontSizes();
        break;
      case "formatBlock": {
        const tag = el.value;
        if (tag === "p") document.execCommand("formatBlock", false, "p");
        else document.execCommand("formatBlock", false, tag);
        break;
      }
      case "bold":
      case "italic":
      case "underline":
      case "strikeThrough":
      case "justifyLeft":
      case "justifyCenter":
      case "justifyRight":
      case "justifyFull":
      case "insertUnorderedList":
      case "insertOrderedList":
      case "indent":
      case "outdent":
        document.execCommand(cmd, false, null);
        break;
      case "foreColor":
      case "hiliteColor":
        this._colorInputClick(el);
        break;
      case "clearFormat":
        document.execCommand("removeFormat", false, null);
        break;
      case "insertTable":
        this._insertTable();
        break;
      case "insertHR":
        document.execCommand("insertHorizontalRule", false, null);
        break;
      case "insertImage":
        this._insertImage();
        break;
      case "insertLink":
        this._insertLink();
        break;
      case "find":
        this._toggleFindBar();
        break;
      case "print":
        this._print();
        break;
      case "zoomIn":
        this._setZoom(this.zoom + 10);
        break;
      case "zoomOut":
        this._setZoom(this.zoom - 10);
        break;
      case "zoomReset":
        this._setZoom(100);
        break;
      default:
        break;
    }
    this._onChange();
    this._updateToolbarState();
  }

  _colorInputClick(btn) {
    const input = btn.querySelector("input[type=color]");
    if (input) input.click();
  }

  _applyColor(cmd, value) {
    this.editor.focus();
    if (cmd === "hiliteColor") {
      document.execCommand("hiliteColor", false, value);
    } else {
      document.execCommand("foreColor", false, value);
    }
    this._onChange();
  }

  _insertTable() {
    const rows = parseInt(prompt("行数", "3"), 10) || 3;
    const cols = parseInt(prompt("列数", "3"), 10) || 3;
    if (rows < 1 || cols < 1 || rows > 50 || cols > 20) return;
    let html = '<table class="oe-table"><tbody>';
    for (let r = 0; r < rows; r++) {
      html += "<tr>";
      for (let c = 0; c < cols; c++) {
        html += r === 0 ? "<th><br></th>" : "<td><br></td>";
      }
      html += "</tr>";
    }
    html += "</tbody></table><p><br></p>";
    document.execCommand("insertHTML", false, html);
  }

  _insertImage() {
    const url = prompt("输入图片 URL：");
    if (url) document.execCommand("insertImage", false, url);
  }

  _insertLink() {
    const sel = window.getSelection();
    const text = sel.toString().trim();
    const url = prompt("输入链接 URL：", "https://");
    if (!url) return;
    if (!text) {
      document.execCommand("insertHTML", false, `<a href="${this._escapeHtml(url)}" target="_blank">${this._escapeHtml(url)}</a>`);
    } else {
      document.execCommand("createLink", false, url);
    }
  }

  _print() {
    const win = window.open("", "_blank");
    if (!win) return;
    win.document.write(this._buildPrintHtml());
    win.document.close();
    win.focus();
    win.print();
  }

  _buildPrintHtml() {
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>打印</title><style>
@page{margin:2.54cm 3.18cm}
body{margin:0;color:#000;font-family:"SimSun","宋体",serif;font-size:12pt;line-height:1.8}
h1{text-align:center;font-family:"SimHei","黑体",sans-serif;font-size:22pt;font-weight:700;margin:0 0 24pt}
h2{font-family:"SimHei","黑体",sans-serif;font-size:16pt;font-weight:700;margin:16pt 0 8pt}
h3{font-family:"SimHei","黑体",sans-serif;font-size:14pt;margin:12pt 0 6pt}
p{margin:0 0 6pt;text-indent:2em}
.clause-line,.list-line{margin-top:8pt;text-indent:0}
.list-line{margin-left:2em;text-indent:-2em}
table{width:100%;border-collapse:collapse;margin:8pt 0}
td,th{border:1px solid #000;padding:6pt;text-align:left;font-size:10.5pt}
th{background:#f0f0f0;font-weight:700}
img{max-width:100%}
hr{border:none;border-top:1px solid #999;margin:12pt 0}
</style></head><body>${this.editor.innerHTML}</body></html>`;
  }

  _bindEditor() {
    this.editor.addEventListener("input", () => {
      this._onChange();
      this._debouncedPushHistory();
    });
    this.editor.addEventListener("mouseup", () => this._onSelectionChange());
    this.editor.addEventListener("keyup", () => this._onSelectionChange());
    document.addEventListener("selectionchange", () => {
      if (!this.editor.contains(document.activeElement) && !this._selInEditor()) return;
      this._onSelectionChange();
    });
    this.editor.addEventListener("paste", (e) => this._onPaste(e));
    this.editor.addEventListener("drop", (e) => {
      e.preventDefault();
      const text = e.dataTransfer.getData("text/plain");
      if (text) document.execCommand("insertText", false, text);
    });
    this.editor.addEventListener("keydown", (e) => {
      if (e.key === "Tab") {
        e.preventDefault();
        if (e.shiftKey) document.execCommand("outdent", false, null);
        else document.execCommand("indent", false, null);
        this._onChange();
      }
    });
  }

  _selInEditor() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return false;
    return this.editor.contains(sel.anchorNode);
  }

  _onPaste(e) {
    e.preventDefault();
    const clip = e.clipboardData || window.clipboardData;
    if (!clip) return;
    const html = clip.getData("text/html");
    const text = clip.getData("text/plain");
    if (html) {
      const cleaned = this._sanitizeHtml(html);
      document.execCommand("insertHTML", false, cleaned);
    } else if (text) {
      document.execCommand("insertText", false, text);
    }
    this._onChange();
  }

  _sanitizeHtml(html) {
    const div = document.createElement("div");
    div.innerHTML = html;
    div.querySelectorAll("script,style,meta,link,iframe,object,embed").forEach((el) => el.remove());
    div.querySelectorAll("*").forEach((el) => {
      [...el.attributes].forEach((attr) => {
        if (/^on/i.test(attr.name)) el.removeAttribute(attr.name);
      });
    });
    return div.innerHTML;
  }

  _bindKeyboard() {
    document.addEventListener("keydown", (e) => {
      if (!this._selInEditor()) return;
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key === "z" && !e.shiftKey) {
        e.preventDefault();
        this._undo();
      }
      if (mod && (e.key === "y" || (e.key === "z" && e.shiftKey))) {
        e.preventDefault();
        this._redo();
      }
      if (mod && e.key === "h") {
        e.preventDefault();
        this._toggleFindBar();
      }
      if (mod && e.key === "p") {
        e.preventDefault();
        this._print();
      }
    });
  }

  _bindFindBar() {
    const fb = this.findBar;
    fb.querySelector(".oe-find-input").addEventListener("input", (e) => {
      this.findState.query = e.target.value;
      this._executeFind();
    });
    fb.querySelector(".oe-find-case").addEventListener("change", (e) => {
      this.findState.caseSensitive = e.target.checked;
      this._executeFind();
    });
    fb.querySelector(".oe-find-next").addEventListener("click", () => this._findNext());
    fb.querySelector(".oe-find-prev").addEventListener("click", () => this._findPrev());
    fb.querySelector(".oe-replace-one").addEventListener("click", () => this._replaceOne());
    fb.querySelector(".oe-replace-all").addEventListener("click", () => this._replaceAll());
    fb.querySelector(".oe-find-close").addEventListener("click", () => this._toggleFindBar());
    fb.querySelector(".oe-find-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        this._findNext();
      }
      if (e.key === "Escape") {
        this._toggleFindBar();
      }
    });
  }

  _toggleFindBar() {
    this.findBar.hidden = !this.findBar.hidden;
    if (!this.findBar.hidden) {
      this.findBar.querySelector(".oe-find-input").focus();
      const sel = window.getSelection();
      if (sel && sel.toString()) {
        this.findBar.querySelector(".oe-find-input").value = sel.toString();
        this.findState.query = sel.toString();
        this._executeFind();
      }
    } else {
      this._clearHighlights();
    }
  }

  _executeFind() {
    this._clearHighlights();
    const q = this.findState.query;
    if (!q) {
      this.findState.matches = [];
      this.findState.current = -1;
      this._updateFindCount();
      return;
    }
    const text = this.editor.innerHTML;
    const flags = this.findState.caseSensitive ? "g" : "gi";
    const escaped = q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(escaped, flags);
    const matches = [];
    let match;
    while ((match = re.exec(text)) !== null) {
      matches.push({ start: match.index, end: match.index + match[0].length });
    }
    this.findState.matches = matches;
    this.findState.current = matches.length ? 0 : -1;
    this._highlightMatches();
    this._updateFindCount();
    if (matches.length) this._scrollToMatch(0);
  }

  _highlightMatches() {
    const walker = document.createTreeWalker(this.editor, NodeFilter.SHOW_TEXT, null, false);
    const text = this.findState.query;
    if (!text) return;
    const flags = this.findState.caseSensitive ? "g" : "gi";
    const escaped = text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(escaped, flags);
    const textNodes = [];
    while (walker.nextNode()) textNodes.push(walker.currentNode);
    let globalIdx = 0;
    let matchIdx = 0;
    for (const node of textNodes) {
      const val = node.nodeValue;
      let m;
      re.lastIndex = 0;
      const fragments = [];
      let last = 0;
      while ((m = re.exec(val)) !== null) {
        fragments.push({ text: val.slice(last, m.index), highlight: false });
        fragments.push({ text: m[0], highlight: true });
        last = m.index + m[0].length;
        globalIdx++;
      }
      if (!fragments.length) continue;
      fragments.push({ text: val.slice(last), highlight: false });
      const parent = node.parentNode;
      for (const frag of fragments) {
        if (!frag.text) continue;
        if (frag.highlight) {
          const mark = document.createElement("mark");
          mark.className = "oe-highlight" + (matchIdx === this.findState.current ? " oe-highlight-active" : "");
          mark.textContent = frag.text;
          mark.dataset.matchIdx = matchIdx++;
          parent.insertBefore(mark, node);
        } else {
          parent.insertBefore(document.createTextNode(frag.text), node);
        }
      }
      parent.removeChild(node);
    }
  }

  _clearHighlights() {
    this.editor.querySelectorAll("mark.oe-highlight").forEach((mark) => {
      const text = document.createTextNode(mark.textContent);
      mark.parentNode.replaceChild(text, mark);
    });
    this.editor.normalize();
  }

  _findNext() {
    if (!this.findState.matches.length) return;
    this.findState.current = (this.findState.current + 1) % this.findState.matches.length;
    this._executeFind();
  }

  _findPrev() {
    if (!this.findState.matches.length) return;
    this.findState.current = (this.findState.current - 1 + this.findState.matches.length) % this.findState.matches.length;
    this._executeFind();
  }

  _scrollToMatch(idx) {
    const active = this.editor.querySelector("mark.oe-highlight-active");
    if (active) active.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  _replaceOne() {
    const replaceVal = this.findBar.querySelector(".oe-replace-input").value;
    const active = this.editor.querySelector("mark.oe-highlight-active");
    if (!active) return;
    active.replaceWith(document.createTextNode(replaceVal));
    this.editor.normalize();
    this._onChange();
    this._executeFind();
  }

  _replaceAll() {
    const replaceVal = this.findBar.querySelector(".oe-replace-input").value;
    const marks = this.editor.querySelectorAll("mark.oe-highlight");
    if (!marks.length) return;
    marks.forEach((mark) => mark.replaceWith(document.createTextNode(replaceVal)));
    this.editor.normalize();
    this._onChange();
    this._executeFind();
  }

  _updateFindCount() {
    const countEl = this.findBar.querySelector(".oe-find-count");
    const total = this.findState.matches.length;
    const cur = total ? this.findState.current + 1 : 0;
    countEl.textContent = `${cur} / ${total}`;
  }

  _onChange() {
    this._updateStatusBar();
    this._updateToolbarState();
    if (this.options.onChange) this.options.onChange(this);
  }

  _onSelectionChange() {
    clearTimeout(this._selectionChangeTimer);
    this._selectionChangeTimer = setTimeout(() => {
      this._updateToolbarState();
      this._updateStatusBar();
      if (this.options.onSelectionChange) this.options.onSelectionChange(this);
    }, 50);
  }

  _updateToolbarState() {
    const bar = this.toolbar;
    const setActive = (cmd, active) => {
      const btn = bar.querySelector(`[data-cmd="${cmd}"]`);
      if (btn) btn.classList.toggle("active", active);
    };
    setActive("bold", document.queryCommandState("bold"));
    setActive("italic", document.queryCommandState("italic"));
    setActive("underline", document.queryCommandState("underline"));
    setActive("strikeThrough", document.queryCommandState("strikeThrough"));
    setActive("justifyLeft", document.queryCommandState("justifyLeft"));
    setActive("justifyCenter", document.queryCommandState("justifyCenter"));
    setActive("justifyRight", document.queryCommandState("justifyRight"));
    setActive("justifyFull", document.queryCommandState("justifyFull"));
    setActive("insertUnorderedList", document.queryCommandState("insertUnorderedList"));
    setActive("insertOrderedList", document.queryCommandState("insertOrderedList"));

    const fontNameSel = bar.querySelector('[data-cmd="fontName"]');
    const currentFont = document.queryCommandValue("fontName");
    if (fontNameSel && currentFont) {
      const match = [...fontNameSel.options].find((o) => currentFont.toLowerCase().includes(o.value.toLowerCase()));
      if (match) fontNameSel.value = match.value;
    }

    const fontSizeSel = bar.querySelector('[data-cmd="fontSize"]');
    const currentSize = document.queryCommandValue("fontSize");
    if (fontSizeSel && currentSize) fontSizeSel.value = currentSize;

    const headingSel = bar.querySelector('[data-cmd="formatBlock"]');
    const block = document.queryCommandValue("formatBlock");
    if (headingSel) {
      const tag = block.toLowerCase().replace(/[<>]/g, "");
      headingSel.value = ["h1", "h2", "h3", "h4"].includes(tag) ? tag : "p";
    }
  }

  _updateStatusBar() {
    const text = this.getText();
    const chars = text.replace(/\s/g, "").length;
    const lines = text.split(/\n/).filter(Boolean);
    const paras = this.editor.querySelectorAll("p,h1,h2,h3,h4,h5,h6,li,td,th,blockquote");
    this.statusBar.querySelector(".oe-status-chars").textContent = `字数：${chars}`;
    this.statusBar.querySelector(".oe-status-words").textContent = `段落：${paras.length || lines.length}`;

    const sel = window.getSelection();
    if (sel && sel.rangeCount && this.editor.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      const pre = range.cloneRange();
      pre.selectNodeContents(this.editor);
      pre.setEnd(range.startContainer, range.startOffset);
      const before = pre.toString();
      const lastNl = before.lastIndexOf("\n");
      const line = before.slice(lastNl + 1);
      const rowNum = (before.match(/\n/g) || []).length + 1;
      const colNum = line.length + 1;
      this.statusBar.querySelector(".oe-status-pos").textContent = `第 ${rowNum} 行，第 ${colNum} 列`;
      const selected = sel.toString();
      if (selected) {
        this.statusBar.querySelector(".oe-status-sel").textContent = `已选 ${selected.length} 字`;
      } else {
        this.statusBar.querySelector(".oe-status-sel").textContent = "";
      }
    } else {
      this.statusBar.querySelector(".oe-status-pos").textContent = "";
      this.statusBar.querySelector(".oe-status-sel").textContent = "";
    }
  }

  _pushHistory() {
    if (this.historyLocked) return;
    const html = this.editor.innerHTML;
    if (this.history.length && this.history[this.historyIndex] === html) return;
    if (this.historyIndex < this.history.length - 1) {
      this.history = this.history.slice(0, this.historyIndex + 1);
    }
    this.history.push(html);
    if (this.history.length > this.maxHistory) this.history.shift();
    this.historyIndex = this.history.length - 1;
  }

  _debouncedPushHistory() {
    clearTimeout(this._inputTimer);
    this._inputTimer = setTimeout(() => this._pushHistory(), 400);
  }

  _undo() {
    if (this.historyIndex <= 0) return;
    this.historyIndex--;
    this._restoreHistory();
  }

  _redo() {
    if (this.historyIndex >= this.history.length - 1) return;
    this.historyIndex++;
    this._restoreHistory();
  }

  _restoreHistory() {
    this.historyLocked = true;
    this.editor.innerHTML = this.history[this.historyIndex] || "";
    this.historyLocked = false;
    this._onChange();
  }

  _setZoom(value) {
    this.zoom = Math.max(this.options.minZoom, Math.min(this.options.maxZoom, value));
    this.pageWrap.style.fontSize = this.zoom + "%";
    this.toolbar.querySelector(".oe-zoom-label").textContent = this.zoom + "%";
  }

  _ensureContractFont() {
    if (!this.options.contractMode) return;
  }

  _normalizeFontSizes() {
    this.editor.querySelectorAll("font[size]").forEach((font) => {
      const size = parseInt(font.getAttribute("size"), 10);
      const ptMap = { 1: "8pt", 2: "9pt", 3: "10pt", 4: "10.5pt", 5: "11pt", 6: "12pt", 7: "14pt" };
      const pt = ptMap[size] || "12pt";
      const span = document.createElement("span");
      span.style.fontSize = pt;
      span.innerHTML = font.innerHTML;
      font.replaceWith(span);
    });
  }

  _escapeHtml(text) {
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  getHtml() {
    return this.editor.innerHTML;
  }

  setHtml(html) {
    this.historyLocked = true;
    this.editor.innerHTML = html || "";
    this.historyLocked = false;
    this._pushHistory();
    this._updateStatusBar();
  }

  getText() {
    return this.editor.innerText.trim();
  }

  setText(text) {
    this.setHtml(this.textToContractHtml(text));
  }

  textToContractHtml(text) {
    const lines = String(text || "").split(/\n/);
    let firstContent = true;
    return lines
      .map((line) => {
        const value = line.trimEnd();
        const compact = value.trim();
        if (!compact) return '<p class="contract-blank"><br></p>';
        if (firstContent && compact.length <= 60 && /合同|协议|承诺书|确认书|订单|补充协议/.test(compact)) {
          firstContent = false;
          return `<h1>${this._escapeHtml(compact)}</h1>`;
        }
        firstContent = false;
        if (/^第[一二三四五六七八九十百]+[章节篇]\s*/.test(compact) || /^[一二三四五六七八九十]+[、.．]\s*/.test(compact)) {
          return `<h2>${this._escapeHtml(compact)}</h2>`;
        }
        if (/^第[一二三四五六七八九十百]+条\s*/.test(compact) || /^\d+[、.．]\s*/.test(compact)) {
          return `<p class="clause-line">${this._escapeHtml(compact)}</p>`;
        }
        if (/^（[一二三四五六七八九十\d]+）/.test(compact) || /^\([一二三四五六七八九十\d]+\)/.test(compact)) {
          return `<p class="list-line">${this._escapeHtml(compact)}</p>`;
        }
        return `<p>${this._escapeHtml(compact)}</p>`;
      })
      .join("");
  }

  insertHtmlAtCursor(html) {
    this.editor.focus();
    document.execCommand("insertHTML", false, html);
    this._onChange();
    this._pushHistory();
  }

  insertTextAtCursor(text) {
    this.editor.focus();
    document.execCommand("insertText", false, text);
    this._onChange();
    this._pushHistory();
  }

  appendHtml(html) {
    this.editor.focus();
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(this.editor);
    range.collapse(false);
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand("insertHTML", false, html);
    this._onChange();
    this._pushHistory();
  }

  focus() {
    this.editor.focus();
  }

  setReadOnly(readonly) {
    this.options.readOnly = readonly;
    this.editor.contentEditable = !readonly;
    this.container.classList.toggle("oe-readonly", readonly);
  }

  getWordExportHtml(title) {
    const body = this.editor.innerHTML;
    return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>${this._escapeHtml(title || "合同文档")}</title>
  <style>
    @page { margin: 2.54cm 3.18cm; size: A4; }
    body { margin: 0; color: #000; font-family: "SimSun", "宋体", serif; font-size: 12pt; line-height: 1.8; }
    .contract-doc { max-width: 780px; margin: 0 auto; }
    h1 { margin: 0 0 24pt; text-align: center; font-family: "SimHei", "黑体", sans-serif; font-size: 22pt; font-weight: 700; }
    h2 { margin: 16pt 0 8pt; font-family: "SimHei", "黑体", sans-serif; font-size: 16pt; font-weight: 700; }
    h3 { margin: 12pt 0 6pt; font-family: "SimHei", "黑体", sans-serif; font-size: 14pt; font-weight: 700; }
    p { margin: 0 0 6pt; text-indent: 2em; }
    .clause-line, .list-line { text-indent: 0; }
    .list-line { margin-left: 2em; text-indent: -2em; }
    .contract-blank { min-height: 1em; text-indent: 0; }
    table { width: 100%; border-collapse: collapse; margin: 8pt 0; }
    td, th { border: 1px solid #000; padding: 6pt; text-align: left; font-size: 10.5pt; }
    th { background: #f0f0f0; font-weight: 700; }
    img { max-width: 100%; }
    hr { border: none; border-top: 1px solid #999; margin: 12pt 0; }
    a { color: #000; text-decoration: underline; }
    mark, .risk-mark { background: transparent; color: inherit; }
  </style>
</head>
<body><div class="contract-doc">${body}</div></body>
</html>`;
  }

  downloadDoc(title) {
    const html = this.getWordExportHtml(title);
    const blob = new Blob([html], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${title || "合同"}.doc`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  async copyContent() {
    const text = this.getText();
    if (navigator.clipboard) {
      await navigator.clipboard.writeText(text);
    }
    return text;
  }

  destroy() {
    this.container.classList.remove("oe-wrapper");
    this.container.innerHTML = "";
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = OfficeEditor;
}

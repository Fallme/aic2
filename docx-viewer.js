/**
 * DocxViewer — 基于 docx-preview 的高保真 Word 文档查看器
 *
 * 用法：
 *   const viewer = new DocxViewer(containerEl, { readOnly: true });
 *   await viewer.loadFromUrl('/api/documents/file/xxx');
 *   await viewer.loadFromArrayBuffer(arrayBuffer);
 *   viewer.destroy();
 */

/* global docx, JSZip */

(function () {
  "use strict";

  /* ── CDN 资源 ── */
  const CDN_BASE = "https://cdn.jsdelivr.net/npm";
  const SCRIPTS = [
    { id: "jszip-cdn", src: `${CDN_BASE}/jszip@3.10.1/dist/jszip.min.js` },
    { id: "docx-preview-cdn", src: `${CDN_BASE}/docx-preview@0.1.15/dist/docx-preview.min.js` },
  ];

  let loadPromise = null;

  function loadScripts() {
    if (loadPromise) return loadPromise;
    loadPromise = (async () => {
      for (const s of SCRIPTS) {
        if (document.getElementById(s.id)) continue;
        await new Promise((resolve, reject) => {
          const el = document.createElement("script");
          el.id = s.id;
          el.src = s.src;
          el.crossOrigin = "anonymous";
          el.onload = resolve;
          el.onerror = () => reject(new Error(`Failed to load ${s.src}`));
          document.head.appendChild(el);
        });
      }
    })();
    return loadPromise;
  }

  /* ── DocxViewer 类 ── */

  class DocxViewer {
    /**
     * @param {HTMLElement|string} container — 容器元素或其 ID
     * @param {object} [opts]
     * @param {boolean} [opts.readOnly=true] — 是否只读
     * @param {function} [opts.onLoad] — 加载完成回调
     * @param {function} [opts.onError] — 错误回调
     */
    constructor(container, opts = {}) {
      this.container = typeof container === "string" ? document.getElementById(container) : container;
      this.opts = { readOnly: true, ...opts };
      this._blob = null;     // 当前文件的原始 ArrayBuffer
      this._fileName = "";
      this._loaded = false;
      this._destroyed = false;

      // 创建渲染容器
      this.wrapper = document.createElement("div");
      this.wrapper.className = "docx-viewer-wrapper";
      this.wrapper.style.cssText = "width:100%;overflow:auto;background:#f5f5f5;padding:20px 0;";
      this.pageContainer = document.createElement("div");
      this.pageContainer.className = "docx-viewer-pages";
      this.pageContainer.style.cssText = "max-width:820px;margin:0 auto;background:#fff;box-shadow:0 2px 12px rgba(0,0,0,0.08);border-radius:4px;min-height:400px;";
      this.wrapper.appendChild(this.pageContainer);
      this.container.appendChild(this.wrapper);
    }

    /**
     * 从 URL 加载 .docx 文件
     */
    async loadFromUrl(url, fileName) {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`Failed to fetch document: ${res.status}`);
      const buf = await res.arrayBuffer();
      return this.loadFromArrayBuffer(buf, fileName || url.split("/").pop());
    }

    /**
     * 从 File 对象加载
     */
    async loadFromFile(file) {
      const buf = await file.arrayBuffer();
      return this.loadFromArrayBuffer(buf, file.name);
    }

    /**
     * 从 ArrayBuffer 加载
     */
    async loadFromArrayBuffer(buffer, fileName) {
      if (this._destroyed) return;
      this._blob = buffer;
      this._fileName = fileName || "document.docx";

      try {
        await loadScripts();
        if (this._destroyed) return;

        this.pageContainer.innerHTML = "";

        // docx.renderAsync(data, container, null, options)
        await docx.renderAsync(buffer, this.pageContainer, null, {
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          ignoreLastRenderedPageBreak: true,
          experimental: true,
          trimXmlDeclaration: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
          inWrapper: true,
          convertImage: undefined, // 使用默认图片处理
        });

        this._loaded = true;
        if (this.opts.onLoad) this.opts.onLoad({ fileName: this._fileName });
      } catch (err) {
        console.error("DocxViewer render error:", err);
        if (this.opts.onError) this.opts.onError(err);
        else throw err;
      }
    }

    /**
     * 获取原始文件 ArrayBuffer（用于导出）
     */
    getOriginalBuffer() {
      return this._blob;
    }

    /**
     * 获取文件名
     */
    getFileName() {
      return this._fileName;
    }

    /**
     * 是否已加载
     */
    isLoaded() {
      return this._loaded;
    }

    /**
     * 滚动到指定位置
     */
    scrollTo(opts) {
      this.wrapper.scrollTo(opts);
    }

    /**
     * 高亮指定文本片段（叠加标注层用）
     */
    highlightText(text, className) {
      if (!this.pageContainer) return [];
      const marks = [];
      const walker = document.createTreeWalker(this.pageContainer, NodeFilter.SHOW_TEXT);
      const textNodes = [];
      while (walker.nextNode()) textNodes.push(walker.currentNode);

      const fullText = textNodes.map(n => n.nodeValue).join("");
      const idx = fullText.indexOf(text);
      if (idx < 0) return marks;

      // 找到对应的文本节点并高亮
      let offset = 0;
      for (const node of textNodes) {
        const end = offset + node.nodeValue.length;
        if (idx < end && idx + text.length > offset) {
          const range = document.createRange();
          const startInNode = Math.max(0, idx - offset);
          const endInNode = Math.min(node.nodeValue.length, idx + text.length - offset);
          range.setStart(node, startInNode);
          range.setEnd(node, endInNode);

          const mark = document.createElement("mark");
          mark.className = className || "docx-highlight";
          mark.style.cssText = "background:#fff3cd;padding:1px 0;border-radius:2px;";
          range.surroundContents(mark);
          marks.push(mark);

          if (marks.length > 50) break; // 安全限制
        }
        offset = end;
        if (offset >= idx + text.length) break;
      }
      return marks;
    }

    /**
     * 清除所有高亮
     */
    clearHighlights() {
      if (!this.pageContainer) return;
      this.pageContainer.querySelectorAll("mark.docx-highlight, mark.docx-risk-mark").forEach(mark => {
        const parent = mark.parentNode;
        while (mark.firstChild) parent.insertBefore(mark.firstChild, mark);
        parent.removeChild(mark);
      });
    }

    /**
     * 销毁实例
     */
    destroy() {
      this._destroyed = true;
      this._blob = null;
      this._loaded = false;
      if (this.pageContainer) this.pageContainer.innerHTML = "";
      if (this.wrapper && this.wrapper.parentNode) this.wrapper.parentNode.removeChild(this.wrapper);
    }
  }

  /* ── 导出 ── */
  if (typeof window !== "undefined") {
    window.DocxViewer = DocxViewer;
  }
})();

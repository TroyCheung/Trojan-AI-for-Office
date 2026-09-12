(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ---- CDN 依赖懒加载 ----
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (loadScript._cache[src]) return resolve();
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => { loadScript._cache[src] = true; resolve(); };
      s.onerror = () => reject(new Error('Failed to load ' + src));
      document.head.appendChild(s);
    });
  }
  loadScript._cache = {};

  // 【P2-3】自托管（与 office.js 同理由：境外 CDN 在受限网络下会让附件解析整体不可用）。
  // 版本与来源见 assets/vendor/README-ATTACH.md；升级时连同 taskpane.html 的 ?v= 一起 bump 击穿缓存。
  const PDFJS_URL = 'assets/vendor/pdfjs/pdf.min.js';
  const PDFJS_WORKER = 'assets/vendor/pdfjs/pdf.worker.min.js';
  const FFLATE_URL = 'assets/vendor/fflate/index.js';

  async function parsePdf(arrayBuffer) {
    await loadScript(PDFJS_URL);
    if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
    const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let out = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const tc = await page.getTextContent();
      out.push(tc.items.map(x => x.str).join(' '));
    }
    return out.join('\n\n');
  }

  function xmlToText(xml, tags) {
    try {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      let parts = [];
      for (const tag of tags) {
        // 不能用 querySelectorAll('a:p')：带命名空间前缀的 CSS 选择器在 XML 文档里会抛
        // SyntaxError（被 catch 静默吞掉导致解析结果为空）。改用按限定名/本地名匹配。
        const local = tag.split(':').pop();
        let nodes = doc.getElementsByTagName(tag);           // 按限定名 'a:p' 直接匹配
        if (!nodes.length) nodes = doc.getElementsByTagNameNS('*', local);
        for (const el of nodes) parts.push(el.textContent);
      }
      return parts.join(tags[0].indexOf('p') > -1 ? '\n' : ' ');
    } catch (e) {
      return '';
    }
  }

  async function unzipArchive(arrayBuffer) {
    await loadScript(FFLATE_URL);
    const fflate = window.fflate;
    const files = fflate.unzipSync(new Uint8Array(arrayBuffer));
    return Object.keys(files).filter(path => path && !path.endsWith('/')).sort().map(path => ({ path, data: files[path] }));
  }

  // PPTX 页序自然排序（P2-2）：字典序会把 slide10.xml 排在 slide2.xml 前面，
  // 附件解析出的页序就乱了。只对 pptx 的 slide 文件启用，DOCX 与其他 ZIP 语义不动。
  function naturalSlideOrder(a, b) {
    const num = item => Number((item.path.match(/slide(\d+)\.xml$/) || [0, 0])[1]);
    return num(a) - num(b) || a.path.localeCompare(b.path);
  }

  async function unzipText(arrayBuffer, innerPathPattern, tags, options = {}) {
    const files = await unzipArchive(arrayBuffer);
    const fflate = window.fflate;
    const selected = files.filter(item => innerPathPattern.test(item.path));
    if (options.naturalSlides) selected.sort(naturalSlideOrder);
    let out = [];
    for (const item of selected) {
      const xml = fflate.strFromU8(item.data);
      const text = xmlToText(xml, tags);
      if (text.trim()) out.push(text);
    }
    return out.join('\n\n');
  }

  // ---- 统一入口 ----
  async function parseFile(file) {
    const name = file.name.toLowerCase();
    if (/\.(txt|md|csv|json)$/.test(name)) {
      return await file.text();
    }
    if (/\.pdf$/.test(name)) {
      return await parsePdf(await file.arrayBuffer());
    }
    if (/\.docx$/.test(name)) {
      return await unzipText(await file.arrayBuffer(), /^word\/document\d*\.xml$/, ['w:p']);
    }
    if (/\.pptx$/.test(name)) {
      return await unzipText(await file.arrayBuffer(), /^ppt\/slides\/slide\d+\.xml$/, ['a:p'], { naturalSlides: true });
    }
    const zh = App.state && App.state.locale === 'zh';
    if (/\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(name)) {
      throw new Error(zh ? '图片不走文档上传：用 ➕ 菜单的「上传图片」或直接把图片拖进对话区，然后用 insert_image(attachmentId) 插入文档' : 'Images are not attachments: use the ➕ "Upload image" option or drag the image into the chat, then insert with insert_image(attachmentId)');
    }
    throw new Error(zh ? '暂不支持该格式（支持：txt/md/csv/json/pdf/docx/pptx；图片请用 ➕ 上传图片）' : 'Unsupported format (supported: txt/md/csv/json/pdf/docx/pptx; images go through the ➕ upload-image path)');
  }

  const MAX_CHARS = 30000;
  async function readAttachment(file) {
    const raw = await parseFile(file);
    const text = raw.length > MAX_CHARS ? raw.slice(0, MAX_CHARS) + '\n…(truncated)' : raw;
    return { name: file.name, text };
  }

  App.readAttachment = readAttachment;
  App.unzipArchive = unzipArchive;
})();

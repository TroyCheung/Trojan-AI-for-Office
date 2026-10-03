'use strict';
// 「仿真机」假宿主：按 Office 16.112（Mac）上实测到的行为复刻，
// 而不是按 Office.js 文档的理想行为。以前的 mock 什么都答应，于是 duplicate 在 mock 里是绿的、
// 真机上不存在，测试白跑。这个 mock 会像那台机器一样刁难：
//   1. 一批 load 里只要有一个请求失败，整批 sync 被取消——这才是「批量 sync 一错全错」的真实机制。
//      纯几何属性（left/top/name/type）在所有形状上都成立，所以整页批量读几何是安全的；
//      但 Group / Picture / Line 没有 textFrame（Group/Line 也没有 fill），
//      一旦把它们和普通形状放进同一批 load，整批就被取消，其余形状的数据一起丢。
//   2. Group 的 shapes.items 恒为空，且没有 ungroup()
//   3. slide.duplicate 不存在
//   4. 集合没有 getItem(index)
//   5. fill 没有 setType()
//   6. pageSetup.slideWidth / slideHeight 读回 undefined
//   7. 从未写过备注的页，访问 notesSlide 抛错
//   8. 图片没有原位替换 API（shape.image 不存在）——replace_image 只能同位置删除+插入
//   9. TableCell.fill 未验证可写，按刁难默认不可写（quirk tableCellFill 可打开）
//   10. Chart 的 series.values 读不到（Chart API 薄），chartType/series.name 可读
//   11. 没有 slide.background 背景写 API——set_slide_background 只能降级为全页矩形置底（quirk slideBackgroundApi 可打开）
//   12. Table 没有行列增删方法——edit_table_structure 只能报错给替代路径（quirk tableStructureApi 可打开）
//   13. 图片没有 cropFormat 裁剪 API（quirk pictureCrop 可打开）
//   14. 图片没有 pictureFormat.transparency 透明度写 API（quirk pictureTransparency 可打开）
//   15. presentation.insertSlidesFromBase64 探测确认存在（duplicate_slide 的文件级复制路径默认走它；
//       quirk insertSlidesFromBase64 关掉可测降级，base64RequiresSuffixId 模拟只认 # 后缀 id 的宿主）
//   16. noGetSubstring 模拟没有 textRange.getSubstring 的旧宿主（applyEdit/set_text 的整框替换降级路径）；
//       noAddTextBox 模拟连 shapes.addTextBox 都没有的宿主（add_table_grid 的明确报错路径）
//   17. 文字级格式的 run/段落模型（close-capability-gaps）：font.italic / font.underline / paragraphFormat.bulletFormat.visible
//       按官方语义仿真——字符级格式挂在 run 上，段落级 bullet 挂在段落上，TextRange 读回全范围一致才有值、混排读 null；
//       getSubstring(...).font / .paragraphFormat 只写命中区间（bullet 是所在整段），区间外 run 对象级不动，
//       测试据此断言「只有命中子串变了」。quirk newTextFormatNoop 模拟冻结旧包宿主的静默 no-op：
//       这三个属性写了不报错也不生效（属性不在代理上，赋值成了普通 JS 属性），只有读回才露馅。
//   18. 样式写入（fill/line/整框 font）的调用不报错、sync 提交整批取消且不落盘（quirk styleWriteSyncFails
//       可打开）——「写入收下、提交被拒」的机器，recolor_slide 修复前在这种机器上谎报「已改 N 处」（B3）
const REAL_HOST_QUIRKS = {
  batchSyncAcrossShapes: true,
  opaqueGroups: true,
  noDuplicate: true,
  noUngroup: true,
  noGetItem: true,
  noFillSetType: true,
  undefinedPageSetup: true,
  lazyNotesSlide: true,
  shapeImageReplace: false,
  tableCellFill: false,
  chartValuesReadable: false,
  noAddImage: false,
  noAddTextBox: false,
  noGetSubstring: false,
  slideBackgroundApi: false,
  tableStructureApi: false,
  pictureCrop: false,
  pictureTransparency: false,
  insertSlidesFromBase64: true,   // 真机 16.112 探测确认存在，默认开；关掉可测 duplicate_slide 的借版降级
  base64RequiresSuffixId: false,  // 打开后 sourceSlideIds 只认 # 后缀，模拟「完整 id 被拒」的宿主
  noShapeTypeField: false,        // 模拟旧包宿主：shape 上 load 含 'type' 的字段列表会让整批 sync 取消
                                   //（触发 get_slide 带样式路径整批失败 → 降级读取，对应 ARCH D.6 场景）
  placeholderSyncFails: false,    // 模拟占位框读取在 sync 阶段才失败的宿主（getter 不抛、批量取消）：
                                   // 验证宿主前置过滤+受控降级，不把读取失败猜成「非占位框」或「空」
  noHyperlinkApi: false,          // 模拟没有 setHyperlink 的旧包宿主（方法不存在，调用即 TypeError）——
                                   // set_hyperlink 的探测摘牌路径用
  hyperlinkCallFails: false,      // 【unify-host-contracts 4】模拟 setHyperlink 方法存在、但本次调用抛 RichApi 异常
                                   //（Office 迟到就绪/受保护文档一类瞬态失败）：探测须记 'pending' 不摘牌
  noAdjustmentsApi: false,        // 模拟没有 Shape.adjustments 集合的宿主（PowerPointApi 1.10 缺失的
                                   // 「声明支持但实测没有」情形）——圆角半径写入失败须删残形并显式报错
  adjustmentsGetLies: false,      // 模拟 adjustments.get 读回值与写入不符的宿主（写后读回校验触发路径）
  deleteCommitFails: false,       // 模拟删除在 sync 提交阶段失败的宿主：验证失败清理「未确认」上报，
                                   // 不得宣称残形已删除
  addShapeSyncFails: false,       // 模拟 addShape 提交在 sync 阶段被 Office 拒绝（InvalidArgument 类，
                                   // 真机展览方案第 5 页「金句文字」同型）：验证错误回传带 code、参数
                                   // 回显与残留核对指引，不吞异常报成功
  childTextSyncFails: false,      // 【49-R2】模拟组合内文字成员读取在 sync 阶段失败的宿主：
                                   // 成员文字异常不得吞成空串，组合证据按不足处理（受限删除）
  newTextFormatNoop: false,       // 模拟冻结旧包宿主：font.italic / font.underline / bulletFormat.visible 写了不报错也不生效
                                   //（静默 no-op）——apply_layout 新格式字段的提交后读回核验靠它触发中文版本提示
  noBulletTypeApi: false,         // 【fix-real-machine-round2 2.1】模拟没有 BulletFormat.type 的宿主（PowerPointApi 1.10
                                   // 声明支持但代理上没有该属性）：赋值成了普通 JS 属性（静默 no-op），读回 undefined——
                                   // bulletType 的「声明支持但实测没有」路径（Mac 16.112 同型），中文版本提示靠它触发
  styleWriteSyncFails: false      // 【B3 show-me 交互审查】模拟样式写入在 sync 提交阶段被宿主整批取消的机器：
                                   // fill/line/整框 font 的写入调用不报错、sync 才拒绝、取消的批次一个字段都不落盘。
                                   // recolor_slide 修复前 changes.push 在 sync 之前，这台机器上会照报「已改 N 处」
};

function createFakeHost(spec = {}, quirks = {}) {
  const q = Object.assign({}, REAL_HOST_QUIRKS, quirks);
  const pending = [];          // 本次 sync 之前累积的 load 请求 { owner, member, type }
  const pendingDeletes = [];   // 【45-R2】delete 只排队，sync 才提交——真实 Office.js 的两段式语义，
                               // 以前即时置位 deleted 让「删残形」测试在未提交状态下误绿
  const log = [];              // 写操作流水，用于断言真的写进去了
  // 【fix-real-machine-round3】type 已随批 sync 成功的形状状态（WeakSet 挂 st，不进 JSON 序列化）。
  // 真机（16.112 实测）：type 是代理标量属性，load 进批且 sync 成功后才可读，未 load 直接读抛
  // 「属性"type"不可用」——以前 fake 读什么都答应，P5 视觉重排「写入成功、读回假失败」在 fake 上测不出来。
  const typeLoadedSts = new WeakSet();

  // 哪些成员在哪些形状类型上不存在——请求它们就会让整批 sync 失败
  const UNSUPPORTED = {
    Group: ['textFrame', 'fill'],
    Picture: ['textFrame'],
    Line: ['textFrame', 'fill'],
    Connector: ['textFrame', 'fill'],
    Chart: ['textFrame'],
    Table: ['textFrame'] // 表格形状顶层没有 textFrame，文字/字体在单元格里（shape.table）
  };
  function markLoad(ownerId, member, type) { if (ownerId) pending.push({ owner: ownerId, member: member || 'self', type: type || '' }); }

  // ---- 文字级格式的 run/段落模型（quirk 17）----
  // 官方语义：字符级格式（font.*）挂在 run 上，段落级格式（bulletFormat.visible）挂在段落上；
  // TextRange 读回时全范围一致才有值，混排读回 null。子串写入只动命中区间，区间外 run 对象级不动。
  function fullTextOf(st) { return Array.isArray(st.runs) ? st.runs.map(r => String(r.text || '')).join('') : String(st.text || ''); }
  function ensureRuns(st) {
    if (!Array.isArray(st.runs)) {
      st.runs = [{ text: String(st.text || ''), font: st.font || '', fontSize: st.fontSize, fontColor: st.fontColor || '', bold: Boolean(st.bold), italic: st.italic === true, underline: st.underline || 'None' }];
    }
    return st.runs;
  }
  // 把 [from, from+len) 区间内的 run 在边界处切开并打上 patch；不相交的 run 原对象保留
  function writeRunFormat(st, from, len, patch) {
    const out = [];
    let offset = 0;
    for (const run of ensureRuns(st)) {
      const t = String(run.text || '');
      const rStart = offset;
      const rEnd = rStart + t.length;
      offset = rEnd;
      const a = Math.max(from, rStart);
      const b = Math.min(from + len, rEnd);
      if (a >= b) { out.push(run); continue; }
      const head = t.slice(0, a - rStart);
      const mid = t.slice(a - rStart, b - rStart);
      const tail = t.slice(b - rStart);
      if (head) out.push(Object.assign({}, run, { text: head }));
      out.push(Object.assign({}, run, patch, { text: mid }));
      if (tail) out.push(Object.assign({}, run, { text: tail }));
    }
    st.runs = out.filter(r => r.text);
    st.text = fullTextOf(st);
  }
  // 区间内各 run 的某字段一致就返回该值，混排返回 null；from == null 表示整范围；没有 run 模型时读 st 自身
  function runConsensus(st, from, len, read) {
    if (!Array.isArray(st.runs)) return read(st);
    const values = [];
    let offset = 0;
    for (const run of st.runs) {
      const t = String(run.text || '');
      const rStart = offset;
      const rEnd = rStart + t.length;
      offset = rEnd;
      if (from == null || (rStart < from + len && rEnd > from)) values.push(read(run));
    }
    if (!values.length) return read(st);
    return values.every(v => v === values[0]) ? values[0] : null;
  }
  function paragraphRangesOf(st) {
    const ranges = [];
    let start = 0;
    for (const part of fullTextOf(st).split('\n')) { ranges.push({ start, end: start + part.length }); start += part.length + 1; }
    return ranges;
  }
  // 段落级 bullet 状态：st.paragraphBullets[i]；spec 给 bulletVisible:true 时作为全段初值
  function bulletVisArray(st) {
    const n = paragraphRangesOf(st).length;
    if (!Array.isArray(st.paragraphBullets) || st.paragraphBullets.length !== n) {
      const prev = Array.isArray(st.paragraphBullets) ? st.paragraphBullets : [];
      st.paragraphBullets = Array.from({ length: n }, (_, i) => (prev[i] == null ? st.bulletVisible === true : prev[i] === true));
    }
    return st.paragraphBullets;
  }
  // from == null 表示整范围；否则只取与 [from, from+len) 相交的段落下标（bullet 是所在整段生效）
  function paragraphIndexesIn(st, from, len) {
    const ranges = paragraphRangesOf(st);
    if (from == null) return ranges.map((_, i) => i);
    return ranges.map((r, i) => (r.start < from + len && r.end > from ? i : -1)).filter(i => i >= 0);
  }
  // 段落级 bullet 编号样式（【fix-real-machine-round2 2.1】BulletFormat.type，PowerPointApi 1.10）：
  // st.paragraphBulletTypes[i] ∈ 'bullet'|'number'|null（未设置读 null，与「混排/未定」同口径）
  function bulletTypeArray(st) {
    const n = paragraphRangesOf(st).length;
    if (!Array.isArray(st.paragraphBulletTypes) || st.paragraphBulletTypes.length !== n) {
      const prev = Array.isArray(st.paragraphBulletTypes) ? st.paragraphBulletTypes : [];
      st.paragraphBulletTypes = Array.from({ length: n }, (_, i) => (prev[i] == null ? (st.bulletType || null) : prev[i]));
    }
    return st.paragraphBulletTypes;
  }
  function makeBulletFormat(st, ownerId, from, len) {
    const fmt = {
      load() { markLoad(ownerId, 'textFrame', st.type); },
      get visible() {
        const arr = bulletVisArray(st);
        const values = paragraphIndexesIn(st, from, len).map(i => arr[i]);
        if (!values.length) return null;
        return values.every(v => v === values[0]) ? values[0] : null;
      },
      set visible(v) {
        if (q.newTextFormatNoop) return;   // 静默 no-op 宿主：写了不报错也不生效
        const arr = bulletVisArray(st);
        for (const i of paragraphIndexesIn(st, from, len)) arr[i] = Boolean(v);
        log.push({ op: 'bulletFormat.visible', shape: st.name, value: Boolean(v), from: from == null ? null : from, len: len == null ? null : len });
      }
    };
    // noBulletTypeApi：代理上没有 type 属性——宿主侧赋值成为普通 JS 属性（静默 no-op），
    // 读回 undefined；与真机「声明 1.10 支持但 BulletFormat.type 不存在」同型。
    // 注意不能用对象 spread 定义：spread 会对 getter 取值落成数据属性，访问器语义就丢了
    if (!q.noBulletTypeApi) {
      Object.defineProperty(fmt, 'type', {
        enumerable: true, configurable: true,
        get() {
          const arr = bulletTypeArray(st);
          const values = paragraphIndexesIn(st, from, len).map(i => arr[i]);
          if (!values.length) return null;
          return values.every(v => v === values[0]) ? values[0] : null;
        },
        set(v) {
          if (q.newTextFormatNoop) return;
          const arr = bulletTypeArray(st);
          for (const i of paragraphIndexesIn(st, from, len)) arr[i] = String(v);
          log.push({ op: 'bulletFormat.type', shape: st.name, value: String(v), from: from == null ? null : from, len: len == null ? null : len });
        }
      });
    }
    return fmt;
  }
  // 子串范围的字体代理：写入只动命中区间的 run（边界处切分），读回按区间内 run 的一致性给值/null
  function makeScopedFont(st, ownerId, from, len) {
    const write = (patch, op, value) => { writeRunFormat(st, from, len, patch); log.push({ op, shape: st.name, value, from, len }); };
    return {
      load() { markLoad(ownerId, 'font', st.type); },
      get name() { return runConsensus(st, from, len, x => x.font || ''); }, set name(v) { write({ font: v }, 'font.name', v); },
      get size() { return runConsensus(st, from, len, x => x.fontSize); }, set size(v) { write({ fontSize: v }, 'font.size', v); },
      get color() { return runConsensus(st, from, len, x => x.fontColor || ''); }, set color(v) { write({ fontColor: v }, 'font.color', v); },
      get bold() { return runConsensus(st, from, len, x => Boolean(x.bold)); }, set bold(v) { write({ bold: Boolean(v) }, 'font.bold', Boolean(v)); },
      get italic() { return runConsensus(st, from, len, x => x.italic === true); },
      set italic(v) { if (q.newTextFormatNoop) return; write({ italic: Boolean(v) }, 'font.italic', Boolean(v)); },
      get underline() { return runConsensus(st, from, len, x => x.underline || 'None'); },
      set underline(v) { if (q.newTextFormatNoop) return; write({ underline: String(v) }, 'font.underline', String(v)); }
    };
  }

  // setHyperlink 落库 + 返回 Hyperlink 代理。真机（vendor 16.0.20416 取证）：TextRange/Shape.setHyperlink
  // 返回 Hyperlink 对象，标量属性 address/screenTip/type，load 后可读（写后读回验证依赖它）。
  // start/length 为 null 表示形状级链接；数组坐标表示 TextRange 子串级。
  function recordHyperlink(st, opts, start, length) {
    // hyperlinkCallFails：方法在（typeof 是函数），但调用在宿主侧失败——瞬态失败，不是能力缺失
    if (q.hyperlinkCallFails) {
      const err = new Error('RichApi.Error: GeneralException — the host is not ready to accept this request');
      err.code = 'GeneralException';
      throw err;
    }
    const entry = {
      address: (opts && opts.address) || '',
      screenTip: (opts && opts.screenTip) || '',
      start: start == null ? null : start,
      length: length == null ? null : length
    };
    if (!Array.isArray(st.hyperlinks)) st.hyperlinks = [];
    st.hyperlinks.push(entry);
    log.push({ op: 'setHyperlink', shape: st.name, value: entry.address, start: entry.start, length: entry.length });
    return {
      load() { markLoad(st.id, 'hyperlink', st.type); },
      get address() { return entry.address; },
      get screenTip() { return entry.screenTip; },
      get type() { return entry.start == null ? 'shape' : 'textRange'; }
    };
  }

  function makeFont(st, ownerId) {
    // styleWriteSyncFails：写入调用被收下但不落盘，sync 阶段整批取消（B3 红测语义，见 quirk 注释）
    const doomedWrite = () => pending.push({ owner: ownerId, member: 'styleWriteCommit', type: 'StyleWrite' });
    return {
      load() { markLoad(ownerId, 'font', st.type); },
      get name() { return st.font; }, set name(v) { if (q.styleWriteSyncFails) { doomedWrite(); return; } st.font = v; log.push({ op: 'font.name', shape: st.name, value: v }); },
      // 混排字号的单元格/形状：真机 font.size 读回 null（不是 0，也不是抛错），写前必须先判
      get size() { return st.mixedFontSize ? null : st.fontSize; }, set size(v) { if (q.styleWriteSyncFails) { doomedWrite(); return; } st.fontSize = v; log.push({ op: 'font.size', shape: st.name, value: v }); },
      get color() { return st.fontColor; }, set color(v) { if (q.styleWriteSyncFails) { doomedWrite(); return; } st.fontColor = v; log.push({ op: 'font.color', shape: st.name, value: v }); },
      get bold() { return st.bold; }, set bold(v) { if (q.styleWriteSyncFails) { doomedWrite(); return; } st.bold = v; log.push({ op: 'font.bold', shape: st.name, value: v }); },
      // 整范围 italic/underline（quirk 17）：读回按 run 一致性（混排 null）；写入盖住全部 run。
      // newTextFormatNoop 模拟静默 no-op 宿主：不写、不报错、不记日志
      get italic() { return runConsensus(st, null, 0, x => x.italic === true); },
      set italic(v) {
        if (q.newTextFormatNoop) return;
        st.italic = Boolean(v);
        if (Array.isArray(st.runs)) st.runs.forEach(r => { r.italic = Boolean(v); });
        log.push({ op: 'font.italic', shape: st.name, value: Boolean(v) });
      },
      get underline() { return runConsensus(st, null, 0, x => x.underline || 'None'); },
      set underline(v) {
        if (q.newTextFormatNoop) return;
        st.underline = String(v);
        if (Array.isArray(st.runs)) st.runs.forEach(r => { r.underline = String(v); });
        log.push({ op: 'font.underline', shape: st.name, value: String(v) });
      }
    };
  }

  function makeShape(init) {
    // 状态要跨「重建」持久：context.presentation.slides 每次访问都重新 map(makeShape)，
    // 若 st 每次新建，几何写入（left/top 等标量）在下一次读取时就丢了——真机上写入是持久的。
    // 【49-R2 保真度】st 直接用 spec 自身（默认值补在 spec 上，__st 不可枚举）：以前 Object.assign
    // 拷贝缓存，spec 在出卡后被改（Astra 49 探针：组合成员文字变了）对代理读取不可见——真实文档
    // 里成员被改一定可见，fake 必须同语义。JSON.stringify 深拷贝因 __st 不可枚举不受循环引用影响。
    if (!init.__st) {
      init.id = init.id || init.name;
      init.name = init.name || 'shape';
      if (init.type === undefined) init.type = 'Shape';
      if (init.left === undefined) init.left = 0;
      if (init.top === undefined) init.top = 0;
      if (init.width === undefined) init.width = 100;
      if (init.height === undefined) init.height = 50;
      if (init.fill === undefined) init.fill = '';
      if (init.line === undefined) init.line = '';
      if (init.lineVisible === undefined) init.lineVisible = true;
      if (init.lineWeight === undefined) init.lineWeight = 1;
      if (init.font === undefined) init.font = '';
      if (init.fontSize === undefined) init.fontSize = 18;
      if (init.fontColor === undefined) init.fontColor = '';
      if (init.bold === undefined) init.bold = false;
      if (init.text === undefined) init.text = '';
      if (init.children === undefined) init.children = [];
      Object.defineProperty(init, '__st', { value: init, enumerable: false, configurable: true, writable: true });
    }
    const isNew = init.__st === init && init.__initialized !== true;
    const st = init.__st;
    // run 级混排模型：spec 给 runs: [{ text, font, fontSize, fontColor, bold }] 时，
    // st.text 由 runs 拼出；getSubstring(...).insertText 只动命中区间的 run，区间外的 run 原样保留
    if (isNew && Array.isArray(st.runs)) st.text = st.runs.map(r => String(r.text || '')).join('');
    st.__initialized = true;
    const ownerId = st.id;
    const shape = {
      _st: st,
      load(fields) {
        markLoad(ownerId, 'self', st.type);
        // type 进批：sync 成功后才解锁 get type（真机 PropertyNotLoaded 语义，见 typeLoadedSts 注释）
        if (String(fields || '').split(',').map(s => s.trim()).includes('type')) pending.push({ owner: ownerId, member: 'typeLoad', st });
        // noShapeTypeField：旧包宿主请求 type 字段整批取消，与 UNSUPPORTED 同机制生效
        if (q.noShapeTypeField && String(fields || '').split(',').map(s => s.trim()).includes('type')) {
          pending.push({ owner: ownerId, member: 'typeField', type: 'UnsupportedTypeHost', fields: String(fields || '') });
        }
      },
      get id() { return st.id; }, get name() { return st.name; }, set name(v) { st.name = v; log.push({ op: 'name', shape: st.id, value: v }); },
      get type() {
        if (!typeLoadedSts.has(st)) throw new Error('属性"type"不可用。读取属性的值之前，请先对包含对象调用 load 方法，再对关联的请求上下文调用 "context.sync()"');
        return st.type;
      },
      get left() { return st.left; }, set left(v) { st.left = v; log.push({ op: 'geometry', shape: st.name, value: v }); },
      get top() { return st.top; }, set top(v) { st.top = v; log.push({ op: 'geometry', shape: st.name, value: v }); },
      get width() { return st.width; }, set width(v) { st.width = v; log.push({ op: 'geometry', shape: st.name, value: v }); },
      get height() { return st.height; }, set height(v) { st.height = v; log.push({ op: 'geometry', shape: st.name, value: v }); },
      get zOrderPosition() { return st.z || 0; },
      setZOrder(v) { log.push({ op: 'setZOrder', shape: st.name, value: v }); },
      delete() {
        log.push({ op: 'delete', shape: st.name });
        // 【45-R2】只排队不提交：真实契约里 delete 在 context.sync() 时才落盘
        pendingDeletes.push(st);
      },
      setHyperlink(opts) { return recordHyperlink(st, opts, null, null); },
      // Shape.adjustments（PowerPointApi 1.10）：调整值是 min(w,h) 的相对比例，roundRect 上限 0.5，
      // 绝对半径 = value × min(w,h)。quirk noAdjustmentsApi 模拟连该集合都没有的旧宿主——
      // 43 的「声明支持≠真支持」第二种情形（Mac 式实测缺失）就靠它触发删残形报错路径。
      ...(q.noAdjustmentsApi ? {} : {
        adjustments: {
          set(index, value) {
            st.adjustments = st.adjustments || [];
            st.adjustments[index] = value;
            log.push({ op: 'adjustments.set', shape: st.name, index, value });
          },
          // adjustmentsGetLies：模拟「写入成功但读回值与请求不符」的宿主（写后读回校验的触发路径）
          get(index) {
            const value = (st.adjustments || [])[index];
            return { value: q.adjustmentsGetLies && Number.isFinite(value) ? value * 0.5 : value };
          }
        }
      }),
      fill: {
        load() { markLoad(ownerId, 'fill', st.type); },
        get type() { return st.fill ? 'Solid' : 'NoFill'; },
        get foregroundColor() { return st.fill; },
        set foregroundColor(v) { if (q.styleWriteSyncFails) { pending.push({ owner: ownerId, member: 'styleWriteCommit', type: 'StyleWrite' }); return; } st.fill = v; log.push({ op: 'fill.foregroundColor', shape: st.name, value: v }); },
        setSolidColor(v) { if (q.styleWriteSyncFails) { pending.push({ owner: ownerId, member: 'styleWriteCommit', type: 'StyleWrite' }); return; } st.fill = v; log.push({ op: 'fill.setSolidColor', shape: st.name, value: v }); },
        clear() { st.fill = ''; log.push({ op: 'fill.clear', shape: st.name }); },
        get transparency() { return st.fillTransparency || 0; }, set transparency(v) { st.fillTransparency = v; }
      },
      lineFormat: {
        load() { markLoad(ownerId, 'lineFormat', st.type); },
        get visible() { return st.lineVisible; }, set visible(v) { st.lineVisible = v; },
        get color() { return st.line; }, set color(v) { if (q.styleWriteSyncFails) { pending.push({ owner: ownerId, member: 'styleWriteCommit', type: 'StyleWrite' }); return; } st.line = v; log.push({ op: 'line.color', shape: st.name, value: v }); },
        get transparency() { return 0; },
        get weight() { return st.lineWeight; }, set weight(v) { st.lineWeight = v; }
      },
      // 占位框身份（37/39-R1/40）：对齐官方契约——非占位框访问 placeholderFormat 抛
      // GeneralException（不是返回 None）；containedType 为占位框包含的 ShapeType，
      // 空占位框返回 null（PowerPointApi 1.8），异常宿主可能读回 undefined（缺失≠空）。
      // 同一形状缓存同一代理对象，探针可覆写；load 进 log 供「非占位框零占位请求」计数。
      get placeholderFormat() {
        if (!st.placeholder) { const e = new Error('GeneralException: Shape is not a placeholder'); e.code = 'GeneralException'; throw e; }
        if (!st.__pf) {
          st.__pf = {
            load(fields) { markLoad(ownerId, 'placeholderFormat', st.type); log.push({ op: 'placeholderFormat.load', shape: ownerId, fields: String(fields || '') }); },
            get type() { return st.placeholder; },
            get containedType() {
              if (st.containedTypeThrows) throw new Error('RichApi.Error: containedType unreadable on this host');
              if (st.containedTypeUndefined) return undefined;
              return st.containedType !== undefined ? st.containedType : null;
            }
          };
        }
        return st.__pf;
      },
      textFrame: {
        load() {
          markLoad(ownerId, 'textFrame', st.type);
          // 【49-R2】组合内文字成员读取在 sync 阶段失败的宿主（getter 不抛、批量取消）
          if (q.childTextSyncFails && st.isGroupChild) pending.push({ owner: ownerId, member: 'childText', type: 'ChildText' });
        },
        get hasText() { return Boolean(st.text); },
        get leftMargin() { return 7; }, get rightMargin() { return 7; },
        get topMargin() { return 4; }, get bottomMargin() { return 4; },
        get verticalAlignment() { return 'Middle'; }, get wordWrap() { return true; },
        get autoSizeSetting() { return 'AutoSizeNone'; },
        textRange: {
          load() {
            markLoad(ownerId, 'textFrame', st.type);
            // 【49-R2】组合内文字成员读取在 sync 阶段失败的宿主（getter 不抛、批量取消）：
            // currentGroupMembers 走的是 textRange.load，注入必须在这一层
            if (q.childTextSyncFails && st.isGroupChild) pending.push({ owner: ownerId, member: 'childText', type: 'ChildText' });
          },
          setHyperlink(opts) { return recordHyperlink(st, opts, 0, String(Array.isArray(st.runs) ? st.runs.map(r => String(r.text || '')).join('') : st.text || '').length); },
          get text() { return Array.isArray(st.runs) ? st.runs.map(r => String(r.text || '')).join('') : st.text; },
          set text(v) {
            st.text = v;
            // 整框直接赋值：真机会把混排 run 抹平成单一格式（近似宿主默认）。保混排要走 getSubstring 子串路径
            if (Array.isArray(st.runs)) st.runs = [{ text: v, font: '', fontSize: 18, fontColor: '', bold: false }];
            log.push({ op: 'text', shape: st.name, value: v });
          },
          // getSubstring(start, length).insertText(text, 'Replace')：只替换命中区间，
          // 新文本继承被替换段第一个字符所在 run 的格式；区间外的 run 原样保留（对象级不动）。
          // quirk noGetSubstring 模拟没有这个 API 的旧宿主。
          getSubstring(start, length) {
            if (q.noGetSubstring) throw new Error('RichApi.Error: getSubstring is not supported on this host');
            const from = Math.max(0, Number(start) || 0);
            const len = Math.max(0, Number(length) || 0);
            return {
              load() { markLoad(ownerId, 'textFrame', st.type); },
              get text() {
                const full = Array.isArray(st.runs) ? st.runs.map(r => String(r.text || '')).join('') : String(st.text || '');
                return full.slice(from, from + len);
              },
              setHyperlink(opts) { return recordHyperlink(st, opts, from, len); },
              insertText(text) {
                const insert = String(text == null ? '' : text);
                log.push({ op: 'insertText', shape: st.name, value: insert, from, len });
                if (!Array.isArray(st.runs)) {
                  const full = String(st.text || '');
                  st.text = full.slice(0, from) + insert + full.slice(from + len);
                  return;
                }
                const out = [];
                let offset = 0;
                let inheritFmt = null;
                let insertedRun = false;
                for (const run of st.runs) {
                  const t = String(run.text || '');
                  const rStart = offset;
                  const rEnd = rStart + t.length;
                  offset = rEnd;
                  const overlapped = rStart < from + len && rEnd > from;
                  if (overlapped && !inheritFmt) inheritFmt = run;
                  const head = t.slice(0, Math.min(Math.max(from - rStart, 0), t.length));
                  const tail = t.slice(Math.min(Math.max(from + len - rStart, 0), t.length));
                  if (head) out.push(Object.assign({}, run, { text: head }));
                  if (overlapped && !insertedRun) { out.push(Object.assign({}, inheritFmt, { text: insert })); insertedRun = true; }
                  if (tail) out.push(Object.assign({}, run, { text: tail }));
                }
                if (!insertedRun) out.push(Object.assign({}, inheritFmt || st.runs[st.runs.length - 1] || {}, { text: insert }));
                st.runs = out.filter(r => r.text);
                st.text = st.runs.map(r => r.text).join('');
              },
              // 子串级格式（quirk 17）：font 只写命中区间的 run，bullet 是命中子串所在的整段
              get font() { return makeScopedFont(st, ownerId, from, len); },
              paragraphFormat: {
                load() { markLoad(ownerId, 'textFrame', st.type); },
                get horizontalAlignment() { return 'Left'; },
                get bulletFormat() { return makeBulletFormat(st, ownerId, from, len); }
              }
            };
          },
          get font() { return makeFont(st, ownerId); },
          paragraphFormat: {
            load() { markLoad(ownerId, 'textFrame', st.type); },
            get horizontalAlignment() { return 'Left'; },
            // 整范围 bullet（quirk 17）：写入盖住全部段落，读回全段一致才有值
            get bulletFormat() { return makeBulletFormat(st, ownerId, null, null); }
          }
        }
      }
    };
    // 组合：真机上子形状读不到
    if (st.type === 'Group') {
      shape.shapes = {
        load() { markLoad(ownerId, 'childShapes', st.type); },
        get items() {
          if (q.opaqueGroups) return [];
          // 【49-R2】标记组合子成员：childTextSyncFails 只对组内文字成员生效
          return st.children.map(spec => { const c = makeShape(spec); c._st.isGroupChild = true; return c; });
        }
      };
      if (!q.noUngroup) shape.ungroup = () => log.push({ op: 'ungroup', shape: st.name });
    }
    // 图片：真机没有原位替换 API；quirk shapeImageReplace 模拟「未来宿主补齐了 image.replace」
    if (st.type === 'Picture' && q.shapeImageReplace) {
      shape.image = {
        load() { markLoad(ownerId, 'image', st.type); },
        replace(base64) { st.imageData = base64; log.push({ op: 'image.replace', shape: st.name }); }
      };
    }
    // 图片裁剪（cropFormat）：真机未验证存在，默认没有；quirk pictureCrop 模拟「宿主补齐了裁剪 API」
    if (st.type === 'Picture' && q.pictureCrop) {
      shape.cropFormat = {
        load() { markLoad(ownerId, 'cropFormat', st.type); },
        get cropTop() { return st.cropTop || 0; }, set cropTop(v) { st.cropTop = v; log.push({ op: 'crop.cropTop', shape: st.name, value: v }); },
        get cropBottom() { return st.cropBottom || 0; }, set cropBottom(v) { st.cropBottom = v; log.push({ op: 'crop.cropBottom', shape: st.name, value: v }); },
        get cropLeft() { return st.cropLeft || 0; }, set cropLeft(v) { st.cropLeft = v; log.push({ op: 'crop.cropLeft', shape: st.name, value: v }); },
        get cropRight() { return st.cropRight || 0; }, set cropRight(v) { st.cropRight = v; log.push({ op: 'crop.cropRight', shape: st.name, value: v }); }
      };
    }
    // 图片透明度（pictureFormat.transparency）：真机未验证可写，默认没有；quirk pictureTransparency 模拟可写宿主
    if (st.type === 'Picture' && q.pictureTransparency) {
      shape.pictureFormat = {
        load() { markLoad(ownerId, 'pictureFormat', st.type); },
        get transparency() { return st.picTransparency || 0; },
        set transparency(v) { st.picTransparency = v; log.push({ op: 'picture.transparency', shape: st.name, value: v }); }
      };
    }
    // 图表：chartType 和 series.name 可读；series.values 真机读不到（Chart API 薄），
    // 读取时走 load+sync 失败机制（markLoad 'chartSeriesValues' → 整批取消），不是 getter 抛错
    if (st.type === 'Chart' && st.chart) {
      const chartSt = st.chart;
      shape.chart = {
        load() { markLoad(ownerId, 'chart', st.type); },
        get chartType() { return chartSt.chartType || ''; },
        series: {
          load() { markLoad(ownerId, 'chart', st.type); },
          get items() {
            return (chartSt.series || []).map(s => ({
              load(fields) { markLoad(ownerId, String(fields || '').includes('values') ? 'chartSeriesValues' : 'chartSeries', st.type); },
              get name() { return s.name; },
              get values() { return s.values; }
            }));
          }
        }
      };
    }
    // 表格：单元格状态直接活在 spec 的数组里，多次 getCell 拿到的是同一份数据（写入才可见）
    if (st.type === 'Table' && Array.isArray(st.table)) {
      const cellStates = st.table;
      const makeCell = (r, c) => {
        const cellSt = cellStates[r][c];
        if (cellSt.text == null) cellSt.text = '';
        if (cellSt.font == null) cellSt.font = '';
        if (cellSt.fontSize == null) cellSt.fontSize = 18;
        if (cellSt.fontColor == null) cellSt.fontColor = '';
        cellSt.name = `${st.name}[${r},${c}]`;
        cellSt.type = 'TableCell'; // 单元格的 textFrame 是可读的，不参与 UNSUPPORTED 判定
        const cell = {
          _st: cellSt,
          textFrame: {
            load() { markLoad(cellSt.name, 'textFrame', cellSt.type); },
            textRange: {
              load() { markLoad(cellSt.name, 'textFrame', cellSt.type); },
              get text() { return cellSt.text; },
              set text(v) { cellSt.text = v; log.push({ op: 'text', shape: cellSt.name, value: v }); },
              get font() { return makeFont(cellSt, cellSt.name); }
            }
          }
        };
        // 单元格底色（TableCell.fill）：真机未验证，默认按刁难不可写
        if (q.tableCellFill) {
          cell.fill = {
            load() { markLoad(cellSt.name, 'fill', cellSt.type); },
            setSolidColor(v) { cellSt.fill = v; log.push({ op: 'cell.fill', shape: cellSt.name, value: v }); }
          };
        }
        return cell;
      };
      shape.table = {
        load() { markLoad(ownerId, 'table', st.type); },
        get rowCount() { return cellStates.length; },
        get columnCount() { return cellStates.length ? cellStates[0].length : 0; },
        getCell(r, c) { return makeCell(r, c); }
      };
      // 行列增删：真机未验证存在，默认没有；quirk tableStructureApi 模拟「宿主补齐了结构操作」
      if (q.tableStructureApi) {
        shape.table.addRow = pos => { const at = pos == null ? cellStates.length : pos; cellStates.splice(at, 0, Array.from({ length: shape.table.columnCount }, () => ({}))); log.push({ op: 'table.addRow', shape: st.name, value: at }); };
        shape.table.deleteRow = pos => { cellStates.splice(pos, 1); log.push({ op: 'table.deleteRow', shape: st.name, value: pos }); };
        shape.table.addColumn = pos => { const at = pos == null ? (cellStates[0] || []).length : pos; for (const row of cellStates) row.splice(at, 0, {}); log.push({ op: 'table.addColumn', shape: st.name, value: at }); };
        shape.table.deleteColumn = pos => { for (const row of cellStates) row.splice(pos, 1); log.push({ op: 'table.deleteColumn', shape: st.name, value: pos }); };
      }
    }
    // noHyperlinkApi：模拟没有 setHyperlink 的旧包宿主——方法不存在（调用即 TypeError），
    // 与真机上冻结在 2022 版的 vendor 包行为一致（set_hyperlink 的摘牌探测路径）
    if (q.noHyperlinkApi) {
      delete shape.setHyperlink;
      if (shape.textFrame && shape.textFrame.textRange) delete shape.textFrame.textRange.setHyperlink;
      if (shape.textFrame && shape.textFrame.textRange && shape.textFrame.textRange.getSubstring) {
        const raw = shape.textFrame.textRange.getSubstring;
        shape.textFrame.textRange.getSubstring = function (s, l) {
          const sub = raw(s, l);
          if (sub && 'setHyperlink' in sub) delete sub.setHyperlink;
          return sub;
        };
      }
    }
    return shape;
  }

  // 真机 slide.id 形如 "2147481232#588170797"（会话 id#页 id），fake 对齐这个格式：
  // duplicate_slide 的文件级复制路径要靠 # 后缀做 id 格式自愈，格式太假就测不到那条分支。
  let slideSerial = 0;
  let shapeSerial = 0; // addGeometricShape 新建形状的唯一 id
  const nextSlideId = () => '2147483648#' + (588170797 + slideSerial++);
  const slideStates = (spec.slides || []).map(s => Object.assign({ id: nextSlideId(), hasNotes: false, notes: '' }, s));
  // 每次 makeSlide 建出的形状代理数组都登记在这里：删除提交后要把代理从「当前 run 已取到的 items」里移走
  //（真实 Office.js：sync 后集合刷新，被删形状不再出现），不能只在 spec 层删
  const liveShapeLists = new Set();

  function makeSlide(state, index) {
    const shapes = (state.shapes || []).map(makeShape);
    liveShapeLists.add(shapes);
    const slide = {
      _state: state,
      load() { markLoad('slide-' + index); },
      get id() { return state.id; },
      get index() { return index; },
      shapes: {
        load() { markLoad('slide-' + index); },
        get items() { return shapes; },
        // addTextBox 是 PowerPointApi 1.4 的 API；quirk noAddTextBox 模拟连它都没有的宿主。
        // 与 addGeometricShape 同款：按 opts 几何建形、登记进页面集合（真机上新文本框随后可被 get_slide 读到，
        // 写后读回也要能按请求的几何/文本核验），不再只返回一个游离代理
        ...(q.noAddTextBox ? {} : { addTextBox(text, opts) {
          log.push({ op: 'addTextBox', text, opts });
          const spec = {
            id: 'new-textbox-' + (++shapeSerial), name: 'new-textbox', type: 'TextBox', text,
            left: opts && opts.left, top: opts && opts.top, width: opts && opts.width, height: opts && opts.height
          };
          const sh = makeShape(spec);
          // 【47.4】与 addGeometricShape 同款提交失败注入（「金句文字」场景走的是 textBox）
          if (q.addShapeSyncFails) pending.push({ owner: sh.id, member: 'addCommit', type: 'AddCommit' });
          (state.shapes = state.shapes || []).push(spec);
          shapes.push(sh);
          return sh;
        } }),
        addGeometricShape(type, opts) {
          log.push({ op: 'addGeometricShape', type, opts });
          // 注册进页面形状列表：set_slide_background 的复用路径要能在下一次调用里按名字找到旧矩形。
          // 【45-R1】真实契约：Shape.type 对 addGeometricShape 建的形状只返回大类 GeometricShape；
          // RoundRectangle/Round2SameRectangle 等是 addGeometricShape 的 GeometricShapeType 入参，
          // 不是 type 的返回值——preset 单独保存供渲染器/断言用，读取契约不得把它当 type。
          const spec = {
            id: 'new-shape-' + (++shapeSerial), name: 'new-shape', type: 'GeometricShape', preset: String(type || ''),
            left: opts && opts.left, top: opts && opts.top, width: opts && opts.width, height: opts && opts.height
          };
          // 【47.4】addShape 提交失败的宿主：形状代理已建（与真实一样，创建发生在 sync 前），
          // 但 sync 阶段被 Office 拒绝——错误必须带 code 回传，不能吞掉报成功
          if (q.addShapeSyncFails) pending.push({ owner: spec.id, member: 'addCommit', type: 'AddCommit' });
          const sh = makeShape(spec);
          (state.shapes = state.shapes || []).push(spec);
          shapes.push(sh);
          return sh;
        },
        // addImage 是 PowerPointApi 1.8 的 API；quirk noAddImage 模拟连它都没有的宿主
        ...(q.noAddImage ? {} : {
          addImage(base64, opts) {
            log.push({ op: 'addImage', opts });
            const pic = makeShape({ id: 'new-picture', name: 'new-picture', type: 'Picture', left: opts && opts.left, top: opts && opts.top, width: opts && opts.width, height: opts && opts.height });
            pic._st.imageData = base64;
            return pic;
          }
        })
      },
      delete() { log.push({ op: 'deleteSlide', index }); },
      getImageAsBase64() { return { value: 'ZmFrZS1wbmc=' }; },
      set customData(v) { state.customData = v; }
    };
    if (!q.noDuplicate) slide.duplicate = () => log.push({ op: 'duplicate', index });
    // 页面背景写 API：真机未验证存在，默认没有；quirk slideBackgroundApi 模拟「宿主补齐了 slide.background」
    if (q.slideBackgroundApi) {
      slide.background = { fill: { load() { markLoad('slide-' + index); }, setSolidColor(v) { state.background = v; log.push({ op: 'background.fill', index, value: v }); } } };
    }
    Object.defineProperty(slide, 'notesSlide', {
      get() {
        if (q.lazyNotesSlide && !state.hasNotes) throw new Error('The notes slide has not been created yet.');
        return { shapes: { load() { markLoad('notes-' + index); }, get items() { return [makeShape({ name: 'notes-body', text: state.notes, placeholder: 'Body' })]; } } };
      }
    });
    return slide;
  }

  const context = {
    _log: log,
    presentation: {
      get slides() {
        const built = slideStates.map(makeSlide);
        return {
          load() { markLoad('collection'); },
          get items() { return built; },
          add(options) { log.push({ op: 'addSlide', options }); slideStates.push({ id: nextSlideId(), shapes: [] }); },
          ...(q.noGetItem ? {} : { getItem: i => built[i] })
        };
      },
      pageSetup: {
        load() { markLoad('pageSetup'); },
        get slideWidth() { return q.undefinedPageSetup ? undefined : 960; },
        get slideHeight() { return q.undefinedPageSetup ? undefined : 540; }
      },
      slideMasters: {
        load() { markLoad('masters'); },
        get items() {
          return (spec.masters || []).map(m => ({
            load() { markLoad('masters'); },
            get id() { return m.id; }, get name() { return m.name; },
            layouts: { load() { markLoad('masters'); }, get items() { return (m.layouts || []).map(l => ({ load() { markLoad('masters'); }, get id() { return l.id; }, get name() { return l.name; } })); } }
          }));
        }
      },
      getSelectedSlides() { return { load() { markLoad('collection'); }, get items() { return [makeSlide(slideStates[0], 0)]; } }; },
      setSelectedSlides() { log.push({ op: 'setSelectedSlides' }); },
      // 选中文本范围（vendor 取证：Presentation.getSelectedTextRange 门控 1.5，返回 TextRange）。
      // spec.selection = { slideIndex, shapeId, start, length } 指定假选中；不给则视为无选中（空文本）
      getSelectedTextRange() {
        log.push({ op: 'getSelectedTextRange' });
        const sel = spec.selection;
        const slide = sel && slideStates[sel.slideIndex || 0];
        const shapeSpec = slide && (slide.shapes || []).find(s => s.id === sel.shapeId);
        if (!shapeSpec) {
          return { load() { markLoad('selection'); }, get text() { return ''; }, setHyperlink() { throw new Error('no selection'); } };
        }
        // 状态统一走 __st（与 makeShape 同语义），写入才能持久并反映到 slideStates
        makeShape(shapeSpec);
        const st = shapeSpec.__st;
        const from = Math.max(0, sel.start || 0);
        const len = Math.max(0, sel.length || 0);
        if (q.noHyperlinkApi) {
          // 旧包宿主：选中范围的 text 可读，但代理上根本没有 setHyperlink 方法（与 shape/textRange 同款：
          // 「API 面不存在」，探测按 typeof 判 false 摘牌；不是「方法存在但调用抛错」的瞬态失败）
          const full0 = Array.isArray(st.runs) ? st.runs.map(r => String(r.text || '')).join('') : String(st.text || '');
          return { load() { markLoad(st.id, 'textFrame', st.type); }, get text() { return full0.slice(from, from + len); } };
        }
        return {
          load() { markLoad(st.id, 'textFrame', st.type); },
          get text() {
            const full = Array.isArray(st.runs) ? st.runs.map(r => String(r.text || '')).join('') : String(st.text || '');
            return full.slice(from, from + len);
          },
          setHyperlink(opts) { return recordHyperlink(st, opts, from, len); }
        };
      },
      // 文件级插入：把 base64 源文件中 sourceSlideIds 指定的页插到 targetSlideId 之后。
      // duplicate_slide 在没有 slide.duplicate() 的宿主上靠它复制页面（源文件就是当前文件）。
      ...(q.insertSlidesFromBase64 ? {
        insertSlidesFromBase64(base64, options) {
          const opts = options || {};
          const sourceId = String((opts.sourceSlideIds || [])[0] || '');
          const targetId = String(opts.targetSlideId || '');
          log.push({ op: 'insertSlidesFromBase64', sourceId, targetSlideId: targetId });
          // id 格式刁难：quirk base64RequiresSuffixId 打开时只认 # 后缀，完整 id 会被拒
          const matches = state => q.base64RequiresSuffixId
            ? String(state.id).split('#').pop() === sourceId
            : String(state.id) === sourceId || String(state.id).split('#').pop() === sourceId;
          const sourceIndex = slideStates.findIndex(matches);
          if (sourceIndex < 0) {
            const err = new Error(`RichApi.Error: insertSlidesFromBase64 找不到 sourceSlideIds 指定的页（"${sourceId}"）`);
            err.code = 'InvalidArgument';
            throw err;
          }
          const targetIndex = slideStates.findIndex(state => String(state.id) === targetId);
          const insertAt = (targetIndex >= 0 ? targetIndex : sourceIndex) + 1;
          // 复制源页 spec（含 __st 形状状态）生成新页：文件级插入的净效果就是源页的完整副本
          const copy = JSON.parse(JSON.stringify(slideStates[sourceIndex]));
          copy.id = nextSlideId();
          slideStates.splice(insertAt, 0, copy);
        }
      } : {})
    },
    async sync() {
      const batch = pending.splice(0, pending.length);
      const deletes = pendingDeletes.splice(0, pendingDeletes.length);
      // type load 解锁只在批次真正通过后发生：整批取消时请求过的属性同样读不到（真机语义）
      const commitTypeLoads = () => { for (const r of batch) if (r.member === 'typeLoad' && r.st) typeLoadedSts.add(r.st); };
      if (!q.batchSyncAcrossShapes) { commitTypeLoads(); for (const st of deletes) commitDelete(st); return; }
      const bad = batch.filter(r => (UNSUPPORTED[r.type] || []).includes(r.member)
        // Chart series.values 的可读性由 quirk 动态决定：不可读时 load('values') 让整批取消
        || (r.member === 'chartSeriesValues' && !q.chartValuesReadable)
        // 占位框读取在 sync 阶段失败的宿主（40-R1-c）：getter 不抛，批量取消
        || (r.member === 'placeholderFormat' && q.placeholderSyncFails)
        || r.member === 'typeField'
        // 【49-R2】addShape 提交被 Office 拒绝（InvalidArgument 类）
        || (r.member === 'addCommit' && q.addShapeSyncFails)
        // 【49-R2】组合内文字成员读取失败
        || (r.member === 'childText' && q.childTextSyncFails)
        // 【B3】样式写入的 sync 提交被宿主拒绝（写入调用不报错、提交才失败）
        || (r.member === 'styleWriteCommit' && q.styleWriteSyncFails));
      if (bad.length) {
        const err = new Error(bad.some(b => b.member === 'addCommit')
          ? 'InvalidArgument'
          : 'RichApi.Error: The batch was canceled because one of its requests failed. [' + bad.map(b => `${b.owner}.${b.member}`).join(', ') + ']');
        err.code = bad.some(b => b.member === 'addCommit') ? 'InvalidArgument' : 'GeneralException';
        throw err;   // 整批取消：排队中的删除同样不提交，type 解锁同样不发生
      }
      commitTypeLoads();
      // 【45-R2】load 批通过后才提交删除；deleteCommitFails 模拟删除提交失败的宿主（排队项丢弃不生效）
      if (q.deleteCommitFails && deletes.length) {
        const err = new Error('RichApi.Error: delete commit failed on this host');
        err.code = 'GeneralException';
        throw err;
      }
      for (const st of deletes) commitDelete(st);
    }
  };
  // 删除提交（真实契约）：形状从集合里消失——同一 run 里已取到的 items 不再有它，下一次 run 重建代理也不再有它。
  // 此前只置位 st.deleted 留下「幽灵形状」，写后读回（unify-host-contracts：deleteShape 提交后重读本页确认目标已不在）
  // 在 fake 上永远判成「删除未生效」，与真机语义相反。st.deleted 仍置位，供「删除是否已提交」的断言使用。
  function commitDelete(st) {
    st.deleted = true;
    const removeFrom = list => {
      if (!Array.isArray(list)) return;
      const at = list.indexOf(st);
      if (at >= 0) list.splice(at, 1);
      for (const item of list) if (item && Array.isArray(item.children)) removeFrom(item.children);
    };
    for (const s of slideStates) removeFrom(s.shapes);
    for (const list of liveShapeLists) {
      const at = list.findIndex(p => p && p._st === st);
      if (at >= 0) list.splice(at, 1);
    }
  }
  // Office.context.document 的 getFileAsync mock：duplicate_slide 的文件级复制路径要把当前 pptx
  // 读成 base64。固定返回 1 个 slice（仿真文件很小），slice.data 是一段假 base64。
  const officeDocument = {
    getFileAsync(fileType, options, callback) {
      log.push({ op: 'getFileAsync', fileType, sliceSize: options && options.sliceSize });
      const file = {
        sliceCount: 1,
        getSliceAsync(i, cb) { cb({ status: 'succeeded', value: { data: i === 0 ? 'ZmFrZS1wcHR4LWJhc2U2NA==' : '' } }); },
        closeAsync(cb) { log.push({ op: 'file.closeAsync' }); if (cb) cb({ status: 'succeeded' }); }
      };
      callback({ status: 'succeeded', value: file });
    }
  };
  return { context, log, slideStates, officeDocument };
}

module.exports = { createFakeHost, REAL_HOST_QUIRKS };

(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── 跨会话记忆（v132）──
  // 主题式工作记忆：agent 在工作中沉淀的用户偏好/习惯/规则，跨会话生效。
  // 与 Excel instructions 的分工：instructions 存「用户显式让存的规则」（全量注入）；
  // 这里存 agent 沉淀的工作经验（清单注入 + 工具按需读全文）。两者语义不同，并存。
  // 存储：localStorage 独立键（体量受护栏约束，见下方常量），不与 sessions 抢同一个键。

  const MAX_TOPICS = 12;          // 主题数上限：记忆是浓缩规则不是资料库，堆多必馊
  const MAX_TOPIC_CHARS = 6000;   // 单主题正文上限：超了说明该拆分或精简
  const MAX_TOTAL_CHARS = 64000;  // 总量上限：给 localStorage 配额留足余量
  const MAX_APPEND_CHARS = 2000;  // 单次追加上限：逼 agent 精炼而不是整段粘贴
  const MAX_TRIGGERS = 8;         // 单主题触发词上限
  const MAX_TRIGGER_CHARS = 24;   // 单个触发词长度上限：是关键词不是句子
  const KEY_RE = /^[a-z0-9][a-z0-9_-]{0,47}$/;

  function storageKey() { return (App.STORAGE_KEYS && App.STORAGE_KEYS.memory) || 'office-ai-trial-agent-memory-v1'; }
  function enabled() {
    const v = App.state && App.state.settings && App.state.settings.memoryEnabled;
    return v === undefined ? true : Boolean(v); // 存量用户默认开启
  }
  function setEnabled(value) {
    const s = App.state && App.state.settings;
    if (!s) return;
    s.memoryEnabled = Boolean(value);
    if (typeof App.saveSettings === 'function') App.saveSettings();
  }

  function blank() { return { version: 1, topics: [], contents: {} }; }

  function load() {
    let raw = null;
    try { raw = (typeof window !== 'undefined' && window.localStorage) ? window.localStorage.getItem(storageKey()) : null; } catch { raw = null; }
    if (!raw) return blank();
    try {
      const data = JSON.parse(raw);
      if (!data || typeof data !== 'object' || !Array.isArray(data.topics) || typeof data.contents !== 'object' || !data.contents) return blank();
      // 元素级损坏自愈：坏条目/孤儿正文直接丢弃，保住工具与设置卡的可用性（坏一次不能瘫全部）
      data.topics = data.topics.filter(t => t && typeof t.key === 'string');
      for (const k of Object.keys(data.contents)) {
        if (!findTopic(data, k) || typeof data.contents[k] !== 'string') delete data.contents[k];
      }
      return data;
    } catch { return blank(); }
  }

  // 保存失败必须报错：静默吞错 = agent 以为记住了（真实机纪律「谎报成功」）
  function save(data) {
    const text = JSON.stringify(data);
    if (text.length > MAX_TOTAL_CHARS * 2) throw new Error(`memory store too large (${text.length} chars)`);
    try {
      if (typeof window === 'undefined' || !window.localStorage) throw new Error('storage unavailable');
      window.localStorage.setItem(storageKey(), text);
    } catch (e) {
      throw new Error(`保存记忆失败（${(e && e.message) || e}）。请整理记忆库（删除过时主题）后重试`);
    }
  }

  function topicChars(content) { return String(content || '').length; }
  function totalChars(data) { return data.topics.reduce((sum, t) => sum + topicChars(data.contents[t.key]), 0); }
  function findTopic(data, key) { return data.topics.find(t => t.key === key) || null; }

  function normalizeKey(value) {
    const key = String(value || '').trim().toLowerCase().replace(/\s+/g, '_');
    if (!KEY_RE.test(key)) {
      throw new Error(`topic key 必须是 1-48 位小写字母/数字/连字符/下划线（如 ppt_style_habits），收到:「${String(value).slice(0, 80)}」`);
    }
    return key;
  }

  // 触发词（A1 文档级知识注入，v145）：关键词数组，prompt 组装时按最新用户消息命中置顶。
  // 接受数组；字符串按中英文逗号/分号/空白拆分（设置卡手工输入的形态）。错误逐项自解释。
  function normalizeTriggers(value) {
    if (value == null || value === '') return [];
    let list = value;
    if (typeof value === 'string') list = value.split(/[,，;；\s]+/);
    if (!Array.isArray(list)) throw new Error(`triggers 必须是关键词数组（如 ["标书", "周报"]），收到:「${String(value).slice(0, 80)}」`);
    const out = [];
    for (const raw of list) {
      const t = String(raw == null ? '' : raw).trim().slice(0, MAX_TRIGGER_CHARS);
      if (t) out.push(t);
    }
    return [...new Set(out)].slice(0, MAX_TRIGGERS);
  }

  function topicTriggers(topic) {
    return Array.isArray(topic && topic.triggers) ? topic.triggers : [];
  }

  // 最新用户消息命中任一触发词 → 该主题与本次任务相关
  function triggerMatches(topic, messageText) {
    const text = String(messageText || '').toLowerCase();
    if (!text) return false;
    return topicTriggers(topic).some(g => text.includes(String(g).toLowerCase()));
  }

  // ── 工具执行器（给模型的结构化返回，报错必须自解释并给出修正方向）──

  async function readMemoryTool(args) {
    if (!enabled()) return { success: false, error: '记忆功能已在设置中关闭。不要尝试读写记忆，直接按对话内容工作。' };
    const data = load();
    const key = args && args.topic ? String(args.topic).trim() : '';
    if (!key) {
      if (!data.topics.length) return { success: true, count: 0, topics: [], note: '记忆库为空。' };
      return {
        success: true,
        count: data.topics.length,
        topics: data.topics.map(t => ({ key: t.key, title: t.title, summary: t.summary, triggers: topicTriggers(t), updatedAt: fmtDate(t.updatedAt) })),
        note: '以上是记忆清单（存储序）。需要某条的完整内容，再调 read_memory 并传 topic。'
      };
    }
    const topic = findTopic(data, normalizeKey(key));
    if (!topic) {
      const known = data.topics.map(t => t.key).join(', ') || '(空)';
      return { success: false, error: `记忆里没有「${key}」这条主题。现有: ${known}` };
    }
    return {
      success: true,
      key: topic.key,
      title: topic.title,
      summary: topic.summary,
      triggers: topicTriggers(topic),
      updatedAt: fmtDate(topic.updatedAt),
      content: data.contents[topic.key] || ''
    };
  }

  async function writeMemoryTool(args) {
    if (!enabled()) return { success: false, error: '记忆功能已在设置中关闭。不要尝试读写记忆，直接按对话内容工作。' };
    const action = String((args && args.action) || '').trim();
    if (!['create', 'append', 'update', 'delete'].includes(action)) {
      throw new Error(`action 必须是 create/append/update/delete（got: ${action || '(empty)'}）。小增量优先用 append`);
    }
    const key = normalizeKey(args && args.topic);
    const data = load();
    const existing = findTopic(data, key);

    if (action === 'delete') {
      if (!existing) return { success: true, deleted: key, note: '该主题本就不存在，无需删除。' };
      data.topics = data.topics.filter(t => t.key !== key);
      delete data.contents[key];
      save(data);
      return { success: true, deleted: key, remaining: data.topics.length };
    }

    const content = String((args && args.content) || '').trim();
    if (!content) throw new Error('content 不能为空。要写进记忆的规则/偏好正文放这里');

    if (action === 'create') {
      if (existing) {
        return { success: false, error: `主题「${key}」已存在（${existing.title}）。追加内容用 append，整理重写用 update。` };
      }
      if (data.topics.length >= MAX_TOPICS) {
        const list = data.topics.map(t => t.key).join(', ');
        throw new Error(`记忆已有 ${MAX_TOPICS} 条主题，先整理再新建（read_memory 看清单 → 把过时的合并或 delete）。现有: ${list}`);
      }
      const title = String((args && args.title) || key).trim().slice(0, 40) || key;
      const summary = String((args && args.summary) || '').trim().slice(0, 80);
      const triggers = normalizeTriggers(args && args.triggers);
      if (content.length > MAX_TOPIC_CHARS) throw new Error(`单条记忆正文上限 ${MAX_TOPIC_CHARS} 字符（收到 ${content.length}）。先精简，或拆成多条主题`);
      if (totalChars(data) + content.length > MAX_TOTAL_CHARS) throw new Error('记忆总量已达上限，先整理（合并/删除过时主题）再写入');
      data.topics.push({ key, title, summary, triggers, updatedAt: Date.now() });
      data.contents[key] = content;
      save(data);
      return { success: true, created: key, title, triggers, topics: data.topics.length, note: '用户会在设置里看到这条记忆并可删除。绑定文档/任务类型的主题建议给 1-3 个触发词，命中时自动置顶。' };
    }

    if (!existing) {
      return { success: false, error: `主题「${key}」不存在。新建用 action=create，或先 read_memory 看现有清单。` };
    }

    if (action === 'append') {
      if (content.length > MAX_APPEND_CHARS) throw new Error(`单次追加上限 ${MAX_APPEND_CHARS} 字符（收到 ${content.length}）。记忆是浓缩规则，请精炼后重写；需要整段重排用 update`);
      const merged = (data.contents[key] || '') + '\n' + content;
      if (topicChars(merged) > MAX_TOPIC_CHARS) throw new Error(`主题「${key}」已到 ${topicChars(merged)}/${MAX_TOPIC_CHARS} 字符上限。用 update 重新整理这一条（合并冗余、删过时）`);
      const others = totalChars(data) - topicChars(data.contents[key]);
      if (others + topicChars(merged) > MAX_TOTAL_CHARS) throw new Error('记忆总量已达上限，先整理（合并/删除过时主题）再写入');
      data.contents[key] = merged;
      existing.updatedAt = Date.now();
      if (args && args.summary) existing.summary = String(args.summary).trim().slice(0, 80) || existing.summary;
      if (args && args.triggers != null) existing.triggers = normalizeTriggers([...topicTriggers(existing), ...normalizeTriggers(args.triggers)]);
      save(data);
      return { success: true, appended: key, triggers: topicTriggers(existing), topicChars: topicChars(merged), maxTopicChars: MAX_TOPIC_CHARS, topics: data.topics.length };
    }

    // update：整条重写（agent 主动整理）
    if (content.length > MAX_TOPIC_CHARS) throw new Error(`单条记忆正文上限 ${MAX_TOPIC_CHARS} 字符（收到 ${content.length}）`);
    const others = totalChars(data) - topicChars(data.contents[key]);
    if (others + content.length > MAX_TOTAL_CHARS) throw new Error('记忆总量已达上限，先精简内容再写入');
    data.contents[key] = content;
    existing.updatedAt = Date.now();
    if (args && args.summary) existing.summary = String(args.summary).trim().slice(0, 80) || existing.summary;
    if (args && args.triggers != null) existing.triggers = normalizeTriggers(args.triggers);
    save(data);
    return { success: true, updated: key, triggers: topicTriggers(existing), topicChars: topicChars(content), topics: data.topics.length };
  }

  // ── 系统提示词注入节（api.js 组装 systemContent 时拼接；禁用/无记忆返回空串）──

  function fmtDate(ts) {
    try { return new Date(ts).toISOString().slice(0, 10); } catch { return ''; }
  }

  function promptSection(latestMessage) {
    if (!enabled()) return '';
    const data = load();
    if (!data.topics.length) {
      return [
        '## 跨会话记忆（可用）',
        '你的记忆库当前为空。每个任务收尾时反观一轮：用户是否纠正过你的做法、明确表达过风格/格式偏好？有稳定偏好就用 write_memory(action=create) 沉淀，下次会话直接生效。绑定文档/任务类型的主题给 1-3 个 triggers（触发词），命中时该条会置顶。',
        '纪律：只有用户明确表达或反复出现的偏好才值得记；单次行为不得当作规则；时间敏感内容标注日期；不记密钥/隐私；用户可在设置中查看并删除记忆。'
      ].join('\n');
    }
    // A1 渐进披露简版：命中最新消息触发词的主题置顶并打 ★，其余保持原序
    const matched = data.topics.filter(t => triggerMatches(t, latestMessage));
    const rest = data.topics.filter(t => !triggerMatches(t, latestMessage));
    const line = t => {
      const tags = topicTriggers(t);
      return `- ${t.key} · ${t.title}${t.summary ? ' · ' + t.summary : ''}（${fmtDate(t.updatedAt)}）${tags.length ? `〔触发词: ${tags.join('、')}〕` : ''}`;
    };
    const lines = [
      ...matched.map(t => `★ ${line(t)}`),
      ...rest.map(line)
    ];
    return [
      '## 跨会话记忆',
      '以下是你在以往工作中沉淀的记忆（跨会话生效）。本次任务涉及其中主题时，先 read_memory(topic=…) 读全文再动手。' + (matched.length ? '带 ★ 的是命中本次消息触发词的主题，与当前任务直接相关，优先读。' : ''),
      lines.join('\n'),
      '沉淀纪律：任务收尾时主动反观本轮是否有值得沉淀的用户偏好（每个 Office 任务收尾自查一次）；用 write_memory 记录，小增量用 append；只有用户明确表达或反复出现的偏好才值得记；单次行为不得当作规则；时间敏感内容标注日期；不记密钥/隐私；用户可在设置中查看并删除记忆，写入内容要经得起用户检查。'
    ].join('\n');
  }

  // ── 设置卡辅助（ui.js 用）──

  function stats() {
    const data = load();
    return {
      enabled: enabled(),
      count: data.topics.length,
      maxTopics: MAX_TOPICS,
      chars: totalChars(data),
      maxChars: MAX_TOTAL_CHARS,
      topics: data.topics.map(t => ({
        key: t.key, title: t.title, summary: t.summary,
        triggers: topicTriggers(t),
        chars: topicChars(data.contents[t.key]),
        updatedAt: t.updatedAt, dateLabel: fmtDate(t.updatedAt)
      }))
    };
  }

  // 设置卡编辑触发词（A1 最小版）：接受逗号/中文逗号分隔的字符串或数组
  function setTopicTriggers(key, value) {
    const data = load();
    const k = normalizeKey(key);
    const topic = findTopic(data, k);
    if (!topic) throw new Error(`记忆里没有「${key}」这条主题`);
    const triggers = normalizeTriggers(value);
    topic.triggers = triggers;
    topic.updatedAt = Date.now();
    save(data);
    return triggers;
  }

  function removeTopic(key) {
    const data = load();
    const k = normalizeKey(key);
    data.topics = data.topics.filter(t => t.key !== k);
    delete data.contents[k];
    save(data);
  }

  function clearAll() {
    save(blank());
  }

  function readTopicText(key) {
    const data = load();
    const k = normalizeKey(key);
    return findTopic(data, k) ? (data.contents[k] || '') : null;
  }

  App.memory = {
    readMemoryTool, writeMemoryTool, promptSection,
    stats, removeTopic, clearAll, setEnabled, enabled, readTopicText, setTopicTriggers,
    MAX_TOPICS, MAX_TOPIC_CHARS, MAX_TOTAL_CHARS, MAX_APPEND_CHARS
  };
})();

(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const TEXT_FILE_RE = /\.(md|txt|ya?ml|json|csv|js|ts|py)$/i;
  const SIMPLE_SKILL_RE = /\.(md|txt)$/i;
  const MAX_SKILL_BODY_CHARS = 180000;

  function normalizePath(value) {
    return String(value || '').replace(/\\/g, '/').replace(/^\.\//, '').replace(/\/{2,}/g, '/').replace(/^\/+|\/+$/g, '');
  }

  function baseName(filePath) {
    const parts = normalizePath(filePath).split('/');
    return parts[parts.length - 1] || '';
  }

  function dirName(filePath) {
    const path = normalizePath(filePath);
    const index = path.lastIndexOf('/');
    return index < 0 ? '' : path.slice(0, index);
  }

  function stem(filePath) {
    return baseName(filePath).replace(/\.[^.]+$/, '');
  }

  function cleanMetadataValue(value) {
    return String(value || '').trim().replace(/^['"]|['"]$/g, '').trim();
  }

  function frontmatter(body) {
    const text = String(body || '').replace(/^\uFEFF/, '');
    const match = text.match(/^---\s*\r?\n([\s\S]*?)\r?\n---\s*(?:\r?\n|$)/);
    if (!match) return { values: {}, rest: text };
    const values = {};
    for (const line of match[1].split(/\r?\n/)) {
      const item = line.match(/^([A-Za-z][\w-]*)\s*:\s*(.+)$/);
      if (item) values[item[1].toLowerCase()] = cleanMetadataValue(item[2]);
    }
    return { values, rest: text.slice(match[0].length) };
  }

  function cleanDisplayName(value) {
    return String(value || '').replace(/^#+\s*/, '').replace(/[*_`]/g, '').replace(/\s+/g, ' ').trim().slice(0, 72);
  }

  function isGenericSkillName(value) {
    return /^(skill|skills|skill\.md|index|untitled|未命名技能?)$/i.test(cleanDisplayName(value));
  }

  function firstUsefulLine(text) {
    for (const raw of String(text || '').split(/\r?\n/)) {
      const line = cleanDisplayName(raw);
      if (!line || /^[-:>{}\[\]]+$/.test(line) || /^```/.test(raw.trim()) || isGenericSkillName(line)) continue;
      return line.slice(0, 48);
    }
    return '';
  }

  function pathFallback(filePath, body) {
    const fileStem = cleanDisplayName(stem(filePath));
    const folder = cleanDisplayName(baseName(dirName(filePath)));
    if (fileStem && !isGenericSkillName(fileStem)) return fileStem;
    if (folder && !isGenericSkillName(folder)) return folder;
    return firstUsefulLine(frontmatter(body).rest) || '未命名技能';
  }

  function skillMetadata(body, fallback) {
    const parsed = frontmatter(body);
    const heading = (parsed.rest.match(/^#\s+(.+)$/m) || [])[1];
    let name = cleanDisplayName(parsed.values.name || parsed.values.title || heading || fallback);
    if (!name || isGenericSkillName(name)) name = cleanDisplayName(fallback || firstUsefulLine(parsed.rest)) || '未命名技能';

    let desc = cleanDisplayName(parsed.values.description || parsed.values.desc || '');
    if (!desc) {
      const lines = parsed.rest.split(/\r?\n/).map(line => line.trim());
      desc = cleanDisplayName(lines.find(line => line && !/^#|^```|^[-*+]\s/.test(line)) || '');
    }
    return { name, desc: desc.slice(0, 88) };
  }

  function identityKey(value) {
    return cleanDisplayName(value).toLowerCase().replace(/[\s._/\\-]+/g, '');
  }

  function makeSkillId(name) {
    let slug = cleanDisplayName(name).toLowerCase()
      .replace(/[’'“”"`]/g, '')
      .replace(/[^a-z0-9\u3400-\u9fff]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 60);
    if (!slug) slug = 'custom-skill';
    return 'cs-' + slug;
  }

  function normalizeCustomSkillEntries(entries) {
    const out = [];
    const seen = new Map();
    for (const raw of (Array.isArray(entries) ? entries : [])) {
      if (!raw || !String(raw.body || '').trim()) continue;
      const fallback = isGenericSkillName(raw.name) ? firstUsefulLine(frontmatter(raw.body).rest) : raw.name;
      const meta = skillMetadata(raw.body, fallback || raw.name);
      const name = isGenericSkillName(raw.name) || !cleanDisplayName(raw.name) ? meta.name : cleanDisplayName(raw.name);
      const genericId = !raw.id || /^cs-(skill|skills|skill-md|untitled)$/i.test(raw.id);
      const entry = Object.assign({}, raw, {
        id: genericId ? makeSkillId(name) : raw.id,
        name,
        desc: cleanDisplayName(raw.desc || meta.desc)
      });
      const key = identityKey(name) || String(entry.id || '').toLowerCase();
      if (seen.has(key)) out[seen.get(key)] = entry;
      else { seen.set(key, out.length); out.push(entry); }
    }
    return out;
  }

  function isInside(path, folder) {
    return !folder || path === folder || path.startsWith(folder + '/');
  }

  function nearestMainFor(record, mains) {
    let best = null;
    let bestLength = -1;
    for (const main of mains) {
      const folder = dirName(main.path);
      if (isInside(record.path, folder) && folder.length > bestLength) {
        best = main;
        bestLength = folder.length;
      }
    }
    return best;
  }

  function buildEntry(main, resources) {
    const folder = dirName(main.path);
    const sections = [String(main.text || '').trim()];
    for (const resource of resources.filter(item => item !== main).sort((a, b) => a.path.localeCompare(b.path))) {
      const relative = folder ? resource.path.slice(folder.length + 1) : resource.path;
      const text = String(resource.text || '').trim();
      if (text) sections.push(`## 附加资料：${relative}\n\n${text}`);
    }
    const body = sections.filter(Boolean).join('\n\n');
    if (body.length > MAX_SKILL_BODY_CHARS) {
      throw new Error(`Skill「${baseName(folder) || baseName(main.path)}」内容超过 ${MAX_SKILL_BODY_CHARS} 字，请删减无关附件后再导入。`);
    }
    const meta = skillMetadata(main.text, pathFallback(main.path, main.text));
    return {
      id: makeSkillId(meta.name),
      name: meta.name,
      desc: meta.desc,
      body,
      sourceKey: normalizePath(main.path).toLowerCase()
    };
  }

  function skillEntriesFromTextRecords(records) {
    const clean = (Array.isArray(records) ? records : [])
      .map(item => ({ path: normalizePath(item.path), text: String(item.text || '') }))
      .filter(item => item.path && TEXT_FILE_RE.test(item.path) && item.text.trim() && !/(^|\/)__MACOSX\//.test(item.path));
    const mains = clean.filter(item => /^skill\.md$/i.test(baseName(item.path)));
    const entries = [];
    const assigned = new Set();

    for (const main of mains) {
      const resources = clean.filter(item => nearestMainFor(item, mains) === main);
      resources.forEach(item => assigned.add(item));
      entries.push(buildEntry(main, resources));
    }

    for (const item of clean) {
      if (!assigned.has(item) && SIMPLE_SKILL_RE.test(item.path) && !/^readme\.md$/i.test(baseName(item.path))) {
        entries.push(buildEntry(item, [item]));
      }
    }
    if (!entries.length) {
      const readme = clean.find(item => /^readme\.md$/i.test(baseName(item.path)));
      if (readme) entries.push(buildEntry(readme, [readme]));
    }
    return entries;
  }

  async function readSkillInputs(fileList) {
    const records = [];
    for (const file of Array.from(fileList || [])) {
      const sourcePath = normalizePath(file.webkitRelativePath || file.relativePath || file.name);
      if (!sourcePath) continue;
      if (/\.zip$/i.test(sourcePath)) {
        if (typeof App.unzipArchive !== 'function') throw new Error('ZIP 读取组件尚未加载，请重新打开插件后再试。');
        const archiveRoot = cleanDisplayName(stem(sourcePath)) || 'imported-skill';
        const archiveFiles = await App.unzipArchive(await file.arrayBuffer());
        for (const item of archiveFiles) {
          const inner = normalizePath(item.path);
          if (!inner || !TEXT_FILE_RE.test(inner) || /(^|\/)__MACOSX\//.test(inner)) continue;
          records.push({ path: `${archiveRoot}/${inner}`, text: new TextDecoder('utf-8').decode(item.data) });
        }
      } else if (TEXT_FILE_RE.test(sourcePath)) {
        records.push({ path: sourcePath, text: await file.text() });
      }
    }
    return skillEntriesFromTextRecords(records);
  }

  function mergeCustomSkills(current, builtin, incoming, options = {}) {
    const custom = normalizeCustomSkillEntries(current);
    const occupied = new Map();
    for (const item of (Array.isArray(builtin) ? builtin : [])) {
      occupied.set('name:' + identityKey(item.name), { kind: 'builtin', item });
      occupied.set('id:' + String(item.id || '').toLowerCase(), { kind: 'builtin', item });
    }
    custom.forEach((item, index) => {
      occupied.set('name:' + identityKey(item.name), { kind: 'custom', item, index });
      occupied.set('id:' + String(item.id || '').toLowerCase(), { kind: 'custom', item, index });
      if (item.sourceKey) occupied.set('source:' + item.sourceKey, { kind: 'custom', item, index });
    });

    let added = 0;
    let updated = 0;
    let existing = 0;
    for (const raw of (Array.isArray(incoming) ? incoming : [])) {
      const entry = normalizeCustomSkillEntries([raw])[0];
      if (!entry) continue;
      const match = (entry.sourceKey && occupied.get('source:' + entry.sourceKey))
        || occupied.get('name:' + identityKey(entry.name))
        || occupied.get('id:' + String(entry.id || '').toLowerCase());
      if (match) {
        if (options.replaceExisting && match.kind === 'custom') {
          entry.id = match.item.id;
          custom[match.index] = Object.assign({}, match.item, entry);
          updated++;
        } else {
          existing++;
        }
        continue;
      }

      let candidate = entry.id;
      let suffix = 2;
      while (occupied.has('id:' + candidate.toLowerCase())) candidate = entry.id + '-' + suffix++;
      entry.id = candidate;
      custom.push(entry);
      const index = custom.length - 1;
      occupied.set('name:' + identityKey(entry.name), { kind: 'custom', item: entry, index });
      occupied.set('id:' + entry.id.toLowerCase(), { kind: 'custom', item: entry, index });
      if (entry.sourceKey) occupied.set('source:' + entry.sourceKey, { kind: 'custom', item: entry, index });
      added++;
    }
    return { skills: custom, added, updated, existing };
  }

  App.skillMetadata = skillMetadata;
  App.skillEntriesFromTextRecords = skillEntriesFromTextRecords;
  App.readSkillInputs = readSkillInputs;
  App.mergeCustomSkills = mergeCustomSkills;
  App.normalizeCustomSkillEntries = normalizeCustomSkillEntries;
  App.skillIdentityKey = identityKey;
})();

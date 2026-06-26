#!/usr/bin/env node
/**
 * generate_tex.js — 从 摘要信息.xls 生成 LaTeX 摘要集正文
 *
 * 用法：node generate_tex.js
 * 输出：papers_content.tex（可直接 \input 到主模板）
 *
 * 处理逻辑：
 *   1. 读取 Excel 中每篇论文的作者1-12及对应单位
 *   2. 将所有单位拆分逗号 → 去重 → 按首次出现顺序编号
 *   3. 通讯作者在姓名后加 \ca（*号上标）
 *   4. 每位作者姓名后按 \ns{编号} 加单位序号
 *   5. 生成 \printpaper{...}{...}{...}{作者行}{单位行}{摘要}{关键词}
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const EXCEL_PATH = path.join(__dirname, '摘要信息.xls');
const OUTPUT_PATH = path.join(__dirname, 'papers_content.tex');

// ---- 工具函数 ----

/** 安全 trim，处理 null/undefined */
function s(v) { return (v || '').toString().trim(); }

/** 转义 LaTeX 特殊字符 */
function escapeLatex(text) {
  return text
    .replace(/\\/g, '\\textbackslash{}')
    .replace(/%/g, '\\%')
    .replace(/#/g, '\\#')
    .replace(/\$/g, '\\$')
    .replace(/&/g, '\\&')
    .replace(/_/g, '\\_')
    .replace(/\^/g, '\\^{}')
    .replace(/~/g, '\\~{}')
    .replace(/\{/g, '\\{')
    .replace(/\}/g, '\\}')
    .replace(/</g, '\\textless{}')
    .replace(/>/g, '\\textgreater{}');
}

/** 将 HTML 实体（&alpha; 等）和希腊字母转为 LaTeX 命令 */
function convertHtmlEntities(text) {
  const map = {
    '&alpha;': '$\\alpha$',  '&Alpha;': '$\\Alpha$',
    '&beta;': '$\\beta$',    '&Beta;': '$\\Beta$',
    '&gamma;': '$\\gamma$',  '&Gamma;': '$\\Gamma$',
    '&delta;': '$\\delta$',  '&Delta;': '$\\Delta$',
    '&mu;': '$\\mu$',        '&micro;': '$\\mu$',
    '&plusmn;': '$\\pm$',    '±': '$\\pm$',
    '&mdash;': '---',        '&ndash;': '--',
    '&hellip;': '…',         '&lsquo;': "'",
    '&rsquo;': "'",          '&ldquo;': '"',
    '&rdquo;': '"',          '&lt;': '<',
    '&gt;': '>',             '&nbsp;': '~',
    '&deg;': '\\textdegree{}',
    '&sup2;': '\\textsuperscript{2}',
    '&sup3;': '\\textsuperscript{3}',
  };
  let result = text;
  for (const [entity, latex] of Object.entries(map)) {
    result = result.split(entity).join(latex);
  }
  return result;
}

/** 转义摘要文本：先转 HTML 实体，再转 LaTeX 特殊字符（注意 & 在实体转换后再转义） */
function sanitizeAbstract(text) {
  let t = s(text);
  // 第一步：转义原始文本中的 LaTeX 特殊字符
  t = t.replace(/%/g, '\\%')
       .replace(/_/g, '\\_')
       .replace(/#/g, '\\#')
       .replace(/~/g, '\\~{}')
       .replace(/</g, '\\textless{}')
       .replace(/>/g, '\\textgreater{}');
  // 第二步：HTML/Unicode 实体 → LaTeX 命令（可能引入 $ 等数学符号）
  t = convertHtmlEntities(t);
  // 第三步：转义孤立的 &（不破坏已有的 LaTeX 命令如 \\textbackslash）
  t = t.replace(/(?<!\\)&(?!\w+;|[a-zA-Z]{2,})/g, '\\&');
  // 第四步：去掉 CJK Ext-E 等 pdflatex 不支持的字符
  t = t.replace(/[\u{2A700}-\u{2F7FF}]/gu, '?');
  return t;
}

/** 转义标题 */
function sanitizeTitle(text) {
  let t = s(text);
  t = t.replace(/%/g, '\\%')
       .replace(/_/g, '\\_')
       .replace(/#/g, '\\#')
       .replace(/~/g, '\\~{}');
  t = convertHtmlEntities(t);
  t = t.replace(/(?<!\\)&(?!\w+;|[a-zA-Z]{2,})/g, '\\&');
  t = t.replace(/^[\u200B-\u200F\uFEFF]+/, '');
  t = t.replace(/<[^>]*>/g, '');
  return t;
}

/** 转义关键词 */
function sanitizeKeywords(text) {
  let t = s(text);
  t = t.replace(/%/g, '\\%')
       .replace(/_/g, '\\_')
       .replace(/#/g, '\\#')
       .replace(/~/g, '\\~{}');
  t = convertHtmlEntities(t);
  t = t.replace(/(?<!\\)&(?!\w+;|[a-zA-Z]{2,})/g, '\\&');
  t = t.split(',').map(k => k.trim()).filter(Boolean).join('；');
  return t;
}

/** 转义作者姓名 */
function sanitizeAuthorName(text) {
  return s(text).trim();
}

// ---- 主处理逻辑 ----

/** 处理单篇论文，返回 LaTeX \printpaper 命令字符串 */
function processPaper(row) {
  const session = s(row['论文专题']);
  const paperId = s(row['论文编号']);
  const title = sanitizeTitle(row['论文标题']);
  const keywords = sanitizeKeywords(row['论文关键字']);
  const abstract = sanitizeAbstract(row['摘要文本']);
  const corrAuthorName = sanitizeAuthorName(row['通讯作者姓名']);

  // 收集所有作者和单位
  const authors = [];  // [{ name, unitStr }]
  for (let i = 1; i <= 12; i++) {
    const name = sanitizeAuthorName(row[`作者姓名${i}`]);
    const unit = s(row[`作者单位${i}`]);
    if (name) {
      authors.push({ name, unitStr: unit });
    }
  }

  if (authors.length === 0) return null;

  // 建立单位去重映射
  const unitMap = new Map();   // 原始单位字符串 → 编号
  const unitList = [];         // (按序) [{ id, text }]
  let nextUnitId = 1;

  function getOrCreateUnitId(unitText) {
    const key = unitText.trim();
    if (!key) return null;
    if (unitMap.has(key)) return unitMap.get(key);
    const id = nextUnitId++;
    unitMap.set(key, id);
    unitList.push({ id, text: key });
    return id;
  }

  // 为每位作者的逗号分隔单位列表分配编号
  const authorEntries = authors.map(a => {
    const unitTexts = a.unitStr
      .split(',')
      .map(u => u.trim())
      .filter(u => u.length > 0);
    // 去重后编号
    const ids = [...new Set(unitTexts)].map(t => getOrCreateUnitId(t)).filter(id => id !== null);
    // 若无单位，分配占位编号0（极少情况）
    return {
      name: a.name,
      unitIds: ids.length > 0 ? ids : [0],
      isCorresponding: (a.name === corrAuthorName),
    };
  });

  // 构建作者 LaTeX 行（仅一个单位时省略上标序号）
  const isSingleUnit = unitList.length <= 1;
  const authorLine = authorEntries.map(a => {
    const name = escapeLatex(a.name);
    const superscripts = a.unitIds.join(',');
    const star = a.isCorresponding ? '\\ca' : '';
    if (isSingleUnit) {
      return `${name}${star}`;
    }
    return `${name}\\ns{${superscripts}}${star}`;
  }).join(',\n    ');

  // 构建单位 LaTeX 行（不加序号前缀，用中文分号分隔）
  const unitLine = unitList.map(u => {
    const escaped = escapeLatex(u.text);
    return escaped;
  }).join('；\\hspace{1em}\n    ');

  // 标题换行处理：过长标题在 40 字处断行
  const titleLine = title.length > 45 ? `${title.slice(0, Math.floor(title.length/2))}\\\\\n   ${title.slice(Math.floor(title.length/2))}` : title;
  const isTitleMultiLine = title.length > 45;

  const labelId = paperId.replace(/[^A-Za-z0-9]/g, '');
  return `\\printpaper
  {${session}}
  {${paperId}}
  {${titleLine}}
  {%
    ${authorLine}
  }
  {%
    ${unitLine}
  }
  {%
    ${abstract}
  }
  {%
    ${keywords}
  }
  {${labelId}}
  {${isTitleMultiLine}}`;
}

// ---- 主流程 ----

console.log('Reading Excel...');
const wb = XLSX.readFile(EXCEL_PATH);
const sheet = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
console.log(`Found ${data.length} papers.`);

// 按专题分组排序
data.sort((a, b) => {
  const sa = s(a['论文专题']);
  const sb = s(b['论文专题']);
  if (sa !== sb) return sa.localeCompare(sb, 'zh-CN');
  return s(a['论文编号']).localeCompare(s(b['论文编号']));
});

const papers = [];
for (const row of data) {
  const result = processPaper(row);
  if (result) papers.push(result);
}

console.log(`Processed ${papers.length} valid papers.`);

// 生成 TOC（按专题分组，带引导符和超链接）
const TOC_PATH = path.join(__dirname, 'toc_content.tex');
const sessions = new Map();
for (const row of data) {
  const sid = s(row['论文编号']);
  const session = s(row['论文专题']);
  if (!sid) continue;
  const labelId = sid.replace(/[^A-Za-z0-9]/g, '');
  const authorNames = [];
  for (let i = 1; i <= 12; i++) {
    const n = sanitizeAuthorName(row[`作者姓名${i}`]);
    if (n) authorNames.push(n);
  }
  const authorStr = authorNames.join(', ');
  const title = sanitizeTitle(row['论文标题']);
  if (!sessions.has(session)) sessions.set(session, []);
  sessions.get(session).push({ id: sid, labelId, title, authors: authorStr });
}

const tocLines = [
  '% ============================================================',
  '%  toc_content.tex — 目录（由 generate_tex.js 自动生成）',
  `%  生成时间：${new Date().toISOString()}`,
  '%  注意：需二次编译才能正确显示页码',
  '% ============================================================',
  '',
  '\\begin{center}',
  '  {\\heiti\\xiaoer 目\\quad 录}',
  '\\end{center}',
  '\\vspace{12pt}',
  '',
];

for (const [session, entries] of sessions) {
  tocLines.push(`\\begin{center}{\\heiti\\xiaosi ${session}}\\end{center}`);
  tocLines.push('\\par\\vspace{6pt}');
  for (const e of entries) {
    // 清理标题中的零宽字符和 HTML 标签
    const titleClean = e.title.replace(/^[\u200B-\u200F\uFEFF]+/, '').replace(/<[^>]*>/g, '');
    // 转义标题和作者中的 LaTeX 特殊字符
    const titleEsc = escapeLatex(titleClean);
    const authEsc = escapeLatex(e.authors);
    tocLines.push(`\\tocentry{${e.id}}{${e.labelId}}{${titleEsc}}{${authEsc}}`);
  }
  tocLines.push('\\vspace{6pt}');
  tocLines.push('');
}

fs.writeFileSync(TOC_PATH, tocLines.join('\n'), 'utf-8');
console.log(`TOC written to ${TOC_PATH}`);

// 生成输出
const output = [
  '% ============================================================',
  '%  papers_content.tex — 由 generate_tex.js 自动生成',
  `%  生成时间：${new Date().toISOString()}`,
  `%  论文数量：${papers.length}`,
  '% ============================================================',
  '',
  papers.join('\n\n'),
  '',
].join('\n');

fs.writeFileSync(OUTPUT_PATH, output, 'utf-8');
console.log(`Written to ${OUTPUT_PATH}`);




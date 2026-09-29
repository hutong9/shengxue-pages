#!/usr/bin/env node
/**
 * generate_tex.js — 从 摘要信息.xls 生成 LaTeX 摘要集正文
 *
 * 用法：node generate_tex.js
 * 输出：papers_content.tex（可直接 \input 到主模板）
 *
 * 处理逻辑：
 *   1. 读取 Excel 中每篇论文的作者1-12及其单位
 *   2. 「作者单位N / 作者单位地区N / 作者单位邮编N」是三个半角逗号分隔的平行列表，
 *      按索引 j 一一对应：单位[j] ↔ 地区[j] ↔ 邮编[j]。
 *      长度不等时按位置配对：多余忽略、缺失留空。
 *      全角「，」出现在单位名内部（如“南京大学物理学院，南京大学声学研究所”），不是分隔符。
 *   3. 缺失邮编（空串 / 全零占位 000000）用全表反查补齐：
 *      「单位名+地区」→ 唯一邮编 自动补（A 级）；多候选（B 级）/ 无候选（C 级）留空。
 *      A 级补齐在编号去重之前执行，使同一单位的“无邮编条目”与“有邮编条目”归并为同一编号。
 *   4. 地区名规范化：若同目录存在 region_map.csv（两列：原值,新值），
 *      将「区/县」等规范化成「市」再用于展示与去重；
 *      但邮编反查索引仍用**原始**地区名，以保证 A 级补齐精度不下降。
 *   5. 人工邮编对照表：若同目录存在 zip_map.csv（三列：单位名,地区,邮编；地区可留空表示不限），
 *      则命中者直接套用（优先级高于一切自动补齐）。
 *      两个映射表既可为 UTF-8（可带 BOM），也可为 GBK/GB18030（中文 Excel 另存的 CSV）。
 *      海外单位的条目可在表里留空「地区 + 邮编」，脚本会保持留空、不做补齐。
 *   6. 按「单位名+地区+邮编」三者全同去重并编号（按首次出现顺序）
 *   7. 通讯作者在姓名后加 \ca（*号上标）
 *   8. 每位作者姓名后按 \ns{编号} 加单位序号（去重且升序）
 *   9. 生成 \printpaper{...}，并输出 unit_report.csv 审计报告
 *  10. 去掉摘要正文开头作者误写的「摘要：」标签（模板已自行印出「摘要」二字），
 *      源表里有 15 篇出现「摘要　摘要：…」的重复。
 *  11. 删掉「数值 单位」之间的空格（如 20 mm → 20mm）：普通空格在 TeX 里是
 *      可拉伸的 interword glue，中文段落两端对齐时会把它拉得很宽。
 *  12. 摘要正文**末尾**的基金 / 致谢句摘出来，作为 \printpaper 的第 9 个参数，
 *      由模板印在页脚「基金项目：」一行（源表里有 4 篇把致谢句写进了正文）。
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { TextDecoder } = require('util');

const EXCEL_PATH = path.join(__dirname, '摘要信息.xls');
const OUTPUT_PATH = path.join(__dirname, 'papers_content.tex');
const REPORT_PATH = path.join(__dirname, 'unit_report.csv');
const REGION_MAP_PATH = path.join(__dirname, 'region_map.csv');
const ZIP_MAP_PATH = path.join(__dirname, 'zip_map.csv');

// ---- 工具函数 ----

/** 安全 trim，处理 null/undefined；并清除零宽/不可见字符（pdfLaTeX 会渲染成空白） */
function s(v) {
  return (v || '').toString()
    .replace(/[\u200B-\u200F\u202A-\u202E\u2060\uFEFF\u00AD]/g, '')
    .trim();
}

/** 按半角逗号切分为列表：去空白、去空项。绝不切分全角「，」（它在单位名内部） */
function splitList(v) {
  return s(v).split(',').map(x => x.trim()).filter(Boolean);
}

/** 判断邮编是否缺失：空串 或 全零占位（000000 / 0000000…，允许前后空白） */
function isMissingZip(z) {
  const t = s(z);
  return t === '' || /^0+$/.test(t);
}

/**
 * 读取文本文件并自动识别编码。
 * 手工维护的 CSV（region_map.csv / zip_map.csv）经常用 Excel/WPS 编辑，
 * 而中文版 Excel 默认另存为 GBK/GB18030（不是 UTF-8），直接按 utf-8 读会变成乱码。
 * 策略：有 UTF-8 BOM 直接按 UTF-8；否则先做**严格** UTF-8 校验，失败再按 GB18030 解码。
 */
function readTextAuto(filePath) {
  const buf = fs.readFileSync(filePath);
  if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
    return buf.slice(3).toString('utf-8');
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buf);
  } catch (e) {
    try {
      console.log(`  注：${path.basename(filePath)} 不是 UTF-8，已按 GB18030（Excel 另存的中文 CSV）读取。`);
      return new TextDecoder('gb18030').decode(buf);
    } catch (e2) {
      return buf.toString('utf-8');
    }
  }
}

// ---- 地区名映射表（可选）----
// 文件 region_map.csv：两列「原值,新值」，UTF-8（可带 BOM）或 GBK，`#` 开头为注释行。
// 用途：把源表里的「区/县」规范化成「市」，例如 闵行区 → 上海市。
// 作用范围：**仅展示与单位去重**；邮编反查索引（buildZipIndex）仍用原始地区名，
//          否则「单位+城市」的键会比「单位+区」粗，可能把唯一映射变成多候选（A 级降为 B 级）。
const regionMap = new Map();

function loadRegionMap() {
  if (!fs.existsSync(REGION_MAP_PATH)) return;
  const text = readTextAuto(REGION_MAP_PATH);
  for (const line of text.split(/\r?\n/)) {
    const t = line.trim();
    if (!t || t.startsWith('#')) continue;
    const cells = t.split(',');
    const from = s(cells[0]);
    const to = s(cells[1]);
    if (!from || !to || from === to) continue;
    if (from === '原值') continue;                 // 跳过表头
    regionMap.set(from, to);
  }
}

/** 把原始地区名映射成规范地区名（表中无该条则原样返回） */
function mapRegion(r) {
  const t = s(r);
  return regionMap.get(t) || t;
}

// ---- 人工邮编对照表（可选，优先级最高）----
// 文件 zip_map.csv：三列「单位名,地区,邮编」（UTF-8 或 GBK 均可）。
//   地区 写**映射后的**地区名（与 unit_report.csv 的「地区」列一致，如 北京市）；
//   地区 留空表示「不限地区」，用于该单位名在全国独一无二的情况。
// 命中规则时直接套用该邮编（无论原文有无邮编）；邮编列留空的行不是规则，会被忽略。
const zipMap = new Map();       // `${单位名}||${地区}` → 邮编
const zipMapUsed = new Set();   // 已被命中的键（用于报告未命中项，防拼写不匹配）
const zipMapStats = { hit: 0, changed: 0 };   // 命中条目数 / 其中真正改写（或补上）的数

function loadZipMap() {
  if (!fs.existsSync(ZIP_MAP_PATH)) return;
  const text = readTextAuto(ZIP_MAP_PATH);
  for (const line of text.split(/\r?\n/)) {
    const t = line.trim();
    if (!t || t.startsWith('#')) continue;
    const cells = t.split(',');
    const unit = s(cells[0]);
    const region = s(cells[1]);
    const zip = s(cells[2]);
    if (!unit || unit === '单位名') continue;      // 跳过表头
    if (!zip || zip === '邮编') continue;          // 未填邮编 → 不是规则
    zipMap.set(`${unit}||${region}`, zip);
  }
}

/** 查人工对照表：先精确匹配「单位名+地区」，再退化为「单位名」（地区不限） */
function lookupZipMap(unit, dispRegion) {
  const exact = `${s(unit)}||${s(dispRegion)}`;
  if (zipMap.has(exact)) { zipMapUsed.add(exact); return zipMap.get(exact); }
  const wide = `${s(unit)}||`;
  if (zipMap.has(wide)) { zipMapUsed.add(wide); return zipMap.get(wide); }
  return null;
}

/**
 * 安全写文件。
 * 输出文件（papers_content.tex / toc_content.tex / unit_report.csv）常被 Excel、
 * WPS 或 PDF 阅读器打开，写入会抛 EBUSY。这里给出可操作的提示而非难懂的堆栈。
 * 注意：unit_report.csv / toc_content.tex / papers_content.tex 都是**生成物**，
 * 往里面手工补充内容一定会被下次运行覆盖；要改数据请改源表 摘要信息.xls 或映射表 region_map.csv。
 */
function writeTextFile(target, content) {
  try {
    fs.writeFileSync(target, content, 'utf-8');
  } catch (err) {
    if (err && (err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES')) {
      console.error('');
      console.error(`✗ 无法写入 ${path.basename(target)} —— 该文件正被其他程序占用（通常是 Excel / WPS）。`);
      console.error('  请先关闭它，然后重新运行：node generate_tex.js');
      console.error('');
      process.exit(1);
    }
    throw err;
  }
}

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

/**
 * 在中英混排边界插入 \allowbreak，提供合法断行点。
 *
 * 背景：pdfLaTeX + CJK 宏包下，汉字与拉丁字母/数字直接相连时（如
 * “以Gauss–Legendre–Lobatto节点”）会被当作一个不可断行的长词，
 * 一旦超出行宽就溢出右页边距（Overfull \hbox）。中文之间的断行由
 * CJK 宏包处理，故只需在“汉字↔拉丁/数字”交界处补上断行点。
 * 破折号（en dash）后同样不可断行，一并在其后插入断行点。
 *
 * 说明：\allowbreak 即 \penalty0，宽度为 0，不会影响 \settowidth
 * 量出的目录标题宽度，也不会引入任何可见空白。
 */
function softenBreaks(text) {
  if (!text) return text;
  const CJK = '\\u3400-\\u4DBF\\u4E00-\\u9FFF\\uF900-\\uFAFF';
  return text
    .replace(new RegExp(`([${CJK}])(?=[A-Za-z0-9])`, 'g'), '$1\\allowbreak{}')
    .replace(new RegExp(`([A-Za-z0-9])(?=[${CJK}])`, 'g'), '$1\\allowbreak{}')
    .replace(/\u2013/g, '\u2013\\allowbreak{}');   // – en dash
}

// ---- Unicode 上标 / 下标字符 → LaTeX 数学模式 ----
// 源表里混进了 Unicode 上标字符（例：`min⁻&sup1;` 的 `⁻`、`10⁻⁴` 的 `⁻⁴`）。
// pdfLaTeX + 中文 CJK 字体没有 U+2070–U+209F 的字形，而 CJK 宏包遇到缺字形
// 会**静默丢弃**（连 "Missing character" 都不报），于是：
//     min⁻¹ → 「min 1」      10⁻⁴ → 「10 ⁴」
// 这里把整段上/下标串转成 $^{-1}$ / $_{2}$，交给 newtxmath 排版。
const SUP_CHARS = '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ';
const SUB_CHARS = '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₕₖₗₘₙₚₛₜ';
const SUP_ASCII = {
  '⁰':'0','¹':'1','²':'2','³':'3','⁴':'4','⁵':'5','⁶':'6','⁷':'7','⁸':'8','⁹':'9',
  '⁺':'+','⁻':'-','⁼':'=','⁽':'(','⁾':')','ⁿ':'n',
};
const SUB_ASCII = {
  '₀':'0','₁':'1','₂':'2','₃':'3','₄':'4','₅':'5','₆':'6','₇':'7','₈':'8','₉':'9',
  '₊':'+','₋':'-','₌':'=','₍':'(','₎':')',
  'ₐ':'a','ₑ':'e','ₒ':'o','ₓ':'x','ₕ':'h','ₖ':'k','ₗ':'l','ₘ':'m','ₙ':'n','ₚ':'p','ₛ':'s','ₜ':'t',
};

// 源表还有「上标减号 + 普通数字」的混合写法（C0317 的 `10⁻4`）。
// true  ⇒ 合并成 $^{-4}$（正确）
// false ⇒ 原样保留，渲染为 `10⁻4`（减号上标、数字不上标，仍不美观）
// 全库只有 C0317 一处命中；更彻底的做法是把源表改成 `10⁻⁴`。
const MERGE_SIGN_DIGIT = true;

/**
 * Unicode / 命名上标、下标 → $^{…}$ / $_{…}$
 * 必须在命名实体替换**之前**调用：先把 `&sup1;` 归一成 `¹`，
 * 才能和源表里已有的 Unicode 上标合并成同一段（`⁻&sup1;` → `⁻¹` → `$^{-1}$`）。
 */
function supSubToMath(text) {
  if (!text) return text;
  let t = text;
  // ① 命名上标实体先归一成 Unicode，便于整段识别
  t = t.replace(/&sup([123]);/g, (_, n) => ({ '1': '¹', '2': '²', '3': '³' }[n]));
  // ② 连续上标串 → $^{…}$   （min⁻¹ → min$^{-1}$）
  t = t.replace(new RegExp(`[${SUP_CHARS}]+`, 'g'),
        m => '$^{' + Array.from(m).map(c => SUP_ASCII[c] || c).join('') + '}$');
  // ③ 连续下标串 → $_{…}$
  t = t.replace(new RegExp(`[${SUB_CHARS}]+`, 'g'),
        m => '$_{' + Array.from(m).map(c => SUB_ASCII[c] || c).join('') + '}$');
  // ④ 「纯符号上标 + 紧跟的普通数字」合并：10⁻4 → 10$^{-4}$
  if (MERGE_SIGN_DIGIT) {
    t = t.replace(/\$\^\{([-+]+)\}\$(\d+)/g, (_, sign, digits) => `$^{${sign}${digits}}$`);
  }
  return t;
}

// ---- 直双引号 " → 中文弯引号 “ ” ----
// 源表里的引号一律是 ASCII 直引号 `"`（例：A0463 的 在"双碳"目标）。
// 若原样写进 .tex，newtxtext(TeX Gyre Termes) 会把这**每一个** `"` 都排成
// 右引号 `”`（Termes 的 `"` 是按位置取字形的 TeX 连字，CJK 标点上下文里判错），
// 于是 PDF 上出现 在”双碳”目标 —— 开引号也是右引号。
//
// 源表里 `"` 的数量**恒为偶数**（全库 51 个字段、158 个引号，已逐一核对），
// 因此按「出现次序奇偶」交替判定开/闭即可；同时跟踪文本中已存在的 `“`/`”`
// （D0384、V0235 等少数条目本来就是真弯引号），避免混排时错位。
function normalizeDoubleQuotes(text) {
  if (!text || text.indexOf('"') === -1) return text;
  let open = true;                       // true ⇒ 下一个直引号是开引号
  return String(text).replace(/["\u201C\u201D]/g, (ch) => {
    if (ch === '\u201C') { open = false; return ch; }   // 已是开引号 → 同步状态
    if (ch === '\u201D') { open = true;  return ch; }   // 已是闭引号 → 同步状态
    const q = open ? '\u201C' : '\u201D';
    open = !open;
    return q;
  });
}

/** 将 HTML 实体（&alpha; 等）和希腊字母转为 LaTeX 命令 */
function convertHtmlEntities(text) {
  const map = {
    '&alpha;': '$\\alpha$',  '&Alpha;': '$\\Alpha$',
    '&beta;': '$\\beta$',    '&Beta;': '$\\Beta$',
    '&gamma;': '$\\gamma$',  '&Gamma;': '$\\Gamma$',
    '&delta;': '$\\delta$',  '&Delta;': '$\\Delta$',
    '&epsilon;': '$\\epsilon$',
    '&theta;': '$\\theta$',
    '&mu;': '$\\mu$',        '&micro;': '$\\mu$',
    '&omega;': '$\\omega$',
    '&pi;': '$\\pi$',
    '&sigma;': '$\\sigma$',
    '&tau;': '$\\tau$',
    '&plusmn;': '$\\pm$',    '±': '$\\pm$',
    '&minus;': '$-$',
    '&times;': '$\\times$',
    '&ge;': '$\\ge$',
    '&le;': '$\\le$',
    '&asymp;': '$\\asymp$',
    '&rarr;': '$\\rightarrow$',
    '&mdash;': '---',        '&ndash;': '--',
    '&hellip;': '…',         '&lsquo;': "'",
    '&rsquo;': "'",          '&ldquo;': '\u201C',
    '&rdquo;': '\u201D',     '&lt;': '<',
    '&gt;': '>',             '&nbsp;': '~',
    '&middot;': '$\\cdot$',
    '&ordm;': '\\textdegree{}',
    '&eacute;': "\\'{e}",
    '&zwnj;': '',
    '&deg;': '\\textdegree{}',
    // 说明：&sup1; &sup2; &sup3; 不在这里处理，
    //       统一交给 supSubToMath() 与 Unicode 上标合并成 $^{…}$
  };
  let result = text;
  // ① 数字实体**先**解码成 Unicode 字符：
  //    这样 &#185; / &#xB9;（即 `¹`）能并入下一步的上标串统一处理；
  //    若仍放在最后解码，解出的 `¹` 会直接漏进 .tex 被 pdfLaTeX 丢弃。
  result = result.replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)));
  result = result.replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)));
  // ② 解码后冒出的零宽 / 不可见字符，pdfLaTeX 不认，直接删掉
  result = result.replace(/[\u200B-\u200F\u202A-\u202E\u2060\uFEFF\u00AD]/g, '');
  // ③ Unicode / 命名上标、下标 → $^{…}$ / $_{…}$
  result = supSubToMath(result);
  // ④ 其余命名实体
  for (const [entity, latex] of Object.entries(map)) {
    result = result.split(entity).join(latex);
  }
  // ⑤ 直双引号 → 中文弯引号（放在最后：实体已展开成 `"` 的也一并归一）
  result = normalizeDoubleQuotes(result);
  return result;
}

// ---- 摘要正文开头的冗余标签 / 重复标题 ----
// 模板 \printpaper 已经自动印出「摘要」二字（main_full.tex 第 148–153 行，注释写明“不加冒号”），
// 但源表里有 15 篇的作者把「摘要：」也写进了正文，PDF 上就成了「摘要　　摘要：为充分考虑…」。
// 另有 D0043 把标题也在正文里重复了一遍，并单独一行写了「摘要」：
//     工业超声清洗机的空化控制研究\n\n摘要\n   \n本项研究是在…
// 这里循环剥离开头的：①「摘要：」/「【摘要】」标签 ②整行重复的论文标题 ③单独一行的「摘要」。
// 只处理**开头**，正文里出现的「摘要」二字不动。
// 兼容写法：摘要：／摘 要：／摘    要：／摘要:（半角）／【摘要】
//
// 注意：源表里 D0043 的「论文标题」以 U+200B 零宽空格开头（Excel 复制粘贴残留），
// 而 JS 的 \s **不匹配**零宽字符，直接比对会导致标题行剥不掉。
// 所以比对前两边都要先去掉零宽字符（ZWSP/ZWNJ/ZWJ/LRM/RLM/WJ/BOM）再比。
const ZERO_WIDTH_RE = /[\u200B-\u200F\u2060\uFEFF]/g;
const LEAD_JUNK_RE = /^[\s\u3000\u200B-\u200F\u2060\uFEFF]+/;
/** 比对用归一化：去掉零宽字符与所有空白 */
function normalizeForCompare(x) {
  return String(x == null ? '' : x).replace(ZERO_WIDTH_RE, '').replace(/\s+/g, '');
}
const ABSTRACT_LABEL_RE = /^[ \t\u3000]*(?:【[ \t\u3000]*摘要[ \t\u3000]*】|摘[ \t\u3000]*要[ \t\u3000]*[:：])[ \t\u3000]*/;
function stripAbstractLabel(text, rawTitle) {
  let t = String(text == null ? '' : text);
  const normTitle = normalizeForCompare(rawTitle);
  for (let guard = 0; guard < 6; guard++) {
    const before = t;
    // ① 去掉开头的空白与零宽字符（含换行）
    t = t.replace(LEAD_JUNK_RE, '');
    // ② 开头的「摘要：」标签
    t = t.replace(ABSTRACT_LABEL_RE, '');
    // ③ 开头整行与论文标题完全相同（去空白/零宽后）→ 删掉整行
    const firstLine = t.match(/^[^\r\n]*/);
    if (normTitle && firstLine && normalizeForCompare(firstLine[0]) === normTitle) {
      t = t.slice(firstLine[0].length);
    }
    // ④ 开头单独一行只有「摘要」（可带全/半角冒号）→ 删掉整行
    t = t.replace(/^[ \t\u3000]*摘[ \t\u3000]*要[ \t\u3000]*[:：]?[ \t\u3000]*(\r?\n|$)/, '');
    if (t === before) break;
  }
  return t;
}

// ---- 摘要正文末尾的基金 / 致谢句 → 页脚 ----
// 模板 \printpaper 的第 9 个参数专印「基金项目」，位置在页面左下角脚线之下的
// 页脚区（与「通讯作者」同处一块）。但源表里有 4 篇把致谢句直接写进了摘要正文末尾：
//     A0463  本课题承蒙国家自然科学基金项目(12274122，12574480)的支持，特此致谢！
//     A0464  本论文承蒙国家自然科学基金项目(12574480，12274122)的支持，特此致谢！
//     A0465  本课题承蒙国家自然科学基金项目(12574480，12274122)的支持，特此致谢！
//     E0461  本论文承蒙国家自然科学基金项目(12274122，12574480)的支持，特此致谢！
// 这里把摘要**末尾**的这类句子摘出来交给页脚渲染，正文里不再重复出现。
//
// 只认这些特征词，**且整句必须落在末尾**：正文中间的「课题」「资助」等词
// （如 B0169 的「水下声隐身成为重要课题」、G0396 的「课题组围绕…」）不受影响。
// 若某篇摘要整篇就是一个基金句（摘完为空），则不动它，避免把正文掏空。
const FUND_HINT_RE = /基金|资助|承蒙|致谢|感谢|重点研发计划|基金委/i;
// 句界（中文句号 / 叹号 / 问号 / 分号 / 换行）
const SENT_END_RE = /[。！？；!?;\n\r]/;

/**
 * 从摘要末尾摘出基金 / 致谢句。
 * @returns {{ abstract: string, funding: string }} funding 为空串表示未找到。
 */
function extractFunding(raw) {
  const text = String(raw == null ? '' : raw);
  const body = text.replace(/[\s\u3000]+$/, '');                 // 去掉尾部空白
  if (!body || !FUND_HINT_RE.test(body)) return { abstract: text, funding: '' };

  // ① 先剥掉整段末尾的句末标点，便于定位「最后一句」的起止
  let end = body.length;
  while (end > 0 && /[。！？；!?;]/.test(body[end - 1])) end--;
  const punct = body.slice(end);                                // 末尾标点（如「！」）
  const core = body.slice(0, end);

  // ② 从后往前找最后一处句界
  let cut = -1;
  for (let k = core.length - 1; k >= 0; k--) {
    if (SENT_END_RE.test(core[k])) { cut = k; break; }
  }
  const tail = core.slice(cut + 1).trim();
  if (!tail || !FUND_HINT_RE.test(tail)) return { abstract: text, funding: '' };

  const funding = (tail + punct).trim();
  const rest = body.slice(0, cut + 1).replace(/[\s\u3000]+$/, '');
  if (!rest) return { abstract: text, funding: '' };            // 摘完为空 → 不动
  return { abstract: rest, funding };
}

// ---- 「数值 + 空格 + 单位」的空格 ----
// 源表里数值与单位之间是普通空格（`厚度仅20 mm`、`（50 m、200 m、500 m）`）。
// TeX 里普通空格是可拉伸的 interword glue：中文段落两端对齐时会被拉得很宽，
// PDF 上就出现「厚度仅20    mm」「（50   m、200   m、500   m）」。
// 处理：直接删掉这个空格（20 mm → 20mm）。
// 若想保留一个**不可拉伸**的细空隙（GB/T 3101 的写法），把下面改成 '\\,'（\thinspace，1/6 em）即可。
const UNIT_SPACE = '';

// 只认白名单里的单位符号（长的排前面），避免把「3 A 级」这类非单位内容误改。
// 末尾的 (?![A-Za-z]) 保证 `20 meters` 里的 m 不会被误判成米。
const UNIT_SYMBOLS = [
  // 长度
  'mm', 'cm', 'dm', 'km', 'µm', 'μm', 'um', 'nm', 'pm', 'fm', 'm',
  // 质量
  'kg', 'mg', 'µg', 'μg', 'ug', 'ng', 'g', 't',
  // 时间
  'ms', 'µs', 'μs', 'us', 'ns', 'ps', 'fs', 's', 'min', 'h',
  // 频率
  'Hz', 'kHz', 'MHz', 'GHz', 'THz',
  // 压力 / 声级
  'Pa', 'kPa', 'MPa', 'GPa', 'hPa', 'dB', 'dBm', 'dBA',
  // 功率 / 能量
  'kW', 'MW', 'GW', 'mW', 'µW', 'μW', 'uW', 'W', 'kJ', 'MJ', 'mJ', 'J',
  // 电学
  'kV', 'mV', 'µV', 'μV', 'uV', 'V', 'mA', 'µA', 'μA', 'uA', 'A', 'kA',
  // 温度
  '°C', '℃', '°F', 'K',
  // 其它
  'mol', 'mmol', 'rad', 'sr', 'rpm', 'L', 'mL', 'µL', 'μL', 'uL', 'dL', 'cL',
  'KB', 'MB', 'GB', 'TB', 'kbps', 'Mbps', 'Gbps', 'kbit', 'Mbit', 'Gbit', 'bit', 'byte',
  'm/s', 'km/h', 'N', 'kN', 'mN',
].sort((a, b) => b.length - a.length);

const UNIT_ALT = UNIT_SYMBOLS
  .map(u => u.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
  .join('|');

// 空格类：半角空格 / 制表符 / 不换行空格 / 细空格 / 窄不换行空格 / 全角空格 / TeX 的 ~
const NUMBER_UNIT_RE = new RegExp(
  `(\\d)[ \\t\\u00A0\\u2009\\u202F\\u3000~]+(?=${UNIT_ALT}(?![A-Za-z]))`,
  'g'
);

/** 删掉「数值 单位」之间的可拉伸空格（20 mm → 20mm） */
function tightenNumberUnit(text) {
  if (!text) return text;
  return String(text).replace(NUMBER_UNIT_RE, `$1${UNIT_SPACE}`);
}

/**
 * 正文类文本（摘要 / 基金行）的公共转义步骤。
 * 顺序不可随意调换：先转义原始 LaTeX 特殊字符，再把 HTML 实体展开
 * （实体展开会引入 `$`、`\` 等），最后才处理孤立 `&` 与数值单位空格。
 */
function escapeBodyText(text) {
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
  t = t.replace(/(?<!\\)&/g, '\\&');
  // 第四步：去掉 CJK Ext-E 等 pdflatex 不支持的字符
  t = t.replace(/[\u{2A700}-\u{2F7FF}]/gu, '?');
  // 第五步：删掉「数值 单位」之间的空格（该空格会被两端对齐拉宽）
  t = tightenNumberUnit(t);
  return t;
}

/** 转义摘要文本（第零步先剥离作者误写的「摘要：」标签 / 重复的标题行） */
function sanitizeAbstract(text, rawTitle) {
  return escapeBodyText(stripAbstractLabel(s(text), rawTitle));
}

/** 转义基金 / 致谢句（页脚「基金项目」行）；与摘要同款转义，但不剥离标签 */
function sanitizeFunding(text) {
  return escapeBodyText(text);
}

/** 转义标题 */
function sanitizeTitle(text) {
  let t = s(text);
  t = t.replace(/%/g, '\\%')
       .replace(/_/g, '\\_')
       .replace(/#/g, '\\#')
       .replace(/~/g, '\\~{}');
  t = convertHtmlEntities(t);
  t = t.replace(/(?<!\\)&/g, '\\&');
  t = t.replace(/^[\u200B-\u200F\uFEFF]+/, '');
  t = t.replace(/<[^>]*>/g, '');
  t = tightenNumberUnit(t);
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
  t = t.replace(/(?<!\\)&/g, '\\&');
  t = tightenNumberUnit(t);
  t = t.split(',').map(k => k.trim()).filter(Boolean).join('，');
  return t;
}

/** 转义作者姓名 */
function sanitizeAuthorName(text) {
  return s(text).trim();
}

// ---- 全表反查表：缺失邮编的 A 级补齐依据 ----
// zipIndex:      `${单位名}||${地区}` → Set<邮编>   （只收有效邮编，排除空与全零占位）
// zipIndexCount: `${单位名}||${地区}` → 有效出现次数（仅用于报告）
// 必须在处理任何论文之前建好；且基于「原始配对结果」而非渲染串，避免污染。
const zipIndex = new Map();
const zipIndexCount = new Map();

function zipKey(unit, region) {
  return `${s(unit)}||${s(region)}`;
}

function buildZipIndex(rows) {
  for (const row of rows) {
    for (let i = 1; i <= 12; i++) {
      if (!sanitizeAuthorName(row[`作者姓名${i}`])) continue;
      const units = splitList(row[`作者单位${i}`]);
      const regions = splitList(row[`作者单位地区${i}`]);
      const zips = splitList(row[`作者单位邮编${i}`]);
      units.forEach((u, j) => {
        const region = regions[j] || '';
        const zip = zips[j] || '';
        if (isMissingZip(zip)) return;
        const k = zipKey(u, region);
        if (!zipIndex.has(k)) zipIndex.set(k, new Set());
        zipIndex.get(k).add(zip);
        zipIndexCount.set(k, (zipIndexCount.get(k) || 0) + 1);
      });
    }
  }
}

// ---- 主处理逻辑 ----

/**
 * 处理单篇论文，返回 LaTeX \printpaper 命令字符串。
 * @param {object} row    Excel 一行
 * @param {Array}  report 审计报告收集数组（按引用追加）
 */
function processPaper(row, report) {
  const session = s(row['论文专题']);
  const paperId = s(row['论文编号']);
  const title = sanitizeTitle(row['论文标题']);
  const keywords = sanitizeKeywords(row['论文关键字']);
  // 摘要末尾的基金 / 致谢句单独摘出，交给页脚渲染（正文里不再重复）
  const { abstract: abstractRaw, funding: fundingRaw } = extractFunding(row['摘要文本']);
  const abstract = sanitizeAbstract(abstractRaw, row['论文标题']);
  const funding = sanitizeFunding(fundingRaw);
  const corrAuthorName = sanitizeAuthorName(row['通讯作者姓名']);
  const corrAuthorEmail = sanitizeAuthorName(row['通讯作者Email']);

  // ── 1. 收集作者：三个平行列表按索引 j 一一配对 ──
  const authors = [];   // [{ name, pairs: [{ unit, region, zip, zipRaw, fillLevel }] }]
  for (let i = 1; i <= 12; i++) {
    const name = sanitizeAuthorName(row[`作者姓名${i}`]);
    if (!name) continue;

    const units   = splitList(row[`作者单位${i}`]);
    const regions = splitList(row[`作者单位地区${i}`]);
    const zips    = splitList(row[`作者单位邮编${i}`]);

    // 只有「地区数与单位数不符」或「邮编非空但与单位数不符」才是真正的配对异常；
    // 邮编整列为空（zips.length === 0）属常见情况，已由 A/B/C 三档逐条报告，不在此重复告警。
    const regionMismatch = units.length !== regions.length;
    const zipPartialMismatch = zips.length > 0 && zips.length !== units.length;
    if (regionMismatch || zipPartialMismatch) {
      report.push({
        type: '长度不一致', paperId, author: name,
        unit: units.join(' | '), region: regions.map(mapRegion).join(' | '),
        zipRaw: zips.join(' | '), zipFill: '',
        candCount: '', occCount: '',
        detail: `单位${units.length}/地区${regions.length}/邮编${zips.length}（按位置配对，多余忽略、缺失留空）`,
      });
    }

    authors.push({
      name,
      pairs: units.map((u, j) => ({
        unit: u,
        region: regions[j] || '',      // 越界 → 留空
        zip: zips[j] || '',            // 越界 → 留空
        zipRaw: zips[j] || '',         // 报告用：保留原值（含 000000）
        fillLevel: '',
      })),
    });
  }

  if (authors.length === 0) return null;

  // ── 2. 人工邮编对照表 zip_map.csv（最高优先级，命中即套用）──
  // 命中即无条件覆盖。但只把「原文缺失」或「与原值不同」的记入报告，
  // 否则同一单位的高频条目（如“中国科学院声学研究所”）会把报告刷成几百行。
  for (const a of authors) {
    for (const p of a.pairs) {
      const forced = lookupZipMap(p.unit, mapRegion(p.region));
      if (forced === null) continue;
      zipMapStats.hit++;
      const changed = s(p.zip) !== forced;
      p.zip = forced;
      p.fillLevel = 'M';
      if (!changed && !isMissingZip(p.zipRaw)) continue;   // 与原文完全一致 → 不必记报告
      zipMapStats.changed++;
      report.push({
        type: '已指定-zip_map', paperId, author: a.name,
        unit: p.unit, region: mapRegion(p.region),
        zipRaw: p.zipRaw, zipFill: forced,
        candCount: 1, occCount: '',
        detail: changed ? `人工对照表指定（原值「${p.zipRaw || '空'}」）` : '人工对照表指定（与原值一致）',
      });
    }
  }

  // ── 3. 缺失邮编补齐（必须在编号去重之前！）──
  for (const a of authors) {
    for (const p of a.pairs) {
      if (!isMissingZip(p.zip)) continue;              // 已有有效邮编，不动
      // 邮编反查用**原始**地区名（键更细、精度更高）；展示与去重用规范化后的地区名
      const k = zipKey(p.unit, p.region);
      const dispRegion = mapRegion(p.region);
      const regionNote = dispRegion === s(p.region) ? '' : `（原地区：${s(p.region)}）`;
      const set = zipIndex.get(k);
      const occ = zipIndexCount.get(k) || 0;

      if (set && set.size === 1) {
        // A 级：同「单位名+地区」在全表只有唯一有效邮编 → 安全回填
        p.zip = [...set][0];
        p.fillLevel = 'A';
        report.push({
          type: '已补齐-A级', paperId, author: a.name,
          unit: p.unit, region: dispRegion,
          zipRaw: p.zipRaw, zipFill: p.zip,
          candCount: set.size, occCount: occ,
          detail: '唯一映射' + regionNote,
        });
      } else if (set && set.size > 1) {
        p.fillLevel = 'B';
        p.zip = '';                                     // 占位/无效值清空，报告里保留 zipRaw
        report.push({
          type: '缺邮编-B级', paperId, author: a.name,
          unit: p.unit, region: dispRegion,
          zipRaw: p.zipRaw, zipFill: '',
          candCount: set.size, occCount: occ,
          detail: '候选不唯一：' + [...set].join(' / ') + regionNote,
        });
      } else {
        p.fillLevel = 'C';
        p.zip = '';                                     // 占位/无效值清空，报告里保留 zipRaw
        report.push({
          type: '缺邮编-C级', paperId, author: a.name,
          unit: p.unit, region: dispRegion,
          zipRaw: p.zipRaw, zipFill: '',
          candCount: 0, occCount: 0,
          detail: '全表无同「单位名+地区」的有效邮编记录' + regionNote,
        });
      }
    }
  }

  // 建立单位去重映射（按单位名称+地区+邮编联合去重）
  const unitMap = new Map();   // 联合键 → 编号
  const unitList = [];         // (按序) [{ id, text, region, zip }]
  let nextUnitId = 1;

  function getOrCreateUnitId(unitText, region, zip) {
    const key = `${unitText.trim()}||${region}||${zip}`;
    if (!unitText.trim() && !region && !zip) return null;
    if (unitMap.has(key)) return unitMap.get(key);
    const id = nextUnitId++;
    unitMap.set(key, id);
    unitList.push({ id, text: unitText.trim(), region, zip });
    return id;
  }

  // ── 4. 为每位作者按配对结果分配编号（去重 + 升序）──
  const authorEntries = authors.map(a => {
    const ids = [];
    for (const p of a.pairs) {
      const id = getOrCreateUnitId(p.unit, mapRegion(p.region), p.zip);
      if (id !== null) ids.push(id);
    }
    // 去重后升序，消除 \ns{5,2} 这类乱序
    const uniq = [...new Set(ids)].sort((x, y) => x - y);
    return {
      name: a.name,
      unitIds: uniq.length > 0 ? uniq : [0],
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

  // 构建单位 LaTeX 行：格式为 (编号 单位名　地区　邮编)，每行一个
  const unitLine = unitList.map(u => {
    const parts = [escapeLatex(u.text)];
    if (u.region) parts.push(escapeLatex(u.region));
    if (u.zip) parts.push(escapeLatex(u.zip));
    return `(${u.id} ${parts.join('\\quad ')})`;
  }).join('\\\\\n    ');

  const corrEmailLine = corrAuthorEmail
    ? `${escapeLatex(corrAuthorName)}，${escapeLatex(corrAuthorEmail)}`
    : '';

 
  // 标题折行：仍按原字符数切分（保持既有版式），切分后再插入断行点，
  // 避免 \allowbreak 被切断而破坏 LaTeX 语法
  const titleLine = softenBreaks(
    title.length > 45
      ? `${title.slice(0, Math.floor(title.length/2))}\\\\\n   ${title.slice(Math.floor(title.length/2))}`
      : title
  );
  const labelId = paperId.replace(/[^A-Za-z0-9]/g, '');
  // \printpaper 只有 9 个参数（TeX 宏参数上限为 9），因此不再单独传 labelId：
  // 锚点名直接用第 2 个参数（论文编号）。全库 432 条编号均为纯字母数字，
  // 若将来出现其他字符，锚点会与目录不一致，此处提前告警。
  if (labelId !== paperId) {
    console.warn(`  ⚠ 论文编号「${paperId}」含非字母数字字符，` +
                 `与目录锚点「paper:${labelId}」可能不一致，请检查。`);
  }
  return `\\printpaper
  {${session}}
  {${paperId}}
  {${titleLine}}
  {%
    ${softenBreaks(authorLine)}
  }
  {%
    ${softenBreaks(unitLine)}
  }
  {%
    ${softenBreaks(corrEmailLine)}
  }
  {%
    ${softenBreaks(abstract)}
  }
  {%
    ${softenBreaks(keywords)}
  }
  {${softenBreaks(funding)}}`;

  // ⚠ 第 9 个参数（基金）**必须写成单行** `{...}`，不能像上面那样写成
  //     {%
  //       <内容>
  //     }
  //   因为当内容为空时，那一行就成了空行，而 TeX 会把「空行」读成 `\par`，
  //   于是参数并非空串而是 `\par`，模板里的空值判断就会被骗过。
  //   全库只有这 4 篇有基金句，其余 428 篇的空参数必须是真正的 `{}`。
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

// 载入地区名映射表（region_map.csv，可选）：把「区/县」规范化成「市」
loadRegionMap();
console.log(regionMap.size > 0
  ? `Region map loaded: ${regionMap.size} rules.`
  : 'Region map: none (region_map.csv not found) - 地区名保持原样.');

// 载入人工邮编对照表（zip_map.csv，可选）
loadZipMap();
console.log(zipMap.size > 0
  ? `Zip map loaded: ${zipMap.size} rules.`
  : 'Zip map: none (zip_map.csv not found or empty).');

// 先建全表邮编反查表（缺失邮编的 A 级补齐依赖它）
buildZipIndex(data);
console.log(`Zip index built: ${zipIndex.size} distinct unit||region keys.`);

const report = [];
const papers = [];
for (const row of data) {
  const result = processPaper(row, report);
  if (result) papers.push(result);
}

console.log(`Processed ${papers.length} valid papers.`);

// ---- 写出审计报告 unit_report.csv（UTF-8 带 BOM，Excel 可直接打开）----
const REPORT_HEADER = ['类型', '论文编号', '作者', '单位', '地区', '邮编原值', '补齐值', '候选数', '出现次数', '备注'];
function csvEscape(v) {
  const t = (v === undefined || v === null) ? '' : String(v);
  return /[",\r\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
}
const reportLines = [REPORT_HEADER.join(',')];
for (const r of report) {
  reportLines.push([
    r.type, r.paperId, r.author, r.unit, r.region,
    r.zipRaw, r.zipFill, r.candCount, r.occCount, r.detail || '',
  ].map(csvEscape).join(','));
}
writeTextFile(REPORT_PATH, String.fromCharCode(0xFEFF) + reportLines.join('\r\n') + '\r\n');

const countBy = t => report.filter(r => r.type === t).length;
console.log(`Report written to ${REPORT_PATH}`);
console.log('  zip_map 命中 ' + zipMapStats.hit + ' 条（其中改写/补上 ' + countBy('已指定-zip_map') +
            ' 条）/ A 级补齐 ' + countBy('已补齐-A级') +
            ' 条 / B 级留空 ' + countBy('缺邮编-B级') +
            ' 条 / C 级留空 ' + countBy('缺邮编-C级') + ' 条 / 长度不一致 ' + countBy('长度不一致') + ' 条');

// 报告 zip_map.csv 中未被任何论文命中的规则（多半是单位名/地区名写法不一致）
if (zipMap.size > 0) {
  const unused = [...zipMap.keys()].filter(k => !zipMapUsed.has(k));
  if (unused.length === 0) {
    console.log(`  zip_map.csv：${zipMap.size} 条规则全部命中。`);
  } else {
    console.log(`  ⚠ zip_map.csv：${unused.length}/${zipMap.size} 条规则未命中，请检查单位名/地区名写法：`);
    for (const k of unused.slice(0, 20)) console.log('     · ' + k.replace('||', '  /  '));
    if (unused.length > 20) console.log(`     · …另有 ${unused.length - 20} 条`);
  }
}

// 生成 TOC（按专题分组，带引导符和超链接）
const TOC_PATH = path.join(__dirname, 'toc_content.tex');
const sessions = new Map();
for (const row of data) {
  const sid = s(row['论文编号']);
  const session = s(row['论文专题']);
  if (!sid) continue;
  const authorNames = [];
  for (let i = 1; i <= 12; i++) {
    const n = sanitizeAuthorName(row[`作者姓名${i}`]);
    if (n) authorNames.push(n);
  }
  const authorStr = authorNames.join(', ');
  const title = sanitizeTitle(row['论文标题']);
  if (!sessions.has(session)) sessions.set(session, []);
  sessions.get(session).push({ id: sid, title, authors: authorStr });
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
    // 版式（单行 / 标题行+姓名行 / 自然流动）由 main_full.tex 中
    // \tocentry 宏内用 \settowidth 实测宽度自动选择，此处无需判断
    // 锚点名与 \printpaper 的 \hypertarget{paper:#2} 保持完全一致（均为论文编号）
    tocLines.push(`\\tocentry{${e.id}}{${e.id}}{${titleEsc}}{${authEsc}}`);
  }
  tocLines.push('\\vspace{6pt}');
  tocLines.push('');
}

writeTextFile(TOC_PATH, tocLines.join('\n'));
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

writeTextFile(OUTPUT_PATH, output);
console.log(`Written to ${OUTPUT_PATH}`);




const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const wb = XLSX.readFile(path.join(__dirname, '摘要信息.xls'));
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, { defval: '' });

const out = [];
let total = 0, empty = 0;
const profile = {};   // 分隔符组合画像
const mixed = [];     // 混用多种分隔符的条目
const dunhao = [];    // 仅用 、 的条目
const fullOnly = [];  // 仅用 ，的条目

for (const row of data) {
  const id = (row['论文编号'] || '').toString().trim();
  if (!id) continue;
  const kw = (row['论文关键字'] || '').toString().trim();
  if (!kw) { empty++; continue; }
  total++;

  const hasHalf = kw.includes(',');
  const hasFull = kw.includes('，');
  const hasDun  = kw.includes('、');
  const hasSemi = /[;；]/.test(kw);

  const tags = [];
  if (hasHalf) tags.push('半角,');
  if (hasFull) tags.push('全角，');
  if (hasDun)  tags.push('顿号、');
  if (hasSemi) tags.push('分号;');
  const key = tags.join('+') || '(无分隔符)';
  profile[key] = (profile[key] || 0) + 1;

  const sepKinds = [hasHalf, hasFull, hasDun, hasSemi].filter(Boolean).length;
  if (sepKinds >= 2) mixed.push(`${id}\t[${key}]\t${kw}`);
  else if (hasDun) dunhao.push(`${id}\t${kw}`);
  else if (hasFull) fullOnly.push(`${id}\t${kw}`);
}

const head = [];
head.push(`有效关键词字段：${total}，空字段：${empty}`);
head.push('');
head.push('=== 分隔符画像（每篇只计一次）===');
for (const [k, v] of Object.entries(profile).sort((a,b)=>b[1]-a[1])) head.push(`  ${k}: ${v}`);

out.push(head.join('\n'));
out.push('');
out.push(`=== 混用多种分隔符（${mixed.length} 篇）===`);
out.push(...mixed);
out.push('');
out.push(`=== 仅用顿号、 分隔（${dunhao.length} 篇）===`);
out.push(...dunhao);
out.push('');
out.push(`=== 仅用全角，分隔（${fullOnly.length} 篇）===`);
out.push(...fullOnly);

fs.writeFileSync(path.join(__dirname, '_kw_audit.txt'), out.join('\n'), 'utf8');
console.log('done, total =', total, 'mixed =', mixed.length, 'dunhao =', dunhao.length, 'fullOnly =', fullOnly.length);

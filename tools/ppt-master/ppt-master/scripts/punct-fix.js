#!/usr/bin/env node
/**
 * ppt-master fullwidth → ASCII punctuation fixer
 *
 * 一次性脚本：把 work/ 下所有 .js 文件中的全角标点替换为 ASCII 等价。
 * （不再扫描 assets/，避免静默修改模板文件）
 * 这是单向操作（无备份），但 SKILL.md 已有提示用户用 GitHub 管理版本。
 *
 * 设计原则：与 minimax-pptx 的 lint 规则保持一致。
 *
 * Usage:
 *   node scripts/punct-fix.js
 *
 * Exit code:
 *   0 = success (no files or all files processed)
 */

const fs = require('fs');
const path = require('path');

const SKILL_ROOT = path.resolve(__dirname, '..');
const SCAN_DIRS = ['work'];

const MAPPING = {
  '\u201c': '"',   // LEFT DOUBLE QUOTATION MARK
  '\u201d': '"',   // RIGHT DOUBLE QUOTATION MARK
  '\u2018': "'",   // LEFT SINGLE QUOTATION MARK
  '\u2019': "'",   // RIGHT SINGLE QUOTATION MARK
  '\u3001': ',',   // IDEOGRAPHIC COMMA (、)
  '\u3002': '.',   // IDEOGRAPHIC FULL STOP (。)
  '\uff0c': ',',   // FULLWIDTH COMMA (，)
  '\uff1b': ';',   // FULLWIDTH SEMICOLON (；)
  '\uff1a': ':',   // FULLWIDTH COLON (：)
  '\uff01': '!',   // FULLWIDTH EXCLAMATION (！)
  '\uff1f': '?',   // FULLWIDTH QUESTION (？)
  '\uff08': '(',   // FULLWIDTH LEFT PAREN (（)
  '\uff09': ')',   // FULLWIDTH RIGHT PAREN (）)
  '\u2014': '--',  // EM DASH (—)
  '\u2013': '-',   // EN DASH (–)
  '\u2026': '...', // HORIZONTAL ELLIPSIS (…)
  '\u00a0': ' ',   // NO-BREAK SPACE
};

function walkDir(dir) {
  const fullDir = path.join(SKILL_ROOT, dir);
  if (!fs.existsSync(fullDir)) return [];

  const results = [];
  const entries = fs.readdirSync(fullDir, { withFileTypes: true });
  for (const entry of entries) {
    if (entry.name.startsWith('.')) continue;
    if (entry.name === 'node_modules') continue;  // NEVER modify dependencies
    const fullPath = path.join(fullDir, entry.name);
    if (entry.isDirectory()) {
      results.push(...walkDir(path.relative(SKILL_ROOT, fullPath)));
    } else if (entry.isFile() && entry.name.endsWith('.js')) {
      results.push(fullPath);
    }
  }
  return results;
}

function processFile(filePath) {
  const original = fs.readFileSync(filePath, 'utf8');
  let modified = original;
  const changes = {};

  for (const [from, to] of Object.entries(MAPPING)) {
    const count = (modified.match(new RegExp(from, 'g')) || []).length;
    if (count > 0) {
      modified = modified.split(from).join(to);
      changes[from] = { count, to };
    }
  }

  if (Object.keys(changes).length > 0) {
    fs.writeFileSync(filePath, modified, 'utf8');
    return { file: filePath, changes };
  }
  return null;
}

function main() {
  const files = SCAN_DIRS.flatMap((d) => walkDir(d));
  console.log(`Scanning ${files.length} JS file(s) in ${SCAN_DIRS.join(', ')}...`);

  let filesChanged = 0;
  let totalReplacements = 0;
  const summary = {};

  for (const file of files) {
    const result = processFile(file);
    if (result) {
      filesChanged++;
      const fileTotal = Object.values(result.changes).reduce((s, c) => s + c.count, 0);
      totalReplacements += fileTotal;
      const rel = path.relative(SKILL_ROOT, result.file);
      console.log(`  ${rel}: ${fileTotal} change(s)`);
      for (const [char, info] of Object.entries(result.changes)) {
        const label = `${char} -> ${info.to}`;
        summary[label] = (summary[label] || 0) + info.count;
      }
    }
  }

  console.log(`\nSummary:`);
  console.log(`  Files changed: ${filesChanged} / ${files.length}`);
  console.log(`  Total replacements: ${totalReplacements}`);
  if (Object.keys(summary).length > 0) {
    console.log(`  By character:`);
    const sorted = Object.entries(summary).sort((a, b) => b[1] - a[1]);
    for (const [k, v] of sorted) {
      console.log(`    ${k}: ${v}`);
    }
  }
  console.log(`\nDone. Run 'node scripts/lint.js' to verify.`);
}

main();

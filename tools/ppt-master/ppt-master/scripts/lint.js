#!/usr/bin/env node
/**
 * ppt-master pre-compile lint
 *
 * Scans assets/ and work/ for .js slide files and reports:
 *   1. Invisible characters (NBSP) — hard fail
 *   2. Smart quotes — warning
 *   3. Hex color values containing "#" (corrupts PptxGenJS)
 *   4. Opacity encoded into hex color strings (e.g. "00000020")
 *   5. async/await inside createSlide() (compile.js does not await)
 *   6. Theme object missing required 5 keys
 *
 * Usage:
 *   node scripts/lint.js
 *
 * Exit code:
 *   0 = pass, 1 = at least one error
 *
 * Reference: minimax-pptx/references/pitfalls.md (Pre-Compile Lint)
 *            ppt-master/references/common-errors.md (技术代码错误)
 */

const fs = require('fs');
const path = require('path');

const SKILL_ROOT = path.resolve(__dirname, '..');
const SCAN_DIRS = ['assets', 'work', 'references/cases/real-project-example'];

const RULES = [
  {
    id: 'invisible-spaces',
    level: 'error',
    description: '不可见空格 (NBSP, \\u00a0) 会破坏文本对齐和换行',
    fix: '替换为普通空格 (U+0020)，或运行 node scripts/punct-fix.js 自动修复',
    regex: /[\u00a0]/g,
  },
  {
    id: 'smart-quotes',
    level: 'warning',
    description: '智能引号 \u201c\u201d\u2018\u2019 — 大多数字体支持，但某些 fallback 会显示异常',
    fix: '运行 node scripts/punct-fix.js 自动替换为 ASCII 引号',
    regex: /[\u201c\u201d\u2018\u2019]/g,
  },
  {
    id: 'color-hash',
    level: 'error',
    description: 'Hex color with "#" prefix (corrupts PptxGenJS)',
    fix: '去掉 #，只留 6 位 hex 值',
    regex: /color\s*:\s*['"]#[0-9A-Fa-f]{6}['"]/g,
  },
  {
    id: 'hex-opacity',
    level: 'error',
    description: 'Opacity encoded in hex (e.g. "00000020"); use opacity property instead',
    fix: '改为 6 位 hex + 单独的 opacity 属性',
    regex: /['"][0-9A-Fa-f]{8}['"]/g,
  },
  {
    id: 'async-createSlide',
    level: 'error',
    description: 'async createSlide() — compile.js will not await it',
    fix: '移除 async 关键字，或改用同步写法',
    regex: /async\s+function\s+createSlide/g,
  },
  {
    id: 'theme-keys',
    level: 'error',
    description: 'Theme 对象缺少必需的 5 键 (primary, secondary, accent, light, bg)',
    fix: '确保 theme 对象包含全部 5 个键，参考 SKILL.md > Step 3',
    check(filePath, lines) {
      const src = lines.join('\n');
      // 检测所有 theme 对象定义（兼容 const / let / var，matchAll 支持多声明）
      const themeMatches = [...src.matchAll(/\b(?:const|let|var)\s+theme\s*=\s*\{([^}]+)\}/g)];
      if (themeMatches.length === 0) return [];
      const keys = ['primary', 'secondary', 'accent', 'light', 'bg'];
      const errors = [];
      for (const m of themeMatches) {
        const body = m[1];
        // 计算每个 theme 定义的实际行号
        const beforeMatch = src.substring(0, m.index);
        const matchLine = (beforeMatch.match(/\n/g) || []).length + 1;
        const missing = keys.filter(k => !new RegExp(`\\b${k}\\s*:`).test(body));
        for (const k of missing) {
          errors.push({
            file: filePath, line: matchLine, rule: 'theme-keys',
            description: `Theme 缺少 "${k}" 键`,
            match: `missing: ${k}`, context: `Theme 对象定义为 { ${body.trim().substring(0, 80)}... }`,
          });
        }
      }
      return errors;
    },
  },
];

const errors = [];
const warnings = [];

function scanFile(filePath) {
  const src = fs.readFileSync(filePath, 'utf8');
  const lines = src.split('\n');

  for (const rule of RULES) {
    // Support function-based checks (for non-regex rules like theme-keys)
    if (rule.check) {
      const results = rule.check(filePath, lines);
      for (const r of results) {
        const entry = { ...r, level: rule.level || 'error', fix: rule.fix || '' };
        if (rule.level === 'warning') warnings.push(entry);
        else errors.push(entry);
      }
      continue;
    }

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const matches = line.match(rule.regex);
      if (!matches) continue;

      for (const m of matches) {
        const entry = {
          file: filePath, line: i + 1, rule: rule.id,
          description: rule.description, match: m,
          context: line.trim().substring(0, 120),
          fix: rule.fix || '', level: rule.level || 'error',
        };
        if (rule.level === 'warning') warnings.push(entry);
        else errors.push(entry);
      }
    }
  }
}

function walkDir(dir) {
  const fullDir = path.join(SKILL_ROOT, dir);
  if (!fs.existsSync(fullDir)) return [];

  const results = [];
  const entries = fs.readdirSync(fullDir, { withFileTypes: true });
  for (const entry of entries) {
    if (entry.name.startsWith('.')) continue;
    if (entry.name === 'node_modules') continue;
    const fullPath = path.join(fullDir, entry.name);
    if (entry.isDirectory()) {
      results.push(...walkDir(path.relative(SKILL_ROOT, fullPath)));
    } else if (entry.isFile() && entry.name.endsWith('.js')) {
      results.push(fullPath);
    }
  }
  return results;
}

function printIssues(list, label) {
  if (list.length === 0) return;
  console.error(`\n${label}: ${list.length} issue(s)`);
  const byFile = new Map();
  for (const e of list) {
    if (!byFile.has(e.file)) byFile.set(e.file, []);
    byFile.get(e.file).push(e);
  }
  for (const [file, items] of byFile) {
    const rel = path.relative(SKILL_ROOT, file);
    console.error(`\n  ${rel}`);
    for (const e of items) {
      console.error(`    L${e.line}  [${e.rule}]  ${e.description}`);
      console.error(`           match: ${JSON.stringify(e.match)}`);
      console.error(`           context: ${e.context}`);
      if (e.fix) console.error(`           fix: ${e.fix}`);
    }
  }
}

function main() {
  const files = SCAN_DIRS.flatMap((d) => walkDir(d));

  if (files.length === 0) {
    console.log('No .js files found under assets/ or work/. Skipping lint.');
    return;
  }

  console.log(`Scanning ${files.length} JS file(s) in ${SCAN_DIRS.join(', ')}...`);

  for (const file of files) {
    scanFile(file);
  }

  printIssues(warnings, 'Warnings');
  printIssues(errors, 'Errors');

  if (errors.length === 0 && warnings.length === 0) {
    console.log('\nPre-compile lint passed.');
    process.exit(0);
  }

  if (errors.length > 0) {
    console.error(`\nLint FAILED: ${errors.length} error(s). Fix them before running compile.`);
    process.exit(1);
  }

  console.log(`\nLint PASSED (${warnings.length} warning(s) — review before delivery).`);
  process.exit(0);
}

main();

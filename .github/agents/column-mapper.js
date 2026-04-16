#!/usr/bin/env node
// .github/agents/column-mapper.js
// Scans repository .gs/.js/.html files for hard-coded numeric column indexes
// and generates a report with suggestions to add HEADERS_COLS mappings.

const fs = require('fs');
const path = require('path');
const lib = require('./lib');
const repoRoot = lib.repoRoot;
function listFiles(dir) { return lib.listFiles(dir); }

const IGNORE_PATH = path.join(__dirname, '.agents-ignore.json');
function loadIgnore() {
  if (!fs.existsSync(IGNORE_PATH)) return { files: [], patterns: [] };
  try {
    const j = JSON.parse(fs.readFileSync(IGNORE_PATH, 'utf8'));
    return {
      files: Array.isArray(j.files) ? j.files : [],
      patterns: Array.isArray(j.patterns) ? j.patterns.map(p => new RegExp(p)) : []
    };
  } catch (e) {
    console.error('Failed to parse ignore file:', e && e.message);
    return { files: [], patterns: [] };
  }
}

function scanFile(filePath, ignore) {
  const rel = path.relative(repoRoot, filePath);
  if (ignore.files && ignore.files.includes(rel)) return [];

  const txt = fs.readFileSync(filePath, 'utf8');
  const lines = txt.split(/\r?\n/);
  const findings = [];

  const patterns = [
    { name: 'getRange_numeric', re: /getRange\(\s*[^,]+,\s*\d+/ },
    { name: 'getRange_two_numeric', re: /getRange\(\s*\d+\s*,\s*\d+/ },
    { name: 'getValues_index', re: /getValues\(\)\s*\[\s*\d+\s*\]\s*\[\s*\d+\s*\]/ },
    { name: 'array_index', re: /\[\s*\d+\s*\](?=\s*\[|\s*;|\s*,|\s*\)|\s*$)/ }
  ];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    for (const p of patterns) {
      if (p.re.test(line)) {
        const skip = (ignore.patterns || []).some(rx => rx.test(line));
        if (skip) continue;
        findings.push({ line: i+1, text: line.trim(), pattern: p.name });
      }
    }
  }
  return findings;
}

function suggestHeaderName(file, lineText) {
  // Heuristic: try to extract a probable column variable name near the code
  // Fallback: generic suggestion including the numeric index found
  const m = lineText.match(/\[\s*(\d+)\s*\]/) || lineText.match(/getRange\(.*,(\s*\d+)/);
  const idx = m ? m[1].replace(/[^0-9]/g,'') : null;
  if (idx) return `SUGGESTED_HEADER_COL_${idx}`;
  return 'SUGGESTED_HEADER_NAME';
}

function run() {
  console.log('Scanning repository for numeric column index usages (column-mapper)...');
  const files = listFiles(repoRoot);
  const report = [];
  const ignore = loadIgnore();

  for (const f of files) {
    const rel = path.relative(repoRoot, f);
    try {
      const findings = scanFile(f, ignore);
      if (findings.length) {
        for (const fin of findings) {
          report.push({ file: rel, line: fin.line, code: fin.text, pattern: fin.pattern, suggestion: suggestHeaderName(rel, fin.text) });
        }
      }
    } catch (e) {
      console.error('Failed to scan', rel, e.message);
    }
  }

  const outDir = path.join(repoRoot, 'reports');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);

  const mdPath = path.join(outDir, 'column-index-report.md');
  const jsonPath = path.join(outDir, 'column-index-report.json');

  const now = new Date().toISOString();
  let md = `# Column Index Report\n\nGenerated: ${now}\n\n`;
  if (!report.length) {
    md += 'No obvious numeric column index usages were found. Good job!\n';
  } else {
    md += `Found ${report.length} potential usages:\n\n`;
    for (const r of report) {
      md += `- ${r.file}:${r.line} — pattern: ${r.pattern} — suggestion: ${r.suggestion}\n  \n    Code: \`${r.code.replace(/`/g,'\\`')}\`\n\n`;
    }
    md += '\n\nSuggested next steps:\n- For each occurrence, inspect and replace numeric index with resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.<SHEET>, CONFIG.COLUMNS.<SHEET>).\n- Add a descriptive key to CONFIG.HEADERS_COLS for the header text you will use.\n';
  }

  fs.writeFileSync(mdPath, md, 'utf8');
  fs.writeFileSync(jsonPath, JSON.stringify({ generated: now, items: report }, null, 2), 'utf8');

  console.log('Report written to', mdPath, jsonPath);
}

run();

#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const repoRoot = path.resolve(__dirname, '..', '..', '..');
const cwd = repoRoot;
const args = process.argv.slice(2);
const cmd = args[0] || 'help';

function runTests() {
  console.log('-> Executando testes (npm test)...');
  const res = spawnSync('npm', ['test'], { cwd, stdio: 'inherit', shell: true });
  if (res.error) {
    console.error('Erro ao executar testes:', res.error.message);
    process.exit(1);
  }
  process.exit(res.status || 0);
}

function listGsFiles(dir) {
  const out = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const e of entries) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) {
      if (['node_modules', '.git'].includes(e.name)) continue;
      out.push(...listGsFiles(full));
    } else {
      if (full.endsWith('.gs') || full.endsWith('.js') || full.endsWith('.html')) out.push(full);
    }
  }
  return out;
}

function findFixedColumns() {
  console.log('-> Procurando usos de índices numéricos de coluna em arquivos .gs/.js/.html...');
  const files = listGsFiles(cwd);
  const regex = /getRange\(\s*\d+\s*,\s*\d+/g;
  let found = 0;
  for (const f of files) {
    const txt = fs.readFileSync(f, 'utf8');
    const matches = txt.match(regex);
    if (matches && matches.length) {
      console.log(`  * ${path.relative(cwd, f)} -> ${matches.length} ocorrência(s)`);
      found += matches.length;
    }
  }
  if (!found) console.log('  Nenhum uso óbvio encontrado.');
  else console.log(`  Total: ${found} ocorrência(s)`);
  // Fail CI when occurrences found so the workflow can block merges
  if (found > 0) process.exit(2);
  process.exit(0);
}

function checkAyUsage() {
  console.log('-> Verificando ocorrências das siglas "AY" e menções a UUID no código...');
  const files = listGsFiles(cwd);
  let total = 0;
  for (const f of files) {
    const txt = fs.readFileSync(f, 'utf8');
    if (/\bAY\b/.test(txt) || /UUID/.test(txt) || /uuid/.test(txt)) {
      console.log(`  * ${path.relative(cwd, f)}`);
      total++;
    }
  }
  if (!total) console.log('  Nenhuma ocorrência encontrada.');
  // Fail CI if any occurrences are found
  if (total > 0) process.exit(2);
  process.exit(0);
}

function status() {
  console.log('Agent status:');
  console.log(' - Repo root: ', cwd);
  console.log(' - Node version:', process.version);
  console.log(' - Available commands: test, find-fixed-columns, check-ay, status, help');
}

function help() {
  console.log(`Uso: node .github\\agents\\run-agent.js <comando>

Comandos:
  test                 Executa os testes (npm test)
  find-fixed-columns   Procura usos de getRange com índices numéricos (potencialmente frágeis)
  check-ay             Lista arquivos que mencionam AY ou UUID
  status               Mostra informações básicas do ambiente
  help                 Mostra esta ajuda

Observação: o agente é conservador por padrão e não aplica mudanças automáticas.
`);
}

switch (cmd) {
  case 'test':
    runTests();
    break;
  case 'find-fixed-columns':
    findFixedColumns();
    break;
  case 'check-ay':
    checkAyUsage();
    break;
  case 'status':
    status();
    break;
  case 'help':
  default:
    help();
}

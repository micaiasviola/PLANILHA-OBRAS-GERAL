#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const repoRoot = path.resolve(__dirname, '..', '..');
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
  if (!total) {
    console.log('  Nenhuma ocorrência encontrada.');
  } else {
    console.warn('  Ocorrências encontradas — isso é apenas um aviso. Revise os arquivos listados para garantir a segurança das colunas AY/UUID.');
    // Do not fail the CI here; apenas avisar para revisão manual
  }
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

// --- Agent logging integration (append) ---
// When this process exits with a non-zero code, trigger `npm run agent-log`
// once, but avoid recursion by setting AGENT_LOGGER=1 in the child env.
(function setupAgentLogger() {
  if (!process || !process.on) return;
  try {
    process.on('exit', (code) => {
      try {
        // Only run the agent-log for non-zero exit codes (diagnostic runs)
        if (!code || code === 0) return;
        if (process.env && process.env.AGENT_LOGGER) return; // avoid loops
        console.log('-> Exit code', code, '- running npm run agent-log for diagnostics...');
        try {
          // Synchronous spawn is necessary inside exit handlers
          spawnSync('npm', ['run', 'agent-log'], {
            cwd,
            stdio: 'inherit',
            shell: true,
            env: Object.assign({}, process.env, { AGENT_LOGGER: '1' })
          });
        } catch (e) {
          try { console.error('agent-log failed:', e && e.message ? e.message : e); } catch (_) {}
        }
      } catch (_) {
        /* swallow */
      }
    });
  } catch (_) {}
})();
// --- end append ---

// --- Agent success logging (explicit) ---
// Run agent-log after successful runs when appropriate. Avoid recursion and duplicate logs.
(function runAgentSuccessLogger() {
  try {
    if (process.env && process.env.AGENT_LOGGER) return; // avoid loops
    // Check last commit message; if it's an agent-log commit, skip to avoid churn
    const res = spawnSync('git', ['log', '-1', '--pretty=%B'], { cwd, encoding: 'utf8', shell: true });
    const lastMsg = (res && res.stdout) ? res.stdout.toString().trim() : '';
    if (/^agent-log:/.test(lastMsg)) return;
    // Run agent-log to record a successful run. Set AGENT_LOGGER=1 to avoid recursion.
    console.log('-> Running agent-log to record successful run...');
    try {
      spawnSync('npm', ['run', 'agent-log', '--', 'COMMIT_PUSH', 'auto'], {
        cwd,
        env: Object.assign({}, process.env, { AGENT_NAME: 'CustomAgent', AGENT_LOGGER: '1' }),
        stdio: 'inherit',
        shell: true
      });
    } catch (e) {
      try { console.error('agent-log (success) failed:', e && e.message ? e.message : e); } catch (_) {}
    }
  } catch (e) {
    /* swallow errors */
  }
})();
// --- end success append ---

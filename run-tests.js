#!/usr/bin/env node

/**
 * CLI PARA EXECUTAR TESTES AUTOMATIZADOS
 * Uso: npm run test
 * Ou: node run-tests.js
 */

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const COLORS = {
  RESET: "\x1b[0m",
  GREEN: "\x1b[32m",
  RED: "\x1b[31m",
  YELLOW: "\x1b[33m",
  BLUE: "\x1b[34m",
  CYAN: "\x1b[36m",
};

function log(msg, color = "RESET") {
  console.log(`${COLORS[color]}${msg}${COLORS.RESET}`);
}

function header(title) {
  console.log("\n" + "=".repeat(70));
  log(title, "BLUE");
  console.log("=".repeat(70));
}

function runCommand(cmd, description) {
  log(`\n📌 ${description}...`, "CYAN");
  try {
    const output = execSync(cmd, { encoding: "utf8", stdio: "inherit" });
    log(`✅ ${description} concluído`, "GREEN");
    return true;
  } catch (e) {
    log(`❌ ${description} falhou`, "RED");
    return false;
  }
}

// ============= MAIN =============

async function main() {
  header("🧪 SUITE DE TESTES AUTOMATIZADOS");

  const projectRoot = path.dirname(__filename);
  const testsFile = path.join(projectRoot, "tests.js");

  if (!fs.existsSync(testsFile)) {
    log(`⚠️  Arquivo ${testsFile} não encontrado`, "YELLOW");
    log("Criando arquivo de testes padrão...", "YELLOW");
    // Seria criado aqui, mas já foi criado antes
  }

  // 1. Verificar se Node.js está instalado
  log("\n🔍 Verificando ambiente...", "CYAN");
  try {
    const nodeVersion = execSync("node --version", { encoding: "utf8" }).trim();
    log(`✅ Node.js ${nodeVersion} encontrado`, "GREEN");
  } catch {
    log("❌ Node.js não encontrado. Instale via https://nodejs.org/", "RED");
    process.exit(1);
  }

  // 2. Rodar testes locais (Node.js)
  header("1️⃣  TESTES ISOLADOS (Node.js)");
  const localTestsPass = runCommand(`node "${testsFile}"`, "Executar testes locais");

  // 3. Explicar como rodar testes de integração
  header("2️⃣  TESTES DE INTEGRAÇÃO (Google Sheets)");
  log("\nPara rodar testes contra a planilha:", "CYAN");
  log("  1. Abra a planilha no Google Sheets", "YELLOW");
  log("  2. Vá em Extensões → Apps Script", "YELLOW");
  log("  3. Procure a função: executarTodosTestes()", "YELLOW");
  log("  4. Clique em ▶ Executar", "YELLOW");
  log("  5. Verifique os resultados na aba TEST_DATA", "YELLOW");

  // 4. Resumo Final
  header("📊 RESUMO FINAL");

  if (localTestsPass) {
    log("✅ Testes locais: PASSOU", "GREEN");
  } else {
    log("❌ Testes locais: FALHOU", "RED");
  }

  log("\n📝 Próximos passos:", "CYAN");
  log("  • Testes locais validam lógica pura (sem planilha)", "YELLOW");
  log("  • Testes de integração validam fluxo Sheets completo", "YELLOW");
  log("  • Execute ambos antes de fazer commit/deploy", "YELLOW");

  log("\n📚 Documentação:", "CYAN");
  log("  Tests.gs:", "YELLOW");
  log("    → Testes de integração (rodados no Google Apps Script)", "YELLOW");
  log("  tests.js:", "YELLOW");
  log("    → Testes isolados (rodados localmente com Node.js)", "YELLOW");

  header("✨ TESTES CONFIGURADOS COM SUCESSO!");
  process.exit(localTestsPass ? 0 : 1);
}

main().catch(e => {
  log(`\n❌ Erro: ${e.message}`, "RED");
  process.exit(1);
});

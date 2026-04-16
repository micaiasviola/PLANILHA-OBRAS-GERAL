/*************************
 * MÓDULO: TESTES AUTOMATIZADOS
 * Executa suite de testes de integração contra a planilha
 *************************/

/**
 * Função raiz para executar todos os testes de integração.
 * Chame esta função do Google Apps Script para validar a automação.
 */
function executarTodosTestes() {
  limparCacheResolucaoColunas_();
  
  const resultados = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const linhasAdicionadas = { obra: [], pedidos: [] }; // Rastrear linhas para limpeza

  // 1. Preparar ambiente de teste
  const abaTest = obterOuCriarAbaTest_(ss);
  limparAbaTest_(abaTest);
  
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput("<p>🧪 Iniciando suite de testes...<br>Verifique a aba TEST_DATA</p>"),
    "Testes"
  );

  // 2. Suite 1: Sincronização Obra → Pedidos
  const resultSuite1 = testarSincronizacaoPedidosHousi_(abaTest, ss);
  resultados.push(...resultSuite1.resultados);
  linhasAdicionadas.obra.push(...resultSuite1.linhasAdicionadas.obra);
  linhasAdicionadas.pedidos.push(...resultSuite1.linhasAdicionadas.pedidos);

  // 3. Suite 2: Busca de Contato Fornecedor (corrigida)
  const resultSuite2 = testarBuscaContatoFornecedor_(abaTest, ss);
  resultados.push(...resultSuite2.resultados);
  linhasAdicionadas.pedidos.push(...resultSuite2.linhasAdicionadas.pedidos);

  // 4. Suite 3: Mapeamento Dinâmico de Colunas
  resultados.push(...testarMapeamentoDinamicoColunas_(abaTest, ss));

  // 5. Relatório Final
  gerarRelatorioTestes_(resultados, abaTest);
  
  // 6. Limpeza: Deletar linhas adicionadas em testes
  limparLinhasAdicionadasEmTestes_(ss, linhasAdicionadas);
  
  ss.toast("✅ Testes concluídos! Linhas de teste removidas.", "Suite de Testes", 10);
}

/**
 * Teste 1: Sincronização Obra → Pedidos para fornecedores HOUSI
 * Retorna { resultados: [], linhasAdicionadas: { obra: [], pedidos: [] } }
 */
function testarSincronizacaoPedidosHousi_(abaTest, ss) {
  const resultados = [];
  const linhasAdicionadas = { obra: [], pedidos: [] };
  const sheetObra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
  
  if (!sheetObra || !abaPedidos) {
    resultados.push({ nome: "Sincronização Obra→Pedidos", status: "SKIP", motivo: "Abas não encontradas" });
    return { resultados, linhasAdicionadas };
  }

  try {
    // Caso 1: Nova obra com HOUSI
    const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
    
    // Insere NO FINAL da aba (não no meio dos dados!)
    const lastObra = sheetObra.getLastRow();
    const novaLinhaObra = lastObra + 1;
    linhasAdicionadas.obra.push(novaLinhaObra); // Rastrear para limpeza
    
    // Array com dados até ATRELADO
    const dadosTesteObra = [
      ["MODERN BUNTANTÃ", "307", "", "", "", "Teste Serviço", "Categoria Teste", "Sub Teste", "HOUSI"]
    ];
    
    sheetObra.getRange(novaLinhaObra, 1, 1, dadosTesteObra[0].length).setValues(dadosTesteObra);
    
    // Dispara sincronização
    sincronizarPedidosHousiPorEdicao_({
      range: sheetObra.getRange(novaLinhaObra, 1, 1, C_OBRA.ATRELADO),
      source: ss
    });

    // Aguarda um pouco para a sincronização processar
    Utilities.sleep(500);
    
    // Verifica se linha foi criada em PEDIDOS
    const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
    const lastPed = abaPedidos.getLastRow();
    const lastPedAntes = lastPed;
    const dadosPedidos = abaPedidos.getRange(iniPed, 1, Math.max(1, lastPed - iniPed + 1), C_PED.EMP).getDisplayValues();
    
    const encontradoEmPedidos = dadosPedidos.some(r => 
      String(r[0] || "").includes("MODERN")
    );
    
    // Rastrear a linha criada em PEDIDOS
    if (encontradoEmPedidos) {
      linhasAdicionadas.pedidos.push(lastPed);
    }

    resultados.push({
      nome: "Sincronização Obra→Pedidos: Nova obra HOUSI sincronizada",
      status: encontradoEmPedidos ? "PASS" : "FAIL",
      detalhes: encontradoEmPedidos 
        ? "Linha criada em PEDIDOS-GERAL" 
        : `Linha não encontrada. Procurou em ${lastPed - iniPed + 1} linhas de PEDIDOS`
    });

    // Caso 2: Não sincroniza se ATRELADO ≠ HOUSI
    // Contar SKY ANTES de inserir
    const lastPed2Before = abaPedidos.getLastRow();
    const dadosPedidos2Before = abaPedidos.getRange(iniPed, 1, Math.max(1, lastPed2Before - iniPed + 1), C_PED.EMP).getDisplayValues();
    const linhasComSkyBefore = dadosPedidos2Before.filter(r => String(r[0] || "").includes("SKY")).length;
    
    const novaLinhaObra2 = lastObra + 2;
    linhasAdicionadas.obra.push(novaLinhaObra2); // Rastrear para limpeza
    const dadosTesteObra2 = [
      ["SKY PINHEIROS", "1501", "", "", "", "Teste Serviço 2", "Categoria Teste 2", "Sub Teste 2", "ECQUA"]
    ];
    
    sheetObra.getRange(novaLinhaObra2, 1, 1, dadosTesteObra2[0].length).setValues(dadosTesteObra2);
    
    sincronizarPedidosHousiPorEdicao_({
      range: sheetObra.getRange(novaLinhaObra2, 1, 1, C_OBRA.ATRELADO),
      source: ss
    });

    Utilities.sleep(500);
    
    // Contar SKY DEPOIS de sincronizar
    const lastPed2After = abaPedidos.getLastRow();
    const dadosPedidos2After = abaPedidos.getRange(iniPed, 1, Math.max(1, lastPed2After - iniPed + 1), C_PED.EMP).getDisplayValues();
    const linhasComSkyAfter = dadosPedidos2After.filter(r => String(r[0] || "").includes("SKY")).length;
    
    // Se a contagem não mudou, não sincronizou (correto!)
    const naoSincronizouNaoHousi = linhasComSkyAfter === linhasComSkyBefore;

    resultados.push({
      nome: "Sincronização Obra→Pedidos: Não sincroniza não-HOUSI",
      status: naoSincronizouNaoHousi ? "PASS" : "FAIL",
      detalhes: naoSincronizouNaoHousi 
        ? "OK - Não sincronizou (correto)" 
        : `Erro: Antes=${linhasComSkyBefore}, Depois=${linhasComSkyAfter}. Sincronizou ${linhasComSkyAfter - linhasComSkyBefore} linha(s) não-HOUSI`
    });

  } catch (e) {
    resultados.push({
      nome: "Sincronização Obra→Pedidos",
      status: "ERROR",
      motivo: e.message
    });
  }

  return { resultados, linhasAdicionadas };
}

/**
 * Teste 2: Busca de Contato Fornecedor (nova implementação)
 * Funciona com Backup com OU sem cabeçalhos
 * Retorna { resultados: [], linhasAdicionadas: { pedidos: [] } }
 */
function testarBuscaContatoFornecedor_(abaTest, ss) {
  const resultados = [];
  const linhasAdicionadas = { pedidos: [] };
  const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
  const abaBackup = obterAbaBackup_(ss);
  
  if (!abaPedidos || !abaBackup) {
    resultados.push({ nome: "Busca Contato Fornecedor", status: "SKIP", motivo: "Abas não encontradas" });
    return { resultados, linhasAdicionadas };
  }

  try {
    const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
    const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
    
    // Caso 1: Busca com fornecedor pré-existente (ex: DECOR FLOOR que está no Backup)
    const novaLinhaTest = iniPed + 2;
    linhasAdicionadas.pedidos.push(novaLinhaTest); // Rastrear para limpeza
    const fornecedorTeste = "DECOR FLOOR"; // Procura por este fornecedor
    
    abaPedidos.getRange(novaLinhaTest, C_PED.FORNECEDOR).setValue(fornecedorTeste);
    
    // Dispara busca
    buscarContatoFornecedor_({
      range: abaPedidos.getRange(novaLinhaTest, C_PED.FORNECEDOR, 1, 1),
      source: ss
    });

    const contatoObtido = String(abaPedidos.getRange(novaLinhaTest, C_PED.CONTATO).getDisplayValue()).trim();
    
    // Validação 1: Contato deve ser texto (não vazio e não data)
    const dataRegex = /^\d{1,2}\/\d{1,2}\/\d{4}|31\/12\/1969|^\d{4}-\d{2}-\d{2}/;
    const naoEhData = contatoObtido.length > 0 && !dataRegex.test(contatoObtido);
    const ehTextoValido = typeof contatoObtido === "string" && naoEhData;

    resultados.push({
      nome: "Busca Contato Fornecedor: Retorna texto válido",
      status: ehTextoValido ? "PASS" : "FAIL",
      detalhes: ehTextoValido 
        ? `Contato encontrado: "${contatoObtido}"` 
        : `Valor incorreto: "${contatoObtido}" (esperado: texto com telefone/contato)`
    });

    // Validação 2: Não retorna data (31/12/1969 era o bug antigo)
    const naoEhDataEpoch = !dataRegex.test(contatoObtido);

    resultados.push({
      nome: "Busca Contato Fornecedor: Não retorna data/epoch",
      status: naoEhDataEpoch ? "PASS" : "FAIL",
      detalhes: naoEhDataEpoch 
        ? "OK - Contato é texto, não data" 
        : `FALHA - Retornou data/hora: "${contatoObtido}"`
    });

  } catch (e) {
    resultados.push({
      nome: "Busca Contato Fornecedor",
      status: "ERROR",
      motivo: e.message
    });
  }

  return { resultados, linhasAdicionadas };
}

/**
 * Teste 3: Mapeamento Dinâmico de Colunas
 */
function testarMapeamentoDinamicoColunas_(abaTest, ss) {
  const resultados = [];
  const sheetObra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  
  if (!sheetObra) {
    resultados.push({ nome: "Mapeamento Dinâmico", status: "SKIP", motivo: "Aba FASE-OBRA não encontrada" });
    return resultados;
  }

  try {
    limparCacheResolucaoColunas_();
    
    const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    
    // Verifica se todas as colunas críticas foram encontradas
    const colunasObrigatorias = ["EMP", "UNI", "CHAVE", "ATRELADO", "CAT", "SUB"];
    const colunasResolvidas = colunasObrigatorias.filter(c => C_OBRA[c] > 0);
    
    resultados.push({
      nome: `Mapeamento Dinâmico de Colunas: ${colunasResolvidas.length}/${colunasObrigatorias.length}`,
      status: colunasResolvidas.length === colunasObrigatorias.length ? "PASS" : "FAIL",
      detalhes: `Encontradas: ${colunasResolvidas.join(", ")}`
    });

    // Verifica cache
    const C_OBRA_2 = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const ehMesmaReferencia = C_OBRA === C_OBRA_2;
    
    resultados.push({
      nome: "Mapeamento Dinâmico: Cache funciona",
      status: ehMesmaReferencia ? "PASS" : "FAIL",
      detalhes: ehMesmaReferencia ? "Cache ativo (mesma referência)" : "Cache não ativo (nova resolução cada vez)"
    });

  } catch (e) {
    resultados.push({
      nome: "Mapeamento Dinâmico",
      status: "ERROR",
      motivo: e.message
    });
  }

  return resultados;
}

/**
 * Cria ou obtém a aba de teste (TEST_DATA)
 */
function obterOuCriarAbaTest_(ss) {
  let abaTest = ss.getSheetByName("TEST_DATA");
  if (!abaTest) {
    abaTest = ss.insertSheet("TEST_DATA");
  }
  return abaTest;
}

/**
 * Limpa linhas que foram adicionadas durante os testes
 * Deleta em ordem reversa para não mudar índices
 */
function limparLinhasAdicionadasEmTestes_(ss, linhasAdicionadas) {
  try {
    // Deletar linhas de PEDIDOS primeiro (índices maiores tendem a ser afetados)
    if (linhasAdicionadas.pedidos && linhasAdicionadas.pedidos.length > 0) {
      const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
      if (abaPedidos) {
        // Ordena em ordem reversa para não mudar índices
        const linhasPedidosOrdenadas = linhasAdicionadas.pedidos.sort((a, b) => b - a);
        for (const linha of linhasPedidosOrdenadas) {
          abaPedidos.deleteRow(linha);
        }
      }
    }
    
    // Deletar linhas de OBRA (em ordem reversa)
    if (linhasAdicionadas.obra && linhasAdicionadas.obra.length > 0) {
      const abaObra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
      if (abaObra) {
        const linhasObraOrdenadas = linhasAdicionadas.obra.sort((a, b) => b - a);
        for (const linha of linhasObraOrdenadas) {
          abaObra.deleteRow(linha);
        }
      }
    }
    
    console.log("✅ Linhas de teste removidas com sucesso");
  } catch (e) {
    console.warn("⚠️ Erro ao limpar linhas de teste: " + e.message);
  }
}

/**
 * Limpa a aba de teste para nova execução
 */
function limparAbaTest_(abaTest) {
  const lastRow = abaTest.getLastRow();
  const lastCol = abaTest.getLastColumn();
  const TEST_START_ROW = 1;
  const TEST_START_COL = 1;
  if (lastRow > 0 && lastCol > 0) {
    abaTest.getRange(TEST_START_ROW, TEST_START_COL, lastRow, lastCol).clearContent();
  }
}

/**
 * Gera relatório de testes na aba TEST_DATA
 */
function gerarRelatorioTestes_(resultados, abaTest) {
  const headers = ["Teste", "Status", "Detalhes", "Data/Hora"];
  const dados = [headers];

  for (const r of resultados) {
    dados.push([
      r.nome,
      r.status,
      r.detalhes || r.motivo || "",
      new Date().toLocaleString("pt-BR")
    ]);
  }

  const TEST_START_ROW = 1;
  const TEST_START_COL = 1;
  const range = abaTest.getRange(TEST_START_ROW, TEST_START_COL, dados.length, 4);
  range.setValues(dados);

  // Formatação
  abaTest.getRange(TEST_START_ROW, TEST_START_COL, 1, 4).setFontWeight("bold").setBackground("#4CAF50").setFontColor("white");
  
  // Color by status
  for (let i = 1; i < dados.length; i++) {
    const status = dados[i][1];
    let bgColor = "#FFF9C4"; // Amarelo padrão
    if (status === "PASS") bgColor = "#C8E6C9"; // Verde
    if (status === "FAIL") bgColor = "#FFCDD2"; // Vermelho
    if (status === "ERROR") bgColor = "#FFCDD2"; // Vermelho
    if (status === "SKIP") bgColor = "#E0E0E0"; // Cinza
    
    abaTest.getRange(TEST_START_ROW + i, TEST_START_COL, 1, 4).setBackground(bgColor);
  }

  abaTest.autoResizeColumns(1, 4);

  // Resumo
  const passCount = resultados.filter(r => r.status === "PASS").length;
  const failCount = resultados.filter(r => r.status === "FAIL" || r.status === "ERROR").length;
  
  const resumo = `\n✅ PASS: ${passCount} | ❌ FAIL: ${failCount} | ⏭️ SKIP: ${resultados.filter(r => r.status === "SKIP").length}`;
  console.log(resumo);
}
/**
 * Basic tests for Payments module (merged from payments-tests.gs)
 */
function testarPagamentos() {
  try {
    // Smoke test: ensure sheet exists and functions callable
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Ensure PAGAMENTOS sheet exists (creates with canonical headers when missing)
    const sh = criarAbaPagamentosSimples();

    // Prepare a minimal test row using header mapping
    const C = resolvePaymentsColMap(sh);
    const cols = sh.getLastColumn();
    const row = new Array(cols).fill('');
    if (C.CHAVE_SERVICO >= 0) row[C.CHAVE_SERVICO] = 'TEST-KEY-123';
    if (C.PRESTADOR >= 0) row[C.PRESTADOR] = 'TESTE';
    if (C.VALOR >= 0) row[C.VALOR] = 100;
    if (C.TOTAL_SERVICO >= 0) row[C.TOTAL_SERVICO] = 100;

    // Append the test row, run validarSoma and then remove the row to clean up
    const before = sh.getLastRow();
    sh.appendRow(row);
    const res = validarSoma('TEST-KEY-123');
    Logger.log('Validar soma: ' + JSON.stringify(res));
    try { sh.deleteRow(before + 1); } catch (e) { Logger.log('Aviso: não foi possível remover linha de teste: ' + e.message); }

  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }
}

/*************************
 * MÓDULO: PEDIDOS-GERAL
 *************************/

/**
 * Handler disparado pelo router onEdit quando a aba PEDIDOS-GERAL é editada.
 */
function handlePedidosEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const C = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);

  // 1) Empreendimento -> Unidade (A -> B)
  if (intervaloInterceptaColuna(range, C.EMP)) {
    processarIntervaloAparaB_(sheet, range);
  }

  // 2) Busca Contato do Fornecedor (I -> J)
  if (intervaloInterceptaColuna(range, C.FORNECEDOR)) {
    buscarContatoFornecedor_(e);
  }

  // 3) Sincronização de Data Agendada (L) -> FASE-OBRA (N)
  if (intervaloInterceptaColuna(range, C.DATA_AGENDADO_ADM)) {
    sincronizarDataAgendadaAdmParaFaseObra_(e);
  }

  // 4) Sincronização de Status/Fornecedor (H, I) -> FASE-OBRA (J, K)
  if (intervaloInterceptaColuna(range, C.STATUS) || intervaloInterceptaColuna(range, C.FORNECEDOR)) {
    sincronizarPedidosParaFaseObra_(e);
  }
}

/**
 * Busca o contato do fornecedor na aba Backup e carimba na coluna J.
 */
function buscarContatoFornecedor_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const ss = e.source;
  const abaBackup = obterAbaBackup_(ss);
  if (!abaBackup) return;

  const C_PED = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);

  const lastBackup = abaBackup.getLastRow();
  if (lastBackup < 1) return;

  const linhaHeaderBackup = 1;
  const colFornecedorBackup = obterColunaPorCabecalho_(abaBackup, ["FORNECEDOR", "NOME FORNECEDOR", "NOME"], linhaHeaderBackup);
  const colContatoBackup = obterColunaPorCabecalho_(abaBackup, ["CONTATO", "CONTATO FORNECEDOR", "TELEFONE", "WHATSAPP"], linhaHeaderBackup);

  const colFornecedor = colFornecedorBackup > 0 ? colFornecedorBackup : 10;
  const colContato = colContatoBackup > 0 ? colContatoBackup : 11;
  const maxCol = Math.max(colFornecedor, colContato);

  // Cache em memória dos fornecedores para performance
  const dadosBackup = abaBackup.getRange(1, 1, lastBackup, maxCol).getValues();
  const mapaFornecedores = new Map();
  dadosBackup.forEach(r => {
    const nome = String(r[colFornecedor - 1] || "").trim().toUpperCase();
    if (nome) mapaFornecedores.set(nome, r[colContato - 1]);
  });

  const numRows = range.getNumRows();
  const rowStart = range.getRow();
  const fornecedores = sheet.getRange(rowStart, C_PED.FORNECEDOR, numRows, 1).getValues();
  const contatos = fornecedores.map(f => {
    const nome = String(f[0] || "").trim().toUpperCase();
    return [mapaFornecedores.get(nome) || ""];
  });

  if (C_PED.CONTATO) {
    sheet.getRange(rowStart, C_PED.CONTATO, numRows, 1).setValues(contatos);
  }
}

/**
 * Sincroniza a data agendada administrativa (L) de volta para a FASE-OBRA (M).
 */
function sincronizarDataAgendadaAdmParaFaseObra_(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const abaObra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!abaObra) return;

  const sheet = range.getSheet();
  const C_PED = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
  const C_OBRA = resolveSheetColumns_(abaObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  
  const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
  const rowStartAdjusted = Math.max(range.getRow(), iniPed);
  const numRowsAdjusted = range.getLastRow() - rowStartAdjusted + 1;
  if (numRowsAdjusted <= 0) return;

  // Lê dados de Pedidos em batch
  const maxColPed = Math.max(C_PED.CHAVE, C_PED.DATA_AGENDADO_ADM);
  const dadosPedidos = sheet.getRange(rowStartAdjusted, 1, numRowsAdjusted, maxColPed).getValues();

  // Mapa chave → novaData
  const mapaNovasDatas = new Map();
  for (let i = 0; i < numRowsAdjusted; i++) {
    const chaveID = String(dadosPedidos[i][C_PED.CHAVE - 1] || "").trim();
    if (!chaveID) continue;
    mapaNovasDatas.set(chaveID, dadosPedidos[i][C_PED.DATA_AGENDADO_ADM - 1] || null);
  }
  if (mapaNovasDatas.size === 0) return;

  // Lê toda a FASE-OBRA em batch (1 chamada)
  const iniObra = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  const lastObra = abaObra.getLastRow();
  if (lastObra < iniObra) return;

  const numRowsObra = lastObra - iniObra + 1;
  const maxColObra = Math.max(C_OBRA.CHAVE, C_OBRA.DATA_AGENDADO_ADM);
  const dadosObra = abaObra.getRange(iniObra, 1, numRowsObra, maxColObra).getValues();

  // Atualiza em memória
  let houveAlteracao = false;
  for (let i = 0; i < dadosObra.length; i++) {
    const chaveObra = String(dadosObra[i][C_OBRA.CHAVE - 1] || "").trim();
    if (chaveObra && mapaNovasDatas.has(chaveObra)) {
      dadosObra[i][C_OBRA.DATA_AGENDADO_ADM - 1] = mapaNovasDatas.get(chaveObra);
      houveAlteracao = true;
    }
  }

  // Grava em batch (1 chamada)
  if (houveAlteracao) {
    const colDatas = dadosObra.map(r => [r[C_OBRA.DATA_AGENDADO_ADM - 1]]);
    abaObra.getRange(iniObra, C_OBRA.DATA_AGENDADO_ADM, numRowsObra, 1).setValues(colDatas);
  }
}


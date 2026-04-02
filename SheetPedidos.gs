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
 * Funciona com ou sem cabeçalhos na aba Backup.
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

  // 1) Tenta buscar cabeçalhos nas primeiras 3 linhas
  const linhasBuscaHeader = [1, 2, 3];
  let colFornecedorBackup = obterColunaPorCabecalhoEmLinhas_(
    abaBackup,
    ["FORNECEDOR", "NOME FORNECEDOR", "NOME"],
    linhasBuscaHeader
  );
  let colContatoBackup = obterColunaPorCabecalhoEmLinhas_(
    abaBackup,
    ["CONTATO", "CONTATO FORNECEDOR", "TELEFONE", "WHATSAPP"],
    linhasBuscaHeader
  );

  // 2) Se não encontrou cabeçalho, assume Backup sem cabeçalho e usa índices fixos
  if (colFornecedorBackup <= 0 || colContatoBackup <= 0) {
    console.log("Aba Backup sem cabeçalhos detectada. Usando índices fixos: FORNECEDOR=col10, CONTATO=col11");
    colFornecedorBackup = 10;  // Coluna J (FORNECEDOR típico)
    colContatoBackup = 11;     // Coluna K (CONTATO típico)
  }

  const maxCol = Math.max(colFornecedorBackup, colContatoBackup);

  // 3) Cache em memória dos fornecedores para performance
  const dadosBackup = abaBackup.getRange(1, 1, lastBackup, maxCol).getDisplayValues();
  const mapaFornecedores = new Map();
  
  dadosBackup.forEach((r, idx) => {
    // Pula linhas de cabeçalho (se existirem) - detecta por padrão
    const fornecedorValor = String(r[colFornecedorBackup - 1] || "").trim();
    if (!fornecedorValor || idx < 3 && /FORNECEDOR|NOME/i.test(fornecedorValor)) {
      return; // Pula linha de cabeçalho
    }
    
    const nome = textoNormalizadoSemAcento_(fornecedorValor);
    if (nome) {
      const contato = String(r[colContatoBackup - 1] || "").trim();
      // Evita guardar valores que parecem ser datas
      if (!contato || /^\d{1,2}\/\d{1,2}\/\d{4}|^\d{4}-\d{2}-\d{2}/.test(contato)) {
        return; // Ignora contatos que são datas
      }
      mapaFornecedores.set(nome, contato);
    }
  });

  if (mapaFornecedores.size === 0) {
    console.warn("Nenhum fornecedor válido encontrado no Backup. Consulte a aba Backup e verifique dados.");
    return;
  }

  // 4) Buscar contatos para fornecedores em PEDIDOS
  const numRows = range.getNumRows();
  const rowStart = range.getRow();
  const fornecedores = sheet.getRange(rowStart, C_PED.FORNECEDOR, numRows, 1).getDisplayValues();
  const contatos = fornecedores.map(f => {
    const nome = textoNormalizadoSemAcento_(f[0]);
    return [mapaFornecedores.get(nome) || ""];
  });

  // 5) Gravar com formato texto (nunca deixa ser interpretado como data)
  if (C_PED.CONTATO) {
    const rangeContato = sheet.getRange(rowStart, C_PED.CONTATO, numRows, 1);
    rangeContato.setNumberFormat("@"); // Força formato texto
    rangeContato.setValues(contatos);
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


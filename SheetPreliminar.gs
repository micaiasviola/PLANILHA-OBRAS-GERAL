/*************************
 * MÓDULO: FASE-PRELIMINAR
 *************************/

/**
 * Handler disparado pelo router onEdit quando a aba FASE-PRELIMINAR é editada.
 */
function handlePreliminarEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const C = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);

  // 1) Empreendimento -> Unidade (A -> B)
  if (intervaloInterceptaColuna(range, C.EMP)) {
    processarIntervaloAparaB_(sheet, range);
  }

  // 2) Sincronização de Status Fase 00 (V)
  const colsFase00 = [C.DATA_VISTORIA, C.STATUS_VISTORIA, C.DATA_REVISTORIA_1, C.STATUS_REVISTORIA_1, C.DATA_REVISTORIA_2, C.STATUS_REVISTORIA_2];
  if (colsFase00.some(c => intervaloInterceptaColuna(range, c))) {
    sincronizarStatusFase00PreliminarPorEdicao_(e);
  }

  // 3) Reaplicar fórmula Status Fase 01 (AO)
  if (range.getRow() >= obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR)) {
    reaplicarFormulaStatusFase01PorEdicao_(e);
  }

  // 4) Resumo de Ocorrências e Pendências -> Sincroniza com INFORMAÇÕES GERAIS
  const colsSyncInfo = [C.EMP, C.UNI, C.RESUMO_PENDENCIAS, C.RESUMO_OCORRENCIAS, C.RESP_OPR];
  if (colsSyncInfo.some(c => intervaloInterceptaColuna(range, c))) {
    sincronizarInformacoesGeraisPorEdicaoPreliminar_(e);
  }
}

/**
 * Calcula e atualiza o Status da Fase 00 (Coluna V).
 */
function atualizarStatusFase00PreliminarPorIntervalo_(pre, primeiraLinha, numLinhas) {
  const C = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const maxCol = Math.max(C.DATA_VISTORIA, C.STATUS_VISTORIA, C.DATA_REVISTORIA_1, C.STATUS_REVISTORIA_1, C.DATA_REVISTORIA_2, C.STATUS_REVISTORIA_2);
  const dados = pre.getRange(primeiraLinha, 1, numLinhas, maxCol).getValues();
  const saida = [];

  for (let i = 0; i < dados.length; i++) {
    const row = dados[i];
    const vData = row[C.DATA_VISTORIA - 1];
    const vSt = row[C.STATUS_VISTORIA - 1];
    const r1Data = row[C.DATA_REVISTORIA_1 - 1];
    const r1St = row[C.STATUS_REVISTORIA_1 - 1];
    const r2Data = row[C.DATA_REVISTORIA_2 - 1];
    const r2St = row[C.STATUS_REVISTORIA_2 - 1];

    saida.push([calcularStatusFase00Preliminar_(vData, vSt, r1Data, r1St, r2Data, r2St)]);
  }

  pre.getRange(primeiraLinha, C.STATUS_FASE00, numLinhas, 1).setValues(saida);
}

/**
 * Disparador para a atualização da fórmula da Fase 01 (Coluna AO).
 */
function reaplicarFormulaStatusFase01PorEdicao_(e) {
  if (!e || !e.range) return;

  const pre = e.range.getSheet();
  const linhaInicialPre = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const primeiraLinha = Math.max(e.range.getRow(), linhaInicialPre);
  const numLinhas = e.range.getLastRow() - primeiraLinha + 1;
  
  if (numLinhas <= 0) return;

  aplicarFormulaStatusFase01PreliminarPorIntervalo_(pre, primeiraLinha, numLinhas);
}

/**
 * Lógica core de cálculo do Status da Fase 01.
 */
function aplicarFormulaStatusFase01PreliminarPorIntervalo_(pre, primeiraLinha, numLinhas) {
  if (!pre || numLinhas <= 0) return;

  const C = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const ultimaColControle = C.CHECKLIST_INI - 1; // Coluna antes do checklist
  const linhaHeader = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR) - 1;
  
  const dadosEmpUni = pre.getRange(primeiraLinha, 1, numLinhas, Math.max(C.EMP, C.UNI)).getDisplayValues();
  const controles = pre.getRange(primeiraLinha, 1, numLinhas, ultimaColControle).getDisplayValues();
  const checklist = pre.getRange(primeiraLinha, C.CHECKLIST_INI, numLinhas, C.CHECKLIST_FIM - C.CHECKLIST_INI + 1).getDisplayValues();
  const cabecalhosChecklist = pre.getRange(linhaHeader, C.CHECKLIST_INI, 1, C.CHECKLIST_FIM - C.CHECKLIST_INI + 1).getDisplayValues()[0];

  const reIgnorarCabecalho = /OBSERVAÇÕES GERAIS DOS ITENS VISTORIADOS|DESCRIÇÃO GERAL APONTAMENTOS COMPATIBILIZAÇÃO DE PROJETO/i;
  const saida = [];

  for (let i = 0; i < numLinhas; i++) {
    const emp = String(dadosEmpUni[i][C.EMP - 1] || "").trim();
    const uni = String(dadosEmpUni[i][C.UNI - 1] || "").trim();

    if (!emp && !uni) {
      saida.push([""]);
      continue;
    }

    // Se nao houver dados de controle (exceto o próprio status), mantém vazio.
    let possuiDadosControle = false;
    for (let c = 0; c < controles[i].length; c++) {
      const colAbsoluta = c + 1;
      if (colAbsoluta === C.STATUS_FASE01) continue;
      if (String(controles[i][c] || "").trim() !== "") {
        possuiDadosControle = true;
        break;
      }
    }

    if (!possuiDadosControle) {
      saida.push([""]);
      continue;
    }

    // Coleta pendências: Vazio, "VERIFICAR" ou "PENDENTE"
    const pendencias = [];
    for (let c = 0; c < cabecalhosChecklist.length; c++) {
      const cab = String(cabecalhosChecklist[c] || "").trim();
      if (!cab || reIgnorarCabecalho.test(cab)) continue;

      const valor = String(checklist[i][c] || "").trim().toUpperCase();
      if (!valor || valor === "VERIFICAR" || valor === "PENDENTE") {
        pendencias.push(cab);
      }
    }

    if (pendencias.length > 0) {
      saida.push(["Pendências a verificar: " + pendencias.join(", ")]);
    } else {
      saida.push(["FASE 01 CONCLUIDA"]);
    }
  }

  pre.getRange(primeiraLinha, C.STATUS_FASE01, numLinhas, 1).setValues(saida);
}

/**
 * Cria e configura a coluna FOTO TOMADA na aba Preliminar.
 * Chamada pelo menu Admin.
 */
function configurarColunaFotoTomadaPreliminar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  if (!pre) return;

  const C = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const linhaHeader = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR) - 1;
  let colFoto = C.FOTO_TOMADA;

  // Se não existir, tenta criar após "FOTOS APONTAMENTOS VISTORIA DE ENTRADA"
  if (colFoto <= 0 || colFoto === CONFIG.COLUMNS.PRELIMINAR.FOTO_TOMADA) {
    const headerAlvo = "FOTOS APONTAMENTOS VISTORIA DE ENTRADA";
    let colBase = obterColunaPorCabecalho_(pre, [headerAlvo], linhaHeader);
    
    // Se não achar o alvo específico, tenta RESP ADM como fallback
    if (colBase <= 0) colBase = Math.max(C.RESP_ADM, 10);

    pre.insertRowBefore(1);
    pre.insertColumnsAfter(colBase, 1);
    pre.deleteRow(1);
    
    colFoto = colBase + 1;
    pre.getRange(linhaHeader, colFoto).setValue("FOTO TOMADA");
    
    // Formata igual ao vizinho
    pre.getRange(linhaHeader, colBase).copyTo(pre.getRange(linhaHeader, colFoto), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const last = pre.getLastRow();
  if (last < ini) {
    SpreadsheetApp.getUi().alert("Coluna configurada, mas não há dados para aplicar validação.");
    return;
  }

  const range = pre.getRange(ini, colFoto, last - ini + 1, 1);
  const regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(["PENDENTE", "OK DRIVE", "NA"], true)
    .setAllowInvalid(false)
    .build();
  
  range.setDataValidation(regra);
  
  // Preenche vazios com PENDENTE
  const vals = range.getValues();
  const novos = vals.map(r => [r[0] || "PENDENTE"]);
  range.setValues(novos);

  SpreadsheetApp.getUi().alert("✅ Coluna FOTO TOMADA configurada com sucesso na posição " + colFoto);
}

/**
 * Cria e configura a coluna FASE-OBRA (Ativador manual) na aba Preliminar.
 */
function configurarColunaEnviarObraPreliminar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  if (!pre) return;

  const C = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const linhaHeader = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR) - 1;
  let colAlvo = C.ENVIAR_OBRA;

  // Se não existir, tenta criar após "FOTO TOMADA"
  if (colAlvo <= 0 || colAlvo === CONFIG.COLUMNS.PRELIMINAR.ENVIAR_OBRA) {
    const headerBase = "FOTO TOMADA";
    let colBase = obterColunaPorCabecalho_(pre, [headerBase], linhaHeader);
    
    if (colBase <= 0) colBase = Math.max(C.RESP_ADM, 10);

    pre.insertRowBefore(1);
    pre.insertColumnsAfter(colBase, 1);
    pre.deleteRow(1);
    
    colAlvo = colBase + 1;
    pre.getRange(linhaHeader, colAlvo).setValue("FASE-OBRA");
    pre.getRange(linhaHeader, colBase).copyTo(pre.getRange(linhaHeader, colAlvo), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const last = pre.getLastRow();
  if (last < ini) {
    SpreadsheetApp.getUi().alert("Coluna configurada sem dados.");
    return;
  }

  const range = pre.getRange(ini, colAlvo, last - ini + 1, 1);
  const regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(["SIM", "NÃO"], true)
    .setAllowInvalid(false)
    .build();
  
  range.setDataValidation(regra);
  
  // Preenche vazios com NÃO
  const vals = range.getValues();
  const novos = vals.map(r => [r[0] || "NÃO"]);
  range.setValues(novos);

  SpreadsheetApp.getUi().alert("✅ Coluna FASE-OBRA (Ativador Manual) configurada na posição " + colAlvo);
}

/**
 * Sincroniza dados editados na PRELIMINAR de volta para INFORMAÇÕES GERAIS.
 * Função chamada pelo handler onEdit de FASE-PRELIMINAR.
 */
function sincronizarInformacoesGeraisPorEdicaoPreliminar_(e) {
  if (!e || !e.range) return;

  const pre = e.range.getSheet();
  const C_PRE = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const primeiraLinha = Math.max(e.range.getRow(), linhaIni);
  const numLinhas = e.range.getLastRow() - primeiraLinha + 1;
  if (numLinhas <= 0) return;

  // Coleta as chaves EMP|UNI das linhas editadas
  const maxCol = Math.max(C_PRE.EMP, C_PRE.UNI);
  const dados = pre.getRange(primeiraLinha, 1, numLinhas, maxCol).getDisplayValues();
  const chavesAlvo = new Set();

  for (let i = 0; i < numLinhas; i++) {
    const emp = String(dados[i][C_PRE.EMP - 1] || "").trim().toUpperCase();
    const uni = String(dados[i][C_PRE.UNI - 1] || "").trim();
    if (emp && uni) chavesAlvo.add(`${emp}|${uni}`);
  }

  if (chavesAlvo.size > 0) {
    sincronizarInformacoesGeraisDesdePreliminar_(false);
  }
}

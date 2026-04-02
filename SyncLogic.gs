/*************************
 * MÓDULO: LÓGICA DE SINCRONIZAÇÃO (SyncLogic)
 *************************/

/**
 * Sincroniza dados da FASE-OBRA para PEDIDOS-GERAL (PEDIDO HOUSI).
 */
function sincronizarPedidosHousiPorEdicao_(e) {
  executarComDocumentLock_(function() {
    if (!e || !e.range) return;
    const sheetObra = e.range.getSheet();
    const rowStart = e.range.getRow();
    const numRows = e.range.getNumRows();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
    if (!abaPedidos) return;

    const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);

    // 1) Lê dados da obra em batch
    const dadosObra = sheetObra.getRange(rowStart, 1, numRows, C_OBRA.CHAVE).getValues();

    // 2) Lê TODOS os pedidos existentes em batch (1 chamada)
    const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
    const lastPed = abaPedidos.getLastRow();
    let dadosPedAll = [];
    const mapaPedIdx = new Map(); // chave → índice no array
    if (lastPed >= iniPed) {
      dadosPedAll = abaPedidos.getRange(iniPed, 1, lastPed - iniPed + 1, C_PED.CHAVE).getValues();
      for (let j = 0; j < dadosPedAll.length; j++) {
        const ch = String(dadosPedAll[j][C_PED.CHAVE - 1] || "").trim();
        if (ch) mapaPedIdx.set(ch, j);
      }
    }

    const linhasParaDeletar = [];
    const chavesGeradas = []; // [linhaObra, chaveID]
    let proximaLinhaLivre = -1;
    const linhasParaEscreverNoPedido = []; // [linhaDestinoSheet, arrayDados]

    for (let i = 0; i < numRows; i++) {
      const rowObra = rowStart + i;
      const vals = dadosObra[i];
      const atrelado = String(vals[C_OBRA.ATRELADO - 1]).trim().toUpperCase();
      const chaveID_Row = String(vals[C_OBRA.CHAVE - 1] || "").trim();

      if (atrelado !== "HOUSI") {
        if (chaveID_Row && mapaPedIdx.has(chaveID_Row)) {
          linhasParaDeletar.push(iniPed + mapaPedIdx.get(chaveID_Row));
          mapaPedIdx.delete(chaveID_Row);
        }
        continue;
      }

      const emp = String(vals[C_OBRA.EMP - 1]).trim();
      const uni = String(vals[C_OBRA.UNI - 1]).trim();
      const cat = String(vals[C_OBRA.CAT - 1]).trim();
      const sub = String(vals[C_OBRA.SUB - 1]).trim();
      if (!emp || !uni || !cat || !sub) continue;

      let chaveID = vals[C_OBRA.CHAVE - 1];
      if (!chaveID || String(chaveID).startsWith("FO_ROW_")) {
        chaveID = gerarUUID_();
        dadosObra[i][C_OBRA.CHAVE - 1] = chaveID;
        chavesGeradas.push([rowObra, chaveID]);
      }

      const chaveStr = String(chaveID).trim();
      let rowParaAtualizar;

      if (mapaPedIdx.has(chaveStr)) {
        // Dados existentes lidos do batch
        rowParaAtualizar = dadosPedAll[mapaPedIdx.get(chaveStr)].slice();
      } else {
        if (proximaLinhaLivre < 0) proximaLinhaLivre = obterPrimeiraLinhaLivrePedidos_(abaPedidos);
        rowParaAtualizar = new Array(C_PED.CHAVE).fill("");
      }

      // Atualiza campos em memória
      rowParaAtualizar[C_PED.EMP - 1] = emp;
      rowParaAtualizar[C_PED.UNI - 1] = uni;
      if (C_PED.OPR > 0) rowParaAtualizar[C_PED.OPR - 1] = vals[C_OBRA.OPR - 1];
      if (C_PED.ADM > 0) rowParaAtualizar[C_PED.ADM - 1] = vals[C_OBRA.ADM - 1];
      if (C_PED.TIPO > 0) rowParaAtualizar[C_PED.TIPO - 1] = vals[C_OBRA.TIPO - 1];
      if (C_PED.CAT > 0) rowParaAtualizar[C_PED.CAT - 1] = cat;
      if (C_PED.SUB > 0) rowParaAtualizar[C_PED.SUB - 1] = sub;
      if (C_PED.DATA_SOLICITADO_OPR > 0) rowParaAtualizar[C_PED.DATA_SOLICITADO_OPR - 1] = vals[C_OBRA.DATA_SOLICITADO_OPR - 1] || null;
      if (C_PED.DATA_AGENDADO_ADM > 0) rowParaAtualizar[C_PED.DATA_AGENDADO_ADM - 1] = vals[C_OBRA.DATA_AGENDADO_ADM - 1] || null;
      rowParaAtualizar[C_PED.CHAVE - 1] = chaveID;

      if (mapaPedIdx.has(chaveStr)) {
        // Atualiza no array existente (será gravado em batch)
        dadosPedAll[mapaPedIdx.get(chaveStr)] = rowParaAtualizar;
      } else {
        const linhaDest = proximaLinhaLivre++;
        linhasParaEscreverNoPedido.push([linhaDest, rowParaAtualizar]);
      }
    }

    // 3) Flush batch — chaves geradas na Obra
    if (chavesGeradas.length > 0) {
      for (const [row, chave] of chavesGeradas) {
        sheetObra.getRange(row, C_OBRA.CHAVE).setValue(chave);
      }
    }

    // 4) Flush batch — atualiza pedidos existentes (1 chamada)
    if (lastPed >= iniPed && dadosPedAll.length > 0) {
      abaPedidos.getRange(iniPed, 1, dadosPedAll.length, C_PED.CHAVE).setValues(dadosPedAll);
    }

    // 5) Flush — novas linhas de pedidos
    if (linhasParaEscreverNoPedido.length > 0) {
      for (const [linhaDest, dados] of linhasParaEscreverNoPedido) {
        garantirLinhasAte_(abaPedidos, linhaDest);
        abaPedidos.getRange(linhaDest, 1, 1, C_PED.CHAVE).setValues([dados]);
      }
    }

    // 6) Deleta órfãos de baixo para cima
    if (linhasParaDeletar.length > 0) {
      linhasParaDeletar.sort((a, b) => b - a).forEach(lp => abaPedidos.deleteRow(lp));
    }
  });
}

/**
 * Versão em lote (Bulk) da sincronização Obra -> Pedidos.
 */
function sincronizarTodosPedidosHousi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  limparCacheResolucaoColunas_(); 
  ss.toast("🔄 Iniciando sincronização em lote dos pedidos...", "Automação", 5);
  
  executarComDocumentLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetObra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
    const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
    if (!sheetObra || !abaPedidos) return;

    const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
    const linhaIniObra = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
    const linhaIniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
    
    const lastObra = sheetObra.getLastRow();
    if (lastObra < linhaIniObra) return;

    // Garante colunas de CHAVE
    garantirColunasAte_(sheetObra, C_OBRA.CHAVE);
    garantirColunasAte_(abaPedidos, C_PED.CHAVE);

    // 1) LÊ TUDO DA OBRA EM UMA REQUISIÇÃO
    const numLinhasObra = lastObra - linhaIniObra + 1;
    const dadosObra = sheetObra.getRange(linhaIniObra, 1, numLinhasObra, C_OBRA.CHAVE).getValues();
    
    // 2) LÊ TUDO DE PEDIDOS EM UMA REQUISIÇÃO
    const lastPed = abaPedidos.getLastRow();
    let dadosPed = [];
    const numRowsPed = lastPed - linhaIniPed + 1;
    if (numRowsPed > 0) {
        dadosPed = abaPedidos.getRange(linhaIniPed, 1, numRowsPed, C_PED.CHAVE).getValues();
    }
    
    // 3) INDEXA PEDIDOS (MAPA) POR CHAVE PARA BUSCA INSTANTÂNEA
    const mapaPedidos = new Map();
    for (let i = 0; i < dadosPed.length; i++) {
        const chave = String(dadosPed[i][C_PED.CHAVE - 1] || "").trim();
        if (chave) mapaPedidos.set(chave, i);
    }
    
    const listaFinalPedidos = [];
    let houveAlteracaoNaObra = false;

    // 4) PROCESSAMENTO NA MEMÓRIA - MONTA A LISTA NA ORDEM DA OBRA
    for (let i = 0; i < numLinhasObra; i++) {
      const vals = dadosObra[i];
      const atrelado = String(vals[C_OBRA.ATRELADO - 1]).trim().toUpperCase();

      if (atrelado !== "HOUSI") continue;

      const emp = String(vals[C_OBRA.EMP - 1]).trim();
      const uni = String(vals[C_OBRA.UNI - 1]).trim();
      const cat = String(vals[C_OBRA.CAT - 1]).trim();
      const sub = String(vals[C_OBRA.SUB - 1]).trim();
      if (!emp || !uni || !cat || !sub) continue;

      let chaveID = vals[C_OBRA.CHAVE - 1];
      
      // Auto-reparo de Chaves
      if (!chaveID || String(chaveID).startsWith("FO_ROW_")) {
        chaveID = gerarUUID_();
        dadosObra[i][C_OBRA.CHAVE - 1] = chaveID;
        houveAlteracaoNaObra = true;
      }

      const opr = vals[C_OBRA.OPR - 1];
      const adm = vals[C_OBRA.ADM - 1];
      const tipo = vals[C_OBRA.TIPO - 1];
      const dataSol = vals[C_OBRA.DATA_SOLICITADO_OPR - 1];
      const dataAge = vals[C_OBRA.DATA_AGENDADO_ADM - 1];

      let rowParaRegistrar;
      if (mapaPedidos.has(chaveID)) {
          const idx = mapaPedidos.get(chaveID);
          rowParaRegistrar = dadosPed[idx];
      } else {
          rowParaRegistrar = new Array(C_PED.CHAVE).fill("");
      }

      // Atualiza/Preenche
      rowParaRegistrar[C_PED.EMP - 1] = emp;
      rowParaRegistrar[C_PED.UNI - 1] = uni;
      if (C_PED.OPR > 0) rowParaRegistrar[C_PED.OPR - 1] = opr;
      if (C_PED.ADM > 0) rowParaRegistrar[C_PED.ADM - 1] = adm;
      if (C_PED.TIPO > 0) rowParaRegistrar[C_PED.TIPO - 1] = tipo;
      if (C_PED.CAT > 0) rowParaRegistrar[C_PED.CAT - 1] = cat;
      if (C_PED.SUB > 0) rowParaRegistrar[C_PED.SUB - 1] = sub;
      if (C_PED.DATA_SOLICITADO_OPR > 0) rowParaRegistrar[C_PED.DATA_SOLICITADO_OPR - 1] = dataSol;
      if (C_PED.DATA_AGENDADO_ADM > 0) rowParaRegistrar[C_PED.DATA_AGENDADO_ADM - 1] = dataAge;
      rowParaRegistrar[C_PED.CHAVE - 1] = chaveID;
      
      listaFinalPedidos.push(rowParaRegistrar);
    }

    // 5) FLUSH PARA O SHEETS (MÁXIMA PERFORMANCE + ORDENAÇÃO)
    if (houveAlteracaoNaObra) {
      const colunasChave = dadosObra.map(r => [r[C_OBRA.CHAVE - 1]]);
      sheetObra.getRange(linhaIniObra, C_OBRA.CHAVE, numLinhasObra, 1).setValues(colunasChave);
    }

    // Determina o limite da escrita respeitando o marcador
    const marcador = localizarLinhaMarcadorBase_(abaPedidos);
    const limite = marcador > 0 ? marcador - 1 : abaPedidos.getLastRow();
    const rowsParaLimpar = Math.max(0, limite - linhaIniPed + 1);

    if (rowsParaLimpar > 0) {
      abaPedidos.getRange(linhaIniPed, 1, rowsParaLimpar, C_PED.CHAVE).clearContent();
    }

    if (listaFinalPedidos.length > 0) {
       // Se o marcador existir, garantimos que temos espaço inserindo linhas se necessário
       // Mas como limpamos o conteúdo, podemos apenas escrever.
       // Se a lista final for MAIOR que o espaço antigo, precisamos inserir linhas antes do marcador.
       const diferenca = listaFinalPedidos.length - rowsParaLimpar;
       if (diferenca > 0 && marcador > 0) {
         abaPedidos.insertRowsBefore(marcador, diferenca);
       }
       
       abaPedidos.getRange(linhaIniPed, 1, listaFinalPedidos.length, C_PED.CHAVE).setValues(listaFinalPedidos);
    }

    ss.toast("✅ Sincronização de pedidos concluída!", "Automação", 5);
    SpreadsheetApp.getUi().alert("Sincronização concluída! Os pedidos agora estão perfeitamente alinhados com a FASE-OBRA.");
  }, 90000);
}

/**
 * Sincroniza especificamente a data editada na Obra para o pedido correspondente.
 */
function sincronizarDataPrevista_(e) {
  if (!e || !e.range) return;
  const sheetObra = e.range.getSheet();
  const rowStart = e.range.getRow();
  const numRows = e.range.getNumRows();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
  if (!abaPedidos) return;

  const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
  if (C_PED.DATA_SOLICITADO_OPR <= 0) return;

  // Lê todas as linhas editadas da obra em batch
  const maxColObra = Math.max(C_OBRA.CHAVE, C_OBRA.DATA_SOLICITADO_OPR);
  const dadosObra = sheetObra.getRange(rowStart, 1, numRows, maxColObra).getValues();

  // Mapa chave → novaData
  const mapaNovasDatas = new Map();
  for (let i = 0; i < numRows; i++) {
    const chaveID = String(dadosObra[i][C_OBRA.CHAVE - 1] || "").trim();
    if (!chaveID) continue;
    mapaNovasDatas.set(chaveID, dadosObra[i][C_OBRA.DATA_SOLICITADO_OPR - 1] || null);
  }
  if (mapaNovasDatas.size === 0) return;

  // Lê toda a coluna CHAVE + DATA dos Pedidos em batch
  const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
  const lastPed = abaPedidos.getLastRow();
  if (lastPed < iniPed) return;

  const numRowsPed = lastPed - iniPed + 1;
  const maxColPed = Math.max(C_PED.CHAVE, C_PED.DATA_SOLICITADO_OPR);
  const dadosPed = abaPedidos.getRange(iniPed, 1, numRowsPed, maxColPed).getValues();

  // Atualiza em memória
  let houveAlteracao = false;
  for (let i = 0; i < dadosPed.length; i++) {
    const chavePed = String(dadosPed[i][C_PED.CHAVE - 1] || "").trim();
    if (chavePed && mapaNovasDatas.has(chavePed)) {
      dadosPed[i][C_PED.DATA_SOLICITADO_OPR - 1] = mapaNovasDatas.get(chavePed);
      houveAlteracao = true;
    }
  }

  // Grava em batch (1 chamada)
  if (houveAlteracao) {
    const colDatas = dadosPed.map(r => [r[C_PED.DATA_SOLICITADO_OPR - 1]]);
    abaPedidos.getRange(iniPed, C_PED.DATA_SOLICITADO_OPR, numRowsPed, 1).setValues(colDatas);
  }
}

/**
 * Versão interna da sincronização para ser chamada em lote sem Lock redundante.
 */


/**
 * Atalho para carregar todos os pedidos em um Map por CHAVE.
 */
function obterMapaPedidosPorChave_(abaPedidos) {
  const mapa = new Map();
  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
  const last = abaPedidos.getLastRow();
  const C = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
  
  if (last >= ini) {
    const dados = abaPedidos.getRange(ini, C.CHAVE, last - ini + 1, 1).getValues();
    for (let i = 0; i < dados.length; i++) {
      const chave = String(dados[i][0]).trim();
      if (chave) mapa.set(chave, ini + i);
    }
  }
  return mapa;
}

/**
 * Localiza a linha de uma obra na aba FASE-OBRA usando a chave estável (ID).
 */
function localizarLinhaObraPorChave_(abaObra, chave) {
  if (!chave) return -1;
  const C = resolveSheetColumns_(abaObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const last = abaObra.getLastRow();
  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  if (last < ini) return -1;

  const chaves = abaObra.getRange(ini, C.CHAVE, last - ini + 1, 1).getValues();
  for (let i = 0; i < chaves.length; i++) {
    if (String(chaves[i][0]).trim() === String(chave).trim()) return ini + i;
  }
  return -1;
}

/**
 * Localiza a linha de um pedido na aba PEDIDOS-GERAL usando a chave estável (ID).
 */
function localizarLinhaPedidoPorChave_(abaPedidos, chave) {
  if (!chave) return -1;
  const C = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
  const last = abaPedidos.getLastRow();
  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
  if (last < ini) return -1;

  const chaves = abaPedidos.getRange(ini, C.CHAVE, last - ini + 1, 1).getValues();
  for (let i = 0; i < chaves.length; i++) {
    if (String(chaves[i][0]).trim() === String(chave).trim()) return ini + i;
  }
  return -1;
}

/**
 * Função para popular a validação de UNID com base no EMP.
 * Esta função é compartilhada por quase todas as abas.
 */
function processarIntervaloAparaB_(aba, intervalo) {
  const nomeAba = aba.getName();
  const rowStart = intervalo.getRow();
  const numRows = intervalo.getNumRows();
  const colEmp = intervalo.getColumn();
  const colUni = colEmp + 1;

  // REQUISITO: Apenas na aba INFORMAÇÕES GERAIS deve existir o menu suspenso de unidades.
  // Nas outras abas, limpamos qualquer validação para manter apenas o texto copiado.
  if (nomeAba !== CONFIG.SHEETS.INFO_GERAIS) {
    aba.getRange(rowStart, colUni, numRows, 1).clearDataValidations();
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBackup = obterAbaBackup_(ss);
  if (!abaBackup) return;

  // Carrega backup uma vez (todas as unidades cadastradas)
  const lastRowBackup = obterUltimaLinhaDados_(abaBackup, 1);
  const dadosBackup = lastRowBackup > 0 
    ? abaBackup.getRange(1, 1, lastRowBackup, 2).getValues() 
    : [];
  
  // Otimização: Carrega todos os empreendimentos do intervalo de uma vez
  const valsEmp = aba.getRange(rowStart, colEmp, numRows, 1).getDisplayValues();

  for (let i = 0; i < numRows; i++) {
    const row = rowStart + i;
    const empOriginal = String(valsEmp[i][0]).trim();
    const empBusca = textoNormalizadoSemAcento_(empOriginal);
    const cellUni = aba.getRange(row, colUni);

    if (!empBusca) {
      cellUni.clearDataValidations();
      continue;
    }

    // Filtra unidades que pertencem a este empreendimento (normalização total)
    const unidades = [...new Set(dadosBackup
      .filter(r => textoNormalizadoSemAcento_(r[0]) === empBusca)
      .map(r => String(r[1]).trim())
      .filter(u => u !== ""))];

    if (unidades.length > 0) {
      const regra = SpreadsheetApp.newDataValidation()
        .requireValueInList(unidades, true)
        .setAllowInvalid(true)
        .build();
      cellUni.setDataValidation(regra);
    } else {
      cellUni.clearDataValidations();
      // Caso não encontre no backup, pode ser um erro de nome ou novo empreendimento
      console.warn("Nenhuma unidade encontrada para: " + empOriginal + " na aba Backup.");
    }
  }
}

/**
 * Normalizações de dados para sincronização
 */
function normalizarValorTexto_(valor) {
  return String(valor || "").trim();
}

function normalizarValorDataOuTexto_(valor) {
  const dataNormalizada = normalizarDataSomenteDia_(valor);
  if (dataNormalizada) return dataNormalizada;
  return normalizarValorTexto_(valor);
}

/**
 * Função interna para obter a próxima linha livre na Fase Preliminar,
 * respeitando o marcador "PUXAR BASE".
 */
function obterProximaLinhaLivrePreliminar_(pre) {
  const C = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const last = obterUltimaLinhaDados_(pre, C.EMP);
  
  // Busca o marcador "PUXAR BASE"
  let rowMarcador = -1;
  if (last >= linhaIni) {
    const valsA = pre.getRange(linhaIni, 1, last - linhaIni + 1, 1).getDisplayValues();
    for (let i = 0; i < valsA.length; i++) {
      if (CONFIG.STATUS.MARCADOR_BASE.test(valsA[i][0])) {
        rowMarcador = linhaIni + i;
        break;
      }
    }
  }

  // Se houver marcador, insere antes dele
  if (rowMarcador > 0) return rowMarcador;
  
  // Se não, usa a primeira vazia ou o fim da aba
  return Math.max(last + 1, linhaIni);
}
function executarSincronizacaoFinalDoDia() {
  executarComDocumentLock_(function () {
    const S = CONFIG.SHEETS;
    // 1) Sincroniza INFORMAÇÕES GERAIS -> PRELIMINAR
    sincronizarPreliminarDesdeInformacoesGerais_(null, false);

    // 2) Sincroniza PEDIDOS -> OBRA (Status/Fornecedor)
    sincronizarPedidosParaFaseObraCompleta_(false);

    // 3) Consolida PRELIMINAR -> INFORMAÇÕES GERAIS
    sincronizarInformacoesGeraisDesdePreliminar_(false);
  }, 60000);
}

function sincronizarPedidosParaFaseObraCompleta_(exibirAlerta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
  if (!abaPedidos) return;

  // Lógica de sincronização em lote simplificada para o novo modelo
  const fakeEvent = {
    range: abaPedidos.getRange(obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS), 1, abaPedidos.getLastRow(), 9),
    source: ss
  };
  sincronizarPedidosParaFaseObra_(fakeEvent);
}

/**
 * Sincroniza dados das INFORMAÇÕES GERAIS para a FASE-PRELIMINAR.
 */
function sincronizarPreliminarDesdeInformacoesGerais_(chavesAlvo, exibirAlerta) {
  executarComDocumentLock_(function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
  const pre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  if (!info || !pre) return;

  const linhaInicialInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
  const lastInfo = info.getLastRow();
  if (lastInfo < linhaInicialInfo) return;

  const C_INFO = resolveSheetColumns_(info, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
  const maxColInfo = Math.max(
    C_INFO.EMP, C_INFO.UNI, C_INFO.DATA_LOTE, C_INFO.DATA_PRAZO,
    C_INFO.FASE_MACRO, C_INFO.PRIORIDADE, C_INFO.RESP_OPR, C_INFO.RESP_ADM
  );
  const registrosInfo = info.getRange(linhaInicialInfo, 1, lastInfo - linhaInicialInfo + 1, maxColInfo).getValues();

  const C_PRE = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const linhaInicialPre = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const linhaHeaderPre = linhaInicialPre - 1;

  // Mapeamento de colunas
  const mapeamento = {
    colDataLote: C_PRE.DATA_LOTE,
    colDataPrazo: C_PRE.DATA_PRAZO,
    colFaseMacro: obterColunaPorCabecalho_(pre, CONFIG.HEADERS.FASE_MACRO, linhaHeaderPre),
    colPrioridade: C_PRE.PRIORIDADE,
    colRespOpr: C_PRE.RESP_OPR,
    colRespAdm: C_PRE.RESP_ADM
  };

  // Mapeamento de Checklist (Dinâmico por Cabeçalho)
  const mapeamentoChecklist = {};
  const lastColPre = pre.getLastColumn();
  const cabecalhosPre = lastColPre > 0
    ? pre.getRange(linhaHeaderPre, 1, 1, lastColPre).getDisplayValues()[0]
    : [];

  for (const headerAlvo in CONFIG.PRELIMINAR_DEFAULTS) {
    const normalizadoAlvo = textoNormalizadoSemAcento_(headerAlvo);
    for (let c = 0; c < cabecalhosPre.length; c++) {
      if (textoNormalizadoSemAcento_(cabecalhosPre[c]) === normalizadoAlvo) {
        mapeamentoChecklist[c + 1] = CONFIG.PRELIMINAR_DEFAULTS[headerAlvo];
        break;
      }
    }
  }

  // Mapa de unidades existentes na Preliminar
  const lastPre = pre.getLastRow();
  const mapaPrePorChave = new Map();
  if (lastPre >= linhaInicialPre) {
    const dadosPre = pre.getRange(linhaInicialPre, 1, lastPre - linhaInicialPre + 1, 2).getDisplayValues();
    for (let i = 0; i < dadosPre.length; i++) {
      const emp = String(dadosPre[i][0] || "").trim().toUpperCase();
      const uni = String(dadosPre[i][1] || "").trim();
      if (emp && uni) mapaPrePorChave.set(`${emp}|${uni}`, linhaInicialPre + i);
    }
  }

  // ============ FASE 1: Identifica novas e existentes ============
  const unidadesNovas = []; // [{emp, uni, infoRow}]
  const atualizacoes = [];  // [{rowPre, infoRow, isNova}]
  const chavesProcessadas = new Set(); // Evita duplicatas de EMP|UNI dentro da mesma execução

  for (let i = 0; i < registrosInfo.length; i++) {
    const row = registrosInfo[i];
    const emp = String(row[C_INFO.EMP - 1] || "").trim();
    const uni = String(row[C_INFO.UNI - 1] || "").trim();
    if (!emp || !uni) continue;

    const chave = `${emp.toUpperCase()}|${uni}`;
    if (chavesAlvo && chavesAlvo.size > 0 && !chavesAlvo.has(chave)) continue;

    // Proteção contra duplicatas: se já processamos essa chave, pula
    if (chavesProcessadas.has(chave)) continue;
    chavesProcessadas.add(chave);

    const rowPre = mapaPrePorChave.get(chave);
    if (rowPre) {
      atualizacoes.push({ rowPre, infoRow: row, isNova: false });
    } else {
      unidadesNovas.push({ emp, uni, infoRow: row });
    }
  }

  // ============ FASE 2: Inserção em lote de novas linhas ============
  if (unidadesNovas.length > 0) {
    const qtd = unidadesNovas.length;
    pre.insertRowsBefore(linhaInicialPre, qtd);

    // Copia formato/validação da linha molde (1 vez para todas)
    const linhaMolde = linhaInicialPre + qtd;
    if (linhaMolde <= pre.getLastRow()) {
      const maxCols = pre.getMaxColumns();
      const origem = pre.getRange(linhaMolde, 1, 1, maxCols);
      const alvo = pre.getRange(linhaInicialPre, 1, qtd, maxCols);
      origem.copyTo(alvo, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      origem.copyTo(alvo, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }

    // Registra as novas linhas para atualização
    for (let n = 0; n < qtd; n++) {
      const rowPre = linhaInicialPre + n;
      atualizacoes.push({ rowPre, infoRow: unidadesNovas[n].infoRow, isNova: true });
    }

    // Atualiza o mapa (linhas existentes deslocaram para baixo)
    for (const [k, v] of mapaPrePorChave.entries()) {
      mapaPrePorChave.set(k, v + qtd);
    }
    // Atualiza referências das existentes no array de atualizações
    for (const a of atualizacoes) {
      if (!a.isNova) a.rowPre += qtd;
    }
  }

  // ============ FASE 3: Acumula escritas por coluna em memória ============
  // Agrupa escritas por coluna para gravar em batch
  const escritasPorColuna = new Map(); // colNum → Map(rowPre → valor)

  function registrarEscrita(rowPre, col, valor) {
    if (col <= 0) return;
    if (!escritasPorColuna.has(col)) escritasPorColuna.set(col, new Map());
    escritasPorColuna.get(col).set(rowPre, valor);
  }

  for (const a of atualizacoes) {
    const { rowPre, infoRow, isNova } = a;

    registrarEscrita(rowPre, 1, String(infoRow[C_INFO.EMP - 1] || "").trim());
    registrarEscrita(rowPre, 2, String(infoRow[C_INFO.UNI - 1] || "").trim());
    registrarEscrita(rowPre, mapeamento.colDataLote, infoRow[C_INFO.DATA_LOTE - 1]);
    registrarEscrita(rowPre, mapeamento.colDataPrazo, infoRow[C_INFO.DATA_PRAZO - 1]);
    registrarEscrita(rowPre, mapeamento.colFaseMacro, infoRow[C_INFO.FASE_MACRO - 1]);
    registrarEscrita(rowPre, mapeamento.colPrioridade, infoRow[C_INFO.PRIORIDADE - 1]);
    registrarEscrita(rowPre, mapeamento.colRespOpr, infoRow[C_INFO.RESP_OPR - 1]);
    registrarEscrita(rowPre, mapeamento.colRespAdm, infoRow[C_INFO.RESP_ADM - 1]);

    // Defaults do checklist (apenas novas)
    if (isNova) {
      for (const col in mapeamentoChecklist) {
        registrarEscrita(rowPre, Number(col), mapeamentoChecklist[col]);
      }
    }
  }

  // ============ FASE 4: Flush batch — 1 setValues por coluna ============
  for (const [col, mapaLinhas] of escritasPorColuna.entries()) {
    // Determina range contíguo mínimo
    const linhas = Array.from(mapaLinhas.keys()).sort((a, b) => a - b);
    const minRow = linhas[0];
    const maxRow = linhas[linhas.length - 1];
    const numRows = maxRow - minRow + 1;

    // Lê valores atuais para preservar linhas não-alvo
    const valoresAtuais = pre.getRange(minRow, col, numRows, 1).getValues();
    for (const [row, valor] of mapaLinhas.entries()) {
      valoresAtuais[row - minRow][0] = valor;
    }
    pre.getRange(minRow, col, numRows, 1).setValues(valoresAtuais);
  }

  // Limpa validações de unidade para aba Preliminar (não INFO_GERAIS)
  if (atualizacoes.length > 0) {
    const minRow = Math.min(...atualizacoes.map(a => a.rowPre));
    const maxRow = Math.max(...atualizacoes.map(a => a.rowPre));
    pre.getRange(minRow, 2, maxRow - minRow + 1, 1).clearDataValidations();
  }

  if (exibirAlerta) {
    SpreadsheetApp.getUi().alert("Sincronização concluída. Novas unidades criadas: " + unidadesNovas.length);
  }
  }); // fim executarComDocumentLock_
}
/**
 * Encontra a primeira linha vazia na aba Pedidos, respeitando o marcador.
 */
function obterPrimeiraLinhaLivrePedidos_(abaPedidos) {
    const linhaInicial = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
    const marcador = localizarLinhaMarcadorBase_(abaPedidos);

    const last = obterUltimaLinhaDados_(abaPedidos, CONFIG.COLUMNS.PEDIDOS.EMP);
    const limiteBusca = marcador > 0 ? marcador - 1 : Math.max(last, linhaInicial);

    if (limiteBusca >= linhaInicial) {
        const valoresA = abaPedidos.getRange(linhaInicial, 1, limiteBusca - linhaInicial + 1, 1).getDisplayValues();
        for (let i = 0; i < valoresA.length; i++) {
            if (String(valoresA[i][0] || "").trim() === "") return linhaInicial + i;
        }
    }

    if (marcador > 0) {
        abaPedidos.insertRowBefore(marcador);
        return marcador;
    }

    // fallback
    const final = Math.max(last + 1, linhaInicial);
    return final;
}

/**
 * Localiza a linha do marcador "PUXAR BASE".
 */
function localizarLinhaMarcadorBase_(aba) {
    const linhaInicial = obterLinhaInicialPorAba(aba.getName());
    const last = aba.getLastRow();
    if (last < linhaInicial) return -1;

    const valsA = aba.getRange(linhaInicial, 1, last - linhaInicial + 1, 1).getDisplayValues();
    for (let i = 0; i < valsA.length; i++) {
        if (CONFIG.STATUS.MARCADOR_BASE.test(String(valsA[i][0] || ""))) return linhaInicial + i;
    }
    return -1;
}

/**
 * Sincroniza em lote os envios da FASE-OBRA para a FASE-ENTREGA.
 */
function sincronizarTodosEnviosParaFaseEntrega() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("🚚 Verificando envios para FASE-ENTREGA...", "Automação", 5);
  
  executarComDocumentLock_(function() {
    const info = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
    if (!info) return;

    const C_INFO = resolveSheetColumns_(info, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
    const colChaveEntrega = C_INFO.ENVIAR_ENTREGA;

    if (colChaveEntrega <= 0) {
        SpreadsheetApp.getUi().alert(
            "Coluna de envio não encontrada na INFORMAÇÕES GERAIS.\n" +
            "Use o menu Admin para configurar o cabeçalho 'ENVIAR FASE-ENTREGA'."
        );
        return;
    }

    const linhaIniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
    const lastInfo = info.getLastRow();
    if (lastInfo < linhaIniInfo) return;

    const numLinhas = lastInfo - linhaIniInfo + 1;
    const numCols = Math.max(C_INFO.UNI, colChaveEntrega);
    const dados = info.getRange(linhaIniInfo, 1, numLinhas, numCols).getDisplayValues();

    const candidatos = [];
    for (let i = 0; i < dados.length; i++) {
        const emp = String(dados[i][C_INFO.EMP - 1]).trim();
        const uni = String(dados[i][C_INFO.UNI - 1]).trim();
        const valorChave = dados[i][colChaveEntrega - 1];

        if (!emp || !uni) continue;
        if (!deveEnviarParaFaseEntrega_(valorChave)) continue;

        candidatos.push([emp, uni]);
    }

    const antes = obterQuantidadeLinhasValidasNaEntrega_();
    inserirUnidadesNaFaseEntregaSemDuplicar_(candidatos);
    const depois = obterQuantidadeLinhasValidasNaEntrega_();
    const inseridas = Math.max(0, depois - antes);

    ss.toast("✅ Sincronização de envios concluída!", "Automação", 5);
    SpreadsheetApp.getUi().alert(
        "Sincronização de envios concluída.\n" +
        "Unidades identificadas para entrega: " + candidatos.length + "\n" +
        "Novas inseridas na FASE-ENTREGA: " + inseridas
    );
  }, 30000);
}

/**
 * Verifica se o valor de uma célula indica que deve ser enviado para entrega.
 */
function deveEnviarParaFaseEntrega_(valor) {
  if (!valor) return false;
  return CONFIG.STATUS.SIM_REGEX.test(String(valor).trim());
}

/**
 * Localiza dinamicamente a coluna de envio para entrega.
 */
function obterColunaChaveEntregaFaseObra_(abaObra) {
  return obterColunaPorCabecalho_(abaObra, CONFIG.HEADERS.ENTREGA_CHAVE, 1);
}

/**
 * Conta quantas linhas possuem empreendimento na aba de entrega.
 */
function obterQuantidadeLinhasValidasNaEntrega_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const entrega = ss.getSheetByName(CONFIG.SHEETS.ENTREGA);
    if (!entrega) return 0;

    const C_ENTREGA = resolveSheetColumns_(entrega, CONFIG.HEADERS_COLS.ENTREGA, CONFIG.COLUMNS.ENTREGA);
    const linhaInicial = obterLinhaInicialPorAba(CONFIG.SHEETS.ENTREGA);
    const last = entrega.getLastRow();
    if (last < linhaInicial) return 0;

    const valoresA = entrega.getRange(linhaInicial, C_ENTREGA.EMP, last - linhaInicial + 1, 1).getDisplayValues();
    let qtd = 0;
    for (let i = 0; i < valoresA.length; i++) {
        if (String(valoresA[i][0] || "").trim() !== "") qtd++;
    }
    return qtd;
}

/**
 * Insere novas unidades na aba de entrega, evitando duplicidades.
 */
function inserirUnidadesNaFaseEntregaSemDuplicar_(listaEmpUni) {
    if (!listaEmpUni || listaEmpUni.length === 0) return;

    executarComDocumentLock_(function () {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const entrega = ss.getSheetByName(CONFIG.SHEETS.ENTREGA);
        if (!entrega) return;

        const C_ENTREGA = resolveSheetColumns_(entrega, CONFIG.HEADERS_COLS.ENTREGA, CONFIG.COLUMNS.ENTREGA);
        const linhaInicialEntrega = obterLinhaInicialPorAba(CONFIG.SHEETS.ENTREGA);
        const lastEntrega = entrega.getLastRow();
        const mapaExistentes = new Set();

        if (lastEntrega >= linhaInicialEntrega) {
            const existentes = entrega.getRange(linhaInicialEntrega, C_ENTREGA.EMP, lastEntrega - linhaInicialEntrega + 1, 2).getDisplayValues();
            for (let i = 0; i < existentes.length; i++) {
                const emp = String(existentes[i][0]).trim().toUpperCase();
                const uni = String(existentes[i][1]).trim();
                if (emp && uni) mapaExistentes.add(`${emp}|${uni}`);
            }
        }

        const novas = [];
        for (let i = 0; i < listaEmpUni.length; i++) {
            const emp = String(listaEmpUni[i][0]).trim();
            const uni = String(listaEmpUni[i][1]).trim();
            if (!emp || !uni) continue;

            const chave = `${emp.toUpperCase()}|${uni}`;
            if (mapaExistentes.has(chave)) continue;

            mapaExistentes.add(chave);
            novas.push([emp, uni]);
        }

        if (novas.length === 0) return;

        const linhaInsercao = Math.max(entrega.getLastRow() + 1, linhaInicialEntrega);
        
        // Garante espaço
        garantirLinhasAte_(entrega, linhaInsercao + novas.length - 1);

        // Limpa e aplica
        const colEmp = C_ENTREGA.EMP;
        entrega.getRange(linhaInsercao, colEmp, novas.length, 2).setValues(novas);
        
        // Reaplica validações A->B se necessário
        processarIntervaloAparaB_(entrega, entrega.getRange(linhaInsercao, colEmp, novas.length, 1));
        
        // Sincroniza dados da Preliminar para essas novas linhas
        if (typeof sincronizarRespPreliminarParaEntregaCompleta_ === "function") {
          sincronizarRespPreliminarParaEntregaCompleta_(false);
        }
    });
}

/**
 * Remove pedidos da aba PEDIDOS-GERAL que não existem mais na FASE-OBRA.
 */
function removerPedidosOrfaos_(sheetObra, abaPedidos) {
  const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
  const iniObra = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
  
  const lastObra = sheetObra.getLastRow();
  const lastPed = abaPedidos.getLastRow();
  
  if (lastPed < iniPed) return;

  // 1. Obtém todas as chaves válidas na OBRA
  const setChavesObra = new Set();
  if (lastObra >= iniObra) {
    const chavesObra = sheetObra.getRange(iniObra, C_OBRA.CHAVE, lastObra - iniObra + 1, 1).getValues();
    for (let i = 0; i < chavesObra.length; i++) {
        const c = String(chavesObra[i][0]).trim();
        if (c) setChavesObra.add(c);
    }
  }

  // 2. Verifica cada linha de PEDIDOS (de baixo para cima para não quebrar o índice)
  const rangeChavesPed = abaPedidos.getRange(iniPed, C_PED.CHAVE, lastPed - iniPed + 1, 1);
  const chavesPedidos = rangeChavesPed.getValues();
  const marcador = localizarLinhaMarcadorBase_(abaPedidos);

  let removidos = 0;
  for (let i = chavesPedidos.length - 1; i >= 0; i--) {
      const rowPed = iniPed + i;
      
      // Regras de segurança: não remove se for a linha do marcador ou acima do inicial
      if (rowPed >= marcador && marcador > 0) continue; 

      const chavePed = String(chavesPedidos[i][0]).trim();
      
      // Se tiver chave mas ela não existe na OBRA, deleta do PEDIDOS
      if (chavePed && !setChavesObra.has(chavePed)) {
          abaPedidos.deleteRow(rowPed);
          removidos++;
      }
  }
  
  if (removidos > 0) {
      console.log("Removidos " + removidos + " pedidos órfãos da PEDIDOS-GERAL.");
  }
}

/**
 * Gatilho para atualizar o Status Geral de Ocorrências.
 */
function sincronizarStatusGeralOcorrenciasPorEdicao_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.OCORRENCIAS);
  const rowStart = Math.max(e.range.getRow(), linhaIni);
  const numRows = e.range.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  atualizarStatusGeralOcorrenciasPorIntervalo_(sheet, rowStart, numRows);
}

/**
 * Lógica de cálculo do Status Geral de Ocorrência.
 */
function calcularStatusGeralOcorrencia_(s1, d1, s2, d2, s3, d3) {
  const st1 = String(s1 || "").trim().toUpperCase();
  const st2 = String(s2 || "").trim().toUpperCase();
  const st3 = String(s3 || "").trim().toUpperCase();

  if (st3 === "CONCLUIDO" || st3 === "APROVADO") return "CONCLUÍDO";
  if (st3 === "CANCELADO") return "CANCELADO";
  if (d3) return "EM REVISTORIA 2";

  if (st2 === "CONCLUIDO" || st2 === "APROVADO") return "CONCLUÍDO";
  if (st2 === "CANCELADO") return "CANCELADO";
  if (d2) return "EM REVISTORIA 1";

  if (st1 === "CONCLUIDO" || st1 === "APROVADO") return "CONCLUÍDO";
  if (st1 === "CANCELADO") return "CANCELADO";
  if (d1) return "EM VISTORIA";

  return "ABERTO";
}

/**
 * Sincroniza a contagem de ocorrências abertas da aba OCORRÊNCIAS para a FASE-PRELIMINAR.
 */
function sincronizarOcorrenciasAbertasParaPreliminarPorEdicaoOcorrencias_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const rowStart = e.range.getRow();
  const numRows = e.range.getNumRows();
  const C = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.OCORRENCIAS, CONFIG.COLUMNS.OCORRENCIAS);
  
  const dados = sheet.getRange(rowStart, 1, numRows, 2).getValues();
  const chavesAlvo = new Set();
  for (let i = 0; i < numRows; i++) {
    const emp = String(dados[i][C.EMP - 1]).trim().toUpperCase();
    const uni = String(dados[i][C.UNI - 1]).trim();
    if (emp && uni) chavesAlvo.add(`${emp}|${uni}`);
  }

  if (chavesAlvo.size > 0) {
    sincronizarOcorrenciasAbertasParaPreliminar_(chavesAlvo, false);
  }
}

/**
 * Lógica em lote para contar ocorrências abertas e atualizar a aba alvo.
 */
function sincronizarOcorrenciasAbertasParaPreliminar_(chavesAlvo, exibirAlerta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaOco = ss.getSheetByName(CONFIG.SHEETS.OCORRENCIAS);
  const abaPre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  if (!abaOco || !abaPre) return;

  const C_OCO = resolveSheetColumns_(abaOco, CONFIG.HEADERS_COLS.OCORRENCIAS, CONFIG.COLUMNS.OCORRENCIAS);
  const C_PRE = resolveSheetColumns_(abaPre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  
  const iniOco = obterLinhaInicialPorAba(CONFIG.SHEETS.OCORRENCIAS);
  const lastOco = abaOco.getLastRow();
  if (lastOco < iniOco) return;

  const dadosOco = abaOco.getRange(iniOco, 1, lastOco - iniOco + 1, Math.max(C_OCO.EMP, C_OCO.UNI, C_OCO.STATUS_GERAL)).getValues();
  
  // Mapa de contagem por unidade
  const mapaContagem = new Map();
  for (let i = 0; i < dadosOco.length; i++) {
    const status = String(dadosOco[i][C_OCO.STATUS_GERAL - 1]).trim().toUpperCase();
    if (status !== "CONCLUÍDO" && status !== "CANCELADO") {
      const emp = String(dadosOco[i][C_OCO.EMP - 1]).trim().toUpperCase();
      const uni = String(dadosOco[i][C_OCO.UNI - 1]).trim();
      if (emp && uni) {
        const chave = `${emp}|${uni}`;
        mapaContagem.set(chave, (mapaContagem.get(chave) || 0) + 1);
      }
    }
  }

  // Atualiza na Preliminar
  const iniPre = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const lastPre = abaPre.getLastRow();
  if (lastPre < iniPre) return;

  const numRowsPre = lastPre - iniPre + 1;
  const maxColPre = Math.max(C_PRE.EMP || 0, C_PRE.UNI || 0, C_PRE.RESUMO_OCORRENCIAS || 0, 1);
  const rangePre = abaPre.getRange(iniPre, 1, numRowsPre, maxColPre);
  const dadosPre = rangePre.getValues();

  // Pré-carrega valores atuais da coluna RESUMO em memória (evita getValue individual)
  const colOcoIdx = C_PRE.RESUMO_OCORRENCIAS - 1;
  const saidaOco = [];

  for (let i = 0; i < dadosPre.length; i++) {
    const emp = String(dadosPre[i][C_PRE.EMP - 1]).trim().toUpperCase();
    const uni = String(dadosPre[i][C_PRE.UNI - 1]).trim();
    const chave = `${emp}|${uni}`;
    
    if (chavesAlvo && !chavesAlvo.has(chave)) {
      // Preserva valor atual lido em batch (antes: getValue individual)
      saidaOco.push([colOcoIdx >= 0 ? dadosPre[i][colOcoIdx] : ""]);
      continue;
    }

    const qtd = mapaContagem.get(chave) || 0;
    saidaOco.push([qtd > 0 ? qtd + " OCORRÊNCIA(S) ABERTAS" : ""]);
  }

  if (C_PRE.RESUMO_OCORRENCIAS > 0) {
    abaPre.getRange(iniPre, C_PRE.RESUMO_OCORRENCIAS, saidaOco.length, 1).setValues(saidaOco);
  } else {
    console.warn("Coluna RESUMO_OCORRENCIAS não encontrada em FASE-PRELIMINAR. Ignorando escrita.");
  }
  if (exibirAlerta) ss.toast("Sincronização de ocorrências concluída!", "Sucesso");
}

/**
 * Sincroniza dados da FASE-PRELIMINAR de volta para INFORMAÇÕES GERAIS.
 */
function sincronizarInformacoesGeraisDesdePreliminar_(exibirAlerta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  const info = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
  if (!pre || !info) return;

  const C_PRE = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const C_INFO = resolveSheetColumns_(info, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);

  const iniPre = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const lastPre = pre.getLastRow();
  if (lastPre < iniPre) return;

  const maxColPre = Math.max(C_PRE.EMP || 0, C_PRE.UNI || 0, C_PRE.RESUMO_PENDENCIAS || 0, C_PRE.RESUMO_OCORRENCIAS || 0, 1);
  const dadosPre = pre.getRange(iniPre, 1, lastPre - iniPre + 1, maxColPre).getValues();
  
  const mapaDados = new Map();
  for (let i = 0; i < dadosPre.length; i++) {
    const emp = String(dadosPre[i][C_PRE.EMP - 1]).trim().toUpperCase();
    const uni = String(dadosPre[i][C_PRE.UNI - 1]).trim();
    if (emp && uni) {
      mapaDados.set(`${emp}|${uni}`, {
        pendencias: C_PRE.RESUMO_PENDENCIAS > 0 ? dadosPre[i][C_PRE.RESUMO_PENDENCIAS - 1] : "",
        ocorrencias: C_PRE.RESUMO_OCORRENCIAS > 0 ? dadosPre[i][C_PRE.RESUMO_OCORRENCIAS - 1] : ""
      });
    }
  }

  const iniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
  const lastInfo = info.getLastRow();
  if (lastInfo < iniInfo) return;

  const rangeInfo = info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, Math.max(C_INFO.EMP, C_INFO.UNI));
  const dadosInfo = rangeInfo.getValues();
  const saidaPend = [];
  const saidaOco = [];

  for (let i = 0; i < dadosInfo.length; i++) {
    const emp = String(dadosInfo[i][C_INFO.EMP - 1]).trim().toUpperCase();
    const uni = String(dadosInfo[i][C_INFO.UNI - 1]).trim();
    const reg = mapaDados.get(`${emp}|${uni}`);
    
    saidaPend.push([reg ? reg.pendencias : ""]);
    saidaOco.push([reg ? reg.ocorrencias : ""]);
  }

  info.getRange(iniInfo, C_INFO.RESUMO_PENDENCIAS, saidaPend.length, 1).setValues(saidaPend);
  info.getRange(iniInfo, C_INFO.RESUMO_OCORRENCIAS, saidaOco.length, 1).setValues(saidaOco);

  if (exibirAlerta) ss.toast("Consolidação para INFORMAÇÕES GERAIS concluída!", "Sucesso");
}

/**
 * Cria o acionador para sincronização final do dia às 23:00.
 */
function criarAcionadorSincronizacaoFinalDoDia() {
  const HANDLER = "executarSincronizacaoFinalDoDia";

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === HANDLER)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(0)
    .create();

  SpreadsheetApp.getUi().alert("✅ Acionador de fechamento diário (23:00) criado com sucesso!");
}

/**
 * Sincroniza dados da aba PEDIDOS-GERAL de volta para a FASE-OBRA.
 */
function sincronizarPedidosParaFaseObra_(e) {
  executarComDocumentLock_(function() {
    if (!e || !e.range) return;
    const range = e.range;
    const abaPedidos = range.getSheet();
    const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    const abaObra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
    if (!abaObra) return;

    const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
    const C_OBRA = resolveSheetColumns_(abaObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
    
    // Calcula o intervalo de edição no Pedidos
    const rowStart = Math.max(range.getRow(), iniPed);
    const numRows = range.getLastRow() - rowStart + 1;
    if (numRows <= 0) return;

    // Obtém Status (I), Fornecedor (J) e ChaveID do Pedidos
    const dadosEdicao = abaPedidos.getRange(rowStart, 1, numRows, abaPedidos.getLastColumn()).getValues();
    
    // Mapeia chaves para novos valores vindos do Pedidos
    const mapaNovosDados = new Map();
    for (let i = 0; i < numRows; i++) {
        const row = dadosEdicao[i];
        const chave = String(row[C_PED.CHAVE - 1]).trim();
        if (chave) {
            mapaNovosDados.set(chave, {
                status: String(row[C_PED.STATUS - 1]).trim(),
                fornecedor: String(row[C_PED.FORNECEDOR - 1]).trim()
            });
        }
    }

    if (mapaNovosDados.size === 0) return;

    // Atualiza a FASE-OBRA baseada no Mapa
    const iniObra = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
    const lastObra = abaObra.getLastRow();
    if (lastObra < iniObra) return;

    const rangeObra = abaObra.getRange(iniObra, 1, lastObra - iniObra + 1, C_OBRA.CHAVE);
    const dadosObra = rangeObra.getValues();
    let houveAlteracao = false;

    for (let i = 0; i < dadosObra.length; i++) {
        const chaveObra = String(dadosObra[i][C_OBRA.CHAVE - 1]).trim();
        if (mapaNovosDados.has(chaveObra)) {
            const novos = mapaNovosDados.get(chaveObra);
            const statusAtual = String(dadosObra[i][C_OBRA.STATUS - 1]).trim();
            const fornecedorAtual = String(dadosObra[i][C_OBRA.FORNECEDOR - 1]).trim();

            if (statusAtual !== novos.status || fornecedorAtual !== novos.fornecedor) {
                // Atualiza em memória
                dadosObra[i][C_OBRA.STATUS - 1] = novos.status;
                dadosObra[i][C_OBRA.FORNECEDOR - 1] = novos.fornecedor;
                houveAlteracao = true;
            }
        }
    }

    if (houveAlteracao) {
        // Grava colunas independentes para evitar falha se não estiverem adjacentes
        const colStatus = dadosObra.map(r => [r[C_OBRA.STATUS - 1]]);
        const colFornecedor = dadosObra.map(r => [r[C_OBRA.FORNECEDOR - 1]]);
        abaObra.getRange(iniObra, C_OBRA.STATUS, dadosObra.length, 1).setValues(colStatus);
        abaObra.getRange(iniObra, C_OBRA.FORNECEDOR, dadosObra.length, 1).setValues(colFornecedor);
    }
  });
}

/**
 * Versão em lote da sincronização Pedidos -> Obra.
 */
function sincronizarPedidosParaFaseObraCompleta_(exibirAlerta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
  if (!abaPedidos) return;

  const fakeEvent = {
    range: abaPedidos.getRange(obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS), 1, abaPedidos.getLastRow(), 9),
    source: ss
  };
  sincronizarPedidosParaFaseObra_(fakeEvent);
  
  if (exibirAlerta) {
      SpreadsheetApp.getUi().alert("Reconciliação completa com Pedidos-Geral finalizada!");
  }
}

/**
 * Sincroniza RESP OPR e RESP ADM da FASE-PRELIMINAR para FASE-ENTREGA.
 */
function sincronizarRespPreliminarParaEntrega_(e) {
  if (!e || !e.range) return;
  const entrega = e.range.getSheet();
  if (entrega.getName() !== CONFIG.SHEETS.ENTREGA) return;

  const C_ENTREGA = resolveSheetColumns_(entrega, CONFIG.HEADERS_COLS.ENTREGA, CONFIG.COLUMNS.ENTREGA);
  
  // Confirma que a edição interceptou EMP ou UNI
  if (!intervaloInterceptaColuna(e.range, C_ENTREGA.EMP) && !intervaloInterceptaColuna(e.range, C_ENTREGA.UNI)) return;

  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const pre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  if (!pre) return;

  const C_PRE = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const rowStart = Math.max(e.range.getRow(), obterLinhaInicialPorAba(CONFIG.SHEETS.ENTREGA));
  const numRows = e.range.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  const dadosEntrega = entrega.getRange(rowStart, 1, numRows, Math.max(C_ENTREGA.EMP, C_ENTREGA.UNI)).getValues();
  
  // Mapa de Preliminar O(1)
  const mapaPre = new Map();
  const preIni = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const preLast = pre.getLastRow();
  
  if (preLast >= preIni) {
    const dadosPre = pre.getRange(preIni, 1, preLast - preIni + 1, Math.max(C_PRE.EMP, C_PRE.UNI, C_PRE.RESP_OPR, C_PRE.RESP_ADM)).getValues();
    for (let i = 0; i < dadosPre.length; i++) {
      const eText = String(dadosPre[i][C_PRE.EMP - 1] || "").trim().toUpperCase();
      const uText = String(dadosPre[i][C_PRE.UNI - 1] || "").trim();
      if (eText && uText) {
        mapaPre.set(`${eText}|${uText}`, {
          opr: String(dadosPre[i][C_PRE.RESP_OPR - 1] || "").trim(),
          adm: String(dadosPre[i][C_PRE.RESP_ADM - 1] || "").trim()
        });
      }
    }
  }

  if (mapaPre.size === 0) return;

  const valsOpr = [];
  const valsAdm = [];

  for (let i = 0; i < numRows; i++) {
    const emp = String(dadosEntrega[i][C_ENTREGA.EMP - 1] || "").trim().toUpperCase();
    const uni = String(dadosEntrega[i][C_ENTREGA.UNI - 1] || "").trim();
    
    const chave = `${emp}|${uni}`;
    if (mapaPre.has(chave)) {
      const resp = mapaPre.get(chave);
      valsOpr.push([resp.opr]);
      valsAdm.push([resp.adm]);
    } else {
      valsOpr.push([""]);
      valsAdm.push([""]);
    }
  }

  if (C_ENTREGA.RESP_OPR_PRE > 0) entrega.getRange(rowStart, C_ENTREGA.RESP_OPR_PRE, numRows, 1).setValues(valsOpr);
  if (C_ENTREGA.RESP_ADM_PRE > 0) entrega.getRange(rowStart, C_ENTREGA.RESP_ADM_PRE, numRows, 1).setValues(valsAdm);
}

/**
 * Versão em lote de sinc de RESP PRELIMINAR -> ENTREGA
 */
function sincronizarRespPreliminarParaEntregaCompleta_(exibirAlerta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrega = ss.getSheetByName(CONFIG.SHEETS.ENTREGA);
  if (!entrega) return;

  const last = entrega.getLastRow();
  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.ENTREGA);
  if (last < ini) return;

  const fakeEvent = {
    range: entrega.getRange(ini, 1, last - ini + 1, 1), // simula alcance englobando coluna A (EMP)
    source: ss
  };
  sincronizarRespPreliminarParaEntrega_(fakeEvent);

  if (exibirAlerta) {
    SpreadsheetApp.getUi().alert("Sincronização de responsáveis OPR/ADM para FASE-ENTREGA concluída!");
  }
}

/**
 * Gatilho de onEdit para calcular Status Fase 00 (Preliminar).
 */
function sincronizarStatusFase00PreliminarPorEdicao_(e) {
  if (!e || !e.range) return;
  const pre = e.range.getSheet();
  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const rowStart = Math.max(e.range.getRow(), linhaIni);
  const numRows = e.range.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  atualizarStatusFase00PreliminarPorIntervalo_(pre, rowStart, numRows);
}

/**
 * Lógica de cálculo do Status Fase 00.
 */
function calcularStatusFase00Preliminar_(vData, vSt, r1Data, r1St, r2Data, r2St) {
  const st1 = String(vSt || "").trim().toUpperCase();
  const st2 = String(r1St || "").trim().toUpperCase();
  const st3 = String(r2St || "").trim().toUpperCase();

  // Prioridade 1: Revistoria 2 (Última tentativa)
  if (st3 === "APROVADO") return "APROVADO";
  if (st3 === "REPROVADO") return "REPROVADO (FIM DE TENTATIVAS)";
  if (r2Data) return "EM REVISTORIA 2";

  // Prioridade 2: Revistoria 1
  if (st2 === "APROVADO") return "APROVADO";
  if (st2 === "REPROVADO") return "REPROVADO, AGUARDANDO MARCAÇÃO DA REVISTORIA 2";
  if (r1Data) return "EM REVISTORIA 1";

  // Prioridade 3: 1ª Vistoria
  if (st1 === "APROVADO") return "APROVADO";
  if (st1 === "REPROVADO") return "REPROVADO, AGUARDANDO MARCAÇÃO DA REVISTORIA 1";
  
  // Se houver qualquer data preenchida mas sem status final, está em processo
  if (vData) return "EM VISTORIA";

  return "PENDENTE";
}

/**
 * Gatilho de onEdit para calcular Status Geral de Entrega.
 */
function sincronizarStatusGeralVistoriaFinalPorEdicao_(e) {
  if (!e || !e.range) return;
  const entrega = e.range.getSheet();
  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.ENTREGA);
  const rowStart = Math.max(e.range.getRow(), linhaIni);
  const numRows = e.range.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  atualizarStatusGeralVistoriaFinalPorIntervalo_(entrega, rowStart, numRows);
}

/**
 * Lógica de cálculo do Status Geral de Entrega (Vistoria Final).
 */
function calcularStatusGeralVistoriaFinal_(s1, d1, sr1, dr1, sr2, dr2) {
  const st1 = String(s1 || "").trim().toUpperCase();
  const st2 = String(sr1 || "").trim().toUpperCase();
  const st3 = String(sr2 || "").trim().toUpperCase();

  // Prioridade 1: Revistoria 2
  if (st3 === "APROVADO") return "APROVADO";
  if (st3 === "REPROVADO") return "REPROVADO (FIM DE TENTATIVAS)";
  if (dr2) return "EM REVISTORIA 2";

  // Prioridade 2: Revistoria 1
  if (st2 === "APROVADO") return "APROVADO";
  if (st2 === "REPROVADO") return "REPROVADO, AGUARDANDO REVISTORIA 2";
  if (dr1) return "EM REVISTORIA 1";

  // Prioridade 3: 1ª Vistoria
  if (st1 === "APROVADO") return "APROVADO";
  if (st1 === "REPROVADO") return "REPROVADO, AGUARDANDO REVISTORIA 1";
  
  if (d1) return "EM VISTORIA";

  return "PENDENTE";
}

/**
 * Obtém um mapa de Chave (EMP|UNI) -> Data Lote a partir da aba Preliminar.
 */
function obterMapaDataLote_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const abaPre = spreadsheet.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  const mapa = new Map();
  if (!abaPre) return mapa;

  const C = resolveSheetColumns_(abaPre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
  const last = abaPre.getLastRow();
  if (last < ini) return mapa;

  const dados = abaPre.getRange(ini, 1, last - ini + 1, Math.max(C.EMP, C.UNI, C.DATA_LOTE)).getValues();
  for (let i = 0; i < dados.length; i++) {
    const emp = String(dados[i][C.EMP - 1]).trim().toUpperCase();
    const uni = String(dados[i][C.UNI - 1]).trim();
    const lote = dados[i][C.DATA_LOTE - 1];
    if (emp && uni && lote instanceof Date) {
      mapa.set(`${emp}|${uni}`, lote);
    }
  }
  return mapa;
}

/**
 * [COMBINADA - PERFORMANCE] Calcula e grava SEMANA CRONOGRAMA e SEMANA DO MÊS
 * em uma única execução: 1 Lock, 1 leitura da PRELIMINAR, 1 leitura da OBRA.
 * Chamada pelo onEdit e pelos wrappers de recalculo em lote.
 *
 * @param {Object} e  - Evento do onEdit (ou fake event para lote).
 * @param {boolean} calcCrono   - Se true, calcula SEMANA CRONOGRAMA.
 * @param {boolean} calcMes     - Se true, calcula SEMANA DO MÊS.
 */
function sincronizarSemanasObraCombinada_(e, calcCrono, calcMes) {
  executarComDocumentLock_(function() {
    if (!e || !e.range) return;
    const sheetObra = e.range.getSheet();
    const rowStart  = e.range.getRow();
    const numRows   = e.range.getNumRows();
    const ss        = SpreadsheetApp.getActiveSpreadsheet();

    const C = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);

    // Valida colunas necessárias
    const precisaCrono = calcCrono && C.DATA_INICIO_PLANEJADO > 0 && C.SEMANA > 0;
    const precisaMes   = calcMes   && C.DATA_INICIO_PLANEJADO > 0 && C.SEMANA_MES > 0;
    if (!precisaCrono && !precisaMes) return;

    // --- 1 LEITURA: PRELIMINAR (só necessária para SEMANA CRONOGRAMA) ---
    const dataLoteMap = precisaCrono ? obterMapaDataLote_(ss) : new Map();

    // --- 1 LEITURA: FASE-OBRA ---
    const numCols = Math.max(C.EMP, C.UNI, C.DATA_INICIO_PLANEJADO);
    const dados   = sheetObra.getRange(rowStart, 1, numRows, numCols).getValues();

    const saidaCrono = precisaCrono ? [] : null;
    const saidaMes   = precisaMes   ? [] : null;

    for (let i = 0; i < numRows; i++) {
      const dataInicio = dados[i][C.DATA_INICIO_PLANEJADO - 1];
      const dtNorm     = normalizarDataSomenteDia_(dataInicio);

      // -- SEMANA CRONOGRAMA --
      if (saidaCrono !== null) {
        if (dtNorm) {
          const emp      = String(dados[i][C.EMP - 1]).trim().toUpperCase();
          const uni      = String(dados[i][C.UNI - 1]).trim();
          const dataLote = normalizarDataSomenteDia_(dataLoteMap.get(`${emp}|${uni}`));
          if (dataLote) {
            const diffDays = Math.floor((dtNorm.getTime() - dataLote.getTime()) / 86400000);
            saidaCrono.push([Math.max(1, Math.ceil(diffDays / 7)) + "ª semana"]);
          } else {
            saidaCrono.push([""]);
          }
        } else {
          saidaCrono.push([""]);
        }
      }

      // -- SEMANA DO MÊS --
      if (saidaMes !== null) {
        saidaMes.push([calcularSemanaMes_(dataInicio)]);
      }
    }

    // --- 2 ESCRITAS (no máximo) ---
    if (saidaCrono !== null) {
      sheetObra.getRange(rowStart, C.SEMANA,     numRows, 1).setValues(saidaCrono);
    }
    if (saidaMes !== null) {
      sheetObra.getRange(rowStart, C.SEMANA_MES, numRows, 1).setValues(saidaMes);
    }
  });
}

/**
 * Compat: mantido para chamadas externas que usem esta assinatura individualmente.
 * Redireciona para a função combinada calculando apenas SEMANA CRONOGRAMA.
 */
function sincronizarSemanasCronogramaObra_(e) {
  sincronizarSemanasObraCombinada_(e, true, false);
}

/**
 * Sincroniza o status operacional da obra (ATIVA/FINALIZADA) para INFORMAÇÕES GERAIS.
 */
function sincronizarStatusObraGeral_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
  const entrega = ss.getSheetByName(CONFIG.SHEETS.ENTREGA);
  if (!info || !entrega) return;

  const C_INFO = resolveSheetColumns_(info, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
  const C_ENT = resolveSheetColumns_(entrega, CONFIG.HEADERS_COLS.ENTREGA, CONFIG.COLUMNS.ENTREGA);
  
  const iniEnt = obterLinhaInicialPorAba(CONFIG.SHEETS.ENTREGA);
  const lastEnt = entrega.getLastRow();
  const mapaFinalizadas = new Set();

  if (lastEnt >= iniEnt) {
    const dadosEnt = entrega.getRange(iniEnt, 1, lastEnt - iniEnt + 1, Math.max(C_ENT.UNI, C_ENT.STATUS_GERAL || 0)).getValues();
    for (let i = 0; i < dadosEnt.length; i++) {
        const status = String(dadosEnt[i][C_ENT.STATUS_GERAL - 1] || "").trim().toUpperCase();
        if (status === "APROVADO" || status === "CONCLUÍDO" || status === "CONCLUIDO") {
            const emp = String(dadosEnt[i][C_ENT.EMP - 1]).trim().toUpperCase();
            const uni = String(dadosEnt[i][C_ENT.UNI - 1]).trim();
            if (emp && uni) mapaFinalizadas.add(`${emp}|${uni}`);
        }
    }
  }

  const iniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
  const lastInfo = info.getLastRow();
  if (lastInfo < iniInfo) return;

  const rangeInfo = info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, Math.max(C_INFO.EMP, C_INFO.UNI));
  const dadosInfo = rangeInfo.getValues();
  const saidaStatus = [];

  for (let i = 0; i < dadosInfo.length; i++) {
      const emp = String(dadosInfo[i][C_INFO.EMP - 1]).trim().toUpperCase();
      const uni = String(dadosInfo[i][C_INFO.UNI - 1]).trim();
      if (!emp || !uni) {
          saidaStatus.push([""]);
          continue;
      }
      const chave = `${emp}|${uni}`;
      saidaStatus.push([mapaFinalizadas.has(chave) ? "FINALIZADA" : "ATIVA"]);
  }

  if (C_INFO.STATUS_OBRA > 0) {
    info.getRange(iniInfo, C_INFO.STATUS_OBRA, saidaStatus.length, 1).setValues(saidaStatus);
  }
}

/**
 * Recalcula as semanas de cronograma para toda a aba FASE-OBRA.
 */
function sincronizarTodaAbaObraSemanasCronograma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!obra) return;
  
  const last = obra.getLastRow();
  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  if (last < ini) return;

  const fakeEvent = {
    range: obra.getRange(ini, 1, last - ini + 1, 1),
    source: ss
  };
  sincronizarSemanasCronogramaObra_(fakeEvent);
  ss.toast("Semana do cronograma recalculada para toda a aba!", "Sucesso");
}

/**
 * Calcula a "Semana do Mês" de uma data, com semanas de Segunda a Domingo.
 *
 * Nomenclatura: "NªS Mmm" (ex: "4ªS Mar").
 * Quando a semana cruza dois meses: "NªS Mmm/ 1ªS Mmm2" (ex: "5ªS Mar/ 1ªS Abr").
 *
 * @param {Date|string} dataInicio - Data de início planejado.
 * @returns {string} A representação da semana do mês.
 */
function calcularSemanaMes_(dataInicio) {
  const dt = normalizarDataSomenteDia_(dataInicio);
  if (!dt) return "";

  const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

  // Encontra a Segunda-feira da semana contendo 'dt'
  // getDay(): Dom=0, Seg=1, ..., Sab=6
  const jsDay = dt.getDay();
  const diasDesdeSegunda = (jsDay === 0) ? 6 : (jsDay - 1);

  const weekMon = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate() - diasDesdeSegunda);
  const weekSun = new Date(weekMon.getFullYear(), weekMon.getMonth(), weekMon.getDate() + 6);

  /**
   * Retorna o número sequencial da semana dentro do mês para uma Segunda-feira.
   * Considera apenas as Segundas-feiras que caem dentro do próprio mês.
   * Ex: Março 2026, primeira Segunda = dia 2 → semana 1.
   */
  function numSemanaNoMes(seg) {
    const ano = seg.getFullYear();
    const mes = seg.getMonth();
    const diaDoMes = seg.getDate();

    // Descobre o dia da semana do dia 1 do mês (para achar a primeira Segunda)
    const jsDayPrimeiro = new Date(ano, mes, 1).getDay();
    // Quantos dias até a próxima Segunda (0 se o dia 1 for Segunda)
    const diasAteSegunda = (jsDayPrimeiro === 1) ? 0 :
                            (jsDayPrimeiro === 0) ? 1 :
                            (8 - jsDayPrimeiro);
    const primeiraSeg = 1 + diasAteSegunda;

    // N-ésima Segunda contando a partir da primeira Segunda do mês
    return Math.floor((diaDoMes - primeiraSeg) / 7) + 1;
  }

  const mesMon = weekMon.getMonth();
  const mesSun = weekSun.getMonth();
  const anoMon = weekMon.getFullYear();
  const anoSun = weekSun.getFullYear();

  if (mesMon === mesSun && anoMon === anoSun) {
    // Semana inteira dentro do mesmo mês
    const num = numSemanaNoMes(weekMon);
    return `${num}ªS ${MESES[mesMon]}`;
  } else {
    // Semana cruza dois meses:
    // - Para o mês da Segunda: é a N-ésima semana daquele mês
    // - Para o mês do Domingo: sempre será a 1ª semana (primeira semana que contém dias daquele mês)
    const numInicio = numSemanaNoMes(weekMon);
    return `${numInicio}ªS ${MESES[mesMon]}/ 1ªS ${MESES[mesSun]}`;
  }
}

/**
 * Calcula e grava a coluna SEMANA DO MÊS nas linhas editadas da FASE-OBRA.
 * Disparado por onEdit quando DATA_INICIO_PLANEJADO é alterada.
 */
/**
 * Compat: redireciona para a função combinada calculando apenas SEMANA DO MÊS.
 */
function sincronizarSemanaMesObra_(e) {
  sincronizarSemanasObraCombinada_(e, false, true);
}

/**
 * Recalcula a coluna SEMANA DO MÊS para toda a aba FASE-OBRA em lote.
 */
function sincronizarTodaAbaObraSemanaMes() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!obra) return;

  const last = obra.getLastRow();
  const ini  = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  if (last < ini) return;

  sincronizarSemanasObraCombinada_(
    { range: obra.getRange(ini, 1, last - ini + 1, 1), source: ss },
    false,   // só SEMANA DO MÊS
    true
  );
  ss.toast("✅ Semana do Mês recalculada para toda a FASE-OBRA!", "Sucesso", 5);
}

/**
 * Wrapper para executar a sincronização de status e limpar SEMANA em FASE-OBRA
 * para todas as unidades marcadas como FINALIZADA na INFORMAÇÕES GERAIS.
 * Use no Editor de Scripts (selecionar função e rodar) ou vincule ao menu.
 */
function executarSincronizarStatusELimparSemana() {
  executarComDocumentLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast("Executando sincronização de status e limpeza de SEMANA...", "Aguarde", 3);
    try {
      // 1) Atualiza STATUS OBRA a partir da FASE-ENTREGA
      sincronizarStatusObraGeral_();

      // 2) Limpa SEMANA para todas as unidades FINALIZADA em INFORMAÇÕES GERAIS
      const info = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
      if (!info) {
        ss.toast("Aba INFORMAÇÕES GERAIS não encontrada.", "Erro", 6);
        return;
      }
      const iniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
      const lastInfo = info.getLastRow();
      if (lastInfo < iniInfo) {
        ss.toast("Nenhuma linha de INFORMAÇÕES GERAIS para processar.", "Info", 3);
        return;
      }

      const fakeEvent = { range: info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, 1) };
      limparSemanaCronogramaPorStatusInformacoesGerais_(fakeEvent);

      ss.toast("Sincronização e limpeza concluídas.", "Sucesso", 4);
    } catch (err) {
      console.error("Erro executarSincronizarStatusELimparSemana: " + err);
      ss.toast("Erro durante a execução. Verifique o log.", "Erro", 8);
    }
  });
}

/**
 * Limpa a coluna SEMANA na aba FASE-OBRA para unidades marcadas como FINALIZADA
 * na aba INFORMAÇÕES GERAIS. Chamado a partir do onEdit de Informações Gerais.
 */
function limparSemanaCronogramaPorStatusInformacoesGerais_(e) {
  if (!e || !e.range) return;
  executarComDocumentLock_(function() {
    const info = e.range.getSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const C_INFO = resolveSheetColumns_(info, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
    if (!C_INFO || C_INFO.STATUS_OBRA <= 0) return;

    const linhaIniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
    const rowStart = Math.max(e.range.getRow(), linhaIniInfo);
    const numRows = e.range.getLastRow() - rowStart + 1;
    if (numRows <= 0) return;

    const maxColInfo = Math.max(C_INFO.EMP, C_INFO.UNI, C_INFO.STATUS_OBRA);
    const dadosInfo = info.getRange(rowStart, 1, numRows, maxColInfo).getDisplayValues();
    const chavesFinalizadas = new Set();
    for (let i = 0; i < dadosInfo.length; i++) {
      const emp = String(dadosInfo[i][C_INFO.EMP - 1] || "").trim().toUpperCase();
      const uni = String(dadosInfo[i][C_INFO.UNI - 1] || "").trim();
      const st  = String(dadosInfo[i][C_INFO.STATUS_OBRA - 1] || "").trim().toUpperCase();
      if (!emp || !uni) continue;
      if (st === "FINALIZADA") chavesFinalizadas.add(`${emp}|${uni}`);
    }
    if (chavesFinalizadas.size === 0) return;

    const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
    if (!obra) return;
    const C_OBRA = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    if (!C_OBRA || C_OBRA.SEMANA <= 0) {
      console.error("Coluna SEMANA não encontrada em FASE-OBRA. Operação cancelada.");
      return;
    }

    const iniObra = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
    const lastObra = obterUltimaLinhaDados_(obra, C_OBRA.EMP);
    if (lastObra < iniObra) return;

    const numObraRows = lastObra - iniObra + 1;
    const maxColObra = Math.max(C_OBRA.EMP, C_OBRA.UNI, C_OBRA.SEMANA);
    const dadosObra = obra.getRange(iniObra, 1, numObraRows, maxColObra).getValues();

    const novoSemanas = [];
    let changed = 0;
    for (let i = 0; i < dadosObra.length; i++) {
      const emp = String(dadosObra[i][C_OBRA.EMP - 1] || "").trim().toUpperCase();
      const uni = String(dadosObra[i][C_OBRA.UNI - 1] || "").trim();
      const chave = `${emp}|${uni}`;
      const atual = dadosObra[i][C_OBRA.SEMANA - 1];
      if (chavesFinalizadas.has(chave)) {
        if (String(atual || "").trim() !== "") {
          novoSemanas.push([""]); 
          changed++;
        } else {
          novoSemanas.push([atual]);
        }
      } else {
        novoSemanas.push([atual]);
      }
    }

    if (changed > 0) {
      obra.getRange(iniObra, C_OBRA.SEMANA, novoSemanas.length, 1).setValues(novoSemanas);
      console.log("Limpou SEMANA em FASE-OBRA para " + changed + " unidade(s) marcadas FINALIZADA.");
    }
  });
}

/**
 * Retorna o texto do indicador de cronograma da coluna E.
 */
function calcularTextoIndicadorCronogramaServico_(statusAprovacao, dataFimPlanejado, hoje) {
  const status = String(statusAprovacao || "").trim().toUpperCase();
  const regexConcluido = /100%|APROVAD|CONCLU|FINALIZ|EXECUTAD|ENTREGUE/;
  if (regexConcluido.test(status)) return "CONCLUÍDO";

  const dataFim = normalizarDataSomenteDia_(dataFimPlanejado);
  if (!dataFim) return "";

  const baseHoje = hoje || normalizarDataSomenteDia_(new Date());
  const msPorDia = 24 * 60 * 60 * 1000;
  const diffDays = Math.floor((dataFim.getTime() - baseHoje.getTime()) / msPorDia);

  if (diffDays > 0) return `Faltam ${diffDays} dia(s)`;
  if (diffDays === 0) return "Vence hoje";
  return `Atrasado ${Math.abs(diffDays)} dia(s)`;
}

/**
 * Atualiza o indicador de cronograma (coluna E) por edição na FASE-OBRA.
 */
function sincronizarIndicadorCronogramaFaseObraPorEdicao_(e) {
  if (!e || !e.range) return;
  const sheetObra = e.range.getSheet();
  if (sheetObra.getName() !== CONFIG.SHEETS.OBRA) return;

  const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const colStatusAprov = obterColunaStatusAprovacaoServicoFaseObra_(sheetObra);
  const colDataFim = obterColunaDataFimPlanejadoFaseObra_(sheetObra);
  if (C_OBRA.CRONOGRAMA <= 0 || colStatusAprov <= 0 || colDataFim <= 0) return;

  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  const rowStart = Math.max(e.range.getRow(), linhaIni);
  const numRows = e.range.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  const maxCol = Math.max(C_OBRA.CRONOGRAMA, colStatusAprov, colDataFim);
  const dados = sheetObra.getRange(rowStart, 1, numRows, maxCol).getValues();
  const hoje = normalizarDataSomenteDia_(new Date());
  const saida = [];

  for (let i = 0; i < dados.length; i++) {
    const statusAprovacao = dados[i][colStatusAprov - 1];
    const dataFimPlanejado = dados[i][colDataFim - 1];
    saida.push([calcularTextoIndicadorCronogramaServico_(statusAprovacao, dataFimPlanejado, hoje)]);
  }

  sheetObra.getRange(rowStart, C_OBRA.CRONOGRAMA, numRows, 1).setValues(saida);
}

/**
 * Recalcula o indicador de cronograma (coluna E) para toda a FASE-OBRA.
 */
function sincronizarTodosIndicadoresCronogramaFaseObra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!obra) return;

  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  const last = obra.getLastRow();
  if (last < ini) return;

  const fakeEvent = {
    range: obra.getRange(ini, 1, last - ini + 1, 1),
    source: ss
  };
  sincronizarIndicadorCronogramaFaseObraPorEdicao_(fakeEvent);
  ss.toast("Indicador de cronograma (coluna E) recalculado para toda a FASE-OBRA!", "Sucesso", 5);
}

/**
 * Ordena a aba FASE-PRELIMINAR de acordo com a ordem exata visual de INFORMAÇÕES GERAIS.
 * A ordem física/numérica da base vira a nova ordem principal na preliminar, 
 * evitando a quebra de DataValidations que ocorre com operações em lote usando setValues na tela toda.
 */
function ordenarPreliminarIgualInformacoesGerais() {
  executarComDocumentLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaInfo = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
    const abaPre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
    if (!abaInfo || !abaPre) return;

    // Toast visível se estiver rodando pelo UI, caso madruagada não exibe (embora .toast ignore).
    try { ss.toast("Lendo ordem atual...", "Ordenação Automática", 3); } catch(e){}

    const C_INFO = resolveSheetColumns_(abaInfo, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
    const C_PRE = resolveSheetColumns_(abaPre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);

    const iniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
    const lastInfo = abaInfo.getLastRow();
    
    // 1. Mapear a ordem atual em INFORMAÇÕES GERAIS para criar o ranking Base
    const mapaOrdem = new Map();
    if (lastInfo >= iniInfo) {
      const maxColInfo = Math.max(C_INFO.EMP, C_INFO.UNI);
      const dadosInfo = abaInfo.getRange(iniInfo, 1, lastInfo - iniInfo + 1, maxColInfo).getDisplayValues();
      for (let i = 0; i < dadosInfo.length; i++) {
        const emp = String(dadosInfo[i][C_INFO.EMP - 1] || "").trim().toUpperCase();
        const uni = String(dadosInfo[i][C_INFO.UNI - 1] || "").trim();
        if (emp && uni) {
          const chave = `${emp}|${uni}`;
          if (!mapaOrdem.has(chave)) {
            mapaOrdem.set(chave, i + 1); // 1º, 2º, 3º...
          }
        }
      }
    }

    if (mapaOrdem.size === 0) return;

    try { ss.toast("Aplicando ordenação...", "Ordenação Automática", 3); } catch(e){}

    // 2. Aplicar etiqueta na Preliminar
    const iniPre = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
    const lastPre = abaPre.getLastRow();
    if (lastPre < iniPre) return;

    const maxColPre = Math.max(C_PRE.EMP, C_PRE.UNI);
    const dadosPre = abaPre.getRange(iniPre, 1, lastPre - iniPre + 1, maxColPre).getDisplayValues();
    const etiquetas = [];

    for (let i = 0; i < dadosPre.length; i++) {
      const emp = String(dadosPre[i][C_PRE.EMP - 1] || "").trim().toUpperCase();
      const uni = String(dadosPre[i][C_PRE.UNI - 1] || "").trim();
      
      const marcador = CONFIG.STATUS.MARCADOR_BASE.test(emp);
      if (marcador) {
         etiquetas.push([999999]); // Marcador vai pro fundo absoluto
         continue;
      }

      if (emp && uni) {
        const chave = `${emp}|${uni}`;
        const ranking = mapaOrdem.has(chave) ? mapaOrdem.get(chave) : 99999;
        etiquetas.push([ranking]);
      } else {
        etiquetas.push([99999]); // Itens não listados, brancos ou lixos caem pro fundo
      }
    }

    // 3. Coluna temporária. Usar o LastColumn "verdadeiro" ou instanciar um distante O(1).
    const maxColsPreRange = abaPre.getMaxColumns();
    // Insere a coluna no limite horizontal para não sobrepor nada
    abaPre.insertColumnAfter(maxColsPreRange);
    const colTemp = parseInt(maxColsPreRange) + 1;
    
    abaPre.getRange(iniPre, colTemp, etiquetas.length, 1).setValues(etiquetas);

    // 4. Sort nativo do Sheets. Arrastará as validações impecavelmente.
    const rangeTotal = abaPre.getRange(iniPre, 1, etiquetas.length, colTemp);
    rangeTotal.sort({ column: colTemp, ascending: true });

    // 5. Limpeza de rastro
    abaPre.deleteColumn(colTemp);

    try { ss.toast("✅ FASE-PRELIMINAR perfeitamente alinhada!", "Sucesso", 5); } catch(e){}
  }, 30000);
}

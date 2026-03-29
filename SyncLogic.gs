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

    // Obtém dados de todas as linhas editadas da obra
    const dadosObra = sheetObra.getRange(rowStart, 1, numRows, C_OBRA.CHAVE).getValues();
    const mapaPedidos = obterMapaPedidosPorChave_(abaPedidos);
    let proximaLinhaLivre = -1;

    const linhasParaDeletar = [];

    for (let i = 0; i < numRows; i++) {
        const rowObra = rowStart + i;
        const vals = dadosObra[i];
        const atrelado = String(vals[C_OBRA.ATRELADO - 1]).trim().toUpperCase();
        const chaveID_Row = String(vals[C_OBRA.CHAVE - 1] || "").trim();

        // Se o atrelado não for HOUSI, remove da aba Pedidos se existir
        if (atrelado !== "HOUSI") {
          const linhaExistente = mapaPedidos.get(chaveID_Row);
          if (linhaExistente > 0) {
            linhasParaDeletar.push(linhaExistente);
            mapaPedidos.delete(chaveID_Row);
          }
          continue;
        }

        // Dados da linha
        const emp = String(vals[C_OBRA.EMP - 1]).trim();
        const uni = String(vals[C_OBRA.UNI - 1]).trim();
        const cat = String(vals[C_OBRA.CAT - 1]).trim();
        const sub = String(vals[C_OBRA.SUB - 1]).trim();
        if (!emp || !uni || !cat || !sub) continue;

        let chaveID = vals[C_OBRA.CHAVE - 1];
        if (!chaveID || String(chaveID).startsWith("FO_ROW_")) {
          chaveID = gerarUUID_();
          sheetObra.getRange(rowObra, C_OBRA.CHAVE).setValue(chaveID);
        }

        let linhaPed = mapaPedidos.get(String(chaveID).trim()) || -1;
        let rowParaAtualizar;
        
        if (linhaPed <= 0) {
          if (proximaLinhaLivre < 0) proximaLinhaLivre = obterPrimeiraLinhaLivrePedidos_(abaPedidos);
          linhaPed = proximaLinhaLivre++;
          garantirLinhasAte_(abaPedidos, linhaPed);
          rowParaAtualizar = new Array(C_PED.CHAVE).fill("");
        } else {
          rowParaAtualizar = abaPedidos.getRange(linhaPed, 1, 1, C_PED.CHAVE).getValues()[0];
        }
        
        // Atualiza campos
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

        // Flush por linha para garantir integridade individual
        abaPedidos.getRange(linhaPed, 1, 1, C_PED.CHAVE).setValues([rowParaAtualizar]);
    }

    // Deleta ordão de baixo para cima para não quebrar índices
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

  // Lê todas as linhas e o mapa de pedidos uma vez
  const maxColObra = Math.max(C_OBRA.CHAVE, C_OBRA.DATA_SOLICITADO_OPR);
  const dadosObra = sheetObra.getRange(rowStart, 1, numRows, maxColObra).getValues();
  const mapaPedidos = obterMapaPedidosPorChave_(abaPedidos);

  for (let i = 0; i < numRows; i++) {
    const valsObra = dadosObra[i];
    const chaveID = valsObra[C_OBRA.CHAVE - 1];
    const dataNova = valsObra[C_OBRA.DATA_SOLICITADO_OPR - 1];

    if (!chaveID) continue;

    const linhaPed = mapaPedidos.get(String(chaveID).trim());
    if (linhaPed > 0 && C_PED.DATA_SOLICITADO_OPR > 0) {
      abaPedidos.getRange(linhaPed, C_PED.DATA_SOLICITADO_OPR).setValue(dataNova || null);
    }
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
  const pre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
  if (!info || !pre) return;

  const linhaInicialInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
  const lastInfo = info.getLastRow();
  if (lastInfo < linhaInicialInfo) return;

  const C_INFO = resolveSheetColumns_(info, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
  const maxColInfo = Math.max(
    C_INFO.EMP,
    C_INFO.UNI,
    C_INFO.DATA_LOTE,
    C_INFO.DATA_PRAZO,
    C_INFO.FASE_MACRO,
    C_INFO.PRIORIDADE,
    C_INFO.RESP_OPR,
    C_INFO.RESP_ADM
  );
  const registrosInfo = info.getRange(linhaInicialInfo, 1, lastInfo - linhaInicialInfo + 1, maxColInfo).getValues();

  const C_PRE = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);

  // Mapeamento de colunas na Preliminar (usando map dinâmico)
  const mapeamento = {
    colDataLote: C_PRE.DATA_LOTE,
    colDataPrazo: C_PRE.DATA_PRAZO,
    colFaseMacro: obterColunaPorCabecalho_(pre, CONFIG.HEADERS.FASE_MACRO, obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR) - 1),
    colPrioridade: C_PRE.PRIORIDADE,
    colRespOpr: C_PRE.RESP_OPR,
    colRespAdm: C_PRE.RESP_ADM
  };

  // Mapeamento de Checklist (Dinâmico por Cabeçalho)
  const mapeamentoChecklist = {};
  const lastColPre = pre.getLastColumn();
  const linhaHeaderPre = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR) - 1;
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

  const linhaInicialPre = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
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

  let criadas = 0;
  const colunasNovas = new Set(); // Para aplicar defaults apenas em novas linhas

  for (let i = 0; i < registrosInfo.length; i++) {
    const row = registrosInfo[i];
    const emp = String(row[C_INFO.EMP - 1] || "").trim();
    const uni = String(row[C_INFO.UNI - 1] || "").trim();
    if (!emp || !uni) continue;

    const chave = `${emp.toUpperCase()}|${uni}`;
    if (chavesAlvo && chavesAlvo.size > 0 && !chavesAlvo.has(chave)) continue;

    let rowPre = mapaPrePorChave.get(chave);
    if (!rowPre) {
      pre.insertRowsBefore(linhaInicialPre, 1);
      
      // Atualiza mapa de linhas existentes (desloca para baixo)
      for (const [k, v] of mapaPrePorChave.entries()) {
        mapaPrePorChave.set(k, v + 1);
      }

      // Copia formato da linha de baixo
      const linhaMolde = pre.getRange(linhaInicialPre + 1, 1, 1, pre.getMaxColumns());
      const alvo = pre.getRange(linhaInicialPre, 1, 1, pre.getMaxColumns());
      linhaMolde.copyTo(alvo, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      linhaMolde.copyTo(alvo, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

      rowPre = linhaInicialPre;
      mapaPrePorChave.set(chave, rowPre);
      colunasNovas.add(rowPre);
      criadas++;
    }

    // Grava dados principais
    pre.getRange(rowPre, 1, 1, 2).setValues([[emp, uni]]);
    if (mapeamento.colDataLote > 0) pre.getRange(rowPre, mapeamento.colDataLote).setValue(row[C_INFO.DATA_LOTE - 1]);
    if (mapeamento.colDataPrazo > 0) pre.getRange(rowPre, mapeamento.colDataPrazo).setValue(row[C_INFO.DATA_PRAZO - 1]);
    if (mapeamento.colFaseMacro > 0) pre.getRange(rowPre, mapeamento.colFaseMacro).setValue(row[C_INFO.FASE_MACRO - 1]);
    
    // Removidos RESUMO_PENDENCIAS e RESUMO_OCORRENCIAS (usuário vai apagar colunas)
    // Se ainda precisar gravar em algum lugar, precisaria de fallback; aqui simplesmente ignoramos
    
    if (mapeamento.colPrioridade > 0) pre.getRange(rowPre, mapeamento.colPrioridade).setValue(row[C_INFO.PRIORIDADE - 1]);
    if (mapeamento.colRespOpr > 0) pre.getRange(rowPre, mapeamento.colRespOpr).setValue(row[C_INFO.RESP_OPR - 1]);
    if (mapeamento.colRespAdm > 0) pre.getRange(rowPre, mapeamento.colRespAdm).setValue(row[C_INFO.RESP_ADM - 1]);

    // Aplica Defaults (apenas se for nova linha)
    if (colunasNovas.has(rowPre)) {
      for (const col in mapeamentoChecklist) {
        pre.getRange(rowPre, Number(col)).setValue(mapeamentoChecklist[col]);
      }
    }
    
    // Limpa validação de unidade
    processarIntervaloAparaB_(pre, pre.getRange(rowPre, 1, 1, 1));
  }

  if (exibirAlerta) {
    SpreadsheetApp.getUi().alert("Sincronização concluída. Novas unidades criadas: " + criadas);
  }
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

  const maxColPre = Math.max(C_PRE.EMP || 0, C_PRE.UNI || 0, C_PRE.RESUMO_OCORRENCIAS || 0, 1);
  const rangePre = abaPre.getRange(iniPre, 1, lastPre - iniPre + 1, maxColPre);
  const dadosPre = rangePre.getValues();
  const saidaOco = [];

  for (let i = 0; i < dadosPre.length; i++) {
    const emp = String(dadosPre[i][C_PRE.EMP - 1]).trim().toUpperCase();
    const uni = String(dadosPre[i][C_PRE.UNI - 1]).trim();
    const chave = `${emp}|${uni}`;
    
    if (chavesAlvo && !chavesAlvo.has(chave)) {
      saidaOco.push([abaPre.getRange(iniPre + i, C_PRE.RESUMO_OCORRENCIAS).getValue()]);
      continue;
    }

    const qtd = mapaContagem.get(chave) || 0;
    saidaOco.push([qtd > 0 ? qtd + " OCORRÊNCIA(S) ABERTAS" : "LIMPO"]);
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

    // Obtém Status (H), Fornecedor (I) e ChaveID (AJ) do Pedidos
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
        // Grava apenas colunas J e K (Status e Fornecedor na Fase Obra)
        const rangeJK = abaObra.getRange(iniObra, C_OBRA.STATUS, dadosObra.length, 2);
        const novosValoresJK = dadosObra.map(r => [r[C_OBRA.STATUS - 1], r[C_OBRA.FORNECEDOR - 1]]);
        rangeJK.setValues(novosValoresJK);
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
 * Calcula e sincroniza a semana do cronograma na Fase Obra.
 */
function sincronizarSemanasCronogramaObra_(e) {
  executarComDocumentLock_(function() {
    if (!e || !e.range) return;
    const sheetObra = e.range.getSheet();
    const rowStart = e.range.getRow();
    const numRows = e.range.getNumRows();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    if (C_OBRA.DATA_INICIO_PLANEJADO <= 0 || C_OBRA.SEMANA <= 0) {
      const msg = "Não foi possível calcular SEMANA CRONOGRAMA: cabeçalhos de Data Início Planejado e/ou Semana não encontrados em FASE-OBRA.";
      console.error(msg);
      ss.toast(msg, "⚠️ Configuração ausente", 6);
      return;
    }

    const dataLoteMap = obterMapaDataLote_(ss);
    const rangeEdicao = sheetObra.getRange(rowStart, 1, numRows, Math.max(C_OBRA.UNI, C_OBRA.DATA_INICIO_PLANEJADO));
    const dados = rangeEdicao.getValues();
    const saidaSemana = [];

    for (let i = 0; i < numRows; i++) {
      const emp = String(dados[i][C_OBRA.EMP - 1]).trim().toUpperCase();
      const uni = String(dados[i][C_OBRA.UNI - 1]).trim();
      const dataInicio = normalizarDataSomenteDia_(dados[i][C_OBRA.DATA_INICIO_PLANEJADO - 1]);
      
      const chave = `${emp}|${uni}`;
      const dataLote = normalizarDataSomenteDia_(dataLoteMap.get(chave));

      if (dataInicio && dataLote) {
        const diffMs = dataInicio.getTime() - dataLote.getTime();
        const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
        
        // Regra: até 7 dias = 1ª semana, até 14 dias = 2ª semana, etc.
        const numSemana = Math.max(1, Math.ceil(diffDays / 7)); 
        // Se diffDays <= 0, ceil(0) = 0 -> max(1, 0) = 1.
        // Se diffDays = 7, ceil(7/7) = 1.
        // Se diffDays = 8, ceil(8/7) = 2.
        
        saidaSemana.push([numSemana + "ª semana"]);
      } else {
        saidaSemana.push([""]);
      }
    }

    sheetObra.getRange(rowStart, C_OBRA.SEMANA, numRows, 1).setValues(saidaSemana);
  });
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

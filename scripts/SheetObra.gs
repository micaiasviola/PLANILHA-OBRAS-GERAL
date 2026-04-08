/*************************
 * MÓDULO: FASE-OBRA
 *************************/

/**
 * Handler disparado pelo router onEdit quando a aba FASE-OBRA é editada.
 */
function handleObraEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const col = range.getColumn();
  const C = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);

  // 1) Categoria -> Subcategoria (G -> H)
  if (intervaloInterceptaColuna(range, C.CAT)) {
    processarSubcategoriasObra_(sheet, range);
  }

  // 2) Sincronização de Pedidos (A, B, C, D, F, G, H, I) -> PEDIDOS-GERAL
  const colsSincPedidos = [C.EMP, C.UNI, C.OPR, C.ADM, C.TIPO, C.CAT, C.SUB, C.ATRELADO];
  if (colsSincPedidos.some(c => intervaloInterceptaColuna(range, c))) {
    sincronizarPedidosHousiPorEdicao_(e);
    
    // Pequeno delay para garantir persistência e limpa cache
    Utilities.sleep(100);
    atualizarLinhaObra_(range.getRow());
  }

  // 3) Data Solicitado (L) -> Sincroniza com Pedidos
  if (intervaloInterceptaColuna(range, C.DATA_SOLICITADO_OPR)) {
    sincronizarDataPrevista_(e);
  }

  // 4) Verba Housi (W) -> Verba Teto (X)
  if (intervaloInterceptaColuna(range, C.VERBA_HOUSI)) {
    sincronizarVerbaTetoPorEdicao_(e);
  }

  // 5) Indicador de Cronograma (E)
  const colStatusAprov = obterColunaStatusAprovacaoServicoFaseObra_(sheet);
  const colDataFim = obterColunaDataFimPlanejadoFaseObra_(sheet);
  if (intervaloInterceptaColuna(range, colStatusAprov) || intervaloInterceptaColuna(range, colDataFim)) {
    sincronizarIndicadorCronogramaFaseObraPorEdicao_(e);
  }

  // 6) Enviar para FASE-ENTREGA (Checkbox)
  const colChaveEntrega = obterColunaChaveEntregaFaseObra_(sheet);
  if (colChaveEntrega > 0 && intervaloInterceptaColuna(range, colChaveEntrega)) {
    sincronizarFaseObraParaFaseEntregaPorChave_(e, colChaveEntrega);
  }

  // 7) Semana do Cronograma + Semana do Mês — COMBINADAS em 1 lock, 1 leitura, 2 escritas
  if (C.DATA_INICIO_PLANEJADO > 0 && intervaloInterceptaColuna(range, C.DATA_INICIO_PLANEJADO)) {
    sincronizarSemanasObraCombinada_(e, true, C.SEMANA_MES > 0);
  }
}

/**
 * Atualiza dropdowns de subcategoria baseados na categoria.
 */
function processarSubcategoriasObra_(abaObra, intervalo, options) {
  options = options || {};
  const revalidate = !!options.revalidate;
  const defaultToFirst = !!options.defaultToFirst;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCad = ss.getSheetByName(CONFIG.SHEETS.CADASTROS);
  if (!abaCad) return;

  const C = resolveSheetColumns_(abaObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  const rowStart = Math.max(intervalo.getRow(), linhaIni);
  const numRows = intervalo.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  const CAD_FIRST_ROW = 6;
  const CAD_FIRST_COL = 2;
  const CAD_NUM_COLS = 2;
  const CAD_NUM_ROWS = Math.max(0, abaCad.getLastRow() - (CAD_FIRST_ROW - 1));
  const dadosCad = CAD_NUM_ROWS > 0 ? abaCad.getRange(CAD_FIRST_ROW, CAD_FIRST_COL, CAD_NUM_ROWS, CAD_NUM_COLS).getValues() : [];

  for (let i = 0; i < numRows; i++) {
    const row = rowStart + i;
    const cat = String(abaObra.getRange(row, C.CAT).getValue()).trim();
    const celulaSub = abaObra.getRange(row, C.SUB);

    if (!cat) {
      celulaSub.clearDataValidations().clearContent();
      continue;
    }

    const subcategorias = dadosCad
      .filter(r => String(r[0]).trim() === cat && String(r[1]).trim() !== "")
      .map(r => String(r[1]).trim());

    if (subcategorias.length === 0) {
      celulaSub.clearDataValidations().clearContent();
      continue;
    }

    const regra = SpreadsheetApp.newDataValidation()
      .requireValueInList(subcategorias, true)
      .setAllowInvalid(true)
      .build();

    celulaSub.setDataValidation(regra);

    // Opcional: revalida o valor atual da célula e limpa (ou substitui) se inválido
    if (revalidate) {
      try {
        const cur = String(celulaSub.getValue() || "").trim();
        if (!cur || subcategorias.indexOf(cur) === -1) {
          if (defaultToFirst) {
            celulaSub.setValue(subcategorias[0]);
          } else {
            celulaSub.clearContent();
          }
        }
      } catch (e) {
        console.error('processarSubcategoriasObra_ revalidate row ' + row + ': ' + e);
      }
    }
  }
}

/**
 * Sincroniza dados de verba teto (90% da verba Housi).
 */
function sincronizarVerbaTetoPorEdicao_(e) {
  const row = e.range.getRow();
  const sheet = e.range.getSheet();
  const C = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const bruto = e.range.getValue();
  const numero = converterParaNumero_(bruto);
  const cellTeto = e.range.getSheet().getRange(row, C.VERBA_TETO);
  
  if (numero === null) cellTeto.clearContent();
  else cellTeto.setValue(numero * 0.9);
}

// ... Outras funções de busca por cabeçalho específicas da Obra ...
function obterColunaStatusAprovacaoServicoFaseObra_(aba) {
  return obterColunaPorCabecalho_(aba, CONFIG.HEADERS.OBRA_STATUS_APROVACAO);
}

function obterColunaDataFimPlanejadoFaseObra_(aba) {
  return obterColunaPorCabecalho_(aba, CONFIG.HEADERS.OBRA_DATA_FIM_PLANEJADO);
}

/**
 * Geração de templates de serviços na FASE-OBRA.
 */
function gerarTemplatesPendentesFaseObra() {
  executarComDocumentLock_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pre = ss.getSheetByName(CONFIG.SHEETS.PRELIMINAR);
    const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
    if (!pre || !obra) return;

    ss.toast("Pesquisando novas unidades para iniciar Obras...", "Aguarde", 5);

    const C_PRE = resolveSheetColumns_(pre, CONFIG.HEADERS_COLS.PRELIMINAR, CONFIG.COLUMNS.PRELIMINAR);
    const fpColOprParaObra = C_PRE.RESP_OPR;
    const fpColAdmParaObra = C_PRE.RESP_ADM;

    // 1) Carrega unidades existentes na OBRA de uma vez (O(1) chamadas à API)
    const C_OBRA = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const obraIni = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
    const obraLast = obra.getLastRow();
    const unidadesObraExistentes = new Set();
    
    if (obraLast >= obraIni) {
        const valsObra = obra.getRange(obraIni, 1, obraLast - obraIni + 1, Math.max(C_OBRA.EMP, C_OBRA.UNI)).getDisplayValues();
        for (let i = 0; i < valsObra.length; i++) {
            const e = String(valsObra[i][C_OBRA.EMP - 1] || "").trim().toUpperCase();
            const u = String(valsObra[i][C_OBRA.UNI - 1] || "").trim();
            if (e && u) unidadesObraExistentes.add(`${e}|${u}`);
        }
    }

    const preIni = obterLinhaInicialPorAba(CONFIG.SHEETS.PRELIMINAR);
    const preLast = pre.getLastRow();
    if (preLast < preIni) return;

    const templateEstatico = [
      ["PROTEÇÕES", "PROTEÇÃO PORTAS E BANCADA", "ECQUA"],
      ["MÁRMORES", "RECORTE COOKTOP", "ECQUA"],
      ["ELÉTRICA", "READEQUAÇÕES ELÉTRICAS", "ECQUA"],
      ["ELÉTRICA", "ILUMINAÇÃO", "ECQUA"],
      ["REVESTIMENTO PISO", "MEDIÇÃO DO VINÍLICO", "HOUSI"],
      ["REVESTIMENTO PISO", "INSTALAÇÃO DO VINÍLICO", "HOUSI"],      
      ["PROTEÇÕES", "PROTEÇÃO PISO", "ECQUA"],
      ["AC", "INSTALAÇÃO AC", "ECQUA"],
      ["AC", "1 SPLIT", "HOUSI"],
      ["VIDROS", "BOX DE ABRIR OU CORRER", "HOUSI"],
      ["VIDROS", "ESPELHO WC", "HOUSI"],
      ["MARCENARIA", "APROVAÇÃO DO CADERNO", "HOUSI"],
      ["LIMPEZA", "LIMPEZA GROSSA", "ECQUA"],
      ["PINTURA", "PINTURA PADRÃO CINZA CABECEIRA E BRANCO GERAL", "ECQUA"],
      ["REVESTIMENTO PISO", "RODAPÉ PVC", "HOUSI"],
      ["TÊXTIL", "CORTINA", "HOUSI"],
      ["ELETROS", "MICROONDAS INOX", "HOUSI"],
      ["ELETROS", "SMART TV 32\"", "HOUSI"],
      ["ELETROS", "COOKTOP EMBUTIR 2 BOCAS", "HOUSI"],
      ["ELETROS", "GELADEIRA INOX", "HOUSI"],
      ["MOBÍLIA", "CAMA CASAL BOX PADRÃO", "HOUSI"],
      ["MOBÍLIA", "CADEIRA HOME-OFFICE", "HOUSI"],
      ["MOBÍLIA", "CADEIRA MESA JANTAR", "HOUSI"],
      ["MOBÍLIA", "MESA JANTAR", "HOUSI"],
      ["ENXOVAL", "ENXOVAL CAMA", "HOUSI"],
      ["ENXOVAL", "ENXOVAL COZINHA", "HOUSI"],
      ["DECORAÇÃO", "DECORAÇÃO PADRÃO", "HOUSI"],
      ["INSTALAÇÕES GERAIS", "DESEMBALAGEM E MONTAGEM GERAL ITENS E ENXOVAL", "ECQUA"],
      ["INSTALAÇÕES GERAIS", "INSTALAÇÃO DE TV", "ECQUA"],
      ["INSTALAÇÕES GERAIS", "INSTALAÇÃO DE KIT WC", "ECQUA"],
      ["INSTALAÇÕES GERAIS", "INSTALAÇÃO DE ASSENTO WC", "ECQUA"],
      ["LIMPEZA", "LIMPEZA FINA", "ECQUA"]
    ];

    const lastColPre = pre.getLastColumn();
    if (lastColPre < 1) return;

    const dadosPre = pre.getRange(preIni, 1, preLast - preIni + 1, lastColPre).getValues();

    let unidadesAbertas = 0;
    const arrayLoteCompleto = [];
    const numMaxColsLinha = C_OBRA.CHAVE;

    // 2) Constrói as linhas virtualmente na memória
    for (let i = 0; i < dadosPre.length; i++) {
        const emp = String(dadosPre[i][C_PRE.EMP - 1] || "").trim();
        const uni = String(dadosPre[i][C_PRE.UNI - 1] || "").trim();
        const enviarParaObra = dadosPre[i][C_PRE.ENVIAR_OBRA - 1];

        // Agora o gatilho é manual via coluna "FASE-OBRA"
        if (!valorEhSim_(enviarParaObra)) continue;
        if (!emp || !uni) continue;

        const chaveUnidade = `${emp.toUpperCase()}|${uni}`;
        if (unidadesObraExistentes.has(chaveUnidade)) continue;

        unidadesObraExistentes.add(chaveUnidade);

        const opr = String(dadosPre[i][fpColOprParaObra - 1] || "").trim();
        const adm = String(dadosPre[i][fpColAdmParaObra - 1] || "").trim();
        
        templateEstatico.forEach(t => {
            const linhaNova = new Array(numMaxColsLinha).fill("");
            linhaNova[C_OBRA.EMP - 1] = emp;
            linhaNova[C_OBRA.UNI - 1] = uni;
            linhaNova[C_OBRA.OPR - 1] = opr;
            linhaNova[C_OBRA.ADM - 1] = adm;
            linhaNova[C_OBRA.TIPO - 1] = "PADRÃO";
            linhaNova[C_OBRA.CAT - 1] = t[0];
            linhaNova[C_OBRA.SUB - 1] = t[1];
            linhaNova[C_OBRA.ATRELADO - 1] = t[2];
            linhaNova[C_OBRA.CHAVE - 1] = gerarUUID_();
            
            arrayLoteCompleto.push(linhaNova);
        });

        unidadesAbertas++;
    }

    if (arrayLoteCompleto.length > 0) {
        ss.toast(`Gerando e inserindo templates para ${unidadesAbertas} novas unidades. Aguarde...`, "Aguarde", 10);
        
        // 3) Insere em Lote
        const iniLivre = Math.max(obterUltimaLinhaDados_(obra, C_OBRA.EMP) + 1, obraIni);
        
        garantirColunasAte_(obra, C_OBRA.CHAVE);
        garantirLinhasAte_(obra, iniLivre + arrayLoteCompleto.length);

        const rangeInsert = obra.getRange(iniLivre, 1, arrayLoteCompleto.length, numMaxColsLinha);
        rangeInsert.clearDataValidations();
        
        rangeInsert.setValues(arrayLoteCompleto);

        // 4) Reaplica Menu Suspenso e Validações
        processarIntervaloAparaB_(obra, rangeInsert);
        processarSubcategoriasObra_(obra, rangeInsert);

        ss.toast(`✅ Templates de ${unidadesAbertas} unidades gerados e importados!`, "Sucesso", 8);
    } else {
        ss.toast("Nenhuma nova unidade qualificada encontrada na FASE-PRELIMINAR.", "Concluído", 3);
    }

  }, 45000);
}

/**
 * Atualiza uma linha da Obra com dados vindos dos Pedidos (Status/Fornecedor).
 */
function atualizarLinhaObra_(linha) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
    if (!obra || linha < obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA)) return;

    const C = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const vals = obra.getRange(linha, 1, 1, Math.max(C.CHAVE, C.ATRELADO)).getValues()[0];
    
    const emp = String(vals[C.EMP - 1]).trim().toUpperCase();
    const uni = String(vals[C.UNI - 1]).trim();
    const cat = String(vals[C.CAT - 1]).trim().toUpperCase();
    const sub = String(vals[C.SUB - 1]).trim().toUpperCase();
    const atr = String(vals[C.ATRELADO - 1]).trim().toUpperCase();

    if (!emp || atr !== "HOUSI") return;

    // Tenta encontrar o pedido correspondente
    const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
    if (!abaPedidos) return;
    
    const chave = vals[C.CHAVE - 1] || gerarUUID_();
    const linhaPed = localizarLinhaPedidoPorChave_(abaPedidos, chave);

    if (linhaPed > 0) {
        const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
        const dadosPed = abaPedidos.getRange(linhaPed, C_PED.STATUS, 1, C_PED.FORNECEDOR - C_PED.STATUS + 1).getValues()[0];
        obra.getRange(linha, C.STATUS).setValue(dadosPed[0]);
        obra.getRange(linha, C.FORNECEDOR).setValue(dadosPed[1]);
    }
}

function statusFase01Concluida_(status) {
    return /FASE 01 CONCLUIDA/i.test(String(status));
}

function valorEhSim_(valor) {
    return CONFIG.STATUS.SIM_REGEX.test(String(valor));
}

/**
 * Cria e configura as colunas de Cronograma (Data Início e Semana) na aba FASE-OBRA.
 */
function configurarColunasCronogramaFaseObra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!obra) return;

  const linhaHeader = 1;
  const headersDataInicio = CONFIG.HEADERS_COLS.OBRA.DATA_INICIO_PLANEJADO;
  const headersSemana = CONFIG.HEADERS_COLS.OBRA.SEMANA;
  let colDataInicio = obterColunaPorCabecalhoEmLinhas_(obra, headersDataInicio, [1, 2, 3]);
  let colSemana = obterColunaPorCabecalhoEmLinhas_(obra, headersSemana, [1, 2, 3]);
  const criadas = [];

  // Cria apenas a coluna SEMANA CRONOGRAMA.
  // Se DATA INÍCIO existir, cria ao lado dela; senão, cria no final.
  if (colSemana <= 0) {
    const colBase = colDataInicio > 0 ? colDataInicio : obra.getLastColumn();
    obra.insertColumnsAfter(colBase, 1);
    const colNew = colBase + 1;
    obra.getRange(linhaHeader, colNew).setValue("SEMANA CRONOGRAMA");
    obra.getRange(linhaHeader, colBase).copyTo(obra.getRange(linhaHeader, colNew), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    criadas.push("SEMANA CRONOGRAMA");
    colSemana = colNew;
  }

  // Garante que próxima execução recarregue os índices de coluna atualizados
  limparCacheResolucaoColunas_();

  if (criadas.length === 0) {
    SpreadsheetApp.getUi().alert("ℹ️ As colunas de Cronograma já estavam configuradas em FASE-OBRA. Nenhuma nova coluna foi criada.");
    return;
  }

  SpreadsheetApp.getUi().alert("✅ Colunas configuradas em FASE-OBRA: " + criadas.join(", "));
}
/**
 * Cria e configura a coluna SEMANA DO MÊS na aba FASE-OBRA,
 * inserindo-a imediatamente antes da coluna SEMANA CRONOGRAMA.
 * Chamada pelo menu Admin ECQUA.
 */
function configurarColunaSemanaMesFaseObra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!obra) return;

  limparCacheResolucaoColunas_(); // garante leitura fresca
  const C = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const linhasHeader = [1, 2, obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA) - 1].filter(v => v > 0);

  // Verifica se já existe
  if (C.SEMANA_MES > 0) {
    SpreadsheetApp.getUi().alert("A coluna SEMANA DO MÊS já existe na FASE-OBRA (coluna " + C.SEMANA_MES + ").");
    return;
  }

  // Posiciona antes de SEMANA CRONOGRAMA; se não existir, usa o final da planilha
  let colBase = C.SEMANA > 0 ? C.SEMANA - 1 : obra.getLastColumn();

  obra.insertColumnsAfter(colBase, 1);
  const colNova = colBase + 1;

  // Grava o cabeçalho em todas as possíveis linhas de header
  for (const lh of linhasHeader) {
    const cellHeader = obra.getRange(lh, colNova);
    // Copia formato do vizinho esquerdo
    obra.getRange(lh, colBase).copyTo(cellHeader, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    if (lh === linhasHeader[linhasHeader.length - 1]) {
      cellHeader.setValue("SEMANA DO MÊS");
    }
  }

  // Limpa cache para que resolveSheetColumns_ encontre a nova coluna
  limparCacheResolucaoColunas_();

  // Popula toda a aba com o cálculo inicial
  ss.toast("Calculando SEMANA DO MÊS para toda a FASE-OBRA...", "⏳ Aguarde", 8);
  sincronizarTodaAbaObraSemanaMes();

  SpreadsheetApp.getUi().alert(
    "✅ Coluna SEMANA DO MÊS criada na posição " + colNova + "\n" +
    "Ela será atualizada automaticamente ao editar DATA INÍCIO PLANEJADO EXECUÇÃO."
  );
}

/**
 * Sincroniza envio unitário para FASE-ENTREGA quando o checkbox é marcado na FASE-OBRA.
 * Chamada pelo handler onEdit de FASE-OBRA.
 */
function sincronizarFaseObraParaFaseEntregaPorChave_(e, colChaveEntrega) {
  if (!e || !e.range) return;
  const sheetObra = e.range.getSheet();
  const rowStart = e.range.getRow();
  const numRows = e.range.getNumRows();
  const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);

  const maxCol = Math.max(C_OBRA.EMP, C_OBRA.UNI, colChaveEntrega);
  const dados = sheetObra.getRange(rowStart, 1, numRows, maxCol).getDisplayValues();

  const candidatos = [];
  for (let i = 0; i < numRows; i++) {
    const valor = dados[i][colChaveEntrega - 1];
    if (!CONFIG.STATUS.SIM_REGEX.test(String(valor).trim())) continue;

    const emp = String(dados[i][C_OBRA.EMP - 1]).trim();
    const uni = String(dados[i][C_OBRA.UNI - 1]).trim();
    if (emp && uni) candidatos.push([emp, uni]);
  }

  if (candidatos.length > 0) {
    inserirUnidadesNaFaseEntregaSemDuplicar_(candidatos);
  }
}

/**
 * Abre o Dashboard Web como um modal ultra premium dentro da planilha.
 */
function abrirDashboardInterface() {
  const payload = coletarDadosDoDashboard();
  
  const template = HtmlService.createTemplateFromFile('DashboardUI');
  template.serverPayload = JSON.stringify(payload);
  
  const htmlOutput = template.evaluate()
      .setTitle('ECQUA Analytics')
      .setWidth(1200)
      .setHeight(800);
      
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ECQUA Analytics');
}

/**
 * Endpoint "API" interno chamado pelo Frontend HTML para buscar os KPIs na nuvem (O(1)).
 */
function coletarDadosDoDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. VISÃO GLOBAL (INFO GERAIS - Base central)
    const abaInfo = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
    let totalObras = 0;
    let obrasAtivas = 0;
    let obrasFinalizadas = 0;
    let qtdObrasAtrasadas = 0;
    
    if (abaInfo) {
      const C_INFO = resolveSheetColumns_(abaInfo, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
      const iniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
      const lastInfo = abaInfo.getLastRow();
      
      if (lastInfo >= iniInfo) {
        const C_EMP = C_INFO.EMP || 1;
        const C_UNI = C_INFO.UNI || 2;
        const C_STATUS = C_INFO.STATUS_OBRA || 11;
        const C_PEND = C_INFO.RESUMO_PENDENCIAS || 6;

        const maxColInfo = Math.max(C_EMP, C_UNI, C_STATUS, C_PEND);
        const dadosInfo = abaInfo.getRange(iniInfo, 1, lastInfo - iniInfo + 1, maxColInfo).getDisplayValues();
        
        for (let i = 0; i < dadosInfo.length; i++) {
          const emp = String(dadosInfo[i][C_EMP - 1] || "").trim();
          const uni = String(dadosInfo[i][C_UNI - 1] || "").trim();
          
          if (!emp || !uni) continue; 
          if (CONFIG.STATUS.MARCADOR_BASE.test(emp)) continue;
          
          totalObras++;
          
          const statusObra = String(dadosInfo[i][C_STATUS - 1] || "").trim().toUpperCase();
          if (statusObra === "FINALIZADA") {
            obrasFinalizadas++;
          } else {
            obrasAtivas++;
          }
          
          const pendencias = String(dadosInfo[i][C_PEND - 1] || "").trim();
          if (pendencias !== "") {
            qtdObrasAtrasadas++;
          }
        }
      }
    }
    
    // 2. LOGÍSTICA (PEDIDOS HOUSI)
    const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
    const contagemStatusPedidos = {};
    
    if (abaPedidos) {
      const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
      const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
      const lastPed = abaPedidos.getLastRow();
      
      if (lastPed >= iniPed) {
        const C_EMP = C_PED.EMP || 1;
        const C_STATUS = C_PED.STATUS || 9;
        const maxColPed = Math.max(C_EMP, C_STATUS);
        const dadosPed = abaPedidos.getRange(iniPed, 1, lastPed - iniPed + 1, maxColPed).getDisplayValues();
        
        for (let i = 0; i < dadosPed.length; i++) {
          const emp = String(dadosPed[i][C_EMP - 1] || "").trim();
          if (!emp || CONFIG.STATUS.MARCADOR_BASE.test(emp)) continue;
          
          let statusText = String(dadosPed[i][C_STATUS - 1] || "PENDENTE").trim().toUpperCase();
          if (statusText === "") statusText = "PENDENTE";
          
          contagemStatusPedidos[statusText] = (contagemStatusPedidos[statusText] || 0) + 1;
        }
      }
    }

    // 3. OCORRÊNCIAS (VILÕES - ABA OCORRÊNCIAS)
    const abaOcor = ss.getSheetByName(CONFIG.SHEETS.OCORRENCIAS);
    const mapaVilaoOcorrencias = {};
    
    if (abaOcor) {
      const C_OCO = resolveSheetColumns_(abaOcor, CONFIG.HEADERS_COLS.OCORRENCIAS, CONFIG.COLUMNS.OCORRENCIAS);
      const iniOco = obterLinhaInicialPorAba(CONFIG.SHEETS.OCORRENCIAS);
      const lastOco = abaOcor.getLastRow();
      
      if (lastOco >= iniOco) {
        const C_EMP = C_OCO.EMP || 1;
        const C_UNI = C_OCO.UNI || 2;
        const C_STATUS = C_OCO.STATUS_GERAL || 9;

        const maxColOco = Math.max(C_EMP, C_UNI, C_STATUS);
        const dadosOco = abaOcor.getRange(iniOco, 1, lastOco - iniOco + 1, maxColOco).getDisplayValues();
        
        for(let i = 0; i < dadosOco.length; i++) {
           const emp = String(dadosOco[i][C_EMP - 1] || "").trim().toUpperCase();
           const uni = String(dadosOco[i][C_UNI - 1] || "").trim();
           const statusGeral = String(dadosOco[i][C_STATUS - 1] || "").trim().toUpperCase();
           
           if(!emp || !uni) continue;
           if(statusGeral === "CONCLUÍDO" || statusGeral === "CONCLUIDO" || statusGeral === "CANCELADO") continue;
           
           const chave = `${emp} ${uni}`;
           mapaVilaoOcorrencias[chave] = (mapaVilaoOcorrencias[chave] || 0) + 1;
        }
      }
    }
    
    const topOcorrencias = Object.keys(mapaVilaoOcorrencias)
      .map(k => ({ label: k, value: mapaVilaoOcorrencias[k] }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 5);

    const pedidosLabels = Object.keys(contagemStatusPedidos).sort();
    const pedidosData = pedidosLabels.map(l => contagemStatusPedidos[l]);

    return {
      success: true,
      data: {
        obras: {
          total: totalObras,
          ativas: obrasAtivas,
          finalizadas: obrasFinalizadas,
          atrasadas: qtdObrasAtrasadas,
          noPrazo: obrasAtivas - qtdObrasAtrasadas >= 0 ? obrasAtivas - qtdObrasAtrasadas : 0
        },
        pedidos: {
          labels: pedidosLabels,
          data: pedidosData
        },
        ocorrencias: topOcorrencias
      }
    };
  } catch (err) {
    return {
      success: false,
      errorMsg: String(err.message),
      errorStack: String(err.stack)
    };
  }
}

/**
 * Basic tests for Payments module
 */

function testarPagamentos() {
  try {
    // Smoke test: ensure sheet exists and functions callable
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('PAGAMENTOS');
    if (!sh) {
      Logger.log('SKIP: PAGAMENTOS sheet not found');
      return;
    }

    // Test criarPagamento with minimal payload
    const id = criarPagamento({ CHAVE_SERVICO: 'TEST-KEY-123', PRESTADOR: 'TESTE', VALOR: 100, TOTAL_SERVICO: 100 });
    Logger.log('Created payment: ' + id);

    // Test validarSoma
    const res = validarSoma('TEST-KEY-123');
    Logger.log('Validar soma: ' + JSON.stringify(res));

  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }
}

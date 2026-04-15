/**
 * Basic tests for Payments module
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

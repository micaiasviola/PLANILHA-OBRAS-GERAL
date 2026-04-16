/**
 * Compatibility wrappers for menu callbacks that reference older function names.
 * Added to ensure menu items in Main.gs resolve to existing implementations.
 */

/**
 * Wrapper: manter nome histórico esperado pelo menu.
 * Redireciona para a implementação existente que configura a coluna de envio
 * na aba Informações Gerais.
 */
function configurarColunaEnvioFaseEntrega() {
  // redireciona para a função existente
  if (typeof configurarColunaEnviarEntregaInfoGerais === 'function') {
    return configurarColunaEnviarEntregaInfoGerais();
  }
  throw new Error('Função configurarColunaEnviarEntregaInfoGerais não encontrada.');
}

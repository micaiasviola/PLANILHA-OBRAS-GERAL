Relatório: usos de getRange com índices numéricos

O agente detectou 9 ocorrências de getRange(row, col) com colunas numéricas — estas são frágeis e devem ser migradas para uso de resolveSheetColumns_ ou referências por nome.

Ocorrências encontradas (resumo):
- SheetObra.gs -> 1 ocorrência
- SheetOcorrencias.gs -> 1 ocorrência
- SyncLogic.gs -> 1 ocorrência
- Tests.gs -> 3 ocorrências
- tests.js -> 2 ocorrências

Sugestão de revisão manual:
- Revisar cada arquivo e substituir getRange(r, c) que refere-se a colunas de dados por resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS, fallback) ou usar constantes em Config.gs.
- Garantir que UUID (coluna AY) não seja alterada durante migrações.

Próximos passos recomendados:
1. Criar uma issue por arquivo com contexto e sugestões de mudança (posso gerar automaticamente se desejar).
2. Implementar as mudanças em uma branch separada e abrir PRs com testes (rodar npm run test).

PRs/Referências:
- Branch atual com arquivos do agente: agent/branch-protection-template
- Agent Guard workflow: .github/workflows/agent-guard.yml

Posso criar issues separadas para cada arquivo agora. Deseja que eu crie issues ou apenas o PR de relatório? (responda pela interface)
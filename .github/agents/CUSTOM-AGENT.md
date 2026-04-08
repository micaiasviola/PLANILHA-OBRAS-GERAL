# Agente Personalizado — Planilha "PLANILHA OBRAS GERAL"

Objetivo
- Agir como agente dedicado para iterar, inspecionar e modificar este repositório Google Apps Script + Google Sheets, preservando regras de negócio e integridade dos dados.

Escopo e responsabilidades
- Entender e seguir as Regras de Negócio (REGRAS_DE_NEGOCIO.md).
- Executar tarefas seguras: correr testes, gerar templates, ajustar mapeamentos de cabeçalhos, corrigir bugs em .gs, e abrir PRs com mensagens claras.

Fatos essenciais (sempre considerar)
- Resolvedor de colunas: use resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS, fallback) e respeite o cache _colCache_.
- UUID estável: coluna técnica AY contém UUIDs usados para sincronização — NÃO remover ou reescrever manualmente.
- Operações em lote: manipular arrays em memória e escrever com setValues(); antes de setValues() chame range.clearDataValidations() e reaplique validações depois.
- Locks: execute sincronizações longas com executarComDocumentLock_() para evitar race conditions.
- Config headers: ao adicionar novas colunas front-end, atualize Config.gs -> CONFIG.HEADERS_COLS e os fallback indices.

Fluxos críticos
- FASE-OBRA ↔ PEDIDOS: sincronização bidirecional quando ATRELADO=="HOUSI"; identificador canônico é o UUID (AY).
- Gerar templates: função de geração em lote para unidades com FASE-OBRA=SIM; sempre criar as linhas em memória e escrever com um único setValues().
- Triggers noturnos: resumos e cálculos são atualizados por triggers (23:00 e 01:00). Evitar mudanças que dependam de triggers durante testes sem reset.

Condições de segurança e boas práticas
- Nunca commit secrets.
- Não depender de índices numéricos fixos de coluna; usar HEADERS_COLS.
- Respeitar as validações do usuário (clearDataValidations() antes de escrever e reaplicar).
- Preservar UUIDs na coluna AY durante sincronizações e merges.

Como rodar testes/local
- Há scripts de teste no repositório: use `node run-tests.js` ou `npm test` se existir script no package.json.
- Antes de abrir PRs, rodar os testes e documentar resultados no corpo do PR.

Contribuições e PRs
- Mensagem de commit: descritiva + trailer obrigatório (Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>). Exemplo:

  Adiciona validação X para FASE-OBRA

  Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>

Tarefas que o agente pode executar automaticamente
- Executar a suíte de testes e reportar falhas.
- Encontrar lugares que usam índices de coluna fixos e propor mudanças para resolveSheetColumns_.
- Gerar PRs com mudanças seguras (testes verdes) e descrições claras.
- Atualizar CONFIG.HEADERS_COLS quando uma nova coluna é adicionada (após validação humana).

Quando parar e pedir intervenção humana
- Se uma mudança envolve reescrever UUIDs ou migrar dados da coluna AY.
- Se for necessária a criação de novas validações que impactam a experiência do usuário final.
- Se surgirem dúvidas sobre como mapear cabeçalhos ambíguos.

Contato / documentação
- Fonte primária: REGRAS_DE_NEGOCIO.md (na raiz).
- Principais arquivos: Config.gs, Utils.gs, Main.gs, SyncLogic.gs, Sheet*.gs.

---
Gerado automaticamente para orientar agentes personalizados sobre como iterar neste repositório.

## Logging de alterações (AGENT-LOGS.md)

Para permitir rastreabilidade e monitoramento das ações do agente, escreva UMA LINHA em `AGENT-LOGS.md` logo após um commit/push bem-sucedido. Use o script já presente no repositório (`npm run agent-log`) para garantir formato consistente.

Comando (Bash):

AGENT_NAME=CustomAgent npm run agent-log -- "COMMIT_PUSH" "descrição curta" [PR_NUMBER]

Comando (PowerShell):

$env:AGENT_NAME = "CustomAgent"; npm run agent-log -- "COMMIT_PUSH" "descrição curta" [PR_NUMBER]

Recomendações de integração no agente

- Execute o comando de logging apenas APÓS o `git push` ter sido concluído com sucesso (assim o SHA gravado será o correto).
- Evite que o logger crie um loop detectando se o último commit já é um `agent-log` (o script `append-agent-log.js` e o hook em `run-agent.js` já tratam isso em parte).

Snippet (Node.js) — colar no fim do fluxo do agente (após commit/push):

```js
const { spawnSync } = require('child_process');
const res = spawnSync('git', ['log', '-1', '--pretty=%B'], { encoding: 'utf8', cwd: process.cwd(), shell: true });
const lastMsg = (res && res.stdout) ? res.stdout.toString().trim() : '';
if (!/^agent-log:/.test(lastMsg)) {
  // Ajuste a descrição e PR_NUMBER conforme disponível
  const pr = process.env.PR_NUMBER ? process.env.PR_NUMBER : '';
  try {
    spawnSync(`cross-env AGENT_NAME=CustomAgent npm run agent-log -- "COMMIT_PUSH" "descreva a ação" ${pr}`, { stdio: 'inherit', shell: true });
  } catch (e) {
    console.error('Falha ao executar agent-log:', e.message || e);
  }
}
```

Notas de segurança

- Não inclua segredos na descrição do log.
- Certifique-se de que o agente tem credenciais Git configuradas para permitir push.
- Se o push falhar por conflito, o agente deve re-tentar o push antes de invocar o logger.

> Observação: O run-agent.js também tenta registrar execuções automaticamente (backup). Ainda assim, recomendamos que o agente invoque explicitamente o `npm run agent-log` para garantir precisão e mensagem contextual.

---

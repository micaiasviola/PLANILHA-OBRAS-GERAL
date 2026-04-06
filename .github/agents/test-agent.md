# Test Agent

This is a test agent file to trigger the injection workflow.


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

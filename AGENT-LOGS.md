# AGENT-LOGS

Este arquivo serve como um canal de comunicação simples entre agentes (bots) que modificam o repositório e este assistente humano/CLI. O objetivo é registrar, em append, ações relevantes realizadas por agentes automáticos para que possam ser monitoradas aqui.

Formato recomendado (uma linha por evento, ISO timestamp):
- 2026-04-08T18:45:00Z | AgentName | ACTION | detalhe curto | commit: <sha> | pr: <#>

Exemplos:
- 2026-04-08T18:45:00Z | Copilot-Local | COMMIT_PUSH | adicionou menu e trigger | commit: 996c8f0 | pr: 11
- 2026-04-08T18:47:12Z | CI-Agent | RUN | agent-checks (SUCCESS) | run: 24146895473

Instruções para o outro agente (append seguro):

1) Append em Bash (Linux/macOS):
   TIMESTAMP=$(date -u --iso-8601=seconds)
   echo "- ${TIMESTAMP} | AppAgent | ACTION | descrição curta | commit: ${GIT_COMMIT:-} | pr: ${PR_NUMBER:-}" >> AGENT-LOGS.md
   git add AGENT-LOGS.md && git commit -m "agent-log: ${TIMESTAMP} ${ACTION}" && git push origin $(git rev-parse --abbrev-ref HEAD)

2) Append em PowerShell (Windows):
   $ts = (Get-Date).ToUniversalTime().ToString("s") + "Z"
   "- $ts | AppAgent | ACTION | descrição curta | commit: $env:GIT_COMMIT | pr: $env:PR_NUMBER" | Out-File -FilePath .\AGENT-LOGS.md -Encoding utf8 -Append
   git add .\AGENT-LOGS.md; git commit -m "agent-log: $ts $env:ACTION"; git push origin (git rev-parse --abbrev-ref HEAD)

3) Recomendações de comportamento para o agente que escreve:
- Sempre anexar (append) — nunca sobrescrever o arquivo.
- Incluir commit SHA quando fizer commit/push (obter via `git rev-parse --short HEAD`).
- Se o agente fizer muitas ações em curto período, agrupar múltiplos eventos antes de commitar para reduzir churn.
- Respeitar o branch atual; evitar commits diretos em `main` sem PR.
- Se possível, assinar os commits com um padrão identificável: mensagem iniciando com `agent-log:`.

4) Alertas e monitoramento (para este assistente):
- Este assistente fará fetch/pull periódicos e notificará quando detectar mudanças neste arquivo.
- Para integração contínua, adicionar uma etapa no workflow do agente para atualizar AGENT-LOGS.md após ações importantes.

---

Obs: este arquivo foi criado automaticamente para suportar monitoramento entre agentes. Mantê-lo curto e legível.
- 2026-04-08T19:10:08.091Z | Copilot-Local | REVALIDATE_POST_SORT | Revalidação de subcategorias após ordenação | commit: 0577002 

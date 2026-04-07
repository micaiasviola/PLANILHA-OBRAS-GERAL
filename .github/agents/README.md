Agente personalizado - instruções de uso

Este diretório contém um agente CLI simples que auxilia em tarefas seguras no repositório.

Como executar localmente:
  node .github\agents\run-agent.js test
  node .github\agents\run-agent.js find-fixed-columns
  node .github\agents\run-agent.js check-ay

Notas de segurança:
- O agente por padrão não altera arquivos. Adições que apliquem mudanças devem ser feitas manualmente ou com flags explícitas.
- Para operações que precisem de autenticação (abrir PRs), use GitHub Actions ou gh/Personal Access Tokens fora do repositório.

Integrando com npm:
- Um script npm "agent" foi adicionado ao package.json para facilitar execução: npm run agent -- <comando>

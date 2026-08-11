# Librevia

Suíte de documentos e planilhas para desktop. Offline, sem dependência de serviços externos.
Formatos-alvo: **DOCX, XLSX, PDF**.

> Nome provisório. Ver `docs/00-plano-tecnico.md` §8, item 7.

**Estado atual: Fase 2 concluída** — fundação, ciclo completo de arquivos e **editor de
documentos** com formatação de texto, títulos, listas, recuo, espaçamento, alinhamento,
tabelas, imagens, links, quebra de página, localizar/substituir e configuração de página
(A4/Carta, retrato/paisagem, margens).

Próxima etapa: Fase 3 — exportação para PDF e impressão.

## Formatos

| Extensão | O que guarda |
| -------- | ------------ |
| `.sdoc`  | formato interno: documento completo, com formatação — sem perda |
| `.txt`   | apenas texto; salvar nele descarta formatação, e o aplicativo avisa antes |

DOCX chega na Fase 4 e XLSX na Fase 7.
O plano completo, com arquitetura, avaliação de bibliotecas, licenças, riscos e as 9 fases,
está em [`docs/00-plano-tecnico.md`](docs/00-plano-tecnico.md).

## Requisitos

- Node.js 22 ou superior

## Comandos

```bash
npm install      # única etapa que precisa de rede
npm run dev      # aplicativo em modo desenvolvimento, com HMR
npm run build    # verificação de tipos + build de produção em out/
npm run verify   # tipos + lint + testes + licenças (o mesmo que o CI roda)
npm test         # testes unitários
```

## Arquitetura em uma tela

```text
main (Node.js)  →  preload (contextBridge)  →  renderer (React)
```

O renderer **não tem acesso ao Node.js**: `contextIsolation` ligado, `nodeIntegration`
desligado, `sandbox` ligado. Ele fala com o processo main exclusivamente por `window.api`,
cuja superfície é declarada em `src/shared/api.ts` — ampliá-la exige editar aquele arquivo
de propósito, e não é efeito colateral de mexer no preload.

| Pasta            | Responsabilidade                                                        |
| ---------------- | ----------------------------------------------------------------------- |
| `src/main/`      | ciclo de vida, janelas, arquivos, impressão, IPC                        |
| `src/preload/`   | ponte `contextBridge` — encaminhador, sem lógica                        |
| `src/renderer/`  | interface React: editor de documentos, planilha, páginas                |
| `src/services/`  | lógica pura: modelos, DOCX/XLSX, fórmulas, PDF                          |
| `src/shared/`    | contratos de IPC, tipos e erros usados pelas três camadas               |

### O renderer não escolhe caminhos de arquivo

Duas travas no processo main, que valem mesmo se o renderer for comprometido por um
documento malicioso:

- **Gravação** (`file:save`) só aceita caminhos autorizados na sessão — isto é, abertos ou
  escolhidos pelo próprio usuário num diálogo nativo.
- **Leitura por atalho** (`file:open-recent`) só aceita caminhos que já estejam na lista de
  recentes.

Qualquer outro caminho é recusado com `PATH_NOT_AUTHORIZED`, sem chegar ao disco.

### Duas regras que o linter faz cumprir

1. **`src/services/` e `src/shared/` não importam `electron`, `react` nem APIs do Node.**
   São camadas puras: rodam nos dois lados e são testáveis sem Electron rodando.
2. **O renderer não importa `electron` nem `node:*`.** Só `window.api`.

Quebrar qualquer uma das duas falha o `npm run lint`, não a revisão de código.

## Portões de qualidade

O `npm run verify` reprova o código se:

- os tipos não fecharem (TypeScript em modo estrito, com `noUncheckedIndexedAccess`);
- as fronteiras de arquitetura acima forem violadas;
- os testes falharem — incluindo os que travam as preferências de segurança da janela;
- alguma dependência trouxer licença fora da allowlist (MIT, BSD, Apache-2.0, ISC e afins).

O portão de licenças existe porque o aplicativo é de uso corporativo: uma dependência
GPL/AGPL entrando sem querer é um problema jurídico, não técnico. Ver §4.4 do plano.

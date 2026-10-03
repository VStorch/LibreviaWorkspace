# AGENTS.md

Instruções para agentes de código que trabalham neste repositório. O contexto do projeto está no
[README](README.md); os padrões, explicados para pessoas, no [CONTRIBUTING](CONTRIBUTING.md). Este
arquivo é a versão curta e operacional.

## O projeto

Editor de documentos e planilhas para desktop, offline. Electron 43 + React 19 + TypeScript
estrito, com um serviço de formatos em .NET 10 (`sidecar/`) que lê e grava DOCX e XLSX. A promessa
central: **salvar regrava só o que foi editado**; o resto do pacote volta byte a byte.

## Comandos

| Comando | Quando |
| --- | --- |
| `npm run typecheck` | depois de qualquer mudança em TypeScript |
| `npm run lint` | idem; cobra as fronteiras entre camadas e os limites de função |
| `npx vitest run <arquivo>` | testes de unidade do que você mexeu |
| `npm run sidecar:test` | mudança em `sidecar/` |
| `npm run sidecar:build` | antes de testes que falam com o sidecar publicado, e antes do e2e |
| `npm run e2e` | build + sidecar + Playwright; só quando a mudança toca a interface |
| `npm run format:check` / `npm run format` | conferir e corrigir a formatação |
| `npm run verify` | tudo que o CI roda, menos a formatação |

Teste só o que a mudança tocou. Rodar a suíte inteira a cada passo é desperdício.

## Regras que o código cobra

- **Fronteiras:** `src/services/` e `src/shared/` não importam `electron`, `react`, `node:*`,
  `@main/*` nem `@renderer/*`. O renderer não importa `electron`, `node:*` nem `@main/*`; fala com
  o main só por `window.api`. O preload só importa de `@shared`.
- **Limites de função** (em `src/`, fora dos testes): `complexity` 20,
  `max-lines-per-function` 80, `max-depth` 4, `max-params` 5. As violações antigas estão
  congeladas em `eslint-suppressions.json`. Não acrescente entradas lá para fazer código novo
  passar; divida a função. Ao corrigir uma função da lista, rode `npm run lint:prune`.
- **`eslint-disable`** só na linha, com o motivo escrito acima, e só quando a causa é externa
  (por exemplo, assinatura de callback do ProseMirror).
- **IPC:** todo canal tem esquema Zod em `src/shared/`.
- **Protocolo do sidecar:** `src/main/sidecar/protocol.ts` e
  `sidecar/src/Librevia.Format/Protocol/Frame.cs` mudam juntos.
- **Formatos próprios:** mudou o que o `.sdoc` ou o `.ssheet` guardam, suba `SDOC_VERSION` ou
  `SSHEET_VERSION` e mantenha a leitura das versões antigas.
- **Avisos de perda:** use as constantes de `sidecar/src/Librevia.Format/Docx/Inventory.cs`;
  nunca reescreva a frase.
- **Textos da interface:** toda frase visível vai no catálogo `src/shared/i18n/catalog/`, com
  `pt` e `en`. O `npm run i18n:check` acusa frase solta no código.
- **Rede:** o aplicativo não acessa a internet. Não acrescente nada que busque algo fora.
- **Dependência nova** precisa de licença permissiva; o portão de licenças reprova o resto.

## Estilo

- Identificadores em **inglês**. Comentários, mensagens ao usuário e nomes de teste em
  **português**.
- Comentários explicam **por quê**, não o quê, e acompanham a densidade do código em volta.
- Comportamento copiado do Word, do Excel ou do LibreOffice leva um comentário dizendo de quem e
  por quê.
- README e MANUAL: linguagem simples e exata, frases curtas, exemplos concretos, números medidos.

## Commits

- [Conventional Commits](https://www.conventionalcommits.org) em inglês:
  `type(scope): description`. Escopos vêm do código (`docx`, `xlsx`, `pagination`, `lists`,
  `comments`, `revisions`, `notes`, `equations`, `sidecar`, `ui`…).
- **Corpo curto ou nenhum:** uma ou duas frases com o porquê, só quando o título não basta.
- **Nenhum rodapé de atribuição** — nem `Co-Authored-By` de agente, nem link de sessão, nem
  "Generated with". Isso vale também para descrições de pull request.
- **Autor e committer são sempre o dono do repositório**, pela configuração do git. Nunca use
  `--author` nem variáveis `GIT_AUTHOR_*` / `GIT_COMMITTER_*`.
- Não faça commit nem push sem que o dono peça.

## O que não entra no repositório

- Os documentos reais usados para calibrar o projeto (o "corpus"). Testes que dependem deles são
  pulados sem `LIBREVIA_CORPUS_DOC`.
- Scripts de comparação visual e de medição feitos durante uma tarefa: ficam fora da árvore.
- A pasta `docs/`, que é ignorada pelo git. Não a cite em documentação publicada.
- Binários do sidecar (`resources/sidecar/`, `bin/`, `obj/`) e saídas de build.

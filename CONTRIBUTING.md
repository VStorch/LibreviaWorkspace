# Contribuindo com o Librevia

Obrigado pelo interesse. Este guia diz como relatar um problema, como preparar uma mudança e
quais padrões o código segue. A maior parte desses padrões é conferida por máquina: se o
`npm run verify` passa, a mudança já está quase toda dentro deles.

---

## 📑 Índice

- [O relato mais útil](#o-relato-mais-útil)
- [Preparar o ambiente](#preparar-o-ambiente)
- [Antes de abrir o pull request](#antes-de-abrir-o-pull-request)
- [Fronteiras entre as camadas](#fronteiras-entre-as-camadas)
- [Tamanho e complexidade das funções](#tamanho-e-complexidade-das-funções)
- [Contratos que mudam juntos](#contratos-que-mudam-juntos)
- [Textos da interface](#textos-da-interface)
- [Testes](#testes)
- [Dependências](#dependências)
- [Idioma](#idioma)
- [Commits](#commits)

---

## O relato mais útil

**É um arquivo que não abriu direito.** O projeto é calibrado com documentos reais, e cada
`.docx` ou `.xlsx` que se comporta de um jeito inesperado ensina mais que uma função nova.

Use o modelo **"Arquivo que abriu errado"** ao abrir a issue. Se puder, anexe o arquivo. Se ele
for confidencial, descreva o que o Word ou o LibreOffice mostram e o que o Librevia mostrou.

Falhas de segurança **não** vão em issue pública: veja a [política de segurança](SECURITY.md).

---

## Preparar o ambiente

Você precisa do **Node.js 22** ou mais novo e do **.NET SDK 10**.

```bash
git clone https://github.com/VStorch/LibreviaWorkspace.git
cd LibreviaWorkspace
npm install
npm run sidecar:build   # publica o serviço de formatos .NET
npm run dev             # abre o aplicativo com recarga automática
```

A lista completa de comandos está no [README](README.md#rodando-localmente).

---

## Antes de abrir o pull request

```bash
npm run format:check
npm run verify
```

São as mesmas verificações que o CI roda: formatação, tipos, lint, testes do TypeScript e do
.NET, licenças e traduções. Se passam na sua máquina, passam no CI. Para corrigir a formatação,
`npm run format`.

Se a mudança mexe na interface, rode também os testes de ponta a ponta:

```bash
npm run e2e
```

O modelo de pull request traz uma lista curta do que conferir.

---

## Fronteiras entre as camadas

O código tem quatro partes, e cada uma só pode importar o que precisa:

| Pasta | Pode importar | Não pode importar |
| --- | --- | --- |
| `src/services/`, `src/shared/` | só código puro | `electron`, `react`, `node:*`, `@main/*`, `@renderer/*` |
| `src/renderer/` | React, `@services/*`, `@shared/*` | `electron`, `node:*`, `@main/*` |
| `src/preload/` | `electron`, `@shared/*` | todo o resto |
| `src/main/` | tudo | — |

Por quê? `services` e `shared` rodam dos dois lados e se testam sem o Electron. O renderer roda
isolado e só fala com o resto por `window.api`.

**O lint cobra isso.** Quem esquecer recebe o erro no `npm run lint`, antes de qualquer revisão.

---

## Tamanho e complexidade das funções

Uma função muito longa ou com muitos caminhos é difícil de ler e de testar. Os limites abaixo
saíram da medição do próprio código e são conferidos pelo ESLint em `src/` (os testes ficam de
fora):

| Regra | Limite | O que mede |
| --- | --- | --- |
| `complexity` | 20 | caminhos possíveis pela função: cada `if`, `case`, `&&`, `??` e laço soma um |
| `max-lines-per-function` | 80 | linhas de código, sem contar comentários e linhas em branco |
| `max-depth` | 4 | blocos aninhados uns dentro dos outros |
| `max-params` | 5 | parâmetros; acima disso, use um objeto de opções |

Para comparar: a função típica do projeto tem complexidade 3 e 7 linhas.

### Os limites sobem em degraus

Quando os limites entraram, 76 funções já passavam deles. Em vez de reescrevê-las de uma vez,
elas foram **congeladas** em `eslint-suppressions.json`, que guarda quantas exceções cada arquivo
tem. A regra passa a valer assim:

- **código novo** acima do limite reprova;
- uma função antiga que **piora** também reprova, porque o arquivo passa a ter mais exceções do
  que o registrado;
- quem **corrige** uma função da lista roda `npm run lint:prune`, que tira a exceção do arquivo.
  Sem isso o lint reclama de uma exceção que não existe mais.

A meta é baixar a complexidade para 15 quando a lista estiver menor.

### Quando a exceção é legítima

Às vezes a assinatura não é nossa. Um callback do ProseMirror, por exemplo, recebe seis
parâmetros porque a biblioteca manda seis. Nesse caso, desligue a regra **só naquela linha** e
escreva o motivo:

```ts
// A assinatura é a do ProseMirror, e não nossa.
// eslint-disable-next-line max-params
handleDoubleClickOn(view, _pos, node, nodePos, _event, direct) {
```

---

## Contratos que mudam juntos

- **IPC:** todo canal tem um esquema Zod em `src/shared/`, conferido nas duas pontas. Canal
  novo, esquema novo.
- **Protocolo do sidecar:** os dois lados, `src/main/sidecar/protocol.ts` e
  `sidecar/src/Librevia.Format/Protocol/Frame.cs`, mudam no mesmo commit. O teste
  `sidecar-real.test.ts` conversa com o executável de verdade e pega a diferença.
- **Formatos próprios:** `.sdoc` e `.ssheet` têm número de versão. Mudou o que eles guardam,
  suba a versão e mantenha a leitura das versões antigas.
- **Avisos de perda:** os rótulos do `Inventory.cs` são constantes. Use a constante; nunca
  escreva a frase de novo em outro lugar.

---

## Textos da interface

Toda frase que a pessoa lê vem do catálogo em `src/shared/i18n/catalog/`, com as duas línguas lado
a lado:

```ts
'comments.action.reply': { pt: 'Responder', en: 'Reply' },
```

O compilador acusa uma entrada sem uma das línguas. O `npm run i18n:check` acusa uma frase
escrita direto no código, fora do catálogo.

---

## Testes

O teste acompanha a mudança, na camada certa:

| O que mudou | Onde testar |
| --- | --- |
| regra de negócio, modelo, fórmula, formato | Vitest, ao lado do arquivo (`*.test.ts`) |
| leitura ou escrita de DOCX e XLSX | xUnit, em `sidecar/tests/` |
| o que só aparece com o aplicativo montado | Playwright, em `e2e/` |

A regra de negócio mora em camada pura justamente para ser testada sem subir o Electron.

Os documentos reais usados para calibrar o projeto **não entram no repositório**. Os testes que
dependem deles são pulados quando a variável `LIBREVIA_CORPUS_DOC` não está definida.

---

## Dependências

Antes de acrescentar uma, pese o que ela traz junto. Toda dependência de produção passa pelo
portão de licenças (`npm run licenses:check` e `npm run sidecar:licenses`), que só aceita licenças
permissivas — MIT, BSD, Apache-2.0, ISC e parecidas — e reprova o pacote que não declara licença.

---

## Idioma

- **Identificadores em inglês:** nomes de arquivo, função, variável e tipo.
- **Comentários, mensagens ao usuário e nomes de teste em português.**
- **Commits em inglês.**

---

## Commits

[Conventional Commits](https://www.conventionalcommits.org), com escopo sempre que houver um:

```text
type(scope): description
```

- **Tipos em uso:** `feat`, `fix`, `refactor`, `perf`, `test`, `docs`, `build`, `ci`, `chore`
  e `revert`.
- **Escopos:** vêm do próprio código, como `docx`, `xlsx`, `styles`, `pagination`, `lists`,
  `sections`, `comments`, `revisions`, `notes`, `equations`, `sidecar`, `print` e `ui`.
- **O corpo é curto, ou não existe.** Uma ou duas frases dizendo *por quê*, quando o título não
  basta. Se o título já diz tudo, não escreva corpo.

```text
fix(docx): line height multiplies the font's height, not its size

LibreOffice puts 12.98 pt in Arial 10 pt; we were applying the factor to the size.
```

- **Sem rodapé:** nada de `Co-Authored-By` de ferramentas ou de assinatura automática no fim da
  mensagem.

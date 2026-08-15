# Librevia

Suíte de documentos e planilhas para desktop. Offline, sem dependência de serviços externos.
Formatos-alvo: **DOCX, XLSX, PDF**.

> Nome provisório.

**Estado atual: editor de documentos e planilhas funcionando.**

| Entregue | |
| --- | --- |
| Documentos | formatação, títulos, listas, recuo, espaçamento, alinhamento, tabelas, imagens, links, quebra de página, localizar/substituir |
| PDF e impressão | cabeçalho, rodapé, numeração, margens e orientação |
| DOCX | abre e grava com **edição cirúrgica** — ver abaixo |
| Planilhas | grade de 10 mil linhas, seleção de intervalo, alça de preenchimento, área de transferência, abas, formatação, congelamento e inserir/excluir linhas e colunas |

Em aberto: motor de fórmulas, XLSX, mesclagem de células e o instalador.

## Formatos

| Extensão | O que guarda |
| -------- | ------------ |
| `.sdoc`   | formato interno do documento: completo, com formatação — sem perda |
| `.ssheet` | formato interno da planilha: valores, fórmulas e formatação — sem perda |
| `.docx`   | Word; abre e grava preservando o que não foi editado |
| `.txt`    | apenas texto; salvar nele descarta formatação, e o aplicativo avisa antes |
| `.pdf`    | saída apenas (exportação e impressão) |

XLSX ainda não abre. **ODT está fora do escopo**: não existe biblioteca madura e
de licença permissiva em nenhum ecossistema — no .NET tudo que presta é
comercial, e no npm a opção viável tem seis meses de vida e um mantenedor.
Documentos `.docx` funcionam vindos tanto do Word quanto do LibreOffice.

## Editar um `.docx` não custa o que você não editou

Regenerar o pacote OOXML a cada salvamento apaga em silêncio comentários,
revisões, notas e formas — tudo que o importador não entendeu. Aqui a gravação é
**cirúrgica**: cada bloco que você não tocou volta para o arquivo exatamente como
estava, e só os editados são gerados de novo.

Medido num documento real de 105 blocos: salvar sem editar reescreve **zero**
blocos; editar um parágrafo reescreve **um**. Do pacote inteiro, só
`word/document.xml` muda — as outras 25 partes saem idênticas byte a byte.

A fidelidade não vem de entender o OOXML. Vem de **não mexer** no que não foi
editado. O mecanismo: cada bloco ganha um id na abertura, e a gravação só
sintetiza aquele cuja impressão digital mudou.

Duas coisas que o aplicativo distingue e a maioria mistura:

- **perda** — some ao salvar (você editou o parágrafo que ancorava um comentário);
- **invisibilidade** — continua no arquivo, mas o editor não mostra (revisões,
  notas de rodapé, o logotipo do cabeçalho).

São avisos diferentes porque são problemas diferentes. Um alerta genérico é um
alerta que o usuário aprende a ignorar.

## Por que há .NET num projeto Electron

Um editor normal preserva o que você não editou. Regenerar o pacote OOXML do zero
a cada salvamento apaga em silêncio comentários, revisões, notas e formas — tudo
que o importador não entendeu. A saída é **edição cirúrgica**: manter o pacote
original e reescrever só as partes tocadas. Isso exige um DOM tipado completo do
OOXML, e o `DocumentFormat.OpenXml` (MIT, da Microsoft) é o único maduro. Daí o
sidecar em `sidecar/`.

Ele é um **serviço de formato**, não uma segunda aplicação: bytes entram por
stdio, JSON sai. Sem rede, sem porta, sem permissão de escrita — quem grava
continua sendo o processo main, com a autorização de caminho e a gravação atômica
já existentes. Interface, edição e estado ficam todos no Electron.

O protocolo é um quadro binário — `[tamanho do JSON][tamanho do binário][JSON][binário]` —
e não uma linha de texto. Um DOCX é um ZIP, cheio de `0x0a` e `0x00`: qualquer
protocolo delimitado por linha se despedaçaria nele. Os dois lados vivem em
`src/main/sidecar/protocol.ts` e `sidecar/src/Librevia.Format/Protocol/Frame.cs`,
e **precisam mudar juntos** — quem garante isso é `sidecar-real.test.ts`, que
conversa com o executável publicado em vez de com uma imitação.

Se o sidecar morrer, o documento aberto não morre junto: a operação em curso
falha com uma frase compreensível e o próximo pedido sobe um processo novo. E ele
encerra ao perder o stdin, então nem um `SIGKILL` no aplicativo deixa processo
órfão para trás.

## PDF e impressão

Não há biblioteca de PDF: quem gera é o próprio Chromium, pelo `printToPDF`. É o mesmo
motor que desenha o editor, então o PDF sai igual à tela — e o texto continua selecionável,
com as fontes embutidas.

O estilo do conteúdo mora em `services/document/content-styles.ts` e é **a mesma folha**
usada pelo editor e pelo HTML de impressão. Duas folhas divergiriam com o tempo, e o PDF
deixaria de bater com a tela.

Cuidado ao mexer em margens: `printToPDF()` recebe **polegadas** e `webContents.print()`
recebe **pixels CSS**. As duas conversões vivem juntas em `services/pdf/page-setup.ts`,
com teste comparando uma com a outra.

## Requisitos

- Node.js 22 ou superior
- .NET SDK 10 (para o sidecar — ver abaixo)

## Comandos

```bash
npm install            # rede: dependências npm
npm run sidecar:build  # rede na primeira vez: pacotes NuGet
npm run dev            # aplicativo em modo desenvolvimento, com HMR
npm run build          # verificação de tipos + build de produção em out/
npm run verify         # tipos + lint + testes (TS e .NET) + licenças — o mesmo que o CI roda
npm test               # testes unitários
```

O `npm run verify` publica o sidecar antes de rodar os testes, porque o teste de
ponta a ponta conversa com o binário de verdade.

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
GPL/AGPL entrando sem querer é um problema jurídico, não técnico. Ele cobre os dois
ecossistemas — npm e NuGet — e **reprova pacote que não declara licença SPDX**, e não
só o que declara uma proibida: foi assim que a Six Labors trocou Apache-2.0 por licença
própria numa dependência transitiva do ClosedXML.

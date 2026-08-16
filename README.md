# Librevia

Suíte de documentos e planilhas para desktop. Offline, sem dependência de serviços externos.
Formatos-alvo: **DOCX, XLSX, PDF**.

> Nome provisório.

**Estado atual: suíte completa, com instalador para Linux e Windows.**

Manual para quem usa o aplicativo: [MANUAL.md](MANUAL.md).

| Entregue | |
| --- | --- |
| Documentos | formatação, títulos, listas, recuo, espaçamento, alinhamento, tabelas, imagens, links, quebra de página, localizar/substituir |
| PDF e impressão | documentos e planilhas, com cabeçalho, rodapé, numeração, margens e orientação |
| DOCX | abre e grava com **edição cirúrgica** — ver abaixo |
| XLSX | abre e grava com a mesma cirurgia, traduzindo as fórmulas na fronteira |
| Planilhas | grade de 10 mil linhas, seleção de intervalo, alça de preenchimento, área de transferência, abas, formatação, congelamento e inserir/excluir linhas e colunas |
| Fórmulas | 37 funções, referências entre abas, barra de fórmulas — ver abaixo |
| Recuperação | rascunho de oito em oito segundos, oferecido de volta depois de uma queda |
| Somente leitura | graduado: trava só quando o arquivo tem o que se perde ao editar |
| Distribuição | AppImage, `.deb` e instalador NSIS, com os avisos de terceiros |

Em aberto: mesclagem de células, interface para filtros de planilha (os que vêm
do arquivo são preservados) e criar `.docx` do zero.

## Formatos

| Extensão | O que guarda |
| -------- | ------------ |
| `.sdoc`   | formato interno do documento: completo, com formatação — sem perda |
| `.ssheet` | formato interno da planilha: valores, fórmulas e formatação — sem perda |
| `.docx`   | Word; abre e grava preservando o que não foi editado |
| `.xlsx`   | Excel; abre e grava preservando o que não foi editado |
| `.txt`    | apenas texto; salvar nele descarta formatação, e o aplicativo avisa antes |
| `.pdf`    | saída apenas (exportação e impressão) |

**ODT e ODS estão fora do escopo**: não existe biblioteca madura e de licença
permissiva em nenhum ecossistema — no .NET tudo que presta é comercial, e no npm
a opção viável tem seis meses de vida e um mantenedor. Arquivos `.docx` e `.xlsx`
funcionam vindos tanto do Office quanto do LibreOffice.

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

E há uma terceira categoria, que é um **subconjunto da invisibilidade**: o que
some se — e só se — o usuário editar justamente o bloco que o ancora. Comentário,
revisão, nota e campo calculado entram; posicionamento de imagem e decoração,
não. É ela que decide o **somente leitura**, que é padrão e não cadeado: a faixa
diz o que está em jogo e um clique libera. Travar por perda de aparência travaria
o uso do dia a dia — quase todo documento do corpus tem uma imagem ancorada — e o
usuário aprenderia a liberar sem ler.

A classificação mora em `Inventory.cs`, e os rótulos são **constantes** usadas
pelos leitores: uma frase mudada num deles deixaria de casar com a lista e o
documento passaria a abrir editável sem que ninguém percebesse. Sendo constantes,
o compilador não deixa.

## Fórmulas escrevem-se em português

`=SOMA(1,5;2)` soma um e meio com dois. **Vírgula é decimal e ponto e vírgula
separa argumentos** — as duas regras do Excel em português, e elas andam juntas:
com a vírgula ocupada pelo decimal, ela não pode separar nada. Aceitar os dois
papéis tornaria `SOMA(1,5)` ambíguo, e a ambiguidade cairia sempre em cima de
quem digitou um número decimal. Quem escreve `SOMA(A1,B1)` recebe a frase
dizendo qual é o separador, e não um erro genérico.

Os nomes respondem nos dois idiomas: `SOMA` e `SUM` são a mesma função, assim
como `SE`/`IF` e `PROCV`/`VLOOKUP`.

Erro de fórmula é **valor**, não exceção: `=A1/0` produz `#DIV/0!`, que se
propaga por quem depende dela — como no Excel, e ao contrário de derrubar o
recálculo da planilha inteira. A exceção é `#CIRC!`, que o Excel não tem: lá uma
referência circular mostra zero e abre um aviso, deixando um número inventado no
meio da planilha.

Comportamentos do Excel reproduzidos **de propósito**, cada um com o motivo
escrito onde ele está: `=-2^2` vale 4; `="a"="A"` é verdadeiro mas `=1="1"` é
falso; texto dentro de um intervalo é ignorado por `SOMA` mas convertido quando
passado direto; `PROCV` procura aproximado por padrão. O critério é sempre o
mesmo: o mesmo arquivo precisa dar o mesmo número nos dois programas.

Inserir uma linha reescreve as referências, inclusive as de outras abas — sem
isso `SOMA(A1:A3)` continuaria somando onde os dados não estão mais. Renomear
uma aba conserta as fórmulas que a citam, em vez de transformá-las em `#REF!`.

Arrastar a alça de preenchimento leva a **fórmula deslocada**, não o resultado:
`=B2*C2` arrastada para baixo vira `=B3*C3`, e a referência com `$` fica onde
está. Copiar o valor calculado daria a coluna inteira repetindo o número da
primeira linha — e o erro só apareceria no fechamento do mês.

Ainda não há: matrizes dinâmicas, referências de coluna inteira (`A:A`) e
intervalos nomeados.

## O `.xlsx` guarda outro idioma, e a fronteira traduz

O arquivo diz `SUM(A1,B1)`; a tela diz `SOMA(A1;B1)`. A tradução acontece num
lugar só — a fronteira entre o processo main e o serviço de formato — e é feita
**caractere a caractere**, não reconstruindo a fórmula a partir da árvore.

O motivo é medido, não estético. A gravação é cirúrgica também aqui: ela relê o
arquivo original, compara célula a célula e só toca no que mudou. Reconstruir a
fórmula normalizaria espaços e maiúsculas, toda célula pareceria diferente, e
abrir e salvar sem editar reescreveria a planilha inteira — apagando fonte,
alinhamento vertical, recuo e bordas diagonais, tudo que o modelo não representa.

Medido numa planilha do LibreOffice: abrir e salvar sem editar escreve **zero**
células e preserva 18; mudar uma quantidade escreve **quatro** — a célula digitada
e as três fórmulas que dependiam dela. O arquivo volta ao LibreOffice com os
valores recalculados e a formatação de moeda intacta.

A diferença para o DOCX vem de uma medição, não de uma suposição: a biblioteca de
planilha **preserva as partes do pacote que não modela** — uma parte de XML
injetada à mão sobrevive à ida e volta, e gráficos e tabelas dinâmicas
sobrevivem pelo mesmo mecanismo. O que ela regenera é a planilha em si. Então
aqui o risco não é o pacote: é a célula.

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
npm run e2e            # build + testes de ponta a ponta no aplicativo montado
npm run dist           # instaladores AppImage e .deb em release/
npm run dist:win       # instalador NSIS (rodando no Windows)
```

O `npm run verify` publica o sidecar antes de rodar os testes, porque o teste de
ponta a ponta conversa com o binário de verdade.

## Recuperação: o rascunho não é uma segunda cópia do arquivo

De oito em oito segundos, o que está na tela é gravado num rascunho — **nunca por
cima do seu arquivo**. Gravar sozinho no arquivo transformaria "não salvei" em
"salvei sem querer", e desfazer isso exigiria o `.bak`, que existe para outro
problema. A decisão de escrever no arquivo continua sendo só sua.

Depois de uma queda, o aplicativo oferece o rascunho de volta numa faixa com duas
ações. Enquanto ela estiver na tela o autosave **não escreve**: sem isso, ignorar
o aviso e começar a digitar apagaria em segundos justamente o trabalho que ele
existe para devolver.

Recuperar reata o que morreu junto com o processo — a autorização de gravação e
os bytes originais do pacote OOXML. Sem os segundos, salvar por cima do `.docx`
recuperado seria recusado, e você ficaria com o trabalho na tela sem poder
gravá-lo onde ele estava.

## Desempenho, medido

Planilha de 20 mil linhas e **120 mil células**, com 20 mil fórmulas, gerada pelo
LibreOffice. Abrir de ponta a ponta no aplicativo: **≈ 3,5 s**, assim repartidos:

| Etapa | Tempo |
| --- | --- |
| serviço de formatos (ClosedXML lendo o pacote) | 1,9 s na primeira vez, 0,9 s depois |
| `parseWorkbook` (validação zod de 120 mil células) | 245 ms |
| `recalculate` (20 mil fórmulas) | 279 ms |
| serializar o modelo (7,4 MB de JSON) | 50 ms |

Duas decisões saíram daí:

**Compilação antecipada (ReadyToRun) no sidecar, ligada.** Sem ela a primeira
abertura custa 3,8 s; com ela, 1,9 s. No regime permanente ela perde ~100 ms,
porque o código pré-compilado é menos otimizado que o que o JIT produz depois de
aquecer — mas a primeira abertura é a que o usuário sente como travamento, e a
diferença no regime permanente ninguém percebe. Custa 41 MB por plataforma.

**Cache de formato numérico.** Uma planilha grande tem meia dúzia de máscaras e
centenas de milhares de células; interpretar a máscara passou a acontecer uma vez
por máscara distinta. O que sobra dos 300 ms daquele trecho não é nosso: é o
ClosedXML materializando `cell.Style` célula a célula.

O autosave, que serializa o documento inteiro a cada oito segundos, custa **4 ms**
num documento de 6 mil parágrafos — não aparece.

## Empacotamento

`npm run dist` gera AppImage e `.deb`; `npm run dist:win` gera o instalador NSIS.
O sidecar fica **fora** do asar — executável dentro de arquivo compactado não
roda — e cada plataforma leva só o seu binário.

Duas coisas são provisórias e devem ser trocadas antes de distribuir de verdade:
o ícone (`build/icon.png`, gerado por `npm run icon`) e os endereços `.internal`
do mantenedor do `.deb` e do `homepage`. O TLD `.internal` é reservado para uso
interno — é honesto por construção, e melhor que inventar um domínio público que
não existe.

`npm run notices` gera o `THIRD-PARTY-NOTICES.md` que viaja no instalador. Ele
cobre três conjuntos, e não um: as dependências npm de produção, o Electron
inteiro (com Chromium e Node.js) e os pacotes NuGet do sidecar, que é publicado
self-contained e leva o runtime do .NET junto. O portão `licenses:check` olha só
o primeiro — deixar os outros dois de fora daria a impressão de conformidade sem
a conformidade.

Os testes de ponta a ponta rodam também contra o **pacote montado**:

```bash
npm run dist
LIBREVIA_E2E_BINARY=release/linux-unpacked/librevia npm run test:e2e
```

É a diferença entre "os testes passam" e "o instalador funciona": dentro do
pacote, caminhos de recurso, asar e o binário do sidecar ficam em outro lugar.

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

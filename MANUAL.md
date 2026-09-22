# Librevia — manual de uso

Suíte de documentos e planilhas para trabalhar **offline**. Abre e grava os arquivos do Office
(`.docx` e `.xlsx`) preservando o que você não editou.

> 📖 Este manual é para quem **usa** o aplicativo.
> Quem for mexer no código deve ler o [README](README.md).

---

<a id="o-que-tem-aqui"></a>

## 📑 O que tem aqui

| | |
| --- | --- |
| [💿 Instalar](#instalar) | [⚠️ Os avisos, e por que são diferentes](#os-avisos) |
| [📄 Os arquivos que ele abre](#os-arquivos-que-ele-abre) | [💾 Se o aplicativo fechar sozinho](#se-o-aplicativo-fechar-sozinho) |
| [⌨️ Atalhos](#atalhos) | [🛟 Salvar não custa o que você não editou](#salvar-não-custa) |
| [✍️ Formatar texto e parágrafo](#formatar) | [🔤 Ortografia, símbolos e contagem](#ferramentas) |
| [📊 Tabelas e imagens](#tabelas-e-imagens) | |
| [🧮 Fórmulas em português](#fórmulas-em-português) | [🖨️ Imprimir e exportar PDF](#imprimir-e-exportar-pdf) |
| [🔒 Arquivos que abrem travados](#arquivos-que-abrem-travados) | [🚧 Limites conhecidos](#limites-conhecidos) |

---

<a id="instalar"></a>

## 💿 Instalar

Nada aqui precisa de internet. **O aplicativo não acessa a rede em momento nenhum.**

### 🐧 Linux — AppImage

Baixe o arquivo, dê permissão de execução e abra:

```bash
chmod +x Librevia-0.0.0.AppImage
./Librevia-0.0.0.AppImage
```

### 🐧 Linux — pacote `.deb`

Instala no menu de aplicativos como qualquer outro programa:

```bash
sudo apt install ./librevia_0.0.0_amd64.deb
```

### 🪟 Windows

Execute o instalador. Ele instala **para o seu usuário**, sem pedir senha de administrador.

---

<a id="os-arquivos-que-ele-abre"></a>

## 📄 Os arquivos que ele abre

| Extensão | O que é |
| --- | --- |
| `.docx` | documento do Word — abre e grava |
| `.xlsx` | planilha do Excel — abre e grava |
| `.sdoc` | documento do Librevia — guarda tudo, sem perda nenhuma |
| `.ssheet` | planilha do Librevia — idem |
| `.txt` | texto puro; salvar nele descarta formatação, e o app avisa antes |
| `.pdf` | só saída: exportar e imprimir, tanto documento quanto planilha |

> ℹ️ **`.odt` e `.ods` não abrem.** Se você recebe arquivos assim, peça para quem enviou
> salvar como `.docx` ou `.xlsx` — o LibreOffice faz isso pelo menu "Salvar como".

---

<a id="atalhos"></a>

## ⌨️ Atalhos

| Atalho | O que faz |
| --- | --- |
| `Ctrl+N` | novo documento |
| `Ctrl+Shift+N` | nova planilha |
| `Ctrl+O` | abrir |
| `Ctrl+S` | salvar |
| `Ctrl+Shift+S` | salvar como |
| `Ctrl+W` | fechar o arquivo |
| `Ctrl+F` | localizar e substituir |
| `Ctrl+P` | imprimir |
| `Ctrl+Enter` | quebra de página manual |
| `Ctrl+F12` | inserir tabela (pergunta linhas e colunas) |
| `Ctrl+B` / `Ctrl+I` / `Ctrl+U` | negrito, itálico, sublinhado |
| `Ctrl+Shift+=` / `Ctrl+=` | sobrescrito e subscrito |
| `Ctrl+L` / `Ctrl+E` / `Ctrl+R` / `Ctrl+J` | alinhar à esquerda, centralizar, à direita, justificar |
| `Ctrl+1` / `Ctrl+5` / `Ctrl+2` | entrelinha simples, 1,5 e dupla |
| `Ctrl+Alt+1`…`Ctrl+Alt+6` | Título 1 a Título 6 |
| `Ctrl+Alt+0` | volta o parágrafo para corpo de texto |
| `Ctrl+]` / `Ctrl+[` | aumentar e diminuir o recuo |
| `Ctrl+Shift+V` | colar sem formatação |
| `Ctrl+Shift+G` | contar palavras |
| `Ctrl+F10` | mostrar e ocultar as marcas de formatação |

> ℹ️ São os atalhos do Word, com duas exceções que valem explicar. **Corpo de texto** é
> `Ctrl+Alt+0` e não `Ctrl+Shift+N`, porque aqui essa tecla abre uma planilha nova — e ela já
> era assim antes. **Ampliar** mudou para o `+` do teclado numérico, para não engolir o
> `Ctrl+Shift+=` do sobrescrito; reduzir segue em `Ctrl+-`, e o menu **Exibir** tem os dois.
> Sobrescrito e subscrito também respondem a `Ctrl+.` e `Ctrl+,`, para o teclado em que o `=`
> exige duas teclas. **Marcas de formatação** são `Ctrl+F10`, como no LibreOffice, e não o
> `Ctrl+*` do Word: essa combinação já é a da lista com marcadores.

**Na planilha**, `Ctrl+B`, `Ctrl+I` e `Ctrl+U` valem para as células selecionadas. Clicar com
o botão direito numa célula abre o menu de inserir e remover linhas e colunas.

**A alça de preenchimento** é o quadradinho no canto inferior direito da seleção: arraste-o e
as células seguintes são preenchidas. Se a célula de origem tiver uma fórmula, ela é
**deslocada** — `=B2*C2` arrastada para baixo vira `=B3*C3`. Referência com `$` não se mexe.

---

<a id="formatar"></a>

## ✍️ Formatar texto e parágrafo

A barra do documento tem, da esquerda para a direita: estilo do parágrafo, fonte, tamanho, a
formatação do texto, o parágrafo, as listas, o que se insere e a página.

### A letra

Além de negrito, itálico, sublinhado e tachado:

| Botão | O que faz |
| --- | --- |
| **Sobrescrito** e **Subscrito** | põem o trecho acima ou abaixo da linha — `m³`, `H₂O`. Vão e voltam do `.docx` |
| **Caixa alta** | desenha em maiúsculas sem trocar o que você digitou: apagar a formatação devolve o texto original |
| **Versalete** | maiúsculas pequenas, do tamanho das minúsculas |
| **Cor do texto** | a cor da letra |
| **Cor de fundo do texto** | pinta o fundo do trecho com qualquer cor |
| **Destaque** | o marca-texto do Word, que tem catorze cores fixas |

> ℹ️ **Por que dois fundos?** No arquivo do Word eles são coisas diferentes: o marca-texto e o
> sombreamento do trecho. Um documento recebido pode trazer os dois, e o aplicativo mantém cada
> um onde estava.

### A lista de fontes

O seletor mostra, nesta ordem: **as fontes que o documento aberto usa**, as cinco que vêm com o
aplicativo — Calibri, Cambria, Arial, Times New Roman e Courier New — e depois **as instaladas
no seu computador**, em ordem alfabética.

As cinco primeiras funcionam igual em qualquer máquina, porque viajam dentro do instalador. As
suas funcionam na sua máquina; num computador que não as tenha, o documento será desenhado com
outra fonte e pode paginar diferente.

### O diálogo de parágrafo

O botão **¶** — ou o menu **Formatar → Parágrafo…** — abre tudo o que é do parágrafo numa tela
só, e já preenchido com o que o parágrafo tem:

- **alinhamento**: esquerda, centro, direita, justificado;
- **entrelinha**: simples, 1,15, 1,5, dupla, um múltiplo qualquer, ou uma medida em pontos;
- **espaçamento** antes e depois, em pontos;
- **recuo** à esquerda e à direita, em milímetros;
- **primeira linha**: recuo (a primeira linha entra) ou deslocamento (ela sai, e as outras
  entram — o formato da lista e da referência bibliográfica);
- **manter com o próximo**: o parágrafo não fica sozinho no pé da página.

Vale para todos os parágrafos que a seleção tocar. `Enter` aplica, `Esc` fecha sem aplicar.

> ℹ️ **"Simples" não é "1,0".** Espaçamento simples é a altura que a própria fonte pede, e é
> por isso que ele aparece como uma opção e não como um número: nenhum fator o imita, e trocar
> um pelo outro muda onde as páginas quebram.

---

<a id="ferramentas"></a>

## 🔤 Ortografia, símbolos e contagem

### A ortografia é em português e não usa internet

O dicionário de português do Brasil **viaja dentro do instalador**. Na primeira vez que o
aplicativo abre, ele é copiado para a pasta de dados do seu usuário, e é ali que o corretor o
procura. Nenhuma máquina precisa estar na rede para a primeira palavra ser conferida — e esta
é a diferença: deixado por conta própria, o corretor embutido baixaria o dicionário de um
servidor na internet, e num computador sem rede ele simplesmente não marcaria nada, sem avisar.

O que está errado ganha o sublinhado ondulado de sempre, **no corpo do texto e também no
cabeçalho e no rodapé** — é lá que um erro de digitação se repete em todas as folhas.

Clique com o **botão direito** sobre a palavra sublinhada para:

- escolher uma das sugestões, que troca a palavra;
- **Adicionar ao dicionário**, para nunca mais ser avisado dela — vale para sempre, em todos
  os documentos;
- **Ignorar nesta sessão**, que vale até você fechar o aplicativo. Serve para o nome de um
  cliente que aparece dez vezes neste documento e não em outro.

Para desligar a verificação: menu **Ferramentas → Verificação ortográfica**.

### O botão direito no documento

Além das sugestões, o menu tem **recortar**, **copiar**, **colar** e **colar sem formatação**.
Dentro de uma tabela ele traz também as ações dela — ver [Tabelas e imagens](#tabelas-e-imagens).
Um item apagado quer dizer que ele não se aplica ali — "colar" fica apagado quando não há nada
na área de transferência, por exemplo.

### Colar sem formatação

`Ctrl+Shift+V`, o menu **Editar** ou o botão direito. O texto entra com a formatação do
documento, não com a de onde veio: é o que evita que um trecho copiado de uma página da web
chegue com outra fonte, outro tamanho e outra cor. Quebras de linha continuam sendo quebras de
linha — cada uma abre um parágrafo.

### Contar palavras

`Ctrl+Shift+G` ou **Ferramentas → Contar palavras…**. Mostra palavras, caracteres com e sem
espaço, parágrafos e páginas, em duas colunas: o **documento** e a **seleção**. Com nada
selecionado, a coluna da seleção mostra um travessão — "nada selecionado" não é a mesma coisa
que "zero palavras". O diálogo acompanha o que você digita, então pode ficar aberto.

Parágrafo vazio não é contado. É tecla `Enter` batida para abrir espaço, e o Word também não o
conta.

### Caracteres especiais

**Inserir → Caractere especial…** abre o painel com o que se procura de verdade: aspas
tipográficas, travessão, símbolos de moeda, matemática, letras gregas e marcas como © e ®. Cada
símbolo é um botão — clique, ou ande pela grade com as setas e insira com `Enter`. O painel
continua aberto, porque quem o abre em geral quer mais de um símbolo.

O `␣` do grupo "Pontuação" é o **espaço inquebrável**: é ele que impede o `R$` de ficar no fim
de uma linha e o valor na linha seguinte.

### Marcas de formatação (¶)

`Ctrl+F10`, o botão `¶` da barra ou **Exibir → Marcas de formatação**. Mostra o que existe no
texto e não tem tinta: o ponto de cada espaço, a seta de cada tabulação, o `¬` de cada quebra
de linha e o `¶` no fim de cada parágrafo. Serve para descobrir por que um alinhamento saiu
errado — quase sempre é um espaço a mais ou uma tabulação onde devia haver recuo.

As marcas **não saem no papel nem no PDF**, como no Word. E não deslocam a paginação: elas são
desenhadas sem ocupar lugar na linha, de propósito.

### Autocorreção tipográfica

Ligada por padrão. Enquanto você digita, ela troca:

| O que você digita | O que sai |
| --- | --- |
| `"aspas"` | `“aspas”` |
| `'aspas'` | `‘aspas’` |
| `--` | `—` (travessão) |
| `...` | `…` |
| `(c)` `(r)` `(tm)` | `©` `®` `™` |
| `1/2` `1/4` `3/4` | `½` `¼` `¾` |
| `+/-` `!=` `<<` `>>` | `±` `≠` `«` `»` |

É o que o Word faz em português. `->` e `3 x 4` **não** são trocados por flecha e por `×`, de
propósito: em texto técnico isso atrapalha mais do que ajuda.

`Backspace` logo depois de uma troca desfaz **só ela**, como no Word: escreva `--silent`, e o
travessão volta a ser dois hifens sem que o resto do texto mude. É a saída para escrever um
comando ou um pedaço de código sem desligar nada.

Para desligar: **Ferramentas → Autocorreção tipográfica**. O que já foi trocado continua
trocado — desligar vale para o que vem depois.

---

<a id="tabelas-e-imagens"></a>

## 📊 Tabelas e imagens

**Inserir tabela** fica no menu **Tabela**, no botão de tabela da barra e em `Ctrl+F12`, e
pergunta quantas linhas e colunas. A linha de cabeçalho vem marcada: no arquivo ela é a linha
que o Word **repete no alto de cada página** quando a tabela quebra.

Com o cursor numa célula, o menu **Tabela** e o **botão direito** oferecem:

- inserir linha acima ou abaixo, coluna à esquerda ou à direita;
- excluir linha, coluna ou a tabela inteira;
- **mesclar células** (selecione-as arrastando o mouse) e **dividir célula**;
- ligar e desligar a **linha de cabeçalho**;
- **Propriedades da tabela…**: a largura da coluna do cursor em milímetros, a borda das células
  selecionadas (estilo, espessura, cor e quais lados), o sombreamento e a repetição do
  cabeçalho.

A divisória entre duas colunas também se **arrasta** com o mouse. A largura vai para o arquivo
nos dois casos.

> ℹ️ O diálogo só oferece o que o `.docx` guarda. Estilos de borda que o Word tem e a tela não
> desenha — linha grossa e fina, ondulada — aparecem como linha simples, e **voltam intactos**
> ao salvar enquanto você não formatar aquela célula. Se formatar, o aviso de perda diz isso.

**Redimensionar imagem**: clique nela e arraste uma das oito alças. Nos cantos a proporção
fica travada; segure `Shift` para esticar livremente. As alças das bordas mexem numa medida só.
A imagem não passa da largura da coluna de texto, e um `Ctrl+Z` desfaz o arrasto inteiro. Com
uma alça em foco, as setas do teclado também redimensionam.

**Texto alternativo e alinhamento**: com a imagem selecionada, o botão de imagem da barra passa a
se chamar **Propriedades da imagem** (também em **Formatar → Imagem…**). O texto alternativo é
o que um leitor de tela lê no lugar da imagem, e vai para o arquivo.

**Mesclagem vertical**: células mescladas de cima para baixo num `.docx` continuam mescladas no
arquivo, mas a tela as mostra separadas — e inserir ou excluir linha numa tabela assim desfaz a
mesclagem, com aviso de perda. Mesclar **na vertical** dentro do Librevia funciona na tela, mas
ainda não vai para o `.docx`: ao salvar, o aviso de perda diz isso. Mesclar na horizontal vai.

---

<a id="fórmulas-em-português"></a>

## 🧮 Fórmulas em português

```excel
=SOMA(1,5;2)
```

Isso soma um e meio com dois. Duas regras andam juntas, e são as mesmas do Excel em português:

- **vírgula é o separador decimal**;
- **ponto e vírgula separa os argumentos** — justamente porque a vírgula já está ocupada.

Se você digitar `=SOMA(A1,B1)`, o aplicativo diz qual é o separador certo, em vez de dar um
erro genérico.

Os nomes funcionam nos **dois idiomas**: `SOMA` e `SUM` são a mesma função, assim como
`SE`/`IF` e `PROCV`/`VLOOKUP`. Quem colou uma fórmula de uma planilha estrangeira não precisa
traduzir na mão.

<a id="funções-disponíveis"></a>

### Funções disponíveis

São 52, em seis grupos:

| Grupo | Funções |
| --- | --- |
| **Contas** | `SOMA` `SOMASE` `ARRED` `ARREDONDAR.PARA.CIMA` `ARREDONDAR.PARA.BAIXO` `ABS` `INT` `TRUNCAR` `RESTO` `RAIZ` `POTÊNCIA` |
| **Estatística** | `MÉDIA` `MÁXIMO` `MÍNIMO` `CONT.NÚM` `CONT.VALORES` `CONTAR.VAZIO` `CONT.SE` |
| **Lógica** | `SE` `SEERRO` `SENÃODISP` `E` `OU` `NÃO` `ÉERROS` `É.NÃO.DISP` `ÉNÚM` `ÉTEXTO` `ÉCÉL.VAZIA` |
| **Texto** | `CONCATENAR` `NÚM.CARACT` `ESQUERDA` `DIREITA` `EXT.TEXTO` `MAIÚSCULA` `MINÚSCULA` `ARRUMAR` `SUBSTITUIR` `PROCURAR` `LOCALIZAR` `VALOR` |
| **Procura** | `PROCV` `PROCH` `CORRESP` `ÍNDICE` |
| **Data** | `HOJE` `AGORA` `DATA` `ANO` `MÊS` `DIA` `DIA.DA.SEMANA` |

### O que ainda não existe

Matrizes dinâmicas, referências de coluna inteira (`A:A`) e intervalos nomeados.

> ✅ Uma fórmula que use função fora da lista mostra `#NOME?` na célula — mas **continua no
> arquivo, intacta**, e o Excel volta a calculá-la normalmente. O aviso na abertura diz quais
> funções são essas.

---

<a id="arquivos-que-abrem-travados"></a>

## 🔒 Alguns arquivos abrem travados — e como destravar

Documento com **comentário, controle de alterações, nota de rodapé ou campo calculado** abre
somente para leitura, com uma faixa laranja no topo dizendo exatamente o que ele tem.

O motivo é concreto: a gravação preserva tudo isso **desde que você não edite o trecho que os
ancora**. Quem só precisa ler não corre risco nenhum. Quem precisa editar clica em
**Editar mesmo assim** e segue — sabendo qual é o risco.

Não é cadeado, é padrão: um clique libera, e vale só para aquele arquivo. Arquivo comum abre
editável, porque travar tudo ensinaria você a clicar sem ler — e aí a proteção deixaria de
proteger.

---

<a id="os-avisos"></a>

## ⚠️ Os avisos: leia, eles são diferentes entre si

O aplicativo distingue duas coisas que a maioria dos programas mistura:

| | O que significa | O que você pode fazer |
| --- | --- | --- |
| **Invisibilidade** | o recurso continua no arquivo, mas não aparece aqui — comentários, controle de alterações, gráficos, filtros, formatação condicional | editar e salvar à vontade: eles voltam intactos |
| **Perda** | some de verdade ao salvar | só acontece se você editar justamente o trecho que ancorava o recurso |

São avisos separados porque exigem reações diferentes. Um alerta genérico é um alerta que se
aprende a fechar sem ler.

---

<a id="se-o-aplicativo-fechar-sozinho"></a>

## 💾 Se o aplicativo fechar sozinho

De oito em oito segundos, o que está na tela é guardado num rascunho — **nunca por cima do seu
arquivo**. Se houver uma queda, na próxima abertura aparece uma faixa azul oferecendo o
trabalho de volta.

- **Recuperar** traz o conteúdo para a tela. Ele fica marcado como *não salvo*, porque é isso
  mesmo que ele é. Confira e salve onde quiser.
- **Descartar** apaga o rascunho de vez.

Enquanto a faixa estiver na tela, o rascunho **não é sobrescrito**. Você pode ignorá-la, abrir
outro arquivo, e o trabalho continua lá.

---

<a id="salvar-não-custa"></a>

## 🛟 Salvar não custa o que você não editou

Ao gravar um `.docx` ou `.xlsx`, o aplicativo **não regenera o arquivo**: ele reescreve só o
que você mexeu e devolve o resto exatamente como estava.

| Medido em | Resultado |
| --- | --- |
| Documento de 105 blocos, salvo sem editar | **zero** blocos reescritos |
| O mesmo, com um parágrafo editado | **um** bloco reescrito |
| Planilha do LibreOffice, aberta e salva sem editar | **zero** células escritas |

Na prática: fonte, alinhamento, bordas, gráficos, tabelas dinâmicas, comentários e filtros
continuam no arquivo depois de você corrigir uma vírgula.

---

<a id="imprimir-e-exportar-pdf"></a>

## 🖨️ Imprimir e exportar PDF

O que sai no papel é **o que está na tela** — o PDF é gerado pelo mesmo motor que desenha o
editor, com o texto selecionável e as fontes embutidas.

- **Documentos** saem com cabeçalho, rodapé, numeração, margens e orientação do arquivo.
- **Planilhas** saem com a aba ativa inteira, ajustada à largura da página; as linhas
  congeladas viram cabeçalho e se repetem em cada folha.

> 💡 Se as colunas da planilha ficarem apertadas, ponha a página em paisagem pelo menu
> **Arquivo → Configuração de página**.

O cabeçalho e o rodapé que vieram do arquivo são **editáveis**: clique no texto e digite, e o
que você escrever volta para o lugar exato de onde saiu. O que não tem texto próprio no
arquivo continua sendo só desenho — o logotipo, a moldura da tabela e o número da página, que
é recalculado a cada abertura.

---

<a id="limites-conhecidos"></a>

## 🚧 Limites conhecidos

| Limite | O que fazer |
| --- | --- |
| **Não cria `.docx` do zero** | documento novo nasce `.sdoc`. Para virar `.docx` seria preciso gerar o pacote inteiro, e aí a promessa acima deixaria de valer. Planilha nova **pode** ser salva direto em `.xlsx` |
| **Mesclagem de células** | ainda não existe na planilha |
| **Filtros de planilha** | são preservados no arquivo, mas não há tela para criar ou alterar |
| **Arquivos acima de 20 MB** | não abrem |
| **A paginação na tela é estimativa** | é o `≈` na barra de status; a paginação exata é a da exportação para PDF |

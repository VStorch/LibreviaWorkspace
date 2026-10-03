# Librevia — manual de uso

Editor de documentos e planilhas que funciona **sem internet**. Abre e grava arquivos do Word
(`.docx`) e do Excel (`.xlsx`) e, ao salvar, regrava só o que você editou.

> 📖 Este manual é para quem **usa** o aplicativo.
> Quem for mexer no código deve ler o [README](README.md).

---

<a id="o-que-tem-aqui"></a>

## 📑 O que tem aqui

| Começar | Escrever | Cuidar do arquivo |
| --- | --- | --- |
| [💿 Instalar](#instalar) | [✍️ Texto e parágrafo](#formatar) | [🔒 Arquivos que abrem travados](#arquivos-que-abrem-travados) |
| [📄 Arquivos e modelos](#os-arquivos-que-ele-abre) | [📋 Listas](#listas) | [⚠️ Os avisos](#os-avisos) |
| [⌨️ Atalhos](#atalhos) | [📐 Páginas, seções e colunas](#paginas) | [💾 Se o aplicativo fechar sozinho](#se-o-aplicativo-fechar-sozinho) |
| [👁️ Ver o documento](#exibir) | [📊 Tabelas e imagens](#tabelas-e-imagens) | [🛟 Salvar não custa o que você não editou](#salvar-não-custa) |
| | [🔗 Sumário e referências](#referencias) | [🖨️ Imprimir e exportar](#imprimir-e-exportar-pdf) |
| | [📝 Notas de rodapé e de fim](#notas) | [🚧 Limites conhecidos](#limites-conhecidos) |
| | [➗ Equações](#equacoes) | |
| | [💬 Comentários e revisões](#revisao) | |
| | [🔤 Ortografia, símbolos e contagem](#ferramentas) | |
| | [🧮 Fórmulas da planilha](#fórmulas-em-português) | |

---

<a id="instalar"></a>

## 💿 Instalar

Nada aqui precisa de internet. **O aplicativo não acessa a rede em momento nenhum**, nem para
conferir a ortografia.

### 🐧 Linux — AppImage

Baixe o arquivo, dê permissão de execução e abra:

```bash
chmod +x Librevia-0.0.0.AppImage
./Librevia-0.0.0.AppImage
```

### 🐧 Linux — pacote `.deb`

Instala no menu de aplicativos, como qualquer outro programa:

```bash
sudo apt install ./librevia_0.0.0_amd64.deb
```

### 🪟 Windows

Execute o instalador. Ele instala **só para o seu usuário** e não pede senha de administrador.
Depois disso, um `.docx` ou um `.dotx` pode ser aberto com o Librevia pelo botão direito do
Explorador de Arquivos.

---

<a id="os-arquivos-que-ele-abre"></a>

## 📄 Os arquivos que ele abre

| Extensão | O que é | O que o Librevia faz com ele |
| --- | --- | --- |
| `.docx` | documento do Word | abre e grava; um documento novo também pode ser salvo assim |
| `.dotx` | modelo do Word | abrir cria um documento novo a partir dele; "Salvar como" também grava modelo |
| `.dotm` | modelo do Word com macros | abre como o `.dotx`, mas as macros ficam de fora, e o aplicativo avisa |
| `.xlsx` | planilha do Excel | abre e grava |
| `.sdoc` | documento do Librevia | guarda tudo, sem perda nenhuma |
| `.ssheet` | planilha do Librevia | guarda tudo, sem perda nenhuma |
| `.txt` | texto puro | abre e grava; salvar nele descarta a formatação, e o aplicativo avisa antes |

E estes, o Librevia só **cria**:

| Extensão | Como |
| --- | --- |
| `.pdf` | **Arquivo → Exportar para PDF…**, para documentos e planilhas |
| `.html`, `.md`, `.odt` | **Arquivo → Exportar como**, só para documentos |

> ℹ️ **`.odt` e `.ods` não abrem.** Se você recebe arquivos assim, peça a quem enviou que salve
> como `.docx` ou `.xlsx`. O LibreOffice faz isso pelo "Salvar como".

### Modelos

Um modelo é um documento pronto para servir de ponto de partida, como um papel timbrado.

**Arquivo → Novo a partir de modelo…** mostra os modelos que vêm com o aplicativo — documento em
branco, carta, relatório com capa e sumário, e ata de reunião — e também os seus, que ficam na
pasta de modelos. O botão **Abrir pasta de modelos** mostra essa pasta, e **Procurar…** aceita
qualquer `.dotx` do computador.

O documento criado é novo e sem título. Estilos, cabeçalho, rodapé e tamanho de página vêm do
modelo. Ao salvar, o aplicativo pergunta onde gravar, e **o modelo nunca é sobrescrito**.

Para criar um modelo seu, use **Salvar como → Modelo do Word (.dotx)** e grave na pasta de
modelos.

### Propriedades do documento

**Arquivo → Propriedades…** tem duas partes:

- **Resumo**: título, assunto, autores, gerente, empresa, categoria, palavras-chave e
  comentários. Você pode preencher e alterar.
- **Informações e estatísticas**: quando o documento foi criado e modificado, por quem, o número
  da revisão e o tempo total de edição.

---

<a id="atalhos"></a>

## ⌨️ Atalhos

São os atalhos do Word, com poucas exceções, explicadas no fim da tabela.

### Arquivo e edição

| Atalho | O que faz |
| --- | --- |
| `Ctrl+N` | novo documento |
| `Ctrl+Shift+N` | nova planilha |
| `Ctrl+O` | abrir |
| `Ctrl+S` | salvar |
| `Ctrl+Shift+S` | salvar como |
| `Ctrl+W` | fechar o arquivo |
| `Ctrl+P` | imprimir |
| `Ctrl+F` | localizar e substituir |
| `Ctrl+Shift+V` | colar sem formatação |
| `Ctrl+Shift+G` | contar palavras |

### Texto e parágrafo

| Atalho | O que faz |
| --- | --- |
| `Ctrl+B` / `Ctrl+I` / `Ctrl+U` | negrito, itálico, sublinhado |
| `Ctrl+Shift+=` / `Ctrl+=` | sobrescrito e subscrito (também `Ctrl+.` e `Ctrl+,`) |
| `Ctrl+L` / `Ctrl+E` / `Ctrl+R` / `Ctrl+J` | alinhar à esquerda, centralizar, à direita, justificar |
| `Ctrl+1` / `Ctrl+5` / `Ctrl+2` | entrelinha simples, 1,5 e dupla |
| `Ctrl+Alt+1` … `Ctrl+Alt+6` | Título 1 a Título 6 |
| `Ctrl+Alt+0` | volta o parágrafo para corpo de texto |
| `Ctrl+]` / `Ctrl+[` | aumentar e diminuir o recuo |
| `Ctrl+Shift+8` / `Ctrl+Shift+7` | lista com marcadores e lista numerada |
| `Tab` / `Shift+Tab` | numa lista, desce o item um nível ou o devolve |

### Inserir

| Atalho | O que faz |
| --- | --- |
| `Ctrl+Enter` | quebra de página |
| `Ctrl+F12` | tabela (pergunta linhas e colunas) |
| `Ctrl+Alt+F` | nota de rodapé |
| `Ctrl+Alt+D` | nota de fim |
| `Ctrl+Alt+M` | comentário |
| `Alt+=` | equação |
| `Ctrl+Shift+F5` | marcador |
| `F9` | atualizar campos (números de página, referências) |

### Exibir e revisar

| Atalho | O que faz |
| --- | --- |
| `Ctrl+F5` | painel de navegação |
| `Ctrl+F10` | mostrar e ocultar as marcas de formatação |
| `Ctrl+F11` | modo de leitura (`Esc` sai) |
| `Ctrl+Shift+E` | ligar e desligar o controle de alterações |
| `Ctrl` + `+` do teclado numérico / `Ctrl+-` | ampliar e reduzir |
| `Ctrl+0` | voltar ao tamanho normal |

> ℹ️ **As exceções ao Word, e por quê.**
>
> - **Corpo de texto** é `Ctrl+Alt+0`, e não `Ctrl+Shift+N`, porque aqui essa combinação abre
>   uma planilha nova.
> - **Ampliar** usa o `+` do teclado numérico, para não ocupar o `Ctrl+Shift+=` do sobrescrito.
>   O menu **Exibir** também tem ampliar e reduzir.
> - **Marcas de formatação** são `Ctrl+F10`, como no LibreOffice. O `Ctrl+*` do Word já é usado
>   pela lista com marcadores.

### Na planilha

`Ctrl+B`, `Ctrl+I` e `Ctrl+U` valem para as células selecionadas. O **botão direito** numa célula
abre o menu de inserir e excluir linhas e colunas.

**A alça de preenchimento** é o quadradinho no canto inferior direito da seleção. Arraste-o, e as
células seguintes são preenchidas. Se a célula de origem tiver uma fórmula, ela **acompanha**:
`=B2*C2` arrastada para baixo vira `=B3*C3`. Uma referência com `$` não se mexe.

---

<a id="formatar"></a>

## ✍️ Texto e parágrafo

A barra do documento tem, da esquerda para a direita: estilo do parágrafo, fonte, tamanho, a
formatação da letra, o parágrafo, as listas, o que se insere e a página.

### A letra

Além de negrito, itálico, sublinhado e tachado:

| Botão | O que faz |
| --- | --- |
| **Sobrescrito** e **Subscrito** | põem o trecho acima ou abaixo da linha: `m³`, `H₂O` |
| **Caixa alta** | mostra em maiúsculas sem trocar o que você digitou; tirar a formatação devolve o texto como era |
| **Versalete** | maiúsculas pequenas, da altura das minúsculas |
| **Cor do texto** | a cor da letra |
| **Cor de fundo do texto** | pinta o fundo do trecho, com qualquer cor |
| **Destaque** | o marca-texto do Word, que tem catorze cores fixas |

> ℹ️ **Por que dois tipos de fundo?** No arquivo do Word são duas coisas diferentes: o
> marca-texto e o sombreamento do trecho. Um documento recebido pode ter os dois, e o aplicativo
> mantém cada um onde estava.

### Os estilos

Um estilo é uma formatação com nome. Em vez de deixar cada título em negrito, 14 pt e azul, um
por um, você aplica o estilo **Título 1**. Se depois decidir que os títulos devem ser verdes,
muda o estilo uma vez e todos os títulos mudam juntos.

Num `.docx` feito no Word, quase toda a formatação mora nos estilos. É por isso que um título pode
aparecer como uma faixa colorida, e não só como letra grande.

O botão **A≡**, ao lado do tamanho da fonte, abre o painel de **Estilos**. Ele mostra os estilos
que **este documento** define e diz qual é o do parágrafo onde está o cursor. Ali você pode:

- **Aplicar** um estilo ao parágrafo ou ao trecho selecionado;
- **Modificar…** um estilo: nome, estilo em que ele se baseia, estilo do parágrafo seguinte,
  fonte e tamanho. **Atualizar a partir da seleção** copia para o estilo a formatação do texto
  selecionado;
- criar um **Novo estilo…**, de parágrafo ou de caractere;
- **Limpar formatação**, que tira a formatação direta e deixa só a do estilo.

A lista pode mostrar todos os estilos, só os de parágrafo ou só os de caractere.

Documentos novos começam em **Calibri 11**, como no Word.

### A lista de fontes

O seletor mostra, nesta ordem:

1. **as fontes que o documento aberto usa**;
2. as cinco que vêm com o aplicativo: Calibri, Cambria, Arial, Times New Roman e Courier New;
3. **as instaladas no seu computador**, em ordem alfabética.

As cinco do meio funcionam igual em qualquer computador, porque vão dentro do instalador. As
suas funcionam no seu. Num computador que não as tenha, o documento aparece com outra fonte, e as
páginas podem quebrar em lugares diferentes.

### O diálogo de parágrafo

O botão **¶**, ou **Formatar → Parágrafo…**, reúne numa tela tudo o que é do parágrafo, já
preenchido com o que ele tem:

- **alinhamento**: esquerda, centro, direita, justificado;
- **entrelinha**: simples, 1,15, 1,5, dupla, um múltiplo qualquer ou uma medida em pontos;
- **espaçamento** antes e depois, em pontos;
- **recuo** à esquerda e à direita, em milímetros;
- **primeira linha**: recuo (a primeira linha entra) ou deslocamento (ela sai, e as outras
  entram, como numa referência bibliográfica);
- **manter com o próximo**: o parágrafo não fica sozinho no pé da página, longe do que vem
  depois.

Vale para todos os parágrafos que a seleção tocar. `Enter` aplica; `Esc` fecha sem aplicar.

> ℹ️ **"Simples" não é "1,0".** A entrelinha simples é a altura que a própria fonte pede. Nenhum
> número a reproduz exatamente, e trocar uma pela outra muda onde as páginas quebram.

---

<a id="listas"></a>

## 📋 Listas

`Ctrl+Shift+8` começa uma lista com marcadores; `Ctrl+Shift+7`, uma lista numerada. Os mesmos
botões estão na barra.

Uma lista pode ter vários níveis, como um índice: `1.`, depois `1.1.`, depois `1.1.1.`. **Tab**
desce o item um nível, e **Shift+Tab** o devolve. A numeração segue as mesmas regras do Word, então
um documento recebido mostra os mesmos números que mostrava lá.

**Formato de lista…** escolhe a aparência:

- **Listas prontas**: `1. a. i.`, `1. 1.1. 1.1.1.`, `I. A. 1. a.`, `1) a) i)` e três tipos de
  marcadores;
- ou nível por nível: o formato do número (`1, 2, 3`, `01, 02`, `a, b, c`, `A, B, C`, `i, ii`,
  `I, II` ou marcador), o texto em volta dele e o número inicial. A **prévia** mostra o resultado
  antes de aplicar.

O **botão direito** num item numerado oferece:

| Opção | Quando usar |
| --- | --- |
| **Reiniciar em 1** | a lista nova começou colada na anterior e herdou a numeração dela |
| **Continuar numeração** | o contrário: um parágrafo comum no meio partiu a lista em duas |
| **Definir valor inicial…** | a lista precisa começar em outro número, como `5.` |

---

<a id="paginas"></a>

## 📐 Páginas, seções e colunas

### A tela mostra as folhas de verdade

O documento aparece dividido em folhas, como vai sair no papel. A quebra é feita **linha por
linha**, do mesmo jeito que o Word e o LibreOffice fazem:

- um parágrafo pode começar numa folha e terminar na seguinte;
- uma linha não fica sozinha no topo ou no pé da folha (são as "viúvas" e "órfãs");
- a linha de cabeçalho de uma tabela se repete no alto de cada folha em que a tabela continua.

O número de páginas da barra de status é o mesmo do PDF e do papel.

### Configuração de página

**Arquivo → Configuração de página…** define:

- **tamanho** do papel e **orientação** (retrato ou paisagem);
- **margens**, em milímetros;
- **cabeçalho e rodapé**, com os botões **Número da página** e **Total de páginas**;
- **Primeira página diferente**, para uma capa sem número;
- **Pares e ímpares diferentes**, para um livro em que o número fica sempre do lado de fora;
- **numeração de página**: o formato (`1, 2, 3`, `i, ii, iii`, `A, B, C`…) e o número em que começa.

O cabeçalho e o rodapé também se editam direto na folha: clique no texto e digite.

### Seções

Uma seção é um pedaço do documento com a própria configuração de página. Serve, por exemplo, para
pôr uma tabela larga numa folha em paisagem, no meio de um relatório em retrato, ou para numerar a
introdução com `i, ii, iii` e o resto com `1, 2, 3`.

**Inserir → Quebra de seção** oferece quatro tipos:

| Tipo | A seção nova começa… |
| --- | --- |
| **Próxima página** | na folha seguinte |
| **Contínua** | na mesma folha, logo abaixo |
| **Página par** | na próxima folha par |
| **Página ímpar** | na próxima folha ímpar, como um capítulo de livro |

Com seções, a configuração de página pergunta onde **Aplicar**: **Nesta seção** ou **No documento
todo**. O rodapé de uma seção pode **vincular ao anterior**, isto é, repetir o da seção de cima, ou
ter texto próprio.

**Inserir → Excluir quebra de seção** junta a seção com a de cima.

### Colunas

**Formatar → Colunas…** divide a seção em colunas, como num jornal: o **número de colunas**, o
**espaço entre elas**, em milímetros, e uma **linha entre colunas**, se quiser.

**Inserir → Quebra de coluna** manda o texto seguinte para o alto da próxima coluna.

---

<a id="tabelas-e-imagens"></a>

## 📊 Tabelas e imagens

**Inserir tabela** fica no menu **Tabela**, no botão de tabela da barra e em `Ctrl+F12`, e
pergunta quantas linhas e colunas você quer. A primeira linha vem marcada como **cabeçalho**: é a
linha que se repete no alto de cada folha quando a tabela continua na próxima.

Com o cursor numa célula, o menu **Tabela** e o **botão direito** oferecem:

- inserir linha acima ou abaixo, coluna à esquerda ou à direita;
- excluir linha, coluna ou a tabela inteira;
- **mesclar células** (selecione-as arrastando o mouse) e **dividir célula**;
- ligar e desligar a **linha de cabeçalho**;
- **Propriedades da tabela…**: a largura da coluna do cursor, em milímetros; a borda das células
  selecionadas (estilo, espessura, cor e quais lados); o sombreamento; e a repetição do
  cabeçalho.

A divisória entre duas colunas também pode ser **arrastada** com o mouse, em qualquer zoom.

> ℹ️ O diálogo só oferece o que o `.docx` guarda. Bordas que o Word tem e a tela não desenha,
> como linha dupla grossa e fina ou ondulada, aparecem como linha simples, e **voltam intactas**
> ao salvar, desde que você não formate aquela célula. Se formatar, o aviso de perda diz isso.

**Redimensionar imagem**: clique nela e arraste uma das oito alças. Nos cantos, a proporção fica
travada; segure `Shift` para esticar livremente. As alças dos lados mudam uma medida só. A imagem
não passa da largura da coluna de texto, e um `Ctrl+Z` desfaz o arrasto inteiro. Com uma alça
selecionada, as setas do teclado também redimensionam.

**Texto alternativo e alinhamento**: com a imagem selecionada, o botão de imagem da barra passa a
se chamar **Propriedades da imagem** (também em **Formatar → Imagem…**). O texto alternativo é o
que um leitor de tela lê no lugar da imagem, e vai para o arquivo.

**Mesclagem vertical**: células mescladas de cima para baixo num `.docx` continuam mescladas no
arquivo. Mesclar **na vertical** dentro do Librevia funciona na tela, mas ainda não vai para o
`.docx`; ao salvar, o aviso de perda diz isso. Mesclar na horizontal vai normalmente.

---

<a id="referencias"></a>

## 🔗 Sumário e referências

### Painel de navegação

**Exibir → Painel de navegação** (`Ctrl+F5`) mostra os títulos do documento numa coluna ao lado,
como um índice. Clique num título para ir até ele. Num documento de cem páginas, é o jeito mais
rápido de achar um capítulo.

### Sumário

**Referências → Inserir sumário** monta o sumário a partir dos parágrafos com estilo **Título 1**,
**Título 2** e assim por diante, com o número da página de cada um.

O sumário não muda sozinho enquanto você escreve. Depois de acrescentar capítulos ou de as páginas
mudarem, use **Referências → Atualizar sumário**.

### Marcadores

Um marcador é um nome dado a um ponto do documento, como um marcador de página num livro.
**Inserir → Marcador…** (`Ctrl+Shift+F5`) cria, lista, leva até um marcador ou o exclui. A lista
pode ser ordenada por nome ou pela posição no documento.

O diálogo de **link** também aponta para dentro do documento: em vez de um endereço da internet,
escolha um título ou um marcador.

### Legendas

**Referências → Inserir legenda…** numera figuras, tabelas e equações: `Figura 1`, `Figura 2`…
Escolha o rótulo, o texto depois do número e se a legenda fica acima ou abaixo do item onde está o
cursor.

### Referência cruzada

Uma referência cruzada é uma frase como "veja a Figura 3, na página 12" em que o **3** e o **12**
se atualizam sozinhos. **Referências → Referência cruzada…** pergunta:

- o **tipo**: título, marcador, nota de rodapé, nota de fim ou legenda;
- **para qual** deles;
- o que inserir: o **texto**, o **número**, o **número da nota** ou o **número da página**;
- e se a referência deve funcionar como **hiperlink**, levando até o ponto ao ser clicada.

### Atualizar campos

Números de página, total de páginas, referências cruzadas e legendas são **campos**: texto que o
aplicativo calcula. `F9`, ou **Referências → Atualizar campos**, recalcula todos.

---

<a id="notas"></a>

## 📝 Notas de rodapé e de fim

Uma **nota de rodapé** fica no pé da mesma folha; uma **nota de fim** fica no fim do documento.

- **Inserir → Nota de rodapé** (`Ctrl+Alt+F`)
- **Inserir → Nota de fim** (`Ctrl+Alt+D`)

O número aparece no texto, e o cursor vai para a nota, **no pé da própria folha**, onde você
escreve. Os números das notas seguem a ordem em que elas aparecem no texto.

O **botão direito** numa nota permite **convertê-la** de rodapé para fim, ou o contrário.

---

<a id="equacoes"></a>

## ➗ Equações

As equações de um documento do Word aparecem desenhadas na folha. As novas são escritas em
**LaTeX**, a forma de escrever matemática usada em artigos científicos: `\frac{a}{b}` é uma fração,
`x^2` é um quadrado, `\sqrt{x}` é uma raiz.

- **Inserir → Equação** (`Alt+=`) põe a equação no meio da linha, junto do texto.
- **Inserir → Equação em destaque** põe a equação numa linha só dela, centralizada.

O diálogo tem o campo de **LaTeX**, uma **Visualização** que acompanha o que você digita e uma
paleta de **Modelos** para quem não sabe os comandos de cor: frações, raízes, índices, somatórios,
integrais, acentos, funções como seno e logaritmo, e símbolos como ∈, ∀ e ⇒.

Para alterar uma equação, **clique duas vezes** nela.

> ℹ️ Uma equação do Word com alguma construção que a tela não sabe desenhar aparece **travada**,
> com o que deu para desenhar. Ela volta ao arquivo inteira, exatamente como veio.

---

<a id="revisao"></a>

## 💬 Comentários e revisões

### Comentários

Selecione o trecho e use **Inserir → Comentário** (`Ctrl+Alt+M`). O comentário aparece no painel
ao lado, ligado ao trecho. **Exibir → Comentários** mostra e esconde esse painel.

Em cada comentário você pode **Responder**, **Editar**, **Resolver** (ele fica marcado como
resolvido, mas continua no arquivo), **Reabrir** e **Excluir**. **Inserir → Próximo comentário** e
**Comentário anterior** andam de um a outro.

O nome que aparece nos seus comentários e revisões é definido em **Ferramentas → Nome do autor…**.

Comentários de um `.docx` vêm junto e voltam ao arquivo, inclusive as respostas.

### Controle de alterações

Com o controle de alterações ligado, cada mudança fica registrada em vez de simplesmente feita:
texto inserido aparece sublinhado, texto apagado aparece riscado, cada um com o nome de quem mudou.
Quem recebe o documento decide o que fica.

**Revisão → Controlar alterações** (`Ctrl+Shift+E`) liga e desliga. O menu **Revisão** também tem:

- **Aceitar alteração** e **Rejeitar alteração**, para a que está no cursor;
- **Aceitar todas** e **Rejeitar todas**;
- **Próxima alteração** e **Alteração anterior**.

**Revisão → Mostrar** escolhe como ver o documento:

| Modo | O que aparece |
| --- | --- |
| **Marcação completa** | tudo à vista: o inserido sublinhado e o excluído riscado |
| **Marcação simples** | o texto como vai ficar, com uma barra na margem ao lado de cada parágrafo alterado |
| **Sem marcação** | o texto como vai ficar, limpo |
| **Original** | o texto de antes das alterações; nesse modo o editor fica só para leitura |

Trocar de modo não altera o documento, e a impressão segue o modo escolhido.

---

<a id="ferramentas"></a>

## 🔤 Ortografia, símbolos e contagem

### A ortografia é em português e não usa internet

O dicionário de português do Brasil **vem dentro do instalador**. Na primeira vez que o aplicativo
abre, ele é copiado para a pasta do seu usuário, e é ali que o corretor o procura.

Por que isso importa? O corretor que acompanha o Chromium, sozinho, baixaria o dicionário da
internet. Num computador sem rede, ele simplesmente não marcaria nada, e não avisaria.

A palavra errada ganha o sublinhado ondulado de sempre, **no texto e também no cabeçalho e no
rodapé**, onde um erro de digitação se repete em todas as folhas.

Clique com o **botão direito** na palavra sublinhada para:

- escolher uma das sugestões, que troca a palavra;
- **Adicionar ao dicionário**, para nunca mais ser avisado dela, em nenhum documento;
- **Ignorar nesta sessão**, que vale até você fechar o aplicativo. Serve para o nome de um cliente
  que aparece dez vezes neste documento e em nenhum outro.

Para desligar: **Ferramentas → Verificação ortográfica**.

### O botão direito no documento

Além das sugestões de ortografia, o menu tem **recortar**, **copiar**, **colar** e **colar sem
formatação**. Dentro de uma tabela, traz também as ações dela. Um item apagado não se aplica ali:
"colar" fica apagado quando não há nada copiado, por exemplo.

### Colar sem formatação

`Ctrl+Shift+V`, o menu **Editar** ou o botão direito. O texto entra com a formatação do documento,
e não com a de onde veio. É o que evita que um trecho copiado de um site chegue com outra fonte,
outro tamanho e outra cor. Cada quebra de linha do texto copiado vira um parágrafo.

### Contar palavras

`Ctrl+Shift+G` ou **Ferramentas → Contar palavras…** mostra palavras, caracteres com e sem
espaço, parágrafos e páginas, em duas colunas: o **documento** e a **seleção**. Sem nada
selecionado, a coluna da seleção mostra um travessão, porque "nada selecionado" não é o mesmo que
"zero palavras". O diálogo acompanha o que você digita, então pode ficar aberto.

Parágrafo vazio não conta: é um `Enter` batido para abrir espaço, e o Word também não o conta.

### Caracteres especiais

**Inserir → Caractere especial…** abre um painel com o que costuma faltar no teclado: aspas
tipográficas, travessão, símbolos de moeda, sinais de matemática, letras gregas e marcas como © e
®. Clique no símbolo, ou ande pela grade com as setas e insira com `Enter`. O painel fica aberto,
porque quem o abre em geral quer mais de um símbolo.

O `␣` do grupo "Pontuação" é o **espaço inquebrável**. É ele que impede que o `R$` fique no fim de
uma linha e o valor no começo da seguinte.

### Marcas de formatação (¶)

`Ctrl+F10`, o botão `¶` da barra ou **Exibir → Marcas de formatação** mostram o que existe no texto
mas não tem tinta: um ponto em cada espaço, uma seta em cada tabulação, um `¬` em cada quebra de
linha e um `¶` no fim de cada parágrafo. Servem para descobrir por que um alinhamento saiu errado.
Quase sempre é um espaço a mais, ou uma tabulação onde devia haver recuo.

As marcas **não saem no papel nem no PDF**, e não mudam onde as páginas quebram.

### Autocorreção tipográfica

Vem ligada. Enquanto você digita, ela troca:

| O que você digita | O que sai |
| --- | --- |
| `"aspas"` | `“aspas”` |
| `'aspas'` | `‘aspas’` |
| `--` | `—` (travessão) |
| `...` | `…` |
| `(c)` `(r)` `(tm)` | `©` `®` `™` |
| `1/2` `1/4` `3/4` | `½` `¼` `¾` |
| `+/-` `!=` `<<` `>>` | `±` `≠` `«` `»` |

É o que o Word faz em português. `->` e `3 x 4` **não** viram flecha e `×`, de propósito: em texto
técnico isso atrapalha mais do que ajuda.

Um `Backspace` logo depois de uma troca desfaz **só aquela troca**. Escreva `--silent`, apague uma
vez, e o travessão volta a ser dois hifens. É a saída para escrever um comando sem desligar nada.

Para desligar: **Ferramentas → Autocorreção tipográfica**. O que já foi trocado continua trocado.

---

<a id="fórmulas-em-português"></a>

## 🧮 Fórmulas da planilha

```excel
=SOMA(1,5;2)
```

Isso soma um e meio com dois, e dá 3,5. Duas regras andam juntas, e são as mesmas do Excel em
português:

- **a vírgula separa os decimais**;
- **o ponto e vírgula separa os argumentos**, justamente porque a vírgula já está ocupada.

Se você digitar `=SOMA(A1,B1)`, o aplicativo diz qual é o separador certo, em vez de mostrar um
erro qualquer.

Os nomes funcionam nos **dois idiomas**: `SOMA` e `SUM` são a mesma função, assim como `SE` e
`IF`, `PROCV` e `VLOOKUP`. Uma fórmula copiada de uma planilha em inglês não precisa ser traduzida
à mão.

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

> ✅ Uma fórmula com uma função fora da lista mostra `#NOME?` na célula, mas **continua no
> arquivo, intacta**, e o Excel volta a calculá-la normalmente. O aviso na abertura diz quais são
> essas funções.

---

<a id="exibir"></a>

## 👁️ Ver o documento

| No menu Exibir | O que faz |
| --- | --- |
| **Ampliar**, **Reduzir**, **Tamanho normal** | zoom de 50 % a 200 %; os botões da barra de status fazem o mesmo |
| **Ajustar à largura** | a folha ocupa a largura da janela |
| **Modo de leitura** (`Ctrl+F11`) | esconde as barras e deixa só o texto; o documento fica só para leitura enquanto durar, e `Esc` sai |
| **Tema** | **Claro**, **Escuro** ou **Do sistema**, que acompanha o que o computador usa |
| **Idioma** | **Português** ou **English**, para menus, botões e mensagens |
| **Mostrar barra de ferramentas**, **Mostrar barra de status** | escondem as barras para ganhar espaço |
| **Tela cheia** | ocupa a tela inteira |

O zoom muda só o tamanho na tela. As páginas quebram nos mesmos lugares em qualquer zoom.

---

<a id="arquivos-que-abrem-travados"></a>

## 🔒 Alguns arquivos abrem travados — e como destravar

Alguns documentos abrem **somente para leitura**, com uma faixa laranja no topo dizendo
exatamente o que eles têm. Acontece quando o arquivo tem algo que o editor mostra, mas não sabe
escrever de volta caso você edite o trecho:

- um **campo** de um tipo que o editor não sabe recalcular. O texto que o Word guardou aparece,
  mas reescrever o parágrafo o transformaria em texto comum;
- **controles de conteúdo**, as caixinhas de formulário do Word;
- **revisões de estrutura**, como uma célula inserida ou excluída numa tabela com o controle de
  alterações ligado.

Quem só precisa ler não corre risco nenhum. Quem precisa editar clica em **Editar mesmo assim** e
segue, sabendo qual é o risco.

Não é um cadeado: um clique libera, e vale só para aquele arquivo. Um arquivo comum abre
editável, mesmo que tenha comentários, revisões ou notas, porque o Librevia sabe gravá-los de
volta. Travar todo arquivo ensinaria você a clicar sem ler, e a proteção deixaria de proteger.

---

<a id="os-avisos"></a>

## ⚠️ Os avisos: leia, eles são diferentes entre si

Ao abrir um arquivo, o aplicativo pode mostrar um aviso. São três tipos, porque cada um pede uma
atitude diferente:

| Aviso | O que significa | O que você pode fazer |
| --- | --- | --- |
| **Invisível** | o recurso continua no arquivo, mas não aparece aqui: gráficos, a sombra ou o degradê de uma forma, formatação condicional e filtros da planilha | editar e salvar à vontade: ele volta intacto |
| **Perda** | algo some de verdade ao salvar | acontece só se você editar justamente o trecho onde ele está; o aviso diz qual é |
| **Trava** | é a perda que fez o arquivo abrir [somente para leitura](#arquivos-que-abrem-travados) | ler à vontade; para editar, clique em **Editar mesmo assim** |

Se todo aviso dissesse a mesma coisa, você aprenderia a fechá-los sem ler.

---

<a id="se-o-aplicativo-fechar-sozinho"></a>

## 💾 Se o aplicativo fechar sozinho

A cada oito segundos, o que está na tela é guardado num rascunho, **nunca por cima do seu
arquivo**. Se o computador desligar ou o aplicativo fechar de repente, na próxima abertura aparece
uma faixa azul oferecendo o trabalho de volta:

- **Recuperar** traz o conteúdo para a tela, marcado como *não salvo*, porque é isso que ele é.
  Confira e salve onde quiser.
- **Descartar** apaga o rascunho.

Enquanto a faixa estiver na tela, o rascunho **não é sobrescrito**. Você pode ignorá-la por um
tempo, abrir outro arquivo, e o trabalho continua guardado.

Quando você salva, o rascunho é apagado: o arquivo no disco já é a versão boa.

---

<a id="salvar-não-custa"></a>

## 🛟 Salvar não custa o que você não editou

Muitos editores, ao salvar um `.docx`, escrevem o arquivo inteiro de novo, e no caminho perdem o
que não entenderam. O Librevia faz diferente: **reescreve só o que você mexeu** e devolve o resto
exatamente como estava.

| Medido em | Resultado |
| --- | --- |
| Documento de 105 blocos, salvo sem editar | **nenhum** bloco reescrito |
| O mesmo, com um parágrafo editado | **um** bloco reescrito |
| Planilha do LibreOffice, aberta e salva sem editar | **nenhuma** célula escrita |

Na prática: fonte, alinhamento, bordas, gráficos, tabelas dinâmicas, comentários e filtros
continuam no arquivo depois que você corrige uma vírgula.

**Documento novo em `.docx`.** Em "Salvar como", escolha "Documento do Word". Como não há arquivo
de origem, o aplicativo cria um pacote mínimo, com estilos, página e propriedades, mas sem o seu
nome dentro, e grava o documento nele. Cada gravação parte do mesmo pacote mínimo, e não do arquivo
gravado da última vez. Assim, salvar dez vezes dá o mesmo arquivo que salvar uma.

---

<a id="imprimir-e-exportar-pdf"></a>

## 🖨️ Imprimir e exportar

### Papel e PDF

O que sai no papel é **o que está na tela**. O PDF é gerado pelo mesmo motor que desenha o editor,
com o texto selecionável e as fontes embutidas.

- **Documentos** saem com cabeçalho, rodapé, numeração, margens, orientação e colunas de cada
  seção.
- **Planilhas** saem com a aba ativa inteira, ajustada à largura da página. As linhas congeladas
  viram cabeçalho e se repetem em cada folha.

**Arquivo → Visualizar impressão** mostra as folhas antes de gastar papel.

> 💡 Se as colunas da planilha ficarem apertadas, ponha a página em paisagem em
> **Arquivo → Configuração de página…**.

### Exportar como HTML, Markdown ou ODT

Em **Arquivo → Exportar como**, o documento vira uma **página da web** (`.html`), um texto em
**Markdown** (`.md`) ou um **documento do OpenDocument** (`.odt`). É um arquivo novo: o documento
continua aberto no mesmo lugar, do jeito que estava.

- **HTML** sai numa página só, com os estilos, as listas, as tabelas e as imagens dentro dela. As
  notas vão para o fim, com um link de volta ao texto.
- **Markdown** leva títulos, ênfase, links, listas, tabelas e notas (`[^1]`). As imagens vão para
  uma pasta ao lado, `nome_arquivos/`. Uma tabela com células mescladas sai em HTML, porque o
  Markdown não tem como mesclar.
- **ODT** abre no LibreOffice e em outros editores que leem OpenDocument, com os estilos, as listas
  numeradas, as tabelas, as imagens, as notas, os marcadores e os links. Cada seção leva o próprio
  papel, margens, colunas, cabeçalho e rodapé, e os comentários vão como anotações.
- As **revisões** saem já aceitas. As **equações** vão junto. No HTML e no Markdown os
  **comentários** ficam de fora.

---

<a id="limites-conhecidos"></a>

## 🚧 Limites conhecidos

| Limite | O que acontece |
| --- | --- |
| **Arquivos acima de 20 MB** | não abrem |
| **`.odt` e `.ods`** | não abrem; o documento pode ser **exportado** para `.odt` |
| **Macros** | um `.dotm` abre, mas sem as macros |
| **Mesclar células na planilha** | ainda não existe |
| **Filtros de planilha** | são preservados no arquivo, mas não há tela para criar ou alterar |
| **Fórmulas** | sem matrizes dinâmicas, referências de coluna inteira (`A:A`) ou intervalos nomeados |
| **Mesclagem vertical feita no editor** | aparece na tela, mas não vai para o `.docx` (o aviso de perda diz isso) |
| **Comentários** | o texto do comentário é texto simples, sem negrito ou cor |
| **Controle de alterações** | mudanças de **formatação** (pôr em negrito, trocar a fonte) não são registradas como revisão |
| **Equações** | usam a fonte de matemática do computador; alguns comandos de espaçamento do LaTeX, como `\,` e `\quad`, se perdem ao salvar em `.docx` |
| **Sumário** | **Atualizar sumário** refaz todas as entradas, inclusive as que você editou à mão |

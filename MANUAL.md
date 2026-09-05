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
| `Ctrl+B` / `Ctrl+I` / `Ctrl+U` | negrito, itálico, sublinhado |

**Na planilha**, `Ctrl+B`, `Ctrl+I` e `Ctrl+U` valem para as células selecionadas. Clicar com
o botão direito numa célula abre o menu de inserir e remover linhas e colunas.

**A alça de preenchimento** é o quadradinho no canto inferior direito da seleção: arraste-o e
as células seguintes são preenchidas. Se a célula de origem tiver uma fórmula, ela é
**deslocada** — `=B2*C2` arrastada para baixo vira `=B3*C3`. Referência com `$` não se mexe.

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

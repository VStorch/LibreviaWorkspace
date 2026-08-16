# Librevia — manual de uso

Suíte de documentos e planilhas para trabalhar offline. Abre e grava os arquivos
do Office (`.docx` e `.xlsx`) preservando o que você não editou.

Este manual é para quem **usa** o aplicativo. Quem for mexer no código deve ler o
[README](README.md).

---

## Instalar

**Linux — AppImage.** Baixe o arquivo, dê permissão de execução e abra:

```bash
chmod +x Librevia-0.0.0.AppImage
./Librevia-0.0.0.AppImage
```

**Linux — pacote `.deb`.** Instala no menu de aplicativos como qualquer outro
programa:

```bash
sudo apt install ./librevia_0.0.0_amd64.deb
```

**Windows.** Execute o instalador. Ele instala **para o seu usuário**, sem pedir
senha de administrador.

Nada aqui precisa de internet. O aplicativo não acessa a rede em momento nenhum.

---

## Os arquivos que ele abre

| Extensão  | O que é                                                        |
| --------- | -------------------------------------------------------------- |
| `.docx`   | documento do Word — abre e grava                                |
| `.xlsx`   | planilha do Excel — abre e grava                                |
| `.sdoc`   | documento do Librevia — guarda tudo, sem perda nenhuma          |
| `.ssheet` | planilha do Librevia — idem                                     |
| `.txt`    | texto puro; salvar nele descarta formatação, e o app avisa antes |
| `.pdf`    | só saída: exportar e imprimir, tanto documento quanto planilha   |

`.odt` e `.ods` **não** abrem. Se você recebe arquivos assim, peça para quem
enviou salvar como `.docx` ou `.xlsx` — o LibreOffice faz isso pelo menu
"Salvar como".

---

## Atalhos

| Atalho             | O que faz                    |
| ------------------ | ---------------------------- |
| `Ctrl+N`           | novo documento               |
| `Ctrl+Shift+N`     | nova planilha                |
| `Ctrl+O`           | abrir                        |
| `Ctrl+S`           | salvar                       |
| `Ctrl+Shift+S`     | salvar como                  |
| `Ctrl+W`           | fechar o arquivo             |
| `Ctrl+F`           | localizar e substituir       |
| `Ctrl+P`           | imprimir                     |
| `Ctrl+Enter`       | quebra de página manual      |
| `Ctrl+B` / `I` / `U` | negrito, itálico, sublinhado |

Na planilha, `Ctrl+B`, `Ctrl+I` e `Ctrl+U` valem para as células selecionadas.
Clicar com o botão direito numa célula abre o menu de inserir e remover linhas e
colunas.

---

## Planilha: as fórmulas são em português

`=SOMA(1,5;2)` soma um e meio com dois. Duas regras andam juntas, e são as mesmas
do Excel em português:

- **vírgula é o separador decimal**;
- **ponto e vírgula separa os argumentos** — justamente porque a vírgula já está
  ocupada.

Se você digitar `=SOMA(A1,B1)`, o aplicativo diz qual é o separador certo em vez
de dar um erro genérico.

Os nomes funcionam nos **dois idiomas**: `SOMA` e `SUM` são a mesma função, assim
como `SE`/`IF` e `PROCV`/`VLOOKUP`. Quem colou uma fórmula de uma planilha
estrangeira não precisa traduzir na mão.

### Funções disponíveis

| Grupo         | Funções                                                                                                                  |
| ------------- | ------------------------------------------------------------------------------------------------------------------------ |
| Contas        | `SOMA` `SOMASE` `ARRED` `ARREDONDAR.PARA.CIMA` `ARREDONDAR.PARA.BAIXO` `ABS` `INT` `TRUNCAR` `RESTO` `RAIZ` `POTÊNCIA`     |
| Estatística   | `MÉDIA` `MÁXIMO` `MÍNIMO` `CONT.NÚM` `CONT.VALORES` `CONTAR.VAZIO` `CONT.SE`                                              |
| Lógica        | `SE` `SEERRO` `SENÃODISP` `E` `OU` `NÃO` `ÉERROS` `É.NÃO.DISP` `ÉNÚM` `ÉTEXTO` `ÉCÉL.VAZIA`                              |
| Texto         | `CONCATENAR` `NÚM.CARACT` `ESQUERDA` `DIREITA` `EXT.TEXTO` `MAIÚSCULA` `MINÚSCULA` `ARRUMAR` `SUBSTITUIR` `PROCURAR` `LOCALIZAR` `VALOR` |
| Procura       | `PROCV` `PROCH` `CORRESP` `ÍNDICE`                                                                                        |
| Data          | `HOJE` `AGORA` `DATA` `ANO` `MÊS` `DIA` `DIA.DA.SEMANA`                                                                   |

### Ainda não existe

Matrizes dinâmicas, referências de coluna inteira (`A:A`), intervalos nomeados, e
a alça de preenchimento copia o **valor**, não a fórmula deslocada.

Uma fórmula que use alguma função fora da lista mostra `#NOME?` na célula — mas
**continua no arquivo**, intacta, e o Excel volta a calculá-la normalmente. O
aviso na abertura diz quais funções são essas.

---

## Os avisos: leia, eles são diferentes entre si

O aplicativo distingue duas coisas que a maioria dos programas mistura:

- **invisibilidade** — o recurso continua no arquivo, mas não aparece aqui.
  Comentários, controle de alterações, gráficos, filtros, formatação
  condicional. Você pode editar e salvar à vontade: eles voltam intactos.
- **perda** — some de verdade ao salvar. Só acontece quando você edita
  justamente o trecho que ancorava o recurso.

São avisos separados porque exigem reações diferentes. Um alerta genérico é um
alerta que se aprende a fechar sem ler.

---

## Se o aplicativo fechar sozinho

De oito em oito segundos, o que está na tela é guardado num rascunho — **nunca
por cima do seu arquivo**. Se houver uma queda, na próxima abertura aparece uma
faixa azul oferecendo o trabalho de volta.

- **Recuperar** traz o conteúdo para a tela. Ele fica marcado como *não salvo* —
  porque é isso mesmo que ele é. Confira e salve onde quiser.
- **Descartar** apaga o rascunho de vez.

Enquanto a faixa estiver na tela, o rascunho não é sobrescrito. Você pode
ignorá-la, abrir outro arquivo, e o trabalho continua lá.

---

## Salvar não custa o que você não editou

Ao gravar um `.docx` ou `.xlsx`, o aplicativo **não regenera o arquivo**: ele
reescreve só o que você mexeu e devolve o resto exatamente como estava.

Medido num documento real de 105 blocos: salvar sem editar reescreve **zero**
blocos; editar um parágrafo reescreve **um**. Numa planilha do LibreOffice,
abrir e salvar sem editar escreve **zero** células.

Na prática: fonte, alinhamento, bordas, gráficos, tabelas dinâmicas, comentários
e filtros continuam no arquivo depois de você corrigir uma vírgula.

---

## Limites conhecidos

- **Não cria `.docx` do zero.** Documento novo nasce `.sdoc`; para virar `.docx`
  seria preciso gerar o pacote inteiro, e aí a promessa acima deixaria de valer.
  Planilha nova **pode** ser salva direto em `.xlsx`.
- **Cabeçalho e rodapé** vindos do arquivo são preservados e desenhados, mas não
  editáveis aqui.
- **Filtros de planilha** são preservados, mas não há tela para criá-los ou
  alterá-los.
- **Ao imprimir uma planilha**, sai a aba ativa inteira, ajustada à largura da
  página; as linhas congeladas viram cabeçalho e se repetem em cada folha. Se as
  colunas ficarem apertadas, ponha a página em paisagem pelo menu Arquivo →
  Configuração de página.
- **Arquivos acima de 20 MB** não abrem.
- **A paginação na tela é estimativa** (o `≈` na barra de status). A paginação
  exata é a da exportação para PDF.

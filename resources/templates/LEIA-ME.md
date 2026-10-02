# Modelos embutidos

Os modelos que a galeria de **Arquivo → Novo a partir de modelo…** oferece:

| Arquivo                    | Modelo                       |
| -------------------------- | ---------------------------- |
| `documento-em-branco.dotx` | Documento em branco          |
| `carta.dotx`               | Carta                        |
| `relatorio.dotx`           | Relatório com capa e sumário |
| `ata-reuniao.dotx`      | Ata de reunião               |

## Origem e licença

Texto, estrutura e estilos são originais deste projeto e seguem a licença dele
(MIT, ver `LICENSE`). Nenhum conteúdo, imagem ou fonte de terceiros vai dentro dos
pacotes: os estilos são os padrões do documento novo (Calibri 11, entrelinha
1,08, 8 pt depois do parágrafo), e a fonte é só citada pelo nome.

Os pacotes foram gerados pelo próprio sidecar — `DocxTemplate.Create` seguido de
`DocxWriter.Write` com o destino de modelo —, e por isso abrem e gravam pelo mesmo
caminho de qualquer documento criado no aplicativo. `docProps/core.xml` não leva
autor nem data.

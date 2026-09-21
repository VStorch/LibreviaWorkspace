# Dicionário de português do Brasil

`pt-BR-3-0.bdic` é o dicionário que a verificação ortográfica usa. Ele viaja
dentro do instalador porque **o aplicativo é offline**: o corretor embutido do
Chromium, deixado por conta própria, baixa o dicionário de um CDN na primeira
execução — e numa máquina sem rede isso significa nenhuma verificação, em
silêncio.

## Origem e licença

| Item      | Valor                                                             |
| --------- | ----------------------------------------------------------------- |
| Projeto   | VERO — Verificador Ortográfico do LibreOffice, dicionário `pt_BR`  |
| Autor     | Raimundo Santos Moura (2006–2013) e o LibreOffice Project          |
| Licença   | LGPL-3 ou MPL                                                     |
| Origem    | https://github.com/LibreOffice/dictionaries/tree/master/pt_BR      |

O arquivo distribuído aqui é o mesmo dicionário Hunspell (`pt_BR.aff` +
`pt_BR.dic`) já convertido para o formato binário que o Chromium lê, o `.bdic`.
A conversão é feita pelo `convert_dict` do próprio Chromium; em Debian e Ubuntu
o resultado já vem pronto no pacote `hunspell-pt-br`, em
`/usr/share/hunspell-bdic/pt_BR.bdic`, e é dali que esta cópia veio.

## Por que este nome de arquivo

`pt-BR-3-0.bdic` não é escolha nossa: é o nome que o Chromium procura dentro da
pasta `Dictionaries` do diretório de dados do usuário — `<código do idioma>` +
`-<revisão do formato>`. Renomear o arquivo faria o corretor não o encontrar e
tentar baixá-lo. Ver `src/main/spellcheck.ts`.

## Como atualizar

1. `apt-get download hunspell-pt-br` (ou baixe o `.bdic` da mesma origem);
2. copie o `.bdic` para cá com o nome que o Chromium espera;
3. `npx vitest run src/services/spell` — o teste confere a assinatura do arquivo;
4. `npm run notices`, para o aviso de terceiros registrar a versão nova.

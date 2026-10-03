# Política de segurança

O Librevia abre arquivos que vêm de outras pessoas. Um `.docx` ou um `.xlsx` pode ter sido feito
de propósito para travar o aplicativo ou para tentar sair do lugar onde ele roda. Por isso,
relatos de segurança são levados a sério e tratados em particular.

## Versões que recebem correção

Só a **versão mais recente**. O projeto ainda não mantém versões antigas em paralelo; a correção
sai numa versão nova.

## Como relatar

**Não abra uma issue pública.** Uma falha descrita em público fica à vista de quem quer
explorá-la antes de haver correção.

Use o **relato privado de vulnerabilidade** do GitHub:

1. abra a aba **Security** do repositório;
2. clique em **Report a vulnerability**;
3. descreva o problema.

Ajuda muito incluir:

- a versão do Librevia e o sistema operacional;
- os passos para reproduzir;
- o arquivo que provoca o problema, se houver um, ou uma descrição de como montá-lo;
- o que acontece, e o que você esperava que acontecesse.

## O que acontece depois

- Você recebe uma resposta confirmando o recebimento.
- O problema é reproduzido e avaliado.
- A correção sai numa versão nova, e o aviso de segurança é publicado junto, com o seu nome nos
  créditos, se você quiser.

O projeto é mantido por uma pessoa só, então os prazos são de boa-fé, e não contratuais.

## O que está no escopo

O aplicativo foi desenhado com algumas promessas. Quebrar qualquer uma delas é uma falha de
segurança:

| Promessa | Exemplo de falha |
| --- | --- |
| **Um arquivo não derruba o aplicativo** | um `.docx` ou `.xlsx` que trava o serviço de formatos, esgota a memória ou abre um ZIP que se expande sem fim |
| **A interface roda isolada** | um documento que consegue executar código fora da sandbox do renderer, ou acessar Node.js |
| **Só se grava onde o usuário escolheu** | gravar ou ler um caminho que não passou por um diálogo do sistema nem pela lista de recentes, furando a trava `PATH_NOT_AUTHORIZED` |
| **O aplicativo não usa a rede** | um documento que faz o aplicativo buscar algo na internet sozinho, ao abrir ou ao imprimir |
| **O serviço de formatos não grava nada** | o sidecar .NET escrevendo no disco ou abrindo conexão |

## O que fica fora do escopo

- Arquivos acima de 20 MB, que o aplicativo já recusa antes de ler.
- Macros de um `.dotm`: elas não são executadas nem copiadas para o documento novo.
- Problemas que exigem que alguém já tenha controle do computador do usuário.
- Falhas no Electron, no Chromium ou no .NET que já tenham correção publicada por eles. Nesse
  caso, avise que o Librevia está com uma versão desatualizada; isso é bem-vindo, mas como issue
  comum.

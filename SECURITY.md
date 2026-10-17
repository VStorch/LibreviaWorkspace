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
| **Um PDF fica no visualizador** | um PDF que, aberto, alcança a página ou as funções do aplicativo, abre uma janela, tira a janela do aplicativo ou lê outro arquivo além dele |
| **Um link só sai com confirmação** | um link num arquivo que abre o navegador sem que o usuário veja o endereço e aceite |
| **Só se grava onde o usuário escolheu** | gravar ou ler um caminho que não passou por um diálogo do sistema nem pela lista de recentes, furando a trava `PATH_NOT_AUTHORIZED` |
| **O aplicativo só usa a rede para o assistente de IA** | um documento que faz o aplicativo buscar algo na internet sozinho, ao abrir ou ao imprimir; qualquer conexão sem um assistente configurado, salvo a consulta a `localhost` quando a pessoa clica em "Procurar assistentes" |
| **O assistente envia só o documento aberto ou a planilha ativa, e só com permissão** | texto de outro arquivo, de outra aba da planilha, o nome do arquivo ou as propriedades indo para o assistente; um envio para fora do computador sem a confirmação feita pelo processo principal, que diz que o documento aberto ou a planilha ativa vai junto como contexto; uma ferramenta da conversa que mude o documento sem o clique da pessoa ou que leia algo além do texto enviado com a mensagem: as ferramentas só leem esse texto e propõem |
| **O Librevia não dá ao agente de IA acesso ao computador** | o Claude Code rodando com alguma ferramenta (ele é chamado com `--tools ""`); o Librevia oferecendo a um agente ACP arquivos ou terminal, ou aceitando um pedido de permissão dele; o agente rodando fora de uma pasta temporária vazia; o programa do agente vindo de um caminho escolhido pela interface, sem passar por diálogo do sistema; o processo ou os filhos dele seguindo vivos depois de parar ou de fechar a janela. As ferramentas próprias de um agente ACP (as do Gemini CLI, do Codex, do Claude Agent…) seguem a configuração dele, não a do Librevia: quem quiser restringi-las configura o agente |
| **A chave do assistente não vaza** | a chave voltando para a interface, aparecendo em log ou gravada sem cifra no disco |
| **O serviço de formatos não grava nada** | o sidecar .NET escrevendo no disco ou abrindo conexão |
| **As propriedades de um PDF só acrescentam** | gravar as propriedades de um PDF muda algum byte do que já estava no arquivo, ou grava num PDF que não foi o aberto |

## O que fica fora do escopo

- Arquivos acima de 20 MB, que o aplicativo já recusa antes de ler.
- Macros de um `.dotm`: elas não são executadas nem copiadas para o documento novo.
- Problemas que exigem que alguém já tenha controle do computador do usuário.
- Falhas no Electron, no Chromium ou no .NET que já tenham correção publicada por eles. Nesse
  caso, avise que o Librevia está com uma versão desatualizada; isso é bem-vindo, mas como issue
  comum.

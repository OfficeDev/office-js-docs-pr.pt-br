---
title: Suplementos do Outlook para o Outlook Mobile
description: Os complementos do Outlook Mobile têm suporte em todas as contas comerciais do Microsoft 365, em Outlook.com e o suporte estará em breve nas contas do Gmail.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 586a473e1036e8480f395da49011f540d87e1b5f
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270704"
---
# <a name="add-ins-for-outlook-mobile"></a>Suplementos do Outlook Mobile

Os suplementos agora funcionam no Outlook Mobile, usando as mesmas APIs disponíveis para outros pontos de extremidade do Outlook. Se você já tiver criado um suplemento para Outlook, é fácil fazê-lo funcionar no Outlook Mobile.

Os complementos do Outlook Mobile têm suporte em todas as contas comerciais do Microsoft 365, Outlook.com e o suporte estará chegando em breve às contas do Gmail.

**Um painel de tarefas de exemplo no Outlook no iOS**

![Uma captura de tela do painel de tarefas no Outlook no iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Um painel de tarefas de exemplo no Outlook no Android**

![Uma captura de tela do painel de tarefas no Outlook no Android](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> Os complementos não funcionam na versão moderna do Outlook em um navegador móvel. Para saber mais, confira [o Outlook em seu navegador móvel que está sendo atualizado.](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)

## <a name="whats-different-on-mobile"></a>Qual é a diferença no celular?

- O tamanho pequeno e as rápidas interações tornam o projeto para celular um desafio. Para garantir experiências de qualidade para nossos clientes, estamos definindo critérios rígidos de validação que devem ser cumpridos por um suplemento que declara suporte a celular de forma a ser aprovado na AppSource.
  - O suplemento **DEVE** cumprir as [diretrizes de interface do usuário](outlook-addin-design.md).
  - O cenário do suplemento **DEVE** [fazer sentido no mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

- Em geral, somente o modo De leitura de mensagem é suportado no momento. Isso significa `MobileMessageReadCommandSurface` que é o único [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) que você deve declarar na seção móvel do manifesto. No entanto, o modo Organizador de Compromissos é suportado para os complementos integrados do provedor de reunião online que, em vez disso, declaram o ponto de extensão [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface). Consulte o [artigo Criar um complemento móvel do Outlook para um provedor](online-meeting.md) de reuniões online para saber mais sobre esse cenário.

- A API [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) não é suportada no celular, já que o aplicativo móvel usa APIs REST para se comunicar com o servidor. Se seu back-end do aplicativo precisa se conectar ao servidor do Exchange, é possível usar o token de retorno de chamada para fazer chamadas de API REST. Para obter detalhes, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](use-rest-api.md).

- Quando você envia o suplemento para a loja com [MobileFormFactor](../reference/manifest/mobileformfactor.md) no manifesto, precisará concordar com nosso adendo de suplementos no iOS e precisará enviar sua ID de desenvolvedor Apple para verificação.

- Por fim, seu manifesto precisará declarar `MobileFormFactor` e ter os tipos corretos de [controles](../reference/manifest/control.md) e [tamanhos de ícone](../reference/manifest/icon.md) incluídos.

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>O que forma um bom cenário para suplementos móveis?

Lembre-se de que o tamanho médio da sessão Outlook em um telefone é bem menor do que em um PC. Isso significa que seu suplemento deve ser rápido e o cenário deve permitir que o usuário entre, saia e prossiga com seu fluxo de email.

Estes são exemplos de cenários que fazem sentido no Outlook Mobile.

- O suplemento traz informações valiosas para o Outlook, para ajudar os usuários na triagem dos emails e a responder adequadamente. Exemplo: um suplemento CRM que permite ao usuário ver informações do cliente e compartilhar informações apropriadas.

- O suplemento agrega valor ao conteúdo do email do usuário, salvando as informações em um controle, uma colaboração ou um sistema semelhante. Exemplo: um suplemento que permite aos usuários ativar emails em itens de tarefa para acompanhamento de projetos, ou tíquetes de ajuda, para uma equipe de suporte.

**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no iOS**

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no Android**

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>Teste seus suplementos no celular

Para testar um suplemento no Outlook Mobile, você pode carregar um suplemento para uma conta do O365 ou do Outlook.com. No Outlook na Web, acesse a engrenagem de configurações e escolha **Gerenciar Integrações** ou **Gerenciar Suplementos**. Perto da parte superior, clique onde diz: **Clique aqui para adicionar um suplemento personalizado** e carregue seu manifesto. Verifique se seu manifesto está formatado corretamente para conter `MobileFormFactor` ou ele não será carregado.

Depois que seu suplemento estiver funcionando, certifique-se de testá-lo em tamanhos de tela diferentes, incluindo celulares e tablets. Você deve verificar se ele atende às diretrizes de acessibilidade de contraste, tamanho da fonte e cor, bem como de usabilidade com um leitor de tela, como o VoiceOver no iOS ou TalkBack no Android.

A solução de problemas em dispositivos móveis pode ser difícil, pois talvez você não tenha as ferramentas com as que está acostumado. No entanto, uma opção para solucionar problemas no iOS é usar o Fiddler (confira este tutorial sobre como [usá-lo com um dispositivo iOS).](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)

## <a name="next-steps"></a>Próximas etapas

Saiba como:

- [Adicionar suporte móvel ao manifesto do seu suplemento](add-mobile-support.md).
- [Projetar uma ótima experiência móvel para seu suplemento](outlook-addin-design.md).
- [Obter um token de acesso e chamar APIs REST do Outlook](use-rest-api.md) do suplemento.

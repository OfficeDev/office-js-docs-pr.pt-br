---
title: Suplementos do Outlook para o Outlook Mobile
description: Os suplementos do Outlook Mobile têm suporte em todas as contas do Office 365 Comercial, Outlook.com e, em breve, haverá suporte para contas do Gmail.
ms.date: 04/13/2020
localization_priority: Normal
ms.openlocfilehash: 4b6341ac1b340ebc46c616ae4274bfdf1e2d0672
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241081"
---
# <a name="add-ins-for-outlook-mobile"></a>Suplementos do Outlook Mobile

Os suplementos agora funcionam no Outlook Mobile, usando as mesmas APIs disponíveis para outros pontos de extremidade do Outlook. Se você já tiver criado um suplemento para Outlook, é fácil fazê-lo funcionar no Outlook Mobile.

Os suplementos do Outlook Mobile têm suporte em todas as contas do Office 365 Comercial, Outlook.com e, em breve, haverá suporte para contas do Gmail.

**Um painel de tarefas de exemplo no Outlook no iOS**

![Uma captura de tela do painel de tarefas no Outlook no iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Um painel de tarefas de exemplo no Outlook no Android**

![Uma captura de tela do painel de tarefas no Outlook no Android](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> Os suplementos não funcionam na versão moderna do Outlook em um navegador móvel. Para obter mais informações, consulte [Outlook em seu navegador móvel está sendo atualizado](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).

## <a name="whats-different-on-mobile"></a>Qual é a diferença no celular?

- O tamanho pequeno e as rápidas interações tornam o projeto para celular um desafio. Para garantir experiências de qualidade para nossos clientes, estamos definindo critérios rígidos de validação que devem ser cumpridos por um suplemento que declara suporte a celular de forma a ser aprovado na AppSource.
    - O suplemento **DEVE** cumprir as [diretrizes de interface do usuário](outlook-addin-design.md).
    - O cenário do suplemento **DEVE** [fazer sentido no mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

- Em geral, só há suporte para o modo de leitura de mensagens no momento. Isso significa `MobileMessageReadCommandSurface` que é o único [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) que você deve declarar na seção móvel do seu manifesto. No entanto, o modo organizador de compromissos tem suporte para suplementos integrados de provedor de reunião online que, em vez disso, declare o [ponto de extensão MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview). Confira o artigo [criar um suplemento do Outlook Mobile para um provedor de reunião online](online-meeting.md) para saber mais sobre esse cenário.

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

A solução de problemas em dispositivos móveis pode ser difícil, já que pode você não ter as ferramentas com as quais está acostumado. Uma opção para solução de problemas é [usar Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md). Ou, se tiver usado o Fiddler antes, confira [este tutorial sobre como usá-lo com um dispositivo iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices).

## <a name="next-steps"></a>Próximas etapas

Saiba como:

- [Adicionar suporte móvel ao manifesto do seu suplemento](add-mobile-support.md).
- [Projetar uma ótima experiência móvel para seu suplemento](outlook-addin-design.md).
- [Obter um token de acesso e chamar APIs REST do Outlook](use-rest-api.md) do suplemento.

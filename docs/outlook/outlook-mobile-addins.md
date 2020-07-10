---
title: Suplementos do Outlook para o Outlook Mobile
description: Os suplementos do Outlook Mobile têm suporte em todas as contas de negócios do Microsoft 365, contas do Outlook.com e o suporte é disponibilizado em breve para contas do gmail.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 34fbb01d596c4da38fe81438088cd71d8c7e152a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093893"
---
# <a name="add-ins-for-outlook-mobile"></a>Suplementos do Outlook Mobile

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

Os suplementos do Outlook Mobile têm suporte em todas as contas de negócios do Microsoft 365, contas do Outlook.com e o suporte é disponibilizado em breve para contas do gmail.

**Um painel de tarefas de exemplo no Outlook no iOS**

![Uma captura de tela do painel de tarefas no Outlook no iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Um painel de tarefas de exemplo no Outlook no Android**

![Uma captura de tela do painel de tarefas no Outlook no Android](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> Os suplementos não funcionam na versão moderna do Outlook em um navegador móvel. Para obter mais informações, consulte [Outlook em seu navegador móvel está sendo atualizado](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).

## <a name="whats-different-on-mobile"></a>Qual é a diferença no celular?

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
    - O suplemento **DEVE** cumprir as [diretrizes de interface do usuário](outlook-addin-design.md).
    - O cenário do suplemento **DEVE** [fazer sentido no mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

- Em geral, só há suporte para o modo de leitura de mensagens no momento. Isso significa que `MobileMessageReadCommandSurface` é o único [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) que você deve declarar na seção móvel do seu manifesto. No entanto, o modo organizador de compromissos tem suporte para suplementos integrados de provedor de reunião online que, em vez disso, declare o [ponto de extensão MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview). Confira o artigo [criar um suplemento do Outlook Mobile para um provedor de reunião online](online-meeting.md) para saber mais sobre esse cenário.

- The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- Quando você envia o suplemento para a loja com [MobileFormFactor](../reference/manifest/mobileformfactor.md) no manifesto, precisará concordar com nosso adendo de suplementos no iOS e precisará enviar sua ID de desenvolvedor Apple para verificação.

- Por fim, seu manifesto precisará declarar `MobileFormFactor` e ter os tipos corretos de [controles](../reference/manifest/control.md) e [tamanhos de ícone](../reference/manifest/icon.md) incluídos.

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>O que forma um bom cenário para suplementos móveis?

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Estes são exemplos de cenários que fazem sentido no Outlook Mobile.

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no iOS**

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no Android**

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>Teste seus suplementos no celular

To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

A solução de problemas no Mobile pode ser difícil, já que você pode não ter as ferramentas para as quais você está acostumado. No entanto, uma opção de solução de problemas no iOS é usar o Fiddler (Confira [este tutorial sobre como usá-lo com um dispositivo IOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).

## <a name="next-steps"></a>Próximas etapas

Saiba como:

- [Adicionar suporte móvel ao manifesto do seu suplemento](add-mobile-support.md).
- [Projetar uma ótima experiência móvel para seu suplemento](outlook-addin-design.md).
- [Obter um token de acesso e chamar APIs REST do Outlook](use-rest-api.md) do suplemento.

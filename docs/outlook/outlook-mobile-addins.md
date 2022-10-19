---
title: Suplementos do Outlook para o Outlook Mobile
description: Os suplementos móveis do Outlook têm suporte em todas as contas comerciais do Microsoft 365 e Outlook.com contas.
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: ca09ba550d8d2ed6e9003e85a8d042f413a6ab52
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607559"
---
# <a name="add-ins-for-outlook-mobile"></a>Suplementos do Outlook Mobile

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

Os suplementos móveis do Outlook têm suporte em todas as contas comerciais do Microsoft 365 e Outlook.com contas. No entanto, o suporte não está disponível atualmente em contas do Gmail.

**Um painel de tarefas de exemplo no Outlook no iOS**

![Captura de tela de um painel de tarefas no Outlook no iOS.](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Um painel de tarefas de exemplo no Outlook no Android**

![Captura de tela de um painel de tarefas no Outlook no Android.](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a>Qual é a diferença no celular?

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
  - O suplemento **DEVE** cumprir as [diretrizes de interface do usuário](outlook-addin-design.md).
  - O cenário do suplemento **DEVE** [fazer sentido no mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

- Em geral, somente o modo de Leitura de Mensagem tem suporte no momento. Isso significa `MobileMessageReadCommandSurface` que é o único [ExtensionPoint](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) que você deve declarar na seção móvel do manifesto. No entanto, há algumas exceções:
  1. O modo Organizador de Compromissos tem suporte para suplementos integrados do provedor de reunião online que, em vez disso, declaram o ponto de extensão [MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface). Consulte o [artigo Criar um suplemento móvel do Outlook para um provedor de reunião online](online-meeting.md) para saber mais sobre esse cenário.
  1. O modo participante do compromisso tem suporte para suplementos integrados criados por provedores de aplicativos crm (gerenciamento de relacionamento com o cliente) e anotações. Esses suplementos devem declarar o ponto de extensão [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee). Consulte as [anotações de compromisso de log para um aplicativo externo no artigo de suplementos](mobile-log-appointments.md) móveis do Outlook para saber mais sobre esse cenário.

- The [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- Quando você envia o suplemento para a loja com [MobileFormFactor](/javascript/api/manifest/mobileformfactor) no manifesto, precisará concordar com nosso adendo de suplementos no iOS e precisará enviar sua ID de desenvolvedor Apple para verificação.

- Por fim, seu manifesto precisará declarar `MobileFormFactor` e ter os tipos corretos de [controles](/javascript/api/manifest/control) e [tamanhos de ícone](/javascript/api/manifest/icon) incluídos.

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>O que forma um bom cenário para suplementos móveis?

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Estes são exemplos de cenários que fazem sentido no Outlook Mobile.

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no iOS**

![GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no iOS.](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no Android**

![GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no Android.](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>Teste seus suplementos no celular

Para testar um suplemento no Outlook Mobile, primeiro realizar o [sideload](sideload-outlook-add-ins-for-testing.md) de um suplemento em uma conta do Microsoft 365 ou Outlook.com na Web, no Windows ou no Mac. Verifique se o manifesto está formatado corretamente para conter `MobileFormFactor` ou se ele não será carregado no cliente do Outlook no celular.

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

A solução de problemas em dispositivos móveis pode ser difícil, pois talvez você não tenha as ferramentas com as que está acostumado. No entanto, uma opção para solucionar problemas no iOS é usar o Fiddler (confira este tutorial sobre como [usá-lo com um dispositivo iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).

> [!NOTE]
> As Outlook na Web modernas em smartphones iPhone e Android não são mais necessárias ou estão disponíveis para testar suplementos do Outlook. Além disso, não há suporte para suplementos no Outlook no Android, no iOS e na Web móvel moderna com contas locais do Exchange. Determinados dispositivos iOS ainda dão suporte a suplementos ao usar contas locais do Exchange com Outlook na Web. Para obter informações sobre os dispositivos suportados, confira [Requisitos para executar Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet).

## <a name="next-steps"></a>Próximas etapas

Saiba como:

- [Adicionar suporte móvel ao manifesto do seu suplemento](add-mobile-support.md).
- [Projetar uma ótima experiência móvel para seu suplemento](outlook-addin-design.md).
- [Obter um token de acesso e chamar APIs REST do Outlook](use-rest-api.md) do suplemento.

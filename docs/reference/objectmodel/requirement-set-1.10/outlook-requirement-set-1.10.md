---
title: Outlook conjunto de requisitos de API de complemento 1.10
description: Conjunto de requisitos 1.10 para Outlook api de complemento.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 9e3e30590279036a08a93d8643cd56c2c73be78c
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671258"
---
# <a name="outlook-add-in-api-requirement-set-110"></a>Outlook conjunto de requisitos de API de complemento 1.10

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

## <a name="whats-new-in-110"></a>Novidades no 1.10?

O conjunto de requisitos 1.10 inclui todos os recursos do conjunto [de requisitos 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs para [ativação baseada em eventos](../../../outlook/autolaunch.md) e recursos de assinatura de email.
- Adicionada a capacidade de incluir uma ação personalizada em uma mensagem de notificação.

### <a name="change-log"></a>Log de mudanças

- Adicionado [o ponto de extensão LaunchEvent:](../../manifest/extensionpoint.md#launchevent)adiciona um novo tipo de ExtensionPoint com suporte. Ele configura a funcionalidade de ativação baseada em evento.
- Elemento [de manifesto LaunchEvents](../../manifest/launchevents.md)adicionado : adiciona um elemento de manifesto para dar suporte à configuração da funcionalidade de ativação baseada em eventos.
- Elemento [de manifesto Runtimes modificado:](../../manifest/runtimes.md)adiciona Outlook suporte. Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.
- Adicionado [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#setSignatureAsync_data__options__callback_): adiciona uma nova função ao `Body` objeto. Ele adiciona ou substitui a assinatura no corpo do item no modo Redação.
- Adicionado [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods): adiciona uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo Redação.
- Adicionado [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getComposeTypeAsync_options__callback_): adiciona uma nova função que obtém o tipo de composição de uma mensagem no modo Redação.
- Adicionado [Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods): adiciona uma nova função que verifica se a assinatura do cliente está habilitada no item no modo Redação.
- Adicionado [Office. MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype): adiciona um novo número. Ele representa o tipo de ação personalizada em uma mensagem de notificação.
- Adicionado [Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true): adiciona um novo número disponível no modo Redação.
- Adicionado [Office. MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype): adiciona um novo tipo ao `ItemNotificationMessageType` número. Ele representa uma mensagem de notificação com uma ação personalizada.
- Adicionado [Office. NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction): adiciona um novo objeto para que você possa definir uma ação personalizada para sua `InsightMessage` notificação.
- Adicionado [Office. NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions): adiciona uma nova propriedade que permite adicionar uma `InsightMessage` notificação com uma ação personalizada.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

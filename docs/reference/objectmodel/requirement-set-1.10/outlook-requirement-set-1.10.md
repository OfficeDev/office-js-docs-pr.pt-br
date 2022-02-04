---
title: Outlook conjunto de requisitos de API de complemento 1.10
description: Conjunto de requisitos 1.10 para Outlook api de complemento.
ms.date: 11/04/2021
ms.localizationpriority: medium
---

# <a name="outlook-add-in-api-requirement-set-110"></a>Outlook conjunto de requisitos de API de complemento 1.10

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o mais recente.

## <a name="whats-new-in-110"></a>Novidades no 1.10?

O conjunto de requisitos 1.10 inclui todos os recursos do conjunto [de requisitos 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs para [ativação baseada em eventos](../../../outlook/autolaunch.md) e recursos de assinatura de email.
- Adicionado suporte para o [objeto OfficeRuntime.Armazenamento](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true) com o recurso de ativação baseada em evento.
- Adicionada a capacidade de incluir uma ação personalizada em uma mensagem de notificação.

### <a name="change-log"></a>Log de mudanças

- Adicionado [o ponto de extensão LaunchEvent](../../manifest/extensionpoint.md#launchevent): adiciona um novo tipo de ExtensionPoint com suporte. Ele configura a funcionalidade de ativação baseada em evento.
- Elemento [de manifesto LaunchEvents](../../manifest/launchevents.md) adicionado: adiciona um elemento de manifesto para dar suporte à configuração da funcionalidade de ativação baseada em eventos.
- Elemento [de manifesto Runtimes modificado](../../manifest/runtimes.md): adiciona Outlook suporte. Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.
- Adicionado [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#outlook-office-body-setsignatureasync-member(1)): adiciona uma nova função ao `Body` objeto. Ele adiciona ou substitui a assinatura no corpo do item no modo Redação.
- Adicionado [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods): adiciona uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo Redação.
- Adicionado [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#outlook-office-messagecompose-getcomposetypeasync-member(1)): adiciona uma nova função que obtém o tipo de composição de uma mensagem no modo Redação.
- Adicionado [Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods): adiciona uma nova função que verifica se a assinatura do cliente está habilitada no item no modo Redação.
- Adicionado [Office. MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype?view=outlook-js-1.10&preserve-view=true): adiciona um novo número. Ele representa o tipo de ação personalizada em uma mensagem de notificação.
- Adicionado [Office. MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true): adiciona um novo número disponível no modo Redação.
- Adicionado [Office. MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.10&preserve-view=true): adiciona um novo tipo ao `ItemNotificationMessageType` número. Ele representa uma mensagem de notificação com uma ação personalizada.
- Adicionado [Office. NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction?view=outlook-js-1.10&preserve-view=true): adiciona um novo objeto para que você possa definir uma ação personalizada para sua `InsightMessage` notificação.
- Adicionado [Office. NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.10&preserve-view=true#outlook-office-notificationmessagedetails-actions-member): adiciona uma nova propriedade `InsightMessage` que permite adicionar uma notificação com uma ação personalizada.
- Modificado [OfficeRuntime.Armazenamento](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true): adiciona Outlook suporte, mas somente com o recurso de ativação baseada em evento.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

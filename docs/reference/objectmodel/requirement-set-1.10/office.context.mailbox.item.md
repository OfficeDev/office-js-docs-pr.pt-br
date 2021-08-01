---
title: Office.context.mailbox.item - conjunto de requisitos 1.10
description: Outlook Conjunto de requisitos da API de Caixa de Correio versão 1.10 do modelo de objeto Item.
ms.date: 07/16/2021
localization_priority: Normal
ms.openlocfilehash: 198fbc27be40bed5b368be2885f55e2ae9aeb71c
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671566"
---
# <a name="item-mailbox-requirement-set-110"></a>item (Conjunto de requisitos de caixa de correio 1.10)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` é usado para acessar a mensagem, solicitação de reunião ou compromisso selecionado no momento. Você pode determinar o tipo do item usando a `itemType` propriedade.

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Nível de permissão mínimo](../../../outlook/understanding-outlook-add-in-permissions.md)|Restrito|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)|Organizador de Compromissos, Participante do Compromisso,<br>Redação de mensagens ou leitura de mensagem|

> [!IMPORTANT]
> Android e iOS: há limitações sobre quando os complementos são ativados e quais APIs estão disponíveis. Para saber mais, consulte [Adicionar suporte móvel a um suplemento do Outlook](../../../outlook/add-mobile-support.md#compose-mode-and-appointments).

## <a name="properties"></a>Propriedades

| Propriedade | Minimum<br>nível de permissão | Detalhes por modo | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|---|:---:|
| anexos | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#bcc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| corpo | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Categorias | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#cc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#conversationId) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#conversationId) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#dateTimeCreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#dateTimeCreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#dateTimeModified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#dateTimeModified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#end) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#end)<br>(Solicitação de Reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#from) | [De](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#internetHeaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#internetMessageId) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemClass) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemClass) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemId) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemId) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| localização | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#location) | [Localização](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#location) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#location)<br>(Solicitação de Reunião) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#normalizedSubject) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#normalizedSubject) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#optionalAttendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#optionalAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#organizer) | [Organizador](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#recurrence)<br>(Solicitação de Reunião) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#requiredAttendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#requiredAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| remetente | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#seriesId) | Cadeia de caracteres | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#seriesId) | Cadeia de caracteres | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#seriesId) | Cadeia de caracteres | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#seriesId) | Cadeia de caracteres | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| iniciar | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#start) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#start)<br>(Solicitação de Reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Assunto | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#subject) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#subject) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| para | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#to) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Detalhes por modo | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| disableClientSignatureAsync([options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| displayReplyAllForm(formData) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync(formData, [options], [callback]) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync(formData, [options], [callback]) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync([options], [callback]) | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getAllInternetHeadersAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync(attachmentId, [options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync([options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getComposeTypeAsync([options], callback) | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getComposeTypeAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| getEntities() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getItemIdAsync([options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync([options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| isClientSignatureEnabledAsync([options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Eventos

Você pode se inscrever e cancelar a assinatura dos seguintes eventos usando `addHandlerAsync` `removeHandlerAsync` e, respectivamente.

> [!IMPORTANT]
> Os eventos só estão disponíveis com a implementação do painel de tarefas.

| [Event](/javascript/api/office/office.eventtype) | Descrição | Minimum<br>conjunto de requisitos |
|---|---|:---:|
|`AppointmentTimeChanged`| A data ou hora do compromisso ou série selecionado foi alterada. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| Um anexo foi adicionado ou removido do item. | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| O local do compromisso selecionado foi alterado. | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`RecipientsChanged`| A lista de destinatários do item ou local do compromisso selecionado foi alterada. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| O padrão de recorrência da série selecionada foi alterado. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

## <a name="example"></a>Exemplo

O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

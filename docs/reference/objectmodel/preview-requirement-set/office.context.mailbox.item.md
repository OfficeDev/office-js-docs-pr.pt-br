---
title: Office.context.mailbox.item - conjunto de requisitos de visualização
description: Outlook Versão do conjunto de requisitos de visualização da API de Caixa de Correio do modelo de objeto Item.
ms.date: 08/27/2021
localization_priority: Normal
ms.openlocfilehash: 60b634ff3ddeeacbcd875086e5041eb8207ef958
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936925"
---
# <a name="item-mailbox-preview-requirement-set"></a>item (Conjunto de requisitos de visualização de caixa de correio)

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
| attachments | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#bcc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| corpo | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#cc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#conversationId) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#dateTimeCreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#dateTimeCreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#dateTimeModified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#dateTimeModified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| delayDeliveryTime | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#delayDeliveryTime) | [DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime) | [Visualização](outlook-requirement-set-preview.md) |
| end | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#end) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#end)<br>(Solicitação de Reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#from) | [De](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#internetHeaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#internetMessageId) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| isAllDayEvent | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#isAllDayEvent) | [IsAllDayEvent](/javascript/api/outlook/office.isalldayevent) | [Visualização](outlook-requirement-set-preview.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent) | Booliano | [Visualização](outlook-requirement-set-preview.md) |
| itemClass | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#location) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#location)<br>(Solicitação de Reunião) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#normalizedSubject) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#normalizedSubject) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalAttendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#optionalattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#organizer) | [Organizador](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#recurrence) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#recurrence)<br>(Solicitação de Reunião) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredAttendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#requiredAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sensitivity | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#sensitivity) | [Sensitivity](/javascript/api/outlook/office.sensitivity) | [Visualização](outlook-requirement-set-preview.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity) | [MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype) | [Visualização](outlook-requirement-set-preview.md) |
| seriesId | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId) | Cadeia de caracteres | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| sessionData | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#sessionData) | [SessionData](/javascript/api/outlook/office.sessiondata) | [Visualização](outlook-requirement-set-preview.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#sessionData) | [SessionData](/javascript/api/outlook/office.sessiondata) | [Visualização](outlook-requirement-set-preview.md) |
| iniciar | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#start) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#start)<br>(Solicitação de Reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| assunto | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#subject) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| para | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#to) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Detalhes por modo | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| disableClientSignatureAsync([options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| displayReplyAllForm(formData) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync(formData, [options], [callback]) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync(formData, [options], [callback]) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync([options], [callback]) | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getAllInternetHeadersAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync(attachmentId, [options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync([options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getComposeTypeAsync([options], callback) | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| getEntities() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getInitializationContextAsync([options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| getItemIdAsync([options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync([options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| isClientSignatureEnabledAsync([options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Events

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

---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,8
description: Conjunto de requisitos da API de caixa de correio do Outlook versão 1,8 do modelo de objeto do item.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: d4a175694375bcddc5bf487107ff4c51bfa5c2ca
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610465"
---
# <a name="item-mailbox-requirement-set-18"></a>Item (conjunto de requisitos de caixa de correio 1,8)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`é usado para acessar a mensagem, solicitação de reunião ou compromisso atualmente selecionado. Você pode determinar o tipo do item usando a `itemType` propriedade.

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Nível de permissão mínimo](../../../outlook/understanding-outlook-add-in-permissions.md)|Restrito|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)|Organizador de compromisso, participante do compromisso<br>Composição de mensagem ou leitura de mensagem|

## <a name="properties"></a>Propriedades

| Propriedade | Mínimo<br>nível de permissão | Detalhes por modo | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|---|:---:|
| attachments | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#bcc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| corpo | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#categories) | [Categorias](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#categories) | [Categorias](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#categories) | [Categorias](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#categories) | [Categorias](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#cc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#cc) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#end)<br>(Solicitação de reunião) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#from) | [De](/javascript/api/outlook/office.from) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Internetheaders: | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| localização | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#location) | [Localização](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#location)<br>(Solicitação de reunião) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#optionalattendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#optionalattendees) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#organizer) | [Organizador](/javascript/api/outlook/office.organizer) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#recurrence) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#recurrence) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#recurrence)<br>(Solicitação de reunião) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#requiredattendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#requiredattendees) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| remetente | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesid | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| iniciar | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#start)<br>(Solicitação de reunião) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| para | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#to) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#to) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| Método | Mínimo<br>nível de permissão | Detalhes por modo | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback]) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getAllInternetHeadersAsync ([opções], [callback]) | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getallinternetheadersasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync (attachmentid, [opções], [callback]) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync ([opções], [callback]) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getentities () | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restricted | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (nome) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getItemIdAsync ([opções], retorno de chamada) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches () | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (nome) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType, [opções], retorno de chamada) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities () | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches () | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync ([opções], retorno de chamada) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.8#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Eventos

Você pode inscrever-se e cancelar a assinatura dos eventos a seguir usando o `addHandlerAsync` e o, `removeHandlerAsync` respectivamente.

| Evento | Descrição | Mínimo<br>conjunto de requisitos |
|---|---|:---:|
|`AppointmentTimeChanged`| A data ou hora do compromisso ou série selecionado foi alterada. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| Um anexo foi adicionado ou removido do item. | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| O local do compromisso selecionado foi alterado. | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`RecipientsChanged`| A lista de destinatários do item selecionado ou local do compromisso foi alterada. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| O padrão de recorrência da série selecionada foi alterado. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

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

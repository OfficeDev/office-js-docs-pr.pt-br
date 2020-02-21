---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 02/03/2020
localization_priority: Normal
ms.openlocfilehash: d4a65fd4e53d86d39c6b2e5bcd2145e32ece2225
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163854"
---
# <a name="item"></a>item

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
| attachments | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#bcc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#cc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#cc) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimecreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimecreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimemodified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimemodified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#end) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#end)<br>(Solicitação de reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#from) | [De](/javascript/api/outlook/office.from) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Internetheaders: | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| localização | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) | [Localização](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#location)<br>(Solicitação de reunião) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#optionalattendees) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) | [Organizador](/javascript/api/outlook/office.organizer) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#recurrence) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#recurrence) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#recurrence)<br>(Solicitação de reunião) | [Recorrência](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#requiredattendees) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| remetente | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesid | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| iniciar | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#start) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#start)<br>(Solicitação de reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| assunto | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| para | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#to) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#to) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| Método | Mínimo<br>nível de permissão | Detalhes por modo | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| addFileAttachmentAsync | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close | Restricted | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getAllInternetHeadersAsync | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getallinternetheadersasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getEntities | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType | Restricted | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getInitializationContextAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| getItemIdAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| loadCustomPropertiesAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Eventos

Você pode inscrever-se e cancelar a assinatura dos eventos `addHandlerAsync` a `removeHandlerAsync` seguir usando o e o, respectivamente.

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

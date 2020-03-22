---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,1
description: Conjunto de requisitos da API de caixa de correio do Outlook versão 1,1 do modelo de objeto do item.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f4f80a14143303de5d455d2618f30e3e7df585d1
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890841"
---
# <a name="item-mailbox-requirement-set-11"></a>Item (conjunto de requisitos de caixa de correio 1,1)

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
| attachments | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#bcc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| corpo | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#cc) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#cc) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#datetimecreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#datetimecreated) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#datetimemodified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#datetimemodified) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#end) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#end)<br>(Solicitação de reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#location)<br>(Solicitação de reunião) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| optionalAttendees | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#optionalattendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#optionalattendees) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#requiredattendees) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#requiredattendees) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| remetente | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| iniciar | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#start) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#start)<br>(Solicitação de reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| assunto | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#subject) | [Assunto](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| para | ReadItem | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#to) | [Destinatários](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#to) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| Método | Mínimo<br>nível de permissão | Detalhes por modo | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllForm(formData) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getentities () | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restricted | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (nome) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches () | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (nome) | ReadItem | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.1#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organizador de compromisso](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composição de mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

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

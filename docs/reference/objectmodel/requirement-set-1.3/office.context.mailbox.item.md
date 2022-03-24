---
title: Office.context.mailbox.item - conjunto de requisitos 1.3
description: Outlook da API de Caixa de Correio do conjunto de requisitos 1.3 do modelo de objeto Item.
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1bdafd9bae5247831e47a491f0a1ed8ab0641a03
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745988"
---
# <a name="item-mailbox-requirement-set-13"></a>item (Conjunto de requisitos de caixa de correio 1.3)

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
| anexos | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-bcc-member) | [Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| corpo | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-cc-member) | [Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-cc-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-conversationid-member) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-conversationid-member) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-datetimecreated-member) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-datetimecreated-member) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-datetimemodified-member) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-datetimemodified-member) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-end-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-end-member) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-end-member)<br>(Solicitação de Reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-from-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-internetmessageid-member) | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-itemclass-member) | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-itemclass-member) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-itemid-member) | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-itemid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-location-member) | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-location-member) | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-location-member)<br>(Solicitação de Reunião) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-normalizedsubject-member) | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-normalizedsubject-member) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) | [Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-optionalattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-organizer-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) | [Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-requiredattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-sender-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| iniciar | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-start-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-start-member) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-start-member)<br>(Solicitação de Reunião) | Data | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| assunto | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-subject-member) | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-subject-member) | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-subject-member) | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-subject-member) | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| para | ReadItem | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-to-member) | [Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-to-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Detalhes por modo | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches() | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participante do Compromisso](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Mensagem lida](/javascript/api/outlook/office.messageread?view=outlook-js-1.3&preserve-view=true#outlook-office-messageread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organizador de Compromissos](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.3&preserve-view=true#outlook-office-appointmentcompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composição da mensagem](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.3&preserve-view=true#outlook-office-messagecompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

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

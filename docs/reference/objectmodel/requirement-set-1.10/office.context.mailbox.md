---
title: Office.context.mailbox - conjunto de requisitos 1.10
description: Outlook Conjunto de requisitos da API de Caixa de Correio versão 1.10 do modelo de objeto Mailbox.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: b5f55e56e846f93910e47e3d2bc5d269011c2a98
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592021"
---
# <a name="mailbox-requirement-set-110"></a>mailbox (conjunto de requisitos 1.10)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](../../../outlook/understanding-outlook-add-in-permissions.md)| Restrito|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Minimum<br>nível de permissão | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|---|:---:|
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#diagnostics) | ReadItem | Escrever<br>Ler | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#ewsurl) | ReadItem | Escrever<br>Ler | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Ler | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#mastercategories) | ReadWriteMailbox | Escrever<br>Ler | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.10&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#resturl) | ReadItem | Escrever<br>Ler | Cadeia de caracteres | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#userprofile) | ReadItem | Escrever<br>Ler | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Modos | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#converttoewsid-itemid--restversion-) | Restricted | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#converttorestid-itemid--restversion-) | Restricted | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#converttoutcclienttime-input-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayappointmentform-itemid-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayappointmentform-itemid--options--callback-) | ReadItem | Escrever<br>Ler | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displaymessageform-itemid-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displaymessageform-itemid--options--callback-) | ReadItem | Escrever<br>Ler | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displaynewappointmentform-parameters--options--callback-) | ReadItem | Ler | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | Ler | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displaynewmessageform-parameters--options--callback-) | ReadItem | Ler | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Eventos

Inscreva-se e cancele a assinatura dos seguintes eventos usando [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) e [removeHandlerAsync,](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) respectivamente.

> [!IMPORTANT]
> Os eventos só estão disponíveis com a implementação do painel de tarefas.

| Evento | Descrição | Minimum<br>conjunto de requisitos |
|---|---|:---:|
|`ItemChanged`| Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
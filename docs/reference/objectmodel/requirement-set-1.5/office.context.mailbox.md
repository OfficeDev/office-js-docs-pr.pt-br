---
title: 'Office.context.mailbox: conjunto de requisitos da versão 1.5'
description: Conjunto de requisitos da API de caixa de correio do Outlook versão 1,5 do modelo de objeto Mailbox.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 9acbaceb8a3928c2b80f8e55caef2243dbb28df8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608660"
---
# <a name="mailbox-requirement-set-15"></a>caixa de correio (conjunto de requisitos 1,5)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](../../../outlook/understanding-outlook-add-in-permissions.md)| Restrito|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Mínimo<br>nível de permissão | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|---|:---:|
| [la](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#diagnostics) | ReadItem | Escrever<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#ewsurl) | ReadItem | Escrever<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Read | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.5) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#resturl) | ReadItem | Escrever<br>Read | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#userprofile) | ReadItem | Escrever<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| Método | Mínimo<br>nível de permissão | Modelos | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Escrever<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#converttoewsid-itemid--restversion-) | Restricted | Escrever<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (TimeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#converttolocalclienttime-timevalue-) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#converttorestid-itemid--restversion-) | Restricted | Escrever<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (entrada)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#converttoutcclienttime-input-) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#displayappointmentform-itemid-) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#displaymessageform-itemid-) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#getcallbacktokenasync-options--callback-) | ReadItem | Escrever<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#getcallbacktokenasync-callback--usercontext-) | ReadItem | Escrever<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#removehandlerasync-eventtype--options--callback-) | ReadItem | Escrever<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Eventos

Você pode assinar e cancelar a assinatura dos eventos a seguir usando o [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#addhandlerasync-eventtype--handler--options--callback-) e o [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#removehandlerasync-eventtype--options--callback-) , respectivamente.

| Evento | Descrição | Mínimo<br>conjunto de requisitos |
|---|---|:---:|
|`ItemChanged`| Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

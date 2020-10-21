---
title: Office. Context. Mailbox – conjunto de requisitos 1,9
description: Conjunto de requisitos da API de caixa de correio do Outlook versão 1,9 do modelo de objeto Mailbox.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: be3ce5a8dded96553884582f6e05fda90d612e12
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628035"
---
# <a name="mailbox-requirement-set-19"></a>caixa de correio (conjunto de requisitos 1,9)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](../../../outlook/understanding-outlook-add-in-permissions.md)| Restrito|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Minimum<br>nível de permissão | Modelos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|---|:---:|
| [la](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#diagnostics) | ReadItem | Escrever<br>Leitura | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#ewsurl) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Leitura | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Nova mastercategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#mastercategories) | ReadWriteMailbox | Escrever<br>Leitura | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.9&preserve-view=true) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#resturl) | ReadItem | Escrever<br>Leitura | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#userprofile) | ReadItem | Escrever<br>Leitura | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Modelos | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Escrever<br>Leitura | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#converttoewsid-itemid--restversion-) | Restricted | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (TimeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#converttorestid-itemid--restversion-) | Restricted | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (entrada)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#converttoutcclienttime-input-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentform-itemid-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync (itemId, [opções], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentform-itemid--options--callback-) | ReadItem | Escrever<br>Leitura | [1,9](outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageform-itemid-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync (itemId, [opções], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageform-itemid--options--callback-) | ReadItem | Escrever<br>Leitura | [1,9](outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync (parâmetros, [opções], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentform-parameters--options--callback-) | ReadItem | Leitura | [1,9](outlook-requirement-set-1.9.md) |
| [displayNewMessageForm (parâmetros)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | Leitura | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync (parâmetros, [opções], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageform-parameters--options--callback-) | ReadItem | Leitura | [1,9](outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | Escrever<br>Leitura | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | Escrever<br>Leitura | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Eventos

Você pode assinar e cancelar a assinatura dos eventos a seguir usando o [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) e o [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#removehandlerasync-eventtype--options--callback-) , respectivamente.

| Evento | Descrição | Minimum<br>conjunto de requisitos |
|---|---|:---:|
|`ItemChanged`| Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

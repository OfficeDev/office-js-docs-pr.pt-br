---
title: Office.context.mailbox - conjunto de requisitos de visualização
description: Outlook Versão do conjunto de requisitos de visualização da API de Caixa de Correio do modelo de objeto mailbox.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 2793491facb48aaf204004ef9f8041d15dd713d1
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893628"
---
# <a name="mailbox-preview-requirement-set"></a>mailbox (conjunto de requisitos de visualização)

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#diagnostics) | ReadItem | Escrever<br>Leitura | [Diagnóstico](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#ewsurl) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Leitura | [Item](/javascript/api/outlook/office.item?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#mastercategories) | ReadWriteMailbox | Escrever<br>Leitura | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#resturl) | ReadItem | Escrever<br>Leitura | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#userprofile) | ReadItem | Escrever<br>Leitura | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Modos | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Escrever<br>Leitura | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#converttoewsid-itemid--restversion-) | Restricted | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#converttorestid-itemid--restversion-) | Restricted | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#converttoutcclienttime-input-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentform-itemid-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentform-itemid--options--callback-) | ReadItem | Escrever<br>Leitura | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageform-itemid-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageform-itemid--options--callback-) | ReadItem | Escrever<br>Leitura | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentform-parameters--options--callback-) | ReadItem | Leitura | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | Leitura | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageform-parameters--options--callback-) | ReadItem | Leitura | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | Escrever<br>Leitura | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | Escrever<br>Leitura | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Eventos

Você pode se inscrever e cancelar a assinatura dos seguintes eventos usando [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) e [removeHandlerAsync,](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#removehandlerasync-eventtype--options--callback-) respectivamente.

> [!IMPORTANT]
> Os eventos só estão disponíveis com a implementação do painel de tarefas.

| [Event](/javascript/api/office/office.eventtype) | Descrição | Minimum<br>conjunto de requisitos |
|---|---|:---:|
|`ItemChanged`| Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
|`OfficeThemeChanged`| O Office tema na caixa de correio foi alterado. | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |

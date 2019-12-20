---
title: Office. Context. Mailbox – conjunto de requisitos 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 3518d3f768157cd8a33976b7a121660bde3154d0
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814714"
---
# <a name="mailbox"></a>mailbox

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Mínimo<br>nível de permissão | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|---|:---:|
| [la](office.context.mailbox.diagnostics.md) | ReadItem | Escrever<br>Leitura | [La](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#ewsurl) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restrito | Escrever<br>Leitura | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#resturl) | ReadItem | Escrever<br>Leitura | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](office.context.mailbox.userProfile.md) | ReadItem | Escrever<br>Leitura | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Mínimo<br>nível de permissão | Modelos | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Escrever<br>Leitura | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttoewsid-itemid--restversion-) | Restrito | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttolocalclienttime-timevalue-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttorestid-itemid--restversion-) | Restrito | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttoutcclienttime-input-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displayappointmentform-itemid-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaymessageform-itemid-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaynewappointmentform-parameters-) | ReadItem | Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaynewmessageform-parameters-) | ReadItem | Escrever<br>Leitura | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getcallbacktokenasync-options--callback-) | ReadItem | Escrever<br>Leitura | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getcallbacktokenasync-callback--usercontext-) | ReadItem | Escrever<br>Leitura | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Escrever<br>Leitura | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#removehandlerasync-eventtype--options--callback-) | ReadItem | Escrever<br>Leitura | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Eventos

Você pode assinar e cancelar a assinatura dos eventos a seguir usando o [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#addhandlerasync-eventtype--handler--options--callback-) e o [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#removehandlerasync-eventtype--options--callback-) , respectivamente.

| Evento | Descrição | Mínimo<br>conjunto de requisitos |
|---|---|:---:|
|`ItemChanged`| Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

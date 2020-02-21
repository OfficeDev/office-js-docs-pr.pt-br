---
title: Office. Context. Mailbox – conjunto de requisitos 1,2
description: ''
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 9d4b30a9b152e755a24788adc779052117492c95
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165451"
---
# <a name="mailbox"></a>mailbox

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
| [la](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#diagnostics) | ReadItem | Escrever<br>Ler | [La](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#ewsurl) | ReadItem | Escrever<br>Ler | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Ler | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#userprofile) | ReadItem | Escrever<br>Ler | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| Método | Mínimo<br>nível de permissão | Modelos | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#converttolocalclienttime-timevalue-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#converttoutcclienttime-input-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#displayappointmentform-itemid-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#displaymessageform-itemid-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#displaynewappointmentform-parameters-) | ReadItem | Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#getcallbacktokenasync-callback--usercontext-) | ReadItem | Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

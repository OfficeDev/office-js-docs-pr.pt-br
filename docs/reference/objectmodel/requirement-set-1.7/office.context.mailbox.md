---
title: Office.context.mailbox - conjunto de requisitos 1.7
description: Outlook Conjunto de requisitos da API de Caixa de Correio versão 1.7 do modelo de objeto mailbox.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3f4b12031611ff5c69d8d4c00474dc77a95f9f0d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151775"
---
# <a name="mailbox-requirement-set-17"></a>mailbox (conjunto de requisitos 1.7)

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#diagnostics) | ReadItem | Escrever<br>Ler | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#ewsUrl) | ReadItem | Escrever<br>Ler | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Ler | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#restUrl) | ReadItem | Escrever<br>Ler | Cadeia de Caracteres | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#userProfile) | ReadItem | Escrever<br>Ler | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Modos | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#convertToEwsId_itemId__restVersion_) | Restricted | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#convertToLocalClientTime_timeValue_) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#convertToRestId_itemId__restVersion_) | Restricted | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#convertToUtcClientTime_input_) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displayAppointmentForm_itemId_) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displayMessageForm_itemId_) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displayNewAppointmentForm_parameters_) | ReadItem | Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displayNewMessageForm_parameters_) | ReadItem | Ler | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#getCallbackTokenAsync_options__callback_) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#getCallbackTokenAsync_callback__userContext_) | ReadItem | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#getUserIdentityTokenAsync_callback__userContext_) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#makeEwsRequestAsync_data__callback__userContext_) | ReadWriteMailbox | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Events

Você pode se inscrever e cancelar a assinatura dos seguintes eventos usando [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) e [removeHandlerAsync,](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#removeHandlerAsync_eventType__options__callback_) respectivamente.

> [!IMPORTANT]
> Os eventos só estão disponíveis com a implementação do painel de tarefas.

| [Event](/javascript/api/office/office.eventtype) | Descrição | Minimum<br>conjunto de requisitos |
|---|---|:---:|
|`ItemChanged`| Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

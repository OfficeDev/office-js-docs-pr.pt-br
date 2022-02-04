---
title: Office.context.mailbox - conjunto de requisitos 1.8
description: Outlook conjunto de requisitos da API de Caixa de Correio 1.8 do modelo de objeto mailbox.
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# <a name="mailbox-requirement-set-18"></a>mailbox (conjunto de requisitos 1.8)

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-diagnostics-member) | ReadItem | Escrever<br>Ler | [Diagnóstico](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-ewsurl-member) | ReadItem | Escrever<br>Ler | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Ler | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-mastercategories-member) | ReadWriteMailbox | Escrever<br>Ler | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-resturl-member) | ReadItem | Escrever<br>Ler | Cadeia de caracteres | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-userprofile-member) | ReadItem | Escrever<br>Ler | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Modos | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-converttoewsid-member(1)) | Restricted | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-converttorestid-member(1)) | Restricted | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | ReadItem | Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-displaynewmessageform-member(1)) | ReadItem | Ler | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(2)) | ReadItem | Escrever<br>Ler | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | ReadItem | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | ReadWriteMailbox | Escrever<br>Ler | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) | ReadItem | Escrever<br>Ler | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Eventos

Você pode se inscrever e cancelar a assinatura dos seguintes eventos usando [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) e [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) , respectivamente.

> [!IMPORTANT]
> Os eventos só estão disponíveis com a implementação do painel de tarefas.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.8&preserve-view=true) | Descrição | Minimum<br>conjunto de requisitos |
|---|---|:---:|
|`ItemChanged`| Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

---
title: Obter e definir dados de item em um formulário de redação no Outlook
description: Obtenha ou defina várias propriedades de um item em um suplemento do Outlook em um cenário de redação, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.
ms.date: 12/10/2019
ms.localizationpriority: medium
ms.openlocfilehash: 03ab6672f8d1dc2bd6dae6fe064da43ee66d9f72
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151863"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Obter e definir dados de item em um formulário de redação no Outlook

Saiba como obter ou definir várias propriedades de um item em um suplemento do Outlook em um cenário de composição, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>Obter e definir propriedades de item de um suplemento de redação

Em um formulário de composição, é possível obter a maioria das propriedades que estão expostas no mesmo tipo de item de um formulário de leitura (por exemplo, participantes, destinatários, assunto e corpo) e acessar algumas propriedades adicionais que são relevantes somente no formulário de composição, mas não de leitura (corpo, cco).

Para a maioria dessas propriedades, como é possível que um suplemento do Outlook e o usuário estejam modificando a mesma propriedade na interface de usuário ao mesmo tempo, os métodos para obtê-las e defini-las é assíncrono. A Tabela 1 lista as propriedades no nível do item e os métodos assíncronos correspondentes para obtê-los e defini-los em um formulário de composição. As propriedades [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [item.conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) são exceções porque os usuários não podem modificá-las. Você pode obtê-las via programação da mesma maneira em um formulário de composição e em um formulário de leitura, diretamente do objeto pai.

Além de acessar propriedades de item na API javaScript Office, você pode acessar propriedades de nível de item usando Exchange Web Services (EWS). Com a permissão **ReadWriteMailbox**, você pode usar o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para acessar as operações [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) dos EWS para obter e definir propriedades de um ou mais itens na caixa de correio do usuário.

A função `makeEwsRequestAsync` está disponível nos formulários de leitura e redação. Para saber mais sobre a permissão **ReadWriteMailbox** e acessar os EWS na plataforma de suplementos do Office, confira [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md) e [Chamar serviços Web de um suplemento do Outlook](web-services.md).

**Tabela 1. Métodos assíncronos para obter ou definir propriedades de item em um formulário de redação**

<br/>

| Propriedade | Tipo de propriedade | Método assíncrono para obter | Método(s) assíncrono(s) para definir |
|:-----|:-----|:-----|:-----|
|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Destinatários](/javascript/api/outlook/office.Recipients)|[Recipients.getAsync](/javascript/api/outlook/office.Recipients#getAsync_options__callback_)|[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addAsync_recipients__options__callback_), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setAsync_recipients__options__callback_)|
|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Body](/javascript/api/outlook/office.Body)|[Body.getAsync](/javascript/api/outlook/office.Body#getAsync_coercionType__options__callback_)|[Body.prependAsync](/javascript/api/outlook/office.Body#prependAsync_data__options__callback_), [Body.setAsync](/javascript/api/outlook/office.Body#setAsync_data__options__callback_), [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setSelectedDataAsync_data__options__callback_)|
|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Time](/javascript/api/outlook/office.Time)|[Time.getAsync](/javascript/api/outlook/office.Time#getAsync_options__callback_)|[Time.setAsync](/javascript/api/outlook/office.Time#setAsync_dateTime__options__callback_)|
|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Location](/javascript/api/outlook/office.Location)|[Location.getAsync](/javascript/api/outlook/office.Location#getAsync_options__callback_)|[Location.setAsync](/javascript/api/outlook/office.Location#setAsync_location__options__callback_)|
|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Hora|Time.getAsync|Time.setAsync|
|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Subject](/javascript/api/outlook/office.Subject)|[Subject.getAsync](/javascript/api/outlook/office.Subject#getAsync_options__callback_)|[Subject.setAsync](/javascript/api/outlook/office.Subject#setAsync_subject__options__callback_)|
|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>Confira também

- [Criar suplementos do Outlook para formulários de redação](compose-scenario.md)
- [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md)
- [Chamar serviços Web de um suplemento do Outlook](web-services.md)
- [Obter e definir dados de item do Outlook em formulários de leitura ou redação](item-data.md)

---
title: Obter e definir dados de item em um formulário de redação no Outlook
description: Obtenha ou defina várias propriedades de um item em um suplemento do Outlook em um cenário de redação, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2ae4b6a30d08199207faf89079c57fbff46d6a0e
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467235"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Obter e definir dados de item em um formulário de redação no Outlook

Saiba como obter ou definir várias propriedades de um item em um suplemento do Outlook em um cenário de composição, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>Obter e definir propriedades de item de um suplemento de redação

Em um formulário de composição, é possível obter a maioria das propriedades que estão expostas no mesmo tipo de item de um formulário de leitura (por exemplo, participantes, destinatários, assunto e corpo) e acessar algumas propriedades adicionais que são relevantes somente no formulário de composição, mas não de leitura (corpo, cco).

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [item.conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

Além de acessar propriedades de item na API JavaScript do Office, você pode acessar propriedades de nível de item usando os Serviços Web do Exchange (EWS). Com a permissão de caixa de correio de leitura **/** gravação, você pode usar o método [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para acessar operações do EWS, [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation), para obter e definir mais propriedades de um item ou itens na caixa de correio do usuário.

O `makeEwsRequestAsync` método está disponível em formulários de composição e leitura. Para obter mais informações sobre a permissão de caixa de correio de leitura **/** gravação e como acessar o EWS por meio da plataforma de Suplementos do Office, consulte Noções básicas sobre permissões de suplemento [do Outlook](understanding-outlook-add-in-permissions.md) e Chamar serviços Web de um suplemento do [Outlook](web-services.md).

**Tabela 1. Métodos assíncronos para obter ou definir propriedades de item em um formulário de redação**

| Propriedade | Tipo de propriedade | Método assíncrono para obter | Métodos assíncronos a serem definidos |
|:-----|:-----|:-----|:-----|
|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Destinatários](/javascript/api/outlook/office.recipients)|[Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))|[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|
|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.body)|[Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)), [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))|
|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Time](/javascript/api/outlook/office.time)|[Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1))|[Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))|
|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Location](/javascript/api/outlook/office.location)|[Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1))|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))|
|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Hora|Time.getAsync|Time.setAsync|
|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Subject](/javascript/api/outlook/office.subject)|[Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1))|[Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))|
|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Destinatários|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>Confira também

- [Criar suplementos do Outlook para formulários de redação](compose-scenario.md)
- [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md)
- [Chamar serviços Web de um suplemento do Outlook](web-services.md)
- [Obter e definir dados de item do Outlook em formulários de leitura ou redação](item-data.md)

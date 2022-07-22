---
title: Obter e definir dados de item em um formulário de redação no Outlook
description: Obtenha ou defina várias propriedades de um item em um suplemento do Outlook em um cenário de redação, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.
ms.date: 12/10/2019
ms.localizationpriority: medium
ms.openlocfilehash: ddc6cd0011060bc49d1fd5cd8e6c9ceebb2a8c08
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958976"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Obter e definir dados de item em um formulário de redação no Outlook

Saiba como obter ou definir várias propriedades de um item em um suplemento do Outlook em um cenário de composição, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>Obter e definir propriedades de item de um suplemento de redação

Em um formulário de composição, é possível obter a maioria das propriedades que estão expostas no mesmo tipo de item de um formulário de leitura (por exemplo, participantes, destinatários, assunto e corpo) e acessar algumas propriedades adicionais que são relevantes somente no formulário de composição, mas não de leitura (corpo, cco).

Para a maioria dessas propriedades, como é possível que um suplemento do Outlook e o usuário estejam modificando a mesma propriedade na interface de usuário ao mesmo tempo, os métodos para obtê-las e defini-las é assíncrono. A Tabela 1 lista as propriedades no nível do item e os métodos assíncronos correspondentes para obtê-los e defini-los em um formulário de composição. As propriedades [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) e [item.conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) são exceções porque os usuários não podem modificá-las. Você pode obtê-las via programação da mesma maneira em um formulário de composição e em um formulário de leitura, diretamente do objeto pai.

Além de acessar propriedades de item na API JavaScript do Office, você pode acessar propriedades de nível de item usando os Serviços Web do Exchange (EWS). Com a permissão **ReadWriteMailbox**, você pode usar o método [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para acessar as operações [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) dos EWS para obter e definir propriedades de um ou mais itens na caixa de correio do usuário.

O `makeEwsRequestAsync` método está disponível em formulários de composição e leitura. Para saber mais sobre a permissão **ReadWriteMailbox** e acessar os EWS na plataforma de suplementos do Office, confira [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md) e [Chamar serviços Web de um suplemento do Outlook](web-services.md).

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

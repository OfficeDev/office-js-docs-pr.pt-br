---
title: Obter ou definir dados de item em um suplemento do Outlook
description: Dependendo da ativação do suplemento ser em um formulário de leitura ou de composição, as propriedades que estão disponíveis para o suplemento no item variam.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: be7d14a6c417d01c0537e3375524da5cc807d749
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165733"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>Obter e configurar dados de item do Outlook em formulários de leitura ou composição

A partir da versão 1.1 do esquema dos manifestos dos suplementos do Office, o Outlook pode ativar suplementos quando o usuário está visualizando ou compondo um item. Dependendo da ativação do suplemento ser em um formulário de leitura ou de composição, as propriedades que estão disponíveis para o suplemento no item também variam.

Por exemplo, as propriedades [dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) são definidas somente para um item que já foi enviado (o item é visualizado em um formulário de leitura), mas não quando o item está sendo criado (em um formulário de composição). Outro exemplo é a propriedade [bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), que só é significativa quando uma mensagem está sendo criada (em um formulário de composição) e não está acessível ao usuário em um formulário de leitura.

## <a name="item-properties-available-in-compose-and-read-forms"></a>Propriedades de item disponíveis nos formulários de leitura e de redação

A Tabela 1 mostra as propriedades a nível de item na API JavaScript para Office que estão disponíveis em cada modo (leitura e redação) de suplementos de email. Normalmente, essas propriedades disponíveis nos formulários de leitura são somente leitura e as disponíveis nos formulários de redação são de leitura/gravação, com exceção das propriedades [itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), [conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), que são sempre somente leitura.

Para as propriedades do nível de item restantes disponíveis em formulários de composição, como o suplemento e o usuário podem estar lendo ou gravando a mesma propriedade ao mesmo tempo, os métodos para obtê-los ou defini-los s no modo redigir são assíncronos e, portanto, o tipo de objeto retornado por essas propriedades também podem ser diferentes entre os formulários de composição e de leitura. Para saber mais sobre como usar métodos assíncronos para obter ou definir propriedades de nível de item no modo de composição, confira [Obter e definir dados de item em um formulário de composição no Outlook](get-and-set-item-data-in-a-compose-form.md).


**Tabela 1. Propriedades de item disponíveis nos formulários de leitura e de redação**

<br/>

|**Tipo de item**|**Propriedade**|**Tipo de propriedade nos formulários de leitura**|**Tipo de propriedade em formulários de redação**|
|:-----|:-----|:-----|:-----|
|Compromissos e mensagens|[dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objeto JavaScript **Date**|Propriedade não disponível|
|Compromissos e mensagens|[dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objeto JavaScript **Date**|Propriedade não disponível|
|Compromissos e mensagens|[itemClass](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|Propriedade não disponível|
|Compromissos e mensagens|[itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|Propriedade não disponível|
|Compromissos e mensagens|[itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Cadeia de caracteres na enumeração [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)|Cadeia de caracteres na enumeração [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) (somente leitura)|
|Compromissos e mensagens|[attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|Propriedade não disponível|
|Compromissos e mensagens|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Body](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|Compromissos e mensagens|[normalizedSubject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|Propriedade não disponível|
|Compromissos e mensagens|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|[Subject](/javascript/api/outlook/office.subject)|
|Compromissos|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objeto JavaScript **Date**|[Time](/javascript/api/outlook/office.time)|
|Compromissos|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|[Location](/javascript/api/outlook/office.location)|
|Compromissos|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Compromissos|[organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizador](/javascript/api/outlook/office.organizer)|
|Compromissos|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Compromissos|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objeto JavaScript **Date**|[Time](/javascript/api/outlook/office.time)|
|Mensagens|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Propriedade não disponível|[Destinatários](/javascript/api/outlook/office.recipients)|
|Mensagens|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinatários](/javascript/api/outlook/office.recipients)|
|Mensagens|[conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|Cadeia de caracteres (somente leitura)|
|Mensagens|[from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[De](/javascript/api/outlook/office.from)|
|Mensagens|[internetMessageId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Inteiro|Propriedade não disponível|
|Mensagens|[sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|Propriedade não disponível|
|Mensagens|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinatários](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>Usar tokens de retorno de chamada do Exchange Server de um suplemento de leitura

Se o suplemento do Outlook é ativado nos formulários de leitura, você pode obter um token de retorno de chamada do Exchange. Esse token pode ser usado no código do lado do servidor para acessar o item completo via EWS (Serviços Web do Exchange).

Ao especificar a permissão **ReadItem** no manifesto do suplemento, você poderá usar o método [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para obter um token de retorno de chamada do Exchange, a propriedade [mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) para obter a URL do ponto de extremidade do EWS para a caixa de correio do usuário e [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para obter a identificação EWS para o item selecionado. Você pode então passar o token de retorno de chamada, a URL de ponto de extremidade de EWS e a ID de item EWS para código do lado do servidor a fim de acessar a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e obter mais propriedades do item.


## <a name="access-ews-from-a-read-or-compose-add-in"></a>Acessar os EWS de um suplemento de leitura ou de redação

Você também pode usar o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para acessar as operações do EWS (Serviços Web do Exchange) [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) diretamente do suplemento. Você pode usar essas operações para obter e definir muitas propriedades de um item especificado. Esse método está disponível para os suplementos do Outlook independentemente de estes serem ativados em formulário de leitura ou de composição, desde que você especifique a permissão **ReadWriteMailbox** no manifesto do suplemento.

Para saber mais sobre o uso de **makeEwsRequestAsync** para acessar as operações EWS, confira [Chamar serviços Web de um suplemento do Outlook](web-services.md).


## <a name="see-also"></a>Confira também

- [Obter e definir dados de item em um formulário de redação no Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Chamar serviços Web de um suplemento do Outlook](web-services.md)

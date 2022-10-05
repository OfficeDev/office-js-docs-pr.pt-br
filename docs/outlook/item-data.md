---
title: Obter ou definir dados de item em um suplemento do Outlook
description: Dependendo da ativação do suplemento ser em um formulário de leitura ou de composição, as propriedades que estão disponíveis para o suplemento no item variam.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8349d81b376aa55d239a88a5d4598381fd8bfc4d
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467269"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>Obter e configurar dados de item do Outlook em formulários de leitura ou composição

A partir da versão 1.1 do esquema dos manifestos dos suplementos do Office, o Outlook pode ativar suplementos quando o usuário está visualizando ou compondo um item. Dependendo da ativação do suplemento ser em um formulário de leitura ou de composição, as propriedades que estão disponíveis para o suplemento no item também variam.

Por exemplo, as propriedades [dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) e [dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) são definidas somente para um item que já foi enviado (o item é visualizado em um formulário de leitura), mas não quando o item está sendo criado (em um formulário de composição). Outro exemplo é a propriedade [bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), que só é significativa quando uma mensagem está sendo criada (em um formulário de composição) e não está acessível ao usuário em um formulário de leitura.

## <a name="item-properties-available-in-compose-and-read-forms"></a>Propriedades de item disponíveis nos formulários de leitura e de redação

A Tabela 1 mostra as propriedades de nível de item na API JavaScript do Office que estão disponíveis em cada modo (leitura e redação) de suplementos de email. Normalmente, essas propriedades disponíveis em formulários de leitura são somente leitura e as disponíveis nos formulários de redação são leitura/gravação, com exceção das propriedades [itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), [conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) e [itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) , que são sempre somente leitura, independentemente.

Para as propriedades do nível de item restantes disponíveis em formulários de composição, como o suplemento e o usuário podem estar lendo ou gravando a mesma propriedade ao mesmo tempo, os métodos para obtê-los ou defini-los s no modo redigir são assíncronos e, portanto, o tipo de objeto retornado por essas propriedades também podem ser diferentes entre os formulários de composição e de leitura. Para saber mais sobre como usar métodos assíncronos para obter ou definir propriedades de nível de item no modo de composição, confira [Obter e definir dados de item em um formulário de composição no Outlook](get-and-set-item-data-in-a-compose-form.md).


**Tabela 1. Propriedades de item disponíveis nos formulários de leitura e de redação**

<br/>

|**Tipo de item**|**Propriedade**|**Tipo de propriedade nos formulários de leitura**|**Tipo de propriedade em formulários de redação**|
|:-----|:-----|:-----|:-----|
|Compromissos e mensagens|[dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objeto JavaScript **Date**|Propriedade não disponível|
|Compromissos e mensagens|[dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objeto JavaScript **Date**|Propriedade não disponível|
|Compromissos e mensagens|[itemClass](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Propriedade não disponível|
|Compromissos e mensagens|[itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Propriedade não disponível|
|Compromissos e mensagens|[itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Cadeia de caracteres na enumeração [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)|Cadeia de caracteres na enumeração [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) (somente leitura)|
|Compromissos e mensagens|[attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|Propriedade não disponível|
|Compromissos e mensagens|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|Compromissos e mensagens|[normalizedSubject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Propriedade não disponível|
|Compromissos e mensagens|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|[Subject](/javascript/api/outlook/office.subject)|
|Compromissos|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objeto JavaScript **Date**|[Time](/javascript/api/outlook/office.time)|
|Compromissos|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|[Location](/javascript/api/outlook/office.location)|
|Compromissos|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Compromissos|[organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizador](/javascript/api/outlook/office.organizer)|
|Compromissos|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Compromissos|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objeto JavaScript **Date**|[Time](/javascript/api/outlook/office.time)|
|Mensagens|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Propriedade não disponível|[Destinatários](/javascript/api/outlook/office.recipients)|
|Mensagens|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinatários](/javascript/api/outlook/office.recipients)|
|Mensagens|[conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Cadeia de caracteres (somente leitura)|
|Mensagens|[from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[De](/javascript/api/outlook/office.from)|
|Mensagens|[internetMessageId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Inteiro|Propriedade não disponível|
|Mensagens|[sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|Propriedade não disponível|
|Mensagens|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinatários](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>Usar tokens de retorno de chamada do Exchange Server de um suplemento de leitura

Se o suplemento do Outlook é ativado nos formulários de leitura, você pode obter um token de retorno de chamada do Exchange. Esse token pode ser usado no código do lado do servidor para acessar o item completo via EWS (Serviços Web do Exchange).

Ao especificar a permissão de [item](understanding-outlook-add-in-permissions.md#read-item-permission) de leitura no manifesto do suplemento, você pode usar o método [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para obter um token de retorno de chamada do Exchange, a propriedade [mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) para obter a URL do ponto de extremidade EWS para a caixa de correio do usuário e [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) para obter a ID do EWS para o item selecionado. Você pode então passar o token de retorno de chamada, a URL de ponto de extremidade de EWS e a ID de item EWS para código do lado do servidor a fim de acessar a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e obter mais propriedades do item.

## <a name="access-ews-from-a-read-or-compose-add-in"></a>Acessar os EWS de um suplemento de leitura ou de redação

Você também pode usar o método [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para acessar as operações do EWS (Serviços Web do Exchange) [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) diretamente do suplemento. Você pode usar essas operações para obter e definir muitas propriedades de um item especificado. Esse método está disponível para suplementos do Outlook, independentemente de o suplemento ter sido ativado em um formulário de leitura ou redação, desde que você especifique a permissão de caixa de correio de leitura **/** gravação no manifesto do suplemento. Para obter mais informações sobre **a permissão de caixa de correio de leitura/** gravação, consulte [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md)

Para saber mais sobre o uso de **makeEwsRequestAsync** para acessar as operações EWS, confira [Chamar serviços Web de um suplemento do Outlook](web-services.md).


## <a name="see-also"></a>Confira também

- [Obter e definir dados de item em um formulário de redação no Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Chamar serviços Web de um suplemento do Outlook](web-services.md)

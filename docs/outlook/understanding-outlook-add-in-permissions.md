---
title: Noções básicas sobre permissões de suplemento do Outlook
description: Suplementos do Outlook especificam o nível de permissão necessário em seu manifesto que incluem o modo restrito, ReadItem, ReadWriteItem ou ReadWriteMailbox.
ms.date: 02/19/2020
ms.localizationpriority: medium
ms.openlocfilehash: b515ef470331a513d6b57007f372b3e4dec1d25b
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660225"
---
# <a name="understanding-outlook-add-in-permissions"></a>Noções básicas sobre permissões de suplemento do Outlook

Os suplementos do Outlook especificam o nível de permissão necessário nos seus manifestos. Os níveis disponíveis são **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**. Esses níveis de permissões são cumulativos: **Restricted** é o nível mais baixo, e cada nível mais alto inclui as permissões dos níveis mais baixos. **ReadWriteMailbox** inclui todas as permissões com suporte.

Você pode ver as permissões solicitadas por um suplemento de email antes de instalá-lo da [AppSource](https://appsource.microsoft.com). Também pode ver as permissões necessárias de suplementos instalados no Centro de Administração do Exchange.

## <a name="restricted-permission"></a>Permissão restrita

A permissão **Restricted** é o nível mais básico de permissão. Especifique a **Restricted** no elemento [Permissions](/javascript/api/manifest/permissions), no manifesto, para solicitar essa permissão. O Outlook atribui essa permissão a um suplemento de email por padrão se o suplemento não solicitar uma permissão específica em seu manifesto.

### <a name="can-do"></a>Pode ser feito

- [Obter somente entidades específicas](match-strings-in-an-item-as-well-known-entities.md) (número de telefone, endereço, URL) do assunto ou corpo do item.

- Especificar uma [regra de ativação ItemIs](activation-rules.md#itemis-rule) que exige que o item atual em um formulário de leitura ou de redação seja um tipo de item específico, ou uma regra [ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md) que corresponde a um subconjunto menor de entidades conhecidas com suporte (número de telefone, endereço, URL) no item selecionado.

- Acessar quaisquer propriedades e métodos que **não** pertencem às informações específicas sobre o usuário ou o item (confira a próxima seção para ver a lista de membros que fazem isso).

### <a name="cant-do"></a>Não pode ser feito

- Use uma [regra ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) no contato, endereço de email, sugestão de reunião ou entidade de sugestão de tarefa.

- Usar a regra [ItemHasAttachment](/javascript/api/manifest/rule#itemhasattachment-rule) ou [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule).

- Acessar os membros na lista a seguir que pertencem às informações do usuário ou do item. A tentativa de acessar os membros nessa lista retorna **null** e resulta em uma mensagem de erro informando que o Outlook que o suplemento de email tenha permissões elevadas.

  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.userProfile](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
  - [Body](/javascript/api/outlook/office.body) e todos os seus membros filhos
  - [Location](/javascript/api/outlook/office.location) e todos os seus membros filhos
  - [Recipients](/javascript/api/outlook/office.recipients) e todos os seus membros filhos
  - [Subject](/javascript/api/outlook/office.subject) e todos os seus membros filhos
  - [Time](/javascript/api/outlook/office.time) e todos os seus membros filhos

## <a name="readitem-permission"></a>Permissão ReadItem

A permissão **ReadItem** é o nível seguinte de permissões no modelo de permissões. **Especifique ReadItem** no **\<Permissions\>** elemento no manifesto para solicitar essa permissão.

### <a name="can-do"></a>Pode ser feito

- [Ler todas as propriedades](item-data.md) do item atual em um formulário de leitura ou [de redação](get-and-set-item-data-in-a-compose-form.md), por exemplo, [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) em um formulário de leitura e [item.to.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) em um formulário de redação.

- [Obter um token de retorno de chamada para obter anexos do item](get-attachments-of-an-outlook-item.md) ou o item completo com os Serviços Web do Exchange (EWS) ou as [APIs REST do Outlook](use-rest-api.md).

- [Gravar propriedades personalizadas](/javascript/api/outlook/office.customproperties) definidas pelo suplemento nesse item.

- [Obter todas as entidades conhecidas existentes](match-strings-in-an-item-as-well-known-entities.md) do assunto ou do corpo do item, e não apenas um subconjunto.

- Usar todas as [entidades conhecidas](activation-rules.md#itemhasknownentity-rule) nas regras [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) ou [expressões regulares](activation-rules.md#itemhasregularexpressionmatch-rule) nas regras [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule). O exemplo a seguir segue a versão 1.1 do esquema. Ele mostra uma regra que ativa o suplemento se uma ou mais das entidades conhecidas forem encontradas no assunto ou no corpo da mensagem selecionada.

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a>Não pode ser feito

- Usar o token fornecido por **mailbox.getCallbackTokenAsync** para:
  - Atualizar ou excluir o item atual usando a API REST do Outlook ou acessar outros itens na caixa de correio do usuário.
  - Obter o item de evento de calendário atual usando a API REST do Outlook.

- Use qualquer uma das APIs a seguir.
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.bcc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.bcc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))
  - [item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))
  - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))
  - [item.cc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.cc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.end.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))
  - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.start.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))
  - [item.to.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.to.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))

## <a name="readwriteitem-permission"></a>Permissão ReadWriteItem

**Especifique ReadWriteItem** no elemento **\<Permissions\>** no manifesto para solicitar essa permissão. Os suplementos de email ativados nos formulários de redação que usam métodos de gravação (**Message.to.addAsync** ou **Message.to.setAsync**) devem usar pelo menos esse nível de permissão.

### <a name="can-do"></a>Pode ser feito

- [Ler e gravar todas as propriedades no nível do item](item-data.md) que está sendo visualizado ou redigido no Outlook.

- [Adicionar ou remover anexos](add-and-remove-attachments-to-an-item-in-a-compose-form.md) desse item.

- Use todos os outros membros da API JavaScript do Office aplicáveis aos suplementos de email, exceto **Mailbox.makeEWSRequestAsync**.

### <a name="cant-do"></a>Não pode ser feito

- Usar o token fornecido por **mailbox.getCallbackTokenAsync** para:
  - Atualizar ou excluir o item atual usando a API REST do Outlook ou acessar outros itens na caixa de correio do usuário.
  - Obter o item de evento de calendário atual usando a API REST do Outlook.

- Usar **mailbox.makeEWSRequestAsync**.

## <a name="readwritemailbox-permission"></a>Permissão ReadWriteMailbox

A permissão **ReadWriteMailbox** é o mais alto nível de permissão. **Especifique ReadWriteMailbox** no **\<Permissions\>** elemento no manifesto para solicitar essa permissão.

Além do suporte que a permissão **ReadWriteItem** oferece, o token fornecido pela **mailbox.getCallbackTokenAsync** fornece acesso para usar as operações dos Serviços Web do Exchange (EWS) ou as APIs REST do Outlook para fazer o seguinte:

- Ler e gravar todas as propriedades de qualquer item na caixa de correio do usuário.
- Criar, ler e gravar em qualquer pasta ou item nessa caixa de correio.
- Enviar um item dessa caixa de correio.

Por **meio de mailbox.makeEWSRequestAsync**, você pode acessar as seguintes operações de EWS.

- [CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)
- [CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)
- [CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)
- [FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)
- [FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)
- [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)
- [GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)
- [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)
- [MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)
- [SendItem](/exchange/client-developer/web-service-reference/senditem-operation)
- [UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)

A tentativa de usar uma operação sem suporte resulta em uma resposta de erro.

## <a name="see-also"></a>Confira também

- [Privacidade, permissões e segurança de suplementos do Outlook](../concepts/privacy-and-security.md)
- [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md)

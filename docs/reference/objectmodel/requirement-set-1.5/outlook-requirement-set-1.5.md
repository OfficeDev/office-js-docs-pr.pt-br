---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 65453a0cff8db682f5f573c25a9afa4e9ff63f67
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163735"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.5

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-15"></a>Novidades na versão 1.5?

O conjunto de requisitos 1.5 inclui todos os recursos do [Conjunto de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) e contém os seguintes recursos adicionais.

- Adicionado suporte para [painéis de tarefas fixáveis](../../../outlook/pinnable-taskpane.md).
- Adicionado suporte para chamar [APIs REST](../../../outlook/use-rest-api.md).
- Adicionada a capacidade de marcar um anexo como embutido.
- Adicionada a capacidade de fechar um painel de tarefas ou uma caixa de diálogo.

### <a name="change-log"></a>Log de alterações

- Adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): adiciona um manipulador de eventos para um evento compatível.
- Foi adicionado o [Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#methods): remove os manipuladores de eventos para um tipo de evento suportado.
- Adicionado o [Office.EventType](office.md#eventtype-string): especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.
- Adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): obtém a URL do ponto de extremidade REST para esta conta de email.
- Modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.
- Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): um novo valor no dicionário `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): um novo valor no dicionário `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

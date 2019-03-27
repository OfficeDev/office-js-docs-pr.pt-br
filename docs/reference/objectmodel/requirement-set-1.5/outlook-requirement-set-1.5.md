---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0489e4efa763b3963fcdc78ec894db46fa06362
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870860"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.5

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-15"></a>Novidades na versão 1.5?

O conjunto de requisitos 1.5 inclui todos os recursos do [Conjunto de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) e contém os seguintes recursos adicionais.

- Adicionado suporte para [painéis de tarefas fixáveis](/outlook/add-ins/pinnable-taskpane).
- Adicionado suporte para chamar [APIs REST](/outlook/add-ins/use-rest-api).
- Adicionada a capacidade de marcar um anexo como embutido.
- Adicionada a capacidade de fechar um painel de tarefas ou uma caixa de diálogo.

### <a name="change-log"></a>Log de alterações

- Adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): adiciona um manipulador de eventos para um evento compatível.
- Foi adicionado o [Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): remove os manipuladores de eventos para um tipo de evento suportado.
- Adicionado o [Office.EventType](office.md#eventtype-string): especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.
- Adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): obtém a URL do ponto de extremidade REST para esta conta de email.
- Modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.
- Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): um novo valor no dicionário `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): um novo valor no dicionário `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/outlook/add-ins/quick-start)

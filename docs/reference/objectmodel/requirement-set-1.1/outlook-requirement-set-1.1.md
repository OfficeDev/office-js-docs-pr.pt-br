---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.1
description: Recursos e APIs que foram introduzidos para suplementos do Outlook e APIs JavaScript do Office como parte da API de caixa de correio 1,1.
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: a6d2d352b2882bf0e5de994c8924bbb99ebb9dfb
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610814"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.1

O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook. API JavaScript do Outlook 1,1 (caixa de correio 1,1) é a primeira versão da API.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o mais recente.

## <a name="whats-new-in-11"></a>Novidades na versão 1.1?

O conjunto de requisitos 1,1 inclui todos os [conjuntos de requisitos de API comuns](../../requirement-sets/office-add-in-requirement-sets.md) com suporte no Outlook. Ele adicionou a capacidade de os suplementos para acessarem o corpo de mensagens e os compromissos e a capacidade de modificar o item atual.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o objeto [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1): Fornece métodos para adicionar e atualizar o conteúdo de um item em um suplemento do Outlook.
- Foi adicionado o objeto [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1): Fornece métodos para obter e definir o local de uma reunião em um suplemento do Outlook.
- Foi adicionado o objeto [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1): Fornece métodos para obter e definir os destinatários de um compromisso ou uma mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1): Fornece métodos para obter e definir o assunto de um compromisso ou uma mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1): Fornece métodos para obter e definir o tempo de início ou fim de uma reunião em um suplemento do Outlook.
- Foi adicionado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.
- Foi adicionado o [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods): Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.
- Foi adicionado o [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods): Remove um anexo de uma mensagem ou de um compromisso.
- Foi adicionado o [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties): Obtém um objeto que fornece métodos para manipular o corpo de um item.
- Foi adicionada a linha [Office. Context. Mailbox. Item. Bcc](office.context.mailbox.item.md#properties) de uma mensagem.
- Adicionado o [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1): especifica o tipo de destinatário para um compromisso.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.1
description: Recursos e APIs que foram introduzidos para Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.1.
ms.date: 12/17/2019
ms.localizationpriority: medium
ms.openlocfilehash: 74a6b2561cf7d0ec28d97810d7337fe909a30837
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746675"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.1

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário. Outlook API JavaScript 1.1 (Caixa de Correio 1.1) é a primeira versão da API.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o mais recente.

## <a name="whats-new-in-11"></a>Novidades na versão 1.1?

O conjunto de requisitos 1.1 inclui todos os conjuntos de requisitos [da API](../../requirement-sets/office-add-in-requirement-sets.md) comum com suporte em Outlook. Ele adicionou a capacidade de os suplementos para acessarem o corpo de mensagens e os compromissos e a capacidade de modificar o item atual.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o objeto [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true): Fornece métodos para adicionar e atualizar o conteúdo de um item em um suplemento do Outlook.
- Foi adicionado o objeto [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true): Fornece métodos para obter e definir o local de uma reunião em um suplemento do Outlook.
- Foi adicionado o objeto [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true): Fornece métodos para obter e definir os destinatários de um compromisso ou uma mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true): Fornece métodos para obter e definir o assunto de um compromisso ou uma mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true): Fornece métodos para obter e definir o tempo de início ou fim de uma reunião em um suplemento do Outlook.
- Foi adicionado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.
- Foi adicionado o [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods): Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.
- Foi adicionado o [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods): Remove um anexo de uma mensagem ou de um compromisso.
- Foi adicionado o [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties): Obtém um objeto que fornece métodos para manipular o corpo de um item.
- Adicionada [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) linha de uma mensagem.
- Adicionado o [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true): especifica o tipo de destinatário para um compromisso.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

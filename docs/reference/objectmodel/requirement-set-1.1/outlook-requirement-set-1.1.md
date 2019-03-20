---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.1
description: ''
ms.date: 10/11/2018
localization_priority: Normal
ms.openlocfilehash: a074d0e38f8d872f0d75a68851aef947989c625e
ms.sourcegitcommit: c4d6ecdc41ea67291b6d155c3b246e31ec2e38b7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/19/2019
ms.locfileid: "30600253"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.1

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o mais recente. 

## <a name="whats-new-in-11"></a>Novidades na versão 1.1?

O conjunto de requisitos 1.1 inclui todos os recursos do Conjunto de requisitos 1.0. Ele adicionou a capacidade de os suplementos para acessarem o corpo de mensagens e os compromissos e a capacidade de modificar o item atual.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o objeto [Body](/javascript/api/outlook_1_1/office.body): Fornece métodos para adicionar e atualizar o conteúdo de um item em um suplemento do Outlook.
- Foi adicionado o objeto [Location](/javascript/api/outlook_1_1/office.location): Fornece métodos para obter e definir o local de uma reunião em um suplemento do Outlook.
- Foi adicionado o objeto [Recipients](/javascript/api/outlook_1_1/office.recipients): Fornece métodos para obter e definir os destinatários de um compromisso ou uma mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Subject](/javascript/api/outlook_1_1/office.subject): Fornece métodos para obter e definir o assunto de um compromisso ou uma mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Time](/javascript/api/outlook_1_1/office.time): Fornece métodos para obter e definir o tempo de início ou fim de uma reunião em um suplemento do Outlook.
- Foi adicionado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.
- Foi adicionado o [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.
- Foi adicionado o [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Remove um anexo de uma mensagem ou de um compromisso.
- Foi adicionado o [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Obtém um objeto que fornece métodos para manipular o corpo de um item.
- Foi adicionada a linha [Office. Context. Mailbox. Item. Bcc](office.context.mailbox.item.md#bcc-recipients) de uma mensagem.
- Adicionado o [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype): especifica o tipo de destinatário para um compromisso.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)

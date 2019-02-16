---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 1767b1b93f13de2c8a0731d2f08a1141b709b734
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068025"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.2

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente. 

## <a name="whats-new-in-12"></a>Novidades na versão 1.2?

O conjunto de requisitos 1.2 inclui todos os recursos do [Conjunto de requisitos 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Ele adicionou a capacidade de os suplementos inserirem texto no cursor do usuário, no assunto ou no corpo da mensagem.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Retorna de forma assíncrona os dados selecionados no corpo ou no assunto de uma mensagem.
- Foi adicionado o [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.
- Foi modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)

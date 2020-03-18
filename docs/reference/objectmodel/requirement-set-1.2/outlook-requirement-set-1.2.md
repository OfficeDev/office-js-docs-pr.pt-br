---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.2
description: Recursos e APIs que foram introduzidos para suplementos do Outlook e APIs JavaScript do Office como parte da API de caixa de correio 1,2.
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: e1605bb2a0d8cc7de0562833cf9cafc9fd932ad4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717779"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.2

O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-12"></a>Novidades na versão 1.2?

O conjunto de requisitos 1.2 inclui todos os recursos do [Conjunto de requisitos 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Ele adicionou a capacidade de os suplementos inserirem texto no cursor do usuário, no assunto ou no corpo da mensagem.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Retorna de forma assíncrona os dados selecionados no corpo ou no assunto de uma mensagem.
- Foi adicionado o [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.
- Foi modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

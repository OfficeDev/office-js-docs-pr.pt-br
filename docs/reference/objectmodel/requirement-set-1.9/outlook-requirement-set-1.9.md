---
title: Conjunto de requisitos de API de suplemento do Outlook 1,9
description: Conjunto de requisitos 1,9 para a API do suplemento do Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: b2174052a60580a895ef82a4b5d8f00ed6899feb
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628038"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Conjunto de requisitos de API de suplemento do Outlook 1,9

O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.

## <a name="whats-new-in-19"></a>Quais são as novidades no 1,9?

O conjunto de requisitos 1,9 inclui todos os recursos do [conjunto de requisitos 1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs para Append-on-Send, propriedades personalizadas e recursos de formulário de exibição.
- Adicionado suporte para `Dialog.messageChild` .

### <a name="change-log"></a>Log de mudanças

- Foi adicionado [CustomProperties. getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--): adiciona uma nova função ao `CustomProperties` objeto que obtém todas as propriedades personalizadas.
- Foi adicionada a [Dialog. messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): Adiciona um novo método que entrega uma mensagem da página host, como um painel de tarefas ou um arquivo de função sem interface do usuário, a uma caixa de diálogo que foi aberta na página.
- Adicionou o [elemento de manifesto ExtendedPermissions](../../manifest/extendedpermissions.md): Adiciona um elemento filho ao elemento de manifesto [VersionOverrides](../../manifest/versionoverrides.md) . Para que um suplemento dê suporte ao [recurso Append-on-Send](../../../outlook/append-on-send.md), a `AppendOnSend` permissão estendida deve ser incluída na coleção de permissões estendidas.
- Foi adicionado o [Office. Context. Mailbox. displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-): adiciona uma nova função ao `Mailbox` objeto que exibe um compromisso existente. Esta é a versão assíncrona do `displayAppointmentForm` método.
- Foi adicionado o [Office. Context. Mailbox. displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-): adiciona uma nova função ao `Mailbox` objeto que exibe uma mensagem existente. Esta é a versão assíncrona do `displayMessageForm` método.
- Foi adicionado o [Office. Context. Mailbox. displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-): adiciona uma nova função ao `Mailbox` objeto que exibe um novo formulário de compromisso. Esta é a versão assíncrona do `displayNewAppointmentForm` método.
- Foi adicionado o [Office. Context. Mailbox. displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-): adiciona uma nova função ao `Mailbox` objeto que exibe um novo formulário de mensagem. Esta é a versão assíncrona do `displayNewMessageForm` método.
- Foi adicionado o [Office. Context. Mailbox. Item. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-): adiciona uma nova função ao `Body` objeto que acrescenta dados ao final do corpo do item no modo de composição.
- Foi adicionado o [Office. Context. Mailbox. Item. displayReplyAllFormAsync](office.context.mailbox.item.md#methods): adiciona uma nova função ao `Item` objeto que exibe o formulário "responder a todos" no modo de leitura. Esta é a versão assíncrona do `displayReplyAllForm` método.
- Foi adicionado o [Office. Context. Mailbox. Item. displayReplyFormAsync](office.context.mailbox.item.md#methods): adiciona uma nova função ao `Item` objeto que exibe o formulário "responder" no modo de leitura. Esta é a versão assíncrona do `displayReplyForm` método.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

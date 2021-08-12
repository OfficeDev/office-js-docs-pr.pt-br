---
title: Outlook conjunto de requisitos de API de complemento 1.9
description: Conjunto de requisitos 1.9 para Outlook api de complemento.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 6ac4e119fea41a6f4bd1b3ab0bfe79f289278f3badeb5842fd895c8635d7f7b4
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087658"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook conjunto de requisitos de API de complemento 1.9

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o mais recente.

## <a name="whats-new-in-19"></a>Novidades no 1.9?

O conjunto de requisitos 1.9 inclui todos os recursos do conjunto [de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs para append-on-send, propriedades personalizadas e recursos de formulário de exibição.
- Adicionado suporte para `Dialog.messageChild` .

### <a name="change-log"></a>Log de mudanças

- Adicionado [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getAll__): adiciona uma nova função ao `CustomProperties` objeto que obtém todas as propriedades personalizadas.
- Adicionado [Dialog.messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): adiciona um novo método que fornece uma mensagem da página host, como um painel de tarefas ou um arquivo de função sem interface do usuário, a uma caixa de diálogo que foi aberta na página.
- Adicionado [elemento de manifesto ExtendedPermissions](../../manifest/extendedpermissions.md): adiciona um elemento filho ao elemento de manifesto [VersionOverrides.](../../manifest/versionoverrides.md) Para que um add-in suporte ao recurso [append-on-send](../../../outlook/append-on-send.md), a permissão estendida deve ser incluída na `AppendOnSend` coleção de permissões estendidas.
- Adicionado [Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayAppointmentFormAsync_itemId__options__callback_): adiciona uma nova função ao objeto que exibe um `Mailbox` compromisso existente. Esta é a versão assíncrona do `displayAppointmentForm` método.
- Adicionado [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayMessageFormAsync_itemId__options__callback_): adiciona uma nova função ao objeto que exibe uma `Mailbox` mensagem existente. Esta é a versão assíncrona do `displayMessageForm` método.
- Adicionado [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayNewAppointmentFormAsync_parameters__options__callback_): adiciona uma nova função ao objeto que exibe `Mailbox` um novo formulário de compromisso. Esta é a versão assíncrona do `displayNewAppointmentForm` método.
- Adicionado [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayNewMessageFormAsync_parameters__options__callback_): adiciona uma nova função ao objeto que exibe `Mailbox` um novo formulário de mensagem. Esta é a versão assíncrona do `displayNewMessageForm` método.
- Adicionado [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendOnSendAsync_data__options__callback_): adiciona uma nova função ao objeto que acrescenta dados ao final do corpo do item no modo `Body` Redação.
- Adicionado [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods): adiciona uma nova função ao objeto que exibe o formulário "Responder a todos" no modo `Item` De leitura. Esta é a versão assíncrona do `displayReplyAllForm` método.
- Adicionado [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods): adiciona uma nova função ao objeto que exibe o formulário "Reply" no modo `Item` De leitura. Esta é a versão assíncrona do `displayReplyForm` método.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

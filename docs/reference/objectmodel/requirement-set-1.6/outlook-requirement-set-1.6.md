---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.6
description: Recursos e APIs que foram introduzidos para os Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.6.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: a552c362e247da7b36d14a0c32f557440a324977
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671734"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.6

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação se aplica a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-16"></a>Novidades na versão 1.6

O conjunto de requisitos 1.6 inclui todos os recursos do conjunto de requisitos [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs para suplementos contextuais para obter a correspondência de entidade ou regex que o usuário selecionou para ativar o suplemento.
- Adicionada uma nova API para abrir um formulário de nova mensagem.
- Adicionada a capacidade de o suplemento determinar o tipo de conta da caixa de correio do usuário.

### <a name="change-log"></a>Log de alterações

- Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada selecionada por um usuário. As correspondências realçadas aplicam-se aos suplementos contextuais.
- Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): adiciona uma nova função que retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se aos suplementos contextuais.
- Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): adiciona uma nova função que abre um novo formulário de mensagem.
- Adicionado [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#accountType): adiciona um novo membro ao perfil de usuário, que indica o tipo de conta do usuário.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

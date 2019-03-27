---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e1f920c259ca1ef8a137bab07132b015d9c75d2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871323"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.6

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação se aplica a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-16"></a>Novidades na versão 1.6

O conjunto de requisitos 1.6 inclui todos os recursos do [Conjunto de requisitos 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs para suplementos contextuais para obter a correspondência de entidade ou regex que o usuário selecionou para ativar o suplemento.
- Adicionada uma nova API para abrir um formulário de nova mensagem.
- Adicionada a capacidade de o suplemento determinar o tipo de conta da caixa de correio do usuário.

### <a name="change-log"></a>Log de alterações

- Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada selecionada por um usuário. As correspondências realçadas aplicam-se aos suplementos contextuais.
- Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): adiciona uma nova função que retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se aos suplementos contextuais.
- Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): adiciona uma nova função que abre um novo formulário de mensagem.
- Adicionado [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): adiciona um novo membro ao perfil de usuário, que indica o tipo de conta do usuário.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/outlook/add-ins/quick-start)

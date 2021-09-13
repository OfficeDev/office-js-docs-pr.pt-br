---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.4
description: Recursos e APIs que foram introduzidos para os Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.4.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 18608e5c105e544783a54eee6fc86df0e0619185
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151778"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.4

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-14"></a>Novidades na versão 1.4

O conjunto de requisitos 1.4 inclui todos os recursos do conjunto [de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Adicionou acesso ao namespace `Office.ui`.

### <a name="change-log"></a>Log de alterações

- Adicionado [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_): exibe uma caixa de diálogo em um Office aplicativo.
- Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageParent_message__messageOptions_): fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.
- Foi adicionado o objeto [Dialog](/javascript/api/office/office.dialog): o objeto retornado quando o método [`displayDialogAsync`](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_) é chamado.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)

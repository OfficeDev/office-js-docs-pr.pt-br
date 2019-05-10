---
title: Elemento Supertip no arquivo de manifesto
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659648"
---
# <a name="supertip"></a>Supertip

Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Title](#title) | Sim | O texto da superdica. |
| [Descrição](#description) | Sim | A descrição da superdica.<br>**Observação**: (Outlook) só há suporte para clientes Windows e Mac. |

### <a name="title"></a>Título

Obrigatório. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).

### <a name="description"></a>Descrição

Obrigatório. A descrição da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **LongStrings** do elemento [Resources](resources.md).

> [!NOTE]
> Para o Outlook, apenas clientes Windows e Mac dão suporte ao elemento **Description** .

## <a name="example"></a>Exemplo

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

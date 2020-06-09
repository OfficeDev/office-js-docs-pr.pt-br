---
title: Elemento Supertip no arquivo de manifesto
description: O elemento Superdica define uma dica de ferramenta rica (título e descrição).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 8061c9dcd7903db0f1265084498d6c86654e1dfa
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608716"
---
# <a name="supertip"></a>Supertip

Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Title](#title) | Sim | O texto da superdica. |
| [Descrição](#description) | Sim | A descrição da superdica.<br>**Observação**: (Outlook) só há suporte para clientes Windows e Mac. |

### <a name="title"></a>Title

Obrigatório. O texto da superdica. O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .

### <a name="description"></a>Descrição

Obrigatório. A descrição da superdica. O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **LongStrings** no elemento [Resources](resources.md) .

> [!NOTE]
> Para o Outlook, apenas clientes Windows e Mac dão suporte ao elemento **Description** .

## <a name="example"></a>Exemplo

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

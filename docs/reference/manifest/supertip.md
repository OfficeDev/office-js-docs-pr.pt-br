---
title: Elemento Supertip no arquivo de manifesto
description: O elemento Supertip define uma dica de ferramenta rica (título e descrição).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 79120cc72aa4804eaaa2330d9298f6521a13552d325d9134814581402ace8210
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093255"
---
# <a name="supertip"></a>Supertip

Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Title](#title) | Sim | O texto da superdica. |
| [Descrição](#description) | Sim | A descrição da superdica.<br>**Observação**: (Outlook) Somente clientes Windows e Mac são suportados. |

### <a name="title"></a>Título

Obrigatório. O texto da superdica. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources.](resources.md)

### <a name="description"></a>Descrição

Obrigatório. A descrição da superdica. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no **elemento LongStrings** no [elemento Resources.](resources.md)

> [!NOTE]
> Para Outlook, somente os clientes Windows e Mac suportam o **elemento Description.**

## <a name="example"></a>Exemplo

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

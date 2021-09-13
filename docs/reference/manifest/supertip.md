---
title: Elemento Supertip no arquivo de manifesto
description: O elemento Supertip define uma dica de ferramenta rica (título e descrição).
ms.date: 05/07/2019
ms.localizationpriority: medium
ms.openlocfilehash: 6c1e73b0aba5923992fba03b78744ae5d34fb5da
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151992"
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

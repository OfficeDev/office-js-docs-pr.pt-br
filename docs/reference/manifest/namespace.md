---
title: Elemento Namespace no arquivo de manifesto
description: O elemento Namespace define o namespace que uma função personalizada usa em Excel.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 3a5afed3d55bde7e9735df534215f96ae1ba7bd3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152152"
---
# <a name="namespace-element"></a>Elemento Namespace

Define o namespace usado por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Não  | Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md). Não pode ter mais de 32 caracteres. |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<Namespace resid="namespace" />
```

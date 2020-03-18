---
title: Elemento Namespace no arquivo de manifesto
description: O elemento namespace define o namespace que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45fd0caa039fdeb885cba4b739750fbd8b642252
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718052"
---
# <a name="namespace-element"></a>Elemento Namespace

Define o namespace usado por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Sim  | Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md). |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<Namespace resid="namespace" />
```

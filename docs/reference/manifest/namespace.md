---
title: Elemento Namespace no arquivo de manifesto
description: O elemento namespace define o namespace que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978666"
---
# <a name="namespace-element"></a>Elemento Namespace

Define o namespace usado por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Não  | Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md). |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<Namespace resid="namespace" />
```

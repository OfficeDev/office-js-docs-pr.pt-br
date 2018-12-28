---
title: Elemento Namespace no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432323"
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

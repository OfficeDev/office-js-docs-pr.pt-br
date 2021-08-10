---
title: Elemento Metadata no arquivo de manifesto
description: O elemento Metadados define as configurações de metadados que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d6b7af8988265baf8fbdea51e1414646f88868ede76ed7194c47db1eb874608d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092484"
---
# <a name="metadata-element"></a>Elemento Metadata

Define as configurações de metadados usados por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|  Elemento  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas. |

## <a name="example"></a>Exemplo

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```

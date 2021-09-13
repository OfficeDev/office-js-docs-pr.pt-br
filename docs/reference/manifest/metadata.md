---
title: Elemento Metadata no arquivo de manifesto
description: O elemento Metadados define as configurações de metadados que uma função personalizada usa no Excel.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6f58b00bb13bde1e2b1742462716119b8b6d369d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151752"
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

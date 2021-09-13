---
title: Elemento Page no arquivo de manifesto
description: O elemento Page define configurações de página HTML que uma função personalizada usa no Excel.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6bde3ba86270874b1d9059b2f1c44952241bf00f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152134"
---
# <a name="page-element"></a>Elemento Page

Define as configurações de página HTML usadas por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|  Elemento  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com o ID de recurso do arquivo HTML usado por funções personalizadas. |

## <a name="example"></a>Exemplo

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```

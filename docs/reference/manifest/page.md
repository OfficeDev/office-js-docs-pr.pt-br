---
title: Elemento Page no arquivo de manifesto
description: O elemento Page define configurações de página HTML que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aa8a2807cbf2549ded680a22b17f24513ea76b9a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937639"
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

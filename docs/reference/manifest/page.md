---
title: Elemento Page no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 83bafd24d0b56322ea5f7d51025f2416be019168
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433729"
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

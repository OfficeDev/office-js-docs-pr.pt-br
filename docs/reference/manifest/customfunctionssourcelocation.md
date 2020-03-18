---
title: Elemento SourceLocation no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 56ebe122853c98a14c52d450bea31fecaefb15d3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720684"
---
# <a name="sourcelocation-element"></a>Elemento SourceLocation

Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.

## <a name="attributes"></a>Atributos

| **Atributo** | **Obrigatório** | **Descrição**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Sim          | O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto. |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<SourceLocation resid="pageURL"/>
```

---
title: Elemento SourceLocation no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432402"
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
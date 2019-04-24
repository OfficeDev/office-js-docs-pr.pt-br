---
title: Elemento SourceLocation no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450685"
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

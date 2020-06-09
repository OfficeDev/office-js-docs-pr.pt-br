---
title: Elemento SourceLocation no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612309"
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

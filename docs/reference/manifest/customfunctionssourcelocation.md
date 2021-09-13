---
title: Elemento SourceLocation para funções personalizadas no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 08/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: 84d5607fbb02c1925137e1a143b7715c7c87c6fa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149071"
---
# <a name="sourcelocation-element-custom-functions"></a>Elemento SourceLocation (funções personalizadas)

Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.

## <a name="attributes"></a>Atributos

| Atributo | Obrigatório | Descrição                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Sim      | O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto. Não pode ter mais de 32 caracteres. |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<SourceLocation resid="pageURL"/>
```

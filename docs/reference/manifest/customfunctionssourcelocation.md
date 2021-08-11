---
title: Elemento SourceLocation para funções personalizadas no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: b18a340d4dd4403b1e5fd2c7d8868a820eef5a241ac3d666926d8f2cb49fcc09
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098296"
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

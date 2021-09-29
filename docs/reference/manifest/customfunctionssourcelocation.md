---
title: Elemento SourceLocation para funções personalizadas no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5f2d881f31f4e46e7f5bb8ab30d78abd0e9b7200
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990681"
---
# <a name="sourcelocation-element-custom-functions"></a>Elemento SourceLocation (funções personalizadas)

Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.

**Tipo de complemento:** Função personalizada

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

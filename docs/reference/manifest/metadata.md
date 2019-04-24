---
title: Elemento Metadata no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a3aecb1983905658f3a55fdb8bf0629a8d5ef474
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452043"
---
# <a name="metadata-element"></a>Elemento Metadata

Define as configurações de metadados usados por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|  Elemento  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas. |

## <a name="example"></a>Exemplo

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```

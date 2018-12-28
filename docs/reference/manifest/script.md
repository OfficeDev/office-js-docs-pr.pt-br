---
title: Elemento Script no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 95e4cbadc35302b4f76108e0ff2a51d31ca89aac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433135"
---
# <a name="script-element"></a>Elemento Script

Define as configurações de script usadas por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|Elementos  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com o ID de recurso do arquivo JavaScript usado por funções personalizadas.|

## <a name="example"></a>Exemplo

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```

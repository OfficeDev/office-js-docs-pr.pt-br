---
title: Elemento Script no arquivo de manifesto
description: O elemento Script define as configurações de script que uma função personalizada usa no Excel.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 259976f752cf3fca72c5012bedd92b9bf021f6aa
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990667"
---
# <a name="script-element"></a>Elemento Script

Define as configurações de script usadas por uma função personalizada no Excel.

**Tipo de complemento:** Função personalizada

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

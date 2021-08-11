---
title: Elemento Script no arquivo de manifesto
description: O elemento Script define as configurações de script que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51902864081e135faed778de1bc6fdee15d67490de8eabc9febf493cb0c09889
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095038"
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

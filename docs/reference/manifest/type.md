---
title: Elemento Type no arquivo de manifesto
description: O elemento Type Especifica se o suplemento equivalente é um suplemento de COM ou um XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 9eeab172ed4ebf06fc93e42f56f8d33f5e7a92db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720313"
---
# <a name="type-element"></a>Elemento Type

Especifica se o suplemento equivalente é um suplemento de COM ou um XLL.

**Tipo de suplemento:** Painel de tarefas, função personalizada

## <a name="syntax"></a>Sintaxe

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Contido em

[EquivalentAdd-in](equivalentaddin.md)

## <a name="add-in-type-values"></a>Valores de tipo de suplemento

Você deve especificar um dos seguintes valores para o `Type` elemento.

- COM: especifica o suplemento equivalente é um suplemento de COM.
- XLL: especifica o suplemento equivalente é um XLL do Excel.

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Tornar seu suplemento do Excel compatível com um suplemento de COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
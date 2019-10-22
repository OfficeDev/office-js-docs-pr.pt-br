---
title: Elemento Type no arquivo de manifesto
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628225"
---
# <a name="type-element"></a>Elemento Type

Especifica se o suplemento equivalente é um suplemento COM ou um XLL.

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
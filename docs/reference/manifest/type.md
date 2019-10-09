---
title: Elemento Type no arquivo de manifesto
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356841"
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

- [Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Tornar o suplemento do Office compatível com um suplemento de COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
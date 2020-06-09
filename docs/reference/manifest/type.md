---
title: Elemento Type no arquivo de manifesto
description: O elemento Type Especifica se o suplemento equivalente é um suplemento de COM ou um XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: b59f903af39facd7543e7384189817d5365cf8c9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604556"
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
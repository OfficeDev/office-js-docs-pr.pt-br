---
title: Elemento Type no arquivo de manifesto
description: O elemento Type especifica se o complemento equivalente é um complemento COM ou um XLL.
ms.date: 03/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 65860ff7aa3e241c227f96c8a8e7c71d7799a04c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151989"
---
# <a name="type-element"></a>Elemento Type

Especifica se o complemento equivalente é um complemento COM ou um XLL.

**Tipo de complemento:** Painel de tarefas, função Personalizada

## <a name="syntax"></a>Sintaxe

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Contido em

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>Valores de tipo de complemento

Você deve especificar um dos seguintes valores para o `Type` elemento.

- COM: Especifica que o complemento equivalente é um complemento COM.
- XLL: especifica que o complemento equivalente é um Excel XLL.

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Torne o seu suplemento do Office compatível com um suplemento COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
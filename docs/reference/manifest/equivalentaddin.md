---
title: Elemento EquivalentAddin no arquivo de manifesto
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356832"
---
# <a name="equivalentaddin-element"></a>Elemento EquivalentAddin

Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.

**Tipo de suplemento:** Painel de tarefas, função personalizada

## <a name="syntax"></a>Sintaxe

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Contido em

[EquivalentAdd-ins](equivalentaddins.md)

## <a name="must-contain"></a>Deve conter

[Type](type.md)

## <a name="can-contain"></a>Pode conter

[](progid.md)
[Nome de arquivo](filename.md) ProgID

## <a name="remarks"></a>Comentários

Para especificar um suplemento de COM como o suplemento equivalente, forneça os `ProgID` elementos e. `Type` Para especificar um XLL como o suplemento equivalente, forneça os `FileName` elementos e. `Type`

## <a name="see-also"></a>Confira também

- [Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Tornar o suplemento do Office compatível com um suplemento de COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
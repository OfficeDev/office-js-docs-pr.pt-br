---
title: Elemento EquivalentAddin no arquivo de manifesto
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059920"
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

[Tipo](type.md)

## <a name="can-contain"></a>Pode conter

[](progid.md)
[Nome de arquivo](filename.md) ProgID

## <a name="remarks"></a>Comentários

Para especificar um suplemento de COM como o suplemento equivalente, forneça os `ProgId` elementos e. `Type` Para especificar um XLL como o suplemento equivalente, forneça os `FileName` elementos e. `Type`

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Tornar seu suplemento do Excel compatível com um suplemento de COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
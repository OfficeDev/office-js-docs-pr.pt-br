---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611558"
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

[ProgID](progid.md) 
 [Nome do arquivo](filename.md)

## <a name="remarks"></a>Comentários

Para especificar um suplemento de COM como o suplemento equivalente, forneça os `ProgId` `Type` elementos e. Para especificar um XLL como o suplemento equivalente, forneça os `FileName` `Type` elementos e.

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Tornar seu suplemento do Excel compatível com um suplemento de COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
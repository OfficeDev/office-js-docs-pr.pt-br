---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939289"
---
# <a name="equivalentaddin-element"></a>Elemento EquivalentAddin

Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.

**Tipo de complemento:** Painel de tarefas, função Personalizada

## <a name="syntax"></a>Sintaxe

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Contido em

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Deve conter

[Tipo](type.md)

## <a name="can-contain"></a>Pode conter

[ProgId](progid.md) 
 [FileName](filename.md)

## <a name="remarks"></a>Comentários

Para especificar um complemento COM como o complemento equivalente, forneça os `ProgId` elementos `Type` e. Para especificar uma XLL como o complemento equivalente, forneça os `FileName` elementos `Type` e.

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Torne o seu suplemento do Office compatível com um suplemento COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
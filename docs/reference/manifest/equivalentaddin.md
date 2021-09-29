---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: f77a70681c8a12674d9e22022276e511552861ad
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990688"
---
# <a name="equivalentaddin-element"></a>Elemento EquivalentAddin

Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Tipo de complemento:** Painel de tarefas, Email, Função Personalizada

## <a name="syntax"></a>Sintaxe

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Contido em

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Deve conter

[Type](type.md)

## <a name="can-contain"></a>Pode conter

[ProgId](progid.md) 
 [FileName](filename.md)

## <a name="remarks"></a>Comentários

Para especificar um complemento COM como o complemento equivalente, forneça os `ProgId` elementos `Type` e. Para especificar uma XLL como o complemento equivalente, forneça os `FileName` elementos `Type` e.

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Torne o seu suplemento do Office compatível com um suplemento COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
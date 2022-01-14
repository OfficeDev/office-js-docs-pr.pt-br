---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: e318a9028ebefdeca9aaf5baac465a1ec1af0a73
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042130"
---
# <a name="equivalentaddin-element"></a>Elemento EquivalentAddin

Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Tipo de complemento:** Painel de tarefas, Email, Função Personalizada

**Válido somente nestes esquemas VersionOverrides:**

- Painel de tarefas 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

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
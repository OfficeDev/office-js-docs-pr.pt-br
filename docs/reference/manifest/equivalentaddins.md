---
title: Elemento EquivalentAddins no arquivo de manifesto
description: Especifica a compatibilidade com compatibilidade com um complemento COM equivalente, XLL ou ambos.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48f3ef86f71ad3d4f0c759df4583af4cd95e5c5a
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042151"
---
# <a name="equivalentaddins-element"></a>Elemento EquivalentAddins

Especifica a compatibilidade com compatibilidade com um complemento COM equivalente, XLL ou ambos.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Tipo de complemento:** Painel de tarefas, Email, Função Personalizada

**Válido somente nestes esquemas VersionOverrides:**

- Painel de tarefas 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="syntax"></a>Sintaxe

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## <a name="contained-in"></a>Contido em

[VersionOverrides](versionoverrides.md)

## <a name="must-contain"></a>Deve conter

[EquivalentAddin](equivalentaddin.md)

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Torne o seu suplemento do Office compatível com um suplemento COM existente](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
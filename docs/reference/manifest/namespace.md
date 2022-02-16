---
title: Elemento Namespace no arquivo de manifesto
description: O elemento Namespace define o namespace que uma função personalizada usa em Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: f9fddaca6ec8ce6128ae638c9b798efb06319ba0
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855622"
---
# <a name="namespace-element"></a>Elemento Namespace

Define o namespace usado por uma função personalizada no Excel.

**Tipo de complemento:** Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Não  | Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md). Não pode ter mais de 32 caracteres. |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<Namespace resid="namespace" />
```

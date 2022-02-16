---
title: Elemento Metadata no arquivo de manifesto
description: O elemento Metadados define as configurações de metadados que uma função personalizada usa no Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52938155442bb5424a170634d1324de77de2b788
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855531"
---
# <a name="metadata-element"></a>Elemento Metadata

Define as configurações de metadados usados por uma função personalizada no Excel.

**Tipo de complemento:** Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|  Elemento  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas. |

## <a name="example"></a>Exemplo

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```

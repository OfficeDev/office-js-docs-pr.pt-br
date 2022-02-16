---
title: Elemento Script no arquivo de manifesto
description: O elemento Script define as configurações de script que uma função personalizada usa no Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0f32314912dd66d8578750bf4818af8483c8ef36
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855538"
---
# <a name="script-element"></a>Elemento Script

Define as configurações de script usadas por uma função personalizada no Excel.

**Tipo de complemento:** Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|Elementos  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com o ID de recurso do arquivo JavaScript usado por funções personalizadas.|

## <a name="example"></a>Exemplo

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```

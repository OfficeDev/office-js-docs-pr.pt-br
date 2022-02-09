---
title: Elemento Items no arquivo de manifesto
description: Especifica os itens em um menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2249bc55db662a36cf3986ebb0b90353237d4985
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467886"
---
# <a name="items-element"></a>Elemento Items

Especifica os itens em um menu.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

## <a name="syntax"></a>Sintaxe

```XML
<Items>
...  
</Items>  
```

## <a name="contained-in"></a>Contido em

[Elemento Control do tipo Menu](control-menu.md)

## <a name="must-contain"></a>Deve conter

[Item](item.md)

## <a name="examples"></a>Exemplos

Por exemplos, consulte [Controle do tipo Menu](control-menu.md).
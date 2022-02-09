---
title: Elemento habilitado no arquivo de manifesto
description: Saiba como especificar que um Comando de Complemento está desabilitado quando o complemento é lançado.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3d83a6d117c498cc4d54dfbe73ae6d800995cb6
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467847"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle [Button](control-button.md) ou [Menu está](control-menu.md) habilitado quando o complemento é lançado. O **elemento Enabled** é um elemento filho de [Control](control.md). Se for omitido, o padrão será `true`.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Esse elemento só é válido em Excel, ou seja, `Name` quando o atributo do [elemento Host](host.md) é "Workbook".

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

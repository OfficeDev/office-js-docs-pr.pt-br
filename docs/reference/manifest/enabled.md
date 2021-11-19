---
title: Elemento habilitado no arquivo de manifesto
description: Saiba como especificar que um Comando de Complemento está desabilitado quando o complemento é lançado.
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4c0107daaf73aee6ba116553a8d01250e9c7d981
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081432"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle [Button](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) está habilitado quando o complemento é lançado. O **elemento Enabled** é um elemento filho de [Control](control.md). Se for omitido, o padrão será `true` .

**Tipo de suplemento:** Painel de tarefas

**Válido somente nestes esquemas VersionOverrides:**

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos:**

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Esse elemento só é válido no Excel, ou seja, quando o atributo do `Name` [elemento Host](host.md) é "Workbook".

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

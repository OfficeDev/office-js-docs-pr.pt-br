---
title: Elemento habilitado no arquivo de manifesto
description: Saiba como especificar que um Comando de Complemento está desabilitado quando o complemento é lançado.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 75a2664143e29c86a05aaf039b0ea7bce659cef9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744762"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle [Button](control-button.md) ou [Menu está](control-menu.md) habilitado quando o complemento é lançado. O **elemento Enabled** é um elemento filho de [Control](control.md). Se for omitido, o padrão será `true`.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Esse elemento só é válido em Excel, PowerPoint e Word; ou seja, `Name` quando o atributo do elemento [Host](host.md) é "Workbook", "Presentation" ou "Document".

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

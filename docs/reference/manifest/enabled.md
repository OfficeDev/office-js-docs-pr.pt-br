---
title: Elemento habilitado no arquivo de manifesto
description: Saiba como especificar que um Comando de Complemento está desabilitado quando o complemento é lançado.
ms.date: 01/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: a14385f7114eb3d35845b5d9873bdd718b46c0e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151814"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle [Button](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) está habilitado quando o complemento é lançado. O **elemento Enabled** é um elemento filho de [Control](control.md). Se for omitido, o padrão será `true` .

Esse elemento só é válido em Excel; ou seja, quando `Name` o atributo do elemento [Host](host.md) for "Workbook".

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

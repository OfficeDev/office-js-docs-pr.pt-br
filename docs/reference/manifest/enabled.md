---
title: Elemento habilitado no arquivo de manifesto
description: Saiba como especificar que um Comando de Complemento está desabilitado quando o complemento é lançado.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937090"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle [Button](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) está habilitado quando o complemento é lançado. O **elemento Enabled** é um elemento filho de [Control](control.md). Se for omitido, o padrão será `true` .

Esse elemento só é válido em Excel; ou seja, quando `Name` o atributo do elemento [Host](host.md) for "Workbook".

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

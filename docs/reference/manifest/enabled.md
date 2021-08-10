---
title: Elemento habilitado no arquivo de manifesto
description: Saiba como especificar que um Comando de Complemento está desabilitado quando o complemento é lançado.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 54d28839a274ff41bab0b1e2cdd2d169e76c5815095950dec67ce2564eade601
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093900"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle [Button](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) está habilitado quando o complemento é lançado. O **elemento Enabled** é um elemento filho de [Control](control.md). Se for omitido, o padrão será `true` .

Esse elemento só é válido em Excel; ou seja, quando `Name` o atributo do elemento [Host](host.md) for "Workbook".

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

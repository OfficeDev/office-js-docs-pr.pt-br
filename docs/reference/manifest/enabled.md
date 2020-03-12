---
title: Elemento Enabled no arquivo de manifesto
description: Saiba como especificar se um comando de suplemento está desabilitado quando o suplemento é iniciado.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 4c2c013c8e55966ba2678755536ce04ae3014ed0
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596897"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) está habilitado quando o suplemento é iniciado. O elemento **Enabled** é um elemento filho do [controle](control.md). Se for omitido, o padrão será `true`.

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

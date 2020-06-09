---
title: Elemento Enabled no arquivo de manifesto
description: Saiba como especificar se um comando de suplemento está desabilitado quando o suplemento é iniciado.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611565"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) está habilitado quando o suplemento é iniciado. O elemento **Enabled** é um elemento filho do [controle](control.md). Se for omitido, o padrão será `true` .

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```

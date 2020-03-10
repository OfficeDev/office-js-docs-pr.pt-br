---
title: Elemento Enabled no arquivo de manifesto
description: Saiba como especificar se um comando de suplemento está desabilitado quando o suplemento é iniciado.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566181"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) está habilitado quando o suplemento é iniciado. O elemento **Enabled** é um elemento filho do [controle](control.md). Se for omitido, o padrão será `true`. 

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [habilitar e desabilitar comandos de suplemento](/office/dev/add-ins/design/disable-add-in-commands).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```


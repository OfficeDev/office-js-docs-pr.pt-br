---
title: Elemento CustomTab no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c1c3c6883a1feb94299feb35c078431e6e2e322c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450629"
---
# <a name="customtab-element"></a>Elemento CustomTab

Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Início**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.

Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.

O atributo **id** deve ser exclusivo dentro do manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Sim |  Define um grupo de comandos  |
|  [Label](#label-tab)      | Sim |  O rótulo para CustomTab ou Group.  |
|  [Control](control.md)    | Sim |  Conjunto de um ou mais objetos Control.  |

### <a name="group"></a>Group

Obrigatório. Confira [Elemento Group](group.md)

### <a name="label-tab"></a>Label (Tab)

Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).


## <a name="customtab-example"></a>Exemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

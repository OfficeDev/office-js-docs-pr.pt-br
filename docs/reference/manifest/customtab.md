---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: a81b64a17eeeb463d55024e189b09048b2eb96ac
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612301"
---
# <a name="customtab-element"></a>Elemento CustomTab

Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser a guia padrão ( **página inicial**, de **mensagem**ou **reunião**) ou em uma guia personalizada definida pelo suplemento.

Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.

O atributo **ID** deve ser exclusivo dentro do manifesto.

> [!IMPORTANT]
> No Outlook no Mac, o `CustomTab` elemento não está disponível, portanto, você terá que usar o [OfficeTab](officetab.md) .

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Sim |  Define um grupo de comandos  |
|  [Label](#label-tab)      | Sim |  O rótulo para CustomTab ou Group.  |

### <a name="group"></a>Group

Obrigatório. Confira [Elemento Group](group.md)

### <a name="label-tab"></a>Label (Tab)

Obrigatório. O rótulo da guia personalizado. O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .


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

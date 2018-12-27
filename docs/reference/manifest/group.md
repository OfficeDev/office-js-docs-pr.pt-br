---
title: Elemento Group no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 13cd9bbe6f602fd1779caea487e34177c3e9d483
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433689"
---
# <a name="group-element"></a>Elemento Group

Define um grupo de controles de interface do usuário em uma guia.  Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Sim  | Identificação exclusiva do grupo.|

### <a name="id-attribute"></a>atributo id

Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.

## <a name="child-elements"></a>Elementos filho
|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Rótulo](#label)      | Sim |  O rótulo para a CustomTab ou um grupo.  |
|  [Control](#control)    | Sim |  Conjunto de um ou mais objetos Control.  |

### <a name="label"></a>Rótulo 

Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).

### <a name="control"></a>Control
Um grupo exige pelo menos um controle.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
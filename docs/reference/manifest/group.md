---
title: Elemento Group no arquivo de manifesto
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 35db4829b40078e97fbfc007e2fb552e00875f9c
ms.sourcegitcommit: 164b11b1e9d2ae20b3d816092025b32a9070450f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2019
ms.locfileid: "39818724"
---
# <a name="group-element"></a>Elemento Group

Define um grupo de controles de interface do usuário em uma guia.  Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Sim  | Identificação exclusiva do grupo.|

### <a name="id-attribute"></a>id attribute

Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.

## <a name="child-elements"></a>Elementos filho
|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Label](#label)      | Sim |  O rótulo para a CustomTab ou um grupo.  |
|  [Icon](icon.md)      | Sim |  A imagem de um grupo.  |
|  [Control](#control)    | Sim |  Conjunto de um ou mais objetos Control.  |

### <a name="label"></a>Label 

Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).

### <a name="icon"></a>Ícone

Obrigatório. Se uma guia contiver muitos grupos e a janela do programa for redimensionada, a imagem especificada poderá ser exibida.

### <a name="control"></a>Control
Um grupo exige pelo menos um controle. Para obter detalhes sobre os tipos de controles suportados, consulte o elemento [Control](control.md) .

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

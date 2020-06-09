---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles da interface do usuário em uma guia.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: a598232f230a120dccd58024e760c2172a769727
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611824"
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

### <a name="label"></a>Rótulo 

Obrigatório. O rótulo do grupo. O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .

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

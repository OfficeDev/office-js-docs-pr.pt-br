---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles de interface do usuário em uma guia.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173959"
---
# <a name="group-element"></a>Elemento Group

Define um grupo de controles de interface do usuário em uma guia. Em guias personalizadas, o complemento pode criar vários grupos. Os suplementos estão limitados a uma guia personalizada.

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
|  [Control](#control)    | Não |  Representa um objeto Control . Pode ser zero ou mais.  |
|  [OfficeControl](#officecontrol)  | Não | Representa um dos controles internos do Office. Pode ser zero ou mais. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Não |  Especifica se o grupo deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas.  |

### <a name="label"></a>Rótulo

Obrigatório. O rótulo do grupo. O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources.](resources.md)

### <a name="icon"></a>Ícone

Obrigatório. Se uma guia contiver muitos grupos e a janela do programa for resizedida, a imagem especificada poderá ser exibida em vez disso.

### <a name="control"></a>Controle

Opcional, mas se não houver deve haver pelo menos um **OfficeControl**. Para obter detalhes sobre os tipos de controles com suporte, consulte o [elemento Control.](control.md) A ordem de **Controle** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon.**

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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

### <a name="officecontrol"></a>OfficeControl

Opcional, mas se não houver deve haver pelo menos um **controle**. Inclua um ou mais controles internos do Office no grupo com `<OfficeControl>` elementos. O `id` atributo especifica a ID do controle office integrado. Para encontrar a ID de um controle, consulte [Encontrar as IDs de controles e grupos de controles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) A ordem de **Controle** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon.**

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Opcional (booliana). Especifica se  o grupo ficará oculto em combinações de aplicativos e plataformas que suportam uma API que instala uma guia contextual personalizada na faixa de opções em tempo de execução. O valor padrão, se não estiver presente, é `false` . Se usado, **OverriddenByRibbonApi** deve ser o *primeiro* filho de **Group**. Para obter mais informações, [consulte OverriddenByRibbonApi](overriddenbyribbonapi.md).

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

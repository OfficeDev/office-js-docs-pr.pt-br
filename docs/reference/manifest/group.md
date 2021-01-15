---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles da interface do usuário em uma guia.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 3872ece926cc399ed2b30d4dabaacfb741e060ab
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771391"
---
# <a name="group-element"></a>Elemento Group

Define um grupo de controles da interface do usuário em uma guia. Nas guias personalizadas, o suplemento pode criar vários grupos. Os suplementos estão limitados a uma guia personalizada.

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
|  [Control](#control)    | Não |  Representa um objeto Control. Pode ser zero ou mais.  |
|  [OfficeControl](#officecontrol)  | Não | Representa um dos controles internos do Office. Pode ser zero ou mais. |

### <a name="label"></a>Rótulo

Obrigatório. O rótulo do grupo. O atributo **Resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .

### <a name="icon"></a>Ícone

Obrigatório. Se uma guia contiver muitos grupos e a janela do programa for redimensionada, a imagem especificada poderá ser exibida.

### <a name="control"></a>Controle

Opcional, mas, se não estiver presente, deve haver pelo menos um **OfficeControl**. Para obter detalhes sobre os tipos de controles suportados, consulte o elemento [Control](control.md) . A ordem de **controle** e **OfficeControl** no manifesto é intercambiável e podem ser mescladas se houver vários elementos, mas todos devem estar abaixo do elemento **Icon** .

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

### <a name="officecontrol"></a>OfficeControl

Opcional, mas, se não estiver presente, deve haver pelo menos um **controle**. Inclua um ou mais controles internos do Office no grupo com `<OfficeControl>` elementos. O `id` atributo especifica a ID do controle interno do Office. Para localizar a ID de um controle, confira [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). A ordem de **controle** e **OfficeControl** no manifesto é intercambiável e podem ser mescladas se houver vários elementos, mas todos devem estar abaixo do elemento **Icon** .

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
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

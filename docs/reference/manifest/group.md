---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles de interface do usuário em uma guia.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4717f6aeff3cd8ac34ee289252054417c489b89
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340460"
---
# <a name="group-element"></a>Elemento Group

Define um grupo de controles de interface do usuário em uma guia. Em guias personalizadas, o complemento pode criar vários grupos. Os suplementos estão limitados a uma guia personalizada.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Sim  | Identificação exclusiva do grupo.|

### <a name="id-attribute"></a>id attribute

Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Isso deve ser exclusivo em todos os elementos Group no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Label](#label)      | Sim |  O rótulo de um grupo.  |
|  [Icon](icon.md)      | Sim |  A imagem de um grupo. Não há suporte em Outlook de complementos. |
|  [Control](#control)    | Não |  Representa um objeto Control. Pode ser zero ou mais.  |
|  [OfficeControl](#officecontrol)  | Não | Representa um dos controles internos Office internos. Pode ser zero ou mais. Não há suporte em Outlook de complementos.|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Não |  Especifica se o grupo deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas. Não há suporte em Outlook de complementos. |

### <a name="label"></a>Rótulo

Obrigatório. O rótulo do grupo. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources](resources.md) .

### <a name="icon"></a>Ícone

Obrigatório. Se uma guia contiver muitos grupos e a janela do programa for resized, a imagem especificada poderá ser exibida.

> [!NOTE]
> Esse elemento filho não é suportado em Outlook de complementos.

### <a name="control"></a>Controle

Opcional, mas se não estiver presente, deve haver pelo menos um **OfficeControl**. Para obter detalhes sobre os tipos de controles com suporte, consulte o [elemento Control](control.md) . A ordem de **Control** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon** .

```xml
<Group id="Contoso.CustomTab1.group1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button1">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a>OfficeControl

Opcional, mas se não estiver presente, deve haver pelo menos um **Controle**. Inclua um ou mais controles internos Office no grupo com `<OfficeControl>` elementos. O `id` atributo especifica a ID do controle Office integrado. Para encontrar a ID de um controle, consulte [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). A ordem de **Control** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon** .

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> Esse elemento filho não é suportado em Outlook de complementos.

```xml
<Group id="Contoso.CustomTab2.group2">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Opcional (booleano). Especifica se o **Grupo** ficará oculto em combinações de aplicativos e plataformas que suportam uma API que instala uma guia contextual personalizada na faixa de opções no tempo de execução. O valor padrão, se não estiver presente, é `false`. Se usado, **OverriddenByRibbonApi** deve ser o *primeiro* filho de **Group**. Para obter mais informações, [consulte OverriddenByRibbonApi](overriddenbyribbonapi.md).

> [!NOTE]
> Esse elemento filho não é suportado em Outlook de complementos.

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.CustomTab3">
    <Group id="Contoso.CustomTab3.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6a9540fd7e98464681a90021a36f7a7529186f7f
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340110"
---
# <a name="customtab-element"></a>Elemento CustomTab

Define uma guia personalizada para a faixa Office faixa de opções. Adicione controles de faixa de opções e grupos para o add-in a uma das guias de Office ou à sua própria guia personalizada. Use o elemento **CustomTab** para adicionar uma guia personalizada à faixa de opções. Em guias personalizadas, o complemento pode ter grupos personalizados ou integrados. Os suplementos estão limitados a uma guia personalizada.

> [!IMPORTANT]
> No Outlook no Mac, o **elemento CustomTab** não está disponível, mas você pode colocar grupos personalizados de  controles em um dos [OfficeTabs](officetab.md) internos. Não é possível *colocar grupos integrados* *em guias Outlook* em qualquer plataforma.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

> [!NOTE]
> Alguns elementos filho não são válidos nos esquemas de email. Consulte [Elementos Filho](#child-elements).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)
- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md). Obrigatório por alguns elementos filho. Consulte [Elementos Filho](#child-elements).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Sim  | Uma ID exclusiva para a guia personalizada.|

### <a name="id-attribute"></a>id attribute

Obrigatório. Identificador exclusivo da guia personalizada. É uma cadeia de caracteres com no máximo 125 caracteres. Isso deve ser exclusivo no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Não |  Define um grupo de comandos  |
|  [OfficeGroup](#officegroup)      | Não |  Representa um grupo de controle Office integrado. **Importante**: não disponível no Outlook. |
|  [Label](#label-tab)      | Sim |  O rótulo do CustomTab.  |
|  [InsertAfter](#insertafter)      | Não |  Especifica que a guia personalizada deve ser imediatamente após uma guia Office de Office. **Importante**: disponível somente no PowerPoint. |
|  [InsertBefore](#insertbefore)      | Não |  Especifica que a guia personalizada deve ser imediatamente antes de uma guia Office de Office. **Importante**: disponível somente no PowerPoint. |

### <a name="group"></a>Group

Opcional, mas se não estiver presente, deve haver pelo menos um **elemento OfficeGroup** . Confira [Elemento Group](group.md) A ordem de **Group** e **OfficeGroup** no manifesto deve ser a ordem que você deseja que eles apareçam na guia personalizada. Eles podem ser intermendados se houver vários elementos, mas todos devem estar acima do **elemento Label** .

### <a name="officegroup"></a>OfficeGroup

Opcional, mas se não estiver presente, deve haver pelo menos um **elemento Group** . Representa um grupo de controle Office integrado. O **atributo id** especifica a ID do grupo de Office integrado. Para encontrar a ID de um grupo integrado, consulte [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). A ordem de **Group** e **OfficeGroup** no manifesto deve ser a ordem que você deseja que eles apareçam na guia personalizada. Eles podem ser intermendados se houver vários elementos, mas todos devem estar acima do **elemento Label** .

> [!IMPORTANT]
> O **elemento OfficeGroup** não está disponível no Outlook. No PowerPoint, ele está em visualização para Mac e Windows; mas está disponível para os complementos de produção no PowerPoint na Web.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="label-tab"></a>Label (Tab)

Obrigatório. O rótulo da guia personalizada. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources](resources.md) .

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertafter"></a>InsertAfter

Opcional. Especifica que a guia personalizada deve ser imediatamente após uma guia Office. O valor do elemento é a ID da guia integrado, como `TabHome` ou `TabReview`.  Para ver uma lista de guias integrados, consulte [OfficeTab](officetab.md). Se presente, deve estar após o **elemento Label** . Não é possível ter **InsertAfter e** **InsertBefore**.

> [!IMPORTANT]
> O **elemento InsertAfter** só está disponível no PowerPoint.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertbefore"></a>InsertBefore

Opcional. Especifica que a guia personalizada deve ser imediatamente antes de uma guia Office. O valor do elemento é a ID da guia integrado, como `TabHome` ou `TabReview`. O valor do elemento é a ID da guia integrado, como `TabHome` ou `TabReview`.  Para ver uma lista de guias integrados, consulte [OfficeTab](officetab.md). Se presente, deve estar após o **elemento Label** . Não é possível ter **InsertAfter e** **InsertBefore**.

> [!IMPORTANT]
> O **elemento InsertBefore** só está disponível no PowerPoint.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)


## <a name="examples"></a>Exemplos

O exemplo de marcação a seguir adiciona o grupo de controle Office Paragraph a uma guia personalizada e posiciona-o para aparecer logo após um grupo personalizado.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.TabCustom1.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

O exemplo de marcação a seguir adiciona o Office sobrescrito a um grupo personalizado e posiciona-o para aparecer logo após um botão personalizado.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

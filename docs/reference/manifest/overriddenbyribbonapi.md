---
title: Elemento OverriddenByRibbonApi no arquivo de manifesto
description: Saiba como especificar que uma guia, grupo, controle ou item de menu personalizado não deve aparecer quando também faz parte de uma guia contextual personalizada.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 62aa484057221f9cd7f41af9c8b9210cdb5b3376
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173994"
---
# <a name="overriddenbyribbonapi-element"></a><span data-ttu-id="43ae0-103">Elemento OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="43ae0-103">OverriddenByRibbonApi element</span></span>

<span data-ttu-id="43ae0-104">Especifica se um [CustomTab](customtab.md) [,](group.md)grupo [,](control.md#button-control) controle de botão, controle de [menu](control.md#menu-dropdown-button-controls) ou item de menu será ocultado em combinações de aplicativo e plataforma que suportam a API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) que instala guias contextuais personalizadas na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="43ae0-104">Specifies whether a [CustomTab](customtab.md), [Group](group.md), [Button](control.md#button-control) control, [Menu](control.md#menu-dropdown-button-controls) control, or menu item will be hidden on application and platform combinations that support the API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) that installs custom contextual tabs on the ribbon.</span></span>

<span data-ttu-id="43ae0-105">Se for omitido, o padrão é `false` .</span><span class="sxs-lookup"><span data-stu-id="43ae0-105">If it is omitted, the default is `false`.</span></span> <span data-ttu-id="43ae0-106">Se for usado, ele deverá ser o *primeiro* elemento filho de seu elemento pai.</span><span class="sxs-lookup"><span data-stu-id="43ae0-106">If it is used, it must be the *first* child element of its parent element.</span></span>

> [!NOTE]
> <span data-ttu-id="43ae0-107">Para ter uma compreensão completa desse elemento, leia Implementar uma experiência de interface do usuário alternativa quando guias [contextuais personalizadas não são suportadas.](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)</span><span class="sxs-lookup"><span data-stu-id="43ae0-107">For a full understanding of this element, please read [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

<span data-ttu-id="43ae0-108">O objetivo desse elemento é criar uma experiência de fallback em um add-in que implemente guias contextuais personalizadas quando o complemento é executado em um aplicativo ou plataforma que não dá suporte a guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="43ae0-108">The purpose of this element is to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> <span data-ttu-id="43ae0-109">A estratégia essencial é duplicar alguns ou todos os grupos e controles da guia contextual personalizada em uma ou mais guias principais personalizadas (ou *seja,* guias personalizadas não textuais).</span><span class="sxs-lookup"><span data-stu-id="43ae0-109">The essential strategy is that you duplicate some or all of the groups and controls from your custom contextual tab onto one or more custom core tabs (that is, *noncontextual* custom tabs).</span></span> <span data-ttu-id="43ae0-110">Em seguida, para garantir que esses grupos e  controles apareçam quando guias contextuais personalizadas não são suportadas, mas não aparecem quando guias contextuais *personalizadas* são suportadas, você adiciona como o primeiro elemento filho dos elementos `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab**, **Group**, **Control** ou **menu Item.**</span><span class="sxs-lookup"><span data-stu-id="43ae0-110">Then, to ensure that these groups and controls appear when custom contextual tabs are *not* supported, but do not appear when custom contextual tabs *are* supported, you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the **CustomTab**, **Group**, **Control**, or menu **Item** elements.</span></span> <span data-ttu-id="43ae0-111">O efeito de fazer isso é o seguinte:</span><span class="sxs-lookup"><span data-stu-id="43ae0-111">The effect of doing so is the following:</span></span>

- <span data-ttu-id="43ae0-112">Se o complemento for executado em um aplicativo e plataforma que suportam guias contextuais personalizadas, as guias, grupos e controles duplicados não aparecerão na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="43ae0-112">If the add-in runs on an application and platform that support custom contextual tabs, then the duplicated tabs, groups, and controls won't appear on the ribbon.</span></span> <span data-ttu-id="43ae0-113">Em vez disso, a guia contextual personalizada será instalada quando o complemento chamar o `requestCreateControls` método.</span><span class="sxs-lookup"><span data-stu-id="43ae0-113">Instead, the custom contextual tab will be installed when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="43ae0-114">Se o complemento for executado em  um aplicativo ou plataforma que não dá suporte a guias contextuais personalizadas, as guias, os grupos e os controles duplicados aparecerão na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="43ae0-114">If the add-in runs on an application or platform that *doesn't* support custom contextual tabs, then the duplicated tabs, groups, and controls will appear on the ribbon.</span></span>

## <a name="examples"></a><span data-ttu-id="43ae0-115">Exemplos</span><span class="sxs-lookup"><span data-stu-id="43ae0-115">Examples</span></span>

### <a name="overriding-an-entire-tab"></a><span data-ttu-id="43ae0-116">Substituindo uma guia inteira</span><span class="sxs-lookup"><span data-stu-id="43ae0-116">Overriding an entire tab</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-group"></a><span data-ttu-id="43ae0-117">Substituindo um grupo</span><span class="sxs-lookup"><span data-stu-id="43ae0-117">Overriding a group</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-control"></a><span data-ttu-id="43ae0-118">Substituindo um controle</span><span class="sxs-lookup"><span data-stu-id="43ae0-118">Overriding a control</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-menu-item"></a><span data-ttu-id="43ae0-119">Substituindo um item de menu</span><span class="sxs-lookup"><span data-stu-id="43ae0-119">Overriding a menu item</span></span>


```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Menu" id="MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

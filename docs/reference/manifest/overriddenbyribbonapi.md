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
# <a name="overriddenbyribbonapi-element"></a>Elemento OverriddenByRibbonApi

Especifica se um [CustomTab](customtab.md) [,](group.md)grupo [,](control.md#button-control) controle de botão, controle de [menu](control.md#menu-dropdown-button-controls) ou item de menu será ocultado em combinações de aplicativo e plataforma que suportam a API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) que instala guias contextuais personalizadas na faixa de opções.

Se for omitido, o padrão é `false` . Se for usado, ele deverá ser o *primeiro* elemento filho de seu elemento pai.

> [!NOTE]
> Para ter uma compreensão completa desse elemento, leia Implementar uma experiência de interface do usuário alternativa quando guias [contextuais personalizadas não são suportadas.](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)

O objetivo desse elemento é criar uma experiência de fallback em um add-in que implemente guias contextuais personalizadas quando o complemento é executado em um aplicativo ou plataforma que não dá suporte a guias contextuais personalizadas. A estratégia essencial é duplicar alguns ou todos os grupos e controles da guia contextual personalizada em uma ou mais guias principais personalizadas (ou *seja,* guias personalizadas não textuais). Em seguida, para garantir que esses grupos e  controles apareçam quando guias contextuais personalizadas não são suportadas, mas não aparecem quando guias contextuais *personalizadas* são suportadas, você adiciona como o primeiro elemento filho dos elementos `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab**, **Group**, **Control** ou **menu Item.** O efeito de fazer isso é o seguinte:

- Se o complemento for executado em um aplicativo e plataforma que suportam guias contextuais personalizadas, as guias, grupos e controles duplicados não aparecerão na faixa de opções. Em vez disso, a guia contextual personalizada será instalada quando o complemento chamar o `requestCreateControls` método.
- Se o complemento for executado em  um aplicativo ou plataforma que não dá suporte a guias contextuais personalizadas, as guias, os grupos e os controles duplicados aparecerão na faixa de opções.

## <a name="examples"></a>Exemplos

### <a name="overriding-an-entire-tab"></a>Substituindo uma guia inteira

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

### <a name="overriding-a-group"></a>Substituindo um grupo

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

### <a name="overriding-a-control"></a>Substituindo um controle

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

### <a name="overriding-a-menu-item"></a>Substituindo um item de menu


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

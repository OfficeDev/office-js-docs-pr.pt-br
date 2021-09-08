---
title: Elemento OverriddenByRibbonApi no arquivo de manifesto
description: Saiba como especificar que uma guia, grupo, controle ou item de menu personalizado não deve aparecer quando também faz parte de uma guia contextual personalizada.
ms.date: 09/02/2021
localization_priority: Normal
ms.openlocfilehash: b2633fac0c83d1e9c2195efd155496a0dacafad7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936229"
---
# <a name="overriddenbyribbonapi-element"></a>Elemento OverriddenByRibbonApi

Especifica se um [grupo](group.md) [,](control.md#button-control) controle button, controle [menu](control.md#menu-dropdown-button-controls) ou item de menu será oculto em combinações de aplicativos e plataformas que suportam a API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)) que instala guias contextuais personalizadas na faixa de opções.

Se for omitido, o padrão será `false` . Se for usado, ele deve ser o *primeiro* elemento filho de seu elemento pai.

> [!NOTE]
> Para uma compreensão completa desse elemento, leia Implementar uma experiência de interface do usuário alternativa quando as guias [contextuais personalizadas não são suportadas](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

O objetivo deste elemento é criar uma experiência de fallback em um add-in que implemente guias contextuais personalizadas quando o add-in está sendo executado em um aplicativo ou plataforma que não oferece suporte a guias contextuais personalizadas. A estratégia essencial é duplicar alguns ou todos os grupos e controles de sua guia contextual personalizada em uma ou mais guias principais personalizadas (ou seja, guias personalizadas *nãocontextuais).* Em seguida, para garantir que esses grupos e  controles apareçam quando as guias contextuais personalizadas não são suportadas, mas não aparecem quando as guias contextuais *personalizadas* são suportadas, você adiciona como o primeiro elemento filho dos elementos `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **Group,** **Control** ou **Item** do menu. O efeito de fazer isso é o seguinte:

- Se o complemento for executado em um aplicativo e plataforma que suportam guias contextuais personalizadas, os grupos e controles duplicados não aparecerão na faixa de opções. Em vez disso, a guia contextual personalizada será instalada quando o complemento chamar o `requestCreateControls` método.
- Se o complemento for executado em  um aplicativo ou plataforma que não oferece suporte a guias contextuais personalizadas, os grupos e controles duplicados aparecerão na faixa de opções.

## <a name="examples"></a>Exemplos

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

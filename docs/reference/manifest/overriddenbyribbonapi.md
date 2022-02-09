---
title: Elemento OverriddenByRibbonApi no arquivo de manifesto
description: Saiba como especificar que uma guia, grupo, controle ou item de menu personalizado não deve aparecer quando também faz parte de uma guia contextual personalizada.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48977691ee4bf2ccd71bc146647dae452ce9e2fc
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467684"
---
# <a name="overriddenbyribbonapi-element"></a>Elemento OverriddenByRibbonApi

Especifica se um [grupo](group.md), controle [button](control-button.md), controle [menu](control-menu.md) ou item de menu será oculto em combinações de aplicativos e plataformas que suportam a API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1))) que instala guias contextuais personalizadas na faixa de opções.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [Faixa de opções 1.2](../requirement-sets/add-in-commands-requirement-sets.md) (Obrigatório para Excel, PowerPoint e Word.)

Se esse elemento for omitido, o padrão será `false`. Se for usado, ele deve ser o *primeiro* elemento filho de seu elemento pai.

> [!NOTE]
> Para uma compreensão completa desse elemento, leia Implementar uma [experiência de interface do usuário alternativa quando guias contextuais personalizadas não são suportadas](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

O objetivo deste elemento é criar uma experiência de fallback em um add-in que implemente guias contextuais personalizadas quando o add-in está sendo executado em um aplicativo ou plataforma que não oferece suporte a guias contextuais personalizadas. A estratégia essencial é duplicar alguns ou todos os grupos e controles de sua guia contextual personalizada em uma ou mais guias principais personalizadas (ou seja, guias personalizadas *nãocontextuais* ). Em seguida, para garantir que esses grupos e controles apareçam quando as guias contextuais personalizadas não são suportadas, mas não aparecem quando as guias contextuais *personalizadas são* suportadas, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` você adiciona como o primeiro elemento filho dos elementos **Group**, **Control** ou menu **Item**. O efeito de fazer isso é o seguinte:

- Se o complemento for executado em um aplicativo e plataforma que suportam guias contextuais personalizadas, os grupos e controles duplicados não aparecerão na faixa de opções. Em vez disso, a guia contextual personalizada será instalada quando o complemento chamar o `requestCreateControls` método.
- Se o complemento for executado em um aplicativo ou plataforma que não  oferece suporte a guias contextuais personalizadas, os grupos e controles duplicados aparecerão na faixa de opções.

## <a name="examples"></a>Exemplos

### <a name="overriding-a-group"></a>Substituindo um grupo

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.CustomTab1.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="Contoso.MyButton1">
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
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.CustomTab2.group2">
      <Control  xsi:type="Button" id="Contoso.MyButton2">
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
  <CustomTab id="Contoso.TabCustom3">
    <Group id="Contoso.CustomTab3.group3">
      <Control  xsi:type="Menu" id="Contoso.MyMenu">
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

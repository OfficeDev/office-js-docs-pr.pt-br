---
title: Integrar botões integrados do Office a grupos de controle personalizados e guias
description: Saiba como incluir botões do Office integrados em seus grupos de comandos personalizados e guias na faixa de opções do Office.
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 8d4e8f39313551d001669b948b146250114f3e06
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505252"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Integrar botões integrados do Office a grupos de controle personalizados e guias

Você pode inserir botões do Office integrados em seus grupos de controle personalizados na faixa de opções do Office usando marcação no manifesto do complemento. (Você não pode inserir seus comandos de complemento personalizados em um grupo do Office integrado.) Você também pode inserir grupos de controle do Office inteiros em suas guias de faixa de opções personalizadas.

> [!NOTE]
> Este artigo supõe que você está familiarizado com o artigo [Conceitos básicos para comandos de complemento.](add-in-commands.md) Revise-o se não tiver feito isso recentemente.

> [!IMPORTANT]
>
> - O recurso de add-in e a marcação descritos neste artigo só estão disponíveis no *PowerPoint na Web*.
> - A marcação descrita neste artigo só funciona em plataformas que suportam o conjunto de **requisitos AddinCommands 1.3**. Consulte a seção Comportamento [posterior em plataformas sem suporte.](#behavior-on-unsupported-platforms)

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Inserir um grupo de controle integrado em uma guia personalizada

Para inserir um grupo de controle do Office integrado em uma guia, adicione um elemento [OfficeGroup](../reference/manifest/customtab.md#officegroup) como um elemento filho no elemento `<CustomTab>` pai. O `id` atributo do elemento é definido como a `<OfficeGroup>` ID do grupo integrado. Consulte [Encontre as IDs de controles e grupos de controle.](#find-the-ids-of-controls-and-control-groups)

O exemplo de marcação a seguir adiciona o grupo de controle Parágrafo do Office a uma guia personalizada e posiciona-o para aparecer logo após um grupo personalizado.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Inserir um controle integrado em um grupo personalizado

Para inserir um controle do Office integrado em um grupo personalizado, adicione um elemento [OfficeControl](../reference/manifest/group.md#officecontrol) como um elemento filho no elemento `<Group>` pai. O `id` atributo do elemento é definido como a `<OfficeControl>` ID do controle integrado. Consulte [Encontre as IDs de controles e grupos de controle.](#find-the-ids-of-controls-and-control-groups)

O exemplo de marcação a seguir adiciona o controle Sobrescrito do Office a um grupo personalizado e posiciona-o para aparecer logo após um botão personalizado.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> Os usuários podem personalizar a faixa de opções no aplicativo do Office. Qualquer personalização do usuário substituirá suas configurações de manifesto. Por exemplo, um usuário pode remover um botão de qualquer grupo e remover qualquer grupo de uma guia.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Encontre as IDs de controles e grupos de controle

As IDs para controles suportados e grupos de controle estão em arquivos nas [IDs](https://github.com/OfficeDev/office-control-ids)de controle do Office de repo . Siga as instruções no arquivo ReadMe desse repo.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento em plataformas sem suporte

Se o seu add-in estiver instalado em uma plataforma que não oferece suporte ao conjunto de [requisitos AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), a marcação descrita neste artigo será ignorada e os controles/grupos internos do Office não aparecerão em seus grupos/guias personalizados. Para impedir que o seu complemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na seção `<Requirements>` do manifesto. Para obter instruções, consulte [Definir o elemento Requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest). Como alternativa, você pode projetar seu complemento para ter uma experiência alternativa quando **AddinCommands 1.3** não tiver suporte, conforme descrito em Usar verificações de tempo de execução em seu código [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). Por exemplo, se o seu add-in contiver instruções que pressuem que os botões integrados estão em seus grupos personalizados, você pode ter uma versão alternativa que presume que os botões integrados estão apenas em seus locais usuais.
